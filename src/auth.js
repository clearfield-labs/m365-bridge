// Per-user sign-in. Each person signs in with their own Microsoft account, so Microsoft
// enforces their own SharePoint/OneDrive permissions. No client secret ever lives on a laptop.
import fs from 'node:fs';
import { PublicClientApplication, CryptoProvider, InteractionRequiredAuthError, LogLevel } from '@azure/msal-node';
import { FILES, REDIRECT_URI } from './paths.js';
import { loadConfig, isConfigured } from './config.js';
import { readEncrypted, writeEncrypted } from './vault.js';
import { scopesFor } from './scopes.js';

export class NotSignedInError extends Error {
  constructor(message) {
    super(message);
    this.name = 'NotSignedInError';
  }
}

export { scopesFor };

/** The scopes Microsoft actually put in this person's token (what the admin consented to). */
export async function grantedScopes() {
  try {
    const token = await getAccessToken();
    const payload = JSON.parse(Buffer.from(token.split('.')[1], 'base64url').toString('utf8'));
    return String(payload.scp || '').split(' ').filter(Boolean).sort();
  } catch {
    return null;
  }
}

const cachePlugin = {
  async beforeCacheAccess(ctx) {
    const data = readEncrypted(FILES.tokens);
    if (data) ctx.tokenCache.deserialize(data);
  },
  async afterCacheAccess(ctx) {
    if (ctx.cacheHasChanged) writeEncrypted(FILES.tokens, ctx.tokenCache.serialize());
  },
};

let pca = null;
let pcaKey = '';

function client() {
  const cfg = loadConfig();
  if (!isConfigured(cfg)) {
    throw new NotSignedInError('M365 Bridge is not set up yet. Open the dashboard and enter the firm Tenant ID and Client ID.');
  }
  const key = `${cfg.tenantId}/${cfg.clientId}`;
  if (!pca || key !== pcaKey) {
    pca = new PublicClientApplication({
      auth: { clientId: cfg.clientId, authority: `https://login.microsoftonline.com/${cfg.tenantId}` },
      cache: { cachePlugin },
      system: { loggerOptions: { loggerCallback() {}, piiLoggingEnabled: false, logLevel: LogLevel.Error } },
    });
    pcaKey = key;
  }
  return { pca, cfg };
}

export async function currentAccount() {
  if (process.env.M365_TEST_TOKEN) {
    return { username: process.env.M365_TEST_USER || 'test@example.com', name: process.env.M365_TEST_NAME || 'Test User' };
  }
  try {
    const { pca } = client();
    const accounts = await pca.getTokenCache().getAllAccounts();
    return accounts[0] || null;
  } catch {
    return null;
  }
}

export async function getAccessToken() {
  if (process.env.M365_TEST_TOKEN) return process.env.M365_TEST_TOKEN;
  const { pca, cfg } = client();
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (!accounts.length) {
    throw new NotSignedInError('Not signed in to Microsoft 365. Use the sign_in tool, or open the dashboard and click "Sign in with Microsoft".');
  }
  try {
    const res = await pca.acquireTokenSilent({ account: accounts[0], scopes: scopesFor(cfg) });
    return res.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      throw new NotSignedInError('Your Microsoft 365 session expired. Sign in again from the dashboard or with the sign_in tool.');
    }
    throw err;
  }
}

// --- Browser sign-in (authorization code + PKCE on localhost) ---
const pending = new Map(); // state -> { verifier, created }

export async function startBrowserSignIn() {
  const { pca, cfg } = client();
  const crypto = new CryptoProvider();
  const { verifier, challenge } = await crypto.generatePkceCodes();
  const state = crypto.createNewGuid();
  pending.set(state, { verifier, created: Date.now() });
  const url = await pca.getAuthCodeUrl({
    scopes: scopesFor(cfg),
    redirectUri: REDIRECT_URI,
    codeChallenge: challenge,
    codeChallengeMethod: 'S256',
    state,
    ...(cfg.loginHint ? { loginHint: cfg.loginHint } : { prompt: 'select_account' }),
  });
  return url;
}

export async function finishBrowserSignIn(code, state) {
  const entry = pending.get(state);
  if (!entry || Date.now() - entry.created > 15 * 60 * 1000) {
    throw new Error('This sign-in link expired. Start again from the dashboard.');
  }
  pending.delete(state);
  const { pca, cfg } = client();
  const res = await pca.acquireTokenByCode({
    code,
    scopes: scopesFor(cfg),
    redirectUri: REDIRECT_URI,
    codeVerifier: entry.verifier,
  });
  return res.account;
}

// --- Device code sign-in (works from inside Claude without a browser redirect) ---
let deviceFlow = null; // { message, done, account, error }

export async function startDeviceCodeSignIn() {
  if (deviceFlow && !deviceFlow.done) return deviceFlow.message;
  const { pca, cfg } = client();
  deviceFlow = { message: null, done: false, account: null, error: null };
  const flow = deviceFlow;
  const messageReady = new Promise((resolve) => {
    pca
      .acquireTokenByDeviceCode({
        scopes: scopesFor(cfg),
        deviceCodeCallback: (r) => {
          flow.message = r.message;
          resolve(r.message);
        },
      })
      .then((res) => { flow.account = res?.account || null; })
      .catch((e) => { flow.error = e.message; resolve(`Sign-in failed: ${e.message}`); })
      .finally(() => { flow.done = true; });
  });
  return messageReady;
}

export function deviceCodeState() {
  return deviceFlow;
}

export async function signOut() {
  try {
    const { pca } = client();
    const cache = pca.getTokenCache();
    for (const a of await cache.getAllAccounts()) await cache.removeAccount(a);
  } catch { /* not configured */ }
  fs.rmSync(FILES.tokens, { force: true });
}
