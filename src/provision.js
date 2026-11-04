// Firm setup, as one guided flow for the Microsoft 365 admin:
//   1. pick permissions          (saved locally first, so the admin approves once)
//   2. admin signs in            (Microsoft's own "Microsoft Graph Command Line Tools" sign-in)
//      -> create the Entra app, grant tenant-wide consent for exactly those permissions
//   3. choose people from the firm's directory -> give them access
//   4. email each person their own setup code from the admin's Outlook (drafts first)
// The admin sign-in lives only in this process's memory and is never written to disk.
import { PublicClientApplication, LogLevel } from '@azure/msal-node';
import { graph, graphAll } from './graph.js';
import { scopesFor } from './scopes.js';
import { REDIRECT_URI } from './paths.js';
import { openBrowser } from './browser.js';
import { audit } from './audit.js';
import { makeSetupCode } from './config.js';

const GRAPH_CLI_APP = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';
const DEFAULT_ACCESS_ROLE = '00000000-0000-0000-0000-000000000000';
export const APP_NAME = 'Claude M365 Bridge';
const G = 'https://graph.microsoft.com/';
const ADMIN_SCOPES = [
  'Application.ReadWrite.All', // create the app
  'DelegatedPermissionGrant.ReadWrite.All', // approve it for the firm
  'AppRoleAssignment.ReadWrite.All', // give chosen people access
  'User.Read.All', // list the team
  'Mail.ReadWrite', // draft invitations in the admin's Outlook
  'Mail.Send', // send them when the admin says so
].map((s) => G + s);
const SESSION_MS = 2 * 60 * 60 * 1000;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const page = (title, text) =>
  `<!doctype html><meta charset="utf-8"><title>${title}</title><body style="font-family:-apple-system,'Segoe UI',system-ui,sans-serif;display:grid;place-items:center;height:100vh;margin:0;background:#f5f5f5;color:#242424"><div style="background:#fff;padding:32px 40px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.12);max-width:420px"><h1 style="font-size:20px;font-weight:600;margin:0 0 8px">${title}</h1><p style="margin:0;color:#616161;line-height:1.5">${text}</p></div>`;

// ---------------------------------------------------------------------------
// Admin session (memory only)

let session = null; // { pca, account, token, expires, me, org, tenantId }

export const adminSession = () => (session && Date.now() < session.started + SESSION_MS ? { me: session.me, org: session.org, tenantId: session.tenantId } : null);
export function endAdminSession() { session = null; }

async function adminToken() {
  if (!session || Date.now() > session.started + SESSION_MS) throw new Error('The admin sign-in has ended. Sign in again to continue.');
  if (session.pca && Date.now() > session.expires - 5 * 60 * 1000) {
    const r = await session.pca.acquireTokenSilent({ account: session.account, scopes: ADMIN_SCOPES });
    session.token = r.accessToken;
    session.expires = r.expiresOn?.getTime() || Date.now() + 3600e3;
  }
  return session.token;
}
const g = async (p, opts = {}) => graph(p, { ...opts, token: await adminToken() });
const gAll = async (p, max) => graphAll(p, max, { token: await adminToken() });

export async function signInAdmin(onUrl = () => {}) {
  let token, tenantId, pca = null, account = null, expires = Date.now() + 3600e3;
  if (process.env.M365_TEST_ADMIN_TOKEN) {
    token = process.env.M365_TEST_ADMIN_TOKEN;
  } else {
    pca = new PublicClientApplication({
      auth: { clientId: GRAPH_CLI_APP, authority: 'https://login.microsoftonline.com/organizations' },
      system: { loggerOptions: { loggerCallback() {}, piiLoggingEnabled: false, logLevel: LogLevel.Error } },
    });
    const res = await pca.acquireTokenInteractive({
      scopes: ADMIN_SCOPES,
      prompt: 'select_account',
      openBrowser: async (url) => { onUrl(url); openBrowser(url); },
      successTemplate: page('Signed in', 'Go back to the M365 Bridge tab. It is setting up your firm now.'),
      errorTemplate: page('Sign-in failed', 'Go back to the M365 Bridge tab and try again.'),
    });
    ({ accessToken: token, tenantId, account } = res);
    expires = res.expiresOn?.getTime() || expires;
  }
  session = { pca, account, token, expires, started: Date.now() };
  const org = (await g('/organization?$select=id,displayName')).value?.[0] || {};
  session.org = { id: org.id || tenantId, name: org.displayName || '' };
  session.tenantId = org.id || tenantId;
  session.me = await g('/me?$select=id,displayName,userPrincipalName,mail');
  return adminSession();
}

// ---------------------------------------------------------------------------
// The app

async function servicePrincipalFor(appId) {
  return (await g(`/servicePrincipals?$filter=appId eq '${appId}'&$select=id,appRoleAssignmentRequired`)).value?.[0] || null;
}

export async function setupApp({ cfg, firmName, onStep = () => {} }) {
  const graphSp = (await g(`/servicePrincipals?$filter=appId eq '${GRAPH_APP_ID}'&$select=id,oauth2PermissionScopes`)).value[0];
  const wanted = scopesFor(cfg);
  const byName = new Map(graphSp.oauth2PermissionScopes.map((s) => [s.value, s.id]));
  const granted = wanted.filter((s) => byName.has(s));
  const unknown = wanted.filter((s) => !byName.has(s));

  const appBody = {
    signInAudience: 'AzureADMyOrg',
    isFallbackPublicClient: true,
    publicClient: { redirectUris: [REDIRECT_URI] },
    requiredResourceAccess: [{ resourceAppId: GRAPH_APP_ID, resourceAccess: granted.map((s) => ({ id: byName.get(s), type: 'Scope' })) }],
    notes: 'Created by M365 Bridge. Lets each person use Microsoft 365 from Claude, signed in as themselves. No secrets.',
  };
  let app = (await g(`/applications?$filter=displayName eq '${APP_NAME}'&$select=id,appId`)).value?.[0];
  if (app) {
    await g(`/applications/${app.id}`, { method: 'PATCH', body: appBody });
    onStep(`Updated the "${APP_NAME}" app in Entra.`);
  } else {
    app = await g('/applications', { method: 'POST', body: { displayName: APP_NAME, ...appBody } });
    onStep(`Created the "${APP_NAME}" app in Entra.`);
  }

  // A brand-new app can take a few seconds to appear in the directory.
  let sp = await servicePrincipalFor(app.appId);
  for (let i = 0; !sp; i++) {
    try {
      sp = await g('/servicePrincipals', { method: 'POST', body: { appId: app.appId, appRoleAssignmentRequired: true, tags: ['WindowsAzureActiveDirectoryIntegratedApp'] } });
    } catch (e) {
      if (i >= 10) throw e;
      await sleep(3000);
    }
  }

  const scope = granted.join(' ');
  const grants = await gAll(`/oauth2PermissionGrants?$filter=clientId eq '${sp.id}'`, 50);
  const existing = grants.find((x) => x.consentType === 'AllPrincipals' && x.resourceId === graphSp.id);
  const token = await adminToken();
  if (existing) await graph(`/oauth2PermissionGrants/${existing.id}`, { method: 'PATCH', body: { scope }, token });
  else await graph('/oauth2PermissionGrants', { method: 'POST', body: { clientId: sp.id, consentType: 'AllPrincipals', resourceId: graphSp.id, scope }, token });
  onStep(`Approved ${granted.length} permissions for the firm.`);

  audit({ tool: 'firm_setup_auto', tenant: session.tenantId, app: app.appId });
  return {
    tenantId: session.tenantId,
    clientId: app.appId,
    firmName: firmName || session.org.name || '',
    admin: session.me.userPrincipalName,
    permissions: granted,
    unknown_permissions: unknown.length ? unknown : undefined,
  };
}

// ---------------------------------------------------------------------------
// People

/** Everyone in the firm's directory, with who already has access. */
export async function listPeople(appId) {
  const users = await gAll('/users?$select=id,displayName,mail,userPrincipalName,jobTitle,department,accountEnabled,userType,assignedLicenses&$top=999', 5000);
  const sp = appId ? await servicePrincipalFor(appId) : null;
  const assigned = sp ? new Set((await gAll(`/servicePrincipals/${sp.id}/appRoleAssignedTo?$top=999`, 5000).catch(() => [])).map((a) => a.principalId)) : new Set();
  return {
    restricted: sp ? sp.appRoleAssignmentRequired !== false : true,
    people: users
      .filter((u) => u.accountEnabled !== false && u.userType !== 'Guest')
      .map((u) => ({
        id: u.id,
        name: u.displayName || u.userPrincipalName,
        email: u.mail || u.userPrincipalName,
        upn: u.userPrincipalName,
        title: [u.jobTitle, u.department].filter(Boolean).join(', '),
        licensed: (u.assignedLicenses || []).length > 0,
        hasAccess: assigned.has(u.id),
        isAdmin: u.id === session?.me?.id,
      }))
      .sort((a, b) => Number(b.licensed) - Number(a.licensed) || a.name.localeCompare(b.name)),
  };
}

/** Give these people access. restricted = only assigned people can sign in. */
export async function grantAccess({ appId, ids, restricted = true }) {
  const sp = await servicePrincipalFor(appId);
  if (!sp) throw new Error('The M365 Bridge app is not in Entra yet. Run setup first.');
  await g(`/servicePrincipals/${sp.id}`, { method: 'PATCH', body: { appRoleAssignmentRequired: restricted } });
  const list = [...new Set([...(ids || []), ...(restricted && session?.me?.id ? [session.me.id] : [])])];
  const added = [];
  for (const principalId of list) {
    try {
      await g(`/servicePrincipals/${sp.id}/appRoleAssignedTo`, { method: 'POST', body: { principalId, resourceId: sp.id, appRoleId: DEFAULT_ACCESS_ROLE } });
      added.push(principalId);
    } catch (e) {
      if (!/already exists/i.test(e.message) && e.status !== 409) throw e;
    }
  }
  audit({ tool: 'grant_access', added: added.length, restricted });
  const chosen = new Set(ids || []);
  return { added: added.filter((id) => chosen.has(id)).length, total: chosen.size, restricted };
}

// ---------------------------------------------------------------------------
// Invitations

export const INSTALL_MAC = 'curl -fsSL https://raw.githubusercontent.com/irfad7/m365-bridge/main/install.sh | bash';
export const INSTALL_WIN = 'irm https://raw.githubusercontent.com/irfad7/m365-bridge/main/install.ps1 | iex';

export const DEFAULT_SUBJECT = 'Claude now works with our Microsoft 365';
export const DEFAULT_MESSAGE = `Hi {first_name},

I've connected Claude to our Microsoft 365. It signs in as you, so it only sees the files and email you can already open. Setup takes about two minutes.

1. Install M365 Bridge.
   Mac: open Terminal and paste
   ${INSTALL_MAC}
   Windows: open PowerShell and paste
   ${INSTALL_WIN}

2. A page opens in your browser. Paste your setup code:
   {setup_code}

3. Click "Sign in with Microsoft", then restart Claude.

You can also ask Claude Code: "Install M365 Bridge from https://github.com/irfad7/m365-bridge and set it up with this code: {setup_code}"

{admin_name}`;

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

/** Fill the template for one person. Returns { subject, html, text }. */
export function renderInvite({ subject = DEFAULT_SUBJECT, message = DEFAULT_MESSAGE, person, cfg, adminName }) {
  const code = makeSetupCode(cfg, { user: person.upn });
  const vars = { first_name: String(person.name || '').split(/\s+/)[0] || 'there', name: person.name, setup_code: code, firm: cfg.firmName || '', admin_name: adminName || '' };
  const fill = (s) => s.replace(/\{(\w+)\}/g, (m, k) => (k in vars ? vars[k] : m));
  const text = fill(message);
  const mono = 'font-family:Consolas,Menlo,monospace;font-size:12px;background:#f3f2f1;border:1px solid #e1dfdd;border-radius:4px;padding:6px 8px;display:inline-block;word-break:break-all';
  const html = `<div style="font-family:'Segoe UI',-apple-system,Helvetica,Arial,sans-serif;font-size:14px;line-height:1.5;color:#242424">${esc(text)
    .split(/\n{2,}/)
    .map((para) => `<p style="margin:0 0 12px">${para
      .split('\n')
      .map((line) => {
        const t = line.trim();
        if (t === code || t === INSTALL_MAC || t === INSTALL_WIN) return `<span style="${mono}">${esc(t)}</span>`;
        return line;
      })
      .join('<br>')}</p>`)
    .join('')}</div>`;
  return { subject: fill(subject), html, text, code };
}

/** Save one draft per person in the admin's Outlook. Nothing is sent. */
export async function draftInvites({ people, subject, message, cfg }) {
  const adminName = session?.me?.displayName || '';
  const drafts = [];
  for (const person of people) {
    const m = renderInvite({ subject, message, person, cfg, adminName });
    const d = await g('/me/messages', {
      method: 'POST',
      body: { subject: m.subject, body: { contentType: 'HTML', content: m.html }, toRecipients: [{ emailAddress: { address: person.email, name: person.name } }] },
    });
    drafts.push({ id: d.id, to: person.email, name: person.name, webLink: d.webLink, sent: false });
  }
  audit({ tool: 'invites_drafted', count: drafts.length });
  return drafts;
}

export async function sendInvites(drafts) {
  let sent = 0;
  const errors = [];
  for (const d of drafts) {
    if (d.sent) continue;
    try {
      await g(`/me/messages/${encodeURIComponent(d.id)}/send`, { method: 'POST' });
      d.sent = true;
      sent++;
    } catch (e) {
      errors.push(`${d.to}: ${e.message}`);
    }
  }
  audit({ tool: 'invites_sent', count: sent, errors: errors.length });
  return { sent, errors };
}

// ---------------------------------------------------------------------------
// Background job for step 2 (sign-in + app), polled by the page

export let job = null;

export function startProvision({ cfg, firmName, onDone, access, people }) {
  if (job?.running) return job;
  job = { running: true, steps: [], error: null, result: null, signInUrl: null, cancelled: false, started: new Date().toISOString() };
  const current = job;
  const onStep = (text) => { if (!current.cancelled) current.steps.push({ text, at: new Date().toISOString() }); };
  const onSignInUrl = (url) => { current.signInUrl = url; };
  current.promise = (async () => {
    // Claude's firm_setup passes access/people: do everything in one go.
    if (access) return provision({ cfg, firmName, access, people, onStep, onSignInUrl });
    onStep('Waiting for an admin to sign in to Microsoft…');
    const s = await signInAdmin(onSignInUrl);
    if (current.cancelled) return null;
    onStep(`Signed in as ${s.me.userPrincipalName} (${s.org.name || 'your organisation'}).`);
    return setupApp({ cfg, firmName, onStep });
  })()
    .then(async (r) => {
      if (current.cancelled || !r) return null;
      current.result = r;
      if (onDone) await onDone(r);
      return r;
    })
    .catch((e) => { if (!current.cancelled) current.error = e.message; return null; })
    .finally(() => { current.running = false; });
  return current;
}

/** Admin sign-in only (manage people later). */
export function startSignIn({ onDone } = {}) {
  if (job?.running) return job;
  job = { running: true, steps: [], error: null, result: null, signedIn: false, signInUrl: null, cancelled: false, started: new Date().toISOString() };
  const current = job;
  current.steps.push({ text: 'Waiting for an admin to sign in to Microsoft…', at: current.started });
  current.promise = signInAdmin((url) => { current.signInUrl = url; })
    .then(async (s) => {
      if (current.cancelled) return null;
      current.signedIn = true;
      current.steps.push({ text: `Signed in as ${s.me.userPrincipalName}.`, at: new Date().toISOString() });
      if (onDone) await onDone(s);
      return s;
    })
    .catch((e) => { if (!current.cancelled) current.error = e.message; return null; })
    .finally(() => { current.running = false; });
  return current;
}

/** Stop waiting for the admin. A sign-in finished after this is ignored. */
export function cancelProvision() {
  if (!job?.running) return false;
  job.cancelled = true;
  job.running = false;
  job.error = 'Cancelled. Nothing was changed in Microsoft 365.';
  return true;
}

/** Everything in one call (used by Claude's firm_setup tool): sign in, set up, optionally limit to named people. */
export async function provision({ firmName, access = 'everyone', people = [], cfg, onStep = () => {}, onSignInUrl }) {
  onStep('Waiting for an admin to sign in to Microsoft…');
  const s = await signInAdmin(onSignInUrl);
  onStep(`Signed in as ${s.me.userPrincipalName} (${s.org.name || 'your organisation'}).`);
  const r = await setupApp({ cfg, firmName, onStep });
  const restricted = access === 'selected';
  const assigned = [];
  const notFound = [];
  const ids = [];
  if (restricted) {
    for (const who of [...new Set(people.map((p) => String(p).trim().toLowerCase()).filter(Boolean))]) {
      try {
        const u = await g(`/users/${encodeURIComponent(who)}?$select=id`);
        ids.push(u.id);
        assigned.push(who);
      } catch { notFound.push(who); }
    }
  }
  await grantAccess({ appId: r.clientId, ids, restricted });
  if (restricted) assigned.unshift(s.me.userPrincipalName.toLowerCase());
  onStep(restricted ? `Gave access to ${assigned.length} people.` : 'Everyone in the firm can sign in.');
  return { ...r, access, assigned: restricted ? [...new Set(assigned)] : undefined, not_found: notFound.length ? notFound : undefined };
}
