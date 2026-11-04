// One-click firm setup. A Microsoft 365 admin signs in once (Microsoft's own
// "Microsoft Graph Command Line Tools" sign-in, the same one Connect-MgGraph uses), and
// this creates the firm's Entra app, grants tenant-wide consent for exactly the chosen
// permissions, and optionally limits it to named people. The admin token lives only in
// memory for these few calls and is never saved.
import { PublicClientApplication, LogLevel } from '@azure/msal-node';
import { graph, graphAll } from './graph.js';
import { scopesFor } from './scopes.js';
import { REDIRECT_URI } from './paths.js';
import { openBrowser } from './browser.js';
import { audit } from './audit.js';

const GRAPH_CLI_APP = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';
const DEFAULT_ACCESS_ROLE = '00000000-0000-0000-0000-000000000000';
export const APP_NAME = 'Claude M365 Bridge';
const ADMIN_SCOPES = [
  'https://graph.microsoft.com/Application.ReadWrite.All',
  'https://graph.microsoft.com/DelegatedPermissionGrant.ReadWrite.All',
  'https://graph.microsoft.com/AppRoleAssignment.ReadWrite.All',
  'https://graph.microsoft.com/User.ReadBasic.All',
];

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const page = (title, text) =>
  `<!doctype html><meta charset="utf-8"><title>${title}</title><body style="font-family:-apple-system,'Segoe UI',system-ui,sans-serif;display:grid;place-items:center;height:100vh;margin:0;background:#f5f5f5;color:#242424"><div style="background:#fff;padding:32px 40px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.12);max-width:420px"><h1 style="font-size:20px;font-weight:600;margin:0 0 8px">${title}</h1><p style="margin:0;color:#616161;line-height:1.5">${text}</p></div>`;

async function adminToken() {
  if (process.env.M365_TEST_ADMIN_TOKEN) return { token: process.env.M365_TEST_ADMIN_TOKEN, tenantId: process.env.M365_TEST_TENANT || '' };
  const pca = new PublicClientApplication({
    auth: { clientId: GRAPH_CLI_APP, authority: 'https://login.microsoftonline.com/organizations' },
    system: { loggerOptions: { loggerCallback() {}, piiLoggingEnabled: false, logLevel: LogLevel.Error } },
  });
  const res = await pca.acquireTokenInteractive({
    scopes: ADMIN_SCOPES,
    prompt: 'select_account',
    openBrowser: async (url) => openBrowser(url),
    successTemplate: page('Signed in', 'Setting up your firm now. Go back to the M365 Bridge tab to watch it finish.'),
    errorTemplate: page('Sign-in failed', 'Go back to the M365 Bridge tab and try again.'),
  });
  return { token: res.accessToken, tenantId: res.tenantId };
}

/** Current setup job, for the dashboard to poll. */
export let job = null;

export function startProvision(options) {
  if (job?.running) return job;
  job = { running: true, steps: [], error: null, result: null, started: new Date().toISOString() };
  const current = job;
  current.promise = provision({ ...options, onStep: (text) => current.steps.push({ text, at: new Date().toISOString() }) })
    .then((r) => { current.result = r; return r; })
    .catch((e) => { current.error = e.message; })
    .finally(() => { current.running = false; });
  return current;
}

export async function provision({ firmName, access = 'everyone', people = [], cfg, onStep = () => {} }) {
  onStep('Waiting for an admin to sign in to Microsoft…');
  const { token, tenantId: tid } = await adminToken();
  const g = (p, opts = {}) => graph(p, { ...opts, token });

  const org = (await g('/organization?$select=id,displayName')).value?.[0] || {};
  const tenantId = org.id || tid;
  const me = await g('/me?$select=id,displayName,userPrincipalName');
  onStep(`Signed in as ${me.userPrincipalName} (${org.displayName || 'your organisation'}).`);

  // Map the chosen permission names to Microsoft's ids.
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
    onStep(`Updated the existing "${APP_NAME}" app.`);
  } else {
    app = await g('/applications', { method: 'POST', body: { displayName: APP_NAME, ...appBody } });
    onStep(`Created the "${APP_NAME}" app in Entra.`);
  }

  // The app's identity in this tenant. A brand-new app can take a few seconds to appear.
  const assignmentRequired = access === 'selected';
  let sp = (await g(`/servicePrincipals?$filter=appId eq '${app.appId}'&$select=id`)).value?.[0];
  for (let i = 0; !sp; i++) {
    try {
      sp = await g('/servicePrincipals', { method: 'POST', body: { appId: app.appId, appRoleAssignmentRequired: assignmentRequired, tags: ['WindowsAzureActiveDirectoryIntegratedApp'] } });
    } catch (e) {
      if (i >= 10) throw e;
      await sleep(3000);
    }
  }
  await g(`/servicePrincipals/${sp.id}`, { method: 'PATCH', body: { appRoleAssignmentRequired: assignmentRequired } });

  // Tenant-wide consent, so nobody sees a permissions prompt.
  const scope = granted.join(' ');
  const grants = await graphAll(`/oauth2PermissionGrants?$filter=clientId eq '${sp.id}'`, 50).catch(() => []);
  const existing = grants.find((x) => x.consentType === 'AllPrincipals' && x.resourceId === graphSp.id);
  if (existing) await g(`/oauth2PermissionGrants/${existing.id}`, { method: 'PATCH', body: { scope } });
  else await g('/oauth2PermissionGrants', { method: 'POST', body: { clientId: sp.id, consentType: 'AllPrincipals', resourceId: graphSp.id, scope } });
  onStep(`Approved ${granted.length} permissions for the whole firm.`);

  const assigned = [];
  const notFound = [];
  if (assignmentRequired) {
    const list = [...new Set([me.userPrincipalName, ...people].map((p) => String(p).trim().toLowerCase()).filter(Boolean))];
    for (const who of list) {
      let user;
      try { user = await g(`/users/${encodeURIComponent(who)}?$select=id,displayName`); } catch { notFound.push(who); continue; }
      try {
        await g(`/servicePrincipals/${sp.id}/appRoleAssignedTo`, { method: 'POST', body: { principalId: user.id, resourceId: sp.id, appRoleId: DEFAULT_ACCESS_ROLE } });
      } catch (e) {
        if (!/already exists/i.test(e.message) && e.status !== 409) throw e;
      }
      assigned.push(who);
    }
    onStep(`Gave access to ${assigned.length} ${assigned.length === 1 ? 'person' : 'people'}.${notFound.length ? ` Not found: ${notFound.join(', ')}.` : ''}`);
  } else {
    onStep('Everyone in the firm can sign in. Each person only sees their own files and mail.');
  }

  audit({ tool: 'firm_setup_auto', tenant: tenantId, app: app.appId, access, people: assigned.length });
  return {
    tenantId,
    clientId: app.appId,
    firmName: firmName || org.displayName || '',
    admin: me.userPrincipalName,
    permissions: granted,
    unknown_permissions: unknown.length ? unknown : undefined,
    access,
    assigned: assignmentRequired ? assigned : undefined,
    not_found: notFound.length ? notFound : undefined,
  };
}
