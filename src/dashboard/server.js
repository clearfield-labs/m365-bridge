// Local admin dashboard. Binds to 127.0.0.1 only, rejects other Host headers (DNS rebinding)
// and cross-origin POSTs (CSRF). Never renders tokens.
import http from 'node:http';
import { spawn } from 'node:child_process';
import { PORT, DASHBOARD_URL, WORKSPACES_ROOT } from '../paths.js';
import { loadConfig, updateConfig, isConfigured, validateIds, makeSetupCode, parseSetupCode } from '../config.js';
import { currentAccount, startBrowserSignIn, finishBrowserSignIn, signOut } from '../auth.js';
import { keyStorage } from '../vault.js';
import { recentAudit, audit } from '../audit.js';
import { graph } from '../graph.js';
import * as ws from '../workspace.js';
import { registrationStatus, registerClaudeCode, registerClaudeDesktop } from '../register.js';
import { renderPage, renderMessage } from './page.js';
import { PACKS, parseExtraScopes, scopesFor } from '../scopes.js';

const ALLOWED_HOSTS = new Set([`localhost:${PORT}`, `127.0.0.1:${PORT}`, `[::1]:${PORT}`]);
let flash = null; // one-shot message shown after a redirect

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', (c) => {
      data += c;
      if (data.length > 1e6) req.destroy();
    });
    req.on('end', () => resolve(Object.fromEntries(new URLSearchParams(data))));
    req.on('error', reject);
  });
}

function send(res, status, html, headers = {}) {
  res.writeHead(status, {
    'Content-Type': 'text/html; charset=utf-8',
    'Cache-Control': 'no-store',
    'X-Frame-Options': 'DENY',
    'Referrer-Policy': 'no-referrer',
    'Content-Security-Policy': "default-src 'self'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; img-src 'self' data:; form-action 'self' https://login.microsoftonline.com",
    ...headers,
  });
  res.end(html);
}
const redirect = (res, to) => { res.writeHead(303, { Location: to, 'Cache-Control': 'no-store' }); res.end(); };
const go = (res, kind, text, anchor = '') => { flash = { kind, text }; redirect(res, `/${anchor}`); };

async function state() {
  const cfg = loadConfig();
  return {
    cfg,
    configured: isConfigured(cfg),
    account: await currentAccount(),
    keyStorage: keyStorage(),
    setupCode: isConfigured(cfg) ? makeSetupCode(cfg) : null,
    registration: registrationStatus(),
    workspaces: ws.listWorkspaces(),
    workspacesRoot: WORKSPACES_ROOT,
    audit: recentAudit(40),
    port: PORT,
  };
}

async function handle(req, res) {
  if (!ALLOWED_HOSTS.has(req.headers.host)) return send(res, 403, 'Forbidden host');
  const url = new URL(req.url, DASHBOARD_URL);

  if (req.method === 'POST') {
    const origin = req.headers.origin;
    if (origin && !ALLOWED_HOSTS.has(origin.replace(/^https?:\/\//, ''))) return send(res, 403, 'Forbidden origin');
    const body = await readBody(req);
    try {
      return await post(url.pathname, body, res);
    } catch (e) {
      return go(res, 'error', e.message);
    }
  }

  switch (url.pathname) {
    case '/': {
      const s = await state();
      const page = renderPage(s, flash);
      flash = null;
      return send(res, 200, page);
    }
    case '/healthz':
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ app: 'm365-bridge' }));
    case '/auth/start':
      if (!isConfigured()) return go(res, 'error', 'Enter the firm Tenant ID and Client ID first.');
      return redirect(res, await startBrowserSignIn());
    case '/auth/callback': {
      const p = url.searchParams;
      if (p.get('admin_consent') || p.get('state') === 'adminconsent') {
        audit({ tool: 'admin_consent', result: p.get('admin_consent') });
        return (p.get('admin_consent') || '').toLowerCase() === 'true'
          ? go(res, 'ok', 'Admin consent granted for the whole firm. Each person can now sign in without a permission prompt.', '#firm')
          : go(res, 'error', `Admin consent was not granted: ${p.get('error_description') || 'cancelled'}`, '#firm');
      }
      if (p.get('error')) return go(res, 'error', `Microsoft sign-in failed: ${p.get('error_description') || p.get('error')}`, '#account');
      try {
        const account = await finishBrowserSignIn(p.get('code'), p.get('state'));
        audit({ tool: 'sign_in', user: account?.username });
        return send(res, 200, renderMessage('Signed in', `You are signed in as <b>${escape(account?.username)}</b>. You can close this tab and go back to Claude.`));
      } catch (e) {
        return go(res, 'error', e.message, '#account');
      }
    }
    default:
      return send(res, 404, renderMessage('Not found', 'Nothing here.'));
  }
}

async function post(route, body, res) {
  switch (route) {
    case '/firm': {
      let patch;
      if (body.setupCode?.trim()) {
        try { patch = parseSetupCode(body.setupCode); } catch { throw new Error('That setup code is not valid. Copy it again from your admin.'); }
      } else {
        patch = {
          tenantId: body.tenantId?.trim(),
          clientId: body.clientId?.trim(),
          firmName: body.firmName?.trim() || '',
          allowWrite: body.allowWrite === 'on',
          allowDelete: body.allowDelete === 'on',
          packs: Object.keys(PACKS).filter((k) => k !== 'files' && body[`pack_${k}`] === 'on'),
          extraScopes: parseExtraScopes(body.extraScopes).join(' '),
        };
      }
      const errors = validateIds(patch.tenantId, patch.clientId);
      if (errors.length) throw new Error(errors.join(' '));
      const before = loadConfig();
      const scopesChanged = scopesFor(before).join(' ') !== scopesFor({ ...before, ...patch }).join(' ');
      if (before.tenantId && (before.tenantId !== patch.tenantId || before.clientId !== patch.clientId || scopesChanged)) await signOut();
      updateConfig(patch);
      audit({ tool: 'firm_setup', tenant: patch.tenantId });
      return go(res, 'ok', 'Saved and encrypted. Next: sign in with your Microsoft account.', '#account');
    }
    case '/signout':
      await signOut();
      audit({ tool: 'sign_out' });
      return go(res, 'ok', 'Signed out. Your tokens were deleted from this computer.', '#account');
    case '/test': {
      const me = await graph('/me?$select=displayName,mail,userPrincipalName');
      const drive = await graph('/me/drive?$select=webUrl,quota');
      const used = drive.quota?.used ? ` (${(drive.quota.used / 1e9).toFixed(1)} GB used)` : '';
      return go(res, 'ok', `Connected as ${me.displayName} (${me.mail || me.userPrincipalName}). OneDrive reachable${used}.`, '#account');
    }
    case '/register': {
      const msgs = [];
      if (body.target === 'code' || body.target === 'both') msgs.push(registerClaudeCode());
      if (body.target === 'desktop' || body.target === 'both') msgs.push(registerClaudeDesktop());
      audit({ tool: 'register', target: body.target });
      return go(res, 'ok', msgs.join(' '), '#claude');
    }
    case '/workspace/add': {
      const r = await ws.pull({ source: body.source?.trim(), name: body.name?.trim() || undefined });
      return go(res, r.conflicts.length ? 'warn' : 'ok', `Pulled ${r.downloaded.length} files into ${r.workspace}.${r.skipped.length ? ` Skipped ${r.skipped.length}.` : ''} Open that folder in Claude Code.`, '#workspaces');
    }
    case '/workspace/pull': {
      const r = await ws.pull({ localPath: body.localPath });
      const n = r.downloaded.length + r.updated.length;
      return go(res, r.conflicts.length ? 'warn' : 'ok', `${n} file${n === 1 ? '' : 's'} updated.${r.conflicts.length ? ` ${r.conflicts.length} conflict(s): SharePoint copies saved next to your files.` : ''}`, '#workspaces');
    }
    case '/workspace/open':
      openPath(body.localPath);
      return go(res, 'ok', `Opened ${body.localPath}`, '#workspaces');
    case '/workspace/forget':
      ws.forgetWorkspace(body.localPath);
      return go(res, 'ok', 'Removed from the list. The local folder and SharePoint were not touched.', '#workspaces');
    default:
      return send(res, 404, renderMessage('Not found', 'Nothing here.'));
  }
}

export function escape(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

export function openBrowser(url) {
  const [cmd, args] =
    process.platform === 'darwin' ? ['open', [url]]
    : process.platform === 'win32' ? ['cmd', ['/c', 'start', '', url]]
    : ['xdg-open', [url]];
  try { spawn(cmd, args, { stdio: 'ignore', detached: true, windowsHide: true }).unref(); } catch { /* headless */ }
}
const openPath = openBrowser;

let server = null;

/** Start the dashboard in this process unless one is already answering on the port. */
export async function ensureDashboard() {
  if (server) return DASHBOARD_URL;
  try {
    const r = await fetch(`${DASHBOARD_URL}healthz`, { signal: AbortSignal.timeout(800) });
    if (r.ok && (await r.json()).app === 'm365-bridge') return DASHBOARD_URL;
  } catch { /* not running */ }
  return startDashboard();
}

export function startDashboard() {
  return new Promise((resolve, reject) => {
    server = http.createServer((req, res) => {
      handle(req, res).catch((e) => send(res, 500, renderMessage('Something went wrong', escape(e.message))));
    });
    server.on('error', (e) => {
      server = null;
      reject(e.code === 'EADDRINUSE' ? new Error(`Port ${PORT} is busy. Set M365_BRIDGE_PORT to another port (and add that redirect URI in Entra).`) : e);
    });
    server.listen(PORT, '127.0.0.1', () => {
      // "localhost" can resolve to IPv6 first; answer there too so the Microsoft redirect always lands.
      const v6 = http.createServer(server.listeners('request')[0]);
      v6.on('error', () => {});
      v6.listen(PORT, '::1');
      resolve(DASHBOARD_URL);
    });
  });
}
