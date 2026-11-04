// Local admin dashboard. Binds to 127.0.0.1 only, rejects other Host headers (DNS rebinding)
// and cross-origin POSTs (CSRF). Never renders tokens.
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { PORT, DASHBOARD_URL, WORKSPACES_ROOT, PACKAGE_ROOT } from '../paths.js';
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
    'Content-Security-Policy': "default-src 'self'; style-src 'self' 'unsafe-inline'; script-src 'self'; img-src 'self' data:; font-src 'self'; connect-src 'self'; form-action 'self' https://login.microsoftonline.com",
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
          ? go(res, 'ok', 'Admin consent granted for the whole firm. Each person can now sign in without a permission prompt.', '#admin')
          : go(res, 'error', `Admin consent was not granted: ${p.get('error_description') || 'cancelled'}`, '#admin');
      }
      if (p.get('error')) return go(res, 'error', `Microsoft sign-in failed: ${p.get('error_description') || p.get('error')}`, '#overview');
      try {
        const account = await finishBrowserSignIn(p.get('code'), p.get('state'));
        audit({ tool: 'sign_in', user: account?.username });
        return go(res, 'ok', `Signed in as ${account?.username}. You can go back to Claude.`, '#overview');
      } catch (e) {
        return go(res, 'error', e.message, '#overview');
      }
    }
    case '/api/workspace-status': {
      const wsPath = url.searchParams.get('path');
      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Cache-Control', 'no-store');
      if (!ws.listWorkspaces().some((w) => w.localPath === wsPath)) {
        res.writeHead(404);
        return res.end(JSON.stringify({ error: 'Unknown workspace' }));
      }
      try {
        const st = await ws.status({ localPath: wsPath, checkRemote: !!(await currentAccount()) });
        res.writeHead(200);
        return res.end(JSON.stringify({ data: st }));
      } catch (e) {
        res.writeHead(200);
        return res.end(JSON.stringify({ error: e.message }));
      }
    }
    default:
      if (url.pathname.startsWith('/assets/fonts/')) return sendFont(res, path.basename(url.pathname));
      if (url.pathname === '/assets/app.css' || url.pathname === '/assets/app.js') return sendStatic(res, path.basename(url.pathname));
      return send(res, 404, renderMessage('Not found', 'Nothing here.'));
  }
}

const FONTS = {
  'serif.woff2': '@fontsource/instrument-serif/files/instrument-serif-latin-400-normal.woff2',
  'serif-italic.woff2': '@fontsource/instrument-serif/files/instrument-serif-latin-400-italic.woff2',
  'sans.woff2': '@fontsource-variable/hanken-grotesk/files/hanken-grotesk-latin-wght-normal.woff2',
  'mono.woff2': '@fontsource/ibm-plex-mono/files/ibm-plex-mono-latin-400-normal.woff2',
};
function sendFont(res, name) {
  const rel = FONTS[name];
  const file = rel && path.join(PACKAGE_ROOT, 'node_modules', rel);
  if (!file || !fs.existsSync(file)) {
    res.writeHead(404);
    return res.end();
  }
  res.writeHead(200, { 'Content-Type': 'font/woff2', 'Cache-Control': 'public, max-age=31536000, immutable' });
  fs.createReadStream(file).pipe(res);
}

const STATIC_TYPES = { '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8' };
function sendStatic(res, name) {
  const file = path.join(PACKAGE_ROOT, 'src', 'dashboard', 'static', name);
  res.writeHead(200, { 'Content-Type': STATIC_TYPES[path.extname(name)], 'Cache-Control': 'no-cache', 'X-Content-Type-Options': 'nosniff' });
  fs.createReadStream(file).pipe(res);
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
      return go(res, 'ok', 'Saved and encrypted. Next: sign in with your Microsoft account.', '#overview');
    }
    case '/signout':
      await signOut();
      audit({ tool: 'sign_out' });
      return go(res, 'ok', 'Signed out. Your tokens were deleted from this computer.', '#overview');
    case '/test': {
      const me = await graph('/me?$select=displayName,mail,userPrincipalName');
      const drive = await graph('/me/drive?$select=webUrl,quota');
      const used = drive.quota?.used ? ` (${(drive.quota.used / 1e9).toFixed(1)} GB used)` : '';
      return go(res, 'ok', `Connected as ${me.displayName} (${me.mail || me.userPrincipalName}). OneDrive reachable${used}.`, '#overview');
    }
    case '/register': {
      const msgs = [];
      if (body.target === 'code' || body.target === 'both') msgs.push(registerClaudeCode());
      if (body.target === 'desktop' || body.target === 'both') msgs.push(registerClaudeDesktop());
      audit({ tool: 'register', target: body.target });
      return go(res, 'ok', msgs.join(' '), '#overview');
    }
    case '/workspace/add': {
      const r = await ws.pull({ source: body.source?.trim(), name: body.name?.trim() || undefined });
      return go(res, r.conflicts.length ? 'warn' : 'ok', `Pulled ${r.downloaded.length} files into ${r.workspace}.${r.skipped.length ? ` Skipped ${r.skipped.length}.` : ''} Open that folder in Claude Code.`, '#matters');
    }
    case '/workspace/pull': {
      const r = await ws.pull({ localPath: body.localPath });
      const n = r.downloaded.length + r.updated.length;
      return go(res, r.conflicts.length ? 'warn' : 'ok', `${n} file${n === 1 ? '' : 's'} updated.${r.conflicts.length ? ` ${r.conflicts.length} conflict(s): SharePoint copies saved next to your files.` : ''}`, '#matters');
    }
    case '/workspace/open':
      openPath(body.localPath);
      return go(res, 'ok', `Opened ${body.localPath}`, '#matters');
    case '/workspace/forget':
      ws.forgetWorkspace(body.localPath);
      return go(res, 'ok', 'Removed from the list. The local folder and SharePoint were not touched.', '#matters');
    case '/workspace/push': {
      const r = await ws.push({ localPath: body.localPath });
      const n = r.uploaded.length + r.created.length;
      if (r.conflicts.length) return go(res, 'warn', `${n} uploaded. ${r.conflicts.length} not uploaded because a colleague changed them in SharePoint. Pull latest, merge, then upload again.`, '#matters');
      if (r.errors.length) return go(res, 'error', `${n} uploaded. ${r.errors[0].file}: ${r.errors[0].error}`, '#matters');
      return go(res, 'ok', n ? `Uploaded ${n} file${n === 1 ? '' : 's'} to SharePoint as new versions.` : 'Nothing to upload. Everything matches SharePoint.', '#matters');
    }
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
