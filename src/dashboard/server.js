// The background service: the setup page for people, and the MCP endpoint for Claude.
// Listens on this computer only (127.0.0.1 and ::1). Rejects other Host headers (DNS
// rebinding) and cross-site requests (CSRF). /mcp also needs the key Claude was given.
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { PORT, DASHBOARD_URL, WORKSPACES_ROOT, PACKAGE_ROOT } from '../paths.js';
import { loadConfig, updateConfig, isConfigured } from '../config.js';
import { currentAccount, startBrowserSignIn, finishBrowserSignIn, signOut } from '../auth.js';
import { keyStorage } from '../vault.js';
import { recentAudit, audit } from '../audit.js';
import { graph } from '../graph.js';
import * as ws from '../workspace.js';
import { registrationStatus, registerClaudeCode, registerClaudeDesktop } from '../register.js';
import { serviceStatus, installService, uninstallService } from '../service.js';
import { applyFirm, applySetupCode, currentSetupCode } from '../setup.js';
import * as provision from '../provision.js';
import { startProvision } from '../provision.js';
import { openBrowser } from '../browser.js';
import { renderPage, renderMessage } from './page.js';
import { PACKS, parseExtraScopes, adminConsentUrl } from '../scopes.js';

export { openBrowser };
const VERSION = JSON.parse(fs.readFileSync(path.join(PACKAGE_ROOT, 'package.json'), 'utf8')).version;
const ALLOWED_HOSTS = new Set([`localhost:${PORT}`, `127.0.0.1:${PORT}`, `[::1]:${PORT}`]);
let flash = null; // one-shot message shown after a redirect

function readRaw(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', (c) => {
      size += c.length;
      if (size > 4e6) req.destroy();
      else chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

const SECURITY_HEADERS = {
  'Cache-Control': 'no-store',
  'X-Frame-Options': 'DENY',
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'same-origin', // no-referrer makes browsers send "Origin: null" on our own forms
  'Content-Security-Policy': "default-src 'self'; style-src 'self' 'unsafe-inline'; script-src 'self'; img-src 'self' data:; connect-src 'self'; form-action 'self' https://login.microsoftonline.com",
};
function send(res, status, html) {
  res.writeHead(status, { 'Content-Type': 'text/html; charset=utf-8', ...SECURITY_HEADERS });
  res.end(html);
}
function sendJson(res, status, data) {
  res.writeHead(status, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' });
  res.end(JSON.stringify(data));
}
const redirect = (res, to) => { res.writeHead(303, { Location: to, 'Cache-Control': 'no-store' }); res.end(); };
const go = (res, kind, text, to = '/') => { flash = { kind, text }; redirect(res, to); };

async function state() {
  const cfg = loadConfig();
  return {
    cfg,
    version: VERSION,
    configured: isConfigured(cfg),
    account: await currentAccount(),
    keyStorage: keyStorage(),
    setupCode: currentSetupCode(),
    consentUrl: isConfigured(cfg) ? adminConsentUrl(cfg, `${DASHBOARD_URL}auth/callback`) : null,
    registration: registrationStatus(),
    service: serviceStatus(),
    workspaces: ws.listWorkspaces(),
    held: ws.heldFiles(),
    workspacesRoot: WORKSPACES_ROOT,
    audit: recentAudit(12),
    job: provision.job ? { running: provision.job.running, steps: provision.job.steps, error: provision.job.error, done: !!provision.job.result, signInUrl: provision.job.signInUrl, options: provision.job.options } : null,
  };
}

/**
 * True when a request comes from some other web page (CSRF). Browsers mark their requests with
 * Sec-Fetch-Site and Origin; some send "Origin: null" for same-page forms, so trust Sec-Fetch-Site
 * there. Programs like Claude Code send neither and are checked by their key instead.
 */
export function isCrossSite(req) {
  const origin = req.headers.origin;
  const site = req.headers['sec-fetch-site'];
  if (site && site !== 'same-origin' && site !== 'none') return true;
  if (origin === 'null') return site !== 'same-origin';
  if (origin && !ALLOWED_HOSTS.has(origin.replace(/^https?:\/\//, ''))) return true;
  return false;
}

function bearerOk(req) {
  const token = loadConfig().mcpToken;
  const got = /^Bearer (.+)$/.exec(req.headers.authorization || '')?.[1] || '';
  if (!token || got.length !== token.length) return false;
  return crypto.timingSafeEqual(Buffer.from(got), Buffer.from(token));
}

async function handle(req, res) {
  if (!ALLOWED_HOSTS.has(req.headers.host)) return send(res, 403, 'Forbidden host');
  const url = new URL(req.url, DASHBOARD_URL);
  const foreignOrigin = isCrossSite(req);

  // Claude's connection. Browsers never get in: no web page can set the key, and any
  // cross-site Origin is refused before the key is even checked.
  if (url.pathname === '/mcp') {
    if (foreignOrigin) return sendJson(res, 403, { error: 'Forbidden origin' });
    if (!bearerOk(req)) return sendJson(res, 401, { error: 'Claude is not connected to this computer\'s M365 Bridge. Open the M365 Bridge page and click Connect Claude.' });
    const raw = req.method === 'POST' ? await readRaw(req) : '';
    let body;
    try { body = raw ? JSON.parse(raw) : undefined; } catch { return sendJson(res, 400, { error: 'Bad JSON' }); }
    const { handleMcpHttp } = await import('../mcp.js');
    return handleMcpHttp(req, res, body);
  }

  if (req.method === 'POST') {
    if (foreignOrigin) return send(res, 403, 'Forbidden origin');
    if (url.pathname === '/api/shutdown') {
      if (!bearerOk(req)) return sendJson(res, 401, { error: 'no' });
      sendJson(res, 200, { ok: true });
      return setTimeout(() => process.exit(0), 100);
    }
    const body = Object.fromEntries(new URLSearchParams(await readRaw(req)));
    const back = body.back === 'admin' ? '/admin' : '/';
    try {
      return await post(url.pathname, body, res, back);
    } catch (e) {
      return go(res, 'error', e.message, back);
    }
  }

  switch (url.pathname) {
    case '/':
    case '/admin': {
      const s = await state();
      const html = renderPage(s, flash, url.pathname === '/admin' ? 'admin' : 'home');
      flash = null;
      return send(res, 200, html);
    }
    case '/healthz':
      return sendJson(res, 200, { app: 'm365-bridge', version: VERSION, pid: process.pid });
    case '/auth/start':
      if (!isConfigured()) return go(res, 'error', 'Paste the setup code from your admin first.');
      return redirect(res, await startBrowserSignIn());
    case '/auth/callback': {
      const p = url.searchParams;
      if (p.get('admin_consent') || p.get('state') === 'adminconsent') {
        audit({ tool: 'admin_consent', result: p.get('admin_consent') });
        return (p.get('admin_consent') || '').toLowerCase() === 'true'
          ? go(res, 'ok', 'Approved for the whole firm. Nobody will see a permissions prompt.', '/admin')
          : go(res, 'error', `Not approved: ${p.get('error_description') || 'cancelled'}`, '/admin');
      }
      if (p.get('error')) return go(res, 'error', `Microsoft sign-in failed: ${p.get('error_description') || p.get('error')}`);
      try {
        const account = await finishBrowserSignIn(p.get('code'), p.get('state'));
        audit({ tool: 'sign_in', user: account?.username });
        return go(res, 'ok', `Signed in as ${account?.username}. You can go back to Claude.`);
      } catch (e) {
        return go(res, 'error', e.message);
      }
    }
    case '/api/workspace-status': {
      const wsPath = url.searchParams.get('path');
      if (!ws.listWorkspaces().some((w) => w.localPath === wsPath)) return sendJson(res, 404, { error: 'Unknown folder' });
      try {
        return sendJson(res, 200, { data: await ws.status({ localPath: wsPath, checkRemote: !!(await currentAccount()) }) });
      } catch (e) {
        return sendJson(res, 200, { error: e.message });
      }
    }
    case '/api/setup-job': {
      const j = provision.job;
      return sendJson(res, 200, { data: j ? { running: j.running, steps: j.steps.map((x) => x.text), error: j.error, done: !!j.result, signInUrl: j.signInUrl } : null });
    }
    case '/assets/app.css':
    case '/assets/app.js':
      return sendStatic(res, path.basename(url.pathname));
    default:
      return send(res, 404, renderMessage('Not found', 'Nothing here.'));
  }
}

const STATIC_TYPES = { '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8' };
function sendStatic(res, name) {
  const file = path.join(PACKAGE_ROOT, 'src', 'dashboard', 'static', name);
  res.writeHead(200, { 'Content-Type': STATIC_TYPES[path.extname(name)], 'Cache-Control': 'no-cache', 'X-Content-Type-Options': 'nosniff' });
  fs.createReadStream(file).pipe(res);
}

const plural = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;

async function post(route, body, res, back) {
  switch (route) {
    case '/setup-code': {
      const cfg = await applySetupCode(body.setupCode);
      return go(res, 'ok', `Connected to ${cfg.firmName || 'your firm'}. Next, sign in with Microsoft.`);
    }
    case '/firm': {
      await applyFirm({ tenantId: body.tenantId?.trim(), clientId: body.clientId?.trim(), firmName: body.firmName?.trim() || loadConfig().firmName });
      return go(res, 'ok', 'Saved. Share the setup code below with your team.', '/admin');
    }
    case '/settings': {
      const cfg = loadConfig();
      const patch = {
        firmName: body.firmName?.trim() || cfg.firmName,
        allowWrite: body.allowWrite === 'on',
        allowDelete: body.allowDelete === 'on',
        packs: Object.keys(PACKS).filter((k) => k !== 'files' && body[`pack_${k}`] === 'on'),
        extraScopes: parseExtraScopes(body.extraScopes).join(' '),
      };
      // Before setup there are no IDs yet: just remember the choices; setup asks Microsoft for exactly these.
      if (!isConfigured(cfg)) {
        updateConfig(patch);
        return go(res, 'ok', 'Saved. Firm setup will ask Microsoft for exactly these permissions.', '/admin');
      }
      await applyFirm({ tenantId: cfg.tenantId, clientId: cfg.clientId, ...patch });
      return go(res, 'ok', 'Settings saved. If you changed permissions, run setup again so Microsoft allows them, then send the team the new setup code.', '/admin');
    }
    case '/admin/auto': {
      const people = String(body.people || '').split(/[\s,;]+/).filter((x) => x.includes('@'));
      if (body.access === 'selected' && !people.length) throw new Error('List at least one email address, or choose "Everyone in the firm".');
      const j = startProvision({ firmName: body.firmName?.trim(), access: body.access === 'selected' ? 'selected' : 'everyone', people, cfg: loadConfig() });
      j.promise.then((r) => r && !j.cancelled && applyFirm({ tenantId: r.tenantId, clientId: r.clientId, firmName: r.firmName })).catch(() => {});
      return redirect(res, '/admin#setup');
    }
    case '/admin/cancel':
      provision.cancelProvision();
      return redirect(res, '/admin#setup');
    case '/signout':
      await signOut();
      audit({ tool: 'sign_out' });
      return go(res, 'ok', 'Signed out. Your sign-in was deleted from this computer.', back);
    case '/test': {
      const me = await graph('/me?$select=displayName,mail,userPrincipalName');
      await graph('/me/drive?$select=id');
      return go(res, 'ok', `Working. Connected to Microsoft 365 as ${me.displayName} (${me.mail || me.userPrincipalName}).`, back);
    }
    case '/connect': {
      const msgs = [];
      if (!serviceStatus().installed) msgs.push(installService());
      msgs.push(registerClaudeCode());
      if (registrationStatus().desktopInstalled) msgs.push(registerClaudeDesktop());
      audit({ tool: 'connect_claude' });
      return go(res, 'ok', msgs.join(' '));
    }
    case '/service/on':
      return go(res, 'ok', installService(), back);
    case '/service/off':
      return go(res, 'ok', `${uninstallService()} Claude can reach M365 Bridge only while this page is running.`, back);
    case '/workspace/add': {
      const r = await ws.pull({ source: body.source?.trim(), name: body.name?.trim() || undefined });
      return go(res, 'ok', `Copied ${plural(r.downloaded.length, 'file')} to ${r.workspace}.${r.skipped.length ? ` Skipped ${r.skipped.length}.` : ''}`);
    }
    case '/workspace/pull': {
      const r = await ws.pull({ localPath: body.localPath });
      const n = r.downloaded.length + r.updated.length;
      return go(res, r.needs_merge.length ? 'warn' : 'ok', `${plural(n, 'file')} updated.${r.needs_merge.length ? ` ${plural(r.needs_merge.length, 'file')} changed on both sides; ask Claude to merge.` : ''}`);
    }
    case '/workspace/push': {
      const r = await ws.push({ localPath: body.localPath, files: body.file ? [body.file] : undefined, note: body.note || 'Updated from the M365 Bridge page' });
      const n = r.uploaded.length + r.created.length;
      const issues = [
        r.held_by_others.length && `${plural(r.held_by_others.length, 'file')} checked out to a colleague`,
        r.needs_merge.length && `${plural(r.needs_merge.length, 'file')} need a merge (ask Claude)`,
        r.errors.length && `${r.errors[0].file}: ${r.errors[0].error}`,
      ].filter(Boolean);
      if (issues.length) return go(res, 'warn', `${n} sent to SharePoint. ${issues.join('. ')}.`);
      return go(res, 'ok', n ? `Sent ${plural(n, 'file')} to SharePoint and checked in.` : r.released.length ? `Released ${plural(r.released.length, 'file')}.` : 'Nothing to send. Everything matches SharePoint.');
    }
    case '/release': {
      const r = await ws.release({ localPath: body.localPath, files: body.file ? [body.file] : undefined });
      return go(res, 'ok', `Released ${plural(r.released.length, 'file')}. Colleagues can edit ${r.released.length === 1 ? 'it' : 'them'} again.${r.edits_kept_locally.length ? ' Your unsent edits are still on this computer.' : ''}`);
    }
    case '/workspace/open':
      if (!ws.listWorkspaces().some((w) => w.localPath === body.localPath)) throw new Error('Unknown folder.');
      openBrowser(body.localPath);
      return go(res, 'ok', `Opened ${path.basename(body.localPath)}.`);
    case '/workspace/forget':
      ws.forgetWorkspace(body.localPath);
      return go(res, 'ok', 'Removed from the list. The folder on this computer and SharePoint were not touched.');
    default:
      return send(res, 404, renderMessage('Not found', 'Nothing here.'));
  }
}

export function escape(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

let server = null;

/** Make sure the page is up: reuse the background service, or start it in this process. */
export async function ensureDashboard() {
  if (server) return DASHBOARD_URL;
  if (await isRunning()) return DASHBOARD_URL;
  return startDashboard();
}

export async function isRunning() {
  try {
    const r = await fetch(`http://127.0.0.1:${PORT}/healthz`, { signal: AbortSignal.timeout(800) });
    const j = r.ok ? await r.json() : null;
    return j?.app === 'm365-bridge' ? j : null;
  } catch {
    return null;
  }
}

export function startDashboard() {
  return new Promise((resolve, reject) => {
    const onRequest = (req, res) => {
      handle(req, res).catch((e) => {
        if (!res.headersSent) send(res, 500, renderMessage('Something went wrong', escape(e.message)));
      });
    };
    server = http.createServer(onRequest);
    server.on('error', (e) => {
      server = null;
      reject(e.code === 'EADDRINUSE' ? Object.assign(new Error(`Port ${PORT} is busy.`), { code: 'EADDRINUSE' }) : e);
    });
    server.listen(PORT, '127.0.0.1', () => {
      // "localhost" can resolve to IPv6 first; answer there too so the Microsoft redirect always lands.
      const v6 = http.createServer(onRequest);
      v6.on('error', () => {});
      v6.listen(PORT, '::1');
      resolve(DASHBOARD_URL);
    });
  });
}

