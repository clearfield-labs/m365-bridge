import { PACKS, scopesFor, adminConsentUrl } from '../scopes.js';
import { REDIRECT_URI, HOME } from '../paths.js';

const e = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);

const CSS = `
:root{--bg:#f6f5f2;--card:#fff;--ink:#1d1d1b;--muted:#6b6a65;--line:#e4e2dc;--accent:#2f5d50;--accent-ink:#fff;--ok:#2f7a4f;--warn:#a26a12;--err:#a83232;--chip:#efede7;--code:#f1efe9}
@media (prefers-color-scheme:dark){:root{--bg:#141413;--card:#1d1d1b;--ink:#eceae4;--muted:#9a988f;--line:#2e2d2a;--accent:#7fb8a4;--accent-ink:#10201a;--ok:#7cc79a;--warn:#e0b25a;--err:#e88b8b;--chip:#2a2926;--code:#26251f}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Inter,sans-serif}
.wrap{max-width:880px;margin:0 auto;padding:32px 16px 80px}
header{display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:12px;margin-bottom:24px}
h1{font:600 24px/1.2 Georgia,"Times New Roman",serif;margin:0}h2{font:600 18px/1.3 Georgia,serif;margin:0 0 4px}
.sub{color:var(--muted);margin:2px 0 0}
.chips{display:flex;flex-wrap:wrap;gap:6px}.chip{background:var(--chip);border-radius:999px;padding:3px 10px;font-size:12.5px;color:var(--muted)}
.chip.on{color:var(--ok)}.chip.off{color:var(--warn)}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:20px;margin:14px 0}
.step{display:inline-flex;width:22px;height:22px;border-radius:50%;background:var(--accent);color:var(--accent-ink);font-size:12px;font-weight:700;align-items:center;justify-content:center;margin-right:8px;vertical-align:2px}
.step.done{background:var(--ok)}
label{display:block;font-size:13px;color:var(--muted);margin:12px 0 4px}
input[type=text],textarea{width:100%;padding:9px 11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit}
textarea{min-height:60px;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:13px}
.row{display:grid;grid-template-columns:1fr 1fr;gap:12px}@media(max-width:640px){.row{grid-template-columns:1fr}}
button,.btn{display:inline-block;border:0;border-radius:8px;padding:9px 16px;font:600 14px/1.2 inherit;cursor:pointer;background:var(--accent);color:var(--accent-ink);text-decoration:none}
button.ghost,.btn.ghost{background:transparent;color:var(--ink);border:1px solid var(--line)}
button.small{padding:6px 11px;font-size:13px}
.actions{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px;align-items:center}
.flash{border-radius:10px;padding:12px 14px;margin:0 0 8px;border:1px solid var(--line);background:var(--card)}
.flash.ok{border-color:var(--ok)}.flash.error{border-color:var(--err);color:var(--err)}.flash.warn{border-color:var(--warn)}
code,.code{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:12.5px;background:var(--code);padding:2px 6px;border-radius:5px;word-break:break-all}
.copy{display:flex;gap:8px;align-items:center;margin:6px 0}.copy code{flex:1;min-width:0;padding:8px 10px}
.card{overflow-wrap:anywhere}td{overflow-wrap:anywhere}
.packs{display:grid;grid-template-columns:1fr 1fr;gap:6px 16px;margin-top:6px}@media(max-width:640px){.packs{grid-template-columns:1fr}}
.check{display:flex;gap:8px;align-items:flex-start;color:var(--ink);font-size:14px;margin:0}
.check input{margin-top:4px}
details{margin-top:12px}summary{cursor:pointer;color:var(--accent);font-weight:600}
ol.guide{padding-left:20px}ol.guide li{margin:8px 0}
table{width:100%;border-collapse:collapse;font-size:13px}td,th{text-align:left;padding:7px 6px;border-bottom:1px solid var(--line);vertical-align:top}th{color:var(--muted);font-weight:500}
.ws{border:1px solid var(--line);border-radius:10px;padding:12px 14px;margin:10px 0}
.ws .name{font-weight:600}.muted{color:var(--muted);font-size:13px}
.scopes{display:flex;flex-wrap:wrap;gap:4px;margin-top:8px}.scopes code{font-size:11.5px}
footer{color:var(--muted);font-size:12.5px;margin-top:28px}
`;

const SCRIPT = `
document.addEventListener('click',function(ev){var b=ev.target.closest('[data-copy]');if(!b)return;
navigator.clipboard.writeText(b.getAttribute('data-copy')).then(function(){var t=b.textContent;b.textContent='Copied';setTimeout(function(){b.textContent=t},1200)})});
document.querySelectorAll('form[data-busy]').forEach(function(f){f.addEventListener('submit',function(){var b=f.querySelector('button[type=submit],button:not([type])');if(b){b.disabled=true;b.textContent=f.getAttribute('data-busy')}})});
`;

const copy = (value, label = 'Copy') => `<div class="copy"><code>${e(value)}</code><button type="button" class="ghost small" data-copy="${e(value)}">${label}</button></div>`;

function shell(title, body) {
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${e(title)}</title><style>${CSS}</style></head><body><div class="wrap">${body}</div><script>${SCRIPT}</script></body></html>`;
}

export function renderMessage(title, html) {
  return shell(title, `<div class="card"><h2>${e(title)}</h2><p>${html}</p><div class="actions"><a class="btn ghost" href="/">Back to dashboard</a></div></div>`);
}

function entraGuide(port) {
  return `<details><summary>Admin: create the Entra app (5 minutes, once per firm)</summary>
<ol class="guide">
<li>Go to <a href="https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank" rel="noopener">entra.microsoft.com → App registrations</a> → <b>New registration</b>.</li>
<li>Name: <code>Claude M365 Bridge</code>. Supported account types: <b>Accounts in this organizational directory only</b>.</li>
<li>Redirect URI: platform <b>Public client/native (mobile &amp; desktop)</b>, value:${copy(REDIRECT_URI)}</li>
<li>Open <b>Authentication</b> → turn on <b>Allow public client flows</b> → Save. (Lets people sign in with a code from inside Claude.)</li>
<li>Copy the <b>Directory (tenant) ID</b> and <b>Application (client) ID</b> from <b>Overview</b> into the form below. <b>No client secret.</b> Nothing secret lives on a laptop.</li>
<li>Save, then click <b>Grant admin consent</b>. That approves every permission you ticked, for the whole firm, in one click.</li>
<li>Optional, to choose who gets it: <b>Enterprise applications</b> → Claude M365 Bridge → <b>Properties</b> → <b>Assignment required = Yes</b>, then <b>Users and groups</b> → add the people or a group.</li>
<li>Optional, recommended: a Conditional Access policy on this app (MFA, managed devices only).</li>
</ol>
<p class="muted">Port ${port}. If you change the port, add the matching redirect URI too.</p></details>`;
}

function firmSection(s) {
  const { cfg, configured } = s;
  const packs = Object.entries(PACKS)
    .map(([key, p]) => `<label class="check"><input type="checkbox" name="pack_${key}" ${p.always ? 'checked disabled' : (cfg.packs || []).includes(key) ? 'checked' : ''}> ${e(p.label)}</label>`)
    .join('');
  const form = `<form method="post" action="/firm" data-busy="Saving…">
<div class="row"><div><label>Directory (tenant) ID</label><input type="text" name="tenantId" value="${e(cfg.tenantId)}" placeholder="00000000-0000-0000-0000-000000000000" autocomplete="off" spellcheck="false"></div>
<div><label>Application (client) ID</label><input type="text" name="clientId" value="${e(cfg.clientId)}" placeholder="00000000-0000-0000-0000-000000000000" autocomplete="off" spellcheck="false"></div></div>
<label>Firm name (shown to your team)</label><input type="text" name="firmName" value="${e(cfg.firmName)}" placeholder="Smith &amp; Barid LLP">
<label>What Claude can reach (each person only ever sees their own data)</label><div class="packs">${packs}</div>
<label>Extra Graph scopes (optional, space separated)</label><input type="text" name="extraScopes" value="${e(cfg.extraScopes)}" placeholder="e.g. Bookings.Read.All" spellcheck="false">
<div class="packs" style="margin-top:12px"><label class="check"><input type="checkbox" name="allowWrite" ${cfg.allowWrite ? 'checked' : ''}> Allow changes (upload, draft email, create events)</label>
<label class="check"><input type="checkbox" name="allowDelete" ${cfg.allowDelete ? 'checked' : ''}> Allow deletes through Graph</label></div>
<div class="actions"><button type="submit">Save and encrypt</button></div></form>`;

  if (!configured) {
    return `<section class="card" id="firm"><h2><span class="step">1</span>Connect your firm</h2>
<p class="sub">Your admin gives you a setup code. Paste it here. Admins: fill in the form below once, then share the code with your team.</p>
<form method="post" action="/firm" data-busy="Saving…"><label>Setup code from your admin</label><textarea name="setupCode" placeholder="M365B-…"></textarea><div class="actions"><button type="submit">Use setup code</button></div></form>
<details><summary>I'm the admin. Enter the IDs from Entra</summary>${form}</details>${entraGuide(s.port)}</section>`;
  }
  return `<section class="card" id="firm"><h2><span class="step done">1</span>${e(cfg.firmName || 'Your firm')} is connected</h2>
<p class="sub">Tenant <code>${e(cfg.tenantId)}</code> · App <code>${e(cfg.clientId)}</code> · ${cfg.allowWrite ? 'read and write' : 'read only'}${cfg.allowDelete ? ' · deletes allowed' : ''}</p>
<div class="scopes">${scopesFor(cfg).map((x) => `<code>${e(x)}</code>`).join('')}</div>
<label>Setup code for your team (no secrets inside: tenant, app and permission choices)</label>${copy(s.setupCode)}
<div class="actions"><a class="btn" href="${e(adminConsentUrl(cfg, REDIRECT_URI))}">Grant admin consent for the firm</a><span class="muted">Global admin only. One click approves these permissions for everyone.</span></div>
<details><summary>Change firm settings</summary>${form}</details>${entraGuide(s.port)}</section>`;
}

function accountSection(s) {
  if (!s.configured) {
    return `<section class="card" id="account"><h2><span class="step">2</span>Sign in with Microsoft</h2><p class="sub">Connect your firm first.</p></section>`;
  }
  if (!s.account) {
    return `<section class="card" id="account"><h2><span class="step">2</span>Sign in with Microsoft</h2>
<p class="sub">Use your own work account. Claude will see only what you can already see in Microsoft 365.</p>
<div class="actions"><a class="btn" href="/auth/start">Sign in with Microsoft</a></div></section>`;
  }
  return `<section class="card" id="account"><h2><span class="step done">2</span>Signed in as ${e(s.account.name || s.account.username)}</h2>
<p class="sub">${e(s.account.username)} · tokens encrypted with ${e(s.keyStorage)}</p>
<div class="actions"><form method="post" action="/test" data-busy="Testing…"><button type="submit">Test connection</button></form>
<form method="post" action="/signout"><button class="ghost" type="submit">Sign out</button></form></div></section>`;
}

function claudeSection(s) {
  const r = s.registration;
  const tag = (v) => `<span class="chip ${v === 'connected' ? 'on' : 'off'}">${e(v)}</span>`;
  return `<section class="card" id="claude"><h2><span class="step ${r.claudeCode === 'connected' ? 'done' : ''}">3</span>Connect to Claude</h2>
<table><tr><td>Claude Code (desktop app and terminal)</td><td>${tag(r.claudeCode)}</td></tr>
<tr><td>Claude Desktop chat</td><td>${tag(r.claudeDesktop)}</td></tr></table>
<form method="post" action="/register" data-busy="Connecting…"><div class="actions">
<button name="target" value="both" type="submit">Connect both</button>
<button class="ghost" name="target" value="code" type="submit">Claude Code only</button>
<button class="ghost" name="target" value="desktop" type="submit">Claude Desktop only</button></div></form>
<p class="muted">Restart Claude afterwards. Then try: "Search my SharePoint for the Henderson engagement letter."</p></section>`;
}

function workspacesSection(s) {
  const list = s.workspaces.length
    ? s.workspaces
        .map(
          (w) => `<div class="ws"><div class="name">${e(w.name)}${w.exists ? '' : ' <span class="chip off">folder missing</span>'}</div>
<div class="muted">${e(w.localPath)}</div><div class="muted"><a href="${e(w.webUrl)}" target="_blank" rel="noopener">Open in SharePoint</a> · last pulled ${e(new Date(w.lastPulled).toLocaleString())}</div>
<div class="actions">
<form method="post" action="/workspace/pull" data-busy="Pulling…"><input type="hidden" name="localPath" value="${e(w.localPath)}"><button class="small" type="submit">Pull latest</button></form>
<form method="post" action="/workspace/open"><input type="hidden" name="localPath" value="${e(w.localPath)}"><button class="ghost small" type="submit">Open folder</button></form>
<button class="ghost small" type="button" data-copy="${e(w.localPath)}">Copy path</button>
<form method="post" action="/workspace/forget"><input type="hidden" name="localPath" value="${e(w.localPath)}"><button class="ghost small" type="submit">Remove from list</button></form>
</div></div>`,
        )
        .join('')
    : '<p class="muted">No workspaces yet.</p>';
  const disabled = s.account ? '' : 'disabled';
  return `<section class="card" id="workspaces"><h2><span class="step ${s.workspaces.length ? 'done' : ''}">4</span>Matter folders</h2>
<p class="sub">A matter folder is a local working copy of a SharePoint folder, like a GitHub repo. SharePoint stays the source of truth: Claude pulls the latest, edits, and pushes back as new versions. If a colleague changed the same file, the push stops instead of overwriting their work.</p>
<form method="post" action="/workspace/add" data-busy="Pulling files…"><div class="row">
<div><label>SharePoint or OneDrive folder link</label><input type="text" name="source" placeholder="https://firm.sharepoint.com/sites/Litigation/Shared Documents/Henderson" ${disabled}></div>
<div><label>Local name (optional)</label><input type="text" name="name" placeholder="Henderson v. Acme" ${disabled}></div></div>
<div class="actions"><button type="submit" ${disabled}>Pull folder</button><span class="muted">Saved under ${e(s.workspacesRoot)}</span></div></form>
${list}</section>`;
}

function activitySection(s) {
  const rows = s.audit
    .map((a) => {
      const what = a.error ? `<span style="color:var(--err)">${e(a.error)}</span>` : e(a.target || a.query || a.path || a.subject || a.local || a.user || a.tenant || a.method || '');
      return `<tr><td>${e(new Date(a.ts).toLocaleString())}</td><td>${e(a.tool)}</td><td>${what}</td></tr>`;
    })
    .join('');
  return `<section class="card" id="activity"><h2>Activity</h2><p class="sub">Everything Claude did through this bridge on this computer. File contents are never logged.</p>
${rows ? `<table><tr><th>When</th><th>Action</th><th>Detail</th></tr>${rows}</table>` : '<p class="muted">Nothing yet.</p>'}</section>`;
}

export function renderPage(s, flash) {
  const chips = [
    `<span class="chip ${s.configured ? 'on' : 'off'}">${s.configured ? 'Firm connected' : 'Firm not set up'}</span>`,
    `<span class="chip ${s.account ? 'on' : 'off'}">${s.account ? e(s.account.username) : 'Not signed in'}</span>`,
    `<span class="chip">${e(s.keyStorage)}</span>`,
  ].join('');
  return shell(
    'M365 Bridge',
    `<header><div><h1>M365 Bridge for Claude</h1><p class="sub">SharePoint, OneDrive, Outlook and Teams in Claude Code, signed in as you.</p></div><div class="chips">${chips}</div></header>
${flash ? `<div class="flash ${e(flash.kind)}">${e(flash.text)}</div>` : ''}
${firmSection(s)}${accountSection(s)}${claudeSection(s)}${workspacesSection(s)}${activitySection(s)}
<footer>Runs only on this computer (localhost). Settings and tokens are encrypted with AES-256-GCM; the key is held by ${e(s.keyStorage)}. Data folder: <code>${e(HOME)}</code></footer>`,
  );
}
