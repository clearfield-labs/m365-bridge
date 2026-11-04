// The M365 Bridge page. Plain and Microsoft-like: one screen for people, one for the admin.
import os from 'node:os';
import { PACKS, scopesFor } from '../scopes.js';
import { REDIRECT_URI } from '../paths.js';

const e = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const tidy = (p) => (p && p.startsWith(os.homedir()) ? `~${p.slice(os.homedir().length)}` : p);

// Fluent-style 20px line icons.
const P = {
  check: '<path d="m4.5 10.5 3.5 3.5 7.5-8"/>',
  folder: '<path d="M2.5 6A1.5 1.5 0 0 1 4 4.5h3.6l1.6 1.6h6.8A1.5 1.5 0 0 1 17.5 7.6V14.5A1.5 1.5 0 0 1 16 16H4a1.5 1.5 0 0 1-1.5-1.5z"/>',
  file: '<path d="M5.5 2.5h6l4 4v10a1 1 0 0 1-1 1h-9a1 1 0 0 1-1-1v-13a1 1 0 0 1 1-1z"/><path d="M11.5 2.5v4h4"/>',
  lock: '<rect x="4" y="9" width="12" height="8.5" rx="1.5"/><path d="M6.5 9V6.5a3.5 3.5 0 0 1 7 0V9"/>',
  up: '<path d="M10 15.5V4.5M5.5 9 10 4.5 14.5 9"/>',
  down: '<path d="M10 4.5v11M5.5 11l4.5 4.5 4.5-4.5"/>',
  open: '<path d="M11.5 3.5h5v5M16.5 3.5 9 11"/><path d="M14.5 11.5v4a1 1 0 0 1-1 1h-9a1 1 0 0 1-1-1v-9a1 1 0 0 1 1-1h4"/>',
  copy: '<rect x="7" y="7" width="10" height="10" rx="1.5"/><path d="M13 7V4.5A1.5 1.5 0 0 0 11.5 3h-7A1.5 1.5 0 0 0 3 4.5v7A1.5 1.5 0 0 0 4.5 13H7"/>',
  plus: '<path d="M10 4v12M4 10h12"/>',
  x: '<path d="m5 5 10 10M15 5 5 15"/>',
  more: '<circle cx="5" cy="10" r=".9"/><circle cx="10" cy="10" r=".9"/><circle cx="15" cy="10" r=".9"/>',
  chat: '<path d="M17 9.5a6.5 6.5 0 0 1-9.6 5.7L3.5 16.5l1.2-3.6A6.5 6.5 0 1 1 17 9.5z"/>',
  power: '<path d="M10 3v6.5"/><path d="M6 5.5a6 6 0 1 0 8 0"/>',
  shield: '<path d="M10 2.5 4 5v4.5c0 3.8 2.6 6.6 6 7.5 3.4-.9 6-3.7 6-7.5V5z"/>',
  warn: '<path d="M10 3 2.5 16.5h15z"/><path d="M10 8.5v3.5M10 14.3v.1"/>',
  merge: '<circle cx="6" cy="4.5" r="1.8"/><circle cx="6" cy="15.5" r="1.8"/><circle cx="14" cy="10" r="1.8"/><path d="M6 6.3v7.4M6 8c0 2 1.5 2 6.2 2"/>',
  key: '<circle cx="7" cy="12.5" r="3.5"/><path d="m9.5 10 6.5-6.5M13.5 6l2 2"/>',
  activity: '<path d="M2.5 10h3l2.5-6 4 12 2.5-6h3"/>',
  admin: '<circle cx="10" cy="10" r="2.5"/><path d="M10 2.5v2M10 15.5v2M2.5 10h2M15.5 10h2M4.7 4.7l1.4 1.4M13.9 13.9l1.4 1.4M4.7 15.3l1.4-1.4M13.9 6.1l1.4-1.4"/>',
};
const icon = (name, cls = 'i') => `<svg class="${cls}" viewBox="0 0 20 20" aria-hidden="true">${P[name] || ''}</svg>`;
const MS = '<span class="ms" aria-hidden="true"><i></i><i></i><i></i><i></i></span>';
const LOGO = '<svg class="logo" viewBox="0 0 32 32" aria-hidden="true"><rect width="32" height="32" rx="7" fill="var(--brand)"/><path d="M7 21c2.6-5.4 5.6-8 9-8s6.4 2.6 9 8" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round"/><path d="M6.5 21h19" stroke="#fff" stroke-width="2" stroke-linecap="round"/></svg>';

const initials = (a) => (a?.name || a?.username || '?').split(/[\s@.]+/).filter(Boolean).slice(0, 2).map((w) => w[0].toUpperCase()).join('');
const form = (action, inner, { busy, cls = '', back } = {}) =>
  `<form method="post" action="${action}"${cls ? ` class="${cls}"` : ''}${busy ? ` data-busy="${e(busy)}"` : ''}>${back ? `<input type="hidden" name="back" value="${back}">` : ''}${inner}</form>`;
const hidden = (name, value) => `<input type="hidden" name="${name}" value="${e(value)}">`;

function shell(title, body, flash, s, view) {
  const toast = flash ? `<div class="toast ${e(flash.kind)}" role="status">${icon(flash.kind === 'ok' ? 'check' : 'warn')}<span>${e(flash.text)}</span><button type="button" aria-label="Close">${icon('x')}</button></div>` : '';
  const who = s?.account
    ? `<div class="me"><span class="avatar">${e(initials(s.account))}</span><span class="who"><b>${e(s.account.name || s.account.username)}</b><small>${e(s.account.username)}</small></span>${form('/signout', '<button class="link" type="submit">Sign out</button>', { back: view === 'admin' ? 'admin' : '' })}</div>`
    : '';
  const nav = s ? `<nav><a href="/"${view === 'home' ? ' aria-current="page"' : ''}>Home</a><a href="/admin"${view === 'admin' ? ' aria-current="page"' : ''}>Admin</a></nav>` : '';
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${e(title)}</title><link rel="stylesheet" href="/assets/app.css"></head>
<body><header class="bar"><a class="brand" href="/">${LOGO}<span>M365 Bridge</span></a>${s?.cfg?.firmName ? `<span class="firm">${e(s.cfg.firmName)}</span>` : ''}${nav}<span class="grow"></span>${who}</header>
${toast}<main>${body}</main><script src="/assets/app.js"></script></body></html>`;
}

export function renderMessage(title, html) {
  return shell(title, `<section class="card"><div class="pad"><h1>${e(title)}</h1><p>${html}</p><p><a class="btn" href="/">Back</a></p></div></section>`, null);
}

const copyBox = (value, label = 'Copy') =>
  `<div class="copybox"><code>${e(value)}</code><button class="btn small" type="button" data-copy="${e(value)}">${icon('copy')}<span class="lbl">${label}</span></button></div>`;

// ---------------------------------------------------------------------------
// Home

function step(n, done, title, desc, action = '', disabled = false) {
  return `<li class="step${done ? ' done' : ''}${disabled ? ' off' : ''}"><span class="num">${done ? icon('check') : n}</span><div class="txt"><b>${title}</b><p>${desc}</p></div><div class="act">${action}</div></li>`;
}

function setup(s) {
  const reg = s.registration;
  const claudeOk = reg.claudeCode === 'connected' && reg.hook;
  const bgOk = s.service.installed;
  const firmDone = s.configured;
  const signedIn = !!s.account;
  const all = firmDone && signedIn && claudeOk && bgOk;

  const firm = firmDone
    ? step(1, true, 'Connected to your firm', e(s.cfg.firmName || 'Firm settings saved and encrypted on this computer.'))
    : step(1, false, 'Connect to your firm', 'Paste the setup code your admin sent you. It starts with <code>M365B-</code>.',
      form('/setup-code', '<input class="input" name="setupCode" placeholder="M365B-…" required autocomplete="off" spellcheck="false"><button class="btn primary" type="submit">Connect</button>', { cls: 'inline', busy: 'Connecting…' }) +
      '<a class="sub" href="/admin">I\'m the Microsoft 365 admin</a>');

  const sign = signedIn
    ? step(2, true, 'Signed in to Microsoft', `As ${e(s.account.username)}. Claude sees only what you can already open.`, form('/test', '<button class="btn small" type="submit">Test</button>', { busy: 'Testing…' }))
    : step(2, false, 'Sign in with Microsoft', 'Use your work account. Your sign-in is encrypted on this computer.',
      firmDone ? `<a class="btn ms-btn" href="/auth/start">${MS}Sign in with Microsoft</a>` : `<span class="btn ms-btn" aria-disabled="true">${MS}Sign in with Microsoft</span>`, !firmDone);

  const claudeDesc = claudeOk
    ? `Claude Code is connected${reg.claudeDesktop === 'connected' ? ', and so is the Claude chat app' : ''}. Restart Claude if you just connected it.`
    : reg.claudeCode === 'needs update' ? 'Claude points at an older setup. Reconnect it.' : 'Lets Claude use your files, email and calendar.';
  const claude = step(3, claudeOk, claudeOk ? 'Claude is connected' : 'Connect Claude', claudeDesc,
    claudeOk ? '' : form('/connect', '<button class="btn primary" type="submit">Connect Claude</button>', { busy: 'Connecting…' }), !signedIn);

  const bg = step(4, bgOk, bgOk ? 'Always running' : 'Keep it running',
    bgOk ? 'Starts by itself when your computer starts, so Claude can always reach it.' : 'Start M365 Bridge with your computer so Claude can always reach it.',
    bgOk ? '' : form('/service/on', '<button class="btn" type="submit">Turn on</button>', { busy: 'Turning on…' }));

  return `<section class="card" id="setup"><div class="card-h"><h2>${all ? `${icon('check', 'i ok')}You're all set` : 'Get set up'}</h2>${all ? '<button class="link" type="button" data-toggle="setup-steps">Show steps</button>' : ''}</div>
<ol class="steps" id="setup-steps"${all ? ' hidden' : ''}>${firm}${sign}${claude}${bg}</ol>
${all ? '<p class="muted pad">Ask Claude about your files, email or calendar. It works as you, in SharePoint and OneDrive.</p>' : ''}</section>`;
}

function held(s) {
  const rows = s.held.map((h) => {
    const state = h.needsMerge
      ? `<span class="pill warn">${icon('merge')}Needs a merge${h.mergeWith ? ` with ${e(h.mergeWith)}` : ''}</span>`
      : h.edited ? '<span class="pill info">Edited, not sent yet</span>' : '<span class="pill">Not changed</span>';
    const folder = `${e(h.workspace)}${h.file.includes('/') ? ` / ${e(h.file.split('/').slice(0, -1).join(' / '))}` : ''}`;
    const send = h.needsMerge
      ? '<span class="muted small">Ask Claude to merge it</span>'
      : h.edited ? form('/workspace/push', hidden('localPath', h.localPath) + hidden('file', h.file) + `<button class="btn small primary" type="submit">${icon('up')}Send to SharePoint</button>`, { busy: 'Sending…' }) : '';
    return `<tr><td><div class="fname">${icon('file')}<div><b>${e(h.file.split('/').pop())}</b><small>${folder}</small></div></div></td>
<td>${state}</td><td class="when"><time data-rel datetime="${e(h.since)}"></time></td>
<td class="row-act">${send}${form('/release', hidden('localPath', h.localPath) + hidden('file', h.file) + '<button class="btn small" type="submit" title="Let colleagues edit it again">Release</button>', { busy: 'Releasing…' })}</td></tr>`;
  });
  return `<section class="card"><div class="card-h"><h2>${icon('lock')}Checked out to you</h2>${s.held.length ? `<span class="count">${s.held.length}</span>` : ''}</div>
${rows.length ? `<table class="list"><tbody>${rows.join('')}</tbody></table>` : '<p class="empty">Nothing checked out. When Claude edits a SharePoint or OneDrive file, the file is checked out to you so nobody else can change it at the same time. It shows here until it is sent back.</p>'}</section>`;
}

function folders(s) {
  const rows = s.workspaces.filter((w) => w.exists).map((w) => {
    const where = /-my\.sharepoint\.com|\/personal\//.test(w.webUrl || '') ? 'OneDrive' : 'SharePoint';
    return `<li data-ws-path="${e(w.localPath)}"><div class="fname">${icon('folder', 'i fold')}<div><b>${e(w.name)}</b><small>${e(tidy(w.localPath))} · synced <time data-rel datetime="${e(w.lastPulled)}"></time></small><div class="status"><span class="pill">Checking…</span></div></div></div>
<div class="row-act">${form('/workspace/pull', hidden('localPath', w.localPath) + `<button class="btn small" type="submit" data-pull>${icon('down')}Get latest</button>`, { busy: 'Getting…' })}
${form('/workspace/push', hidden('localPath', w.localPath) + `<button class="btn small primary" type="submit" data-upload hidden>${icon('up')}Send changes</button>`, { busy: 'Sending…' })}
<details class="menu"><summary class="btn small icon" aria-label="More">${icon('more')}</summary><div class="pop">
${form('/workspace/open', hidden('localPath', w.localPath) + `<button type="submit">${icon('folder')}Open on this computer</button>`)}
${w.webUrl ? `<a href="${e(w.webUrl)}" target="_blank" rel="noopener">${icon('open')}Open in ${where}</a>` : ''}
<button type="button" data-copy="${e(w.localPath)}">${icon('copy')}<span class="lbl">Copy folder path</span></button>
${form('/workspace/forget', hidden('localPath', w.localPath) + `<button type="submit">${icon('x')}Remove from list</button>`)}</div></details></div></li>`;
  });
  return `<section class="card"><div class="card-h"><h2>${icon('folder')}Synced folders</h2><button class="btn small" type="button" data-open="add">${icon('plus')}Add folder</button></div>
${rows.length ? `<ul class="rows">${rows.join('')}</ul>` : '<p class="empty">No folders yet. Ask Claude <q>Get the Smith matter folder from SharePoint</q> and paste the folder link, or click Add folder.</p>'}</section>
<dialog id="add"><form method="post" action="/workspace/add" data-busy="Copying…"><h3>Add a folder</h3><p class="muted">Paste a link to a SharePoint or OneDrive folder. M365 Bridge copies it to <code>${e(tidy(s.workspacesRoot))}</code> so Claude can work on it.</p>
<label class="lbl-f" for="src">Folder link</label><input class="input" id="src" name="source" required placeholder="https://yourfirm.sharepoint.com/sites/…" autocomplete="off">
<label class="lbl-f" for="nm">Name on this computer <span class="muted">(optional)</span></label><input class="input" id="nm" name="name" autocomplete="off">
<div class="dlg-act"><button class="btn" type="button" data-close>Cancel</button><button class="btn primary" type="submit">Add folder</button></div></form></dialog>`;
}

const PROMPTS = [
  'Get the Smith matter folder from SharePoint. [paste the folder link]',
  'Update the timeline with the dates from the deposition transcript, show me the changes, then send it back to SharePoint.',
  'Organise the Smith folder: put pleadings, discovery and correspondence in their own folders.',
  'Find the settlement email from opposing counsel and draft a reply for me to review.',
];
function prompts() {
  return `<section class="card"><div class="card-h"><h2>${icon('chat')}Try asking Claude</h2><span class="muted small">Click to copy</span></div><ul class="prompts">${PROMPTS.map((p) => `<li><button type="button" data-copy="${e(p)}"><span>${e(p)}</span><span class="lbl">Copy</span></button></li>`).join('')}</ul></section>`;
}

const ACT = {
  sign_in: 'Signed in', sign_out: 'Signed out', firm_setup: 'Saved firm settings', firm_setup_auto: 'Set up the firm in Entra', admin_consent: 'Approved for the firm',
  connect_claude: 'Connected Claude', search: 'Searched', read_file: 'Read', download: 'Downloaded', upload: 'Uploaded', create_folder: 'Created a folder', move: 'Moved',
  workspace_pull: 'Got latest', workspace_push: 'Sent to SharePoint', workspace_lock: 'Checked out', workspace_release: 'Released', workspace_merge: 'Merging',
  workspace_move: 'Moved', restore_version: 'Restored a version', mail_search: 'Searched email', mail_read: 'Read an email', mail_draft: 'Drafted an email',
  mail_send: 'Sent an email', graph_request: 'Microsoft 365 request',
};
function detailOf(a) {
  if (a.error) return a.error;
  if (a.tool === 'workspace_push') return [Array.isArray(a.uploaded) && a.uploaded.join(', '), a.note].filter(Boolean).join(' · ') || 'Nothing to send';
  if (a.tool === 'workspace_lock' && Array.isArray(a.locked)) return a.locked.join(', ');
  if (a.tool === 'workspace_move') return `${a.from} → ${a.to}`;
  if (a.target && /^https?:/.test(a.target)) {
    try { return decodeURIComponent(new URL(a.target).pathname).split('/').filter(Boolean).slice(-2).join(' / '); } catch { /* fall through */ }
  }
  if (a.query) return `“${a.query}”`;
  return a.subject || (a.method ? `${a.method} ${a.path}` : '') || a.user || '';
}
function activity(s) {
  if (!s.audit.length) return '';
  return `<section class="card"><div class="card-h"><h2>${icon('activity')}Recent activity</h2><span class="muted small">Names only, never contents</span></div><ul class="acts">${s.audit
    .slice(0, 8)
    .map((a) => `<li class="${a.error ? 'err' : ''}"><b>${e(a.error ? `${ACT[a.tool] || a.tool} failed` : ACT[a.tool] || a.tool)}</b><span>${e(detailOf(a))}</span><time data-rel datetime="${e(a.ts)}"></time></li>`)
    .join('')}</ul></section>`;
}

function home(s) {
  const ready = s.configured && s.account;
  return `<div class="head"><h1>Microsoft 365 for Claude</h1><p>Claude works in your SharePoint, OneDrive, Outlook and Teams as you. It sees only what you can already open.</p></div>
${setup(s)}${ready ? held(s) + folders(s) + prompts() : ''}${activity(s)}`;
}

// ---------------------------------------------------------------------------
// Admin

function autoSetup(s) {
  const j = s.job;
  const again = j?.running
    ? `<p class="again"><span>Didn't see the Microsoft tab?</span> <a href="${e(j.signInUrl || '#')}" target="_blank" rel="noopener" data-signin${j.signInUrl ? '' : ' hidden'}>Open sign-in again</a>${form('/admin/cancel', '<button class="link" type="submit">Cancel</button>', { cls: 'inline-form' })}</p>`
    : '';
  const progress = j
    ? `<div class="progress" data-job${j.running ? ' data-running' : ''}><ul>${j.steps.map((x) => `<li>${icon('check')}<span>${e(x.text)}</span></li>`).join('')}${j.running ? '<li class="now"><span class="spin"></span><span>Working…</span></li>' : ''}</ul>${again}${j.error ? `<p class="bad">${icon('warn')}${e(j.error)}</p>` : ''}${j.done ? `<p class="good">${icon('check')}Done. Share the setup code below with your team.</p>` : ''}</div>`
    : '';
  const o = j?.options || {};
  const selected = o.access === 'selected';
  return `<section class="card" id="setup"><div class="card-h"><h2>${icon('admin')}Set up your firm</h2>${s.configured ? '<span class="pill ok">Done</span>' : ''}</div>
<div class="pad"><p>Sign in with a Microsoft 365 admin account. M365 Bridge creates its app in Entra and approves it for the firm. There is nothing to copy and no secret.</p>
${progress}
${form('/admin/auto', `<label class="lbl-f" for="fn">Firm name</label><input class="input short" id="fn" name="firmName" value="${e(o.firmName || s.cfg.firmName)}" placeholder="Smith &amp; Partners LLP">
<fieldset class="choice"><legend class="lbl-f">Who can use it</legend>
<label><input type="radio" name="access" value="everyone"${selected ? '' : ' checked'}><span><b>Everyone in the firm</b><small>Each person only sees their own files and email.</small></span></label>
<label><input type="radio" name="access" value="selected"${selected ? ' checked' : ''}><span><b>Only these people</b><small>Start small. Run this again later to add more.</small></span></label>
<textarea class="input" name="people" rows="2" placeholder="jeremy@firm.com, maria@firm.com">${e((o.people || []).join(', '))}</textarea></fieldset>
<button class="btn primary" type="submit"${j?.running ? ' disabled' : ''}>${s.configured ? 'Run setup again' : 'Sign in as admin and set up'}</button>`, { busy: 'Opening Microsoft sign-in…', back: 'admin' })}</div></section>`;
}

function code(s) {
  if (!s.setupCode) return '';
  return `<section class="card"><div class="card-h"><h2>${icon('key')}Setup code for your team</h2></div><div class="pad">
<p>Send this to everyone who will use Claude. They paste it into M365 Bridge, or tell Claude <q>set up M365 with this code</q>. It holds no secret, only your firm's IDs and settings.</p>
${copyBox(s.setupCode, 'Copy code')}
<dl class="facts"><dt>Tenant ID</dt><dd><code>${e(s.cfg.tenantId)}</code></dd><dt>Client ID</dt><dd><code>${e(s.cfg.clientId)}</code></dd></dl></div></section>`;
}

function settings(s) {
  const cfg = s.cfg;
  const packs = Object.entries(PACKS).map(([k, p]) => `<label class="check"><input type="checkbox" name="pack_${k}"${k === 'files' || (cfg.packs || []).includes(k) ? ' checked' : ''}${p.always ? ' disabled' : ''}><span>${e(p.label)}</span></label>`).join('');
  return `<details class="card"><summary class="card-h"><h2>${icon('shield')}Permissions</h2><span class="muted small">${scopesFor(cfg).length} Microsoft permissions</span></summary><div class="pad">
${form('/settings', `<label class="lbl-f" for="fn2">Firm name</label><input class="input short" id="fn2" name="firmName" value="${e(cfg.firmName)}">
<p class="lbl-f">What Claude can reach</p><div class="checks">${packs}</div>
<p class="lbl-f">What Claude can do</p><div class="checks">
<label class="check"><input type="checkbox" name="allowWrite"${cfg.allowWrite ? ' checked' : ''}><span>Edit and upload files, check files out, draft email</span></label>
<label class="check"><input type="checkbox" name="allowDelete"${cfg.allowDelete ? ' checked' : ''}><span>Delete through Microsoft Graph (off is safer)</span></label></div>
<label class="lbl-f" for="xs">Extra Microsoft Graph permissions <span class="muted">(optional)</span></label><input class="input mono" id="xs" name="extraScopes" value="${e(cfg.extraScopes)}" placeholder="Bookings.Read.All">
<div class="row"><button class="btn primary" type="submit">Save</button>${s.consentUrl ? `<a class="btn" href="${e(s.consentUrl)}">Approve for the firm again</a>` : ''}</div>`, { back: 'admin' })}
<p class="muted small">After changing permissions, run setup again or approve again, then send the team the new setup code.</p></div></details>`;
}

function manual(s) {
  return `<details class="card"><summary class="card-h"><h2>Set it up by hand instead</h2><span class="muted small">If IT prefers the Entra admin center</span></summary><div class="pad">
<ol class="howto">
<li>In <a href="https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank" rel="noopener">Entra admin center, App registrations</a>, choose <b>New registration</b>. Name it <code>Claude M365 Bridge</code>, single tenant.</li>
<li>Add a redirect URI with platform <b>Public client/native</b>:${copyBox(REDIRECT_URI)}</li>
<li>Under <b>Authentication</b>, set <b>Allow public client flows</b> to Yes.</li>
<li>Copy the <b>Directory (tenant) ID</b> and <b>Application (client) ID</b> into the form below. No client secret.</li>
<li>Save, then open Permissions above and click <b>Approve for the firm again</b>.</li>
<li>Optional: in Enterprise applications, set <b>Assignment required</b> to Yes and add people or a group.</li></ol>
${form('/firm', `<div class="grid2"><div><label class="lbl-f" for="t">Directory (tenant) ID</label><input class="input mono" id="t" name="tenantId" value="${e(s.cfg.tenantId)}" required></div>
<div><label class="lbl-f" for="c">Application (client) ID</label><input class="input mono" id="c" name="clientId" value="${e(s.cfg.clientId)}" required></div></div>
<button class="btn" type="submit">Save IDs</button>`, { back: 'admin' })}</div></details>`;
}

function computer(s) {
  const bg = s.service;
  return `<section class="card"><div class="card-h"><h2>${icon('power')}This computer</h2><span class="muted small">Version ${e(s.version)}</span></div><div class="pad">
<dl class="facts"><dt>Background service</dt><dd>${bg.installed ? `On${bg.running === false ? ' (not running right now)' : ''}` : 'Off'} ${form(bg.installed ? '/service/off' : '/service/on', `<button class="link" type="submit">${bg.installed ? 'Turn off' : 'Turn on'}</button>`, { back: 'admin', cls: 'inline-form' })}</dd>
<dt>Claude Code</dt><dd>${e(s.registration.claudeCode)}${s.registration.hook ? ', check-out on edit is on' : ''}</dd>
<dt>Encryption key</dt><dd>${e(s.keyStorage)}</dd>
<dt>Synced folders</dt><dd><code>${e(tidy(s.workspacesRoot))}</code></dd></dl></div></section>`;
}

function admin(s) {
  return `<div class="head"><h1>Firm setup</h1><p>For the Microsoft 365 admin. Done once for the whole firm.</p></div>
${autoSetup(s)}${code(s)}${settings(s)}${manual(s)}${computer(s)}`;
}

export function renderPage(s, flash, view = 'home') {
  return shell(view === 'admin' ? 'Firm setup · M365 Bridge' : 'M365 Bridge', view === 'admin' ? admin(s) : home(s), flash, s, view);
}
