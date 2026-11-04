import { PACKS, scopesFor, adminConsentUrl } from '../scopes.js';
import os from 'node:os';
import { REDIRECT_URI } from '../paths.js';

const tidy = (p) => (p && p.startsWith(os.homedir()) ? `~${p.slice(os.homedir().length)}` : p);

const e = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);

// ---------- Icons (1.6px line, 24 grid) ----------
const P = {
  folder: '<path d="M3 7.5A2.5 2.5 0 0 1 5.5 5H9l2 2h7.5A2.5 2.5 0 0 1 21 9.5v8a2.5 2.5 0 0 1-2.5 2.5h-13A2.5 2.5 0 0 1 3 17.5z"/>',
  mail: '<rect x="3" y="5" width="18" height="14" rx="2.5"/><path d="m3.5 7 8.5 6 8.5-6"/>',
  calendar: '<rect x="3" y="5" width="18" height="16" rx="2.5"/><path d="M16 3v4M8 3v4M3 10.5h18"/>',
  users: '<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0"/><path d="M16 4.6a3.5 3.5 0 0 1 0 6.8M21.5 20a6.5 6.5 0 0 0-4-6"/>',
  chat: '<path d="M20.5 12a8 8 0 0 1-11.7 7.1L4 20.5l1.4-4.6A8 8 0 1 1 20.5 12z"/>',
  video: '<rect x="2.5" y="6" width="13" height="12" rx="2.5"/><path d="m15.5 10.5 6-3.5v10l-6-3.5"/>',
  tasks: '<rect x="3.5" y="3.5" width="17" height="17" rx="3"/><path d="m8 12 3 3 5-6"/>',
  notebook: '<path d="M6.5 3H17a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6.5z"/><path d="M6.5 3v18M4 7.5h2.5M4 12h2.5M4 16.5h2.5"/>',
  person: '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="10" r="3"/><path d="M6.6 18.6a6 6 0 0 1 10.8 0"/>',
  shield: '<path d="M12 3 4.5 6v6c0 4.6 3.2 7.9 7.5 9 4.3-1.1 7.5-4.4 7.5-9V6z"/><path d="m9 12 2 2 4-4"/>',
  key: '<circle cx="8" cy="15" r="4"/><path d="m10.8 12.2 8.7-8.7M16 6l3 3M14 8l2 2"/>',
  upload: '<path d="M12 15V4M7.5 8.5 12 4l4.5 4.5"/><path d="M4 15v3a2.5 2.5 0 0 0 2.5 2.5h11A2.5 2.5 0 0 0 20 18v-3"/>',
  download: '<path d="M12 4v11M7.5 10.5 12 15l4.5-4.5"/><path d="M4 15v3a2.5 2.5 0 0 0 2.5 2.5h11A2.5 2.5 0 0 0 20 18v-3"/>',
  external: '<path d="M14 4h6v6M20 4l-9 9"/><path d="M18 14v4.5a1.5 1.5 0 0 1-1.5 1.5h-11A1.5 1.5 0 0 1 4 18.5v-11A1.5 1.5 0 0 1 5.5 6H10"/>',
  copy: '<rect x="8" y="8" width="12.5" height="12.5" rx="2.5"/><path d="M16 8V5.5A2.5 2.5 0 0 0 13.5 3h-8A2.5 2.5 0 0 0 3 5.5v8A2.5 2.5 0 0 0 5.5 16H8"/>',
  home: '<path d="m3 11 9-7.5 9 7.5"/><path d="M5.5 9.5V20h13V9.5"/><path d="M10 20v-5.5h4V20"/>',
  activity: '<path d="M3 12h4l3-8 4 16 3-8h4"/>',
  sliders: '<path d="M4 6h9M18 6h2M4 12h3M12 12h8M4 18h11M19 18h1"/><circle cx="15.5" cy="6" r="2.2"/><circle cx="9.5" cy="12" r="2.2"/><circle cx="17" cy="18" r="2.2"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
  check: '<path d="m5 12.5 4.5 4.5L19 7.5"/>',
  alert: '<path d="M12 3.5 2.5 20h19z"/><path d="M12 10v4.5M12 17.2v.1"/>',
  lock: '<rect x="4.5" y="10.5" width="15" height="10" rx="2.5"/><path d="M8 10.5V7.5a4 4 0 0 1 8 0v3"/>',
  logout: '<path d="M15 4h3a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2h-3"/><path d="m10 16.5-4.5-4.5L10 7.5M5.5 12H16"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  x: '<path d="M6 6l12 12M18 6 6 18"/>',
  eye: '<path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="3"/>',
  pen: '<path d="M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z"/><path d="m13.5 6.5 4 4"/>',
  history: '<path d="M3.5 12a8.5 8.5 0 1 0 2.6-6.1"/><path d="M3.5 4v4.5H8"/><path d="M12 8v4.5l3 2"/>',
  terminal: '<rect x="3" y="4" width="18" height="16" rx="2.5"/><path d="m7.5 9.5 3 2.5-3 2.5M13 15h4"/>',
  link: '<path d="M10 14a4.5 4.5 0 0 0 6.4 0l3-3a4.5 4.5 0 0 0-6.4-6.4l-1 1"/><path d="M14 10a4.5 4.5 0 0 0-6.4 0l-3 3a4.5 4.5 0 0 0 6.4 6.4l1-1"/>',
  sparkle: '<path d="M12 3.5 13.8 9 19.5 10.8 13.8 12.6 12 18.5 10.2 12.6 4.5 10.8 10.2 9z"/>',
  cloud: '<path d="M7 18.5a4.5 4.5 0 0 1-.6-9A6 6 0 0 1 18 8.5a4 4 0 0 1-.5 10z"/>',
};
const icon = (name, cls = 'i') => `<svg class="${cls}" viewBox="0 0 24 24" aria-hidden="true">${P[name] || P.sparkle}</svg>`;

const BRAND = `<svg viewBox="0 0 40 40" aria-hidden="true"><rect x="1" y="1" width="38" height="38" rx="11" fill="none" stroke="#D6B26A" stroke-opacity=".55"/><path d="M8 27c3.5-8 7.5-12 12-12s8.5 4 12 12" fill="none" stroke="#D6B26A" stroke-width="1.8" stroke-linecap="round"/><path d="M8 27h24M12.5 27v-5.5M27.5 27v-5.5M20 27V15" stroke="#EDE6D3" stroke-width="1.6" stroke-linecap="round"/></svg>`;
const MS = (cls = '') => `<span class="ms-logo ${cls}" aria-hidden="true"><i></i><i></i><i></i><i></i></span>`;

const PACK_ICON = { files: 'folder', mail: 'mail', calendar: 'calendar', contacts: 'users', teams: 'chat', meetings: 'video', tasks: 'tasks', notes: 'notebook', people: 'person' };
const PACK_COPY = {
  files: ['Search, read and work on documents in SharePoint and OneDrive.', 'Search and read documents in SharePoint and OneDrive.'],
  mail: ['Search and read email, including shared inboxes. Drafts replies and sends only when you say so.', 'Search and read email, including shared inboxes.'],
  calendar: ['See your calendar and shared calendars. Books meetings when asked.', 'See your calendar and shared calendars.'],
  contacts: ['Look up and update your contacts.', 'Look up your contacts.'],
  teams: ['Read your chats and channels. Posts only when asked.', 'Read your chats and channels.'],
  meetings: ['Find meetings and read their transcripts and recordings.', 'Find meetings and read their transcripts.'],
  tasks: ['Read and update To Do and Planner tasks.', 'Read To Do and Planner tasks.'],
  notes: ['Read and write OneNote notebooks.', 'Read OneNote notebooks.'],
  people: ['Look up colleagues and see who is available.', 'Look up colleagues and see who is available.'],
};
const packOn = (cfg, key) => key === 'files' || (cfg.packs || []).includes(key);

const initials = (name, email) => {
  const src = (name || email || '?').replace(/@.*/, '').split(/[\s._-]+/).filter(Boolean);
  return ((src[0]?.[0] || '?') + (src[1]?.[0] || '')).toUpperCase();
};
const firstName = (a) => (a?.name || a?.username || '').split(/[\s@.]/)[0] || '';

// ---------- Shell ----------
function shell(title, body, flash) {
  const t = flash
    ? `<div class="toasts" role="status"><div class="toast ${e(flash.kind)}">${icon(flash.kind === 'ok' ? 'check' : 'alert')}<div>${e(flash.text)}</div><button type="button" aria-label="Dismiss">×</button></div></div>`
    : '';
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${e(title)}</title><link rel="preload" href="/assets/fonts/serif.woff2" as="font" type="font/woff2" crossorigin>
<link rel="stylesheet" href="/assets/app.css"></head><body>${body}${t}<script src="/assets/app.js" defer></script></body></html>`;
}

export function renderMessage(title, html) {
  return shell(title, `<div class="welcome" style="grid-template-columns:1fr"><div class="panel" style="align-items:center"><div class="card"><div class="eyebrow">M365 Bridge</div><h2 style="margin:8px 0 10px">${e(title)}</h2><p class="lede">${html}</p><p style="margin-top:22px"><a class="btn btn-primary" href="/">Back to the dashboard</a></p></div></div></div>`);
}

// ---------- Shared pieces ----------
function copyLine(value) {
  return `<div class="copyline"><code>${e(value)}</code><button type="button" class="btn btn-quiet btn-sm" data-copy="${e(value)}">${icon('copy')}<span class="lbl">Copy</span></button></div>`;
}

function firmForm(cfg) {
  const toggles = Object.entries(PACKS)
    .map(([key, p]) => `<label class="toggle"><input type="checkbox" name="pack_${key}" ${p.always ? 'checked disabled' : packOn(cfg, key) ? 'checked' : ''}><span class="box">${icon('check')}</span>${icon(PACK_ICON[key])}${e(p.label.split(' (')[0])}</label>`)
    .join('');
  return `<form class="form" method="post" action="/firm" data-busy="Saving">
<div class="grid"><div class="field"><label for="t">Directory (tenant) ID</label><input id="t" class="input mono" name="tenantId" value="${e(cfg.tenantId)}" placeholder="00000000-0000-0000-0000-000000000000" autocomplete="off" spellcheck="false" required></div>
<div class="field"><label for="c">Application (client) ID</label><input id="c" class="input mono" name="clientId" value="${e(cfg.clientId)}" placeholder="00000000-0000-0000-0000-000000000000" autocomplete="off" spellcheck="false" required></div></div>
<div class="field"><label for="f">Firm name</label><input id="f" class="input" name="firmName" value="${e(cfg.firmName)}" placeholder="Henderson &amp; Vale LLP"></div>
<div class="field"><label>What Claude can reach</label><div class="toggles">${toggles}</div><div class="hint">Each person only ever reaches their own data and what is already shared with them.</div></div>
<div class="field"><label>Changes</label><div class="toggles">
<label class="toggle"><input type="checkbox" name="allowWrite" ${cfg.allowWrite ? 'checked' : ''}><span class="box">${icon('check')}</span>${icon('pen')}Allow changes and uploads</label>
<label class="toggle"><input type="checkbox" name="allowDelete" ${cfg.allowDelete ? 'checked' : ''}><span class="box">${icon('check')}</span>${icon('x')}Allow deletes</label></div></div>
<div class="field"><label for="x">Extra Microsoft Graph scopes <span class="muted">(optional)</span></label><input id="x" class="input mono" name="extraScopes" value="${e(cfg.extraScopes)}" placeholder="Bookings.Read.All" spellcheck="false"></div>
<div><button class="btn btn-primary" type="submit">${icon('lock')}Save and encrypt</button></div></form>`;
}

function entraSteps() {
  return `<ol class="steps">
<li><b>Create an app registration</b><p>In <a href="https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank" rel="noopener">Entra admin center → App registrations</a>, choose <b>New registration</b>. Name it <code>Claude M365 Bridge</code> and choose <b>this organizational directory only</b>.</p></li>
<li><b>Add the redirect URI</b><p>Platform <b>Public client/native (mobile &amp; desktop)</b>, value:</p>${copyLine(REDIRECT_URI)}</li>
<li><b>Allow public client flows</b><p>Under <b>Authentication</b>, switch <b>Allow public client flows</b> to Yes. This lets people sign in from inside Claude with a code.</p></li>
<li><b>Copy two IDs</b><p>From <b>Overview</b>: the <b>Directory (tenant) ID</b> and the <b>Application (client) ID</b>. No client secret. Nothing secret ever sits on a laptop.</p></li>
<li><b>Grant consent once</b><p>Save the IDs here, then press <b>Grant admin consent</b>. One click approves the chosen permissions for the whole firm.</p></li>
<li><b>Choose who gets it</b><p>Optional. Under <b>Enterprise applications → Properties</b>, set <b>Assignment required</b> to Yes, then add people or a group under <b>Users and groups</b>.</p></li></ol>`;
}

// ---------- Onboarding ----------
function onboarding(s) {
  return `<div class="welcome">
<section class="story"><a class="brand" href="/">${BRAND}<span><b>M365 Bridge</b><small>for Claude</small></span></a>
<h1>Your firm's Microsoft 365, <em>inside Claude.</em></h1>
<p class="lede">SharePoint, OneDrive, Outlook and Teams, signed in as you. Claude sees what you can see, and nothing more.</p>
<ol><li><span class="n">i</span><div><b>Connect your firm</b><span>Paste the setup code from your IT admin.</span></div></li>
<li><span class="n">ii</span><div><b>Sign in with Microsoft</b><span>Your own work account. Your own permissions.</span></div></li>
<li><span class="n">iii</span><div><b>Open Claude</b><span>Ask for a document, a matter folder, an email. Claude does the rest.</span></div></li></ol>
<div class="seal-line">${icon('lock')}Runs only on this computer. Everything is encrypted with ${e(s.keyStorage)}.</div></section>
<section class="panel"><div class="card">
<div class="eyebrow">Step one</div><h2 style="margin:8px 0 6px">Connect your firm</h2>
<p class="muted" style="margin:0 0 20px">Your IT admin sends one setup code for the whole firm.</p>
<form class="form" method="post" action="/firm" data-busy="Connecting"><div class="field"><label for="code">Setup code</label><textarea id="code" class="input" name="setupCode" placeholder="M365B-…" required></textarea></div>
<div><button class="btn btn-primary" type="submit">Connect firm ${icon('check')}</button></div></form>
<div class="or">or</div>
<details class="more" style="border:0;margin:0;padding:0"><summary>I'm the IT admin, setting this up for the firm</summary>
<div style="margin-top:18px">${firmForm(s.cfg)}</div>
<details class="more"><summary>How to create the Entra app (5 minutes)</summary><div style="margin-top:8px">${entraSteps()}</div></details></details>
</div></section></div>`;
}

// ---------- App ----------
function rail(s) {
  const { cfg, account } = s;
  const n = s.workspaces.length;
  const nav = [
    ['overview', 'home', 'Overview'],
    ['matters', 'folder', 'Matters', n ? `<span class="count">${n}</span>` : ''],
    ['access', 'shield', 'Access'],
    ['activity', 'activity', 'Activity'],
    ['admin', 'sliders', 'Admin'],
  ]
    .map(([id, ic, label, extra = '']) => `<a href="#${id}">${icon(ic)}<span class="lbl">${label}</span>${extra}</a>`)
    .join('');
  const me = account
    ? `<div class="me"><span class="avatar">${e(initials(account.name, account.username))}</span><span class="who"><b>${e(account.name || account.username)}</b><span>${e(account.username)}</span></span>
<form method="post" action="/signout"><button type="submit" title="Sign out">${icon('logout')}</button></form></div>`
    : `<a class="ms-btn" href="/auth/start" style="height:40px;font-size:14px">${MS()}Sign in</a>`;
  return `<aside class="rail"><a class="brand" href="#overview">${BRAND}<span><b>M365 Bridge</b><small>for Claude</small></span></a>
<div class="firm"><div class="k">Firm</div><div class="v">${e(cfg.firmName || 'Your firm')}</div></div>
<nav class="nav">${nav}</nav>
<div class="rail-foot">${me}<div class="lockline">${icon('lock')}Encrypted · ${e(s.keyStorage.replace(/ \(.*\)/, ''))}</div></div></aside>`;
}

function prompts(s) {
  const matter = s.workspaces[0]?.name || 'Henderson v. Acme';
  const list = [
    [`Pull the ${matter} folder from SharePoint so we can work on it.`, 'files'],
    ['Find the latest email from opposing counsel and draft a reply for me to review.', 'mail'],
    ['Who changed the complaint last, and what did the previous version say?', 'files'],
    ["What's on my calendar tomorrow? Summarise yesterday's client call from its transcript.", 'calendar'],
  ].filter(([, pack]) => packOn(s.cfg, pack)).map(([p]) => p);
  return `<div class="grid g${Math.min(list.length, 4)}">${list.map((p) => `<button type="button" class="prompt"><q>${e(p)}</q><span class="tag">${icon('copy')}<span class="lbl">Click to copy</span></span></button>`).join('')}</div>`;
}

function bridge(s) {
  const signed = !!s.account;
  const claude = s.registration.claudeCode === 'connected' || s.registration.claudeDesktop === 'connected';
  const youOrb = signed ? `<span class="avatar">${e(initials(s.account.name, s.account.username))}</span>` : icon('person');
  const youAct = signed
    ? `<div class="state">${e(s.account.username)}</div>`
    : `<div class="state">Not signed in yet</div><div class="act"><a class="ms-btn" href="/auth/start">${MS()}Sign in with Microsoft</a></div>`;
  const where = [s.registration.claudeCode === 'connected' && 'Claude Code', s.registration.claudeDesktop === 'connected' && 'Claude Desktop'].filter(Boolean).join(' & ');
  const claudeAct = claude
    ? `<div class="state">${e(where)}</div>`
    : `<div class="state">Not connected yet</div><div class="act"><form method="post" action="/register" data-busy="Connecting"><button class="btn btn-brass" name="target" value="both" type="submit">${icon('link')}Connect to Claude</button></form></div>`;
  const all = signed && claude;
  return `<div class="bridge"><div class="span">
<div class="node on"><div class="orb">${MS('lg')}</div><div class="label">Microsoft 365</div><div class="state">${e(s.cfg.firmName || 'Your firm')}</div></div>
<div class="wire ${signed ? 'live' : ''}"></div>
<div class="node ${signed ? 'on' : ''}"><div class="orb">${youOrb}</div><div class="label">${signed ? e(firstName(s.account)) : 'You'}</div>${youAct}</div>
<div class="wire ${all ? 'live rev' : ''}"></div>
<div class="node ${claude ? 'on' : ''}"><div class="orb"><span class="claude-mark">C</span></div><div class="label">Claude</div>${claudeAct}</div>
</div><div class="bridge-foot">${icon(all ? 'shield' : 'lock')}${all ? 'Claude works as you. It can open only what your Microsoft account can open.' : signed ? 'Connect Claude and it can work with your files, as you. Nothing more.' : 'Nothing reaches Claude until you sign in. Your access stays your access.'}</div></div>`;
}

function dossier(w, { mini = false } = {}) {
  return `<article class="dossier ${mini ? 'mini' : ''}" data-ws-path="${e(w.localPath)}"><span class="tab">${icon('cloud')}SharePoint</span>
<h3>${e(w.name)}</h3><div class="path" title="${e(w.localPath)}">${e(tidy(w.localPath))}</div>
<div class="status"><span class="pill idle loading">Checking SharePoint</span></div><ul class="files"></ul>
<div class="meta"><span>${icon('history')}Pulled <time data-rel datetime="${e(w.lastPulled)}"></time></span><span data-count></span></div>
${mini ? '' : `<div class="actions">
<form method="post" action="/workspace/pull" data-busy="Pulling"><input type="hidden" name="localPath" value="${e(w.localPath)}"><button class="btn btn-quiet btn-sm" data-pull type="submit">${icon('download')}Pull latest</button></form>
<form method="post" action="/workspace/push" data-busy="Uploading" data-upload hidden><input type="hidden" name="localPath" value="${e(w.localPath)}"><button class="btn btn-brass btn-sm" type="submit">${icon('upload')}Upload edits</button></form>
<span class="spacer"></span>
<form method="post" action="/workspace/open"><input type="hidden" name="localPath" value="${e(w.localPath)}"><button class="btn btn-ghost btn-sm" type="submit" title="Open the folder">${icon('folder')}</button></form>
<button class="btn btn-ghost btn-sm" type="button" data-copy="${e(w.localPath)}" title="Copy the folder path for Claude">${icon('copy')}</button>
<a class="btn btn-ghost btn-sm" href="${e(w.webUrl)}" target="_blank" rel="noopener" title="Open in SharePoint">${icon('external')}</a>
<form method="post" action="/workspace/forget"><input type="hidden" name="localPath" value="${e(w.localPath)}"><button class="btn btn-ghost btn-sm" type="submit" title="Remove from this list (deletes nothing)">${icon('x')}</button></form>
</div>`}</article>`;
}

const ACT = {
  sign_in: ['key', 'Signed in to Microsoft'], sign_out: ['logout', 'Signed out'], firm_setup: ['sliders', 'Saved firm settings'],
  admin_consent: ['shield', 'Admin consent'], register: ['link', 'Connected to Claude'], search: ['search', 'Searched SharePoint'],
  read_file: ['eye', 'Read a document'], download: ['download', 'Downloaded a file'], upload: ['upload', 'Uploaded a file'],
  create_folder: ['folder', 'Created a folder'], workspace_pull: ['download', 'Pulled a matter'], workspace_push: ['upload', 'Uploaded edits'],
  restore_version: ['history', 'Restored an earlier version'], mail_search: ['mail', 'Searched email'], mail_read: ['mail', 'Read an email'],
  mail_draft: ['pen', 'Drafted an email'], mail_send: ['mail', 'Sent an email'], graph_request: ['terminal', 'Microsoft Graph request'],
};
function detailOf(a) {
  if (a.error) return a.error;
  if (a.tool === 'workspace_push' && Array.isArray(a.uploaded)) return a.uploaded.length ? a.uploaded.join(', ') : 'Nothing to upload';
  if (a.target && /^https?:/.test(a.target)) {
    try { return decodeURIComponent(new URL(a.target).pathname).split('/').filter(Boolean).slice(-2).join(' / '); } catch { /* fall through */ }
  }
  if (a.query) return `“${a.query}”`;
  if (a.subject) return a.subject;
  if (a.method) return `${a.method} ${a.path}`;
  return a.user || a.target || a.local || a.tenant || '';
}
function event(a, compact = false) {
  const [ic, label] = ACT[a.tool] || ['activity', a.tool];
  return `<li class="ev ${a.error ? 'err' : ''} ${compact ? 'compact' : ''}"><span class="dot">${icon(a.error ? 'alert' : ic)}</span><div><b>${e(a.error ? `${label} failed` : label)}</b><div class="d">${e(detailOf(a))}</div></div><time data-rel datetime="${e(a.ts)}"></time></li>`;
}

function overview(s) {
  const signed = !!s.account;
  const recent = s.workspaces.slice(0, 2);
  return `<section class="view" data-view="overview">
<div><div class="eyebrow" data-today></div>
<h1>${signed ? `<span data-greet>Hello</span>, <em>${e(firstName(s.account))}.</em>` : `Welcome to <em>${e(s.cfg.firmName || 'the bridge')}.</em>`}</h1>
<p class="lede">${signed ? 'Your Microsoft 365 is connected to Claude. Ask in plain English, and Claude works with your files, email and calendar as you.' : 'One more step. Sign in with your Microsoft work account and Claude can start working with your files.'}</p></div>
${bridge(s)}
<div><div class="section-title"><h3>Try asking Claude</h3><span class="muted small">Click a card, then paste it into Claude</span></div>${prompts(s)}</div>
<div class="grid g2" style="margin-top:44px;align-items:start">
<div><div class="section-title" style="margin-top:0"><h3>Matters</h3><a href="#matters">All matters →</a></div>
${recent.length ? `<div class="grid" style="gap:28px">${recent.map((w) => dossier(w, { mini: true })).join('')}</div>` : `<div class="empty" style="padding:36px 20px"><p style="margin:0 0 16px">No matter folders yet. Pull one from SharePoint to work on it with Claude.</p><button class="btn btn-primary" type="button" data-open="add-matter" ${signed ? '' : 'disabled'}>${icon('plus')}Add a matter</button></div>`}</div>
<div><div class="section-title" style="margin-top:0"><h3>Recent activity</h3><a href="#activity">Full log →</a></div>
<div class="card feed">${s.audit.length ? `<ul class="timeline">${s.audit.slice(0, 6).map((a) => event(a, true)).join('')}</ul>` : '<p class="muted">Nothing yet. Everything Claude does through the bridge appears here.</p>'}</div></div>
</div></section>`;
}

const FOLDERS_ART = `<svg class="art" viewBox="0 0 120 96" fill="none" aria-hidden="true"><path d="M14 30a6 6 0 0 1 6-6h22l7 7h51a6 6 0 0 1 6 6v43a6 6 0 0 1-6 6H20a6 6 0 0 1-6-6z" fill="var(--paper-2)" stroke="var(--line-2)" stroke-width="1.5"/><path d="M8 40a6 6 0 0 1 6-6h24l7 7h55a6 6 0 0 1 6 6v37a6 6 0 0 1-6 6H14a6 6 0 0 1-6-6z" fill="var(--card)" stroke="var(--brass)" stroke-width="1.5"/><path d="M36 66h48M36 74h30" stroke="var(--line-2)" stroke-width="3" stroke-linecap="round"/></svg>`;

function matters(s) {
  const signed = !!s.account;
  return `<section class="view" data-view="matters">
<div class="head"><div><div class="eyebrow">Source of truth: SharePoint</div><h2>Matters</h2><p class="lede" style="margin-top:8px">A matter is a working copy of a SharePoint folder on this computer. Claude edits here. SharePoint keeps every version.</p></div>
<button class="btn btn-primary" type="button" data-open="add-matter" ${signed ? '' : 'disabled title="Sign in first"'}>${icon('plus')}Add a matter</button></div>
${s.workspaces.length ? `<div class="dossiers">${s.workspaces.map((w) => dossier(w)).join('')}</div>` : `<div class="empty">${FOLDERS_ART}<h3>No matters yet</h3><p>Paste a SharePoint folder link and the bridge brings the files to this computer, ready for Claude. Your colleagues keep working in SharePoint as usual.</p><button class="btn btn-primary" type="button" data-open="add-matter" ${signed ? '' : 'disabled'}>${icon('plus')}Add your first matter</button></div>`}
<div class="howto"><div><span class="n">1</span><b>Pull</b><p>Bring the latest files down from SharePoint. Your own edits are never overwritten.</p></div>
<div><span class="n">2</span><b>Work with Claude</b><p>Open the folder in Claude Code and ask for drafts, summaries, timelines, redlines.</p></div>
<div><span class="n">3</span><b>Upload</b><p>Edits go back as new SharePoint versions. If a colleague changed the same file, nothing is overwritten.</p></div></div>
</section>`;
}

function access(s) {
  const { cfg } = s;
  const w = cfg.allowWrite ? 0 : 1;
  const tiles = Object.keys(PACKS)
    .map((k) => {
      const on = packOn(cfg, k);
      return `<div class="pack ${on ? 'on' : 'off'}"><span class="ic">${icon(PACK_ICON[k])}</span><div><b>${e(PACKS[k].label.split(' (')[0])}</b><p>${e(on ? PACK_COPY[k][w] : 'Not turned on for this firm.')}</p></div><span class="st">${on ? 'On' : 'Off'}</span></div>`;
    })
    .join('');
  return `<section class="view" data-view="access">
<div class="head"><div><div class="eyebrow">Permissions</div><h2>What Claude can <em>reach</em></h2><p class="lede" style="margin-top:8px">Claude acts as you. It opens only what your Microsoft account can already open, and your firm's walls and matter permissions still apply.</p></div></div>
<div class="rules">
<div class="rule">${icon('eye')}<b>Reads</b><span>Whatever you can open in Microsoft 365.</span></div>
<div class="rule">${icon('pen')}<b>${cfg.allowWrite ? 'Edits, on request' : 'Read only'}</b><span>${cfg.allowWrite ? 'Uploads, drafts and events, when you ask.' : 'Your admin turned changes off.'}</span></div>
<div class="rule">${icon('x')}<b>${cfg.allowDelete ? 'Deletes allowed' : 'Never deletes'}</b><span>${cfg.allowDelete ? 'Only with your explicit OK.' : 'Nothing in Microsoft 365 is ever deleted.'}</span></div>
<div class="rule">${icon('history')}<b>Every version kept</b><span>Each upload is a new SharePoint version you can restore.</span></div></div>
<div class="section-title"><h3>Connected services</h3></div><div class="packs">${tiles}</div>
<div class="vault"><span class="seal">${icon('lock')}</span><div><h3>Sealed on this computer</h3><p>Your Microsoft sign-in is encrypted with AES-256 and the key is held by ${e(s.keyStorage)}. The bridge runs only on this machine. It sends your sign-in to Microsoft and nowhere else, and it never stores a password.</p></div></div>
</section>`;
}

function activity(s) {
  if (!s.audit.length) {
    return `<section class="view" data-view="activity"><div class="head"><div><div class="eyebrow">Audit trail</div><h2>Activity</h2></div></div><div class="empty"><h3>Nothing yet</h3><p>Every search, read, pull and upload Claude makes through the bridge is recorded here, on this computer.</p></div></section>`;
  }
  const groups = [];
  for (const a of s.audit) {
    const day = a.ts.slice(0, 10);
    if (!groups.length || groups[groups.length - 1].day !== day) groups.push({ day, items: [] });
    groups[groups.length - 1].items.push(a);
  }
  return `<section class="view" data-view="activity">
<div class="head"><div><div class="eyebrow">Audit trail</div><h2>Activity</h2><p class="lede" style="margin-top:8px">Everything Claude did through the bridge on this computer. File names, never file contents.</p></div></div>
<div class="card" style="padding:8px 26px 18px">${groups.map((g) => `<div class="day" data-day="${e(g.items[0].ts)}">${e(g.day)}</div><ul class="timeline">${g.items.map((a) => event(a)).join('')}</ul>`).join('')}</div>
</section>`;
}

function admin(s) {
  const { cfg } = s;
  return `<section class="view" data-view="admin">
<div class="head"><div><div class="eyebrow">For the IT admin</div><h2>Firm setup</h2><p class="lede" style="margin-top:8px">One Entra app for the whole firm. Each person signs in as themselves. No secrets on laptops.</p></div></div>
<div class="split">
<div class="grid" style="gap:16px">
<div class="record"><div class="eyebrow">Setup code</div><h3>Send this to your team</h3><p>It carries the firm's tenant, app and permission choices. It contains no secret, so it is safe to email.</p>
<div class="code"><output>${e(s.setupCode)}</output><button type="button" class="btn btn-brass" data-copy="${e(s.setupCode)}">${icon('copy')}<span class="lbl">Copy</span></button></div></div>
<div class="card"><div class="eyebrow">Once per firm</div><h3 style="margin:6px 0 6px">Grant admin consent</h3><p class="muted" style="margin:0 0 16px">A Global Administrator approves these permissions for everyone in one step. Without it, each person sees a consent prompt or is blocked.</p>
<a class="btn btn-primary" href="${e(adminConsentUrl(cfg, REDIRECT_URI))}">${icon('shield')}Grant admin consent</a>
<details class="more"><summary>${scopesFor(cfg).length} Microsoft Graph permissions, all delegated</summary><div class="chips">${scopesFor(cfg).map((x) => `<code>${e(x)}</code>`).join('')}</div></details></div>
<div class="card"><div class="eyebrow">Settings</div><h3 style="margin:6px 0 18px">Firm and permissions</h3>${firmForm(cfg)}</div>
</div>
<div class="card"><div class="eyebrow">Guide</div><h3 style="margin:6px 0 4px">Create the Entra app</h3><p class="muted" style="margin:0 0 6px">About five minutes, once.</p>${entraSteps()}</div>
</div></section>`;
}

function addMatterDialog(s) {
  return `<dialog id="add-matter"><div class="dlg"><button class="btn btn-ghost btn-sm close" type="button" data-close aria-label="Close">${icon('x')}</button>
<div class="eyebrow">New matter</div><h2>Bring a folder from SharePoint</h2>
<p class="muted" style="margin:0 0 20px">In SharePoint, open the matter folder and copy the address from the browser bar, or use <b>Copy link</b>.</p>
<form class="form" method="post" action="/workspace/add" data-busy="Pulling files">
<div class="field"><label for="src">SharePoint or OneDrive folder link</label><input id="src" class="input" name="source" placeholder="https://firm.sharepoint.com/sites/Litigation/Shared Documents/Henderson" required></div>
<div class="field"><label for="nm">Name on this computer <span class="muted">(optional)</span></label><input id="nm" class="input" name="name" placeholder="Henderson v. Acme"><div class="hint">Saved under ${e(tidy(s.workspacesRoot))}</div></div>
<div style="display:flex;gap:10px;justify-content:flex-end"><button class="btn btn-quiet" type="button" data-close>Cancel</button><button class="btn btn-primary" type="submit">${icon('download')}Pull folder</button></div></form></div></dialog>`;
}

export function renderPage(s, flash) {
  if (!s.configured) return shell('M365 Bridge · Welcome', onboarding(s), flash);
  return shell(
    `M365 Bridge · ${s.cfg.firmName || 'Dashboard'}`,
    `<div class="app">${rail(s)}<main class="main">${overview(s)}${matters(s)}${access(s)}${activity(s)}${admin(s)}</main></div>${addMatterDialog(s)}`,
    flash,
  );
}
