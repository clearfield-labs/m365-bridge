// A workspace is a local working copy of a SharePoint or OneDrive folder, used the way
// developers use a git clone. SharePoint stays the source of truth.
//
//   pull    bring down what changed in SharePoint. Never overwrites your edits.
//   lock    check a file out in SharePoint before editing it. Colleagues see
//           "Checked out to <you>" and cannot save it until you push or release.
//   push    upload your edits and check them in with a note. The note shows in
//           SharePoint version history, like a commit message.
//   merge   if SharePoint has a newer version than the one you edited, Claude merges
//           the two into your file. There is never a second copy in the folder.
//   move    rename or move a file or folder in SharePoint and locally together.
//
// Nothing is ever deleted in SharePoint by this tool.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import {
  resolveItem, listChildren, downloadItem, uploadBytes, createFolder, graph, enc, personOf,
  getItem, checkout, checkin, discardCheckout, moveItem, contentTag, holderOf, SELECT,
} from './graph.js';
import { loadConfig, updateConfig } from './config.js';
import { WORKSPACES_ROOT } from './paths.js';
import { audit } from './audit.js';

const META_DIR = '.m365';
const MANIFEST = path.join(META_DIR, 'manifest.json');
const BASE_DIR = path.join(META_DIR, 'base'); // each file as it was at the last sync
const MERGE_DIR = path.join(META_DIR, 'merge'); // SharePoint's newer version, while a merge is open
const MAX_FILES = Number(process.env.M365_MAX_FILES || 3000);
const MAX_FILE_BYTES = Number(process.env.M365_MAX_FILE_MB || 250) * 1024 * 1024;
const LEGACY_COPY_TAG = ' (SharePoint copy ';

const IGNORED_DIRS = new Set([META_DIR, '.claude', '.git', 'node_modules']);
export function isIgnored(rel) {
  const parts = rel.split('/');
  if (parts.some((p) => IGNORED_DIRS.has(p) || p.startsWith('.') || p === '__pycache__')) return true;
  const base = parts[parts.length - 1];
  return (
    base === '.DS_Store' || base === 'Thumbs.db' || base === 'desktop.ini' ||
    base.startsWith('~$') || base.endsWith('.tmp') || base.endsWith('.pyc') ||
    base.includes(LEGACY_COPY_TAG)
  );
}

const sha256 = (buf) => crypto.createHash('sha256').update(buf).digest('hex');
// Background saving checks folders often, so remember each file's hash until it changes on disk.
const hashCache = new Map();
function hashFile(file) {
  const st = fs.statSync(file);
  const hit = hashCache.get(file);
  if (hit && hit.mtimeMs === st.mtimeMs && hit.size === st.size) return hit.hash;
  const hash = sha256(fs.readFileSync(file));
  // A file changed within the last two seconds may change again within the same clock tick.
  if (Date.now() - st.mtimeMs > 2000) hashCache.set(file, { mtimeMs: st.mtimeMs, size: st.size, hash });
  return hash;
}

// One operation at a time per folder in this process (background saving, Claude's tools, the page).
const queues = new Map();
function serial(localPath, fn) {
  const key = localPath ? path.resolve(localPath) : '';
  const prev = queues.get(key) || Promise.resolve();
  const run = prev.then(fn, fn);
  const tail = run.catch(() => {});
  queues.set(key, tail);
  tail.then(() => { if (queues.get(key) === tail) queues.delete(key); });
  return run;
}
const toPosix = (p) => p.split(path.sep).join('/');
const clean = (rel) => toPosix(String(rel)).replace(/^\.\//, '').replace(/^\/+/, '');
const abs = (root, rel, dir = '') => path.join(root, dir, ...rel.split('/'));
const now = () => new Date().toISOString();
// Compare content, not metadata. Prefer quickXorHash when both sides have it; SharePoint can add
// the hash a moment after an upload, so fall back to the content tag (cTag) rather than flag a change.
const qxOf = (item) => item?.file?.hashes?.quickXorHash || null;
function sameContent(item, rec) {
  if (!rec) return false;
  if (rec.qx !== undefined) {
    const qx = qxOf(item);
    if (qx && rec.qx) return qx === rec.qx;
    return (item.cTag || item.eTag) === (rec.cTag || rec.eTag);
  }
  return contentTag(item) === (rec.tag || rec.cTag || rec.eTag); // manifests from before 0.7
}
const remoteChanged = (item, known) => !sameContent(item, known);
const isCheckedOut = (item) => item?.publication?.level === 'checkout';

export function safeName(name) {
  return String(name).replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').trim() || 'Workspace';
}

function readManifest(root) {
  const file = path.join(root, MANIFEST);
  if (!fs.existsSync(file)) {
    throw new Error(`${root} is not a synced folder yet. Run workspace_pull with a SharePoint or OneDrive folder link first.`);
  }
  const m = JSON.parse(fs.readFileSync(file, 'utf8'));
  Object.defineProperty(m, '_snap', { value: snapshot(m), enumerable: false, writable: true });
  return m;
}
const snapshot = (m) => Object.fromEntries(Object.entries(m.files).map(([k, v]) => [k, JSON.stringify(v)]));

/**
 * Save the manifest. The background service, Claude's tools and the edit hook can all work on the
 * same folder at once, so re-read what is on disk and apply only the entries this operation
 * changed. Written atomically, so a reader never sees half a file.
 */
function writeManifest(root, m) {
  const file = path.join(root, MANIFEST);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  let out = m;
  if (m._snap && fs.existsSync(file)) {
    let disk = null;
    try { disk = JSON.parse(fs.readFileSync(file, 'utf8')); } catch { /* unreadable: ours wins */ }
    if (disk?.files) {
      for (const k of new Set([...Object.keys(m._snap), ...Object.keys(m.files)])) {
        const mine = m.files[k] === undefined ? undefined : JSON.stringify(m.files[k]);
        if (mine === m._snap[k]) continue; // not changed by this operation
        if (mine === undefined) delete disk.files[k];
        else disk.files[k] = m.files[k];
      }
      disk.source = m.source;
      disk.version = m.version;
      if (m.pulledAt) disk.pulledAt = m.pulledAt;
      out = disk;
    }
  }
  const tmp = `${file}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(out, null, 2));
  for (let i = 0; ; i++) {
    try { fs.renameSync(tmp, file); break; } catch (e) {
      if (i >= 20 || !['EPERM', 'EBUSY', 'EACCES'].includes(e.code)) { fs.rmSync(tmp, { force: true }); throw e; }
      Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 50);
    }
  }
  m.files = out.files;
  m._snap = snapshot(out);
}

function writeAt(file, buf) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, buf);
}
const writeLocal = (root, rel, buf) => writeAt(abs(root, rel), buf);
const writeBase = (root, rel, buf) => writeAt(abs(root, rel, BASE_DIR), buf);
const dropMergeCopy = (root, rel) => fs.rmSync(abs(root, rel, MERGE_DIR), { force: true });

function entryFor(item, hash) {
  return {
    itemId: item.id,
    eTag: item.eTag,
    cTag: item.cTag,
    tag: contentTag(item),
    qx: qxOf(item),
    size: item.size,
    remoteModified: item.lastModifiedDateTime,
    hash,
  };
}

/** Walk the remote folder tree. Returns Map(relPath -> driveItem) for files. */
async function walkRemote(driveId, rootId) {
  const files = new Map();
  const skipped = [];
  const skippedRels = new Set(); // present in SharePoint but not synced (OneNote, too large)
  const queue = [['', rootId]];
  while (queue.length) {
    const [prefix, id] = queue.shift();
    for (const item of await listChildren(driveId, id)) {
      const rel = prefix ? `${prefix}/${item.name}` : item.name;
      if (item.folder) queue.push([rel, item.id]);
      else if (item.package || /\.(one|onetoc2|onepkg)$/i.test(item.name)) { skipped.push(`${rel} (OneNote, open it in OneNote)`); skippedRels.add(rel); }
      else if (item.size > MAX_FILE_BYTES) { skipped.push(`${rel} (larger than ${MAX_FILE_BYTES / 1024 / 1024} MB)`); skippedRels.add(rel); }
      else if (!isIgnored(rel)) files.set(rel, item);
      if (files.size >= MAX_FILES) {
        skipped.push(`stopped after ${MAX_FILES} files; pick a smaller folder`);
        return { files, skipped, skippedRels, complete: false };
      }
    }
  }
  return { files, skipped, skippedRels, complete: true };
}

function walkLocal(root) {
  const out = [];
  const visit = (dir) => {
    for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, ent.name);
      const rel = toPosix(path.relative(root, full));
      if (isIgnored(rel)) continue;
      if (ent.isDirectory()) visit(full);
      else if (ent.isFile()) out.push(rel);
    }
  };
  if (fs.existsSync(root)) visit(root);
  return out;
}

const INSTRUCTIONS_VERSION = 'v4';
function writeClaudeInstructions(root, source, firmName, autoSave) {
  const mark = `<!-- m365-bridge:instructions ${INSTRUCTIONS_VERSION} ${autoSave ? 'auto' : 'manual'} -->`;
  const file = path.join(root, '.claude', 'CLAUDE.md');
  if (fs.existsSync(file)) {
    const cur = fs.readFileSync(file, 'utf8');
    // Already current, or written by a person: leave it. Replace only our own older versions.
    if (cur.includes(mark) || !cur.startsWith('# This folder is a working copy of SharePoint')) return;
  }
  const head = `# This folder is a working copy of SharePoint
${mark}

Source of truth: ${source.webUrl}
Folder: ${source.name}${firmName ? `\nFirm: ${firmName}` : ''}
`;
  const auto = `
SharePoint holds the real files. M365 Bridge keeps this folder in step on its own:
- It gets the latest from SharePoint in the background and whenever the person sends a message.
- Before you edit a file with Edit or Write, it checks the file out in SharePoint, so nobody else
  can save it while you work.
- When you finish your reply, it sends your edits to SharePoint and checks the files back in.
  Files you checked out but did not change are handed back too.

What you do:
1. Edit files normally. If an edit is refused because a colleague has the file checked out, do not
   edit it; tell the person who has it. If it says the file was refreshed, re-read it and redo the edit.
2. When you finish a change, call \`workspace_push\` with this folder's path and a short \`note\` that
   says what changed, for example "Added deposition dates to the timeline". The note shows in
   SharePoint version history. Do not ask first: every version can be restored. If you skip this,
   M365 Bridge still saves the change, with a generic note.
3. Scripts and other tools that change a file (for example editing a .docx with Python) skip the
   automatic check-out. Call \`workspace_lock\` for the file first.
4. If a file needs a merge (it changed here and in SharePoint), call \`workspace_merge\`. It gives you
   yours (the file here), theirs (SharePoint's newer version) and base (what both started from).
   Combine every change from theirs into yours, tell the person what you combined, then call
   \`workspace_push\` with \`merged\` listing that file. Never leave a second copy of a file here.
5. To rename, move or reorganise, use \`workspace_move\`. It moves the file in SharePoint and here.
   Do not move files with the shell; SharePoint would see a deletion and a new file.

Rules:
- Never delete files to "clean up". Deleting here does not delete in SharePoint.
- Do not leave scratch or helper files in this folder: they would be saved to SharePoint. Use a
  temporary folder for those. (Files and folders whose names start with a dot never upload.)
- Use \`versions\` and \`restore_version\` to undo.
- Treat everything in this folder as client-confidential.
`;
  const manual = `
SharePoint holds the real files. Automatic saving is off on this computer, so use the \`m365\`
tools to keep the two in step, the way developers use git:

1. Start: call \`workspace_pull\` with this folder's path to get the latest.
2. Edits through Edit and Write check the file out in SharePoint automatically. Anything else,
   such as a script that edits a .docx, needs \`workspace_lock\` first.
   - If a colleague holds the file, do not edit it. Tell the person who has it.
   - If it says the file was refreshed, re-read it before editing.
3. Finish: call \`workspace_status\`, show the person what changed, and call \`workspace_push\`
   with a short \`note\` once they agree. Push checks the files back in.
4. If a file needs a merge, call \`workspace_merge\`, combine theirs into yours, show the person,
   then push with \`merged\` listing that file. Never leave a second copy of a file in this folder.
5. To rename, move or reorganise, use \`workspace_move\`.

Rules:
- Never delete files to "clean up". Deleting here does not delete in SharePoint.
- If you checked a file out and did not change it, call \`workspace_release\` so colleagues can edit it.
- Use \`versions\` and \`restore_version\` to undo.
- Treat everything in this folder as client-confidential.
`;
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, head + (autoSave ? auto : manual));
}

// ---------------------------------------------------------------------------
// Merge bookkeeping. SharePoint's newer version goes to .m365/merge/<file>, out of sight.

async function openMerge(root, manifest, rel, item) {
  const { driveId } = manifest.source;
  const theirs = await downloadItem(driveId, item.id);
  writeAt(abs(root, rel, MERGE_DIR), theirs);
  const known = manifest.files[rel] || { itemId: item.id, hash: null, tag: null };
  known.merge = { eTag: item.eTag, cTag: item.cTag, qx: qxOf(item), tag: contentTag(item), by: personOf(item) || undefined, at: now() };
  manifest.files[rel] = known;
  return { file: rel, changed_by: known.merge.by, next: 'Call workspace_merge for this file.' };
}

// ---------------------------------------------------------------------------
// Locks

/**
 * Check one tracked file out to the signed-in person.
 * Returns { status, ... } where status is one of:
 *   locked, yours, held, refreshed, needs_merge, unsupported, untracked
 */
async function lockOne(root, manifest, rel) {
  const known = manifest.files[rel];
  if (!known) return { status: 'untracked' };
  if (known.lock) return { status: 'yours' };
  const { driveId } = manifest.source;
  let real = true;
  try {
    await checkout(driveId, known.itemId);
  } catch (e) {
    if (e.status === 423) return { status: 'held', by: holderOf(e.raw) || 'someone else' };
    if (e.status === 401 || e.status === 403 || e.status === 404 || !e.status) throw e;
    // Check-out refused for another reason. If the file is checked out, it is someone else's
    // (we cannot prove it is ours), so never treat it as our lock. Otherwise the library does
    // not support check-out: fall back to version checks at push.
    const it = await getItem(driveId, known.itemId);
    if (isCheckedOut(it)) return { status: 'held', by: holderOf(e.raw) || 'someone else' };
    real = false;
  }
  known.lock = { at: now(), real }; // recorded at once, so a later failure cannot orphan it
  const item = await getItem(driveId, known.itemId);
  if (!remoteChanged(item, known)) {
    known.eTag = item.eTag; // check-out changes the eTag but not the content
    return { status: real ? 'locked' : 'unsupported' };
  }
  const full = abs(root, rel);
  const localHash = fs.existsSync(full) ? hashFile(full) : null;
  if (localHash === known.hash || localHash === null) {
    const buf = await downloadItem(driveId, item.id);
    writeLocal(root, rel, buf);
    writeBase(root, rel, buf);
    manifest.files[rel] = { ...entryFor(item, sha256(buf)), lock: known.lock };
    return { status: 'refreshed', by: personOf(item) || undefined };
  }
  const m = await openMerge(root, manifest, rel, item);
  return { status: 'needs_merge', by: m.changed_by };
}

async function lockImpl({ localPath, files }) {
  if (!loadConfig().allowWrite) throw new Error('This firm set M365 Bridge to read-only.');
  const root = path.resolve(localPath);
  const manifest = readManifest(root);
  const report = { workspace: root, locked: [], already_yours: [], held_by_others: [], refreshed_first: [], needs_merge: [], no_lock_needed: [], lock_not_supported: [], errors: [] };
  try {
  for (const raw of files || []) {
    const rel = clean(raw);
    let r;
    try { r = await lockOne(root, manifest, rel); } catch (e) { report.errors.push({ file: rel, error: e.message }); continue; }
    if (r.status === 'locked') report.locked.push(rel);
    else if (r.status === 'yours') report.already_yours.push(rel);
    else if (r.status === 'held') report.held_by_others.push({ file: rel, checked_out_to: r.by });
    else if (r.status === 'refreshed') report.refreshed_first.push({ file: rel, changed_by: r.by, next: 'Re-read this file before editing. It now has the latest version from SharePoint.' });
    else if (r.status === 'needs_merge') report.needs_merge.push({ file: rel, changed_by: r.by, next: 'Call workspace_merge for this file.' });
    else if (r.status === 'unsupported') report.lock_not_supported.push(rel);
    else report.no_lock_needed.push(rel);
  }
  } finally {
    writeManifest(root, manifest);
  }
  audit({ tool: 'workspace_lock', target: manifest.source.webUrl, locked: report.locked, held: report.held_by_others.length });
  return report;
}

async function releaseImpl({ localPath, files }) {
  const root = path.resolve(localPath);
  const manifest = readManifest(root);
  const only = files?.length ? new Set(files.map(clean)) : null;
  const report = { workspace: root, released: [], edits_kept_locally: [], errors: [] };
  for (const [rel, known] of Object.entries(manifest.files)) {
    if (!known.lock || (only && !only.has(rel))) continue;
    try {
      if (known.lock.real && isCheckedOut(await getItem(manifest.source.driveId, known.itemId))) {
        await discardCheckout(manifest.source.driveId, known.itemId).catch((e) => { if (e.status !== 400) throw e; });
      }
      delete known.lock;
      report.released.push(rel);
      const full = abs(root, rel);
      if (fs.existsSync(full) && hashFile(full) !== known.hash) report.edits_kept_locally.push(rel);
    } catch (e) {
      report.errors.push({ file: rel, error: e.message });
    }
  }
  writeManifest(root, manifest);
  audit({ tool: 'workspace_release', target: manifest.source.webUrl, released: report.released });
  return report;
}

// ---------------------------------------------------------------------------
// Pull

async function pullImpl({ source, localPath, name, auto = false } = {}) {
  let root = localPath ? path.resolve(localPath) : null;
  let manifest = null;
  if (root && fs.existsSync(path.join(root, MANIFEST))) manifest = readManifest(root);

  if (!manifest) {
    if (!source) throw new Error('Give a SharePoint or OneDrive folder link (source) for a new synced folder.');
    const folder = await resolveItem(source);
    if (!folder.folder && !folder.root) throw new Error(`"${folder.name}" is a file. Point workspace_pull at a folder.`);
    const driveId = folder.parentReference?.driveId;
    const label = folder.root ? name || (folder.webUrl?.includes('-my.sharepoint.com') ? 'OneDrive' : 'Documents') : folder.name;
    root = root || path.join(WORKSPACES_ROOT, safeName(name || label));
    if (fs.existsSync(path.join(root, MANIFEST))) {
      manifest = readManifest(root);
    } else {
      if (fs.existsSync(root) && fs.readdirSync(root).length) {
        throw new Error(`${root} already has files and is not a synced folder. Choose an empty folder or a new name.`);
      }
      manifest = { version: 3, source: { driveId, itemId: folder.id, name: label, webUrl: folder.webUrl }, files: {} };
    }
  }
  const { driveId, itemId: rootId } = manifest.source;
  fs.mkdirSync(root, { recursive: true });

  const { files: remote, skipped, skippedRels, complete } = await walkRemote(driveId, rootId);
  const report = {
    workspace: root, source: manifest.source.webUrl,
    downloaded: [], updated: [], unchanged: 0, kept_local_edits: [], needs_merge: [],
    removed_in_sharepoint: [], held_by_others: [], skipped, errors: [],
  };
  // A fresh entry for a file, keeping a lock this person holds.
  const keep = (rel, item, hash) => {
    const prev = manifest.files[rel];
    manifest.files[rel] = { ...entryFor(item, hash), ...(prev?.lock ? { lock: prev.lock } : {}) };
    dropMergeCopy(root, rel);
  };

  // New files: download a few at a time (each Microsoft request can take a second or two).
  const fresh = [...remote].filter(([rel]) => !fs.existsSync(abs(root, rel)));
  await pool(fresh, 4, async ([rel, item]) => {
    try {
      const buf = await downloadItem(driveId, item.id);
      writeLocal(root, rel, buf);
      writeBase(root, rel, buf);
      keep(rel, item, sha256(buf));
      report.downloaded.push(rel);
    } catch (e) {
      report.errors.push({ file: rel, error: e.message });
    }
  });
  const freshSet = new Set(fresh.map(([rel]) => rel));

  for (const [rel, item] of remote) {
    const full = abs(root, rel);
    const known = manifest.files[rel];
    if (isCheckedOut(item) && !known?.lock) report.held_by_others.push({ file: rel });
    if (freshSet.has(rel)) continue;
    if (known?.lock?.real) { report.unchanged++; continue; } // checked out to you: nobody else can change it
    try {
    const localHash = hashFile(full);
    // A merge left open with no local edits: just take SharePoint's version.
    if (known && (!known.merge || localHash === known.hash)) {
      const changedThere = remoteChanged(item, known);
      const changedHere = localHash !== known.hash;
      if (!changedThere) {
        report.unchanged++;
        if (changedHere) report.kept_local_edits.push(rel);
        continue;
      }
      if (!changedHere) {
        const buf = await downloadItem(driveId, item.id);
        writeLocal(root, rel, buf);
        writeBase(root, rel, buf);
        keep(rel, item, sha256(buf));
        report.updated.push({ file: rel, by: personOf(item) || undefined });
        continue;
      }
    }
    if (known?.merge && sameContent(item, known.merge)) {
      report.needs_merge.push({ file: rel, changed_by: known.merge.by, next: 'Call workspace_merge for this file.' });
      continue;
    }
    // Both sides changed (or an untracked local file has the same name).
    const buf = await downloadItem(driveId, item.id);
    if (sha256(buf) === localHash) {
      writeBase(root, rel, buf);
      keep(rel, item, localHash);
      report.unchanged++;
      continue;
    }
    report.needs_merge.push(await openMerge(root, manifest, rel, item));
    } catch (e) {
      // For example a file open in Word on Windows. Report it and carry on with the rest.
      report.errors.push({ file: rel, error: /EBUSY|EPERM/.test(e.code || e.message) ? 'The file is open in another app. Close it and get latest again.' : e.message });
    }
  }

  // Files that SharePoint no longer has were moved or deleted there (it keeps them in its
  // recycle bin). Drop untouched local copies; keep any you edited.
  // Only when the listing was complete: a cut-off listing must never look like a deletion.
  const remoteLower = new Set([...remote.keys()].map((k) => k.toLowerCase()));
  for (const [rel, known] of Object.entries(complete ? manifest.files : {})) {
    if (remote.has(rel) || skippedRels.has(rel)) continue;
    const full = abs(root, rel);
    if (remoteLower.has(rel.toLowerCase())) {
      // Renamed in SharePoint by letter case only. On Mac and Windows it is the same file on disk.
      delete manifest.files[rel];
      continue;
    }
    const edited = fs.existsSync(full) && hashFile(full) !== known.hash;
    if (edited || known.lock) { report.kept_local_edits.push(rel); continue; }
    fs.rmSync(full, { force: true });
    fs.rmSync(abs(root, rel, BASE_DIR), { force: true });
    delete manifest.files[rel];
    report.removed_in_sharepoint.push(rel);
  }
  removeEmptyDirs(root);

  manifest.version = 3;
  manifest.pulledAt = now();
  writeManifest(root, manifest);
  { const c = loadConfig(); writeClaudeInstructions(root, manifest.source, c.firmName, c.autoSave !== false); }
  rememberWorkspace(root, manifest.source);
  const changed = report.downloaded.length + report.updated.length + report.needs_merge.length + report.removed_in_sharepoint.length;
  if (!auto || changed || report.errors.length) {
    audit({ tool: 'workspace_pull', target: manifest.source.webUrl, local: root, downloaded: report.downloaded.length + report.updated.length, needs_merge: report.needs_merge.length, auto: auto || undefined });
  }
  return report;
}

async function pool(items, size, fn) {
  let i = 0;
  const workers = Array.from({ length: Math.min(size, items.length) }, async () => {
    while (i < items.length) await fn(items[i++]);
  });
  await Promise.all(workers);
}

function removeEmptyDirs(root) {
  const visit = (dir) => {
    for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
      if (!ent.isDirectory() || IGNORED_DIRS.has(ent.name)) continue;
      const full = path.join(dir, ent.name);
      visit(full);
      if (!fs.readdirSync(full).length) fs.rmdirSync(full);
    }
  };
  if (fs.existsSync(root)) visit(root);
}

// ---------------------------------------------------------------------------
// Status

export async function status({ localPath, checkRemote = true }) {
  const root = path.resolve(localPath);
  const manifest = readManifest(root);
  const local = walkLocal(root);
  const localSet = new Set(local);
  const report = {
    workspace: root, source: manifest.source.webUrl, last_pull: manifest.pulledAt,
    modified: [], new_files: [], deleted_locally: [], locked_by_you: [], needs_merge: [],
    changed_in_sharepoint: [], new_in_sharepoint: [], held_by_others: [],
  };
  for (const rel of local) {
    const known = manifest.files[rel];
    if (!known) report.new_files.push(rel);
    else if (hashFile(abs(root, rel)) !== known.hash) report.modified.push(rel);
  }
  for (const [rel, known] of Object.entries(manifest.files)) {
    if (!localSet.has(rel)) report.deleted_locally.push(rel);
    if (known.lock) report.locked_by_you.push(rel);
    if (known.merge) report.needs_merge.push({ file: rel, changed_by: known.merge.by });
  }
  if (checkRemote) {
    const { files: remote } = await walkRemote(manifest.source.driveId, manifest.source.itemId);
    for (const [rel, item] of remote) {
      const known = manifest.files[rel];
      if (isCheckedOut(item) && !known?.lock) report.held_by_others.push({ file: rel });
      if (!known) report.new_in_sharepoint.push(rel);
      else if (!known.lock?.real && !known.merge && remoteChanged(item, known)) report.changed_in_sharepoint.push({ file: rel, by: personOf(item) || undefined });
    }
  }
  report.tracked_files = Object.keys(manifest.files).length;
  report.summary = summarizeStatus(report);
  return report;
}

function summarizeStatus(r) {
  const bits = [];
  if (r.needs_merge.length) bits.push(`${r.needs_merge.length} need a merge`);
  if (r.modified.length) bits.push(`${r.modified.length} edited here`);
  if (r.new_files.length) bits.push(`${r.new_files.length} new here`);
  if (r.locked_by_you.length) bits.push(`${r.locked_by_you.length} checked out to you`);
  if (r.changed_in_sharepoint.length) bits.push(`${r.changed_in_sharepoint.length} changed in SharePoint (pull first)`);
  if (r.new_in_sharepoint.length) bits.push(`${r.new_in_sharepoint.length} new in SharePoint`);
  if (r.held_by_others.length) bits.push(`${r.held_by_others.length} checked out to colleagues`);
  if (r.deleted_locally.length) bits.push(`${r.deleted_locally.length} deleted here (SharePoint keeps them)`);
  return bits.length ? bits.join(', ') : 'Everything matches SharePoint.';
}

// ---------------------------------------------------------------------------
// Merge

async function mergeImpl({ localPath, file }) {
  const root = path.resolve(localPath);
  const manifest = readManifest(root);
  const rel = clean(file);
  const known = manifest.files[rel];
  if (!known) throw new Error(`${rel} is not a SharePoint file in this folder.`);
  const { driveId } = manifest.source;
  const item = await getItem(driveId, known.itemId);
  if (!known.merge || !sameContent(item, known.merge)) {
    const full = abs(root, rel);
    const changedHere = fs.existsSync(full) && hashFile(full) !== known.hash;
    if (!known.merge && !(remoteChanged(item, known) && changedHere)) {
      throw new Error(`${rel} does not need a merge. Push it, or pull to get the latest.`);
    }
    await openMerge(root, manifest, rel, item);
  }
  // Hold the file while merging so it cannot change again underneath you.
  let lockNote;
  if (!known.lock) {
    try {
      await checkout(driveId, known.itemId);
      known.lock = { at: now(), real: true };
    } catch (e) {
      if (e.status === 423) lockNote = `Checked out to ${holderOf(e.raw) || 'someone else'} right now. Merge, but push only after they check it in.`;
      else known.lock = { at: now(), real: false };
    }
  }
  writeManifest(root, manifest);
  const base = abs(root, rel, BASE_DIR);
  audit({ tool: 'workspace_merge', target: manifest.source.webUrl, file: rel });
  return {
    file: rel,
    yours: abs(root, rel),
    theirs: abs(root, rel, MERGE_DIR),
    base: fs.existsSync(base) ? base : null,
    changed_by: known.merge.by,
    note: lockNote,
    how: 'Write the merged result into "yours" (the file in the folder). Keep every change from "theirs" plus every change the person made, using "base" to tell who changed what. For .docx files, edit the document so formatting survives. Show the person the result, then call workspace_push with merged: ["' + rel + '"]. Do not save the merged text anywhere else.',
  };
}

// ---------------------------------------------------------------------------
// Push

async function ensureRemoteDir(driveId, rootId, relDir, cache) {
  if (!relDir) return rootId;
  if (cache.has(relDir)) return cache.get(relDir);
  const parentRel = relDir.includes('/') ? relDir.slice(0, relDir.lastIndexOf('/')) : '';
  const parentId = await ensureRemoteDir(driveId, rootId, parentRel, cache);
  const folder = await createFolder(driveId, parentId, relDir.split('/').pop());
  cache.set(relDir, folder.id);
  return folder.id;
}

async function pushImpl({ localPath, files, note, merged = [], via = true, auto = false }) {
  const cfg = loadConfig();
  if (!cfg.allowWrite) throw new Error('This firm set M365 Bridge to read-only. Ask your admin to allow uploads.');
  const root = path.resolve(localPath);
  const manifest = readManifest(root);
  const { driveId, itemId: rootId } = manifest.source;
  const st = await status({ localPath: root, checkRemote: false });
  const only = files?.length ? new Set(files.map(clean)) : null;
  const mergedSet = new Set(merged.map(clean));
  const inScope = (f) => !only || only.has(f);
  const plain = (note || 'Updated').trim().replace(/\s*\(via Claude\)$/i, '');
  const comment = via ? `${plain} (via Claude)` : plain;
  const report = { workspace: root, note: comment, uploaded: [], created: [], released: [], held_by_others: [], needs_merge: [], errors: [], not_pushed_deletions: st.deleted_locally };
  const dirCache = new Map();

  for (const rel of st.modified.filter(inScope)) {
    const known = manifest.files[rel];
    if (known.merge && !mergedSet.has(rel)) {
      report.needs_merge.push({ file: rel, changed_by: known.merge.by, next: 'Call workspace_merge, merge, then push with merged.' });
      continue;
    }
    try {
      const r = await lockOne(root, manifest, rel);
      if (r.status === 'held') { report.held_by_others.push({ file: rel, checked_out_to: r.by }); continue; }
      if (r.status === 'needs_merge') { report.needs_merge.push({ file: rel, changed_by: r.by, next: 'Call workspace_merge for this file.' }); continue; }
      const item = await getItem(driveId, known.itemId);
      if (!sameContent(item, known.merge || known)) {
        report.needs_merge.push(await openMerge(root, manifest, rel, item));
        continue;
      }
      const buf = fs.readFileSync(abs(root, rel));
      try {
        await uploadBytes(driveId, buf, { itemId: known.itemId, ifMatch: item.eTag });
      } catch (e) {
        if (e.status === 412) { report.needs_merge.push(await openMerge(root, manifest, rel, await getItem(driveId, known.itemId))); continue; }
        if (e.status === 423) { report.held_by_others.push({ file: rel, checked_out_to: holderOf(e.raw) || 'someone else' }); continue; }
        throw e;
      }
      // Record the upload first, so a check-in failure cannot make it look like someone else's change.
      const uploaded = await getItem(driveId, known.itemId);
      const lockRec = known.lock;
      manifest.files[rel] = { ...entryFor(uploaded, sha256(buf)), ...(lockRec ? { lock: lockRec } : {}) };
      writeBase(root, rel, buf);
      dropMergeCopy(root, rel);
      if (lockRec?.real) {
        try {
          await checkin(driveId, known.itemId, comment);
          delete manifest.files[rel].lock;
        } catch (e) {
          report.errors.push({ file: rel, error: `Uploaded, but SharePoint did not check it back in (${e.message}). It is still checked out to you; colleagues will not see the new version until it is checked in.` });
          continue;
        }
      } else {
        delete manifest.files[rel].lock;
      }
      const fresh = await getItem(driveId, known.itemId);
      manifest.files[rel] = { ...entryFor(fresh, sha256(buf)), ...(manifest.files[rel].lock ? { lock: manifest.files[rel].lock } : {}) };
      report.uploaded.push({ file: rel, url: fresh.webUrl });
    } catch (e) {
      report.errors.push({ file: rel, error: e.message });
    }
  }

  for (const rel of st.new_files.filter(inScope)) {
    const buf = fs.readFileSync(abs(root, rel));
    try {
      const relDir = rel.includes('/') ? rel.slice(0, rel.lastIndexOf('/')) : '';
      const parentId = await ensureRemoteDir(driveId, rootId, relDir, dirCache);
      let item;
      try {
        item = await uploadBytes(driveId, buf, { parentId, name: rel.split('/').pop() });
      } catch (e) {
        if (e.status !== 409) throw e;
        const existing = await graph(`/drives/${driveId}/items/${parentId}:/${encodeURIComponent(rel.split('/').pop())}?$select=${SELECT}`);
        report.needs_merge.push(await openMerge(root, manifest, rel, existing));
        continue;
      }
      // Libraries that require check-out leave new files checked out (invisible to others).
      let stillOut = null;
      await checkin(driveId, item.id, comment).catch((e) => { if (e.status !== 400) stillOut = e; }); // 400 = was not checked out
      const fresh = await getItem(driveId, item.id);
      manifest.files[rel] = { ...entryFor(fresh, sha256(buf)), ...(stillOut ? { lock: { at: now(), real: true } } : {}) };
      if (stillOut) report.errors.push({ file: rel, error: `Created, but SharePoint did not check it in (${stillOut.message}). It is checked out to you until it is.` });
      writeBase(root, rel, buf);
      report.created.push({ file: rel, url: fresh.webUrl });
    } catch (e) {
      report.errors.push({ file: rel, error: e.message });
    }
  }

  // Locked but unchanged: hand them back so colleagues can edit.
  const modifiedSet = new Set(st.modified);
  for (const [rel, known] of Object.entries(manifest.files)) {
    if (!known.lock || modifiedSet.has(rel) || !inScope(rel)) continue;
    try {
      if (known.lock.real) await discardCheckout(driveId, known.itemId).catch((e) => { if (e.status !== 400) throw e; });
      delete known.lock;
      report.released.push(rel);
    } catch (e) {
      report.errors.push({ file: rel, error: `Could not release: ${e.message}` });
    }
  }

  writeManifest(root, manifest);
  if (!report.uploaded.length && !report.created.length && !report.needs_merge.length && !report.held_by_others.length && !report.errors.length) {
    report.message = 'Nothing to push. Everything matches SharePoint.';
  }
  audit({
    tool: 'workspace_push', target: manifest.source.webUrl, local: root, note: comment, auto: auto || undefined,
    uploaded: report.uploaded.map((u) => u.file).concat(report.created.map((c) => c.file)),
    needs_merge: report.needs_merge.length, held: report.held_by_others.length, errors: report.errors.length,
  });
  return report;
}

// ---------------------------------------------------------------------------
// Move / rename (organising)

async function moveImpl({ localPath, from, to }) {
  if (!loadConfig().allowWrite) throw new Error('This firm set M365 Bridge to read-only.');
  const root = path.resolve(localPath);
  const manifest = readManifest(root);
  const { driveId, itemId: rootId } = manifest.source;
  const src = clean(from).replace(/\/$/, '');
  const dst = clean(to).replace(/\/$/, '');
  if (!src || !dst || src === dst) throw new Error('Give a different "from" and "to", relative to the folder.');
  const caseOnly = src.toLowerCase() === dst.toLowerCase();
  if (!caseOnly && fs.existsSync(abs(root, dst))) throw new Error(`${dst} already exists here.`);

  const item = await graph(`/drives/${driveId}/items/${rootId}:/${enc(src)}?$select=${SELECT}`);
  const prefix = `${src}/`;
  const affected = Object.keys(manifest.files).filter((rel) => rel === src || rel.startsWith(prefix));
  for (const rel of affected) {
    if (manifest.files[rel].merge) throw new Error(`${rel} needs a merge first.`);
  }
  if (!item.folder && isCheckedOut(item) && !manifest.files[src]?.lock) {
    throw new Error(`${src} is checked out to a colleague. Ask them to check it in first.`);
  }
  const dstDir = dst.includes('/') ? dst.slice(0, dst.lastIndexOf('/')) : '';
  const parentId = await ensureRemoteDir(driveId, rootId, dstDir, new Map());
  const moved = await moveItem(driveId, item.id, { parentId, name: dst.split('/').pop() });

  // Mirror locally.
  try {
    if (fs.existsSync(abs(root, src))) {
      fs.mkdirSync(path.dirname(abs(root, dst)), { recursive: true });
      fs.renameSync(abs(root, src), abs(root, dst));
    }
  } catch (e) {
    // Moved in SharePoint, but the local copy is busy (open in Word). The next "get latest"
    // brings the moved file down and tidies the old one, so leave the record as it was.
    audit({ tool: 'workspace_move', target: manifest.source.webUrl, from: src, to: dst, local: 'busy' });
    return { moved: src, to: dst, url: moved.webUrl, note: 'Moved in SharePoint. The copy on this computer is open in another app; close it and get latest.' };
  }
  for (const dir of [BASE_DIR]) {
    if (fs.existsSync(abs(root, src, dir))) {
      fs.mkdirSync(path.dirname(abs(root, dst, dir)), { recursive: true });
      fs.renameSync(abs(root, src, dir), abs(root, dst, dir));
    }
  }
  for (const rel of affected) {
    const next = rel === src ? dst : dst + rel.slice(src.length);
    manifest.files[next] = manifest.files[rel];
    delete manifest.files[rel];
  }
  if (!item.folder && manifest.files[dst]) manifest.files[dst].eTag = moved.eTag;
  removeEmptyDirs(root);
  writeManifest(root, manifest);
  audit({ tool: 'workspace_move', target: manifest.source.webUrl, from: src, to: dst });
  return { moved: src, to: dst, url: moved.webUrl, files_moved: affected.length || 1 };
}

// ---------------------------------------------------------------------------
// Registry of synced folders

function rememberWorkspace(localPath, source) {
  const cfg = loadConfig();
  const list = cfg.workspaces.filter((w) => w.localPath !== localPath);
  list.unshift({ name: path.basename(localPath), localPath, webUrl: source.webUrl, driveId: source.driveId, itemId: source.itemId, lastPulled: now() });
  updateConfig({ workspaces: list.slice(0, 50) });
}

export function listWorkspaces() {
  return loadConfig().workspaces.map((w) => ({ ...w, exists: fs.existsSync(path.join(w.localPath, MANIFEST)) }));
}

export function forgetWorkspace(localPath) {
  const cfg = loadConfig();
  updateConfig({ workspaces: cfg.workspaces.filter((w) => w.localPath !== localPath) });
}

/** Every file checked out to you, across all synced folders. */
export function heldFiles() {
  const out = [];
  for (const w of listWorkspaces()) {
    if (!w.exists) continue;
    let manifest;
    try { manifest = readManifest(w.localPath); } catch { continue; }
    for (const [rel, known] of Object.entries(manifest.files)) {
      if (!known.lock && !known.merge) continue;
      const full = abs(w.localPath, rel);
      out.push({
        workspace: w.name, localPath: w.localPath, file: rel, since: known.lock?.at || known.merge?.at,
        edited: fs.existsSync(full) && hashFile(full) !== known.hash,
        needsMerge: !!known.merge, mergeWith: known.merge?.by,
      });
    }
  }
  return out;
}

/** The synced folder that contains a local file, and the file's path inside it. */
export function locateLocalFile(file) {
  const target = path.resolve(file);
  let dir = path.dirname(target);
  while (dir !== path.dirname(dir)) {
    if (fs.existsSync(path.join(dir, MANIFEST))) {
      const manifest = JSON.parse(fs.readFileSync(path.join(dir, MANIFEST), 'utf8'));
      const rel = toPosix(path.relative(dir, target));
      return { root: dir, rel, manifest, known: manifest.files[rel] || null };
    }
    dir = path.dirname(dir);
  }
  return null;
}

export function workspaceRefForLocalFile(file) {
  const hit = locateLocalFile(file);
  return hit?.known ? `${hit.manifest.source.driveId}|${hit.known.itemId}` : null;
}

export const lock = (a) => serial(a.localPath, () => lockImpl(a));
export const release = (a) => serial(a.localPath, () => releaseImpl(a));
export const pull = (a = {}) => serial(a.localPath, () => pullImpl(a));
export const merge = (a) => serial(a.localPath, () => mergeImpl(a));
export const push = (a) => serial(a.localPath, () => pushImpl(a));
export const move = (a) => serial(a.localPath, () => moveImpl(a));

/** Files checked out to you in one synced folder: when, and whether you changed them. */
export function locks(localPath) {
  const root = path.resolve(localPath);
  const manifest = readManifest(root);
  return Object.entries(manifest.files).filter(([, k]) => k.lock).map(([rel, k]) => {
    const full = abs(root, rel);
    return { file: rel, at: k.lock.at, merging: !!k.merge, edited: fs.existsSync(full) && hashFile(full) !== k.hash };
  });
}

/** The synced folder (its root) that contains a path, or null. */
export function workspaceRootOf(p) {
  let dir = path.resolve(p);
  while (dir !== path.dirname(dir)) {
    if (fs.existsSync(path.join(dir, MANIFEST))) return dir;
    dir = path.dirname(dir);
  }
  return null;
}

export { enc };
