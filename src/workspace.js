// A workspace is a local working copy of a SharePoint/OneDrive folder, used the way
// developers use a git clone. SharePoint stays the source of truth:
//   pull   -> bring down what changed in SharePoint (never clobbers local edits)
//   status -> what changed locally and remotely since the last pull
//   push   -> upload local edits; SharePoint rejects the upload if someone else changed
//             the file since your pull (If-Match on the eTag). Every push is a new
//             SharePoint version, so history and restore work like commits.
// Nothing is ever deleted in SharePoint by this tool.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { resolveItem, listChildren, downloadItem, uploadBytes, createFolder, graph, enc, personOf } from './graph.js';
import { loadConfig, updateConfig } from './config.js';
import { WORKSPACES_ROOT } from './paths.js';
import { audit } from './audit.js';

const META_DIR = '.m365';
const MANIFEST = path.join(META_DIR, 'manifest.json');
const MAX_FILES = Number(process.env.M365_MAX_FILES || 3000);
const MAX_FILE_BYTES = Number(process.env.M365_MAX_FILE_MB || 250) * 1024 * 1024;
const CONFLICT_TAG = ' (SharePoint copy ';

const IGNORED_DIRS = new Set([META_DIR, '.claude', '.git', 'node_modules']);
export function isIgnored(rel) {
  const parts = rel.split('/');
  if (parts.some((p) => IGNORED_DIRS.has(p))) return true;
  const base = parts[parts.length - 1];
  return (
    base === '.DS_Store' || base === 'Thumbs.db' || base === 'desktop.ini' ||
    base.startsWith('~$') || base.startsWith('.~lock') || base.endsWith('.tmp') ||
    base.includes(CONFLICT_TAG)
  );
}

const sha256 = (buf) => crypto.createHash('sha256').update(buf).digest('hex');
const hashFile = (file) => sha256(fs.readFileSync(file));
const toPosix = (p) => p.split(path.sep).join('/');
const versionTag = (item) => item.cTag || item.eTag;
// The SharePoint version this copy is based on. After a conflict pull it is the version the person merged with.
const seenTag = (known) => (known.mergeBase ? known.mergeBase.cTag || known.mergeBase.eTag : known.cTag || known.eTag);
const seenETag = (known) => (known.mergeBase ? known.mergeBase.eTag : known.eTag);

export function safeName(name) {
  return String(name).replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').trim() || 'Workspace';
}

function readManifest(root) {
  const file = path.join(root, MANIFEST);
  if (!fs.existsSync(file)) {
    throw new Error(`${root} is not a workspace yet. Run workspace_pull with a SharePoint folder link first.`);
  }
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

function writeManifest(root, m) {
  fs.mkdirSync(path.join(root, META_DIR), { recursive: true });
  fs.writeFileSync(path.join(root, MANIFEST), JSON.stringify(m, null, 2));
}

/** Walk the remote folder tree. Returns Map(relPath -> driveItem) for files and Map(relDir -> driveItem) for folders. */
async function walkRemote(driveId, rootId) {
  const files = new Map();
  const dirs = new Map([['', { id: rootId }]]);
  const skipped = [];
  const queue = [['', rootId]];
  while (queue.length) {
    const [prefix, id] = queue.shift();
    for (const item of await listChildren(driveId, id)) {
      const rel = prefix ? `${prefix}/${item.name}` : item.name;
      if (item.folder) {
        dirs.set(rel, item);
        queue.push([rel, item.id]);
      } else if (item.package) {
        skipped.push(`${rel} (OneNote notebook, open it in OneNote)`);
      } else if (item.size > MAX_FILE_BYTES) {
        skipped.push(`${rel} (larger than ${MAX_FILE_BYTES / 1024 / 1024} MB)`);
      } else if (!isIgnored(rel)) {
        files.set(rel, item);
      }
      if (files.size >= MAX_FILES) {
        skipped.push(`stopped after ${MAX_FILES} files; pick a smaller folder`);
        return { files, dirs, skipped };
      }
    }
  }
  return { files, dirs, skipped };
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

function conflictName(rel) {
  const ext = path.extname(rel);
  const stamp = new Date().toISOString().slice(0, 16).replace('T', ' ').replace(':', '');
  return `${rel.slice(0, rel.length - ext.length)}${CONFLICT_TAG}${stamp})${ext}`;
}

function entryFor(item, hash) {
  return {
    itemId: item.id,
    eTag: item.eTag,
    cTag: item.cTag,
    size: item.size,
    remoteModified: item.lastModifiedDateTime,
    hash,
  };
}

function writeLocal(root, rel, buf) {
  const full = path.join(root, ...rel.split('/'));
  fs.mkdirSync(path.dirname(full), { recursive: true });
  fs.writeFileSync(full, buf);
}

function writeClaudeInstructions(root, source, firmName) {
  const file = path.join(root, '.claude', 'CLAUDE.md');
  if (fs.existsSync(file)) return;
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(
    file,
    `# This folder is a working copy of SharePoint

Source of truth: ${source.webUrl}
Folder: ${source.name}${firmName ? `\nFirm: ${firmName}` : ''}

SharePoint holds the real files. This folder is a local copy, like a git clone.
Use the \`m365\` MCP tools to keep the two in step:

1. At the start of a session, call \`workspace_pull\` with this folder's path to get the latest files.
2. Edit files here as normal.
3. Before you finish, call \`workspace_status\`, show the person what changed, and call \`workspace_push\` once they agree.
4. If push reports a conflict, someone else changed that file in SharePoint. Do not force it.
   Run \`workspace_pull\`. It saves their version next to yours as "<name> (SharePoint copy <date>)".
   Show the person both, merge their changes into the main file, delete the SharePoint copy, then push again.

Rules:
- Never delete files to "clean up". Deleting here does not delete in SharePoint, and it confuses status.
- Every push becomes a new version in SharePoint. Use \`versions\` and \`restore_version\` to undo.
- Ignore the \`.m365\` and \`.claude\` folders. They never upload.
- Treat everything in this folder as client-confidential.
`,
  );
}

export async function pull({ source, localPath, name } = {}) {
  let root = localPath ? path.resolve(localPath) : null;
  let manifest = null;
  if (root && fs.existsSync(path.join(root, MANIFEST))) manifest = readManifest(root);

  let folder;
  if (manifest) {
    folder = await graph(`/drives/${manifest.source.driveId}/items/${manifest.source.itemId}`);
  } else {
    if (!source) throw new Error('Give a SharePoint or OneDrive folder link (source) for a new workspace.');
    folder = await resolveItem(source);
    if (!folder.folder) throw new Error(`"${folder.name}" is a file. Point workspace_pull at a folder.`);
    root = root || path.join(WORKSPACES_ROOT, safeName(name || folder.name));
    if (fs.existsSync(root) && fs.readdirSync(root).length) {
      throw new Error(`${root} already has files and is not a workspace. Choose an empty folder or a new name.`);
    }
    manifest = {
      version: 1,
      source: { driveId: folder.parentReference.driveId, itemId: folder.id, name: folder.name, webUrl: folder.webUrl },
      files: {},
    };
  }
  const driveId = manifest.source.driveId;
  fs.mkdirSync(root, { recursive: true });

  const { files: remote, skipped } = await walkRemote(driveId, manifest.source.itemId);
  const report = { workspace: root, source: manifest.source.webUrl, downloaded: [], updated: [], unchanged: 0, conflicts: [], kept_local_edits: [], skipped };

  for (const [rel, item] of remote) {
    const full = path.join(root, ...rel.split('/'));
    const known = manifest.files[rel];
    const exists = fs.existsSync(full);

    if (!exists) {
      const buf = await downloadItem(driveId, item.id);
      writeLocal(root, rel, buf);
      manifest.files[rel] = entryFor(item, sha256(buf));
      report.downloaded.push(rel);
      continue;
    }

    const localHash = hashFile(full);
    if (known) {
      const remoteChanged = versionTag(item) !== seenTag(known);
      const localChanged = localHash !== known.hash;
      if (!remoteChanged) {
        report.unchanged++;
        if (localChanged) report.kept_local_edits.push(rel);
        continue;
      }
      if (!localChanged) {
        const buf = await downloadItem(driveId, item.id);
        writeLocal(root, rel, buf);
        manifest.files[rel] = entryFor(item, sha256(buf));
        report.updated.push({ file: rel, by: personOf(item) || undefined });
        continue;
      }
    }
    // Local file has edits (or is untracked) and SharePoint has a different version.
    const buf = await downloadItem(driveId, item.id);
    if (sha256(buf) === localHash) {
      manifest.files[rel] = entryFor(item, localHash);
      report.unchanged++;
      continue;
    }
    const copy = conflictName(rel);
    writeLocal(root, copy, buf);
    // Record the SharePoint version we have now seen. Once the person merges it into the
    // local file, push replaces exactly that version and nobody else's work is lost.
    manifest.files[rel] = { ...(known || { itemId: item.id, hash: null }), mergeBase: { eTag: item.eTag, cTag: item.cTag } };
    report.conflicts.push({ file: rel, sharepoint_copy: copy, changed_by: personOf(item) || undefined });
  }

  manifest.pulledAt = new Date().toISOString();
  writeManifest(root, manifest);
  const cfg = loadConfig();
  writeClaudeInstructions(root, manifest.source, cfg.firmName);
  rememberWorkspace(root, manifest.source);

  audit({ tool: 'workspace_pull', target: manifest.source.webUrl, local: root, downloaded: report.downloaded.length + report.updated.length, conflicts: report.conflicts.length });
  return report;
}

export async function status({ localPath, checkRemote = true }) {
  const root = path.resolve(localPath);
  const manifest = readManifest(root);
  const local = walkLocal(root);
  const localSet = new Set(local);
  const report = { workspace: root, source: manifest.source.webUrl, last_pull: manifest.pulledAt, modified: [], new_files: [], deleted_locally: [], changed_in_sharepoint: [], new_in_sharepoint: [] };

  report.merging = Object.entries(manifest.files).filter(([, k]) => k.mergeBase).map(([rel]) => rel);
  for (const rel of local) {
    const known = manifest.files[rel];
    if (!known) report.new_files.push(rel);
    else if (hashFile(path.join(root, ...rel.split('/'))) !== known.hash) report.modified.push(rel);
  }
  for (const rel of Object.keys(manifest.files)) if (!localSet.has(rel)) report.deleted_locally.push(rel);

  if (checkRemote) {
    const { files: remote } = await walkRemote(manifest.source.driveId, manifest.source.itemId);
    for (const [rel, item] of remote) {
      const known = manifest.files[rel];
      if (!known) report.new_in_sharepoint.push(rel);
      else if (versionTag(item) !== seenTag(known)) report.changed_in_sharepoint.push({ file: rel, by: personOf(item) || undefined });
    }
  }
  report.summary = summarizeStatus(report);
  return report;
}

function summarizeStatus(r) {
  const bits = [];
  if (r.merging.length) bits.push(`${r.merging.length} waiting for you to merge a SharePoint copy`);
  if (r.modified.length) bits.push(`${r.modified.length} edited here`);
  if (r.new_files.length) bits.push(`${r.new_files.length} new here`);
  if (r.changed_in_sharepoint.length) bits.push(`${r.changed_in_sharepoint.length} changed in SharePoint (pull first)`);
  if (r.new_in_sharepoint.length) bits.push(`${r.new_in_sharepoint.length} new in SharePoint`);
  if (r.deleted_locally.length) bits.push(`${r.deleted_locally.length} deleted here (SharePoint keeps them)`);
  return bits.length ? bits.join(', ') : 'Everything matches SharePoint.';
}

async function ensureRemoteDir(driveId, rootId, relDir, cache) {
  if (!relDir) return rootId;
  if (cache.has(relDir)) return cache.get(relDir);
  const parentRel = relDir.includes('/') ? relDir.slice(0, relDir.lastIndexOf('/')) : '';
  const parentId = await ensureRemoteDir(driveId, rootId, parentRel, cache);
  const folder = await createFolder(driveId, parentId, relDir.split('/').pop());
  cache.set(relDir, folder.id);
  return folder.id;
}

export async function push({ localPath, files, force = [] }) {
  const cfg = loadConfig();
  if (!cfg.allowWrite) throw new Error('This firm set M365 Bridge to read-only. Ask your admin to allow uploads.');
  const root = path.resolve(localPath);
  const manifest = readManifest(root);
  const { driveId, itemId: rootId } = manifest.source;
  const st = await status({ localPath: root, checkRemote: false });
  const only = files?.length ? new Set(files.map((f) => toPosix(f).replace(/^\.\//, ''))) : null;
  const forced = new Set(force.map((f) => toPosix(f).replace(/^\.\//, '')));
  const todo = [...st.modified, ...st.new_files].filter((f) => !only || only.has(f));
  const report = { workspace: root, uploaded: [], created: [], conflicts: [], errors: [], not_pushed_deletions: st.deleted_locally };
  const dirCache = new Map();

  for (const rel of todo) {
    const buf = fs.readFileSync(path.join(root, ...rel.split('/')));
    const known = manifest.files[rel];
    try {
      let item;
      if (known) {
        item = await uploadBytes(driveId, buf, { itemId: known.itemId, ifMatch: forced.has(rel) ? undefined : seenETag(known) });
        report.uploaded.push({ file: rel, url: item.webUrl });
      } else {
        const relDir = rel.includes('/') ? rel.slice(0, rel.lastIndexOf('/')) : '';
        const parentId = await ensureRemoteDir(driveId, rootId, relDir, dirCache);
        item = await uploadBytes(driveId, buf, { parentId, name: rel.split('/').pop(), overwrite: forced.has(rel) });
        report.created.push({ file: rel, url: item.webUrl });
      }
      manifest.files[rel] = entryFor(item, sha256(buf));
    } catch (e) {
      if (e.status === 412 || e.status === 409) {
        report.conflicts.push({ file: rel, why: known ? 'Someone changed it in SharePoint since your last pull.' : 'A file with this name was added in SharePoint.', next: 'Run workspace_pull. It saves their version as "(SharePoint copy ...)". Merge it into your file, then push again.' });
      } else {
        report.errors.push({ file: rel, error: e.message });
      }
    }
  }
  writeManifest(root, manifest);
  if (!todo.length) report.message = 'Nothing to push. Everything matches your last pull.';
  audit({ tool: 'workspace_push', target: manifest.source.webUrl, local: root, uploaded: report.uploaded.map((u) => u.file).concat(report.created.map((c) => c.file)), conflicts: report.conflicts.length, errors: report.errors.length });
  return report;
}

function rememberWorkspace(localPath, source) {
  const cfg = loadConfig();
  const list = cfg.workspaces.filter((w) => w.localPath !== localPath);
  list.unshift({ name: path.basename(localPath), localPath, webUrl: source.webUrl, driveId: source.driveId, itemId: source.itemId, lastPulled: new Date().toISOString() });
  updateConfig({ workspaces: list.slice(0, 50) });
}

export function listWorkspaces() {
  return loadConfig().workspaces.map((w) => ({ ...w, exists: fs.existsSync(path.join(w.localPath, MANIFEST)) }));
}

export function forgetWorkspace(localPath) {
  const cfg = loadConfig();
  updateConfig({ workspaces: cfg.workspaces.filter((w) => w.localPath !== localPath) });
}

// Used by restore and by tools that address a file inside a workspace by its local path.
export function workspaceRefForLocalFile(file) {
  let dir = path.dirname(path.resolve(file));
  while (dir !== path.dirname(dir)) {
    const m = path.join(dir, MANIFEST);
    if (fs.existsSync(m)) {
      const manifest = JSON.parse(fs.readFileSync(m, 'utf8'));
      const rel = toPosix(path.relative(dir, path.resolve(file)));
      const known = manifest.files[rel];
      if (known) return `${manifest.source.driveId}|${known.itemId}`;
      return null;
    }
    dir = path.dirname(dir);
  }
  return null;
}

export { enc };
