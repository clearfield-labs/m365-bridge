// Automatic sync, so nobody has to remember "get latest" and "send".
//
//   When Claude finishes a reply (Claude Code's Stop hook): send every edited file in synced
//     folders to SharePoint, check it back in, and hand back files checked out but not changed.
//   When the person sends a message (UserPromptSubmit hook): get the latest for the folder Claude
//     is working in, and tell Claude what changed. Once per session, tell Claude about the day.
//   Every minute in the background service: the same safety net for anything the hooks missed
//     (a script that edited a file, a closed laptop, Claude quitting mid-reply), and refresh every
//     synced folder every few minutes.
//
// Nothing here can lose work: sending never overwrites a colleague's newer version (that becomes
// a merge), getting latest never overwrites your edits, and every save is a new SharePoint version.
// Merges are never finished automatically; Claude is told to do them.
import fs from 'node:fs';
import path from 'node:path';
import * as ws from './workspace.js';
import { loadConfig, isConfigured } from './config.js';
import { currentAccount } from './auth.js';

const ms = (name, def) => Number(process.env[name] || def);
const TICK = () => ms('M365_AUTOSYNC_TICK_MS', 60_000);
const SAVE_IDLE = () => ms('M365_AUTOSAVE_IDLE_MS', 3 * 60_000); // background: a file untouched for 3 minutes
const LOCK_IDLE = () => ms('M365_LOCK_IDLE_MS', 10 * 60_000); // background: a check-out unused for 10 minutes
const PULL_EVERY = () => ms('M365_AUTOPULL_MS', 5 * 60_000);
const PROMPT_PULL_EVERY = 60_000;
const BRIEF_EVERY = 3 * 60 * 60_000;

/** Per folder: last refresh, last save, last problem. Shown on the page. */
export const syncState = new Map();
const stateOf = (root) => {
  if (!syncState.has(root)) syncState.set(root, { lastPull: 0, lastSave: null, saved: [], problem: null, backoffUntil: 0 });
  return syncState.get(root);
};

const plural = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;
const names = (list, n = 3) => list.slice(0, n).map((f) => f.split('/').pop()).join(', ') + (list.length > n ? ` and ${list.length - n} more` : '');

/**
 * Send what changed in one synced folder and hand back unused check-outs.
 *   idleMs      only files not touched for this long (0 when Claude has just finished)
 *   lockIdleMs  only check-outs older than this
 */
export async function settle(root, { idleMs = SAVE_IDLE(), lockIdleMs = LOCK_IDLE(), note, via = true } = {}) {
  const save = loadConfig().autoSave !== false;
  // merges: found while saving just now. open: merges already waiting before this.
  const out = { saved: [], released: [], held: [], merges: [], open: [], errors: [] };
  const st = await ws.status({ localPath: root, checkRemote: false });
  const merging = new Set(st.needs_merge.map((m) => m.file));
  out.open.push(...st.needs_merge.map((m) => ({ file: m.file, by: m.changed_by })));
  const now = Date.now();
  const idle = (rel) => {
    try { return now - fs.statSync(path.join(root, ...rel.split('/'))).mtimeMs >= idleMs; } catch { return false; }
  };

  if (save) {
    const files = [...st.modified, ...st.new_files].filter((f) => !merging.has(f) && idle(f));
    if (files.length) {
      const r = await ws.push({ localPath: root, files, note: note || 'Saved automatically', via, auto: true });
      out.saved.push(...r.uploaded.map((u) => u.file), ...r.created.map((c) => c.file));
      out.released.push(...r.released);
      out.held.push(...r.held_by_others.map((h) => ({ file: h.file, by: h.checked_out_to })));
      out.merges.push(...r.needs_merge.map((m) => ({ file: m.file, by: m.changed_by })));
      out.errors.push(...r.errors.map((e) => `${e.file}: ${e.error}`));
    }
  }

  // Checked out but not changed: hand back so colleagues can edit. Never during a merge.
  const stale = ws.locks(root).filter((l) => !l.edited && !l.merging && now - Date.parse(l.at) >= lockIdleMs && !out.released.includes(l.file));
  if (stale.length) {
    const r = await ws.release({ localPath: root, files: stale.map((l) => l.file) });
    out.released.push(...r.released);
    out.errors.push(...r.errors.map((e) => `${e.file}: ${e.error}`));
  }

  const s = stateOf(root);
  if (out.saved.length) { s.lastSave = new Date().toISOString(); s.saved = out.saved; }
  return out;
}

/** One background pass over every synced folder. */
let ticking = false;
export async function tick({ force = false } = {}) {
  if (ticking) return;
  ticking = true;
  try {
    if (!isConfigured() || !(await currentAccount())) return;
    for (const w of ws.listWorkspaces()) {
      if (!w.exists) continue;
      const s = stateOf(w.localPath);
      if (!force && s.backoffUntil > Date.now()) continue;
      try {
        await settle(w.localPath, { note: 'Saved automatically by M365 Bridge', via: false });
        if (force || Date.now() - s.lastPull >= PULL_EVERY()) {
          await ws.pull({ localPath: w.localPath, auto: true });
          s.lastPull = Date.now();
        }
        s.problem = null;
        s.backoffUntil = 0;
      } catch (e) {
        // Signed out, offline, or Microsoft busy: try again in a few minutes, quietly.
        s.problem = e.message;
        s.backoffUntil = Date.now() + 5 * 60_000;
      }
    }
  } finally {
    ticking = false;
  }
}

let timer = null;
export function startAutoSync() {
  if (timer || process.env.M365_AUTOSYNC === '0' || process.env.M365_DEMO === '1') return;
  const run = () => tick().catch(() => {});
  setTimeout(run, 15_000).unref(); // shortly after login or wake
  timer = setInterval(run, TICK());
  timer.unref();
}

// ---------------------------------------------------------------------------
// Claude Code hooks (run here, in the background service, so all work on a folder is in one place)

const briefed = new Map(); // Claude session id -> when it was last told about the day

export async function onHookEvent(input = {}) {
  const event = input.hook_event_name;
  if (event === 'Stop') return onStop(input);
  if (event === 'UserPromptSubmit') return onPrompt(input);
  return null;
}

async function onStop(input) {
  const saved = [];
  const held = [];
  const merges = [];
  const waiting = []; // merges already open: not saved until Claude merges
  const errors = [];
  // Only folders Claude worked in: the one it runs in, and any with files checked out to you.
  // (Other edits, say in Word, are saved by the background pass, without "via Claude".)
  const here = input.cwd ? ws.workspaceRootOf(input.cwd) : null;
  for (const w of ws.listWorkspaces()) {
    if (!w.exists) continue;
    let active = w.localPath === here;
    try { active = active || ws.locks(w.localPath).length > 0; } catch { /* unreadable: skip */ }
    if (!active) continue;
    try {
      const r = await settle(w.localPath, { idleMs: 0, lockIdleMs: 0, note: 'Saved automatically' });
      saved.push(...r.saved);
      held.push(...r.held);
      merges.push(...r.merges.map((m) => ({ ...m, root: w.localPath })));
      waiting.push(...r.open.map((m) => m.file));
      errors.push(...r.errors);
    } catch (e) {
      errors.push(`${w.name}: ${e.message}`);
    }
  }
  // Saving just found a file changed in SharePoint too: have Claude merge it now (once per reply,
  // and only in the folder this Claude is working in).
  const mine = merges.filter((m) => m.root === here);
  if (mine.length && !input.stop_hook_active && loadConfig().autoSave !== false) {
    const list = mine.map((m) => `${m.file}${m.by ? ` (changed by ${m.by})` : ''} in ${m.root}`).join('; ');
    return {
      decision: 'block',
      reason: `Not saved yet: ${list} changed in SharePoint and on this computer. Call workspace_merge for each, combine both versions into the one file, tell the person what you combined, then call workspace_push with merged listing the file.`,
    };
  }
  const bits = [];
  if (saved.length) bits.push(`Saved to SharePoint: ${names(saved)}.`);
  if (held.length) bits.push(`Not saved, checked out to a colleague: ${held.map((h) => `${h.file.split('/').pop()} (${h.by})`).join(', ')}. It will save once they check it in.`);
  const unmerged = [...new Set([...merges.map((m) => m.file), ...waiting])];
  if (unmerged.length) bits.push(`Not saved yet, ${unmerged.length === 1 ? 'it needs' : 'they need'} a merge: ${names(unmerged)}. Ask Claude to merge.`);
  if (errors.length) bits.push(`Could not save: ${errors[0]}`);
  return bits.length ? { systemMessage: `M365 Bridge: ${bits.join(' ')}` } : null;
}

async function onPrompt(input) {
  const cfg = loadConfig();
  const parts = [];
  // 1. The folder Claude is working in: get the latest first.
  const root = input.cwd ? ws.workspaceRootOf(input.cwd) : null;
  if (root && cfg.autoSave !== false) {
    const s = stateOf(root);
    try {
      let r = null;
      if (Date.now() - s.lastPull >= PROMPT_PULL_EVERY) {
        // Large folders can take a while; never hold the person up for more than a few seconds.
        const job = ws.pull({ localPath: root, auto: true }).then((x) => { s.lastPull = Date.now(); return x; });
        r = await Promise.race([job, new Promise((res) => setTimeout(() => res(null), 8000).unref())]);
        job.catch(() => {});
      }
      if (r) {
        const got = [...r.updated.map((u) => `${u.file}${u.by ? ` (by ${u.by})` : ''}`), ...r.downloaded];
        if (got.length) parts.push(`M365 Bridge got the latest from SharePoint for this folder: ${got.slice(0, 10).join(', ')}${got.length > 10 ? ` and ${got.length - 10} more` : ''}. Re-read these before relying on earlier reads.`);
        if (r.removed_in_sharepoint.length) parts.push(`Moved or deleted in SharePoint (removed here): ${r.removed_in_sharepoint.slice(0, 10).join(', ')}.`);
      }
      const st = await ws.status({ localPath: root, checkRemote: false });
      if (st.needs_merge.length) parts.push(`Needs a merge before it can be saved: ${st.needs_merge.map((m) => `${m.file}${m.changed_by ? ` (changed by ${m.changed_by})` : ''}`).join(', ')}. Call workspace_merge for it when convenient.`);
      if (r?.held_by_others?.length) parts.push(`Checked out to colleagues right now (read-only for you): ${r.held_by_others.slice(0, 10).map((h) => h.file).join(', ')}.`);
    } catch (e) {
      parts.push(`M365 Bridge could not refresh this folder from SharePoint (${e.message}). Local files may be out of date.`);
    }
  }
  // 2. Once per session (and every few hours): the person's day.
  const sid = input.session_id || 'none';
  if (cfg.dayBrief !== false && isConfigured(cfg) && Date.now() - (briefed.get(sid) || 0) >= BRIEF_EVERY) {
    try {
      if (await currentAccount()) {
        const { myDay, briefText } = await import('./day.js');
        const day = await Promise.race([myDay(), new Promise((res) => setTimeout(() => res(null), 6000).unref())]);
        if (day) {
          parts.push(briefText(day));
          briefed.set(sid, Date.now());
        }
        if (briefed.size > 500) briefed.delete(briefed.keys().next().value);
      }
    } catch { /* never block a message on this */ }
  }
  return parts.length ? { hookSpecificOutput: { hookEventName: 'UserPromptSubmit', additionalContext: parts.join('\n\n') } } : null;
}
