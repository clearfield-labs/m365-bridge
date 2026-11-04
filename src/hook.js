// Claude Code PreToolUse hook. Before Claude edits a file inside a synced folder, check it
// out in SharePoint. If a colleague holds it, or SharePoint has a newer version, stop the
// edit and tell Claude why. Files outside synced folders pass straight through.
import fs from 'node:fs';
import path from 'node:path';

// Cheap check before loading anything heavy: is this file inside a synced folder at all?
function insideSyncedFolder(target) {
  let dir = path.dirname(target);
  while (dir !== path.dirname(dir)) {
    if (fs.existsSync(path.join(dir, '.m365', 'manifest.json'))) return true;
    dir = path.dirname(dir);
  }
  return false;
}

function readStdin() {
  return new Promise((resolve) => {
    let data = '';
    process.stdin.setEncoding('utf8');
    process.stdin.on('data', (c) => { data += c; });
    process.stdin.on('end', () => resolve(data));
    setTimeout(() => resolve(data), 5000).unref();
  });
}

const decide = (decision, reason) => {
  process.stdout.write(JSON.stringify({ hookSpecificOutput: { hookEventName: 'PreToolUse', permissionDecision: decision, permissionDecisionReason: reason } }));
};

export async function runHook() {
  let input = {};
  try { input = JSON.parse(await readStdin()); } catch { return; }
  const ti = input.tool_input || {};
  const file = ti.file_path || ti.notebook_path;
  if (!file) return;
  const target = path.resolve(input.cwd || process.cwd(), file);
  if (!insideSyncedFolder(target)) return;

  let ws;
  let hit;
  try {
    ws = await import('./workspace.js');
    hit = ws.locateLocalFile(target);
  } catch (e) {
    // Inside a synced folder but its record could not be read: stop rather than edit unlocked.
    return decide('deny', `Could not read this synced folder's record (${e.message}). Try again in a moment.`);
  }
  if (!hit || !hit.known || hit.known.lock) return; // not ours, new file, or already checked out to you
  if (hit.known.merge) return; // a merge is open: Claude is writing the merged file
  try {
    const r = await ws.lock({ localPath: hit.root, files: [hit.rel] });
    if (r.held_by_others.length) {
      return decide('deny', `${hit.rel} is checked out to ${r.held_by_others[0].checked_out_to} in SharePoint, so it cannot be edited right now. Tell the person who has it, and work on something else or wait.`);
    }
    if (r.refreshed_first.length) {
      return decide('deny', `${hit.rel} had a newer version in SharePoint${r.refreshed_first[0].changed_by ? ` from ${r.refreshed_first[0].changed_by}` : ''}. I downloaded it and checked it out to you. Read the file again, then redo the edit.`);
    }
    if (r.needs_merge.length) {
      return decide('deny', `${hit.rel} changed in SharePoint and here. Call the m365 workspace_merge tool for this file, then edit.`);
    }
    // locked, already yours, or lock not supported (push still checks versions): allow.
  } catch (e) {
    return decide('deny', `Could not check ${hit.rel} out in SharePoint: ${e.message}`);
  }
}
