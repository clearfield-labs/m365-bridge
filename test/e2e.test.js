// End to end: the real MCP server (stdio and HTTP), the hook and firm setup, against a mock Graph.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { createMockGraph } from './mock-graph.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const CLI = path.join(ROOT, 'bin', 'm365-bridge.js');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'm365b-'));
const WS_ROOT = path.join(tmp, 'workspaces');
const PORT = '47365';
const mock = createMockGraph();
let client;
let env;
let background;

async function call(name, args = {}, c = client) {
  const r = await c.callTool({ name, arguments: args });
  const text = r.content[0].text;
  let data;
  try { data = JSON.parse(text); } catch { data = text; }
  return { data, isError: !!r.isError, text };
}

before(async () => {
  const base = await mock.listen();
  env = {
    ...process.env,
    M365_GRAPH_BASE: `${base}/v1.0`,
    M365_TEST_TOKEN: 'test-token',
    M365_BRIDGE_HOME: path.join(tmp, 'home'),
    M365_BRIDGE_KEYSTORE: 'file',
    M365_WORKSPACES: WS_ROOT,
    M365_BRIDGE_PORT: PORT,
    M365_TENANT_ID: '11111111-1111-1111-1111-111111111111',
    M365_CLIENT_ID: '22222222-2222-2222-2222-222222222222',
    M365_SIMPLE_UPLOAD_MAX: '200',
    M365_NO_SERVICE: '1',
    M365_NO_BROWSER: '1',
  };
  client = new Client({ name: 'test', version: '1.0.0' });
  await client.connect(new StdioClientTransport({ command: process.execPath, args: [CLI, 'serve'], env, stderr: 'inherit' }));
});

after(async () => {
  await client?.close();
  background?.kill();
  await mock.close();
  try { fs.rmSync(tmp, { recursive: true, force: true }); } catch { /* Windows may hold a file briefly */ }
});

const WS = () => path.join(WS_ROOT, 'Henderson v Acme');
const MATTER = 'Matters/Henderson v Acme';
const MATTER_URL = () => `${mock.HOST}/${MATTER}`;
const remote = (rel) => mock.find(`${MATTER}/${rel}`);
const noCopies = () => {
  const all = [];
  const walk = (d) => fs.readdirSync(d, { withFileTypes: true }).forEach((x) => (x.isDirectory() ? x.name !== '.m365' && x.name !== '.claude' && walk(path.join(d, x.name)) : all.push(x.name)));
  walk(WS());
  return all.filter((n) => /copy|conflict/i.test(n));
};

test('tools are registered', async () => {
  const { tools } = await client.listTools();
  const names = tools.map((t) => t.name);
  for (const n of ['status', 'setup', 'firm_setup', 'sign_in', 'search', 'list_folder', 'read_file', 'upload', 'move', 'versions', 'restore_version', 'workspace_pull', 'workspace_status', 'workspace_lock', 'workspace_push', 'workspace_merge', 'workspace_release', 'workspace_move', 'mail_search', 'mail_draft', 'calendar_view', 'graph_request']) {
    assert.ok(names.includes(n), `missing ${n}`);
  }
});

test('status reports setup', async () => {
  const { data } = await call('status');
  assert.equal(data.configured, true);
  assert.ok(data.signed_in_as.includes('test@example.com'));
});

test('search and read_file', async () => {
  const { data } = await call('search', { query: 'engagement' });
  assert.equal(data[0].name, 'Engagement Letter.txt');
  const read = await call('read_file', { location: data[0].ref });
  assert.match(read.data.text, /Fee: hourly/);
});

test('list_folder by SharePoint AllItems.aspx address-bar link', async () => {
  const id = encodeURIComponent('/sites/Lit/Shared Documents/Matters/Henderson v Acme');
  const { data, isError, text } = await call('list_folder', { location: `https://mock.sharepoint.com/sites/Lit/Shared%20Documents/Forms/AllItems.aspx?id=${id}` });
  assert.ok(!isError, text);
  assert.ok(data.items.some((i) => i.name === 'Pleadings' && i.type === 'folder'));
});

test('pull copies the folder and writes Claude instructions', async () => {
  const { data, isError, text } = await call('workspace_pull', { source: MATTER_URL() });
  assert.ok(!isError, text);
  assert.equal(data.workspace, WS());
  assert.equal(data.downloaded.length, 3);
  assert.ok(fs.existsSync(path.join(WS(), 'Pleadings', 'Complaint.txt')));
  assert.ok(!fs.existsSync(path.join(WS(), '~$Lockfile.txt')), 'office lock files are ignored');
  assert.match(fs.readFileSync(path.join(WS(), '.claude', 'CLAUDE.md'), 'utf8'), /workspace_lock/);
});

test('lock checks the file out; push uploads, checks in with the note, and releases', async () => {
  const lock = await call('workspace_lock', { local_path: WS(), files: ['Timeline.md'] });
  assert.deepEqual(lock.data.locked, ['Timeline.md'], lock.text);
  assert.equal(remote('Timeline.md').checkout, 'test@example.com');

  fs.appendFileSync(path.join(WS(), 'Timeline.md'), '- 2026-02-01 demand letter\n');
  const st = await call('workspace_status', { local_path: WS() });
  assert.deepEqual(st.data.modified, ['Timeline.md']);
  assert.deepEqual(st.data.locked_by_you, ['Timeline.md']);

  const push = await call('workspace_push', { local_path: WS(), note: 'Added demand letter date' });
  assert.equal(push.data.uploaded.length, 1, push.text);
  assert.match(remote('Timeline.md').content.toString(), /demand letter/);
  assert.equal(remote('Timeline.md').checkout, null, 'checked back in');
  assert.equal(remote('Timeline.md').versions[0].comment, 'Added demand letter date (via Claude)');
  const st2 = await call('workspace_status', { local_path: WS() });
  assert.equal(st2.data.summary, 'Everything matches SharePoint.');
});

test('push locks on its own if Claude did not', async () => {
  fs.appendFileSync(path.join(WS(), 'Timeline.md'), '- 2026-03-01 mediation\n');
  const push = await call('workspace_push', { local_path: WS(), note: 'Mediation date' });
  assert.equal(push.data.uploaded.length, 1, push.text);
  assert.equal(remote('Timeline.md').versions[0].comment, 'Mediation date (via Claude)');
  assert.equal(remote('Timeline.md').checkout, null);
});

test('a file checked out to a colleague cannot be locked or overwritten', async () => {
  mock.remoteCheckout(`${MATTER}/Engagement Letter.txt`, 'maria@mockfirm.com');
  const lock = await call('workspace_lock', { local_path: WS(), files: ['Engagement Letter.txt'] });
  assert.equal(lock.data.held_by_others[0].checked_out_to, 'maria@mockfirm.com');

  fs.writeFileSync(path.join(WS(), 'Engagement Letter.txt'), 'my edit');
  const push = await call('workspace_push', { local_path: WS() });
  assert.equal(push.data.held_by_others[0].checked_out_to, 'maria@mockfirm.com');
  assert.match(remote('Engagement Letter.txt').content.toString(), /Fee: hourly/, 'colleague file untouched');

  const st = await call('workspace_status', { local_path: WS() });
  assert.deepEqual(st.data.held_by_others.map((h) => h.file), ['Engagement Letter.txt']);
  mock.remoteCheckin(`${MATTER}/Engagement Letter.txt`);
  fs.writeFileSync(path.join(WS(), 'Engagement Letter.txt'), 'Engagement letter for Henderson. Fee: hourly.');
});

test('pull brings down a colleague edit without touching local work', async () => {
  mock.remoteEdit(`${MATTER}/Pleadings/Complaint.txt`, 'COMPLAINT draft v2 (partner edits)', 'Sarah Chen');
  const st = await call('workspace_status', { local_path: WS() });
  assert.equal(st.data.changed_in_sharepoint[0].file, 'Pleadings/Complaint.txt');
  assert.equal(st.data.changed_in_sharepoint[0].by, 'Sarah Chen');
  const pull = await call('workspace_pull', { local_path: WS() });
  assert.equal(pull.data.updated[0].file, 'Pleadings/Complaint.txt');
  assert.match(fs.readFileSync(path.join(WS(), 'Pleadings', 'Complaint.txt'), 'utf8'), /v2/);
});

test('lock refreshes a stale but unedited file first', async () => {
  mock.remoteEdit(`${MATTER}/Pleadings/Complaint.txt`, 'COMPLAINT draft v3', 'Sarah Chen');
  const lock = await call('workspace_lock', { local_path: WS(), files: ['Pleadings/Complaint.txt'] });
  assert.equal(lock.data.refreshed_first[0].file, 'Pleadings/Complaint.txt', lock.text);
  assert.match(fs.readFileSync(path.join(WS(), 'Pleadings', 'Complaint.txt'), 'utf8'), /v3/);
  const rel = await call('workspace_release', { local_path: WS() });
  assert.deepEqual(rel.data.released, ['Pleadings/Complaint.txt']);
  assert.equal(remote('Pleadings/Complaint.txt').checkout, null);
});

test('both sides changed: one file, merged by Claude, never a second copy', async () => {
  const file = path.join(WS(), 'Engagement Letter.txt');
  fs.writeFileSync(file, 'Engagement letter for Henderson. Fee: flat $5,000.');
  mock.remoteEdit(`${MATTER}/Engagement Letter.txt`, 'Engagement letter for Henderson. Fee: hourly. Retainer $2,500.', 'Maria Lopez');

  const pull = await call('workspace_pull', { local_path: WS() });
  assert.equal(pull.data.needs_merge[0].file, 'Engagement Letter.txt', pull.text);
  assert.deepEqual(noCopies(), [], 'no copy files in the folder');
  assert.match(fs.readFileSync(file, 'utf8'), /flat/, 'local edit kept');

  const blocked = await call('workspace_push', { local_path: WS() });
  assert.equal(blocked.data.needs_merge.length, 1);
  assert.match(remote('Engagement Letter.txt').content.toString(), /Retainer/, 'colleague work kept');

  const m = await call('workspace_merge', { local_path: WS(), file: 'Engagement Letter.txt' });
  assert.match(fs.readFileSync(m.data.theirs, 'utf8'), /Retainer/);
  assert.match(fs.readFileSync(m.data.base, 'utf8'), /Fee: hourly\.$/);
  assert.equal(m.data.changed_by, 'Maria Lopez');
  assert.equal(remote('Engagement Letter.txt').checkout, 'test@example.com', 'held while merging');

  fs.writeFileSync(file, 'Engagement letter for Henderson. Fee: flat $5,000. Retainer $2,500.');
  const push = await call('workspace_push', { local_path: WS(), merged: ['Engagement Letter.txt'], note: 'Merged fee terms with Maria' });
  assert.equal(push.data.uploaded.length, 1, push.text);
  assert.match(remote('Engagement Letter.txt').content.toString(), /flat \$5,000\. Retainer/);
  assert.equal(remote('Engagement Letter.txt').checkout, null);
  assert.deepEqual(noCopies(), []);
});

test('move organises in SharePoint and locally', async () => {
  const r = await call('workspace_move', { local_path: WS(), from: 'Pleadings/Complaint.txt', to: 'Court/Pleadings/Complaint.txt' });
  assert.ok(!r.isError, r.text);
  assert.ok(remote('Court/Pleadings/Complaint.txt'));
  assert.ok(!remote('Pleadings/Complaint.txt'));
  assert.ok(fs.existsSync(path.join(WS(), 'Court', 'Pleadings', 'Complaint.txt')));
  assert.ok(!fs.existsSync(path.join(WS(), 'Pleadings', 'Complaint.txt')));
  const st = await call('workspace_status', { local_path: WS() });
  assert.equal(st.data.summary, 'Everything matches SharePoint.', st.text);
});

test('a move made in SharePoint shows up locally on pull', async () => {
  const it = remote('Timeline.md');
  it.parentId = remote('Court').id;
  const pull = await call('workspace_pull', { local_path: WS() });
  assert.deepEqual(pull.data.removed_in_sharepoint, ['Timeline.md']);
  assert.deepEqual(pull.data.downloaded, ['Court/Timeline.md']);
  assert.ok(!fs.existsSync(path.join(WS(), 'Timeline.md')));
});

test('hook: checks a file out before Claude edits it, and blocks a colleague\'s file', async () => {
  const target = path.join(WS(), 'Engagement Letter.txt');
  // Async: the mock Graph runs in this process, so a blocking spawn would deadlock.
  const run = (file) => new Promise((resolve) => {
    const child = spawn(process.execPath, [CLI, 'hook'], { env });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (d) => { stdout += d; });
    child.stderr.on('data', (d) => { stderr += d; });
    child.on('close', (status) => resolve({ status, stdout, stderr }));
    child.stdin.end(JSON.stringify({ tool_name: 'Edit', tool_input: { file_path: file }, cwd: WS() }));
  });

  const ok = await run(target);
  assert.equal(ok.status, 0, ok.stderr);
  assert.equal(ok.stdout, '', 'allowed silently');
  assert.equal(remote('Engagement Letter.txt').checkout, 'test@example.com');

  mock.remoteCheckout(`${MATTER}/Court/Timeline.md`, 'maria@mockfirm.com');
  const no = await run(path.join(WS(), 'Court', 'Timeline.md'));
  const out = JSON.parse(no.stdout);
  assert.equal(out.hookSpecificOutput.permissionDecision, 'deny');
  assert.match(out.hookSpecificOutput.permissionDecisionReason, /maria@mockfirm\.com/);
  mock.remoteCheckin(`${MATTER}/Court/Timeline.md`);

  const outside = await run(path.join(tmp, 'notes.txt'));
  assert.equal(outside.stdout, '', 'files outside synced folders pass through');
  await call('workspace_release', { local_path: WS() });
});

test('new files in new subfolders are created; large files use an upload session', async () => {
  fs.mkdirSync(path.join(WS(), 'Discovery', 'Set 1'), { recursive: true });
  fs.writeFileSync(path.join(WS(), 'Discovery', 'Set 1', 'Interrogatories.txt'), 'Q'.repeat(1000));
  const push = await call('workspace_push', { local_path: WS() });
  assert.equal(push.data.created.length, 1, push.text);
  assert.equal(remote('Discovery/Set 1/Interrogatories.txt').content.length, 1000);
  assert.ok(mock.log.some((l) => l.startsWith('PUT /upload/')));
});

test('deleting locally never deletes in SharePoint', async () => {
  fs.rmSync(path.join(WS(), 'Court', 'Timeline.md'));
  const push = await call('workspace_push', { local_path: WS() });
  assert.deepEqual(push.data.not_pushed_deletions, ['Court/Timeline.md']);
  assert.ok(remote('Court/Timeline.md'));
});

test('versions and restore_version', async () => {
  const loc = `${MATTER_URL()}/Engagement Letter.txt`;
  const v = await call('versions', { location: loc });
  assert.ok(v.data.versions.length >= 3);
  const r = await call('restore_version', { location: loc, version: v.data.versions.at(-1).version });
  assert.ok(!r.isError, r.text);
  assert.match(remote('Engagement Letter.txt').content.toString(), /^Engagement letter for Henderson\. Fee: hourly\.$/);
});

test('upload refuses to clobber unless replace is set', async () => {
  const src = path.join(tmp, 'Complaint.txt');
  fs.writeFileSync(src, 'local complaint');
  const r = await call('upload', { local_path: src, folder: `${MATTER_URL()}/Court/Pleadings` });
  assert.ok(r.isError);
  assert.match(r.text, /already exists/);
  const r2 = await call('upload', { local_path: src, folder: `${MATTER_URL()}/Court/Pleadings`, replace: true });
  assert.ok(!r2.isError, r2.text);
});

test('mail tools and graph_request', async () => {
  const s = await call('mail_search', { query: 'settlement' });
  assert.equal(s.data[0].subject, 'Henderson settlement offer');
  const m = await call('mail_read', { id: 'M1' });
  assert.match(m.data.body, /\$50,000/);
  const d = await call('mail_draft', { to: ['client@example.com'], subject: 'Update', body: 'Draft only' });
  assert.equal(d.data.draft_saved, true);
  const chats = await call('graph_request', { path: '/me/chats' });
  assert.match(chats.text, /Henderson team/);
  const del = await call('graph_request', { method: 'DELETE', path: '/me/messages/M1' });
  assert.ok(del.isError);
  const evil = await call('graph_request', { path: 'https://evil.example.com/steal' });
  assert.match(evil.text, /only talks to Microsoft Graph/);
});

test('background service: Claude Code connects over HTTP with its key; browsers and strangers cannot', async () => {
  // Give this computer a Claude key, the way install does.
  const { ensureMcpToken } = await (async () => {
    Object.assign(process.env, { M365_BRIDGE_HOME: env.M365_BRIDGE_HOME, M365_BRIDGE_KEYSTORE: 'file', M365_BRIDGE_PORT: PORT });
    return import('../src/config.js');
  })();
  const token = ensureMcpToken();
  background = spawn(process.execPath, [CLI, 'background'], { env, stdio: 'ignore' });
  for (let i = 0; i < 50; i++) {
    try { if ((await fetch(`http://127.0.0.1:${PORT}/healthz`)).ok) break; } catch { /* starting */ }
    await new Promise((r) => setTimeout(r, 100));
  }
  const url = `http://127.0.0.1:${PORT}/mcp`;
  const http = new Client({ name: 'claude-code', version: '1.0.0' });
  await http.connect(new StreamableHTTPClientTransport(new URL(url), { requestInit: { headers: { Authorization: `Bearer ${token}` } } }));
  const st = await call('status', {}, http);
  assert.equal(st.data.configured, true);
  assert.equal(st.data.synced_folders[0].local_path, WS());
  await http.close();

  const body = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/list' });
  const headers = { 'Content-Type': 'application/json', Accept: 'application/json, text/event-stream' };
  assert.equal((await fetch(url, { method: 'POST', headers, body })).status, 401, 'no key');
  assert.equal((await fetch(url, { method: 'POST', headers: { ...headers, Authorization: 'Bearer wrong' }, body })).status, 401, 'wrong key');
  assert.equal((await fetch(url, { method: 'POST', headers: { ...headers, Authorization: `Bearer ${token}`, Origin: 'https://evil.example.com' }, body })).status, 403, 'web pages refused');

  // Our own page's forms: some browsers send "Origin: null" with Sec-Fetch-Site same-origin.
  const form = (headers) => fetch(`http://127.0.0.1:${PORT}/workspace/forget`, { method: 'POST', redirect: 'manual', headers: { 'Content-Type': 'application/x-www-form-urlencoded', ...headers }, body: 'localPath=%2Fnowhere' }).then((r) => r.status);
  assert.equal(await form({ Origin: 'null', 'Sec-Fetch-Site': 'same-origin' }), 303, 'own form with Origin null');
  assert.equal(await form({ Origin: `http://localhost:${PORT}`, 'Sec-Fetch-Site': 'same-origin' }), 303, 'own form');
  assert.equal(await form({ Origin: 'null' }), 403, 'opaque origin without fetch metadata');
  assert.equal(await form({ Origin: 'https://evil.example.com', 'Sec-Fetch-Site': 'cross-site' }), 403, 'other site');
  assert.equal(await form({ 'Sec-Fetch-Site': 'cross-site' }), 403, 'other site without Origin');

  const page = await (await fetch(`http://127.0.0.1:${PORT}/`)).text();
  assert.match(page, /Checked out to you/);
  assert.match(page, /Synced folders/);
});

test('firm setup creates the Entra app, approves it, and limits it to named people', async () => {
  const c = new Client({ name: 'admin', version: '1.0.0' });
  const adminEnv = { ...env, M365_TEST_ADMIN_TOKEN: 'admin-token', M365_BRIDGE_HOME: path.join(tmp, 'admin-home') };
  delete adminEnv.M365_TENANT_ID;
  delete adminEnv.M365_CLIENT_ID;
  await c.connect(new StdioClientTransport({ command: process.execPath, args: [CLI, 'serve'], env: adminEnv, stderr: 'inherit' }));
  const before = await call('status', {}, c);
  assert.equal(before.data.configured, false);
  assert.match(before.data.next, /setup code/);

  const r = await call('firm_setup', { firm_name: 'Henderson Law', access: 'selected', people: ['maria@mockfirm.com', 'nobody@mockfirm.com'] }, c);
  assert.ok(!r.isError, r.text);
  assert.equal(r.data.done, true);
  assert.match(r.data.setup_code, /^M365B-/);
  assert.deepEqual(r.data.not_found, ['nobody@mockfirm.com']);
  const app = mock.entra.apps[0];
  assert.equal(app.displayName, 'Claude M365 Bridge');
  assert.equal(app.isFallbackPublicClient, true);
  assert.ok(app.requiredResourceAccess[0].resourceAccess.length > 20);
  assert.equal(mock.entra.grants[0].consentType, 'AllPrincipals');
  assert.match(mock.entra.grants[0].scope, /Files\.ReadWrite\.All/);
  assert.equal(mock.entra.sps[0].appRoleAssignmentRequired, true);
  assert.deepEqual(mock.entra.assignments.map((a) => a.principalId).sort(), ['U-admin', 'U-maria']);

  const after = await call('status', {}, c);
  assert.equal(after.data.configured, true);
  assert.equal(after.data.firm, 'Henderson Law');

  // A teammate pastes the code.
  const t = new Client({ name: 'teammate', version: '1.0.0' });
  const teamEnv = { ...env, M365_BRIDGE_HOME: path.join(tmp, 'team-home') };
  delete teamEnv.M365_TENANT_ID;
  delete teamEnv.M365_CLIENT_ID;
  await t.connect(new StdioClientTransport({ command: process.execPath, args: [CLI, 'serve'], env: teamEnv, stderr: 'inherit' }));
  const s = await call('setup', { setup_code: r.data.setup_code }, t);
  assert.match(s.text, /Connected to Henderson Law/);
  assert.equal((await call('status', {}, t)).data.configured, true);
  await c.close();
  await t.close();
});

test('audit log records actions, never contents', async () => {
  const log = fs.readFileSync(path.join(tmp, 'home', 'audit.log'), 'utf8');
  assert.match(log, /workspace_push/);
  assert.match(log, /workspace_lock/);
  assert.doesNotMatch(log, /Retainer/);
});
