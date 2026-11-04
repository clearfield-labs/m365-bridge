// End to end: real MCP server over stdio, talking to the mock Graph.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { createMockGraph } from './mock-graph.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'm365b-'));
const WS_ROOT = path.join(tmp, 'workspaces');
const mock = createMockGraph();
let client;

async function call(name, args = {}) {
  const r = await client.callTool({ name, arguments: args });
  const text = r.content[0].text;
  let data;
  try { data = JSON.parse(text); } catch { data = text; }
  return { data, isError: !!r.isError, text };
}

before(async () => {
  const base = await mock.listen();
  client = new Client({ name: 'test', version: '1.0.0' });
  await client.connect(
    new StdioClientTransport({
      command: process.execPath,
      args: [path.join(ROOT, 'bin', 'm365-bridge.js'), 'serve'],
      env: {
        ...process.env,
        M365_GRAPH_BASE: `${base}/v1.0`,
        M365_TEST_TOKEN: 'test-token',
        M365_BRIDGE_HOME: path.join(tmp, 'home'),
        M365_BRIDGE_KEYSTORE: 'file',
        M365_WORKSPACES: WS_ROOT,
        M365_BRIDGE_PORT: '47365',
        M365_TENANT_ID: '11111111-1111-1111-1111-111111111111',
        M365_CLIENT_ID: '22222222-2222-2222-2222-222222222222',
        M365_SIMPLE_UPLOAD_MAX: '200',
      },
      stderr: 'inherit',
    }),
  );
});

after(async () => {
  await client?.close();
  await mock.close();
  fs.rmSync(tmp, { recursive: true, force: true });
});

const WS = () => path.join(WS_ROOT, 'Henderson v Acme');
const MATTER_URL = () => `${mock.HOST}/Matters/Henderson v Acme`;

test('tools are registered', async () => {
  const { tools } = await client.listTools();
  const names = tools.map((t) => t.name);
  for (const n of ['status', 'sign_in', 'search', 'list_folder', 'read_file', 'upload', 'versions', 'restore_version', 'workspace_pull', 'workspace_status', 'workspace_push', 'mail_search', 'mail_read', 'mail_draft', 'calendar_view', 'graph_request']) {
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
  const byUrl = await call('read_file', { location: `${MATTER_URL()}/Timeline.md` });
  assert.match(byUrl.data.text, /incident/);
});

test('list_folder by SharePoint AllItems.aspx address-bar link', async () => {
  const id = encodeURIComponent('/sites/Lit/Shared Documents/Matters/Henderson v Acme');
  const { data, isError, text } = await call('list_folder', { location: `https://mock.sharepoint.com/sites/Lit/Shared%20Documents/Forms/AllItems.aspx?id=${id}` });
  assert.ok(!isError, text);
  assert.ok(data.items.some((i) => i.name === 'Pleadings' && i.type === 'folder'));
});

test('workspace pull clones the folder and writes Claude instructions', async () => {
  const { data, isError, text } = await call('workspace_pull', { source: MATTER_URL() });
  assert.ok(!isError, text);
  assert.equal(data.workspace, WS());
  assert.equal(data.downloaded.length, 3);
  assert.ok(fs.existsSync(path.join(WS(), 'Pleadings', 'Complaint.txt')));
  assert.ok(!fs.existsSync(path.join(WS(), '~$Lockfile.txt')), 'office lock files are ignored');
  assert.match(fs.readFileSync(path.join(WS(), '.claude', 'CLAUDE.md'), 'utf8'), /source of truth/i);
});

test('status, then push a local edit as a new SharePoint version', async () => {
  fs.appendFileSync(path.join(WS(), 'Timeline.md'), '- 2026-02-01 demand letter\n');
  const st = await call('workspace_status', { local_path: WS() });
  assert.deepEqual(st.data.modified, ['Timeline.md']);
  const push = await call('workspace_push', { local_path: WS() });
  assert.equal(push.data.uploaded.length, 1);
  assert.match(mock.find('Matters/Henderson v Acme/Timeline.md').content.toString(), /demand letter/);
  const st2 = await call('workspace_status', { local_path: WS() });
  assert.equal(st2.data.summary, 'Everything matches SharePoint.');
});

test('pull brings down a colleague edit without touching local work', async () => {
  mock.remoteEdit('Matters/Henderson v Acme/Pleadings/Complaint.txt', 'COMPLAINT draft v2 (partner edits)');
  const st = await call('workspace_status', { local_path: WS() });
  assert.equal(st.data.changed_in_sharepoint[0].file, 'Pleadings/Complaint.txt');
  const pull = await call('workspace_pull', { local_path: WS() });
  assert.equal(pull.data.updated[0].file, 'Pleadings/Complaint.txt');
  assert.match(fs.readFileSync(path.join(WS(), 'Pleadings', 'Complaint.txt'), 'utf8'), /v2/);
});

test('conflict: push refuses to overwrite a colleague, pull saves their copy, merged push succeeds', async () => {
  const file = path.join(WS(), 'Engagement Letter.txt');
  fs.writeFileSync(file, 'Engagement letter for Henderson. Fee: flat $5,000.');
  mock.remoteEdit('Matters/Henderson v Acme/Engagement Letter.txt', 'Engagement letter for Henderson. Fee: hourly. Retainer $2,500.');

  const push = await call('workspace_push', { local_path: WS() });
  assert.equal(push.data.conflicts.length, 1);
  assert.match(mock.find('Matters/Henderson v Acme/Engagement Letter.txt').content.toString(), /Retainer/, 'colleague work kept');

  const pull = await call('workspace_pull', { local_path: WS() });
  assert.equal(pull.data.conflicts.length, 1);
  const copy = path.join(WS(), pull.data.conflicts[0].sharepoint_copy);
  assert.match(fs.readFileSync(copy, 'utf8'), /Retainer/);
  assert.match(fs.readFileSync(file, 'utf8'), /flat/, 'local edit kept');

  fs.writeFileSync(file, 'Engagement letter for Henderson. Fee: flat $5,000. Retainer $2,500.');
  fs.rmSync(copy);
  const push2 = await call('workspace_push', { local_path: WS() });
  assert.equal(push2.data.conflicts.length, 0, JSON.stringify(push2.data));
  assert.equal(push2.data.uploaded.length, 1);
  assert.match(mock.find('Matters/Henderson v Acme/Engagement Letter.txt').content.toString(), /flat \$5,000\. Retainer/);
});

test('new files in new subfolders are created; large files use an upload session', async () => {
  fs.mkdirSync(path.join(WS(), 'Discovery', 'Set 1'), { recursive: true });
  fs.writeFileSync(path.join(WS(), 'Discovery', 'Set 1', 'Interrogatories.txt'), 'Q'.repeat(1000));
  const push = await call('workspace_push', { local_path: WS() });
  assert.equal(push.data.created.length, 1, JSON.stringify(push.data));
  assert.equal(mock.find('Matters/Henderson v Acme/Discovery/Set 1/Interrogatories.txt').content.length, 1000);
  assert.ok(mock.log.some((l) => l.startsWith('PUT /upload/')));
});

test('deleting locally never deletes in SharePoint', async () => {
  fs.rmSync(path.join(WS(), 'Timeline.md'));
  const push = await call('workspace_push', { local_path: WS() });
  assert.deepEqual(push.data.not_pushed_deletions, ['Timeline.md']);
  assert.ok(mock.find('Matters/Henderson v Acme/Timeline.md'));
});

test('versions and restore_version', async () => {
  const loc = `${MATTER_URL()}/Engagement Letter.txt`;
  const v = await call('versions', { location: loc });
  assert.ok(v.data.versions.length >= 3);
  const first = v.data.versions.at(-1).version;
  const r = await call('restore_version', { location: loc, version: first });
  assert.ok(!r.isError, r.text);
  assert.match(mock.find('Matters/Henderson v Acme/Engagement Letter.txt').content.toString(), /^Engagement letter for Henderson\. Fee: hourly\.$/);
});

test('upload refuses to clobber unless replace is set', async () => {
  const src = path.join(tmp, 'Complaint.txt');
  fs.writeFileSync(src, 'local complaint');
  const r = await call('upload', { local_path: src, folder: `${MATTER_URL()}/Pleadings` });
  assert.ok(r.isError);
  assert.match(r.text, /already exists/);
  const r2 = await call('upload', { local_path: src, folder: `${MATTER_URL()}/Pleadings`, replace: true });
  assert.ok(!r2.isError, r2.text);
});

test('mail tools and graph_request', async () => {
  const s = await call('mail_search', { query: 'settlement' });
  assert.equal(s.data[0].subject, 'Henderson settlement offer');
  const m = await call('mail_read', { id: 'M1' });
  assert.match(m.data.body, /\$50,000/);
  const d = await call('mail_draft', { to: ['client@example.com'], subject: 'Update', body: 'Draft only' });
  assert.equal(d.data.draft_saved, true);
  assert.equal(mock.drafts.length, 1);

  const chats = await call('graph_request', { path: '/me/chats' });
  assert.match(chats.text, /Henderson team/);
  const del = await call('graph_request', { method: 'DELETE', path: '/me/messages/M1' });
  assert.ok(del.isError);
  const evil = await call('graph_request', { path: 'https://evil.example.com/steal' });
  assert.ok(evil.isError);
  assert.match(evil.text, /only talks to Microsoft Graph/);
});

test('audit log records actions, never contents', async () => {
  const log = fs.readFileSync(path.join(tmp, 'home', 'audit.log'), 'utf8');
  assert.match(log, /workspace_push/);
  assert.doesNotMatch(log, /Retainer/);
});
