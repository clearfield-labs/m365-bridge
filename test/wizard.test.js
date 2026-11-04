// The admin wizard end to end through the real page: permissions -> one approval ->
// choose people -> drafts with personal setup codes -> send -> overview.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMockGraph } from './mock-graph.js';

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'm365w-'));
const PORT = '47420';
const BASE = `http://127.0.0.1:${PORT}`;
const mock = createMockGraph();
let parseSetupCode;

before(async () => {
  const base = await mock.listen();
  Object.assign(process.env, {
    M365_GRAPH_BASE: `${base}/v1.0`,
    M365_TEST_ADMIN_TOKEN: 'admin-token',
    M365_BRIDGE_HOME: path.join(tmp, 'home'),
    M365_BRIDGE_KEYSTORE: 'file',
    M365_WORKSPACES: path.join(tmp, 'ws'),
    M365_BRIDGE_PORT: PORT,
    M365_NO_SERVICE: '1',
    M365_NO_BROWSER: '1',
    M365_DEMO: '1', // never touch the real Claude config
  });
  ({ parseSetupCode } = await import('../src/config.js'));
  const { startDashboard } = await import('../src/dashboard/server.js');
  await startDashboard();
});

after(async () => {
  await mock.close();
  try { fs.rmSync(tmp, { recursive: true, force: true }); } catch { /* ignore */ }
  setTimeout(() => process.exit(0), 50); // the in-process page server keeps the loop alive
});

const page = async (p = '/admin') => (await fetch(BASE + p)).text();
async function post(p, fields) {
  const body = new URLSearchParams();
  for (const [k, v] of Object.entries(fields)) for (const x of [].concat(v)) body.append(k, x);
  const r = await fetch(BASE + p, { method: 'POST', redirect: 'manual', headers: { 'Content-Type': 'application/x-www-form-urlencoded', Origin: `http://localhost:${PORT}`, 'Sec-Fetch-Site': 'same-origin' }, body });
  assert.equal(r.status, 303, `${p} -> ${r.status}`);
  return r;
}
async function waitJob() {
  for (let i = 0; i < 100; i++) {
    const j = (await (await fetch(`${BASE}/api/setup-job`)).json()).data;
    if (j && !j.running) return j;
    await new Promise((r) => setTimeout(r, 100));
  }
  throw new Error('job did not finish');
}

test('step 1: permissions come first', async () => {
  const html = await page();
  assert.match(html, /What should Claude be able to reach\?/);
  assert.match(html, /Optional extras/);
  assert.match(html, /Continue and approve with Microsoft/);
});

test('step 2: one approval creates the app with exactly the chosen permissions', async () => {
  await post('/admin/start', { firmName: 'Henderson Law', pack_mail: 'on', pack_groups: 'on', allowWrite: 'on' });
  const j = await waitJob();
  assert.equal(j.error, null);
  const scopes = mock.entra.grants[0].scope.split(' ');
  assert.ok(scopes.includes('Group.ReadWrite.All'), 'optional pack included');
  assert.ok(scopes.includes('Mail.Send'));
  assert.ok(!scopes.includes('Calendars.ReadWrite'), 'unticked pack left out');
});

test('step 3: the team appears; choose people and give them access', async () => {
  const html = await page();
  assert.match(html, /Choose people/);
  assert.match(html, /Maria Lopez/);
  assert.match(html, /Evan Clarke/);
  assert.ok(mock.entra.assignments.some((a) => a.principalId === 'U-admin'), 'the admin gets access without picking themselves');
  assert.doesNotMatch(html, /value="U-admin"/, 'no checkbox for yourself');
  await post('/admin/access', { ids: ['U-maria', 'U-evan'], restricted: 'on' });
  assert.deepEqual(mock.entra.assignments.map((a) => a.principalId).sort(), ['U-admin', 'U-evan', 'U-maria']);
  assert.equal(mock.entra.sps[0].appRoleAssignmentRequired, true);
});

test('step 4: drafts in the admin Outlook, each with a personal setup code, then send', async () => {
  const html = await page();
  assert.match(html, /Invite 2 people/);
  assert.match(html, /Preview for/);
  await post('/admin/drafts', { subject: 'Claude now works with our Microsoft 365', message: 'Hi {first_name}, your code: {setup_code}' });
  assert.equal(mock.drafts.length, 2);
  const maria = mock.drafts.find((d) => d.toRecipients[0].emailAddress.address === 'maria@mockfirm.com');
  assert.match(maria.body.content, /Hi Maria/);
  const code = /M365B-[A-Za-z0-9_-]+/.exec(maria.body.content)[0];
  const parsed = parseSetupCode(code);
  assert.equal(parsed.loginHint, 'maria@mockfirm.com', 'personal code pre-fills her sign-in');
  assert.equal(parsed.firmName, 'Henderson Law');
  const { renderInvite } = await import('../src/provision.js');
  const inv = renderInvite({ person: { name: 'Maria Lopez', upn: 'maria@mockfirm.com' }, cfg: { tenantId: 't', clientId: 'c', firmName: 'X', packs: [] }, adminName: 'Admin' });
  assert.match(inv.html, /href="https:\/\/clearfield-labs\.github\.io\/m365-bridge\/#code=M365B-/, 'default email links to the install page with the code after #');
  assert.match(await page(), /Send all 2/);
  await post('/admin/send', {});
  assert.ok(mock.drafts.every((d) => d.sent));
});

test('if the admin sign-in has ended, Give access signs in first and then finishes', async () => {
  const { endAdminSession } = await import('../src/provision.js');
  endAdminSession();
  await post('/admin/access', { ids: ['U-evan'], restricted: 'on' });
  const j = await waitJob();
  assert.equal(j.error, null);
  assert.match(await page(), /Invite Evan Clarke/);
  await post('/admin/back', {});
});

test('afterwards: the admin page shows the setup and who was invited; add someone later', async () => {
  let html = await page();
  assert.match(html, /Henderson Law is set up/);
  assert.match(html, /Has access/);
  assert.match(html, /Invited/);
  assert.match(html, /Change permissions/);
  await post('/admin/refresh', {});
  await post('/admin/access', { ids: ['U-test'], restricted: 'on' });
  assert.ok(mock.entra.assignments.some((a) => a.principalId === 'U-test'));
  html = await page();
  assert.match(html, /Invite Test User/);
});
