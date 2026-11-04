#!/usr/bin/env node
// Demo mode: the real dashboard and tools against a fake Microsoft 365 with a seeded
// matter folder. No tenant needed. Nothing touches your real settings or Keychain.
//   npm run demo
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMockGraph } from '../test/mock-graph.js';

const home = path.join(os.tmpdir(), 'm365-bridge-demo');
fs.rmSync(home, { recursive: true, force: true });

const mock = createMockGraph();
const base = await mock.listen();
Object.assign(process.env, {
  M365_GRAPH_BASE: `${base}/v1.0`,
  M365_TEST_TOKEN: 'test-token',
  M365_TEST_USER: 'jordan.avery@demofirm.com',
  M365_TEST_NAME: 'Jordan Avery',
  M365_BRIDGE_HOME: path.join(home, 'data'),
  M365_BRIDGE_KEYSTORE: 'file',
  M365_WORKSPACES: path.join(home, 'M365 Workspaces'),
  M365_BRIDGE_PORT: process.env.M365_BRIDGE_PORT || '4366',
  M365_DEMO: '1',
});

const { updateConfig } = await import('../src/config.js');
const ws = await import('../src/workspace.js');
const { audit } = await import('../src/audit.js');
const { startDashboard, openBrowser } = await import('../src/dashboard/server.js');

updateConfig({
  tenantId: '9f3c2a1e-4b7d-4e2a-9c1f-2d8e6a5b4c3d',
  clientId: '5e1d7c9b-2a3f-4c6e-8b0d-1f2e3a4b5c6d',
  firmName: 'Demo Firm LLP',
});

// Simulate a morning of work so the dashboard has something to show.
const matter = `${mock.HOST}/Matters/Henderson v Acme`;
audit({ tool: 'sign_in', user: 'jordan.avery@demofirm.com' });
audit({ tool: 'search', query: 'Henderson engagement letter', results: 3 });
const pulled = await ws.pull({ source: matter });
fs.appendFileSync(path.join(pulled.workspace, 'Timeline.md'), '- 2026-02-01 demand letter sent\n');
await ws.push({ localPath: pulled.workspace });
audit({ tool: 'mail_search', query: 'settlement offer', results: 4 });
audit({ tool: 'mail_draft', subject: 'Re: Henderson settlement offer' });
mock.remoteEdit('Matters/Henderson v Acme/Pleadings/Complaint.txt', 'COMPLAINT draft v2', 'Sarah Chen');

const url = await startDashboard();
console.log(`Demo dashboard: ${url}`);
console.log(`Demo workspace: ${pulled.workspace}`);
console.log('Fake Microsoft 365 is running. Press Ctrl+C to stop.');
if (!process.argv.includes('--no-open')) openBrowser(url);
