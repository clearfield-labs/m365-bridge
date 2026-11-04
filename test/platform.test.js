// Platform checks that touch the operating system: the key store (Keychain / DPAPI), the
// start-at-login service, writing Claude's config, and launching the hook through a shell.
// They run only with M365_PLATFORM_TEST=1 (CI), always inside a throwaway home folder.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn, execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ON = process.env.M365_PLATFORM_TEST === '1';
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const CLI = path.join(ROOT, 'bin', 'm365-bridge.js');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'm365p-'));
const HOME = path.join(tmp, 'home');
const PORT = '47410';
const env = {
  ...process.env,
  HOME, USERPROFILE: HOME, APPDATA: path.join(HOME, 'AppData', 'Roaming'),
  M365_BRIDGE_HOME: path.join(HOME, '.m365-bridge'),
  M365_BRIDGE_PORT: PORT,
  M365_WORKSPACES: path.join(HOME, 'M365 Workspaces'),
  M365_NO_BROWSER: '1',
};
delete env.M365_BRIDGE_KEYSTORE;
// The Mac keychain needs the real home folder, so the install checks below use the key file there.
if (process.platform === 'darwin') env.M365_BRIDGE_KEYSTORE = 'file';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function run(args, extra = {}) {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, [CLI, ...args], { env: { ...env, ...extra } });
    let out = '';
    child.stdout.on('data', (d) => { out += d; });
    child.stderr.on('data', (d) => { out += d; });
    child.on('close', (code) => resolve({ code, out }));
  });
}
async function up(ms = 15000) {
  for (const end = Date.now() + ms; Date.now() < end; await sleep(300)) {
    try { if ((await fetch(`http://127.0.0.1:${PORT}/healthz`)).ok) return true; } catch { /* not yet */ }
  }
  return false;
}

before(() => fs.mkdirSync(env.APPDATA, { recursive: true }));
after(async () => {
  if (ON) await run(['uninstall']);
  if (ON && process.platform === 'darwin') {
    try { execFileSync(process.execPath, ['--input-type=module', '-e', `import('${pathToUrl(path.join(ROOT, 'src', 'vault.js'))}').then((v) => v.keychainDelete())`], { env: keyEnv, timeout: 30000 }); } catch { /* best effort */ }
  }
  try { fs.rmSync(tmp, { recursive: true, force: true }); } catch { /* Windows may hold a file briefly */ }
});

// The Mac keychain lives under the real home folder; point only the key-store check there.
// The key's keychain entry is named after the throwaway data folder, so it never touches a real one.
const keyEnv = process.platform === 'darwin' ? { ...env, HOME: os.homedir(), M365_BRIDGE_HOME: path.join(tmp, 'keytest'), M365_BRIDGE_KEYSTORE: '' } : env;

test('OS key store encrypts, and a second process can decrypt', { skip: !ON }, async () => {
  const script = (code) => ['--input-type=module', '-e', `import('${pathToUrl(path.join(ROOT, 'src', 'config.js'))}').then(async (c) => { ${code} })`];
  const write = await new Promise((resolve) => {
    const c = spawn(process.execPath, script(`c.updateConfig({ firmName: 'Platform Test LLP' }); const v = await import('${pathToUrl(path.join(ROOT, 'src', 'vault.js'))}'); console.log(v.keyStorage());`), { env: keyEnv });
    let o = ''; c.stdout.on('data', (d) => { o += d; }); c.stderr.on('data', (d) => { o += d; }); c.on('close', () => resolve(o.trim()));
  });
  console.log('key storage:', write);
  if (process.platform === 'win32') assert.equal(write, 'Windows DPAPI');
  if (process.platform === 'darwin') assert.equal(write, 'macOS Keychain');
  const read = await new Promise((resolve) => {
    const c = spawn(process.execPath, script('console.log(c.loadConfig().firmName)'), { env: keyEnv });
    let o = ''; c.stdout.on('data', (d) => { o += d; }); c.on('close', () => resolve(o.trim()));
  });
  assert.equal(read, 'Platform Test LLP');
  assert.doesNotMatch(fs.readFileSync(path.join(keyEnv.M365_BRIDGE_HOME, 'config.enc'), 'utf8'), /Platform Test/);
});

test('install: starts in the background, connects Claude, adds the hook', { skip: !ON }, async () => {
  const r = await run(['install', '--no-open']);
  console.log(r.out);
  assert.equal(r.code, 0, r.out);
  assert.ok(await up(), 'background service answers');
  const claude = JSON.parse(fs.readFileSync(path.join(HOME, '.claude.json'), 'utf8')).mcpServers.m365;
  assert.equal(claude.type, 'http');
  assert.match(claude.headers.Authorization, /^Bearer .{20,}/);
  const res = await fetch(claude.url, { method: 'POST', headers: { 'Content-Type': 'application/json', Accept: 'application/json, text/event-stream', Authorization: claude.headers.Authorization }, body: JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'ci', version: '1' } } }) });
  assert.equal(res.status, 200, 'Claude Code can connect with the key');
  if (process.platform === 'win32') assert.ok(fs.existsSync(path.join(env.APPDATA, 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup', 'M365 Bridge.vbs')));
});

test('the hook command launches from the shells Claude Code uses, and exits fast', { skip: !ON }, async () => {
  const settings = JSON.parse(fs.readFileSync(path.join(HOME, '.claude', 'settings.json'), 'utf8'));
  const cmd = settings.hooks.PreToolUse.find((m) => m.matcher.includes('Edit')).hooks[0].command;
  const input = JSON.stringify({ tool_name: 'Edit', tool_input: { file_path: path.join(tmp, 'outside.txt') }, cwd: tmp });
  const shells = process.platform === 'win32' ? [['bash', ['-c', cmd]], ['cmd.exe', ['/d', '/s', '/c', `"${cmd}"`]]] : [['/bin/sh', ['-c', cmd]]];
  for (const [sh, args] of shells) {
    const t = Date.now();
    const r = await new Promise((resolve) => {
      const c = spawn(sh, args, { env, windowsVerbatimArguments: sh === 'cmd.exe' });
      let o = ''; c.stdout.on('data', (d) => { o += d; }); c.stderr.on('data', (d) => { o += d; });
      c.on('close', (code) => resolve({ code, o }));
      c.on('error', (e) => resolve({ code: -1, o: e.message }));
      c.stdin.end(input);
    });
    console.log(sh, r.code, `${Date.now() - t} ms`, r.o);
    assert.equal(r.code, 0, `${sh}: ${r.o}`);
    assert.equal(r.o, '', 'files outside synced folders pass silently');
    assert.ok(Date.now() - t < 5000, 'fast');
  }
});

test('uninstall stops it and disconnects Claude', { skip: !ON }, async () => {
  const r = await run(['uninstall']);
  assert.equal(r.code, 0, r.out);
  await sleep(1500);
  assert.equal(await up(1000), false, 'stopped');
  assert.equal(JSON.parse(fs.readFileSync(path.join(HOME, '.claude.json'), 'utf8')).mcpServers.m365, undefined);
});

function pathToUrl(p) {
  return new URL(`file://${p.startsWith('/') ? '' : '/'}${p.replace(/\\/g, '/')}`).href;
}
