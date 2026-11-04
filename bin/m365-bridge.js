#!/usr/bin/env node
import fs from 'node:fs';
import { spawn } from 'node:child_process';
import { HOME, DASHBOARD_URL, CLI_PATH, FILES, PORT, ensureHome } from '../src/paths.js';

const cmd = process.argv[2] || 'open';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function waitUntilUp(ms = 10000) {
  const { isRunning } = await import('../src/dashboard/server.js');
  for (const end = Date.now() + ms; Date.now() < end; await sleep(300)) {
    if (await isRunning()) return true;
  }
  return false;
}

/** Ask an older background copy to stop, so the new code takes over. */
async function stopRunning() {
  const { loadConfig } = await import('../src/config.js');
  const token = loadConfig().mcpToken;
  if (!token) return;
  try {
    await fetch(`http://127.0.0.1:${PORT}/api/shutdown`, { method: 'POST', headers: { Authorization: `Bearer ${token}` }, signal: AbortSignal.timeout(1500) });
    await sleep(500);
  } catch { /* not running */ }
}

function spawnBackground() {
  ensureHome();
  const log = fs.openSync(FILES.daemonLog, 'a');
  spawn(process.execPath, [CLI_PATH, 'background'], { detached: true, stdio: ['ignore', log, log], windowsHide: true }).unref();
}

async function main() {
  switch (cmd) {
    case 'serve': {
      // MCP over stdio (used by the Claude chat app). Nothing may print to stdout except the protocol.
      const { startMcp } = await import('../src/mcp.js');
      return startMcp();
    }
    case 'background': {
      // The always-on service: setup page + MCP endpoint for Claude Code.
      const { startDashboard, isRunning } = await import('../src/dashboard/server.js');
      try {
        await startDashboard();
        console.log(`${new Date().toISOString()} M365 Bridge running at ${DASHBOARD_URL} (pid ${process.pid})`);
      } catch (e) {
        if (e.code === 'EADDRINUSE' && (await isRunning())) {
          console.log(`${new Date().toISOString()} Already running; this copy exits.`);
          process.exit(0);
        }
        throw e;
      }
      return;
    }
    case 'hook': {
      const { runHook } = await import('../src/hook.js');
      await runHook();
      // Exit promptly (open network sockets would otherwise hold Claude's edit), after stdout drains.
      return process.stdout.write('', () => process.exit(0));
    }
    case 'open':
    case 'dashboard': {
      const { ensureDashboard, isRunning } = await import('../src/dashboard/server.js');
      const { openBrowser } = await import('../src/browser.js');
      const already = await isRunning();
      await ensureDashboard();
      console.log(`M365 Bridge: ${DASHBOARD_URL}`);
      if (!already) console.log('Running in this window. Close it with Ctrl+C. (Run "m365-bridge install" to keep it running in the background.)');
      const { pageUrl } = await import('../src/config.js');
      if (!process.argv.includes('--no-open')) openBrowser(pageUrl(DASHBOARD_URL));
      return;
    }
    case 'url': {
      // The unlocking address for this computer (used by the M365 Bridge app).
      const { pageUrl } = await import('../src/config.js');
      return console.log(pageUrl(DASHBOARD_URL, process.argv[3] || ''));
    }
    case 'install': {
      const { ensureMcpToken } = await import('../src/config.js');
      const { installService } = await import('../src/service.js');
      const { registerClaudeCode, registerClaudeDesktop, registrationStatus } = await import('../src/register.js');
      const { openBrowser } = await import('../src/browser.js');
      await stopRunning();
      ensureMcpToken();
      try {
        console.log(installService());
      } catch (e) {
        console.log(`Could not set up the background service (${e.message}). Starting it for now.`);
        spawnBackground();
      }
      console.log(registerClaudeCode());
      if (registrationStatus().desktopInstalled) console.log(registerClaudeDesktop());
      if (!(await waitUntilUp())) spawnBackground();
      if (!(await waitUntilUp())) {
        console.log(`M365 Bridge did not start. See ${FILES.daemonLog}`);
        process.exit(1);
      }
      console.log(`\nM365 Bridge is running at ${DASHBOARD_URL}`);
      console.log('Next: paste your setup code and sign in (opening it now). Then restart Claude.');
      const { pageUrl } = await import('../src/config.js');
      if (!process.argv.includes('--no-open')) openBrowser(pageUrl(DASHBOARD_URL));
      return;
    }
    case 'start': {
      const { isRunning } = await import('../src/dashboard/server.js');
      if (await isRunning()) return console.log(`Already running at ${DASHBOARD_URL}`);
      const { serviceStatus, installService } = await import('../src/service.js');
      if (serviceStatus().installed) installService();
      else spawnBackground();
      console.log((await waitUntilUp()) ? `Running at ${DASHBOARD_URL}` : `Did not start. See ${FILES.daemonLog}`);
      return;
    }
    case 'uninstall': {
      const { uninstallService } = await import('../src/service.js');
      const { unregisterAll } = await import('../src/register.js');
      await stopRunning();
      console.log(uninstallService());
      unregisterAll();
      console.log('Disconnected from Claude. Your settings and synced folders are still here (run "m365-bridge reset" to delete settings).');
      return;
    }
    case 'status': {
      const { loadConfig, isConfigured } = await import('../src/config.js');
      const { currentAccount } = await import('../src/auth.js');
      const { registrationStatus } = await import('../src/register.js');
      const { serviceStatus } = await import('../src/service.js');
      const { isRunning } = await import('../src/dashboard/server.js');
      const { keyStorage } = await import('../src/vault.js');
      const { heldFiles } = await import('../src/workspace.js');
      const cfg = loadConfig();
      const acct = await currentAccount();
      console.log(JSON.stringify({
        configured: isConfigured(cfg),
        firm: cfg.firmName,
        signedInAs: acct?.username || null,
        running: !!(await isRunning()),
        backgroundService: serviceStatus(),
        claude: registrationStatus(),
        checkedOutToYou: heldFiles().map((h) => `${h.workspace}/${h.file}`),
        keyStorage: keyStorage(),
        workspaces: cfg.workspaces.map((w) => w.localPath),
        dataFolder: HOME,
      }, null, 2));
      return;
    }
    case 'reset': {
      const { keychainDelete } = await import('../src/vault.js');
      keychainDelete();
      fs.rmSync(HOME, { recursive: true, force: true });
      console.log('Removed all M365 Bridge settings and sign-ins from this computer. Synced folders were not touched.');
      return;
    }
    default:
      console.log(`m365-bridge <command>

  (none)      Open the M365 Bridge page
  install     Run in the background from login, connect Claude, open the page
  start       Start the background service if it is not running
  status      Show setup state
  uninstall   Stop the background service and disconnect Claude
  reset       Delete settings and sign-ins from this computer
  serve       MCP over stdio (the Claude chat app starts this)
  background  The background service itself (started at login)
  hook        Claude Code hook that checks files out before edits`);
  }
}

main().catch((e) => {
  console.error(e.message);
  process.exit(1);
});
