#!/usr/bin/env node
import fs from 'node:fs';
import { HOME, DASHBOARD_URL } from '../src/paths.js';

const cmd = process.argv[2] || 'dashboard';

async function main() {
  switch (cmd) {
    case 'serve': {
      // MCP over stdio. Nothing may print to stdout except the protocol.
      const { startMcp } = await import('../src/mcp.js');
      return startMcp();
    }
    case 'dashboard': {
      const { ensureDashboard, openBrowser } = await import('../src/dashboard/server.js');
      await ensureDashboard();
      console.log(`M365 Bridge dashboard: ${DASHBOARD_URL}`);
      console.log('Keep this window open while you set things up. Press Ctrl+C to stop.');
      if (!process.argv.includes('--no-open')) openBrowser(DASHBOARD_URL);
      return;
    }
    case 'install': {
      const { registerClaudeCode, registerClaudeDesktop, registrationStatus } = await import('../src/register.js');
      console.log(registerClaudeCode());
      if (registrationStatus().desktopInstalled) console.log(registerClaudeDesktop());
      const { ensureDashboard, openBrowser } = await import('../src/dashboard/server.js');
      await ensureDashboard();
      console.log(`\nNext: finish setup at ${DASHBOARD_URL} (opening it now).`);
      console.log('Keep this window open while you set things up. Press Ctrl+C when done.');
      openBrowser(DASHBOARD_URL);
      return;
    }
    case 'status': {
      const { loadConfig, isConfigured } = await import('../src/config.js');
      const { currentAccount } = await import('../src/auth.js');
      const { registrationStatus } = await import('../src/register.js');
      const { keyStorage } = await import('../src/vault.js');
      const cfg = loadConfig();
      const acct = await currentAccount();
      console.log(JSON.stringify({
        configured: isConfigured(cfg),
        firm: cfg.firmName,
        signedInAs: acct?.username || null,
        keyStorage: keyStorage(),
        claude: registrationStatus(),
        workspaces: cfg.workspaces.map((w) => w.localPath),
        dataFolder: HOME,
      }, null, 2));
      return;
    }
    case 'reset': {
      const { keychainDelete } = await import('../src/vault.js');
      keychainDelete();
      fs.rmSync(HOME, { recursive: true, force: true });
      console.log('Removed all M365 Bridge settings and tokens from this computer. Your workspace folders were not touched.');
      return;
    }
    default:
      console.log(`m365-bridge <command>

  dashboard   Open the setup dashboard (default)
  install     Connect to Claude Code + Claude Desktop, then open the dashboard
  serve       Run the MCP server (Claude starts this for you)
  status      Show setup state
  reset       Delete settings and tokens from this computer`);
  }
}

main().catch((e) => {
  console.error(e.message);
  process.exit(1);
});
