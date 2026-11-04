import os from 'node:os';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';

export const HOME = process.env.M365_BRIDGE_HOME || path.join(os.homedir(), '.m365-bridge');
export const PORT = Number(process.env.M365_BRIDGE_PORT || 4365);
export const WORKSPACES_ROOT = process.env.M365_WORKSPACES || path.join(os.homedir(), 'M365 Workspaces');
export const REDIRECT_URI = `http://localhost:${PORT}/auth/callback`;
export const DASHBOARD_URL = `http://localhost:${PORT}/`;
export const MCP_URL = `http://127.0.0.1:${PORT}/mcp`;
export const PACKAGE_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const CLI_PATH = path.join(PACKAGE_ROOT, 'bin', 'm365-bridge.js');

/**
 * A node path that survives upgrades. Homebrew's process.execPath points into a versioned
 * Cellar folder that disappears on the next "brew upgrade"; its opt/ link does not.
 */
export function stableNodePath(execPath = process.execPath) {
  const m = /^(.*)\/Cellar\/([^/]+)\/[^/]+\/bin\/node$/.exec(execPath);
  if (m) {
    const opt = `${m[1]}/opt/${m[2]}/bin/node`;
    if (fs.existsSync(opt)) return opt;
    if (fs.existsSync(`${m[1]}/bin/node`)) return `${m[1]}/bin/node`;
  }
  return execPath;
}
export const NODE_PATH = stableNodePath();

export const FILES = {
  config: path.join(HOME, 'config.enc'),
  tokens: path.join(HOME, 'tokens.enc'),
  audit: path.join(HOME, 'audit.log'),
  keyFile: path.join(HOME, 'master.key'),
  dpapiKey: path.join(HOME, 'master.key.dpapi'),
  daemonLog: path.join(HOME, 'background.log'),
};

export function ensureHome() {
  fs.mkdirSync(HOME, { recursive: true, mode: 0o700 });
}

/** Write a file atomically with owner-only permissions. */
export function writePrivate(file, data) {
  ensureHome();
  const tmp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, data, { mode: 0o600 });
  fs.renameSync(tmp, file);
}
