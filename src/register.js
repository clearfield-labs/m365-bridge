// Connect M365 Bridge to Claude.
//   Claude Code (desktop app and terminal share ~/.claude.json): an HTTP MCP server pointing
//     at the background service on this computer, plus a hook that checks a SharePoint file
//     out before Claude edits it.
//   Claude Desktop (the chat app): a local stdio server, since it only runs local commands.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { CLI_PATH, MCP_URL, NODE_PATH } from './paths.js';
import { ensureMcpToken, loadConfig } from './config.js';

export const SERVER_NAME = 'm365';

const claudeCodeConfig = () => path.join(os.homedir(), '.claude.json');
const claudeSettings = () => path.join(os.homedir(), '.claude', 'settings.json');

export function claudeDesktopConfig() {
  if (process.platform === 'darwin') return path.join(os.homedir(), 'Library', 'Application Support', 'Claude', 'claude_desktop_config.json');
  if (process.platform === 'win32') return path.join(process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'), 'Claude', 'claude_desktop_config.json');
  return path.join(os.homedir(), '.config', 'Claude', 'claude_desktop_config.json');
}

function readJson(file) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return null; }
}
/** For editing: a missing file is {}, but a file that exists and does not parse is never overwritten. */
function readJsonForEdit(file) {
  if (!fs.existsSync(file)) return {};
  const json = readJson(file);
  if (!json) throw new Error(`${file} has a formatting error, so M365 Bridge left it alone. Fix it (or ask Claude to), then connect again.`);
  return json;
}
function writeJson(file, json) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  if (fs.existsSync(file)) fs.copyFileSync(file, `${file}.m365-bridge.bak`);
  const tmp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(json, null, 2));
  fs.renameSync(tmp, file);
}

function findClaudeCli() {
  try {
    const cmd = process.platform === 'win32' ? 'where' : 'which';
    const found = execFileSync(cmd, ['claude'], { stdio: ['ignore', 'pipe', 'ignore'] }).toString().split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
    // Node refuses to run .cmd/.bat files without a shell; on Windows use claude.exe or edit the config directly.
    return found.find((f) => process.platform !== 'win32' || /\.exe$/i.test(f)) || null;
  } catch {
    return null;
  }
}

// Windows: forward slashes in double quotes work in cmd, PowerShell and Git Bash alike.
const q = (s) => (process.platform === 'win32' ? `"${String(s).replace(/\\/g, '/')}"` : `'${String(s).replace(/'/g, `'\\''`)}'`);
export const httpSpec = (token) => ({ type: 'http', url: MCP_URL, headers: { Authorization: `Bearer ${token}` } });
export const stdioSpec = () => ({ command: NODE_PATH, args: [CLI_PATH, 'serve'], env: {} });
export const hookCommand = () => `${q(NODE_PATH)} ${q(CLI_PATH)} hook`;
// Before an edit: check the file out. After a reply: save. On a message: get latest, and the day.
const HOOKS = {
  PreToolUse: { matcher: 'Edit|Write|MultiEdit|NotebookEdit', timeout: 60 },
  Stop: { timeout: 120 },
  UserPromptSubmit: { timeout: 30 },
};
const isOurHook = (h) => typeof h?.command === 'string' && /m365-bridge\.js['"]? hook$/.test(h.command);

// Demo mode never touches the real Claude config.
const demo = () => process.env.M365_DEMO === '1';
const demoState = { code: false, desktop: false, hook: false };

export function registrationStatus() {
  if (demo()) return { claudeCode: demoState.code ? 'connected' : 'not connected', claudeDesktop: demoState.desktop ? 'connected' : 'not connected', desktopInstalled: true, hook: demoState.hook };
  const token = loadConfig().mcpToken;
  const code = readJson(claudeCodeConfig())?.mcpServers?.[SERVER_NAME];
  const desktop = readJson(claudeDesktopConfig())?.mcpServers?.[SERVER_NAME];
  const codeOk = code?.type === 'http' && code.url === MCP_URL && token && code.headers?.Authorization === `Bearer ${token}`;
  const all = readJson(claudeSettings())?.hooks || {};
  const has = (ev) => (all[ev] || []).some((m) => (m.hooks || []).some((h) => isOurHook(h) && h.command === hookCommand()));
  return {
    claudeCode: code ? (codeOk ? 'connected' : 'needs update') : 'not connected',
    claudeDesktop: desktop ? (desktop.args?.[0] === CLI_PATH ? 'connected' : 'needs update') : 'not connected',
    desktopInstalled: fs.existsSync(path.dirname(claudeDesktopConfig())),
    hook: Object.keys(HOOKS).every(has),
  };
}

export function registerClaudeCode() {
  if (demo()) { demoState.code = true; demoState.hook = true; return 'Demo: connected to Claude Code (your real Claude settings were not changed).'; }
  const spec = httpSpec(ensureMcpToken());
  const cli = findClaudeCli();
  let how;
  try {
    if (!cli) throw new Error('no cli');
    try { execFileSync(cli, ['mcp', 'remove', '--scope', 'user', SERVER_NAME], { stdio: 'ignore' }); } catch { /* absent */ }
    execFileSync(cli, ['mcp', 'add-json', '--scope', 'user', SERVER_NAME, JSON.stringify(spec)], { stdio: 'ignore' });
    how = 'with the claude command';
  } catch {
    // Desktop-only install, or an older CLI: edit ~/.claude.json directly.
    const file = claudeCodeConfig();
    const json = readJsonForEdit(file);
    json.mcpServers = { ...(json.mcpServers || {}), [SERVER_NAME]: spec };
    writeJson(file, json);
    how = `in ${file}`;
  }
  installHook();
  return `Connected to Claude Code ${how}. Restart Claude to load it.`;
}

/** A PreToolUse hook: before Claude edits a file in a synced folder, check it out in SharePoint. */
export function installHook() {
  const file = claudeSettings();
  const json = readJsonForEdit(file);
  json.hooks = json.hooks || {};
  for (const [event, { matcher, timeout }] of Object.entries(HOOKS)) {
    const list = withoutOurs(json.hooks[event]);
    list.push({ ...(matcher ? { matcher } : {}), hooks: [{ type: 'command', command: hookCommand(), timeout }] });
    json.hooks[event] = list;
  }
  writeJson(file, json);
}

const withoutOurs = (list) => (list || [])
  .map((m) => ({ ...m, hooks: (m.hooks || []).filter((h) => !isOurHook(h)) }))
  .filter((m) => m.hooks.length);

export function registerClaudeDesktop() {
  if (demo()) { demoState.desktop = true; return 'Demo: connected to Claude Desktop.'; }
  const file = claudeDesktopConfig();
  const json = readJsonForEdit(file);
  json.mcpServers = { ...(json.mcpServers || {}), [SERVER_NAME]: stdioSpec() };
  writeJson(file, json);
  return 'Connected to the Claude Desktop chat app. Quit and reopen it to load.';
}

export function unregisterAll() {
  if (demo()) { Object.assign(demoState, { code: false, desktop: false, hook: false }); return; }
  const cli = findClaudeCli();
  if (cli) { try { execFileSync(cli, ['mcp', 'remove', '--scope', 'user', SERVER_NAME], { stdio: 'ignore' }); } catch { /* absent */ } }
  for (const file of [claudeCodeConfig(), claudeDesktopConfig()]) {
    const json = readJson(file);
    if (json?.mcpServers?.[SERVER_NAME]) { delete json.mcpServers[SERVER_NAME]; writeJson(file, json); }
  }
  const settings = readJson(claudeSettings());
  if (settings?.hooks) {
    for (const event of Object.keys(HOOKS)) {
      if (!settings.hooks[event]) continue;
      settings.hooks[event] = withoutOurs(settings.hooks[event]);
      if (!settings.hooks[event].length) delete settings.hooks[event];
    }
    writeJson(claudeSettings(), settings);
  }
}
