// Add the MCP server to Claude Code (CLI + desktop share ~/.claude.json) and Claude Desktop.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { CLI_PATH } from './paths.js';

export const SERVER_NAME = 'm365';

export function serverSpec() {
  // Absolute node path: GUI apps on macOS do not inherit the shell PATH.
  return { type: 'stdio', command: process.execPath, args: [CLI_PATH, 'serve'], env: {} };
}

const claudeCodeConfig = () => path.join(os.homedir(), '.claude.json');

export function claudeDesktopConfig() {
  if (process.platform === 'darwin') return path.join(os.homedir(), 'Library', 'Application Support', 'Claude', 'claude_desktop_config.json');
  if (process.platform === 'win32') return path.join(process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'), 'Claude', 'claude_desktop_config.json');
  return path.join(os.homedir(), '.config', 'Claude', 'claude_desktop_config.json');
}

function readJson(file) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return null; }
}

function findClaudeCli() {
  try {
    const cmd = process.platform === 'win32' ? 'where' : 'which';
    return execFileSync(cmd, ['claude'], { stdio: ['ignore', 'pipe', 'ignore'] }).toString().split(/\r?\n/)[0].trim() || null;
  } catch {
    return null;
  }
}

// Demo mode never touches the real Claude config.
const demo = () => process.env.M365_DEMO === '1';
const demoState = { code: false, desktop: false };

export function registrationStatus() {
  if (demo()) return { claudeCode: demoState.code ? 'connected' : 'not connected', claudeDesktop: demoState.desktop ? 'connected' : 'not connected', desktopInstalled: true };
  const code = readJson(claudeCodeConfig())?.mcpServers?.[SERVER_NAME];
  const desktop = readJson(claudeDesktopConfig())?.mcpServers?.[SERVER_NAME];
  const current = (s) => !!s && s.args?.[0] === CLI_PATH;
  return {
    claudeCode: code ? (current(code) ? 'connected' : 'points to another copy') : 'not connected',
    claudeDesktop: desktop ? (current(desktop) ? 'connected' : 'points to another copy') : 'not connected',
    desktopInstalled: fs.existsSync(path.dirname(claudeDesktopConfig())),
  };
}

export function registerClaudeCode() {
  if (demo()) { demoState.code = true; return 'Demo: connected to Claude Code (your real Claude settings were not changed).'; }
  const spec = serverSpec();
  const cli = findClaudeCli();
  if (cli) {
    try { execFileSync(cli, ['mcp', 'remove', '--scope', 'user', SERVER_NAME], { stdio: 'ignore' }); } catch { /* absent */ }
    execFileSync(cli, ['mcp', 'add-json', '--scope', 'user', SERVER_NAME, JSON.stringify(spec)], { stdio: 'ignore' });
    return 'Added to Claude Code with the claude CLI (user scope, every project).';
  }
  // No CLI on PATH (desktop-only install): edit ~/.claude.json directly.
  const file = claudeCodeConfig();
  const json = readJson(file) || {};
  if (fs.existsSync(file)) fs.copyFileSync(file, `${file}.m365-bridge.bak`);
  json.mcpServers = { ...(json.mcpServers || {}), [SERVER_NAME]: spec };
  fs.writeFileSync(file, JSON.stringify(json, null, 2));
  return `Added to ${file}. Restart Claude to load it.`;
}

export function registerClaudeDesktop() {
  if (demo()) { demoState.desktop = true; return 'Demo: connected to Claude Desktop.'; }
  const file = claudeDesktopConfig();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const json = readJson(file) || {};
  if (fs.existsSync(file)) fs.copyFileSync(file, `${file}.m365-bridge.bak`);
  const { type, ...spec } = serverSpec();
  json.mcpServers = { ...(json.mcpServers || {}), [SERVER_NAME]: spec };
  fs.writeFileSync(file, JSON.stringify(json, null, 2));
  return `Added to Claude Desktop (${file}). Quit and reopen Claude Desktop to load it.`;
}
