import crypto from 'node:crypto';
import { FILES, writePrivate } from './paths.js';
import { readEncrypted, writeEncrypted } from './vault.js';
import { DEFAULT_PACKS } from './scopes.js';

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const DEFAULTS = {
  tenantId: '',
  clientId: '',
  firmName: '',
  allowWrite: true,
  allowDelete: false,
  packs: DEFAULT_PACKS,
  extraScopes: '',
  workspaces: [], // { name, localPath, webUrl, driveId, itemId, lastPulled }
  mcpToken: '', // shared secret between Claude and the background server on this computer
  loginHint: '', // from a personal setup code: pre-fills the Microsoft sign-in
  team: null, // admin only: last loaded directory { people, restricted, fetchedAt }
  invites: {}, // admin only: userId -> { at, sent }
  autoSave: true, // send Claude's edits to SharePoint when it finishes, hand back check-outs, refresh folders
  dayBrief: true, // tell Claude about the person's day (calendar, unread mail) at the start of a session
};

export function loadConfig() {
  let saved = {};
  const raw = readEncrypted(FILES.config);
  if (raw) {
    try { saved = JSON.parse(raw); } catch { saved = {}; }
  }
  const cfg = { ...DEFAULTS, ...saved };
  if (process.env.M365_TENANT_ID) cfg.tenantId = process.env.M365_TENANT_ID;
  if (process.env.M365_CLIENT_ID) cfg.clientId = process.env.M365_CLIENT_ID;
  return cfg;
}

export function saveConfig(cfg) {
  writeEncrypted(FILES.config, JSON.stringify(cfg));
  writeHookState(cfg);
}

/**
 * Claude's hooks run often (every message, every reply), so they read this small file instead of
 * unlocking the encrypted settings. It holds the same key that ~/.claude.json already holds for
 * Claude Code, and the synced folder paths. Owner-only permissions.
 */
export function writeHookState(cfg = loadConfig()) {
  try {
    const data = JSON.stringify({ token: cfg.mcpToken || '', workspaces: (cfg.workspaces || []).map((w) => w.localPath), dayBrief: cfg.dayBrief !== false });
    writePrivate(FILES.hookState, data);
  } catch { /* the hooks fall back to doing the work themselves */ }
}

export function updateConfig(patch) {
  const cfg = { ...loadConfig(), ...patch };
  saveConfig(cfg);
  return cfg;
}

/** The page key: unlocks the M365 Bridge page in a browser on this computer (see dashboard/server.js). */
export function ensurePageKey() {
  const cfg = loadConfig();
  if (cfg.pageKey) return cfg.pageKey;
  const key = crypto.randomBytes(24).toString('base64url');
  updateConfig({ pageKey: key });
  return key;
}

/** The address to open the page with, unlocking it in that browser. */
export const pageUrl = (base, p = '') => `${base}${p}${p.includes('?') ? '&' : '?'}k=${encodeURIComponent(ensurePageKey())}`;

/** The key Claude Code presents to the background server. Created once per computer. */
export function ensureMcpToken() {
  const cfg = loadConfig();
  if (cfg.mcpToken) return cfg.mcpToken;
  const token = crypto.randomBytes(32).toString('base64url');
  updateConfig({ mcpToken: token });
  return token;
}

export function isConfigured(cfg = loadConfig()) {
  return GUID.test(cfg.tenantId) && GUID.test(cfg.clientId);
}

export function validateIds(tenantId, clientId) {
  const errors = [];
  if (!GUID.test(tenantId)) errors.push('Tenant ID should look like 00000000-0000-0000-0000-000000000000.');
  if (!GUID.test(clientId)) errors.push('Client ID should look like 00000000-0000-0000-0000-000000000000.');
  return errors;
}

/** A setup code bundles the firm's non-secret IDs so an admin can hand them to the team in one paste. */
export function makeSetupCode(cfg, { user } = {}) {
  const payload = JSON.stringify({ t: cfg.tenantId, c: cfg.clientId, n: cfg.firmName, w: cfg.allowWrite, d: cfg.allowDelete, p: cfg.packs, x: cfg.extraScopes, ...(user ? { u: user } : {}) });
  return `M365B-${Buffer.from(payload).toString('base64url')}`;
}

export function parseSetupCode(code) {
  const body = String(code || '').trim().replace(/^M365B-/, '');
  const data = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
  return {
    tenantId: data.t,
    clientId: data.c,
    firmName: data.n || '',
    allowWrite: data.w !== false,
    allowDelete: data.d === true,
    packs: Array.isArray(data.p) ? data.p : DEFAULT_PACKS,
    extraScopes: data.x || '',
    loginHint: data.u || '',
  };
}
