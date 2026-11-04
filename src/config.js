import { FILES } from './paths.js';
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
}

export function updateConfig(patch) {
  const cfg = { ...loadConfig(), ...patch };
  saveConfig(cfg);
  return cfg;
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
export function makeSetupCode(cfg) {
  const payload = JSON.stringify({ t: cfg.tenantId, c: cfg.clientId, n: cfg.firmName, w: cfg.allowWrite, d: cfg.allowDelete, p: cfg.packs, x: cfg.extraScopes });
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
  };
}
