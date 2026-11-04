// Apply firm settings from a setup code or from the provisioning result.
import { loadConfig, updateConfig, validateIds, parseSetupCode, makeSetupCode, isConfigured } from './config.js';
import { signOut } from './auth.js';
import { scopesFor } from './scopes.js';
import { audit } from './audit.js';

export async function applyFirm(patch) {
  const errors = validateIds(patch.tenantId, patch.clientId);
  if (errors.length) throw new Error(errors.join(' '));
  const before = loadConfig();
  const scopesChanged = scopesFor(before).join(' ') !== scopesFor({ ...before, ...patch }).join(' ');
  if (before.tenantId && (before.tenantId !== patch.tenantId || before.clientId !== patch.clientId || scopesChanged)) await signOut();
  const cfg = updateConfig(patch);
  audit({ tool: 'firm_setup', tenant: patch.tenantId });
  return cfg;
}

export async function applySetupCode(code) {
  let patch;
  try { patch = parseSetupCode(code); } catch { throw new Error('That setup code is not valid. Copy it again from your admin.'); }
  return applyFirm(patch);
}

export function currentSetupCode() {
  const cfg = loadConfig();
  return isConfigured(cfg) ? makeSetupCode(cfg) : null;
}
