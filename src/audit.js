// Append-only activity log. Records what was touched, never file contents.
import fs from 'node:fs';
import { FILES, ensureHome } from './paths.js';

export function audit(entry) {
  try {
    ensureHome();
    fs.appendFileSync(FILES.audit, JSON.stringify({ ts: new Date().toISOString(), ...entry }) + '\n', { mode: 0o600 });
  } catch { /* logging must never break a tool */ }
}

export function recentAudit(limit = 50) {
  if (!fs.existsSync(FILES.audit)) return [];
  const lines = fs.readFileSync(FILES.audit, 'utf8').trim().split('\n').filter(Boolean);
  return lines.slice(-limit).reverse().map((l) => {
    try { return JSON.parse(l); } catch { return null; }
  }).filter(Boolean);
}
