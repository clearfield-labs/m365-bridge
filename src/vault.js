// Encryption at rest. A random 256-bit master key lives in the OS key store
// (macOS Keychain, Windows DPAPI). Files on disk hold only AES-256-GCM ciphertext.
import crypto from 'node:crypto';
import fs from 'node:fs';
import { execFileSync } from 'node:child_process';
import { FILES, HOME, writePrivate } from './paths.js';

const SERVICE = 'm365-bridge';
// Separate key per data folder, so a test folder never touches the real key.
const ACCOUNT = `master-key-${crypto.createHash('sha256').update(HOME).digest('hex').slice(0, 10)}`;

let cachedKey = null;
let storageLabel = null;

function backend() {
  const forced = process.env.M365_BRIDGE_KEYSTORE;
  if (forced) return forced;
  if (process.platform === 'darwin') return 'keychain';
  if (process.platform === 'win32') return 'dpapi';
  return 'file';
}

// macOS Keychain via the `security` tool. The key goes through stdin, never argv.
function keychainGet() {
  try {
    const out = execFileSync('security', ['find-generic-password', '-s', SERVICE, '-a', ACCOUNT, '-w'], {
      stdio: ['ignore', 'pipe', 'ignore'],
    });
    const hex = out.toString().trim();
    return /^[0-9a-f]{64}$/.test(hex) ? Buffer.from(hex, 'hex') : null;
  } catch {
    return null;
  }
}
function keychainSet(key) {
  execFileSync('security', ['-i'], {
    input: `add-generic-password -U -s ${SERVICE} -a ${ACCOUNT} -l "M365 Bridge" -w ${key.toString('hex')}\n`,
    stdio: ['pipe', 'ignore', 'pipe'],
  });
  if (!keychainGet()) throw new Error('Keychain write did not stick');
}
export function keychainDelete() {
  try {
    execFileSync('security', ['delete-generic-password', '-s', SERVICE, '-a', ACCOUNT], { stdio: 'ignore' });
  } catch { /* not present */ }
}

// Windows DPAPI (CurrentUser scope) via Windows PowerShell 5.1, which ships with every Windows.
function dpapi(op, b64) {
  const script =
    'Add-Type -AssemblyName System.Security; ' +
    '$i=[Console]::In.ReadToEnd().Trim(); ' +
    `$o=[Security.Cryptography.ProtectedData]::${op}([Convert]::FromBase64String($i),$null,[Security.Cryptography.DataProtectionScope]::CurrentUser); ` +
    '[Console]::Out.Write([Convert]::ToBase64String($o))';
  return execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], {
    input: b64,
    stdio: ['pipe', 'pipe', 'pipe'],
    windowsHide: true,
  }).toString().trim();
}
function dpapiGet() {
  if (!fs.existsSync(FILES.dpapiKey)) return null;
  try {
    return Buffer.from(dpapi('Unprotect', fs.readFileSync(FILES.dpapiKey, 'utf8')), 'base64');
  } catch {
    return null;
  }
}
function dpapiSet(key) {
  writePrivate(FILES.dpapiKey, dpapi('Protect', key.toString('base64')));
}

function fileGet() {
  if (!fs.existsSync(FILES.keyFile)) return null;
  const hex = fs.readFileSync(FILES.keyFile, 'utf8').trim();
  return /^[0-9a-f]{64}$/.test(hex) ? Buffer.from(hex, 'hex') : null;
}
function fileSet(key) {
  writePrivate(FILES.keyFile, key.toString('hex'));
}

const LABELS = { keychain: 'macOS Keychain', dpapi: 'Windows DPAPI', file: 'Local key file (owner-only)' };

export function getMasterKey() {
  if (cachedKey) return cachedKey;
  const b = backend();
  const stores = {
    keychain: [keychainGet, keychainSet],
    dpapi: [dpapiGet, dpapiSet],
    file: [fileGet, fileSet],
  };
  const [get, set] = stores[b] || stores.file;
  let key = null;
  try {
    key = get();
  } catch { /* fall through */ }
  if (key) {
    storageLabel = LABELS[b];
  } else if (b !== 'file' && fileGet()) {
    // An earlier run fell back to the key file. Keep using it so old data stays readable.
    key = fileGet();
    storageLabel = `${LABELS.file} (fallback)`;
  } else {
    key = crypto.randomBytes(32);
    try {
      set(key);
      storageLabel = LABELS[b];
    } catch {
      fileSet(key);
      storageLabel = `${LABELS.file} (fallback)`;
    }
  }
  cachedKey = key;
  return key;
}

export function keyStorage() {
  getMasterKey();
  return storageLabel;
}

export function encrypt(text) {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv('aes-256-gcm', getMasterKey(), iv);
  const data = Buffer.concat([cipher.update(text, 'utf8'), cipher.final()]);
  return JSON.stringify({
    v: 1,
    alg: 'aes-256-gcm',
    iv: iv.toString('base64'),
    tag: cipher.getAuthTag().toString('base64'),
    data: data.toString('base64'),
  });
}

export function decrypt(blob) {
  const { iv, tag, data } = JSON.parse(blob);
  const decipher = crypto.createDecipheriv('aes-256-gcm', getMasterKey(), Buffer.from(iv, 'base64'));
  decipher.setAuthTag(Buffer.from(tag, 'base64'));
  return Buffer.concat([decipher.update(Buffer.from(data, 'base64')), decipher.final()]).toString('utf8');
}

/** Read and decrypt a file. Returns null if missing or unreadable (e.g. the key was reset). */
export function readEncrypted(file) {
  if (!fs.existsSync(file)) return null;
  try {
    return decrypt(fs.readFileSync(file, 'utf8'));
  } catch {
    return null;
  }
}

export function writeEncrypted(file, text) {
  writePrivate(file, encrypt(text));
}
