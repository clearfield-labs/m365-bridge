// Keep M365 Bridge running in the background from login, so Claude can always reach it.
//   macOS    a LaunchAgent (restarts it if it ever crashes)
//   Windows  a hidden start-up entry in the Startup folder
//   Linux    a systemd user service
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync, spawn } from 'node:child_process';
import { CLI_PATH, FILES, HOME, NODE_PATH, DASHBOARD_URL, ensureHome } from './paths.js';
import { pageUrl } from './config.js';

const LABEL = 'com.m365bridge.agent';
const demo = () => process.env.M365_DEMO === '1' || process.env.M365_NO_SERVICE === '1';
let demoInstalled = false;

const plistPath = () => path.join(os.homedir(), 'Library', 'LaunchAgents', `${LABEL}.plist`);
const startupVbs = () => path.join(process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup', 'M365 Bridge.vbs');
const systemdUnit = () => path.join(os.homedir(), '.config', 'systemd', 'user', 'm365-bridge.service');
const uid = () => process.getuid?.() ?? 0;
const run = (cmd, args) => execFileSync(cmd, args, { stdio: ['ignore', 'pipe', 'pipe'] }).toString();
const tryRun = (cmd, args) => { try { return run(cmd, args); } catch { return null; } };

// Settings that should follow the background copy (non-default port, data folder, workspaces).
function passEnv() {
  const keep = ['M365_BRIDGE_PORT', 'M365_BRIDGE_HOME', 'M365_WORKSPACES', 'M365_BRIDGE_KEYSTORE'];
  return Object.fromEntries(keep.filter((k) => process.env[k]).map((k) => [k, process.env[k]]));
}

const xml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export function serviceStatus() {
  if (demo()) return { installed: demoInstalled, running: demoInstalled, kind: 'demo' };
  if (process.platform === 'darwin') {
    const installed = fs.existsSync(plistPath());
    const out = installed ? tryRun('launchctl', ['print', `gui/${uid()}/${LABEL}`]) : null;
    return { installed, running: !!out && /state = running/.test(out), kind: 'LaunchAgent' };
  }
  if (process.platform === 'win32') return { installed: fs.existsSync(startupVbs()), running: null, kind: 'Startup folder' };
  const installed = fs.existsSync(systemdUnit());
  return { installed, running: installed && tryRun('systemctl', ['--user', 'is-active', 'm365-bridge'])?.trim() === 'active', kind: 'systemd' };
}

// A way back to the page that doesn't need the address: an "M365 Bridge" app on the Mac
// (Spotlight and Launchpad find it), Start menu and desktop shortcuts on Windows.
const macApp = () => path.join(os.homedir(), 'Applications', 'M365 Bridge.app');
const winShortcuts = () => [
  path.join(process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'M365 Bridge.url'),
  path.join(os.homedir(), 'Desktop', 'M365 Bridge.url'),
];
const linuxDesktop = () => path.join(os.homedir(), '.local', 'share', 'applications', 'm365-bridge.desktop');

export function installShortcuts() {
  if (demo()) return;
  try {
    if (process.platform === 'darwin') {
      const as = (v) => `"${String(v).replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
      const cmd = `quoted form of ${as(NODE_PATH)} & " " & quoted form of ${as(CLI_PATH)}`;
      const script = `try\n  do shell script ${cmd} & " start >/dev/null 2>&1"\nend try\nset u to ${as(DASHBOARD_URL)}\ntry\n  set u to do shell script ${cmd} & " url"\nend try\nopen location u`;
      fs.mkdirSync(path.dirname(macApp()), { recursive: true });
      fs.rmSync(macApp(), { recursive: true, force: true });
      execFileSync('osacompile', ['-o', macApp(), '-e', script], { stdio: 'ignore', timeout: 30000 });
    } else if (process.platform === 'win32') {
      for (const f of winShortcuts()) {
        fs.mkdirSync(path.dirname(f), { recursive: true });
        // In this person's own profile, which other Windows users cannot read.
        fs.writeFileSync(f, `[InternetShortcut]\r\nURL=${pageUrl(DASHBOARD_URL)}\r\n`);
      }
    } else {
      fs.mkdirSync(path.dirname(linuxDesktop()), { recursive: true });
      fs.writeFileSync(linuxDesktop(), `[Desktop Entry]\nType=Link\nName=M365 Bridge\nURL=${DASHBOARD_URL}\n`);
    }
  } catch { /* a shortcut is a convenience; never fail the install over it */ }
}

export function removeShortcuts() {
  if (demo()) return;
  fs.rmSync(macApp(), { recursive: true, force: true });
  for (const f of winShortcuts()) fs.rmSync(f, { force: true });
  fs.rmSync(linuxDesktop(), { force: true });
}

export function installService() {
  if (demo()) { demoInstalled = true; return 'Demo: background service turned on (nothing was installed).'; }
  ensureHome();
  installShortcuts();
  const node = NODE_PATH;
  if (process.platform === 'darwin') {
    const env = { PATH: '/usr/local/bin:/opt/homebrew/bin:/usr/bin:/bin', ...passEnv() };
    const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>${LABEL}</string>
  <key>ProgramArguments</key>
  <array><string>${xml(node)}</string><string>${xml(CLI_PATH)}</string><string>background</string></array>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><dict><key>SuccessfulExit</key><false/></dict>
  <key>ThrottleInterval</key><integer>10</integer>
  <key>ProcessType</key><string>Background</string>
  <key>EnvironmentVariables</key>
  <dict>${Object.entries(env).map(([k, v]) => `<key>${xml(k)}</key><string>${xml(v)}</string>`).join('')}</dict>
  <key>StandardOutPath</key><string>${xml(FILES.daemonLog)}</string>
  <key>StandardErrorPath</key><string>${xml(FILES.daemonLog)}</string>
</dict>
</plist>
`;
    fs.mkdirSync(path.dirname(plistPath()), { recursive: true });
    fs.writeFileSync(plistPath(), plist);
    tryRun('launchctl', ['bootout', `gui/${uid()}/${LABEL}`]);
    if (tryRun('launchctl', ['bootstrap', `gui/${uid()}`, plistPath()]) === null) {
      tryRun('launchctl', ['unload', plistPath()]);
      run('launchctl', ['load', '-w', plistPath()]);
    }
    return 'Runs in the background and starts when you log in (macOS LaunchAgent).';
  }
  if (process.platform === 'win32') {
    const env = Object.entries(passEnv()).map(([k, v]) => `sh.Environment("Process")("${k}") = "${v.replace(/"/g, '""')}"\r\n`).join('');
    const vbs = `Set sh = CreateObject("WScript.Shell")\r\n${env}sh.Run """${node}"" ""${CLI_PATH}"" background", 0, False\r\n`;
    fs.mkdirSync(path.dirname(startupVbs()), { recursive: true });
    // Windows Script Host reads UTF-16LE with a BOM correctly; plain UTF-8 breaks names like José.
    fs.writeFileSync(startupVbs(), Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(vbs, 'utf16le')]));
    spawn('wscript.exe', [startupVbs()], { detached: true, stdio: 'ignore', windowsHide: true }).unref();
    return 'Runs in the background and starts when you sign in to Windows.';
  }
  const env = Object.entries(passEnv()).map(([k, v]) => `Environment=${k}=${v}\n`).join('');
  fs.mkdirSync(path.dirname(systemdUnit()), { recursive: true });
  fs.writeFileSync(systemdUnit(), `[Unit]\nDescription=M365 Bridge for Claude\n\n[Service]\nExecStart="${node}" "${CLI_PATH}" background\nRestart=on-failure\n${env}\n[Install]\nWantedBy=default.target\n`);
  tryRun('systemctl', ['--user', 'daemon-reload']);
  run('systemctl', ['--user', 'enable', '--now', 'm365-bridge']);
  tryRun('systemctl', ['--user', 'restart', 'm365-bridge']);
  return 'Runs in the background and starts when you log in (systemd user service).';
}

export function uninstallService() {
  if (demo()) { demoInstalled = false; return 'Demo: background service turned off.'; }
  removeShortcuts();
  if (process.platform === 'darwin') {
    tryRun('launchctl', ['bootout', `gui/${uid()}/${LABEL}`]);
    fs.rmSync(plistPath(), { force: true });
  } else if (process.platform === 'win32') {
    fs.rmSync(startupVbs(), { force: true });
  } else {
    tryRun('systemctl', ['--user', 'disable', '--now', 'm365-bridge']);
    fs.rmSync(systemdUnit(), { force: true });
  }
  return 'Background service removed.';
}

/** Restart the background copy so it runs the newest code (after an update). */
export function restartService() {
  if (demo()) return;
  if (process.platform === 'darwin' && fs.existsSync(plistPath())) tryRun('launchctl', ['kickstart', '-k', `gui/${uid()}/${LABEL}`]);
  else if (process.platform === 'linux' && fs.existsSync(systemdUnit())) tryRun('systemctl', ['--user', 'restart', 'm365-bridge']);
}

export { HOME };
