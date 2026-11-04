import { spawn } from 'node:child_process';

/** Open a URL (or a folder) with the computer's default app. */
export function openBrowser(url) {
  if (process.env.M365_NO_BROWSER === '1') return;
  const [cmd, args] =
    process.platform === 'darwin' ? ['open', [url]]
    : process.platform === 'win32' ? (/^https?:/i.test(url) ? ['rundll32', ['url.dll,FileProtocolHandler', url]] : ['explorer', [url]])
    : ['xdg-open', [url]];
  try { spawn(cmd, args, { stdio: 'ignore', detached: true, windowsHide: true }).unref(); } catch { /* headless */ }
}
