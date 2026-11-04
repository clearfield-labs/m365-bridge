# Security policy

M365 Bridge handles sign-ins to Microsoft 365 for law firms, so we take reports seriously.

## Reporting a vulnerability

Please do not open a public issue. Use GitHub's private reporting: **Security → Report a vulnerability** on this repository.

Include what you found, how to reproduce it, and what an attacker could do. We aim to reply within three working days and to ship a fix for confirmed high-severity issues within a week.

## Design summary

- Delegated Microsoft Graph tokens only. Each person's token reaches only what they can already open. No app-only credentials or client secrets.
- Settings and tokens are encrypted with AES-256-GCM. The key lives in macOS Keychain or Windows DPAPI.
- The local service listens on 127.0.0.1 and ::1 only. It rejects foreign Host headers and cross-site requests. Claude's endpoint needs a per-computer key, and the page needs a per-computer key held in an HttpOnly cookie.
- The sync engine never deletes in SharePoint or OneDrive. Every upload is a new version.
- Admin setup uses Microsoft's own Graph Command Line Tools sign-in; that token stays in memory and is never written to disk.
