# Changelog

## 0.3.3 (2026-11-04)

- fix: saving Permissions before firm setup failed (no tenant IDs yet). It now stores the choices, and setup requests exactly those

## 0.3.2 (2026-11-04)

- feat: while firm setup waits for the admin, a quiet "Open sign-in again" link (the Microsoft sign-in URL) and a Cancel link; `firm_setup` returns the link to Claude too
- feat: the setup form keeps the firm name, access choice and people from the last run

## 0.3.1 (2026-11-04)

- fix: page buttons failed with "Forbidden origin" in real browsers. `Referrer-Policy: no-referrer` made Chromium browsers send `Origin: null` on our own forms. Now `same-origin`, and the cross-site check trusts `Sec-Fetch-Site: same-origin` for opaque origins
- test: own-form POSTs (normal and `Origin: null`) pass; cross-site and opaque-without-metadata are refused

## 0.3.0 (2026-11-04)

- feat: SharePoint check-out as the lock. `workspace_lock` checks files out before editing; colleagues see "Checked out to <you>" and cannot save. `workspace_push` uploads and checks in with a note shown in version history; files locked but unchanged are released (`discardCheckout`). `workspace_release` hands files back
- feat: Claude Code hook (PreToolUse on Edit/Write/MultiEdit/NotebookEdit) checks a file out before Claude edits it, blocks edits to files a colleague holds, and refreshes stale files first. Files outside synced folders pass straight through
- feat: merges instead of copies. When a file changed on both sides, `workspace_merge` gives Claude yours, theirs and base (kept in `.m365/`) and push takes the merged file with `merged`. No "(SharePoint copy)" files any more
- feat: `workspace_move` and `move` rename and organise in SharePoint and locally together; pull removes local copies of files moved or deleted in SharePoint unless they were edited
- feat: content-hash change detection (quickXorHash), so check-out and other metadata changes never look like edits
- feat: background service. Runs at login (macOS LaunchAgent with crash restart, Windows Startup folder, systemd user unit); Claude Code connects over HTTP MCP at `127.0.0.1:4365/mcp` with a per-computer key; the Claude chat app keeps a stdio server
- feat: one-click firm setup. The admin signs in once (Microsoft Graph Command Line Tools app); M365 Bridge creates the Entra app, grants tenant-wide consent for the chosen permissions, and optionally limits access to named people. Also available to Claude as `firm_setup`
- feat: `setup` tool lets Claude apply a setup code; `status` reports the background service, Claude connection and files checked out to you
- feat: new page in a plain Microsoft 365 style (Fluent colours, system font): setup checklist, files checked out to you, synced folders with live status, example prompts, recent activity; admin page with automatic setup, setup code, permissions, manual steps and this computer's state. Bundled fonts removed
- feat: CLI `install` (service + Claude + hook), `start`, `uninstall`, `background`, `hook`
- fix: the service and hook use Homebrew's stable `opt/` node path, so `brew upgrade` does not break them
- test: 22 end-to-end tests, including locks, colleague check-outs, merges, moves, the hook, HTTP MCP auth (no key, wrong key, foreign origin) and automatic firm setup against a mock Entra

## 0.2.0 (2026-11-04)

- feat: redesigned dashboard: library-green rail, paper and brass palette, bundled Instrument Serif + Hanken Grotesk + Plex Mono (no outside requests), light and dark
- feat: overview with a live "bridge" (Microsoft 365, you, Claude), copyable prompts, recent matters and activity
- feat: matters as dossier cards with live SharePoint status (updates from colleagues, edits to upload), Pull latest and Upload edits buttons
- feat: Access view in plain language, Activity as a dated timeline, Admin view with setup code, consent and a guided Entra checklist
- feat: first-run welcome screen; add-matter dialog; toasts
- feat: /api/workspace-status (known workspaces only) and /workspace/push
- security: scripts and styles served as files, CSP drops inline scripts
- fix: demo mode never writes to the real Claude config

## 0.1.1 (2026-11-04)

- feat: `npm run demo` runs the dashboard against a fake Microsoft 365 with a seeded matter, for presenting without a tenant
- refactor: the firm card shows one plain-language line; setup code, consent, scopes and settings move under "Admin tools"
- fix: links readable in dark mode; long paths and codes wrap at narrow widths

## 0.1.0 (2026-11-04)

- feat: local MCP server for Microsoft 365 with per-person delegated sign-in (auth code + PKCE, device code fallback)
- feat: permission packs (files, mail, calendar, contacts, Teams, meetings and transcripts, tasks, OneNote, people) plus extra scopes; v2 one-click admin consent
- feat: matter-folder workspaces: pull, status and push with If-Match conflict protection, merge-base tracking, no deletes, and a per-workspace CLAUDE.md
- feat: tools for search, browse, read (docx/pdf/text), upload, versions, restore, mail search/read/draft, calendar view, and a generic graph_request
- feat: localhost dashboard for firm setup, setup codes, sign-in, Claude registration, workspaces and an activity log
- feat: AES-256-GCM encryption at rest with the key in macOS Keychain or Windows DPAPI
- feat: installers for macOS (install.sh) and Windows (install.ps1); registers with Claude Code and Claude Desktop
- test: end-to-end suite over stdio against a mock Graph (14 tests)
