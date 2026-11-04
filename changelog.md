# Changelog

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
