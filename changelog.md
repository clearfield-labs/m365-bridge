# Changelog

## 0.9.0 (2026-11-04)

- feat: automatic saving. When Claude finishes a reply (Claude Code Stop hook), edits in synced folders are sent to SharePoint as a new version, checked back in with the note "Saved automatically (via Claude)", and files checked out but not changed are handed back. A short message shows what was saved. Claude is told to push with a meaningful note when it finishes a change, so the generic note is only the fallback
- feat: automatic refresh. When the person sends a message (UserPromptSubmit hook), the synced folder Claude works in gets the latest (8-second limit, then it carries on in the background), and Claude is told which files colleagues changed and which need a merge
- feat: background safety net in the always-on service, every minute: saves edits nobody touched for 3 minutes (made outside Claude, say in Word, or by a script), hands back check-outs unused for 10 minutes (a closed laptop, Claude quitting mid-reply), and refreshes every synced folder every 5 minutes. Errors pause that folder quietly for 5 minutes and show on the page
- feat: when saving finds a file that changed in SharePoint too, Claude is asked once (per reply, only in the folder it works in) to merge it. Nothing is ever overwritten; merges are never finished automatically
- feat: Claude knows the person's day. Once per Claude Code session (and every 3 hours), Claude gets what is left on the calendar today and tomorrow, the unread email count and the latest subjects, marked as information, never instructions. New my_day tool for the Claude chat app and on request
- feat: email is draft first, send on the person's word. mail_draft returns a draft_id, changes an existing draft (draft_id), replies (reply_to_id, reply_all, keeps the quoted thread) and forwards (forward_id). New mail_send sends a draft only after the person says so
- feat: calendar_create (optional Teams link; invitations need the person's OK) and calendar_respond (accept, tentative, decline); calendar times come back in the computer's time zone
- feat: the page has an Automatic section with two switches (saving, the day) for advanced users who want push and pull by hand. Claude's folder instructions follow the switch (v4)
- fix: files and folders whose names start with a dot, Python caches and .pyc files never upload
- perf: file hashes are remembered until a file changes on disk; one operation at a time per synced folder inside the service
- The background service upgrades Claude Code's hooks on start after an update

## 0.8.2 (4 Nov 2026)

- Install page: removed stale steps (Mac password, installing Node.js or Git). The installers need neither.

## 0.8.1 (2026-11-04)

- feat: the install page is live at https://clearfield-labs.github.io/m365-bridge/ (org Pages site). Invitation emails now carry a personal "Install M365 Bridge" button linking to it with the setup code after # (never sent to a server), plus the code itself

## 0.8.0 (2026-11-04)

- feat: one-line installers need nothing installed first. The code downloads as a zip (no Git), and if the computer has no Node.js 18+, a private copy is downloaded and checksum-verified into ~/.m365-bridge-node (no Homebrew, no winget, no admin rights, no pop-ups)
- feat: install page (docs/index.html, for GitHub Pages): detects Mac or Windows, one Copy button, three steps with the exact keys, and shows the person's setup code when the link carries #code=
- feat: `m365-bridge stop`; the Windows installer stops a running copy before replacing its files
- ci: installers run on Windows, macOS and Linux the way a person would (no Git, no Node.js beforehand)

## 0.7.1 (2026-11-04)

- docs: install is one command in Terminal (Mac) or PowerShell (Windows); removed "ask Claude to install it" from the README and the invitation email, since the Claude chat app runs in the cloud and cannot install on the computer

## 0.7.0 (2026-11-04)

Hardening from a QA review and read-only tests on a real law firm tenant.

- security: the page and its API only answer a browser unlocked with this computer's page key (HttpOnly cookie), so other users on a shared computer (Remote Desktop, Citrix, fast user switching) cannot open it. The M365 Bridge app, shortcuts, installer and Claude open it unlocked; bookmarks keep working afterwards
- security: `graph_request` judges the path Graph will receive (no `..` tricks past read-only), and treats `$batch`, `permanentDelete` and `discardCheckout` as deletes
- fix: a key store that is slow or locked (Keychain before unlock, slow DPAPI) no longer leads to a new encryption key and reset settings; undecryptable settings are an error, never "empty"
- fix: the sync record is written atomically and merged per file, so the background service, Claude's tools and the edit hook can work on one folder at once without losing each other's changes
- fix: a check-out is recorded the moment it succeeds; one failing file no longer orphans the others' check-outs
- fix: a refused check-out on a file that is checked out is treated as someone else's, never as yours
- fix: release only discards a check-out that is still in place
- fix: push records the upload before check-in; a failed check-in is reported and the file stays listed as checked out to you (also for new files in libraries that require check-out)
- fix: a cut-off folder listing (very large folders) is never treated as deletions; skipped OneNote and oversized files count as present
- fix: renames that only change letter case no longer delete the local copy; `workspace_move` allows them
- fix: a file open in Word on Windows no longer aborts the whole pull; move falls back cleanly if the local copy is busy
- fix: an open merge with no local edits resolves to SharePoint's version on pull; the edit hook lets Claude write the merged file
- fix: libraries without check-out no longer hide colleagues' changes
- fix: change detection keeps working when SharePoint adds a file's hash a moment after upload
- fix: Claude's config files are never overwritten when they contain a formatting error, and are written atomically
- fix: Windows start-up script is UTF-16 so profile names like José work
- feat: pull downloads new files four at a time
- feat: "M365 Bridge" app on the Mac (Spotlight and Launchpad) and Start menu and desktop shortcuts on Windows; favicon; footer with the bookmark address
- feat: OneNote section files are skipped, not synced
- docs: repo moved to clearfield-labs/m365-bridge; new README, llms.txt, SECURITY.md, CONTRIBUTING.md; credits to Irfad Imtiaz, Clearfield Labs and My Legal Academy

## 0.6.1 (2026-11-04)

- fix: `list_places` no longer fails for an account without OneDrive (for example an unlicensed admin account); it still lists SharePoint sites, falls back to all visible sites when none are followed, and says plainly when no sites are visible
- fix: the folder picker explains a missing OneDrive instead of showing a raw error

## 0.6.0 (2026-11-04)

- feat: "Sync a folder" is now a folder picker: My OneDrive, Shared with me and SharePoint sites (followed first, searchable) on the left, folders with a path bar on the right, and one "Sync <folder>" button. Pasting a link is still there as "Paste a link instead". `/#sync` opens it directly
- feat: `/api/places` and `/api/browse` (same-page only, the person's own sign-in)
- fix: a SharePoint library root synced from the picker is named after the library, not "OneDrive"

## 0.5.2 (2026-11-04)

- feat: the admin who sets up the firm gets access automatically (also checked on every admin sign-in); their row shows "You" with no checkbox
- feat: if the admin sign-in has ended (it lives in memory and ends on restart), Give access, Refresh and Save drafts sign in first and then finish the action
- fix: errors from Admin actions return to the Admin page, not Home

## 0.5.1 (2026-11-04)

- fix: clearer setup-code wording on the Admin overview, and the long code is folded behind "Show the firm's setup code"

## 0.5.0 (2026-11-04)

- feat: admin setup is now a four-step wizard: 1) firm name and permissions, 2) one admin approval for exactly those, 3) choose people from the firm's directory and give them access, 4) email each person their own setup code from the admin's Outlook (drafts first, then Send all). The admin sign-in stays in memory only
- feat: after setup, the Admin page is a firm overview: what Claude can reach, who can use it, the setup code, and a People list (has access / invited) with Refresh. Add someone later by ticking them and clicking Give access and invite. The last loaded list is kept encrypted on this computer
- feat: personal setup codes carry the person's email, so their Microsoft sign-in is pre-filled
- feat: editable invitation template with {first_name} and {setup_code}, and a preview
- fix: paged Graph calls during setup now use the admin sign-in (re-running setup could add a second approval)
- fix: keychain and DPAPI calls time out instead of hanging when the OS key store waits on a dialog
- test: wizard end to end (permissions -> approve -> people -> drafts with personal codes -> send -> overview -> add later)

## 0.4.0 (2026-11-04)

- feat: optional permission packs (off by default): Microsoft 365 groups, create SharePoint libraries and lists, Bookings, meeting attendance, meeting rooms, sensitivity labels, Graph connectors, Copilot meeting recaps, full SharePoint site control. Shown under "Optional" in Permissions
- fix(windows): installer runs npm.cmd (npm.ps1 is blocked by the default execution policy) and stops if npm fails
- fix(windows): hook command uses forward slashes so it launches from cmd, PowerShell and Git Bash
- fix(windows): skip claude.cmd (Node will not spawn .cmd without a shell) and edit Claude's config directly
- fix(windows): retry atomic file saves when antivirus or a reader briefly holds the file
- ci: GitHub Actions on Windows, macOS and Linux running the end-to-end suite plus platform checks (DPAPI/Keychain, start-at-login install, Claude config, hook launch through cmd and Git Bash, uninstall)

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
