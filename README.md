# M365 Bridge for Claude

Use SharePoint, OneDrive, Outlook and Teams from Claude Code, signed in as yourself. Matter folders in SharePoint become the source of truth, the way GitHub is for developers, with no GitHub involved.

- **Every person signs in as themselves.** Claude sees only what that person can already see in Microsoft 365. There are no shared keys and no client secrets on laptops.
- **One admin setup for the whole firm.** The admin creates one Entra app and clicks one consent button, then shares a setup code with the team.
- **Matter folders work like repos.** Claude pulls a SharePoint folder to the laptop, edits it, and pushes changes back as new SharePoint versions. If a colleague changed the same file, the push stops instead of overwriting their work.
- **Encrypted on disk.** Settings and tokens use AES-256-GCM. The key is held by macOS Keychain or Windows DPAPI.
- **Local dashboard** at `http://localhost:4365` for setup, sign-in, matter folders and an activity log.

## Install

**Mac** (Terminal):
```bash
curl -fsSL https://raw.githubusercontent.com/irfad7/m365-bridge/main/install.sh | bash
```

**Windows** (PowerShell):
```powershell
irm https://raw.githubusercontent.com/irfad7/m365-bridge/main/install.ps1 | iex
```

**Or ask Claude Code to install it:**
> Install the M365 Bridge from https://github.com/irfad7/m365-bridge. Clone it, run npm install, then run `node bin/m365-bridge.js install`.

The installer connects the bridge to Claude Code (desktop app and terminal share one config) and to Claude Desktop, then opens the dashboard.

## Setup

1. **Admin, once per firm:** follow [SETUP-ENTRA.md](SETUP-ENTRA.md). You need the **tenant ID** and **client ID**, nothing else.
2. **Everyone:** paste the setup code from your admin into the dashboard, click **Sign in with Microsoft**, then click **Connect both**.
3. Restart Claude.

## What you can say to Claude

- "Search SharePoint for the Henderson engagement letter and summarise the fee terms."
- "Pull the Henderson v Acme folder from SharePoint so we can work on it." (paste the folder link)
- "Update the timeline with the dates from the deposition transcript, then show me what changed before you push."
- "Who changed the complaint last, and what did the previous version say?"
- "Find the settlement email from opposing counsel and draft a reply for me to review."
- "What's on my calendar tomorrow, and pull the transcript from yesterday's Henderson call."

## Tools

| Tool | What it does |
|---|---|
| `status`, `sign_in`, `open_dashboard` | Setup and sign-in from inside Claude (browser or device code) |
| `search` | Search every site and OneDrive you can access |
| `list_places`, `list_folder` | Browse sites, libraries and folders |
| `read_file` | Text from Word, PDF and text files |
| `download`, `upload`, `create_folder` | Single files. Upload never replaces an existing file unless asked to |
| `versions`, `restore_version` | Version history and undo |
| `workspace_pull`, `workspace_status`, `workspace_push`, `workspace_list` | Matter folders as repos |
| `mail_search`, `mail_read`, `mail_draft` | Outlook, including shared mailboxes. `mail_draft` writes drafts and sends only when told to |
| `calendar_view` | Your calendar, or a colleague's calendar you can see |
| `graph_request` | Any other Microsoft Graph call as you: Teams, transcripts, contacts, To Do, Planner, OneNote |

## How matter folders work

```
SharePoint folder  ──pull──▶  ~/M365 Workspaces/Henderson v Acme  ──push──▶  SharePoint (new version)
```

- **pull** downloads new and changed files. It never overwrites a file you edited. If both sides changed, SharePoint's version is saved next to yours as `<name> (SharePoint copy <date>)`.
- **status** lists what you edited, what's new, and what colleagues changed in SharePoint since your pull.
- **push** uploads with an If-Match check on the file's version. If the file changed in SharePoint since your pull, that file is skipped and reported as a conflict. After you merge, the push goes through.
- **Never deletes.** Deleting a local file does not delete it in SharePoint.
- Each workspace gets a `.claude/CLAUDE.md` that tells Claude this workflow, so you can open the folder in Claude Code and start working.

## Commands

```
m365-bridge             open the dashboard
m365-bridge install     connect to Claude + open the dashboard
m365-bridge status      show setup state
m365-bridge reset       delete settings and tokens from this computer
m365-bridge serve       the MCP server (Claude runs this)
```

## Security

- The dashboard listens on 127.0.0.1 only. It rejects foreign Host headers (DNS rebinding) and cross-origin POSTs (CSRF), and it never displays tokens.
- `graph_request` sends the token only to graph.microsoft.com.
- The activity log (`~/.m365-bridge/audit.log`) records actions and file names, never contents.
- Sign-in uses the authorization code flow with PKCE, which is the standard for desktop apps.

## Development

```bash
npm install
npm test        # end-to-end against a mock Microsoft Graph
```
