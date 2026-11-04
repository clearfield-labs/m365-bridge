# M365 Bridge for Claude

Use SharePoint, OneDrive, Outlook and Teams from Claude Code, signed in as yourself. Matter folders in SharePoint or OneDrive stay the source of truth, the way GitHub is for developers, with no GitHub involved.

- **Each person signs in as themselves.** Claude sees only what that person can already open in Microsoft 365. There are no shared keys and no secrets on laptops.
- **One admin setup for the firm.** The admin clicks **Set up my firm** and signs in once. M365 Bridge creates the Entra app and approves it, then hands back a setup code for the team.
- **Always on.** M365 Bridge runs in the background from the moment the computer starts. Claude Code connects to it by itself.
- **No two people edit the same file.** Before Claude changes a file, M365 Bridge checks it out in SharePoint, so colleagues see "Checked out to Jeremy". When Claude finishes, the file is checked back in with a note that shows in version history.
- **One file, never two copies.** If a file changed in SharePoint and on the laptop, Claude merges the two into one file. M365 Bridge never leaves "copy" files behind.

## How it works

```
 Claude Code ──► M365 Bridge (background, this computer) ──► Microsoft 365
                 signed in as you                            SharePoint, OneDrive, Outlook, Teams
```

Matter folders work like a git repo:

```
SharePoint folder ──get latest──► ~/M365 Workspaces/Smith v Jones ──send changes──► SharePoint (new version, with a note)
                                    │
                                    └── before Claude edits a file: check it out in SharePoint
```

| Step | What happens |
|---|---|
| **Get latest** | Copies new and changed files down. Never overwrites your edits. Files moved or deleted in SharePoint are removed here too, unless you edited them. |
| **Check out** | Happens automatically before Claude edits. Colleagues can open the file but not save it. If a colleague already holds it, Claude stops and tells you who has it. If SharePoint has a newer version, Claude gets it first. |
| **Send changes** | Uploads the file and checks it in with a note such as "Added deposition dates (via Claude)". Files Claude checked out but did not change are released. |
| **Merge** | Only needed if a file changed on both sides, for example someone edited it in SharePoint before it was checked out. Claude gets your version, their version and the original, writes one merged file, shows you, and sends it. |
| **Organise** | Claude renames and moves files and folders in SharePoint and on your computer together. |

People who don't use Claude keep working in Word and SharePoint as usual. They see a file as checked out while Claude is working on it.

## Install

**Mac** (Terminal):
```bash
curl -fsSL https://raw.githubusercontent.com/irfad7/m365-bridge/main/install.sh | bash
```

**Windows** (PowerShell):
```powershell
irm https://raw.githubusercontent.com/irfad7/m365-bridge/main/install.ps1 | iex
```

**Or ask Claude Code:**
> Install the M365 Bridge from https://github.com/irfad7/m365-bridge. Clone it, run npm install, then run `node bin/m365-bridge.js install`.

The installer:
- sets M365 Bridge to start with the computer: a LaunchAgent on Mac, the Startup folder on Windows
- connects it to Claude Code and adds the check-out-before-edit hook
- connects it to the Claude chat app too, if that is installed
- opens `http://localhost:4365`

Re-running the installer updates M365 Bridge and restarts it.

## Rolling it out to a firm

**The admin, once (about 5 minutes):** open `http://localhost:4365/admin` and follow four steps:
1. **Firm and permissions.** Name the firm and tick what Claude can reach (optional extras included).
2. **Approve.** Sign in once with a Microsoft 365 admin account. On Microsoft's screen, leave "Consent on behalf of your organization" unticked.
3. **Choose people.** Your team appears from the directory. Tick who gets Claude.
4. **Invite.** M365 Bridge saves an email per person in your Outlook, each with their own setup code. Check one, then **Send all**.

Afterwards the Admin page shows your setup and everyone's status. To add someone later, click **Refresh**, tick them, and click **Give access and invite**.

**Each person, once (about 2 minutes):**
1. Install M365 Bridge.
2. Paste the setup code and click **Sign in with Microsoft**. Or ask Claude: "Set up M365 with this code: M365B-…"
3. Restart Claude.

**Every day, people just talk to Claude:**
- "Get the latest on the Smith matter." (The first time, paste the SharePoint folder link.)
- "Update the timeline with the dates from the deposition transcript, show me, then send it back."
- "Organise the Smith folder into pleadings, discovery and correspondence."
- "Find the settlement email from opposing counsel and draft a reply for me to review."

Each person also needs the Claude Code desktop app and a Claude plan.

## Tools Claude gets

| Tool | What it does |
|---|---|
| `status`, `setup`, `firm_setup`, `sign_in`, `open_dashboard` | Setup from inside Claude: paste a code, set up the firm as admin, sign in |
| `search`, `list_places`, `list_folder`, `read_file` | Find and read anything you can access (Word, PDF and text) |
| `download`, `upload`, `create_folder`, `move` | Single files. Upload never replaces an existing file unless told to |
| `versions`, `restore_version` | Version history and undo |
| `workspace_pull`, `workspace_status` | Get latest and see what changed |
| `workspace_lock`, `workspace_release` | Check files out in SharePoint, and hand them back |
| `workspace_push`, `workspace_merge` | Send changes with a note; merge when both sides changed |
| `workspace_move`, `workspace_list` | Organise synced folders; list folders and checked-out files |
| `mail_search`, `mail_read`, `mail_draft` | Outlook, including shared mailboxes. Drafts by default; sends only when told to |
| `calendar_view` | Your calendar, or a colleague's calendar you can see |
| `graph_request` | Any other Microsoft Graph call as you: Teams, transcripts, contacts, To Do, Planner, OneNote |

## Commands

```
m365-bridge             open the M365 Bridge page
m365-bridge install     run in the background from login, connect Claude, open the page
m365-bridge start       start the background service if it stopped
m365-bridge status      show setup state
m365-bridge uninstall   stop the background service and disconnect Claude
m365-bridge reset       delete settings and sign-ins from this computer
```

## Security

- M365 Bridge listens on this computer only (127.0.0.1). It refuses other Host headers (DNS rebinding) and requests from web pages (CSRF and cross-origin).
- Claude Code's connection (`/mcp`) needs a random key that the installer writes into Claude's config. Web pages cannot send it.
- Settings and sign-ins are encrypted with AES-256-GCM. The key is held by macOS Keychain or Windows DPAPI.
- `graph_request` sends the token only to graph.microsoft.com.
- The activity log (`~/.m365-bridge/audit.log`) records actions and file names, never contents.
- Admin setup signs in through Microsoft's own Graph Command Line Tools app. That token stays in memory and is never saved.

## Status

- **Tested:** an end-to-end suite (22 tests) against a fake Microsoft Graph. It covers check-out, check-in with notes, colleague locks, merges, moves, the edit hook, the HTTP connection and its key, and automatic firm setup.
- **Verified on a Mac:** the background LaunchAgent (including restart after a crash) and Claude Code connecting over HTTP (`claude mcp list` shows Connected).
- **Not yet tested on a real Microsoft 365 tenant or on Windows.** Try it on a trial tenant or a pilot group first.

## Demo without a tenant

```bash
npm run demo
```

This runs the real page and tools against a fake Microsoft 365 with a synced matter folder, a file checked out to you, a colleague's update and a colleague's check-out. It opens at `http://localhost:4366` and never touches your real settings.

## Development

```bash
npm install
npm test        # end-to-end against a mock Microsoft Graph
```

See [SETUP-ENTRA.md](SETUP-ENTRA.md) for the admin details and other ways to configure Entra.
