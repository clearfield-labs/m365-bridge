# Admin setup: Microsoft Entra (once per firm, about 5 minutes)

You need two values from Entra, and neither is a secret:

| Value | Where it is |
|---|---|
| **Directory (tenant) ID** | App registration → Overview |
| **Application (client) ID** | App registration → Overview |

No client secret and no certificate. Each person signs in as themselves, so nothing on a laptop can reach anyone else's mail or files.

## Steps

1. Go to **entra.microsoft.com → Identity → Applications → App registrations → New registration**.
   - Name: `Claude M365 Bridge`
   - Supported account types: **Accounts in this organizational directory only (single tenant)**
   - Redirect URI: platform **Public client/native (mobile & desktop)**, value `http://localhost:4365/auth/callback`
2. **Authentication** → set **Allow public client flows** to **Yes** → Save. This lets people sign in with a code from inside Claude if the browser redirect is blocked.
3. Copy the **tenant ID** and **client ID** from **Overview**.
4. On your own machine, run the installer, open the dashboard (`http://localhost:4365`), and paste the two IDs into **Connect your firm**. Tick the permission packs the firm wants. Save.
5. Click **Grant admin consent for the firm** in the dashboard. You need to be a Global Administrator (or Privileged Role Administrator). That one click approves every ticked permission for the whole tenant through v2 dynamic consent, so you do not have to add each permission in the portal by hand.
   - If you prefer the portal: **API permissions → Add a permission → Microsoft Graph → Delegated**, add the scopes the dashboard lists, then **Grant admin consent for <tenant>**.
6. Copy the **setup code** from the dashboard and send it to your team. It holds the tenant ID, client ID and permission choices. It contains no secret.

## Choose who gets it (recommended)

**Enterprise applications → Claude M365 Bridge → Properties → Assignment required = Yes**, then **Users and groups → Add** the pilot group. Anyone not assigned is refused at sign-in.

Also recommended: a **Conditional Access** policy targeting this app (require MFA, compliant or managed devices, block risky sign-ins).

## Permission packs

Every scope is **delegated**. A delegated token carries the person's identity, so Microsoft lets it reach only what that person can already reach: their own mailbox, mailboxes and calendars that someone shared with them, and the SharePoint sites they belong to. Ethical walls and matter-level permissions in SharePoint still apply.

| Pack | Read-only mode | Read and write mode |
|---|---|---|
| Files (always on) | Files.Read.All, Sites.Read.All | Files.ReadWrite.All, Sites.ReadWrite.All |
| Mail | Mail.Read, Mail.Read.Shared, MailboxSettings.Read | Mail.ReadWrite, Mail.ReadWrite.Shared, Mail.Send, Mail.Send.Shared, MailboxSettings.ReadWrite |
| Calendar | Calendars.Read, Calendars.Read.Shared | Calendars.ReadWrite, Calendars.ReadWrite.Shared |
| Contacts | Contacts.Read, Contacts.Read.Shared | Contacts.ReadWrite, Contacts.ReadWrite.Shared |
| Teams | Chat.Read, ChannelMessage.Read.All, Team.ReadBasic.All, Channel.ReadBasic.All | Chat.ReadWrite, ChannelMessage.Send (+ the read scopes) |
| Meetings | OnlineMeetings.Read, OnlineMeetingTranscript.Read.All, OnlineMeetingRecording.Read.All | OnlineMeetings.ReadWrite (+ transcripts, recordings) |
| Tasks | Tasks.Read, Tasks.Read.Shared | Tasks.ReadWrite, Tasks.ReadWrite.Shared |
| OneNote | Notes.Read.All | Notes.ReadWrite.All |
| People | People.Read, User.ReadBasic.All, Presence.Read.All | same |

**Extra scopes**: type any other delegated Graph scope (for example `Bookings.Read.All`) in the dashboard. It is added to the sign-in request and the consent link.

**Deletes** are off by default, even in read and write mode. Turn on "Allow deletes through Graph" in the dashboard to change that.

## The other ways Entra can do this

| Model | What it is | Per-person isolation | Verdict |
|---|---|---|---|
| **Delegated + tenant-wide admin consent** (this tool) | One app. The admin consents once. Each person signs in and gets their own token. | Yes. Microsoft enforces each person's own permissions. | **Use this.** |
| **Delegated + per-person grants** | The admin grants scopes to specific people only (`POST /oauth2PermissionGrants` with `consentType: "Principal"`). Partners can get Mail.Send while paralegals get read-only, all on one app. | Yes, and the scopes can differ per person. | Good for tiered access. Script it with Graph PowerShell. |
| **Two app registrations** | For example "Bridge – Partners" (full) and "Bridge – Staff" (files only), each assigned to a group. | Yes | The simplest form of tiering. Give each group its own setup code. |
| **App-only (client secret or certificate)** | Microsoft's version of Google domain-wide delegation. One credential reads every mailbox and every site. | No. Whoever holds the secret reaches everyone. | Do not put this on laptops. If needed for a server, limit mail with Exchange **RBAC for Applications** (management scopes) and SharePoint with **Sites.Selected**. |

## Revoking access

- **One person**: Entra → Users → the person → **Revoke sessions**, or remove them from the app assignment.
- **Everyone**: Enterprise applications → Claude M365 Bridge → **Properties → Enabled for users to sign in = No**.
- **One capability**: remove the scope under Enterprise applications → Permissions. For example, removing Mail.Send stops all sending.
- **On a laptop**: the dashboard's **Sign out** button deletes the person's tokens, and `m365-bridge reset` wipes everything.
