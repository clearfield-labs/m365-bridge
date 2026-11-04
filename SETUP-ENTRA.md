# Admin setup: Microsoft Entra (once per firm)

## Automatic (recommended, about 2 minutes)

1. Install M365 Bridge on your own computer (see the README).
2. Open `http://localhost:4365/admin` and go to **Set up your firm**. You can also ask Claude: "Set up M365 Bridge for my firm."
3. Enter the firm name and choose who can use it: **Everyone in the firm**, or **Only these people** with a list of email addresses.
4. Click **Sign in as admin and set up**, and sign in with a Global Administrator account. (Privileged Role Administrator also works.)
   The sign-in uses Microsoft's own "Microsoft Graph Command Line Tools" app, which is the same one `Connect-MgGraph` uses.
   M365 Bridge then:
   - creates the **Claude M365 Bridge** app registration: single tenant, public client, redirect `http://localhost:4365/auth/callback`, no secret
   - adds the delegated Microsoft Graph permissions for the chosen permission packs
   - grants tenant-wide admin consent for exactly those permissions
   - for **Only these people**, turns on **Assignment required** and assigns you plus the listed people
5. Copy the **setup code** it shows and send it to the team.

The admin token stays in memory for these calls and is never saved. To add people later, run setup again with more email addresses. Running it again updates the same app and does not create a second one.

## By hand, in the Entra admin center

You need two values from Entra, and neither is a secret:

| Value | Where it is |
|---|---|
| **Directory (tenant) ID** | App registration → Overview |
| **Application (client) ID** | App registration → Overview |

There is no client secret and no certificate. Each person signs in as themselves, so nothing on a laptop can reach anyone else's mail or files.

1. Go to **entra.microsoft.com → Identity → Applications → App registrations → New registration**.
   - Name: `Claude M365 Bridge`
   - Supported account types: **Accounts in this organizational directory only (single tenant)**
   - Redirect URI: platform **Public client/native (mobile & desktop)**, value `http://localhost:4365/auth/callback`
2. **Authentication** → set **Allow public client flows** to **Yes** → Save. This lets people sign in with a code from inside Claude if the browser redirect is blocked.
3. Copy the **tenant ID** and **client ID** from **Overview**.
4. Open `http://localhost:4365/admin` → **Set it up by hand instead**, paste the two IDs, and save.
5. Open **Permissions** and click **Approve for the firm again**. This needs a Global Administrator. One click approves every permission for the whole tenant through v2 dynamic consent.
   - If you prefer the portal: **API permissions → Add a permission → Microsoft Graph → Delegated**, add the scopes, then **Grant admin consent for <tenant>**.
6. Copy the **setup code** and send it to your team. It holds the tenant ID, client ID and permission choices, and no secret.

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

**Optional packs** (off unless ticked):

| Pack | Read-only mode | Read and write mode |
|---|---|---|
| Microsoft 365 groups | Group.Read.All | Group.ReadWrite.All |
| Create SharePoint libraries and lists | none | Sites.Manage.All |
| Microsoft Bookings | Bookings.Read.All | BookingsAppointment.ReadWrite.All |
| Teams meeting attendance reports | OnlineMeetingArtifact.Read.All | same |
| Meeting rooms | Place.Read.All | same |
| Sensitivity labels | InformationProtectionPolicy.Read | same |
| Graph connectors content | ExternalItem.Read.All | same |
| Copilot meeting recaps (needs Copilot licences) | OnlineMeetingAiInsight.Read.All | same |
| Full control of SharePoint sites | none | Sites.FullControl.All |

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
