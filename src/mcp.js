import fs from 'node:fs';
import path from 'node:path';
import { z } from 'zod';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { loadConfig, isConfigured, pageUrl } from './config.js';
import { currentAccount, startDeviceCodeSignIn, deviceCodeState, NotSignedInError, grantedScopes } from './auth.js';
import { hasPack } from './scopes.js';
import { GRAPH, graph, graphAll, resolveItem, listChildren, downloadItem, uploadBytes, createFolder, moveItem, summarize, refOf, personOf } from './graph.js';
import { extractText } from './extract.js';
import * as ws from './workspace.js';
import { audit } from './audit.js';
import { DASHBOARD_URL } from './paths.js';
import { openBrowser } from './browser.js';
import { applySetupCode, applyFirm, currentSetupCode } from './setup.js';
import { startProvision } from './provision.js';
import { serviceStatus } from './service.js';
import { registrationStatus } from './register.js';

// The dashboard lives in the background service. Start it here only if it is not running.
async function ensureDashboard() {
  const { ensureDashboard: ensure } = await import('./dashboard/server.js');
  return ensure();
}

const VERSION = JSON.parse(fs.readFileSync(new URL('../package.json', import.meta.url), 'utf8')).version;

const ok = (data) => ({ content: [{ type: 'text', text: typeof data === 'string' ? data : JSON.stringify(data, null, 2) }] });
const fail = (msg) => ({ content: [{ type: 'text', text: msg }], isError: true });

function tool(server, name, description, shape, handler, { write = false } = {}) {
  server.registerTool(name, { description, inputSchema: shape }, async (args) => {
    try {
      if (write && !loadConfig().allowWrite) return fail('This firm set M365 Bridge to read-only. Ask your admin to allow uploads.');
      return ok(await handler(args || {}));
    } catch (e) {
      audit({ tool: name, error: e.message });
      if (e instanceof NotSignedInError) return fail(`${e.message}\nDashboard: ${DASHBOARD_URL}`);
      return fail(e.message);
    }
  });
}

const location = z.string().describe('A SharePoint/OneDrive link, a "ref" returned by another tool, or a path in your own OneDrive such as "/Documents/Clients".');

const INSTRUCTIONS = `Microsoft 365 for this firm, signed in as the person using Claude. SharePoint and OneDrive are the source of truth.

Setup: if a tool says M365 is not set up, ask the person for the setup code from their admin and call setup. If they are the admin and there is no code yet, call firm_setup. If not signed in, call sign_in.

Synced folders (local copies of SharePoint or OneDrive folders in ~/M365 Workspaces):
- workspace_pull makes a synced folder from a folder link. After that M365 Bridge keeps it current on its own.
- In Claude Code, saving is automatic: edits through Edit/Write check the file out in SharePoint first, and when you finish your reply the edits are sent and checked back in. When you finish a change, call workspace_push with a short note saying what changed (it shows in SharePoint version history); you do not need to ask first, since every version can be restored.
- Scripts that edit a file (for example a .docx through Python) skip the automatic check-out: call workspace_lock first. If a colleague holds the file, do not edit it; say who has it.
- If a file needs a merge, call workspace_merge and combine both versions into the one file. Never leave a second copy.
- workspace_move renames or moves files and folders in SharePoint and locally together. Use it to organise.
- workspace_status and workspace_release are there for when the person asks.

Email: always draft first with mail_draft and show the person the draft (to, subject, the text). Change it with mail_draft and draft_id when they ask. Call mail_send only after the person clearly says to send that email. Never send on your own initiative.

Calendar: my_day gives today and tomorrow at a glance; calendar_view for any range. Before calendar_create with attendees (it sends invitations) or calendar_respond, confirm with the person.

For quick lookups without a local copy use search, list_folder and read_file.`;

export function createServer() {
  const server = new McpServer({ name: 'm365', version: VERSION }, { instructions: INSTRUCTIONS });

  tool(server, 'status', 'Check M365 Bridge: firm setup, who is signed in, whether it runs in the background, files checked out to you, and synced folders. Call this first if anything seems off.', {}, async () => {
    const cfg = loadConfig();
    const account = await currentAccount();
    const configured = isConfigured(cfg);
    const held = ws.heldFiles();
    return {
      configured,
      firm: cfg.firmName || undefined,
      signed_in_as: account ? `${account.name || ''} <${account.username}>`.trim() : null,
      can_write: cfg.allowWrite,
      background_service: serviceStatus(),
      claude: registrationStatus(),
      checked_out_to_you: held.map((h) => ({ folder: h.localPath, file: h.file, since: h.since, edited: h.edited, needs_merge: h.needsMerge || undefined })),
      synced_folders: ws.listWorkspaces().map((w) => ({ name: w.name, local_path: w.localPath, sharepoint: w.webUrl, last_pulled: w.lastPulled })),
      permissions: account ? await grantedScopes() : undefined,
      dashboard: DASHBOARD_URL,
      next: !configured
        ? 'Ask the person for the setup code from their admin and call setup. If they are the Microsoft 365 admin and have no code, call firm_setup.'
        : !account ? 'Call sign_in.' : held.length ? 'Some files are checked out to this person. Push or release them when done.' : undefined,
    };
  });

  tool(
    server,
    'setup',
    'Connect this computer to the firm using the setup code the admin shared (starts with M365B-). Saved encrypted on this computer.',
    { setup_code: z.string() },
    async ({ setup_code }) => {
      const cfg = await applySetupCode(setup_code);
      return `Connected to ${cfg.firmName || 'the firm'}. Next: call sign_in so the person can sign in with their Microsoft account.`;
    },
  );

  tool(
    server,
    'firm_setup',
    'For the firm\'s Microsoft 365 admin, once per firm: creates the Entra app and approves it for the firm. Opens a Microsoft sign-in in the browser; the admin must sign in with an admin account. Returns the setup code to share with the team. Call again with no arguments to check progress.',
    {
      firm_name: z.string().optional(),
      access: z.enum(['everyone', 'selected']).optional().describe('everyone (default): anyone in the firm can sign in. selected: only the people listed (plus the admin).'),
      people: z.array(z.string()).optional().describe('Email addresses to give access when access is "selected".'),
    },
    async ({ firm_name, access = 'everyone', people = [] }) => {
      const { job } = await import('./provision.js');
      let j = job?.running ? job : null;
      if (!j) {
        if (job?.result && !firm_name && !people.length) return finished(job.result);
        j = startProvision({ firmName: firm_name, access, people, cfg: loadConfig() });
      }
      const deadline = Date.now() + 120000;
      while (j.running && Date.now() < deadline) await new Promise((r) => setTimeout(r, 500));
      if (j.running) return { status: 'Waiting for the admin to finish signing in to Microsoft in the browser. Call firm_setup again to check.', sign_in_link: j.signInUrl || undefined, progress: j.steps.map((x) => x.text) };
      if (j.error) throw new Error(`Firm setup failed: ${j.error}`);
      return finished(j.result);
      async function finished(r) {
        await applyFirm({ tenantId: r.tenantId, clientId: r.clientId, firmName: r.firmName });
        return { done: true, setup_code: currentSetupCode(), share: 'Send this setup code to the team. Each person installs M365 Bridge, pastes the code, and signs in.', ...r, next: 'Call sign_in so the admin can sign in as themselves.' };
      }
    },
  );

  tool(
    server,
    'sign_in',
    'Sign in to Microsoft 365. Default opens the sign-in page in the browser. method "device_code" returns a code to enter at microsoft.com/devicelogin instead.',
    { method: z.enum(['browser', 'device_code']).optional() },
    async ({ method = 'browser' }) => {
      const cfg = loadConfig();
      if (!isConfigured(cfg)) {
        return 'M365 Bridge is not connected to a firm yet. Ask the person for the setup code from their admin and call setup (or firm_setup if they are the admin).';
      }
      if (method === 'device_code') {
        const state = deviceCodeState();
        if (state?.done && state.account) return `Signed in as ${state.account.username}.`;
        const msg = await startDeviceCodeSignIn();
        return `${msg}\nAfter you finish in the browser, call status to confirm.`;
      }
      await ensureDashboard();
      openBrowser(pageUrl(DASHBOARD_URL, 'auth/start'));
      return 'I opened the Microsoft sign-in page in the browser. Finish signing in there, then call status to confirm.';
    },
  );

  tool(server, 'open_dashboard', 'Open the M365 Bridge page in the browser (setup, sign-in, files checked out to you, synced folders).', {}, async () => {
    await ensureDashboard();
    openBrowser(pageUrl(DASHBOARD_URL));
    return `Opened ${DASHBOARD_URL}`;
  });

  tool(
    server,
    'search',
    'Search every SharePoint site and OneDrive you have access to, by words in the file name or contents.',
    { query: z.string(), limit: z.number().int().min(1).max(50).optional() },
    async ({ query, limit = 15 }) => {
      const res = await graph('/search/query', {
        method: 'POST',
        body: { requests: [{ entityTypes: ['driveItem'], query: { queryString: query }, from: 0, size: limit }] },
      });
      const hits = res?.value?.[0]?.hitsContainers?.[0]?.hits || [];
      audit({ tool: 'search', query, results: hits.length });
      return hits.map((h) => ({ ...summarize(h.resource), snippet: h.summary?.replace(/<\/?c0>/g, '').replace(/<ddd\/>/g, '…') }));
    },
  );

  tool(
    server,
    'list_places',
    'List the places you can open: your OneDrive, SharePoint sites you follow or that match a search, and the document libraries inside a site.',
    { search: z.string().optional().describe('Find SharePoint sites by name.'), site: z.string().optional().describe('A site URL or site id. Lists its document libraries.') },
    async ({ search, site }) => {
      if (site) {
        const s = /^https?:\/\//.test(site)
          ? await graph(`/sites/${new URL(site).hostname}:${new URL(site).pathname.replace(/\/$/, '')}`)
          : await graph(`/sites/${site}`);
        const drives = await graphAll(`/sites/${s.id}/drives`);
        return { site: s.displayName, url: s.webUrl, libraries: drives.map((d) => ({ name: d.name, ref: `${d.id}|root`, url: d.webUrl })) };
      }
      // An account without a OneDrive licence has no OneDrive; still list the sites.
      const me = await graph('/me/drive?$select=id,name,webUrl').catch(() => null);
      let sites = search
        ? await graphAll(`/sites?search=${encodeURIComponent(search)}`, 50)
        : await graphAll('/me/followedSites', 50).catch(() => []);
      if (!search && !sites.length) sites = await graphAll('/sites?search=*', 50).catch(() => []);
      return {
        my_onedrive: me ? { ref: `${me.id}|root`, url: me.webUrl } : 'This account has no OneDrive (usually no Microsoft 365 licence with OneDrive).',
        sites: sites.map((s) => ({ name: s.displayName, site: s.id, url: s.webUrl })),
        tip: sites.length ? (search ? undefined : 'Pass search to find other sites, or site to list a site\'s libraries.') : 'No SharePoint sites are visible to this account. People only see sites they are members of.',
      };
    },
  );

  tool(server, 'list_folder', 'List the files and folders inside a SharePoint or OneDrive folder.', { location }, async ({ location: loc }) => {
    const folder = await resolveFolderOrRoot(loc);
    const items = await listChildren(folder.parentReference?.driveId || folder.driveId, folder.id, 500);
    return { folder: folder.name, url: folder.webUrl, ref: refOf(folder), items: items.map(summarize) };
  });

  tool(
    server,
    'read_file',
    'Read a Word, PDF or text file from SharePoint/OneDrive and return its text. For spreadsheets and slides, use download and open the local copy.',
    { location, max_chars: z.number().int().min(1000).max(400000).optional() },
    async ({ location: loc, max_chars = 60000 }) => {
      const item = await resolveItem(loc);
      if (item.folder) throw new Error(`"${item.name}" is a folder. Use list_folder.`);
      const buf = await downloadItem(item.parentReference.driveId, item.id);
      const text = await extractText(item.name, buf);
      audit({ tool: 'read_file', target: item.webUrl });
      if (text == null) return { name: item.name, url: item.webUrl, note: 'This file type has no text to extract. Use download to save a local copy.' };
      return {
        name: item.name,
        url: item.webUrl,
        modified: item.lastModifiedDateTime,
        modified_by: personOf(item) || undefined,
        truncated: text.length > max_chars || undefined,
        text: text.slice(0, max_chars),
      };
    },
  );

  tool(
    server,
    'download',
    'Save one SharePoint/OneDrive file to a local path. Use workspace_pull instead for a whole folder you plan to edit.',
    { location, local_path: z.string().describe('Where to save. A folder or a full file path.') },
    async ({ location: loc, local_path }) => {
      const item = await resolveItem(loc);
      if (item.folder) throw new Error(`"${item.name}" is a folder. Use workspace_pull for folders.`);
      let dest = path.resolve(local_path);
      if (fs.existsSync(dest) && fs.statSync(dest).isDirectory()) dest = path.join(dest, item.name);
      fs.mkdirSync(path.dirname(dest), { recursive: true });
      fs.writeFileSync(dest, await downloadItem(item.parentReference.driveId, item.id));
      audit({ tool: 'download', target: item.webUrl, local: dest });
      return { saved: dest, size: item.size, url: item.webUrl };
    },
  );

  tool(
    server,
    'upload',
    'Upload a local file into a SharePoint/OneDrive folder. Refuses to replace an existing file unless replace is true; a replaced file keeps its old versions in SharePoint.',
    {
      local_path: z.string(),
      folder: location.describe('The destination folder: a link, a ref, or an OneDrive path.'),
      name: z.string().optional().describe('File name in SharePoint. Defaults to the local name.'),
      replace: z.boolean().optional(),
    },
    async ({ local_path, folder, name, replace = false }) => {
      const src = path.resolve(local_path);
      if (!fs.existsSync(src)) throw new Error(`No file at ${src}`);
      const dest = await resolveItem(folder);
      if (!dest.folder) throw new Error(`"${dest.name}" is a file, not a folder.`);
      const item = await uploadBytes(dest.parentReference.driveId, fs.readFileSync(src), { parentId: dest.id, name: name || path.basename(src), overwrite: replace }).catch((e) => {
        if (e.status === 409) throw new Error(`A file named "${name || path.basename(src)}" already exists there. Pass replace: true to make a new version of it, or choose another name.`);
        throw e;
      });
      audit({ tool: 'upload', target: item.webUrl, local: src });
      return { uploaded: item.name, url: item.webUrl, ref: refOf(item) };
    },
    { write: true },
  );

  tool(
    server,
    'create_folder',
    'Create a folder inside a SharePoint/OneDrive folder.',
    { parent: location, name: z.string() },
    async ({ parent, name }) => {
      const p = await resolveFolderOrRoot(parent);
      const f = await createFolder(p.parentReference?.driveId || p.driveId, p.id, name);
      audit({ tool: 'create_folder', target: f.webUrl });
      return { created: f.name, url: f.webUrl, ref: refOf(f) };
    },
    { write: true },
  );

  tool(
    server,
    'move',
    'Move and/or rename a file or folder in SharePoint or OneDrive (within the same library). For files inside a synced folder, use workspace_move instead.',
    { location, to_folder: location.optional().describe('Destination folder in the same library.'), new_name: z.string().optional() },
    async ({ location: loc, to_folder, new_name }) => {
      if (!to_folder && !new_name) throw new Error('Give to_folder, new_name, or both.');
      const item = await resolveItem(loc);
      let parentId;
      if (to_folder) {
        const dest = await resolveFolderOrRoot(to_folder);
        if ((dest.parentReference?.driveId || dest.driveId) !== item.parentReference.driveId) throw new Error('That folder is in a different library. Download and upload instead.');
        parentId = dest.id;
      }
      const moved = await moveItem(item.parentReference.driveId, item.id, { parentId, name: new_name });
      audit({ tool: 'move', target: moved.webUrl });
      return { moved: item.name, now: moved.name, url: moved.webUrl, ref: refOf(moved) };
    },
    { write: true },
  );

  tool(
    server,
    'versions',
    'Show the version history of a file: who saved each version and when. Accepts a link, a ref, or a local file inside a workspace.',
    { location },
    async ({ location: loc }) => {
      const item = await resolveItem(fs.existsSync(loc) ? ws.workspaceRefForLocalFile(loc) || loc : loc);
      const versions = await graphAll(`/drives/${item.parentReference.driveId}/items/${item.id}/versions`, 200);
      return {
        name: item.name,
        url: item.webUrl,
        versions: versions.map((v) => ({ version: v.id, saved: v.lastModifiedDateTime, by: v.lastModifiedBy?.user?.displayName, size: v.size })),
      };
    },
  );

  tool(
    server,
    'restore_version',
    'Restore an earlier version of a SharePoint/OneDrive file. The current version stays in history, so this can be undone. Pull the workspace afterwards.',
    { location, version: z.string().describe('A version id from the versions tool, such as "3.0".') },
    async ({ location: loc, version }) => {
      const item = await resolveItem(fs.existsSync(loc) ? ws.workspaceRefForLocalFile(loc) || loc : loc);
      await graph(`/drives/${item.parentReference.driveId}/items/${item.id}/versions/${encodeURIComponent(version)}/restoreVersion`, { method: 'POST' });
      audit({ tool: 'restore_version', target: item.webUrl, version });
      return `Restored ${item.name} to version ${version}. Run workspace_pull to bring it down.`;
    },
    { write: true },
  );

  tool(
    server,
    'workspace_pull',
    'Make or refresh a local copy of a SharePoint or OneDrive folder (like git clone / git pull). New: pass source (a folder link, or an OneDrive path like "/Documents/Clients/Smith"). Existing: pass local_path. Never overwrites your edits; if both sides changed a file, it reports needs_merge.',
    {
      source: z.string().optional().describe('SharePoint/OneDrive folder link, ref, or OneDrive path. Needed only the first time.'),
      local_path: z.string().optional().describe('The local folder. Defaults to ~/M365 Workspaces/<folder name>.'),
      name: z.string().optional().describe('Local folder name for a new copy.'),
    },
    async ({ source, local_path, name }) => ws.pull({ source, localPath: local_path, name }),
  );

  tool(
    server,
    'workspace_status',
    'What changed since the last pull: edited here, new here, checked out to you, needing a merge, changed or checked out by colleagues in SharePoint (like git status).',
    { local_path: z.string() },
    async ({ local_path }) => ws.status({ localPath: local_path }),
  );

  tool(
    server,
    'workspace_lock',
    'Check files out in SharePoint before editing them, so nobody else can save them meanwhile. Reports files a colleague holds (do not edit those), files refreshed with a newer version first (re-read them), and files that need a merge.',
    { local_path: z.string(), files: z.array(z.string()).describe('Paths relative to the folder.') },
    async ({ local_path, files }) => ws.lock({ localPath: local_path, files }),
    { write: true },
  );

  tool(
    server,
    'workspace_push',
    'Send edits to SharePoint now and check them in with a note saying what changed (like git commit + push). Each file becomes one new version; the note shows in version history. Files checked out but unchanged are released. Never deletes anything in SharePoint. Claude Code also does this automatically when you finish a reply, with a generic note, so calling it is how you give the version a meaningful note.',
    {
      local_path: z.string(),
      note: z.string().optional().describe('What changed, in a few words. Shown in SharePoint version history.'),
      files: z.array(z.string()).optional().describe('Only push these files (relative paths).'),
      merged: z.array(z.string()).optional().describe('Files you merged with workspace_merge and the person approved.'),
    },
    async ({ local_path, note, files, merged }) => ws.push({ localPath: local_path, note, files, merged }),
    { write: true },
  );

  tool(
    server,
    'workspace_merge',
    'When a file changed both here and in SharePoint: returns paths to yours (the file in the folder), theirs (SharePoint\'s newer version, kept out of sight) and base (what both started from). Merge into yours, then push with merged.',
    { local_path: z.string(), file: z.string() },
    async ({ local_path, file }) => ws.merge({ localPath: local_path, file }),
    { write: true },
  );

  tool(
    server,
    'workspace_release',
    'Hand back files checked out to you without uploading (others can edit them again). Local edits stay in the folder.',
    { local_path: z.string(), files: z.array(z.string()).optional().describe('Defaults to every file checked out to you in this folder.') },
    async ({ local_path, files }) => ws.release({ localPath: local_path, files }),
  );

  tool(
    server,
    'workspace_move',
    'Rename or move a file or folder inside a synced folder, in SharePoint and locally at once. Creates destination folders as needed. Use this to organise.',
    { local_path: z.string(), from: z.string().describe('Current path relative to the folder.'), to: z.string().describe('New path relative to the folder.') },
    async ({ local_path, from, to }) => ws.move({ localPath: local_path, from, to }),
    { write: true },
  );

  tool(server, 'workspace_list', 'List the synced folders on this computer, and every file checked out to you.', {}, async () => ({
    folders: ws.listWorkspaces().map((w) => ({ name: w.name, local_path: w.localPath, sharepoint: w.webUrl, last_pulled: w.lastPulled, present: w.exists })),
    checked_out_to_you: ws.heldFiles().map((h) => ({ folder: h.localPath, file: h.file, since: h.since, edited: h.edited })),
  }));

  const cfg = loadConfig();
  if (hasPack(cfg, 'mail')) registerMail(server);
  if (hasPack(cfg, 'calendar')) registerCalendar(server);
  registerGraphRequest(server);
  return server;
}

export async function startMcp() {
  await createServer().connect(new StdioServerTransport());
}

/** MCP over HTTP for the background service. Stateless: one server per request. */
export async function handleMcpHttp(req, res, body) {
  const server = createServer();
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined, enableJsonResponse: true });
  res.on('close', () => { transport.close(); server.close(); });
  await server.connect(transport);
  await transport.handleRequest(req, res, body);
}

async function resolveFolderOrRoot(loc) {
  if (/\|root$/.test(loc)) {
    const driveId = loc.split('|')[0];
    return graph(`/drives/${driveId}/root`);
  }
  const item = await resolveItem(loc);
  if (!item.folder && !item.root) throw new Error(`"${item.name}" is a file, not a folder.`);
  return item;
}

const stripHtml = (html) =>
  String(html || '').replace(/<style[\s\S]*?<\/style>/gi, '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/[ \t]+/g, ' ').replace(/\n\s*\n+/g, '\n\n').trim();
const who = (r) => (r?.emailAddress ? `${r.emailAddress.name || ''} <${r.emailAddress.address}>`.trim() : undefined);
const mailbox = (m) => (m ? `/users/${encodeURIComponent(m)}` : '/me');

function registerMail(server) {
  tool(
    server,
    'mail_search',
    'Search Outlook mail. Searches your own mailbox, or a shared mailbox you have rights to (pass mailbox).',
    {
      query: z.string().optional().describe('Words to search for. Supports from:, to:, subject:, hasattachments:true.'),
      folder: z.string().optional().describe('inbox, sentitems, drafts, archive, or a folder id.'),
      mailbox: z.string().optional().describe('Shared mailbox address, e.g. intake@firm.com.'),
      limit: z.number().int().min(1).max(50).optional(),
    },
    async ({ query, folder, mailbox: mb, limit = 15 }) => {
      const base = `${mailbox(mb)}${folder ? `/mailFolders/${encodeURIComponent(folder)}` : ''}/messages`;
      const select = '$select=id,subject,from,toRecipients,receivedDateTime,hasAttachments,bodyPreview,webLink,conversationId';
      const url = query
        ? `${base}?$search=${encodeURIComponent(`"${query.replace(/"/g, '')}"`)}&$top=${limit}&${select}`
        : `${base}?$orderby=receivedDateTime desc&$top=${limit}&${select}`;
      const res = await graph(url);
      audit({ tool: 'mail_search', query, mailbox: mb, results: res.value.length });
      return res.value.map((m) => ({
        id: m.id,
        subject: m.subject,
        from: who(m.from),
        to: (m.toRecipients || []).map(who),
        received: m.receivedDateTime,
        attachments: m.hasAttachments || undefined,
        preview: m.bodyPreview,
        url: m.webLink,
      }));
    },
  );

  tool(
    server,
    'mail_read',
    'Read one email in full, with its attachment list. Pass save_attachments_to to save the attachments locally.',
    { id: z.string(), mailbox: z.string().optional(), save_attachments_to: z.string().optional() },
    async ({ id, mailbox: mb, save_attachments_to }) => {
      const base = `${mailbox(mb)}/messages/${encodeURIComponent(id)}`;
      const m = await graph(`${base}?$select=subject,from,toRecipients,ccRecipients,receivedDateTime,body,webLink,hasAttachments`, {
        headers: { Prefer: 'outlook.body-content-type="text"' },
      });
      let attachments = [];
      if (m.hasAttachments) {
        const list = await graph(`${base}/attachments`);
        for (const a of list.value) {
          const entry = { name: a.name, size: a.size, type: a.contentType };
          if (save_attachments_to && a.contentBytes) {
            const dir = path.resolve(save_attachments_to);
            fs.mkdirSync(dir, { recursive: true });
            entry.saved = path.join(dir, path.basename(a.name));
            fs.writeFileSync(entry.saved, Buffer.from(a.contentBytes, 'base64'));
          }
          attachments.push(entry);
        }
      }
      audit({ tool: 'mail_read', mailbox: mb, subject: m.subject, saved_attachments: attachments.filter((a) => a.saved).length });
      return {
        subject: m.subject,
        from: who(m.from),
        to: (m.toRecipients || []).map(who),
        cc: (m.ccRecipients || []).map(who),
        received: m.receivedDateTime,
        url: m.webLink,
        body: m.body?.contentType === 'html' ? stripHtml(m.body.content) : m.body?.content,
        attachments,
      };
    },
  );

  tool(
    server,
    'mail_draft',
    'Write an email as a DRAFT in the person\'s Outlook Drafts folder, then show them the draft. To change it, call again with draft_id. To send it once they say so, call mail_send with the draft id. Also for replies (reply_to_id, reply_all) and forwards (forward_id).',
    {
      to: z.array(z.string()).optional().describe('Recipients. For a reply, leave empty to reply to the sender (or everyone with reply_all).'),
      subject: z.string().optional().describe('Leave empty on a reply or forward to keep "RE:"/"FW:" and the original subject.'),
      body: z.string().describe('Plain text or HTML. On a reply or forward, the original message is kept below it.'),
      cc: z.array(z.string()).optional(),
      bcc: z.array(z.string()).optional(),
      reply_to_id: z.string().optional().describe('Reply to this message id (keeps the thread).'),
      reply_all: z.boolean().optional().describe('With reply_to_id: reply to everyone on the thread.'),
      forward_id: z.string().optional().describe('Forward this message id.'),
      draft_id: z.string().optional().describe('Change this existing draft instead of making a new one.'),
      attach: z.array(z.string()).optional().describe('Local file paths to attach (each under 3 MB).'),
      mailbox: z.string().optional().describe('Draft in (and later send from) a shared mailbox you have Send As rights to.'),
      send_now: z.boolean().optional().describe('Only when the person has already approved sending this exact email. Prefer mail_send after showing the draft.'),
    },
    async ({ to = [], subject, body, cc = [], bcc = [], reply_to_id, reply_all = false, forward_id, draft_id, attach = [], mailbox: mb, send_now = false }) => {
      const rcpt = (list) => list.map((address) => ({ emailAddress: { address } }));
      const isHtml = /<[a-z][\s\S]*>/i.test(body);
      const html = isHtml ? body : `<div>${body.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>')}</div>`;
      const patch = {
        ...(subject ? { subject } : {}),
        ...(to.length ? { toRecipients: rcpt(to) } : {}),
        ...(cc.length ? { ccRecipients: rcpt(cc) } : {}),
        ...(bcc.length ? { bccRecipients: rcpt(bcc) } : {}),
      };
      let draft;
      if (draft_id) {
        draft = await graph(`${mailbox(mb)}/messages/${encodeURIComponent(draft_id)}?$select=id,isDraft,webLink,subject,toRecipients`);
        if (draft.isDraft === false) throw new Error('That email was already sent. Make a new draft instead.');
        await graph(`${mailbox(mb)}/messages/${draft.id}`, { method: 'PATCH', body: { ...patch, body: { contentType: 'HTML', content: html } } });
      } else if (reply_to_id || forward_id) {
        // Outlook builds the reply or forward with the original message quoted; our text goes on top.
        const id = encodeURIComponent(reply_to_id || forward_id);
        const action = forward_id ? 'createForward' : reply_all ? 'createReplyAll' : 'createReply';
        draft = await graph(`${mailbox(mb)}/messages/${id}/${action}`, { method: 'POST', body: {} });
        const full = await graph(`${mailbox(mb)}/messages/${draft.id}?$select=body`).catch(() => null);
        const quoted = full?.body?.content || '';
        await graph(`${mailbox(mb)}/messages/${draft.id}`, { method: 'PATCH', body: { ...patch, body: { contentType: 'HTML', content: `${html}${quoted}` } } });
      } else {
        if (!to.length || !subject) throw new Error('A new email needs to and subject.');
        draft = await graph(`${mailbox(mb)}/messages`, { method: 'POST', body: { ...patch, body: { contentType: isHtml ? 'HTML' : 'Text', content: body } } });
      }
      for (const file of attach) {
        const buf = fs.readFileSync(path.resolve(file));
        await graph(`${mailbox(mb)}/messages/${draft.id}/attachments`, {
          method: 'POST',
          body: { '@odata.type': '#microsoft.graph.fileAttachment', name: path.basename(file), contentBytes: buf.toString('base64') },
        });
      }
      if (send_now) {
        await graph(`${mailbox(mb)}/messages/${draft.id}/send`, { method: 'POST' });
        audit({ tool: 'mail_send', to, subject, mailbox: mb });
        return `Sent "${subject || draft.subject || 'the email'}"${to.length ? ` to ${to.join(', ')}` : ''}.`;
      }
      audit({ tool: 'mail_draft', to, subject: subject || draft.subject, mailbox: mb });
      return {
        draft_saved: true,
        draft_id: draft.id,
        subject: subject || draft.subject,
        open_in_outlook: draft.webLink,
        next: 'Show the person the draft. Send it with mail_send and this draft_id only when they say so.',
      };
    },
    { write: true },
  );

  tool(
    server,
    'mail_send',
    'Send a draft that the person has seen and told you to send. Pass the draft_id from mail_draft. Never call this without the person clearly saying to send.',
    { draft_id: z.string(), mailbox: z.string().optional() },
    async ({ draft_id, mailbox: mb }) => {
      const d = await graph(`${mailbox(mb)}/messages/${encodeURIComponent(draft_id)}?$select=id,subject,isDraft,toRecipients,ccRecipients`);
      if (d.isDraft === false) return `"${d.subject}" was already sent.`;
      const to = [...(d.toRecipients || []), ...(d.ccRecipients || [])].map((r) => r.emailAddress?.address).filter(Boolean);
      if (!to.length) throw new Error('This draft has no recipients yet. Add them with mail_draft and draft_id.');
      await graph(`${mailbox(mb)}/messages/${encodeURIComponent(draft_id)}/send`, { method: 'POST' });
      audit({ tool: 'mail_send', to, subject: d.subject, mailbox: mb });
      return `Sent "${d.subject}" to ${to.join(', ')}. It is in Sent Items.`;
    },
    { write: true },
  );
}

function registerCalendar(server) {
  tool(
    server,
    'calendar_view',
    'List calendar events between two dates, including Teams meeting links. Pass calendar_owner for a colleague\'s or a shared calendar you can see.',
    {
      start: z.string().describe('ISO date or date-time, e.g. 2026-11-04'),
      end: z.string().describe('ISO date or date-time'),
      calendar_owner: z.string().optional(),
    },
    async ({ start, end, calendar_owner }) => {
      const q = new URLSearchParams({ startDateTime: new Date(start).toISOString(), endDateTime: new Date(end).toISOString(), $top: '100', $orderby: 'start/dateTime', $select: 'id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeeting,webLink,bodyPreview,responseStatus' });
      const events = await graphAll(`${mailbox(calendar_owner)}/calendarView?${q}`, 300, { headers: { Prefer: `outlook.timezone="${localTz()}"` } });
      return events.map((e) => ({
        id: e.id,
        your_response: e.responseStatus?.response,
        subject: e.subject,
        start: e.start?.dateTime,
        end: e.end?.dateTime,
        timezone: e.start?.timeZone,
        location: e.location?.displayName || undefined,
        organizer: who(e.organizer),
        attendees: (e.attendees || []).map((a) => who(a)),
        teams_link: e.onlineMeeting?.joinUrl,
        preview: e.bodyPreview || undefined,
        url: e.webLink,
      }));
    },
  );

  tool(server, 'my_day', 'The person\'s day at a glance: what is left on their calendar today, tomorrow\'s meetings, and unread email. Use it when they ask about their day or what to work on.', {}, async () => {
    const { myDay } = await import('./day.js');
    return myDay();
  });

  tool(
    server,
    'calendar_create',
    'Add an event to the person\'s calendar. With attendees, Outlook sends them invitations as soon as it is created, so confirm the details with the person first. teams: true adds a Teams meeting link.',
    {
      subject: z.string(),
      start: z.string().describe('Local date-time, e.g. 2026-11-05T14:00'),
      end: z.string().describe('Local date-time'),
      attendees: z.array(z.string()).optional().describe('Email addresses. They get an invitation.'),
      location: z.string().optional(),
      body: z.string().optional(),
      teams: z.boolean().optional(),
      all_day: z.boolean().optional(),
    },
    async ({ subject, start, end, attendees = [], location, body, teams = false, all_day = false }) => {
      const tz = localTz();
      const ev = await graph('/me/events', {
        method: 'POST',
        headers: { Prefer: `outlook.timezone="${tz}"` },
        body: {
          subject,
          start: { dateTime: start, timeZone: tz },
          end: { dateTime: end, timeZone: tz },
          isAllDay: all_day,
          ...(location ? { location: { displayName: location } } : {}),
          ...(body ? { body: { contentType: 'Text', content: body } } : {}),
          attendees: attendees.map((address) => ({ emailAddress: { address }, type: 'required' })),
          ...(teams ? { isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness' } : {}),
        },
      });
      audit({ tool: 'calendar_create', subject, attendees: attendees.length });
      return { created: true, id: ev.id, subject, start: ev.start?.dateTime || start, end: ev.end?.dateTime || end, invitations_sent_to: attendees, teams_link: ev.onlineMeeting?.joinUrl, url: ev.webLink };
    },
    { write: true },
  );

  tool(
    server,
    'calendar_respond',
    'Accept, tentatively accept or decline a meeting invitation (event id from calendar_view). The organizer is told unless send_response is false. Confirm with the person first.',
    {
      event_id: z.string(),
      response: z.enum(['accept', 'tentative', 'decline']),
      comment: z.string().optional(),
      send_response: z.boolean().optional(),
    },
    async ({ event_id, response, comment, send_response = true }) => {
      const action = { accept: 'accept', tentative: 'tentativelyAccept', decline: 'decline' }[response];
      await graph(`/me/events/${encodeURIComponent(event_id)}/${action}`, { method: 'POST', body: { comment: comment || '', sendResponse: send_response } });
      audit({ tool: 'calendar_respond', response });
      return `Done: ${response === 'tentative' ? 'tentatively accepted' : `${response}ed`}.${send_response ? ' The organizer was told.' : ''}`;
    },
    { write: true },
  );
}

const localTz = () => Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';

function registerGraphRequest(server) {
  tool(
    server,
    'graph_request',
    'Call any Microsoft Graph API endpoint as the signed-in person (their own token, their own permissions). Use for anything the other tools do not cover: Teams chats and channels, meeting transcripts, contacts, To Do, Planner, OneNote, people, presence, mail rules. Confirm with the person before any POST, PATCH, PUT or DELETE that changes data or sends something.',
    {
      method: z.enum(['GET', 'POST', 'PATCH', 'PUT', 'DELETE']).optional(),
      path: z.string().describe('Graph path such as /me/chats?$top=20 or /me/onlineMeetings/{id}/transcripts'),
      body: z.any().optional().describe('JSON body for POST/PATCH/PUT.'),
      version: z.enum(['v1.0', 'beta']).optional(),
      save_to: z.string().optional().describe('Save a binary response (file content, transcript, photo) to this local path.'),
    },
    async ({ method = 'GET', path: p, body, version = 'v1.0', save_to }) => {
      const cfg = loadConfig();
      const graphOrigin = new URL(GRAPH).origin;
      let rel = p.trim();
      if (/^https?:\/\//i.test(rel)) {
        if (new URL(rel).origin !== graphOrigin) throw new Error('graph_request only talks to Microsoft Graph.');
        rel = rel.slice(graphOrigin.length).replace(/^\/(v1\.0|beta)/, '');
      }
      if (!rel.startsWith('/')) rel = `/${rel}`;
      // Judge the path Graph will actually receive ("/search/query/../../me/sendMail" is sendMail).
      const resolved = new URL(`${graphOrigin}/${version}${rel}`);
      if (resolved.origin !== graphOrigin) throw new Error('graph_request only talks to Microsoft Graph.');
      rel = resolved.pathname.replace(/^\/(v1\.0|beta)/, '') + resolved.search;
      const pathOnly = resolved.pathname.replace(/^\/(v1\.0|beta)/, '').toLowerCase();
      const isSearch = method === 'POST' && pathOnly === '/search/query';
      const deleteLike = method === 'DELETE' || /\/(permanentdelete|discardcheckout)$/.test(pathOnly) || pathOnly === '/$batch';
      if (deleteLike && !cfg.allowDelete) throw new Error('Deleting through Microsoft Graph is turned off for this firm (this includes batch requests). An admin can allow it.');
      if (method !== 'GET' && !isSearch && !cfg.allowWrite) throw new Error('This firm set M365 Bridge to read-only.');
      const res = await graph(`${graphOrigin}/${version}${rel}`, { method, body, raw: true });
      audit({ tool: 'graph_request', method, path: rel.split('?')[0] });
      if (res.status === 204) return { status: 204, ok: true };
      const type = res.headers.get('content-type') || '';
      if (type.includes('json')) {
        const text = await res.text();
        return text.length > 100000 ? `${text.slice(0, 100000)}\n…truncated. Narrow with $select, $top or $filter.` : text;
      }
      const buf = Buffer.from(await res.arrayBuffer());
      if (save_to) {
        const dest = path.resolve(save_to);
        fs.mkdirSync(path.dirname(dest), { recursive: true });
        fs.writeFileSync(dest, buf);
        return { saved: dest, bytes: buf.length, content_type: type };
      }
      if (/^text\/|vtt|xml/.test(type)) return buf.toString('utf8').slice(0, 100000);
      return { bytes: buf.length, content_type: type, note: 'Binary response. Pass save_to to save it.' };
    },
  );
}
