import fs from 'node:fs';
import path from 'node:path';
import { z } from 'zod';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { loadConfig, isConfigured } from './config.js';
import { currentAccount, startDeviceCodeSignIn, deviceCodeState, NotSignedInError, grantedScopes } from './auth.js';
import { hasPack } from './scopes.js';
import { GRAPH, graph, graphAll, resolveItem, listChildren, downloadItem, uploadBytes, createFolder, summarize, refOf, personOf } from './graph.js';
import { extractText } from './extract.js';
import * as ws from './workspace.js';
import { audit } from './audit.js';
import { DASHBOARD_URL } from './paths.js';
import { ensureDashboard, openBrowser } from './dashboard/server.js';

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

export async function startMcp() {
  const server = new McpServer(
    { name: 'm365', version: VERSION },
    {
      instructions:
        'Microsoft 365 files for this firm. SharePoint and OneDrive are the source of truth. ' +
        'For work on a matter folder, use workspace_pull to make a local working copy, edit files there, then workspace_status and workspace_push (like git pull/status/push). ' +
        'For quick lookups use search, list_folder and read_file. Always show the person what will upload before calling workspace_push or upload. ' +
        'If a tool says you are not signed in, call sign_in.',
    },
  );

  tool(server, 'status', 'Show who is signed in to Microsoft 365, the firm setup, and the local workspaces.', {}, async () => {
    const cfg = loadConfig();
    const account = await currentAccount();
    return {
      configured: isConfigured(cfg),
      firm: cfg.firmName || undefined,
      signed_in_as: account ? `${account.name || ''} <${account.username}>`.trim() : null,
      can_write: cfg.allowWrite,
      can_delete: cfg.allowDelete,
      permission_packs: ['files', ...cfg.packs.filter((p) => p !== 'files')],
      token_scopes: account ? await grantedScopes() : undefined,
      workspaces: ws.listWorkspaces().map((w) => ({ name: w.name, local_path: w.localPath, sharepoint: w.webUrl, last_pulled: w.lastPulled })),
      dashboard: DASHBOARD_URL,
      next: !isConfigured(cfg) ? 'Open the dashboard to enter the firm Tenant ID and Client ID.' : !account ? 'Call sign_in.' : undefined,
    };
  });

  tool(
    server,
    'sign_in',
    'Sign in to Microsoft 365. Default opens the local dashboard in the browser. method "device_code" returns a code to enter at microsoft.com/devicelogin instead.',
    { method: z.enum(['browser', 'device_code']).optional() },
    async ({ method = 'browser' }) => {
      const cfg = loadConfig();
      if (!isConfigured(cfg)) {
        await ensureDashboard();
        openBrowser(DASHBOARD_URL);
        return `M365 Bridge is not set up yet. I opened ${DASHBOARD_URL}. Enter the firm Tenant ID and Client ID (or a setup code from your admin), then sign in.`;
      }
      if (method === 'device_code') {
        const state = deviceCodeState();
        if (state?.done && state.account) return `Signed in as ${state.account.username}.`;
        const msg = await startDeviceCodeSignIn();
        return `${msg}\nAfter you finish in the browser, call status to confirm.`;
      }
      await ensureDashboard();
      openBrowser(`${DASHBOARD_URL}auth/start`);
      return `I opened the Microsoft sign-in page in your browser (${DASHBOARD_URL}auth/start). Finish signing in there, then call status to confirm.`;
    },
  );

  tool(server, 'open_dashboard', 'Open the M365 Bridge dashboard (setup, sign-in, workspaces, activity log) in the browser.', {}, async () => {
    await ensureDashboard();
    openBrowser(DASHBOARD_URL);
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
      const me = await graph('/me/drive?$select=id,name,webUrl');
      const sites = search
        ? await graphAll(`/sites?search=${encodeURIComponent(search)}`, 50)
        : await graphAll('/me/followedSites', 50);
      return {
        my_onedrive: { ref: `${me.id}|root`, url: me.webUrl },
        sites: sites.map((s) => ({ name: s.displayName, site: s.id, url: s.webUrl })),
        tip: search ? undefined : 'These are sites you follow. Pass search to find others, or site to list a site\'s libraries.',
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
    'Make or refresh a local working copy of a SharePoint/OneDrive folder (like git clone / git pull). New: pass source (a folder link). Existing: pass local_path. Never overwrites your local edits; if both sides changed, SharePoint\'s version is saved next to yours as "(SharePoint copy ...)".',
    {
      source: z.string().optional().describe('SharePoint/OneDrive folder link or ref. Needed only for a new workspace.'),
      local_path: z.string().optional().describe('Workspace folder. Defaults to ~/M365 Workspaces/<folder name>.'),
      name: z.string().optional().describe('Local folder name for a new workspace.'),
    },
    async ({ source, local_path, name }) => ws.pull({ source, localPath: local_path, name }),
  );

  tool(
    server,
    'workspace_status',
    'Show what changed in a workspace since the last pull: edited here, new here, changed in SharePoint by others (like git status).',
    { local_path: z.string() },
    async ({ local_path }) => ws.status({ localPath: local_path }),
  );

  tool(
    server,
    'workspace_push',
    'Upload edits from a workspace to SharePoint (like git push). Each upload becomes a new SharePoint version. If someone else changed a file since your pull, that file is not uploaded and comes back as a conflict. Never deletes anything in SharePoint.',
    {
      local_path: z.string(),
      files: z.array(z.string()).optional().describe('Only push these files (paths relative to the workspace).'),
      force: z.array(z.string()).optional().describe('Replace these files even if they changed in SharePoint. Their version stays in history. Only with the person\'s explicit OK.'),
    },
    async ({ local_path, files, force }) => ws.push({ localPath: local_path, files, force }),
    { write: true },
  );

  tool(server, 'workspace_list', 'List the local workspaces on this computer and the SharePoint folders they mirror.', {}, async () =>
    ws.listWorkspaces().map((w) => ({ name: w.name, local_path: w.localPath, sharepoint: w.webUrl, last_pulled: w.lastPulled, present: w.exists })),
  );

  const cfg = loadConfig();
  if (hasPack(cfg, 'mail')) registerMail(server);
  if (hasPack(cfg, 'calendar')) registerCalendar(server);
  registerGraphRequest(server);

  await server.connect(new StdioServerTransport());
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
    'Write an email as a DRAFT in Outlook (the person reviews and sends it). Set send_now only when the person has explicitly approved sending this exact email.',
    {
      to: z.array(z.string()),
      subject: z.string(),
      body: z.string().describe('Plain text or HTML.'),
      cc: z.array(z.string()).optional(),
      reply_to_id: z.string().optional().describe('Reply to this message id instead of starting a new thread.'),
      attach: z.array(z.string()).optional().describe('Local file paths to attach (each under 3 MB).'),
      mailbox: z.string().optional().describe('Send from a shared mailbox you have Send As rights to.'),
      send_now: z.boolean().optional(),
    },
    async ({ to, subject, body, cc = [], reply_to_id, attach = [], mailbox: mb, send_now = false }) => {
      const rcpt = (list) => list.map((address) => ({ emailAddress: { address } }));
      const isHtml = /<[a-z][\s\S]*>/i.test(body);
      const message = { subject, body: { contentType: isHtml ? 'HTML' : 'Text', content: body }, toRecipients: rcpt(to), ccRecipients: rcpt(cc) };
      let draft;
      if (reply_to_id) {
        draft = await graph(`${mailbox(mb)}/messages/${encodeURIComponent(reply_to_id)}/createReply`, { method: 'POST', body: {} });
        await graph(`${mailbox(mb)}/messages/${draft.id}`, { method: 'PATCH', body: message });
      } else {
        draft = await graph(`${mailbox(mb)}/messages`, { method: 'POST', body: message });
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
        return `Sent "${subject}" to ${to.join(', ')}.`;
      }
      audit({ tool: 'mail_draft', to, subject, mailbox: mb });
      return { draft_saved: true, subject, open_in_outlook: draft.webLink };
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
      const q = new URLSearchParams({ startDateTime: new Date(start).toISOString(), endDateTime: new Date(end).toISOString(), $top: '100', $orderby: 'start/dateTime', $select: 'subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeeting,webLink,bodyPreview' });
      const events = await graphAll(`${mailbox(calendar_owner)}/calendarView?${q}`, 300);
      return events.map((e) => ({
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
}

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
      const isSearch = method === 'POST' && /^\/search\/query/.test(rel);
      if (method === 'DELETE' && !cfg.allowDelete) throw new Error('Deleting through Microsoft Graph is turned off for this firm. An admin can allow it in the dashboard.');
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
