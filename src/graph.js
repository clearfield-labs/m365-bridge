// Thin Microsoft Graph client: retries, paging, location parsing, downloads and uploads.
import { getAccessToken } from './auth.js';

export const GRAPH = (process.env.M365_GRAPH_BASE || 'https://graph.microsoft.com/v1.0').replace(/\/$/, '');
const SIMPLE_UPLOAD_MAX = Number(process.env.M365_SIMPLE_UPLOAD_MAX || 4 * 1024 * 1024);
const CHUNK = 320 * 1024 * 16; // 5 MiB, a multiple of 320 KiB as Graph requires

export class GraphError extends Error {
  constructor(status, code, message) {
    super(message);
    this.name = 'GraphError';
    this.status = status;
    this.code = code;
  }
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export async function graph(path, { method = 'GET', body, headers = {}, raw = false, auth = true, token } = {}) {
  const url = path.startsWith('http') ? path : GRAPH + path;
  const isBinary = body instanceof Uint8Array;
  const h = { ...headers };
  if (auth) h.Authorization = `Bearer ${token || (await getAccessToken())}`;
  if (body !== undefined && !isBinary) h['Content-Type'] = 'application/json';
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(url, {
      method,
      headers: h,
      body: body === undefined ? undefined : isBinary ? body : JSON.stringify(body),
    });
    if ((res.status === 429 || res.status === 503 || res.status === 504) && attempt < 4) {
      const wait = Number(res.headers.get('retry-after')) * 1000 || 1000 * 2 ** attempt;
      await sleep(Math.min(wait, 30000));
      continue;
    }
    if (!res.ok) {
      let code = String(res.status);
      let message = res.statusText;
      try {
        const j = await res.json();
        code = j.error?.code || code;
        message = j.error?.message || message;
      } catch { /* not json */ }
      const e = new GraphError(res.status, code, friendly(res.status, code, message));
      e.raw = message;
      throw e;
    }
    if (raw) return res;
    if (res.status === 204) return null;
    const text = await res.text();
    return text ? JSON.parse(text) : null;
  }
}

function friendly(status, code, message) {
  if (status === 401) return 'Microsoft rejected the sign-in token. Sign in again from the dashboard.';
  if (status === 403) return `Access denied. Your Microsoft account does not have permission for this item (${message}).`;
  if (status === 404) return `Not found. The file or folder may have moved, or you may not have access (${message}).`;
  if (status === 412) return 'Conflict: someone changed this file in SharePoint after you last pulled it.';
  if (status === 409) return `Conflict: ${message}`;
  if (status === 423) {
    const who = holderOf(message);
    return `The file is checked out to ${who || 'someone else'} in SharePoint.`;
  }
  return `${code}: ${message}`;
}

export async function graphAll(path, max = 5000) {
  const out = [];
  let next = path;
  while (next && out.length < max) {
    const page = await graph(next);
    out.push(...(page.value || []));
    next = page['@odata.nextLink'] || null;
  }
  return out.slice(0, max);
}

export const enc = (p) => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');

/** SharePoint/OneDrive share or browser URL -> Graph sharing id. */
export function shareId(url) {
  return 'u!' + Buffer.from(url, 'utf8').toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
}

/**
 * A folder URL copied from the SharePoint address bar looks like
 * https://firm.sharepoint.com/sites/X/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FX%2FShared%20Documents%2FSmith
 * Graph cannot resolve the AllItems.aspx form, so rebuild the direct folder URL from the id parameter.
 */
export function normalizeSharePointUrl(url) {
  try {
    const u = new URL(url);
    const id = u.searchParams.get('id') || u.searchParams.get('RootFolder');
    if (id && /\.aspx$/i.test(u.pathname)) return `${u.origin}${encodeURI(id)}`;
  } catch { /* not a URL */ }
  return url;
}

/** SharePoint names the holder of a check-out in its error text, e.g. "i:0#.f|membership|maria@firm.com". */
export function holderOf(message) {
  const m = /([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})/.exec(String(message || ''));
  return m ? m[1] : '';
}

export const refOf = (item) => `${item.parentReference?.driveId || item.driveId}|${item.id}`;

export const SELECT = 'id,name,size,file,folder,package,eTag,cTag,webUrl,lastModifiedDateTime,lastModifiedBy,parentReference,publication';

/**
 * Resolve any way a person might point at a file or folder:
 *  - a SharePoint / OneDrive link copied from the browser or "Copy link"
 *  - a ref returned by another tool ("driveId|itemId")
 *  - a path in the signed-in person's own OneDrive ("/Documents/Clients")
 */
export async function resolveItem(location) {
  const loc = String(location || '').trim();
  if (!loc) throw new Error('No location given.');
  if (/^https?:\/\//i.test(loc)) {
    return graph(`/shares/${shareId(normalizeSharePointUrl(loc))}/driveItem?$select=${SELECT}`);
  }
  if (loc.includes('|')) {
    const [driveId, itemId] = loc.split('|');
    return graph(`/drives/${driveId}/items/${itemId}?$select=${SELECT}`);
  }
  const p = loc.replace(/^(me:|onedrive:)/i, '');
  if (p === '' || p === '/') return graph(`/me/drive/root?$select=${SELECT}`);
  return graph(`/me/drive/root:/${enc(p)}?$select=${SELECT}`);
}

export async function listChildren(driveId, itemId, max = 2000) {
  return graphAll(`/drives/${driveId}/items/${itemId}/children?$top=200&$select=${SELECT}`, max);
}

export async function downloadItem(driveId, itemId) {
  const res = await graph(`/drives/${driveId}/items/${itemId}/content`, { raw: true });
  return Buffer.from(await res.arrayBuffer());
}

/**
 * Upload bytes. Either replace an existing item (itemId + optional ifMatch eTag),
 * or create a new file under parentId with name (fails if the name already exists unless overwrite).
 */
export async function uploadBytes(driveId, buf, { itemId, parentId, name, ifMatch, overwrite = false }) {
  const target = itemId
    ? `/drives/${driveId}/items/${itemId}`
    : `/drives/${driveId}/items/${parentId}:/${encodeURIComponent(name)}:`;
  const headers = ifMatch ? { 'If-Match': ifMatch } : {};
  const behavior = overwrite ? 'replace' : 'fail';
  if (buf.length <= SIMPLE_UPLOAD_MAX) {
    const q = itemId ? '' : `?@microsoft.graph.conflictBehavior=${behavior}`;
    return graph(`${target}/content${q}`, { method: 'PUT', body: new Uint8Array(buf), headers });
  }
  const session = await graph(`${target}/createUploadSession`, {
    method: 'POST',
    headers,
    body: itemId ? {} : { item: { '@microsoft.graph.conflictBehavior': behavior } },
  });
  let result = null;
  for (let start = 0; start < buf.length; start += CHUNK) {
    const end = Math.min(start + CHUNK, buf.length);
    const res = await fetch(session.uploadUrl, {
      method: 'PUT',
      headers: { 'Content-Length': String(end - start), 'Content-Range': `bytes ${start}-${end - 1}/${buf.length}` },
      body: new Uint8Array(buf.subarray(start, end)),
    });
    if (!res.ok && res.status !== 202) {
      throw new GraphError(res.status, String(res.status), friendly(res.status, String(res.status), await res.text()));
    }
    if (res.status === 200 || res.status === 201) result = await res.json();
  }
  return result;
}

export async function createFolder(driveId, parentId, name) {
  try {
    return await graph(`/drives/${driveId}/items/${parentId}/children`, {
      method: 'POST',
      body: { name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' },
    });
  } catch (e) {
    if (e.status === 409) return graph(`/drives/${driveId}/items/${parentId}:/${encodeURIComponent(name)}?$select=${SELECT}`);
    throw e;
  }
}

export const getItem = (driveId, itemId) => graph(`/drives/${driveId}/items/${itemId}?$select=${SELECT}`);

// Check-out is SharePoint's lock. While a file is checked out to you, nobody else can save it,
// and your uploads stay private until check-in publishes them as one new version.
export const checkout = (driveId, itemId) => graph(`/drives/${driveId}/items/${itemId}/checkout`, { method: 'POST' });
export const checkin = (driveId, itemId, comment) =>
  graph(`/drives/${driveId}/items/${itemId}/checkin`, { method: 'POST', body: { comment: String(comment || '').slice(0, 1000) } });
export const discardCheckout = (driveId, itemId) => graph(`/drives/${driveId}/items/${itemId}/discardCheckout`, { method: 'POST' });

/** Move and/or rename an item inside its drive. */
export function moveItem(driveId, itemId, { parentId, name }) {
  const body = {};
  if (parentId) body.parentReference = { id: parentId };
  if (name) body.name = name;
  return graph(`/drives/${driveId}/items/${itemId}`, { method: 'PATCH', body });
}

/** A content fingerprint that survives metadata-only changes (check-out bumps the eTag). */
export const contentTag = (item) => item?.file?.hashes?.quickXorHash || item?.file?.hashes?.sha256Hash || item?.cTag || item?.eTag;

export function personOf(item) {
  const u = item.lastModifiedBy?.user;
  return u?.displayName || u?.email || '';
}

export function summarize(item) {
  return {
    name: item.name,
    type: item.folder ? 'folder' : 'file',
    ref: refOf(item),
    size: item.folder ? undefined : item.size,
    items_inside: item.folder ? item.folder.childCount : undefined,
    modified: item.lastModifiedDateTime,
    modified_by: personOf(item) || undefined,
    url: item.webUrl,
  };
}
