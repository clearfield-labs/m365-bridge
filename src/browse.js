// A folder picker over the signed-in person's OneDrive and SharePoint, for the "Sync a folder" dialog.
// Every call uses that person's own sign-in, so it only shows what they can already open.
import { graph, graphAll, listChildren } from './graph.js';

const folderRow = (it, driveId) => ({
  ref: `${driveId || it.parentReference?.driveId}|${it.id}`,
  name: it.name,
  count: it.folder?.childCount ?? null,
  modified: it.lastModifiedDateTime,
});

/** The left-hand list: OneDrive, Shared with me, and SharePoint sites (followed first, or a search). */
export async function places(query = '') {
  const q = String(query).trim();
  const [followed, found] = await Promise.all([
    q ? Promise.resolve([]) : graphAll('/me/followedSites?$select=id,displayName,webUrl', 100).catch(() => []),
    graphAll(`/sites?search=${encodeURIComponent(q || '*')}&$select=id,displayName,webUrl&$top=50`, 50).catch(() => []),
  ]);
  const seen = new Set();
  const sites = [...followed, ...found]
    .filter((s) => s.displayName && !seen.has(s.id) && seen.add(s.id))
    .map((s) => ({ ref: `site:${s.id}`, name: s.displayName, url: s.webUrl, followed: followed.some((f) => f.id === s.id) }));
  return { sites };
}

/** One level of folders inside a place. ref: "onedrive" | "shared" | "site:<id>" | "<driveId>|<itemId>". */
export async function browse(ref) {
  if (ref === 'onedrive') {
    const root = await graph('/me/drive/root?$select=id,name,parentReference,webUrl').catch((e) => {
      if (e.status === 403 || e.status === 404) throw new Error('This account has no OneDrive. That usually means it has no Microsoft 365 licence with OneDrive. Pick a SharePoint site on the left instead.');
      throw e;
    });
    const driveId = root.parentReference?.driveId || (await graph('/me/drive?$select=id')).id;
    const kids = await listChildren(driveId, root.id, 1000);
    // The whole OneDrive is usually far too big to copy; pick a folder inside it.
    return { title: 'My OneDrive', ref: null, syncable: false, folders: kids.filter((k) => k.folder).map((k) => folderRow(k, driveId)) };
  }
  if (ref === 'shared') {
    const items = await graphAll('/me/drive/sharedWithMe?$top=200', 500).catch(() => []);
    const folders = items
      .map((it) => it.remoteItem || it)
      .filter((it) => it.folder)
      .map((it) => ({ ref: `${it.parentReference?.driveId}|${it.id}`, name: it.name, count: it.folder?.childCount ?? null, by: it.shared?.owner?.user?.displayName || it.createdBy?.user?.displayName }));
    return { title: 'Shared with me', ref: null, syncable: false, folders };
  }
  if (ref.startsWith('site:')) {
    const id = ref.slice(5);
    const [site, drives] = await Promise.all([graph(`/sites/${id}?$select=displayName,webUrl`), graphAll(`/sites/${id}/drives?$select=id,name,webUrl`, 100)]);
    if (drives.length === 1) return browse(`${drives[0].id}|root`).then((r) => ({ ...r, title: site.displayName }));
    return { title: site.displayName, ref: null, syncable: false, folders: drives.map((d) => ({ ref: `${d.id}|root`, name: d.name, count: null, library: true })) };
  }
  const [driveId, itemId] = ref.split('|');
  const item = await graph(`/drives/${driveId}/items/${itemId}?$select=id,name,root,folder,parentReference`);
  const kids = await listChildren(driveId, item.id, 1000);
  return {
    title: item.root ? 'Documents' : item.name,
    ref: `${driveId}|${item.id}`,
    syncable: true,
    files: kids.filter((k) => !k.folder).length,
    folders: kids.filter((k) => k.folder).map((k) => folderRow(k, driveId)),
  };
}
