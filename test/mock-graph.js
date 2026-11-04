// A small in-memory Microsoft Graph for tests: one SharePoint library with folders,
// files, eTags, versions, check-out/check-in, moves, upload sessions, sharing links,
// search, a mailbox, and the Entra calls that firm setup makes.
import http from 'node:http';
import crypto from 'node:crypto';
import { PACKS } from '../src/scopes.js';

export function createMockGraph({ token = 'test-token', user = { name: 'Test User', email: 'test@example.com' } } = {}) {
  const USERS = {
    [token]: { id: 'U-test', displayName: user.name, userPrincipalName: user.email, mail: user.email },
    'admin-token': { id: 'U-admin', displayName: 'Firm Admin', userPrincipalName: 'admin@mockfirm.com', mail: 'admin@mockfirm.com' },
  };
  const directory = [...Object.values(USERS), { id: 'U-maria', displayName: 'Maria Lopez', userPrincipalName: 'maria@mockfirm.com' }];
  const entra = { apps: [], sps: [], grants: [], assignments: [] };
  const GRAPH_APP = '00000003-0000-0000-c000-000000000000';
  const scopeNames = new Set(['User.Read']);
  for (const p of Object.values(PACKS)) for (const x of [...p.read, ...p.write]) scopeNames.add(x);
  const graphSp = { id: 'SP-graph', appId: GRAPH_APP, oauth2PermissionScopes: [...scopeNames].map((v) => ({ id: `scope-${v}`, value: v })) };
  const DRIVE = 'b!mockdrive';
  const HOST = 'https://mock.sharepoint.com/sites/Lit/Shared Documents';
  const items = new Map();
  let seq = 0;
  let base = '';
  const sessions = new Map();

  const now = () => new Date().toISOString();
  function add(parentId, name, content, by = 'Seed') {
    const id = `I${++seq}`;
    const it = { id, name, parentId, ver: 1, meta: 0, modified: now(), by, checkout: null };
    if (content === null) it.folder = true;
    else { it.content = Buffer.from(content); it.versions = [{ id: '1.0', content: it.content, modified: it.modified, by }]; }
    items.set(id, it);
    return it;
  }
  items.set('root', { id: 'root', name: 'root', folder: true, parentId: null, ver: 1, modified: now() });
  const pathOf = (it) => (it.parentId && it.parentId !== 'root' ? `${pathOf(items.get(it.parentId))}/${it.name}` : it.parentId ? it.name : '');
  const children = (id) => [...items.values()].filter((i) => i.parentId === id);
  const childByName = (id, name) => children(id).find((c) => c.name.toLowerCase() === name.toLowerCase());
  const json = (it) => ({
    id: it.id,
    name: it.name,
    size: it.content ? it.content.length : 0,
    eTag: `"{${it.id}},${it.ver}.${it.meta || 0}"`,
    cTag: `"c:{${it.id}},${it.ver}"`,
    webUrl: `${HOST}/${pathOf(it)}`.replace(/\/$/, ''),
    lastModifiedDateTime: it.modified,
    lastModifiedBy: { user: { displayName: it.by } },
    parentReference: { driveId: DRIVE, id: it.parentId || undefined },
    publication: { level: it.checkout ? 'checkout' : 'published' },
    ...(it.folder
      ? { folder: { childCount: children(it.id).length } }
      : { file: { mimeType: 'application/octet-stream', hashes: { quickXorHash: crypto.createHash('sha1').update(it.content).digest('base64') } } }),
    ...(it.id === 'root' ? { root: {} } : {}),
  });
  function update(it, buf, by = user.name) {
    it.ver++;
    it.content = buf;
    it.modified = now();
    it.by = by;
    it.versions.unshift({ id: `${it.ver}.0`, content: buf, modified: it.modified, by });
  }
  function walk(startId, segs) {
    let cur = items.get(startId);
    for (const s of segs) {
      cur = cur && childByName(cur.id, decodeURIComponent(s));
      if (!cur) return null;
    }
    return cur;
  }

  // Seed: a matter folder.
  const matters = add('root', 'Matters', null);
  const henderson = add(matters.id, 'Henderson v Acme', null);
  add(henderson.id, 'Engagement Letter.txt', 'Engagement letter for Henderson. Fee: hourly.');
  add(henderson.id, 'Timeline.md', '# Timeline\n- 2026-01-04 incident\n');
  const pleadings = add(henderson.id, 'Pleadings', null);
  add(pleadings.id, 'Complaint.txt', 'COMPLAINT draft v1');
  add(henderson.id, '~$Lockfile.txt', 'office lock'); // should be ignored

  const messages = [
    { id: 'M1', subject: 'Henderson settlement offer', from: { emailAddress: { name: 'Opposing Counsel', address: 'oc@lawfirm.com' } }, toRecipients: [], receivedDateTime: now(), bodyPreview: 'We propose...', body: { contentType: 'html', content: '<p>We propose <b>$50,000</b>.</p>' }, webLink: 'https://outlook/M1', hasAttachments: false },
  ];
  const drafts = [];
  const log = [];

  function send(res, status, body, headers = {}) {
    if (Buffer.isBuffer(body)) {
      res.writeHead(status, { 'Content-Type': 'application/octet-stream', ...headers });
      return res.end(body);
    }
    res.writeHead(status, { 'Content-Type': 'application/json', ...headers });
    res.end(body === undefined ? '' : JSON.stringify(body));
  }
  const err = (res, status, code, message) => send(res, status, { error: { code, message } });

  const server = http.createServer(async (req, res) => {
    let mm0;
    const chunks = [];
    for await (const c of req) chunks.push(c);
    const raw = Buffer.concat(chunks);
    const url = new URL(req.url, 'http://x');
    const p = decodeURIComponent(url.pathname);
    log.push(`${req.method} ${p}`);

    if (p.startsWith('/upload/')) {
      const s = sessions.get(p.split('/')[2]);
      const [, start, end, total] = /bytes (\d+)-(\d+)\/(\d+)/.exec(req.headers['content-range']).map(Number);
      raw.copy(s.buf, start);
      s.received += end - start + 1;
      if (s.received < total) return send(res, 202, { nextExpectedRanges: [`${end + 1}-`] });
      const it = s.finish(s.buf);
      return send(res, 201, json(it));
    }
    const caller = USERS[(req.headers.authorization || '').replace(/^Bearer /, '')];
    if (!caller) return err(res, 401, 'InvalidAuthenticationToken', 'bad token');
    const m = p.replace(/^\/(v1\.0|beta)/, '');
    const q = url.searchParams;
    const filterValue = () => /eq '([^']+)'/.exec(q.get('$filter') || '')?.[1];

    if (m === '/me') return send(res, 200, caller);

    // --- Entra (firm setup) ---
    if (m === '/organization') return send(res, 200, { value: [{ id: 'aaaaaaaa-1111-2222-3333-444444444444', displayName: 'Mock Firm LLP' }] });
    if (m === '/servicePrincipals' && req.method === 'GET') {
      const appId = filterValue();
      if (appId === GRAPH_APP) return send(res, 200, { value: [graphSp] });
      return send(res, 200, { value: entra.sps.filter((x) => x.appId === appId) });
    }
    if (m === '/servicePrincipals' && req.method === 'POST') {
      const b = JSON.parse(raw);
      const sp = { id: `SP-${entra.sps.length + 1}`, ...b };
      entra.sps.push(sp);
      return send(res, 201, sp);
    }
    if ((mm0 = /^\/servicePrincipals\/([^/]+)$/.exec(m)) && req.method === 'PATCH') {
      Object.assign(entra.sps.find((x) => x.id === mm0[1]), JSON.parse(raw));
      return send(res, 204);
    }
    if ((mm0 = /^\/servicePrincipals\/([^/]+)\/appRoleAssignedTo$/.exec(m))) {
      const b = JSON.parse(raw);
      if (entra.assignments.some((a) => a.principalId === b.principalId && a.resourceId === b.resourceId)) return err(res, 400, 'Request_BadRequest', 'Permission being assigned already exists on the object');
      entra.assignments.push(b);
      return send(res, 201, b);
    }
    if (m === '/applications' && req.method === 'GET') return send(res, 200, { value: entra.apps.filter((a) => a.displayName === filterValue()) });
    if (m === '/applications' && req.method === 'POST') {
      const app = { id: `APP-${entra.apps.length + 1}`, appId: 'bbbbbbbb-1111-2222-3333-444444444444', ...JSON.parse(raw) };
      entra.apps.push(app);
      return send(res, 201, app);
    }
    if ((mm0 = /^\/applications\/([^/]+)$/.exec(m)) && req.method === 'PATCH') {
      Object.assign(entra.apps.find((a) => a.id === mm0[1]), JSON.parse(raw));
      return send(res, 204);
    }
    if (m === '/oauth2PermissionGrants' && req.method === 'GET') return send(res, 200, { value: entra.grants.filter((g) => g.clientId === filterValue()) });
    if (m === '/oauth2PermissionGrants' && req.method === 'POST') {
      const g = { id: `G-${entra.grants.length + 1}`, ...JSON.parse(raw) };
      entra.grants.push(g);
      return send(res, 201, g);
    }
    if ((mm0 = /^\/oauth2PermissionGrants\/([^/]+)$/.exec(m)) && req.method === 'PATCH') {
      Object.assign(entra.grants.find((g) => g.id === mm0[1]), JSON.parse(raw));
      return send(res, 204);
    }
    if ((mm0 = /^\/users\/([^/]+)$/.exec(m))) {
      const u = directory.find((x) => x.userPrincipalName === decodeURIComponent(mm0[1]).toLowerCase());
      return u ? send(res, 200, u) : err(res, 404, 'Request_ResourceNotFound', 'no user');
    }
    if (m === '/me/drive') return send(res, 200, { id: DRIVE, name: 'Documents', webUrl: HOST, quota: { used: 1e9 } });
    if (m === '/me/followedSites') return send(res, 200, { value: [{ id: 'site1', displayName: 'Litigation', webUrl: 'https://mock.sharepoint.com/sites/Lit' }] });
    if (m === '/me/chats') return send(res, 200, { value: [{ id: 'chat1', topic: 'Henderson team' }] });
    if (m === '/me/messages' && req.method === 'GET') return send(res, 200, { value: messages });
    if (m === '/me/messages' && req.method === 'POST') {
      const d = { id: `D${drafts.length + 1}`, webLink: 'https://outlook/draft', ...JSON.parse(raw) };
      drafts.push(d);
      return send(res, 201, d);
    }
    if (/^\/me\/messages\/[^/]+$/.test(m)) return send(res, 200, messages.find((x) => x.id === m.split('/')[3]) || {});
    if (m === '/search/query') {
      const q = JSON.parse(raw).requests[0].query.queryString.toLowerCase();
      const hits = [...items.values()].filter((i) => !i.folder && (i.name.toLowerCase().includes(q) || i.content?.toString().toLowerCase().includes(q)));
      return send(res, 200, { value: [{ hitsContainers: [{ hits: hits.map((i) => ({ resource: json(i), summary: `…<c0>${q}</c0>…` })) }] }] });
    }
    let mm;
    if ((mm = /^\/shares\/([^/]+)\/driveItem$/.exec(m))) {
      const urlStr = Buffer.from(mm[1].slice(2).replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8');
      const rel = decodeURI(urlStr).replace(`${HOST}/`, '').replace(HOST, '');
      const it = rel ? walk('root', rel.split('/')) : items.get('root');
      return it ? send(res, 200, json(it)) : err(res, 404, 'itemNotFound', 'not found');
    }
    if (m === '/me/drive/root' || m === `/drives/${DRIVE}/root`) return send(res, 200, json(items.get('root')));
    if ((mm = /^\/me\/drive\/root:\/(.+?):?$/.exec(m))) {
      const it = walk('root', mm[1].split('/'));
      return it ? send(res, 200, json(it)) : err(res, 404, 'itemNotFound', 'not found');
    }
    if (!(mm = new RegExp(`^/drives/${DRIVE.replace('!', '\\!')}/items/([^/:]+)(.*)$`).exec(m))) return err(res, 404, 'notFound', `no route ${m}`);
    const [, id, rest] = mm;
    const it = items.get(id);
    if (!it) return err(res, 404, 'itemNotFound', 'item not found');

    // Path addressing: /items/{parent}:/a/b:[/content|/createUploadSession]
    if (rest.startsWith(':/')) {
      const [, relPath, tail = ''] = /^:\/(.+?):?(\/content|\/createUploadSession)?$/.exec(rest);
      const segs = relPath.split('/');
      const name = decodeURIComponent(segs.pop());
      const parent = walk(id, segs);
      const existing = parent && childByName(parent.id, name);
      if (!tail) return existing ? send(res, 200, json(existing)) : err(res, 404, 'itemNotFound', 'not found');
      const behavior = url.searchParams.get('@microsoft.graph.conflictBehavior') || (raw.length && tail === '/createUploadSession' ? JSON.parse(raw).item?.['@microsoft.graph.conflictBehavior'] : 'replace');
      if (existing && behavior === 'fail') return err(res, 409, 'nameAlreadyExists', 'An item with that name already exists.');
      const finish = (buf) => {
        if (existing) { update(existing, buf); return existing; }
        const created = add(parent.id, name, buf, 'Test User');
        return created;
      };
      if (tail === '/content') return send(res, 201, json(finish(raw)));
      const sid = `S${sessions.size + 1}`;
      sessions.set(sid, { buf: Buffer.alloc(0), received: 0, finish });
      const s = sessions.get(sid);
      s.buf = Buffer.alloc(0);
      // Allocate lazily on first chunk using total from content-range.
      s.finish = finish;
      return send(res, 200, { uploadUrl: `${base}/upload/${sid}`, expirationDateTime: now() });
    }
    const ifMatch = req.headers['if-match'];
    const etagOk = !ifMatch || ifMatch === json(it).eTag || ifMatch === json(it).cTag;
    const me = caller.userPrincipalName;
    const lockedOut = () => it.checkout && it.checkout !== me;
    const locked = () => err(res, 423, 'resourceLocked', `The file "${it.name}" is checked out for editing by i:0#.f|membership|${it.checkout}.`);
    if (rest === '' && req.method === 'GET') return send(res, 200, json(it));
    if (rest === '' && req.method === 'PATCH') {
      if (lockedOut()) return locked();
      const b = JSON.parse(raw);
      const parentId = b.parentReference?.id || it.parentId;
      const name = b.name || it.name;
      const clash = childByName(parentId, name);
      if (clash && clash !== it) return err(res, 409, 'nameAlreadyExists', 'An item with that name already exists.');
      it.parentId = parentId;
      it.name = name;
      it.meta = (it.meta || 0) + 1;
      return send(res, 200, json(it));
    }
    if (rest === '/checkout' && req.method === 'POST') {
      if (lockedOut()) return locked();
      it.checkout = me;
      it.meta = (it.meta || 0) + 1; // check-out changes the eTag, not the content
      return send(res, 204);
    }
    if (rest === '/checkin' && req.method === 'POST') {
      if (!it.checkout) return err(res, 400, 'invalidRequest', 'The file is not checked out.');
      if (lockedOut()) return locked();
      it.checkout = null;
      it.meta++;
      if (it.versions?.[0]) it.versions[0].comment = JSON.parse(raw || '{}').comment;
      return send(res, 204);
    }
    if (rest === '/discardCheckout' && req.method === 'POST') {
      if (!it.checkout) return err(res, 400, 'invalidRequest', 'The file is not checked out.');
      if (lockedOut()) return locked();
      it.checkout = null;
      it.meta++;
      return send(res, 204);
    }
    if (rest === '/children' && req.method === 'GET') return send(res, 200, { value: children(id).map(json) });
    if (rest === '/children' && req.method === 'POST') {
      const body = JSON.parse(raw);
      if (childByName(id, body.name)) return err(res, 409, 'nameAlreadyExists', 'exists');
      return send(res, 201, json(add(id, body.name, null, 'Test User')));
    }
    if (rest === '/content' && req.method === 'GET') return send(res, 200, it.content);
    if (rest === '/content' && req.method === 'PUT') {
      if (lockedOut()) return locked();
      if (!etagOk) return err(res, 412, 'resourceModified', 'ETag mismatch');
      update(it, raw);
      return send(res, 200, json(it));
    }
    if (rest === '/createUploadSession') {
      if (lockedOut()) return locked();
      if (!etagOk) return err(res, 412, 'resourceModified', 'ETag mismatch');
      const sid = `S${sessions.size + 1}`;
      sessions.set(sid, { buf: Buffer.alloc(0), received: 0, finish: (buf) => { update(it, buf); return it; } });
      return send(res, 200, { uploadUrl: `${base}/upload/${sid}` });
    }
    if (rest === '/versions') return send(res, 200, { value: it.versions.map((v) => ({ id: v.id, lastModifiedDateTime: v.modified, lastModifiedBy: { user: { displayName: v.by } }, size: v.content.length, comment: v.comment })) });
    if ((mm = /^\/versions\/([^/]+)\/restoreVersion$/.exec(rest))) {
      const v = it.versions.find((x) => x.id === mm[1]);
      if (!v) return err(res, 404, 'itemNotFound', 'no version');
      update(it, v.content, 'Test User');
      return send(res, 204);
    }
    return err(res, 404, 'notFound', `no route ${req.method} ${m}`);
  });

  // Upload session buffers need the total size up front.
  const origEmit = server.emit.bind(server);
  server.emit = (ev, req, res) => {
    if (ev === 'request' && req.url.startsWith('/upload/')) {
      const s = sessions.get(req.url.split('/')[2]);
      const total = Number(/\/(\d+)$/.exec(req.headers['content-range'] || '')?.[1] || 0);
      if (s && s.buf.length !== total) s.buf = Buffer.alloc(total);
    }
    return origEmit(ev, req, res);
  };

  return {
    HOST,
    DRIVE,
    items,
    drafts,
    log,
    find: (relPath) => walk('root', relPath.split('/')),
    entra,
    remoteEdit(relPath, text, by = 'Colleague') {
      update(walk('root', relPath.split('/')), Buffer.from(text), by);
    },
    remoteCheckout(relPath, email) {
      const it = walk('root', relPath.split('/'));
      it.checkout = email;
      it.meta++;
    },
    remoteCheckin(relPath) {
      const it = walk('root', relPath.split('/'));
      it.checkout = null;
      it.meta++;
    },
    listen: () =>
      new Promise((resolve) =>
        server.listen(0, '127.0.0.1', () => {
          base = `http://127.0.0.1:${server.address().port}`;
          resolve(base);
        }),
      ),
    close: () => new Promise((r) => server.close(r)),
  };
}
