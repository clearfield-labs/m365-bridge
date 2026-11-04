// A small in-memory Microsoft Graph for tests: one SharePoint library with folders,
// files, eTags, versions, upload sessions, sharing links, search, and a mailbox.
import http from 'node:http';

export function createMockGraph({ token = 'test-token' } = {}) {
  const DRIVE = 'b!mockdrive';
  const HOST = 'https://mock.sharepoint.com/sites/Lit/Shared Documents';
  const items = new Map();
  let seq = 0;
  let base = '';
  const sessions = new Map();

  const now = () => new Date().toISOString();
  function add(parentId, name, content, by = 'Seed') {
    const id = `I${++seq}`;
    const it = { id, name, parentId, ver: 1, modified: now(), by };
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
    eTag: `"{${it.id}},${it.ver}"`,
    cTag: `"c:{${it.id}},${it.ver}"`,
    webUrl: `${HOST}/${pathOf(it)}`.replace(/\/$/, ''),
    lastModifiedDateTime: it.modified,
    lastModifiedBy: { user: { displayName: it.by } },
    parentReference: { driveId: DRIVE, id: it.parentId || undefined },
    ...(it.folder ? { folder: { childCount: children(it.id).length } } : { file: { mimeType: 'application/octet-stream' } }),
    ...(it.id === 'root' ? { root: {} } : {}),
  });
  function update(it, buf, by = 'Test User') {
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
    if (req.headers.authorization !== `Bearer ${token}`) return err(res, 401, 'InvalidAuthenticationToken', 'bad token');
    const m = p.replace(/^\/(v1\.0|beta)/, '');

    if (m === '/me') return send(res, 200, { displayName: 'Test User', mail: 'test@example.com', userPrincipalName: 'test@example.com' });
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
    if (rest === '' && req.method === 'GET') return send(res, 200, json(it));
    if (rest === '/children' && req.method === 'GET') return send(res, 200, { value: children(id).map(json) });
    if (rest === '/children' && req.method === 'POST') {
      const body = JSON.parse(raw);
      if (childByName(id, body.name)) return err(res, 409, 'nameAlreadyExists', 'exists');
      return send(res, 201, json(add(id, body.name, null, 'Test User')));
    }
    if (rest === '/content' && req.method === 'GET') return send(res, 200, it.content);
    if (rest === '/content' && req.method === 'PUT') {
      if (!etagOk) return err(res, 412, 'resourceModified', 'ETag mismatch');
      update(it, raw);
      return send(res, 200, json(it));
    }
    if (rest === '/createUploadSession') {
      if (!etagOk) return err(res, 412, 'resourceModified', 'ETag mismatch');
      const sid = `S${sessions.size + 1}`;
      sessions.set(sid, { buf: Buffer.alloc(0), received: 0, finish: (buf) => { update(it, buf); return it; } });
      return send(res, 200, { uploadUrl: `${base}/upload/${sid}` });
    }
    if (rest === '/versions') return send(res, 200, { value: it.versions.map((v) => ({ id: v.id, lastModifiedDateTime: v.modified, lastModifiedBy: { user: { displayName: v.by } }, size: v.content.length })) });
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
    remoteEdit(relPath, text, by = 'Colleague') {
      update(walk('root', relPath.split('/')), Buffer.from(text), by);
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
