(function () {
  'use strict';
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

  // Relative times in the viewer's own time zone.
  function rel(iso) {
    var t = new Date(iso).getTime();
    if (!t) return '';
    var s = Math.round((Date.now() - t) / 1000);
    if (s < 45) return 'just now';
    var m = Math.round(s / 60); if (m < 60) return m + ' min ago';
    var h = Math.round(m / 60); if (h < 24) return h + ' h ago';
    var d = Math.round(h / 24); if (d < 7) return d + ' d ago';
    return new Date(t).toLocaleDateString(undefined, { day: 'numeric', month: 'short' });
  }
  function paintTimes() {
    $$('time[data-rel]').forEach(function (el) {
      el.textContent = rel(el.getAttribute('datetime'));
      el.title = new Date(el.getAttribute('datetime')).toLocaleString();
    });
  }
  paintTimes();
  setInterval(paintTimes, 30000);

  // Copy, dialogs, toggles, menus.
  function copy(text, done) {
    if (navigator.clipboard) navigator.clipboard.writeText(text).then(done, done);
    else { var t = document.createElement('textarea'); t.value = text; document.body.appendChild(t); t.select(); document.execCommand('copy'); t.remove(); done(); }
  }
  document.addEventListener('click', function (ev) {
    var b = ev.target.closest('[data-copy]');
    if (b) {
      copy(b.getAttribute('data-copy'), function () {
        var lbl = b.querySelector('.lbl') || b;
        var was = lbl.textContent;
        lbl.textContent = 'Copied';
        setTimeout(function () { lbl.textContent = was; }, 1400);
      });
      return;
    }
    var o = ev.target.closest('[data-open]');
    if (o) { var d = document.getElementById(o.getAttribute('data-open')); if (d) { d.showModal(); var f = d.querySelector('input:not([type=hidden])'); if (f) f.focus(); } }
    var c = ev.target.closest('[data-close]');
    if (c) c.closest('dialog').close();
    var tg = ev.target.closest('[data-toggle]');
    if (tg) { var el = document.getElementById(tg.getAttribute('data-toggle')); el.hidden = !el.hidden; tg.textContent = el.hidden ? 'Show steps' : 'Hide steps'; }
    $$('details.menu[open]').forEach(function (m) { if (!m.contains(ev.target)) m.open = false; });
  });
  $$('dialog').forEach(function (d) { d.addEventListener('click', function (ev) { if (ev.target === d) d.close(); }); });

  // Busy buttons.
  $$('form[data-busy]').forEach(function (f) {
    f.addEventListener('submit', function (ev) {
      var b = ev.submitter || f.querySelector('button[type=submit]');
      if (!b) return;
      setTimeout(function () { b.disabled = true; b.innerHTML = '<span class="spin"></span>' + f.getAttribute('data-busy'); }, 0);
    });
  });

  // Toasts.
  $$('.toast').forEach(function (t) {
    var close = function () { t.classList.add('out'); setTimeout(function () { t.remove(); }, 400); };
    var x = t.querySelector('button'); if (x) x.addEventListener('click', close);
    setTimeout(close, 8000);
  });

  // Live status for each synced folder.
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function pill(kind, text) { return '<span class="pill ' + kind + '">' + esc(text) + '</span>'; }
  function plural(n, one, many) { return n + ' ' + (n === 1 ? one : many || one + 's'); }
  $$('[data-ws-path]').forEach(function (row) {
    fetch('/api/workspace-status?path=' + encodeURIComponent(row.getAttribute('data-ws-path')), { headers: { Accept: 'application/json' } })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        var box = row.querySelector('.status');
        if (j.error || !j.data) { box.innerHTML = pill('bad', j.error || 'Could not check'); return; }
        var d = j.data;
        var remote = d.changed_in_sharepoint.length + d.new_in_sharepoint.length;
        var local = d.modified.length + d.new_files.length;
        var html = '';
        if (d.needs_merge.length) html += pill('warn', plural(d.needs_merge.length, 'needs a merge', 'need a merge'));
        if (remote) {
          var by = d.changed_in_sharepoint.map(function (c) { return c.by; }).filter(Boolean)[0];
          html += pill('info', plural(remote, 'update') + ' in SharePoint' + (by ? ' from ' + by : ''));
        }
        if (local) html += pill('warn', plural(local, 'change') + ' to send');
        if (d.held_by_others.length) html += pill('', plural(d.held_by_others.length, 'file') + ' checked out to colleagues');
        if (!html) html = pill('ok', 'Up to date');
        box.innerHTML = html;
        var up = row.querySelector('[data-upload]');
        if (up) up.hidden = !local;
        var pull = row.querySelector('[data-pull]');
        if (pull && remote) pull.classList.add('primary');
      })
      .catch(function () { row.querySelector('.status').innerHTML = pill('', 'Status unavailable'); });
  });

  // Firm setup progress.
  var job = document.querySelector('[data-job][data-running]');
  // Folder picker: browse OneDrive, Shared with me and SharePoint sites, then sync the folder you're in.
  var picker = document.getElementById('add');
  if (picker && picker.classList.contains('picker')) {
    var FOLDER = '<svg class="i fold" viewBox="0 0 20 20"><path d="M2.5 6A1.5 1.5 0 0 1 4 4.5h3.6l1.6 1.6h6.8A1.5 1.5 0 0 1 17.5 7.6V14.5A1.5 1.5 0 0 1 16 16H4a1.5 1.5 0 0 1-1.5-1.5z"/></svg>';
    var SITE = '<svg class="i" viewBox="0 0 20 20"><rect x="3" y="3" width="14" height="14" rx="2"/><path d="M7 7h6M7 10h6M7 13h3"/></svg>';
    var listEl = picker.querySelector('[data-list]');
    var crumbsEl = picker.querySelector('[data-crumbs]');
    var sitesEl = picker.querySelector('[data-sites]');
    var syncBtn = picker.querySelector('[data-sync]');
    var srcIn = picker.querySelector('[data-source]');
    var nameIn = picker.querySelector('[data-name]');
    var linkIn = picker.querySelector('[data-link]');
    var dest = picker.querySelector('[data-dest]');
    var root = (document.querySelector('[data-dest-root]') || {}).getAttribute ? document.querySelector('[data-dest-root]').getAttribute('data-dest-root') : '';
    var stack = [];
    var seq = 0;
    var started = false;
    var api = function (u) { return fetch(u, { headers: { Accept: 'application/json' } }).then(function (r) { return r.json(); }); };
    var setSync = function (ref, name) {
      if (linkIn && linkIn.value.trim()) { syncBtn.disabled = false; syncBtn.textContent = 'Sync this link'; dest.textContent = ''; return; }
      srcIn.value = ref || ''; nameIn.value = name || '';
      syncBtn.disabled = !ref;
      syncBtn.textContent = ref ? 'Sync “' + name + '”' : 'Choose a folder';
      dest.textContent = ref && root ? 'Copies to ' + root + '/' + name : '';
    };
    var crumbs = function () {
      crumbsEl.innerHTML = stack.map(function (c, i) { return (i ? '<span class="sep">›</span>' : '') + '<button type="button" data-crumb="' + i + '">' + esc(c.name) + '</button>'; }).join('');
    };
    var open = function () {
      var top = stack[stack.length - 1];
      var mine = ++seq;
      crumbs();
      setSync(null);
      listEl.innerHTML = '<li class="pk-note"><span class="spin"></span></li>';
      api('/api/browse?ref=' + encodeURIComponent(top.ref)).then(function (j) {
        if (mine !== seq) return;
        if (j.error) { listEl.innerHTML = '<li class="pk-note">' + esc(j.error) + '</li>'; return; }
        var d = j.data;
        if (d.syncable) setSync(d.ref, top.name);
        if (!d.folders.length) { listEl.innerHTML = '<li class="pk-note">' + (d.syncable ? 'No folders inside' + (d.files ? ' (' + plural(d.files, 'file') + ' here)' : '') + '. You can sync this one.' : 'No folders here.') + '</li>'; return; }
        listEl.innerHTML = d.folders.map(function (f) {
          var meta = f.library ? 'Library' : f.by ? 'From ' + f.by : f.count != null ? plural(f.count, 'item') : '';
          return '<li><button type="button" data-ref="' + esc(f.ref) + '" data-name="' + esc(f.name) + '">' + (f.library ? SITE : FOLDER) + '<span class="nm">' + esc(f.name) + '</span><span class="meta">' + esc(meta) + '</span><span class="chev">›</span></button></li>';
        }).join('');
      }).catch(function () { listEl.innerHTML = '<li class="pk-note">Could not load. Check that you are signed in.</li>'; });
    };
    var mark = function (el) { $$('.pk-place, .pk-site', picker).forEach(function (b) { if (b === el) b.setAttribute('aria-current', 'true'); else b.removeAttribute('aria-current'); }); };
    var loadSites = function (q) {
      api('/api/places?q=' + encodeURIComponent(q || '')).then(function (j) {
        var sites = (j.data && j.data.sites) || [];
        sitesEl.innerHTML = sites.length ? sites.map(function (s) { return '<button type="button" class="pk-site" data-place="' + esc(s.ref) + '" data-label="' + esc(s.name) + '" title="' + esc(s.url || '') + '">' + SITE + '<span>' + esc(s.name) + '</span></button>'; }).join('') : '<span class="muted small">' + (j.error ? esc(j.error) : 'No sites found') + '</span>';
      });
    };
    picker.addEventListener('click', function (ev) {
      var p = ev.target.closest('[data-place]');
      if (p) { stack = [{ name: p.getAttribute('data-label'), ref: p.getAttribute('data-place') }]; mark(p); open(); return; }
      var f = ev.target.closest('[data-ref]');
      if (f) { stack.push({ name: f.getAttribute('data-name'), ref: f.getAttribute('data-ref') }); open(); return; }
      var c = ev.target.closest('[data-crumb]');
      if (c) { var i = Number(c.getAttribute('data-crumb')); if (i < stack.length - 1) { stack = stack.slice(0, i + 1); open(); } }
    });
    var timer;
    var search = picker.querySelector('[data-site-search]');
    if (search) search.addEventListener('input', function () { clearTimeout(timer); timer = setTimeout(function () { loadSites(search.value); }, 300); });
    if (linkIn) linkIn.addEventListener('input', function () { var top = stack[stack.length - 1]; setSync(srcIn.value, nameIn.value); if (!linkIn.value.trim() && top) open(); });
    document.querySelectorAll('[data-open="add"]').forEach(function (b) {
      b.addEventListener('click', function () {
        if (started) return;
        started = true;
        loadSites('');
        var first = picker.querySelector('[data-place="onedrive"]');
        stack = [{ name: 'My OneDrive', ref: 'onedrive' }]; mark(first); open();
      });
    });
  }

  // /#sync opens the folder picker straight away.
  if (location.hash === '#sync') { var ob = document.querySelector('[data-open="add"]'); if (ob) setTimeout(function () { ob.click(); }, 0); }

  // People picker: search, unlicensed toggle, selected count, confirm before sending.
  var filter = document.querySelector('[data-filter]');
  var plist = document.querySelector('ul.people');
  if (filter && plist) {
    filter.addEventListener('input', function () {
      var q = filter.value.trim().toLowerCase();
      $$('li', plist).forEach(function (li) { li.classList.toggle('hide', !!q && li.getAttribute('data-name').indexOf(q) === -1); });
    });
  }
  var unlic = document.querySelector('[data-show-unlic]');
  if (unlic && plist) unlic.addEventListener('change', function () { plist.classList.toggle('show-unlic', unlic.checked); });
  var count = document.querySelector('[data-count]');
  if (count && plist) {
    var upd = function () { count.textContent = $$('input[name=ids]:checked', plist).length + ' selected'; };
    plist.addEventListener('change', upd);
    upd();
  }
  $$('[data-confirm]').forEach(function (b) {
    b.addEventListener('click', function (ev) { if (!window.confirm(b.getAttribute('data-confirm'))) ev.preventDefault(); });
  });

  if (job) {
    var poll = setInterval(function () {
      fetch('/api/setup-job').then(function (r) { return r.json(); }).then(function (j) {
        if (!j.data || j.data.running) {
          var link = job.querySelector('[data-signin]');
          if (link && j.data && j.data.signInUrl) { link.href = j.data.signInUrl; link.hidden = false; }
          if (j.data) job.querySelector('ul').innerHTML = j.data.steps.map(function (s) { return '<li><svg class="i" viewBox="0 0 20 20"><path d="m4.5 10.5 3.5 3.5 7.5-8"/></svg><span>' + esc(s) + '</span></li>'; }).join('') + '<li class="now"><span class="spin"></span><span>Working…</span></li>';
          return;
        }
        clearInterval(poll);
        location.reload();
      });
    }, 1500);
  }
})();
