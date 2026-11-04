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
