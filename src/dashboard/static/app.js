(function () {
  'use strict';
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  var store = {
    get: function (k) { try { return sessionStorage.getItem(k); } catch (e) { return null; } },
    set: function (k, v) { try { sessionStorage.setItem(k, v); } catch (e) {} },
  };

  // ---- Views (hash routing) ----
  var views = $$('[data-view]');
  function show(name) {
    if (!views.length) return;
    var target = views.filter(function (v) { return v.getAttribute('data-view') === name; })[0] || views[0];
    views.forEach(function (v) { v.hidden = v !== target; });
    var current = target.getAttribute('data-view');
    $$('.nav a').forEach(function (a) {
      if (a.getAttribute('href') === '#' + current) a.setAttribute('aria-current', 'page');
      else a.removeAttribute('aria-current');
    });
    store.set('m365b.view', current);
    if (current === 'matters' || current === 'overview') loadStatuses();
  }
  window.addEventListener('hashchange', function () { show(location.hash.slice(1)); window.scrollTo(0, 0); });

  // ---- Greeting and dates in the viewer's own time zone ----
  var h = new Date().getHours();
  $$('[data-greet]').forEach(function (el) {
    el.textContent = h < 5 ? 'Working late' : h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening';
  });
  $$('[data-today]').forEach(function (el) {
    el.textContent = new Date().toLocaleDateString(undefined, { weekday: 'long', day: 'numeric', month: 'long' });
  });

  function rel(iso) {
    var t = new Date(iso).getTime();
    if (!t) return '';
    var s = Math.round((Date.now() - t) / 1000);
    if (s < 45) return 'just now';
    var m = Math.round(s / 60); if (m < 60) return m + ' min ago';
    var hr = Math.round(m / 60); if (hr < 24) return hr + ' h ago';
    var d = Math.round(hr / 24); if (d < 7) return d + ' d ago';
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
  $$('[data-day]').forEach(function (el) {
    var d = new Date(el.getAttribute('data-day'));
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var diff = Math.round((today - new Date(d.getFullYear(), d.getMonth(), d.getDate())) / 86400000);
    el.textContent = diff === 0 ? 'Today' : diff === 1 ? 'Yesterday' : d.toLocaleDateString(undefined, { weekday: 'long', day: 'numeric', month: 'long' });
  });

  // ---- Copy ----
  function copy(text, done) {
    var ok = function () { if (done) done(); };
    if (navigator.clipboard) navigator.clipboard.writeText(text).then(ok, ok);
    else { var t = document.createElement('textarea'); t.value = text; document.body.appendChild(t); t.select(); document.execCommand('copy'); t.remove(); ok(); }
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
    var p = ev.target.closest('.prompt');
    if (p) {
      copy(p.querySelector('q').textContent, function () {
        p.classList.add('copied');
        var tag = p.querySelector('.tag .lbl');
        tag.textContent = 'Copied. Paste it into Claude';
        setTimeout(function () { p.classList.remove('copied'); tag.textContent = 'Click to copy'; }, 2200);
      });
    }
    var o = ev.target.closest('[data-open]');
    if (o) { var d = document.getElementById(o.getAttribute('data-open')); if (d) { d.showModal(); var f = d.querySelector('input'); if (f) f.focus(); } }
    var c = ev.target.closest('[data-close]');
    if (c) c.closest('dialog').close();
  });
  $$('dialog').forEach(function (d) {
    d.addEventListener('click', function (ev) { if (ev.target === d) d.close(); });
  });

  // ---- Busy buttons ----
  $$('form[data-busy]').forEach(function (f) {
    f.addEventListener('submit', function (ev) {
      var b = ev.submitter || f.querySelector('button[type=submit]');
      if (!b) return;
      setTimeout(function () {
        b.disabled = true;
        b.innerHTML = '<span class="spin"></span>' + f.getAttribute('data-busy');
      }, 0);
    });
  });

  // ---- Toasts ----
  $$('.toast').forEach(function (t, i) {
    var close = function () { t.classList.add('out'); setTimeout(function () { t.remove(); }, 400); };
    var x = t.querySelector('button'); if (x) x.addEventListener('click', close);
    setTimeout(close, 7000 + i * 600);
  });

  // ---- Live matter status ----
  var loaded = false;
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function pill(kind, text) { return '<span class="pill ' + kind + '">' + esc(text) + '</span>'; }
  function plural(n, one, many) { return n + ' ' + (n === 1 ? one : many); }
  function loadStatuses() {
    if (loaded) return;
    loaded = true;
    var cache = {};
    $$('[data-ws-path]').forEach(function (card) {
      var p = card.getAttribute('data-ws-path');
      var req = cache[p] || (cache[p] = fetch('/api/workspace-status?path=' + encodeURIComponent(p), { headers: { Accept: 'application/json' } }).then(function (r) { return r.json(); }));
      req.then(function (j) {
        var box = card.querySelector('.status');
        var list = card.querySelector('.files');
        if (j.error || !j.data) { box.innerHTML = pill('err', j.error || 'Could not check'); return; }
        var d = j.data;
        var remote = d.changed_in_sharepoint.map(function (c) { return { f: c.file, by: c.by }; }).concat(d.new_in_sharepoint.map(function (f) { return { f: f, by: null, isNew: true }; }));
        var local = d.modified.concat(d.new_files);
        var html = '';
        if (d.merging.length) html += pill('err', plural(d.merging.length, 'file to merge', 'files to merge'));
        if (remote.length) {
          var names = remote.map(function (r) { return r.by; }).filter(Boolean);
          html += pill('blue', plural(remote.length, 'update', 'updates') + ' in SharePoint' + (names.length ? ' from ' + names[0] + (names.length > 1 ? ' + others' : '') : ''));
        }
        if (local.length) html += pill('warn', plural(local.length, 'edit', 'edits') + ' to upload');
        if (!html) html = pill('ok', 'In step with SharePoint');
        box.innerHTML = html;
        var count = card.querySelector('[data-count]');
        if (count && d.tracked_files != null) count.textContent = plural(d.tracked_files, 'file', 'files');
        if (list) {
          var items = remote.slice(0, 2).map(function (r) { return '<li title="' + esc(r.f) + '">↓ <span>' + esc(r.f) + '</span></li>'; })
            .concat(local.slice(0, 2).map(function (f) { return '<li title="' + esc(f) + '">↑ <span>' + esc(f) + '</span></li>'; }));
          list.innerHTML = items.join('');
        }
        var up = card.querySelector('[data-upload]');
        if (up) up.hidden = !local.length;
        var pull = card.querySelector('[data-pull]');
        if (pull && remote.length) { pull.classList.remove('btn-quiet'); pull.classList.add('btn-primary'); }
      }).catch(function () {
        card.querySelector('.status').innerHTML = pill('idle', 'Status unavailable');
      });
    });
  }

  show(location.hash.slice(1) || store.get('m365b.view') || 'overview');
})();
