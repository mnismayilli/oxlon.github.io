/* site.js — the only runtime script on the site.
   Three jobs, no framework, no network:
     1. the mobile menu (the sidebar collapses below 900px),
     2. the in-page section highlight in the sidebar,
     3. the figure loader — reads assets/figdata/<id>.json when Phase C has
        written it, and leaves a labelled placeholder when it has not.
   Mathematics is pre-rendered at build time; nothing here touches it. */
(function () {
  'use strict';

  /* ---------- 1. mobile menu ------------------------------------------- */
  var btn = document.getElementById('menu-button');
  var side = document.getElementById('sidebar');
  if (btn && side) {
    btn.addEventListener('click', function () {
      var open = side.classList.toggle('open');
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
    side.addEventListener('click', function (e) {
      if (e.target.tagName === 'A' && window.matchMedia('(max-width: 900px)').matches) {
        side.classList.remove('open');
        btn.setAttribute('aria-expanded', 'false');
      }
    });
  }

  /* ---------- 2. section highlight ------------------------------------- */
  var links = Array.prototype.slice.call(
    document.querySelectorAll('.nav-sections a'));
  if (links.length && 'IntersectionObserver' in window) {
    var byId = {};
    links.forEach(function (a) { byId[a.getAttribute('href').slice(1)] = a; });
    var heads = Object.keys(byId)
      .map(function (id) { return document.getElementById(id); })
      .filter(Boolean);
    var obs = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        if (!en.isIntersecting) return;
        links.forEach(function (a) { a.classList.remove('active'); });
        var a = byId[en.target.id];
        if (a) a.classList.add('active');
      });
    }, { rootMargin: '-10% 0px -75% 0px', threshold: 0 });
    heads.forEach(function (h) { obs.observe(h); });
  }

  /* ---------- 3. figures ------------------------------------------------ */
  function draw(mount, spec) {
    if (window.Plotly && spec && spec.data) {
      mount.textContent = '';
      mount.classList.remove('is-loading', 'is-pending');
      mount.classList.add('is-ready');
      window.Plotly.newPlot(mount, spec.data, spec.layout || {},
        Object.assign({ displaylogo: false, responsive: true,
                        modeBarButtonsToRemove: ['lasso2d', 'select2d'] },
                      spec.config || {}));
      return;
    }
    /* data present, renderer not vendored yet — say so rather than sit blank */
    mount.classList.remove('is-loading');
    mount.classList.add('is-pending');
    mount.textContent = 'chart data present · renderer not loaded';
  }

  Array.prototype.forEach.call(
    document.querySelectorAll('.fig-mount'), function (mount) {
      var id = mount.getAttribute('data-fig');
      var inline = document.getElementById('figdata-' + id);
      if (inline) {                       /* build-time inlined by build_site.py */
        try { draw(mount, JSON.parse(inline.textContent)); }
        catch (e) { mount.textContent = 'figure data could not be read (' + id + ')'; }
        return;
      }
      var src = mount.getAttribute('data-src');
      if (!src || !window.fetch) return;  /* leave the pending placebo as built */
      fetch(src).then(function (r) {
        if (!r.ok) throw new Error('missing');
        return r.json();
      }).then(function (spec) {
        draw(mount, spec);
      }).catch(function () {
        /* file:// blocks fetch, or Phase C has not run — keep the placeholder */
      });
    });

  /* ---------- 4. register filter (register.html only) ------------------- */
  var table = document.querySelector('table[data-filterable]');
  var bar = document.querySelector('.filter-bar');
  if (table && bar) {
    var rows = Array.prototype.slice.call(table.tBodies[0].rows);
    var sels = Array.prototype.slice.call(bar.querySelectorAll('select'));
    var count = bar.querySelector('.count');
    var apply = function () {
      var shown = 0;
      rows.forEach(function (tr) {
        var ok = sels.every(function (s) {
          if (!s.value) return true;
          var cell = tr.cells[parseInt(s.getAttribute('data-col'), 10)];
          if (!cell) return false;
          var txt = cell.textContent.trim();
          return s.hasAttribute('data-exact')
            ? txt === s.value
            : txt.indexOf(s.value) !== -1;
        });
        tr.style.display = ok ? '' : 'none';
        if (ok) shown++;
      });
      if (count) count.textContent = shown + ' of ' + rows.length + ' rows';
    };
    sels.forEach(function (s) { s.addEventListener('change', apply); });
    apply();
  }
})();
