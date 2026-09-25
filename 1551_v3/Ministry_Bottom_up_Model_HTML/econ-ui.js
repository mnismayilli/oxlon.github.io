/* Ekonometrika — the ministry's EViews equations, re-estimated in the browser on the model's own data.
   Shows coefficients with diagnostics, lets them be changed by hand, and lets the admin add a realised year.
   Requires ev-engine.js and data/econ.js; everything model-specific arrives through init(cfg). */
(function (root) {
  'use strict';
  var E = root.EvEngine, DATA = null, CFG = null;
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function nf(v, d) { return CFG.nf ? CFG.nf(v, d) : (isNum(v) ? v.toFixed(d == null ? 3 : d) : '—'); }
  function sig(v, d) {
    if (!isNum(v)) return '—';
    var a = Math.abs(v);
    var t = (a !== 0 && (a < 1e-4 || a >= 1e7)) ? v.toExponential(2) : v.toFixed(d == null ? 4 : d);
    return t.replace('.', ',');
  }

  /* ------------------------------------------------------------ data from the model */
  var SERIES = null, SRC = null, BUILT = 0;
  function buildData(force) {
    if (SERIES && !force) return SERIES;
    var d = {}, src = {};
    var Y = CFG.years();
    Object.keys(DATA.vars).forEach(function (name) {
      var v = DATA.vars[name], gs = CFG.sheetIdx(v.b, v.s);
      if (gs < 0) return;
      var o = {}, cells = {}, n = 0;
      for (var y = Y.first; y <= Y.last; y++) {
        var c = CFG.colForYear(gs, v.r, y); if (c == null) continue;
        var id = CFG.cellId(gs, v.r, c); if (id < 0) continue;
        var val = CFG.value(id);
        if (isNum(val)) { o[y] = val; cells[y] = id; n++; }
      }
      if (n) { d[name] = o; src[name] = { kind: 'model', at: v.b + ' › ' + v.s + ' · sətir ' + v.r, gs: gs, row: v.r, cells: cells }; }
    });
    Object.keys(DATA.inline).forEach(function (name) {
      if (d[name]) return;
      d[name] = DATA.inline[name]; src[name] = { kind: 'csv', at: 'eviews_data.csv' };
    });
    // ECM / identity series computed from their own equations
    var ctx = new E.Ctx(d, DATA.alias);
    for (var pass = 0; pass < 3; pass++) {
      DATA.eqs.forEach(function (e) {
        var dep = E.parse(e.dep).ast;
        if (!dep || dep.k !== 'var') return;
        var nm = dep.name;
        if (d[nm] && src[nm] && src[nm].kind !== 'derived') return;
        var o = {}, n = 0;
        for (var y = Y.first - 2; y <= Y.last; y++) {
          var s = 0, ok = true;
          for (var i = 0; i < e.terms.length; i++) {
            var t = e.terms[i], val = t.t === '1' ? 1 : E.ev(pAst(t.t), y, ctx);
            if (val === undefined || !isFinite(val)) { ok = false; break; }
            s += t.c * val;
          }
          if (ok) { o[y] = s; n++; }
        }
        if (n >= 8) { d[nm] = o; src[nm] = { kind: 'derived', at: e.id + ' tənliyindən hesablanır' }; }
      });
    }
    SERIES = d; SRC = src; BUILT++;
    return d;
  }
  var AST = {};
  function pAst(s) { if (!(s in AST)) AST[s] = E.parse(s).ast; return AST[s]; }

  /* ----------------------------------------------------------------- estimation */
  var RES = null;
  function estimate(force) {
    if (RES && !force) return RES;
    var d = buildData(force), ctx = new E.Ctx(d, DATA.alias), Y = CFG.years();
    var out = {};
    DATA.eqs.forEach(function (e) {
      var r = { id: e.id };
      var dp = pAst(e.dep);
      // ECM_* rows are long-run definitions: the ministry computes the variable with this very formula,
      // so regressing it on its own terms would fit by construction
      if (dp && dp.k === 'var' && (/^ECM/i.test(dp.name) || (SRC[dp.name] && SRC[dp.name].kind === 'derived'))) {
        r.ok = false; r.identity = 1; r.why = 'təyinat tənliyidir — əmsallar qiymətləndirilmir, dəyişən elə bu düsturla hesablanır';
        out[e.id] = r; return;                     // estimating it against its own output would be circular
      }
      var vs = (e.vars || []).map(function (v) { return DATA.alias[v] || v; });
      var miss = vs.filter(function (v) { return !d[v]; });
      if (miss.length) { r.ok = false; r.why = 'məlumat yoxdur: ' + miss.slice(0, 4).join(', '); r.miss = miss; out[e.id] = r; return; }
      var dep = pAst(e.dep), y = [], X = e.terms.map(function () { return []; }), yrs = [];
      for (var yy = Y.first; yy <= Y.lastActual; yy++) {
        var dv = E.ev(dep, yy, ctx);
        if (dv === undefined || !isFinite(dv)) continue;
        var row = e.terms.map(function (t) { return t.t === '1' ? 1 : E.ev(pAst(t.t), yy, ctx); });
        var bad = row.some(function (v) { return v === undefined || !isFinite(v); });
        if (bad) continue;
        y.push(dv); row.forEach(function (v, i) { X[i].push(v); }); yrs.push(yy);
      }
      if (y.length < e.terms.length + 2) { r.ok = false; r.why = 'müşahidə azdır (' + y.length + ' il, ' + e.terms.length + ' hədd)'; out[e.id] = r; return; }
      var o = E.ols(y, X, e.terms.map(function (t) { return t.t; }));
      if (!o.ok) { r.ok = false; r.why = o.why; out[e.id] = r; return; }
      o.years = yrs; o.y = y;
      o.z = o.beta.map(function (b, i) { var s = e.terms[i].c, sd = o.se[i]; return (isFinite(sd) && sd > 0) ? Math.abs(b - s) / sd : NaN; });
      r.ok = true; r.est = o;
      out[e.id] = r;
    });
    RES = out;
    return out;
  }
  function flagsOf(e, r) {
    var f = [];
    if (!r || !r.ok) return f;
    var o = r.est;
    if (isFinite(o.dw) && (o.dw < 1.5 || o.dw > 2.5)) f.push('dw');
    if (isFinite(o.r2) && o.r2 < 0.3) f.push('fit');
    if (o.p.some(function (p, i) { return isFinite(p) && p > 0.1 && !e.terms[i].f; })) f.push('insig');
    if (o.z.some(function (z) { return isFinite(z) && z > 2; })) f.push('drift');
    return f;
  }

  /* ---------------------------------------------------------- coefficient cells */
  function coefId(e, t) {
    if (!t.cr) return -1;
    var gs = CFG.sheetIdx(e.b, e.s); if (gs < 0) return -1;
    return CFG.cellId(gs, t.cr, t.cc);
  }
  function currentCoef(e, t) { var id = coefId(e, t); return id >= 0 ? CFG.value(id) : t.c; }

  /* ------------------------------------------------------------------- styles */
  var SHEET = [
    '.ec-tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(168px,1fr));gap:10px;margin:16px 0 20px}',
    '.ec-tile{text-align:left;background:var(--bg,#fff);border:1px solid var(--line,#E7EBEF);border-radius:10px;padding:12px 13px;cursor:pointer;font:inherit;color:inherit;transition:border-color .12s,box-shadow .12s}',
    '.ec-tile:hover{border-color:#B8C2CC;box-shadow:0 1px 4px rgba(21,32,43,.07)}',
    '.ec-tile.on{border-color:#1F6FB2;box-shadow:inset 0 0 0 1px #1F6FB2}',
    '.ec-tile b{display:block;font-size:24px;line-height:1.1;font-variant-numeric:tabular-nums}',
    '.ec-tile span{display:block;font-size:12px;color:var(--muted,#71808D);margin-top:3px;line-height:1.35}',
    '.ec-tile i{font-style:normal;display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:5px;vertical-align:1px}',
    '.ec-bar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:12px}',
    '.ec-bar input[type=text]{font:inherit;font-size:13px;padding:7px 10px;border:1px solid var(--line,#E7EBEF);border-radius:7px;min-width:220px;background:var(--bg,#fff);color:inherit}',
    '.ec-list{display:flex;flex-direction:column;gap:7px}',
    '.ec-row{border:1px solid var(--line,#E7EBEF);border-radius:9px;background:var(--bg,#fff);overflow:hidden}',
    '.ec-head{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:10px;align-items:center;padding:10px 13px;cursor:pointer}',
    '.ec-head:hover{background:var(--bg2,#F5F7FC)}',
    '.ec-head h4{margin:0;font-size:13.5px;font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
    '.ec-head .sub{font-size:11.5px;color:var(--muted,#71808D);margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
    '.ec-stats{display:flex;gap:9px;align-items:center;font-size:11.5px;color:var(--muted,#71808D);white-space:nowrap;font-variant-numeric:tabular-nums}',
    '.ec-stats b{color:var(--ink,#15202B);font-weight:600}',
    '.ec-chip{font-size:10.5px;padding:2px 6px;border-radius:5px;border:1px solid var(--line,#E7EBEF);color:var(--muted,#71808D)}',
    '.ec-chip.warn{border-color:#E2B96B;color:#8A6216;background:#FDF7EA}',
    '.ec-chip.bad{border-color:#E0A6A0;color:#A3352A;background:#FCF1F0}',
    '.ec-chip.ok{border-color:#9FC4A8;color:#2F6B3D;background:#F1F8F3}',
    '.ec-body{border-top:1px solid var(--line,#E7EBEF);padding:13px;background:var(--bg2,#F9FBFD)}',
    '.ec-eq{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:11.5px;line-height:1.7;background:var(--bg,#fff);border:1px solid var(--line,#E7EBEF);border-radius:7px;padding:9px 11px;overflow-x:auto;white-space:pre-wrap;word-break:break-word}',
    '.ec-diag{display:grid;grid-template-columns:repeat(auto-fit,minmax(104px,1fr));gap:8px;margin:11px 0}',
    '.ec-d{background:var(--bg,#fff);border:1px solid var(--line,#E7EBEF);border-radius:7px;padding:7px 9px}',
    '.ec-d b{display:block;font-size:15px;font-variant-numeric:tabular-nums}',
    '.ec-d span{display:block;font-size:10.5px;color:var(--muted,#71808D);margin-top:1px}',
    '.ec-tbl{width:100%;border-collapse:collapse;font-size:12px;font-variant-numeric:tabular-nums;background:var(--bg,#fff)}',
    '.ec-tbl th,.ec-tbl td{padding:5px 7px;border-bottom:1px solid var(--line,#E7EBEF);text-align:right;white-space:nowrap}',
    '.ec-tbl th.l,.ec-tbl td.l{text-align:left;white-space:normal;word-break:break-word}',
    '.ec-tbl thead th{font-size:10.5px;text-transform:uppercase;letter-spacing:.03em;color:var(--muted,#71808D);font-weight:600}',
    '.ec-tbl input{font:inherit;font-size:12px;width:112px;padding:3px 6px;border:1px solid var(--line,#E7EBEF);border-radius:5px;text-align:right;background:var(--bg,#fff);color:inherit;font-variant-numeric:tabular-nums}',
    '.ec-tbl input.mod{border-color:#1F6FB2;background:#F0F6FC;font-weight:600}',
    '.ec-tbl .p0{color:#2F6B3D;font-weight:600}.ec-tbl .p1{color:#8A6216}.ec-tbl .p2{color:var(--muted,#71808D)}',
    '.ec-tw{overflow-x:auto;border:1px solid var(--line,#E7EBEF);border-radius:7px}',
    '.ec-act{display:flex;gap:7px;flex-wrap:wrap;margin-top:10px;align-items:center}',
    '.ec-note{font-size:11.5px;color:var(--muted,#71808D);margin-top:9px;line-height:1.5}',
    '.ec-none{padding:26px;text-align:center;color:var(--muted,#71808D);font-size:13px;border:1px dashed var(--line,#E7EBEF);border-radius:9px}',
    '.ec-fit{width:100%;height:54px;display:block}',
    '.ec-src{font-size:11px;color:var(--muted,#71808D)}',
    '.ec-dgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:9px}',
    '.ec-dcard{border:1px solid var(--line,#E7EBEF);border-radius:8px;padding:9px 11px;background:var(--bg,#fff)}',
    '.ec-dcard label{display:block;font-size:11.5px;color:var(--muted,#71808D);margin-bottom:4px}',
    '.ec-dcard input{font:inherit;font-size:13px;width:100%;padding:5px 8px;border:1px solid var(--line,#E7EBEF);border-radius:6px;text-align:right;background:var(--bg,#fff);color:inherit;font-variant-numeric:tabular-nums}',
    '.ec-dcard input.mod{border-color:#1F6FB2;background:#F0F6FC;font-weight:600}',
    '.ec-dcard .h{font-size:12.5px;font-weight:600;margin-bottom:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
  ].join('\n');
  function injectCSS() { if ($('#ec-css')) return; var s = document.createElement('style'); s.id = 'ec-css'; s.textContent = SHEET; document.head.appendChild(s); }

  /* --------------------------------------------------------------------- view */
  var FILTER = 'all', Q = '', OPEN = {}, TAB = 'eq';
  function tiles(res) {
    var eqs = DATA.eqs, nEst = 0, nIns = 0, nDw = 0, nDrift = 0, nFit = 0, nCoef = 0, nSigCoef = 0, nIdent = 0;
    eqs.forEach(function (e) {
      var r = res[e.id];
      if (r && r.identity) nIdent++;
      if (r && r.ok) {
        nEst++;
        var f = flagsOf(e, r);
        if (f.indexOf('insig') >= 0) nIns++;
        if (f.indexOf('dw') >= 0) nDw++;
        if (f.indexOf('drift') >= 0) nDrift++;
        if (f.indexOf('fit') >= 0) nFit++;
        r.est.p.forEach(function (p, i) { if (e.terms[i].f) return; nCoef++; if (isFinite(p) && p <= 0.05) nSigCoef++; });
      }
    });
    var T = [
      ['all', eqs.length, 'Tənlik', 'Nazirliyin EViews kataloqundan modelin «eq» vərəqlərinə köçürülüb', '#44525F'],
      ['est', nEst, 'Yenidən qiymətləndirildi', 'Bütün dəyişənləri modeldə var — OLS ilə yenidən hesablandı', '#1F6FB2'],
      ['ident', nIdent, 'Təyinat (ECM)', 'Uzunmüddətli əlaqəni təyin edən sətirlər — reqressiya deyil', '#7D6B12'],
      ['sig', nSigCoef + '/' + nCoef, 'Statistik əhəmiyyətli əmsal', 'p ≤ 0,05 — yeni qiymətləndirmədə', '#2F6B3D'],
      ['insig', nIns, 'Əhəmiyyətsiz həddi olan tənlik', 'Ən azı bir əmsal p > 0,10', '#8A6216'],
      ['drift', nDrift, 'Modeldəki əmsal fərqlənir', 'Modeldə istifadə olunan əmsal yeni qiymətin 95% intervalından kənardadır', '#A3352A'],
      ['dw', nDw, 'Avtokorrelyasiya şübhəsi', 'Durbin-Watson 1,5–2,5 aralığından kənar', '#8A6216'],
      ['fit', nFit, 'Zəif uyğunluq', 'R² < 0,30', '#71808D']
    ];
    return '<div class="ec-tiles">' + T.map(function (t) {
      return '<button class="ec-tile' + (FILTER === t[0] ? ' on' : '') + '" data-f="' + t[0] + '" title="' + esc(t[3]) + '">' +
        '<b>' + t[1] + '</b><span><i style="background:' + t[4] + '"></i>' + esc(t[2]) + '</span></button>';
    }).join('') + '</div>';
  }
  function fitChart(o) {
    if (!o || !o.years || o.years.length < 3) return '';
    var y = o.y, f = o.fitted, n = y.length, lo = Math.min.apply(null, y.concat(f)), hi = Math.max.apply(null, y.concat(f));
    var W = 280, H = 54, P = 4, sp = (hi - lo) || 1;
    function path(a) { return a.map(function (v, i) { return (i ? 'L' : 'M') + (P + (W - 2 * P) * i / (n - 1)).toFixed(1) + ' ' + (H - P - (H - 2 * P) * (v - lo) / sp).toFixed(1); }).join(''); }
    return '<svg class="ec-fit" viewBox="0 0 ' + W + ' ' + H + '" preserveAspectRatio="none" aria-hidden="true">' +
      '<path d="' + path(y) + '" fill="none" stroke="#15202B" stroke-width="1.3"/>' +
      '<path d="' + path(f) + '" fill="none" stroke="#1F6FB2" stroke-width="1.3" stroke-dasharray="3 2"/></svg>';
  }
  function eqHTML(e, res) {
    var r = res[e.id], open = !!OPEN[e.id], f = flagsOf(e, r);
    var stats = '';
    if (r && r.ok) {
      var o = r.est;
      stats = '<span>n <b>' + o.n + '</b></span><span>R² <b>' + sig(o.r2, 3) + '</b></span><span>DW <b>' + sig(o.dw, 2) + '</b></span>';
      if (f.indexOf('drift') >= 0) stats += '<span class="ec-chip bad">əmsal fərqlənir</span>';
      else stats += '<span class="ec-chip ok">uyğun</span>';
      if (f.indexOf('insig') >= 0) stats += '<span class="ec-chip warn">əhəmiyyətsiz hədd</span>';
    } else if (r && r.identity) stats = '<span class="ec-chip">təyinat</span>';
    else stats = '<span class="ec-chip">yalnız sənəd</span>';
    var h = '<div class="ec-row" id="ec-' + esc(e.id) + '"><div class="ec-head" data-eq="' + esc(e.id) + '" role="button" tabindex="0" aria-expanded="' + open + '">' +
      '<div><h4>' + esc(e.dep) + '</h4><div class="sub">' + esc(e.desc || e.id) + ' · ' + esc(e.b + ' › ' + e.s) + ' · sətir ' + e.row + '</div></div>' +
      '<div class="ec-stats">' + stats + '</div></div>';
    if (open) h += '<div class="ec-body">' + bodyHTML(e, r) + '</div>';
    return h + '</div>';
  }
  function bodyHTML(e, r) {
    var h = '<div class="ec-eq">' + esc(e.dep) + ' = ' + e.terms.map(function (t, i) {
      return (i ? '\n    + ' : '') + sig(currentCoef(e, t), 6) + (t.t === '1' ? '' : ' · ' + t.t);
    }).join('') + '</div>';
    if (!r || !r.ok) {
      h += r && r.identity
        ? '<p class="ec-note"><b>Təyinat (identity).</b> Bu sətir ' + esc(e.dep) + ' dəyişənini təyin edir — uzunmüddətli (ECM) əlaqədir, reqressiya deyil. Əmsallar Nazirliyin qiymətləndirməsindən gəlir və aşağıda əl ilə dəyişdirilə bilər; dəyişiklik ondan asılı bütün tənliklərə ötürülür.</p>'
        : '<p class="ec-note"><b>Yenidən qiymətləndirilmədi:</b> ' + esc(r ? r.why : '—') + '. Əmsallar Nazirliyin EViews nəticəsindən götürülüb və aşağıda əl ilə dəyişdirilə bilər.</p>';
      h += coefTable(e, null);
      return h;
    }
    var o = r.est;
    h += '<div class="ec-diag">' +
      d('R²', sig(o.r2, 4)) + d('Düzəlişli R²', sig(o.adjR2, 4)) + d('Reqressiyanın S.X.', sig(o.ser, 5)) +
      d('Qalıqların kvadratı', sig(o.ssr, 5)) + d('Durbin-Watson', sig(o.dw, 3)) + d('F-statistika', sig(o.f, 3)) +
      d('Prob(F)', isFinite(o.fp) ? sig(o.fp, 4) : '—') + d('Müşahidə', String(o.n)) + d('Sərbəstlik dər.', String(o.df)) +
      d('Nümunə', o.years[0] + '–' + o.years[o.years.length - 1]) + '</div>';
    h += coefTable(e, o);
    h += '<div style="display:flex;gap:14px;align-items:center;margin-top:11px;flex-wrap:wrap">' + fitChart(o) +
      '<span class="ec-src"><span style="display:inline-block;width:14px;border-top:1.3px solid #15202B;vertical-align:4px"></span> faktiki &nbsp; <span style="display:inline-block;width:14px;border-top:1.3px dashed #1F6FB2;vertical-align:4px"></span> tənliyin verdiyi</span></div>';
    h += '<div class="ec-act"><button class="btn sm" data-apply="' + esc(e.id) + '">Yeni qiymətləri modelə yaz</button>' +
      '<button class="btn sm" data-revert="' + esc(e.id) + '">Nazirliyin əmsallarına qaytar</button>' +
      '<span class="ec-note" style="margin:0">Dəyişiklik modelin «' + esc(e.s) + '» vərəqindəki xanalara yazılır və bütün proqnoza ötürülür.</span></div>';
    return h;
  }
  function d(t, v) { return '<div class="ec-d"><b>' + v + '</b><span>' + t + '</span></div>'; }
  function coefTable(e, o) {
    var h = '<div class="ec-tw"><table class="ec-tbl"><thead><tr><th class="l">Hədd</th><th>Modeldə (dəyişdirilə bilər)</th>' +
      (o ? '<th>Yeni qiymət</th><th>Std. xəta</th><th>t</th><th>Prob.</th><th>|Δ|/s.x.</th>' : '') + '<th class="l">Xana</th></tr></thead><tbody>';
    e.terms.forEach(function (t, i) {
      var id = coefId(e, t), cur = currentCoef(e, t), ed = id >= 0 && CFG.isEdited(id);
      h += '<tr><td class="l">' + (t.t === '1' ? '<i>sabit</i>' : esc(t.t)) + '</td>';
      h += '<td>' + (t.f ? '<span class="ec-src">1 (sabit)</span>' : id >= 0
        ? '<input type="text" data-coef="' + esc(e.id) + '|' + i + '" value="' + sig(cur, 6) + '"' + (ed ? ' class="mod"' : '') + ' aria-label="Əmsal">'
        : '<span class="ec-src">' + sig(cur, 6) + '</span>') + '</td>';
      if (o) {
        var p = o.p[i], cls = !isFinite(p) ? 'p2' : p <= 0.05 ? 'p0' : p <= 0.1 ? 'p1' : 'p2';
        h += '<td>' + sig(o.beta[i], 6) + '</td><td>' + sig(o.se[i], 5) + '</td><td>' + sig(o.t[i], 2) + '</td>' +
          '<td class="' + cls + '">' + (isFinite(p) ? sig(p, 4) : '—') + '</td>' +
          '<td>' + (isFinite(o.z[i]) ? (o.z[i] > 2 ? '<b style="color:#A3352A">' + sig(o.z[i], 1) + '</b>' : sig(o.z[i], 1)) : '—') + '</td>';
      }
      h += '<td class="l"><span class="ec-src">' + (t.cr ? esc(colA(t.cc) + t.cr) : '—') + '</span></td></tr>';
    });
    return h + '</tbody></table></div>';
  }
  function colA(n) { var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

  /* ------------------------------------------------- admin: add a realised year */
  var DQ = '', DYEAR = null;
  function dataYear() {
    if (DYEAR == null) DYEAR = CFG.years().lastActual + 1;
    return DYEAR;
  }
  function dataHTML() {
    buildData();
    var Y = CFG.years(), y = dataYear();
    var names = Object.keys(SRC).filter(function (n) { return SRC[n].kind === 'model'; }).sort();
    var q = DQ.toLowerCase();
    var show = names.filter(function (n) { return !q || n.toLowerCase().indexOf(q) >= 0 || SRC[n].at.toLowerCase().indexOf(q) >= 0; });
    var filled = 0, total = 0;
    names.forEach(function (n) { total++; var gs = SRC[n].gs, c = CFG.colForYear(gs, SRC[n].row, y); var id = c == null ? -1 : CFG.cellId(gs, SRC[n].row, c); if (id >= 0 && CFG.isEdited(id)) filled++; });
    var yrs = [];
    for (var yy = Y.lastActual - 1; yy <= Y.last; yy++) yrs.push(yy);
    var h = '<p class="lead">Faktiki məlumat gələndə onu burada göstərici üzrə daxil edin. Dəyər modelin öz xanasına yazılır, bütün proqnoz yenilənir və ekonometrik tənliklər yeni müşahidə ilə yenidən qiymətləndirilə bilər.</p>';
    h += '<div class="ec-bar"><label class="ec-src" for="ec-dy">İl:</label><select id="ec-dy" style="font:inherit;font-size:13px;padding:6px 9px;border:1px solid var(--line,#E7EBEF);border-radius:7px;background:var(--bg,#fff);color:inherit">' +
      yrs.map(function (a) { return '<option value="' + a + '"' + (a === y ? ' selected' : '') + '>' + a + (a <= Y.lastActual ? ' (faktiki)' : ' (proqnoz)') + '</option>'; }).join('') + '</select>' +
      '<input type="text" id="ec-dq" placeholder="Göstərici axtar…" value="' + esc(DQ) + '">' +
      '<span class="ec-src">' + show.length + ' göstərici · bu ildə ' + filled + ' dəyişiklik</span>' +
      '<button class="btn sm" id="ec-rerun" style="margin-left:auto">Tənlikləri yenidən qiymətləndir</button></div>';
    if (!show.length) return h + '<div class="ec-none">Uyğun göstərici tapılmadı.</div>';
    h += '<div class="ec-dgrid">' + show.map(function (n) {
      var s = SRC[n], c = CFG.colForYear(s.gs, s.row, y), id = c == null ? -1 : CFG.cellId(s.gs, s.row, c);
      var v = id >= 0 ? CFG.value(id) : null, ed = id >= 0 && CFG.isEdited(id);
      var prev = SERIES[n] ? SERIES[n][y - 1] : null;
      return '<div class="ec-dcard"><div class="h" title="' + esc(s.at) + '">' + esc(n) + '</div>' +
        '<label>' + esc(s.at) + '</label>' +
        (id < 0 ? '<span class="ec-src">bu il üçün xana yoxdur</span>'
          : '<input type="text" data-din="' + esc(n) + '" value="' + (isNum(v) ? sig(v, 6) : '') + '"' + (ed ? ' class="mod"' : '') + ' aria-label="' + esc(n) + ' ' + y + '">') +
        '<div class="ec-src" style="margin-top:4px">' + (y - 1) + ': ' + (isNum(prev) ? sig(prev, 4) : '—') + '</div></div>';
    }).join('') + '</div>';
    return h;
  }

  /* --------------------------------------------------------------------- page */
  function page(v) {
    injectCSS();
    if (!DATA) DATA = root.ECON_DATA;
    if (!DATA) { v.innerHTML = '<div class="ec-none">Ekonometrika məlumatı yüklənmədi (data/econ.js).</div>'; return; }
    var res = estimate();
    var h = '<div class="eyebrow">Ekonometrika</div><h1 class="h1">Tənliklər və əmsallar</h1>';
    h += '<p class="lead">Nazirliyin EViews kataloqundakı <b>' + DATA.eqs.length + '</b> davranış tənliyi modelin «eq» vərəqlərindən oxunur. Dəyişənlərin məlumatı ayrıca fayldan deyil, <b>modelin öz vərəqlərindən</b> götürülür (' + DATA.meta.fromModel + ' sıra), ona görə faktiki məlumat yeniləndikdə qiymətləndirmə də yenilənir.</p>';
    h += '<div class="seg" id="ec-tab" role="group" aria-label="Bölmə" style="margin:12px 0 4px">' +
      '<button data-t="eq" class="' + (TAB === 'eq' ? 'on' : '') + '">Tənliklər</button>' +
      '<button data-t="data" class="' + (TAB === 'data' ? 'on' : '') + '">Faktiki məlumatın daxil edilməsi</button></div>';
    if (TAB === 'data') { v.innerHTML = h + dataHTML(); wireData(v); return; }
    h += tiles(res);
    h += '<div class="ec-bar"><input type="text" id="ec-q" placeholder="Tənlik və ya dəyişən axtar…" value="' + esc(Q) + '">' +
      '<button class="btn sm" id="ec-reest">Yenidən qiymətləndir</button>' +
      '<span class="ec-src" id="ec-count"></span></div>';
    var list = filtered(res);
    h += '<div class="ec-list" id="ec-list">' + (list.length ? list.map(function (e) { return eqHTML(e, res); }).join('') : '<div class="ec-none">Seçimə uyğun tənlik yoxdur.</div>') + '</div>';
    h += '<p class="ec-note" style="margin-top:14px">Qeyd: «Yeni qiymət» sütunu modelin cari faktiki məlumatı ilə sadə EKÜ (OLS) nəticəsidir. Nazirliyin əmsalları başqa vaxtda, başqa məlumat versiyası ilə qiymətləndirilib, ona görə fərq ola bilər — fərqin statistik ölçüsü «|Δ|/s.x.» sütunundadır (2-dən böyükdürsə, fərq təsadüfi deyil).</p>';
    v.innerHTML = h;
    var cnt = $('#ec-count'); if (cnt) cnt.textContent = list.length + ' tənlik';
    wireEq(v, res);
  }
  function filtered(res) {
    var q = Q.toLowerCase();
    return DATA.eqs.filter(function (e) {
      var r = res[e.id], f = flagsOf(e, r);
      if (FILTER === 'est' && !(r && r.ok)) return false;
      if (FILTER === 'ident' && !(r && r.identity)) return false;
      if (FILTER === 'insig' && f.indexOf('insig') < 0) return false;
      if (FILTER === 'dw' && f.indexOf('dw') < 0) return false;
      if (FILTER === 'drift' && f.indexOf('drift') < 0) return false;
      if (FILTER === 'fit' && f.indexOf('fit') < 0) return false;
      if (FILTER === 'sig' && !(r && r.ok)) return false;
      if (q) {
        var hay = (e.id + ' ' + e.dep + ' ' + (e.desc || '') + ' ' + (e.vars || []).join(' ')).toLowerCase();
        if (hay.indexOf(q) < 0) return false;
      }
      return true;
    });
  }
  function wireTabs(v) {
    var t = $('#ec-tab');
    if (t) t.addEventListener('click', function (e) { var b = e.target.closest('[data-t]'); if (!b) return; TAB = b.getAttribute('data-t'); page(v); });
  }
  function wireEq(v, res) {
    wireTabs(v);
    var tl = v.querySelector('.ec-tiles');
    if (tl) tl.addEventListener('click', function (e) { var b = e.target.closest('[data-f]'); if (!b) return; FILTER = b.getAttribute('data-f'); page(v); });
    var q = $('#ec-q');
    if (q) q.addEventListener('input', function () { Q = q.value; var l = $('#ec-list'); var list = filtered(res); l.innerHTML = list.length ? list.map(function (x) { return eqHTML(x, res); }).join('') : '<div class="ec-none">Seçimə uyğun tənlik yoxdur.</div>'; var c = $('#ec-count'); if (c) c.textContent = list.length + ' tənlik'; });
    var re = $('#ec-reest');
    if (re) re.addEventListener('click', function () { estimate(true); CFG.toast('Tənliklər modelin cari məlumatı ilə yenidən qiymətləndirildi'); page(v); });
    var l = $('#ec-list');
    if (!l) return;
    l.addEventListener('click', function (e) {
      var hd = e.target.closest('[data-eq]');
      if (hd && !e.target.closest('input')) { var id = hd.getAttribute('data-eq'); OPEN[id] = !OPEN[id]; page(v); return; }
      var ap = e.target.closest('[data-apply]');
      if (ap) { applyAll(ap.getAttribute('data-apply'), res); page(v); return; }
      var rv = e.target.closest('[data-revert]');
      if (rv) { revertAll(rv.getAttribute('data-revert')); page(v); return; }
    });
    l.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter') return;
      var inp = e.target.closest('[data-coef]');
      if (inp) { e.preventDefault(); commitCoef(inp); page(v); }
    });
    l.addEventListener('change', function (e) { var inp = e.target.closest('[data-coef]'); if (inp) { commitCoef(inp); page(v); } });
  }
  function parseNum(s) { var t = String(s).trim().replace(/\s/g, '').replace(',', '.'); return t === '' ? null : (isFinite(+t) ? +t : NaN); }
  function eqById(id) { for (var i = 0; i < DATA.eqs.length; i++) if (DATA.eqs[i].id === id) return DATA.eqs[i]; return null; }
  function commitCoef(inp) {
    var k = inp.getAttribute('data-coef').split('|'), e = eqById(k[0]), i = +k[1];
    if (!e) return;
    var t = e.terms[i], id = coefId(e, t); if (id < 0) return;
    var v = parseNum(inp.value);
    if (v === null || isNaN(v)) { CFG.toast('Rəqəm daxil edin'); return; }
    CFG.setCell(id, v); CFG.changed();
    CFG.toast(e.dep + ' · əmsal dəyişdirildi: ' + sig(v, 6));
  }
  function applyAll(id, res) {
    var e = eqById(id), r = res[id];
    if (!e || !r || !r.ok) return;
    var n = 0;
    e.terms.forEach(function (t, i) { var cid = coefId(e, t); if (cid < 0 || t.f) return; CFG.setCell(cid, r.est.beta[i]); n++; });
    CFG.changed();
    CFG.toast(n + ' əmsal yeni qiymətləndirmə ilə əvəz olundu');
  }
  function revertAll(id) {
    var e = eqById(id); if (!e) return;
    var n = 0;
    e.terms.forEach(function (t) { var cid = coefId(e, t); if (cid < 0) return; if (CFG.isEdited(cid)) { CFG.resetCell(cid); n++; } });
    CFG.changed();
    CFG.toast(n ? n + ' əmsal Nazirliyin qiymətinə qaytarıldı' : 'Dəyişiklik yox idi');
  }
  function wireData(v) {
    wireTabs(v);
    var sel = $('#ec-dy');
    if (sel) sel.addEventListener('change', function () { DYEAR = +sel.value; page(v); });
    var q = $('#ec-dq');
    if (q) q.addEventListener('input', function () { DQ = q.value; var g = v.querySelector('.ec-dgrid'); if (g) { page(v); var n = $('#ec-dq'); if (n) { n.focus(); n.setSelectionRange(n.value.length, n.value.length); } } });
    var rr = $('#ec-rerun');
    if (rr) rr.addEventListener('click', function () { estimate(true); TAB = 'eq'; CFG.toast('Yeni məlumatla yenidən qiymətləndirildi'); page(v); });
    v.addEventListener('change', function (e) {
      var inp = e.target.closest('[data-din]'); if (!inp) return;
      var n = inp.getAttribute('data-din'), s = SRC[n], y = dataYear();
      var c = CFG.colForYear(s.gs, s.row, y), id = c == null ? -1 : CFG.cellId(s.gs, s.row, c);
      if (id < 0) return;
      var val = parseNum(inp.value);
      if (val === null) { CFG.resetCell(id); CFG.changed(); SERIES = null; RES = null; CFG.toast(n + ' ilkin dəyərə qaytarıldı'); page(v); return; }
      if (isNaN(val)) { CFG.toast('Rəqəm daxil edin'); return; }
      CFG.setCell(id, val); CFG.changed(); SERIES = null; RES = null;
      CFG.toast(n + ' · ' + y + ' = ' + sig(val, 6));
      inp.classList.add('mod');
    });
  }

  root.EconUI = {
    init: function (cfg) { CFG = cfg; DATA = root.ECON_DATA; SERIES = null; RES = null; },
    page: page,
    invalidate: function () { SERIES = null; RES = null; },
    estimate: estimate
  };
})(typeof window !== 'undefined' ? window : globalThis);
