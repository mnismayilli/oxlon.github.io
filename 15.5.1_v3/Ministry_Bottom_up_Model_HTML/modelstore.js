/* Makro model — model store: active model selection (bundled or uploaded), year extension (insert column + fill right),
   realised data (actuals) and baseline. Shared by index.html and panel.html. Requires engine.js and model.js. */
(function (root) {
  'use strict';
  var LS_STRUCT = 'makroModel.structure.v1', LS_ACT = 'makroModel.actuals.v1', LS_EDITS = 'makroModel.edits.v1', LS_SCEN = 'makroModel.scenarios.v1';
  var BASE_END = 2030, BASE_LAST_ACTUAL = 2024;
  function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); return true; } catch (e) { return false; } }
  function lsJSON(k, d) { try { var s = lsGet(k); return s ? JSON.parse(s) : d; } catch (e) { return d; } }

  // ------------------------------------------------------------------ IndexedDB (uploaded model)
  var DB = 'makroModel', STORE = 'models';
  function idb() {
    return new Promise(function (res, rej) {
      if (!root.indexedDB) { rej(new Error('IndexedDB yoxdur')); return; }
      var rq = indexedDB.open(DB, 1);
      rq.onupgradeneeded = function () { rq.result.createObjectStore(STORE); };
      rq.onsuccess = function () { res(rq.result); };
      rq.onerror = function () { rej(rq.error); };
    });
  }
  function idbOp(mode, fn) {
    return idb().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction(STORE, mode), st = tx.objectStore(STORE), rq = fn(st);
        tx.oncomplete = function () { res(rq && rq.result); db.close(); };
        tx.onerror = function () { rej(tx.error); db.close(); };
      });
    });
  }
  function idbGet(k) { return idbOp('readonly', function (s) { return s.get(k); }); }
  function idbPut(k, v) { return idbOp('readwrite', function (s) { return s.put(v, k); }); }
  function idbDel(k) { return idbOp('readwrite', function (s) { return s.delete(k); }); }

  // ------------------------------------------------------------------ helpers on flat model state
  function colName(n) { var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function yearOfValue(v) {
    if (isNum(v) && v === Math.floor(v) && v >= 1990 && v <= 2060) return v;
    if (typeof v === 'string') { var m = /^\s*((?:19|20)\d\d)\s*\D{0,12}$/.exec(v); if (m) return +m[1]; }
    return null;
  }
  // year header rows of one sheet: rows with >=3 consecutive increasing years in consecutive columns
  function yearRowsOf(ids, R, C, V) {
    var rows = {};
    ids.forEach(function (id) { var y = yearOfValue(V[id]); if (y) (rows[R[id]] || (rows[R[id]] = {}))[C[id]] = y; });
    var out = [];
    Object.keys(rows).forEach(function (r) {
      var d = rows[r], cs = Object.keys(d).map(Number).sort(function (a, b) { return a - b; }), run = 0;
      for (var i = 1; i < cs.length; i++) if (cs[i] === cs[i - 1] + 1 && d[cs[i]] === d[cs[i - 1]] + 1) run++;
      if (run >= 3) out.push({ r: +r, d: d });
    });
    return out.sort(function (a, b) { return a.r - b.r; });
  }
  function colYearIn(yr, r, c) {
    var best = null, first = null;
    for (var i = 0; i < yr.length; i++) { if (yr[i].d[c] === undefined) continue; if (first === null) first = yr[i].d[c]; if (yr[i].r <= r) best = yr[i].d[c]; }
    return best !== null ? best : first;
  }
  function flatten(core, parts) {
    var S = { core: core, N: core.ncells, sh: [], r: [], c: [], v: [], st: [], F: {}, K: core.K.slice(), ranges: core.ranges.map(function (x) { return x.slice(); }),
      sheets: JSON.parse(JSON.stringify(core.sheets)) };
    parts.forEach(function (P) {
      for (var i = 0; i < P.ids.length; i++) { var id = P.ids[i]; S.sh[id] = P.sh[i]; S.r[id] = P.r[i]; S.c[id] = P.c[i]; S.v[id] = P.v[i] === undefined ? null : P.v[i]; S.st[id] = P.st[i]; }
      for (var k in P.f) S.F[k] = P.f[k];
    });
    for (var j = 0; j < S.N; j++) if (S.sh[j] === undefined) { S.sh[j] = 0; S.r[j] = 1; S.c[j] = 1; S.v[j] = null; S.st[j] = 0; }
    S.tplNL = core.disp.map(function (d) { return (d.match(/\x01/g) || []).length; });
    S.tplFlags = core.disp.map(function (d) { var out = [], re = /\x02([^\x03]*)\x03/g, m; while ((m = re.exec(d))) out.push(m[1]); return out; });
    return S;
  }
  function buildSmap(S) {
    S.smap = S.sheets.map(function () { return new Map(); });
    for (var i = 0; i < S.N; i++) S.smap[S.sh[i]].set(S.r[i] * 20000 + S.c[i], i);
  }
  function a1Map(a1, f) { var m = /^([A-Z]+)(\d+)$/.exec(a1); return m ? colName(f(colNum(m[1]))) + m[2] : a1; }

  // ------------------------------------------------------------------ year extension: insert a column after the last year and fill right
  // label text of a row (string cells left of the data) — used to recognise percentage / growth-rate rows
  var RATE_RE = /%|artım|tempi|temp\b|growth|yoy|faiz|dərəcə|rate|nisbət|ratio|çəki|share|\bpay\b|deflyator|deflator|əmsal|coef/i;
  function extendOne(S, mode) {
    var END = S.end, core = S.core, byS = S.sheets.map(function () { return []; });
    for (var i = 0; i < S.N; i++) byS[S.sh[i]].push(i);
    var plan = {}, skipped = [];
    S.sheets.forEach(function (sm, gs) {
      if (core.books[sm.b].kind !== 'model') return;
      var yr = yearRowsOf(byS[gs], S.r, S.c, S.v); if (!yr.length) return;
      // irregular layouts (the same year in several columns, e.g. forecast-version comparisons) cannot be filled right
      // (the years just before the last one appear in several columns); separate historical blocks in the same row are fine
      var irregular = yr.some(function (y) {
        var cnt = {}; Object.keys(y.d).forEach(function (c) { cnt[y.d[c]] = (cnt[y.d[c]] || 0) + 1; });
        return cnt[END] >= 1 && ((cnt[END - 1] || 0) > 1 || (cnt[END - 2] || 0) > 1);
      });
      if (irregular) { skipped.push(core.books[sm.b].n + ' › ' + sm.n); return; }
      var yrSet = {}; yr.forEach(function (y) { yrSet[y.r] = 1; });
      var src = byS[gs].filter(function (id) { return (S.v[id] !== null || S.F[id]) && colYearIn(yr, S.r[id], S.c[id]) === END; });
      if (!src.length) return;
      var X = []; src.forEach(function (id) { var x = S.c[id] + 1; if (X.indexOf(x) < 0) X.push(x); });
      X.sort(function (a, b) { return a - b; });
      plan[gs] = { X: X, src: src, yrSet: yrSet };
    });
    function cmap(X, c) { var n = 0; for (var i = 0; i < X.length; i++) if (X[i] <= c) n++; return c + n; }
    var stepInfo = {};
    var rowText = {};
    Object.keys(plan).forEach(function (k) {
      var gs = +k; byS[gs].forEach(function (id) { if (typeof S.v[id] === 'string' && S.c[id] <= 8) { var key = gs + ':' + S.r[id]; rowText[key] = (rowText[key] || '') + ' ' + S.v[id]; } });
    });
    var leftVal = {};
    Object.keys(plan).forEach(function (k) {
      var gs = +k, m = new Map(); byS[gs].forEach(function (id) { m.set(S.r[id] * 20000 + S.c[id], id); });
      plan[gs].src.forEach(function (id) { var l = m.get(S.r[id] * 20000 + S.c[id] - 1); leftVal[id] = l === undefined ? null : S.v[l]; });
    });
    Object.keys(plan).forEach(function (k) {
      var gs = +k, X = plan[gs].X, sm = S.sheets[gs];
      stepInfo[core.books[sm.b].n + '|' + sm.n] = X.slice();
      byS[gs].forEach(function (id) { S.c[id] = cmap(X, S.c[id]); });
      S.ranges.forEach(function (R) { if (R[0] === gs) { R[2] = cmap(X, R[2]); R[4] = cmap(X, R[4]); } });
      if (sm.cols) {
        sm.cols.forEach(function (cd) { cd.min = cmap(X, cd.min); cd.max = cmap(X, cd.max); });
        var widthAt = function (c) { var w = null; sm.cols.forEach(function (cd) { if (c >= cd.min && c <= cd.max) w = cd.w; }); return w; };
        X.forEach(function (x, i) { var nc = x + i, w = widthAt(nc - 1); if (w !== null) sm.cols.push({ min: nc, max: nc, w: w, hidden: false }); });
      }
      if (sm.fz && sm.fz[1]) sm.fz[1] = cmap(X, sm.fz[1]);
      if (sm.cm) { var cm = {}; Object.keys(sm.cm).forEach(function (a) { cm[a1Map(a, function (c) { return cmap(X, c); })] = sm.cm[a]; }); sm.cm = cm; }
      if (sm.merges) sm.merges = sm.merges.map(function (m) { return m.split(':').map(function (a) { return a1Map(a, function (c) { return cmap(X, c); }); }).join(':'); });
    });
    buildSmap(S);
    // pass 1: create the new year cells
    var pairs = [];
    Object.keys(plan).forEach(function (k) {
      var gs = +k;
      plan[gs].src.forEach(function (id) {
        var nid = S.N++;
        S.sh[nid] = gs; S.r[nid] = S.r[id]; S.c[nid] = S.c[id] + 1; S.st[nid] = S.st[id]; S.v[nid] = null;
        S.smap[gs].set(S.r[nid] * 20000 + S.c[nid], nid);
        pairs.push([id, nid, plan[gs].yrSet[S.r[id]] === 1]);
      });
    });
    function cellAt(gs, r, c) {
      var k = r * 20000 + c, x = S.smap[gs].get(k);
      if (x !== undefined) return x;
      x = S.N++; S.sh[x] = gs; S.r[x] = r; S.c[x] = c; S.v[x] = null; S.st[x] = 0; S.smap[gs].set(k, x); return x;
    }
    var rIdx = {}; S.ranges.forEach(function (R, i) { rIdx[R.join(',')] = i; });
    function rangeAt(R) { var k = R.join(','); if (rIdx[k] !== undefined) return rIdx[k]; rIdx[k] = S.ranges.length; S.ranges.push(R); return rIdx[k]; }
    // pass 2: formulas (relative column references shift by one) and carried-forward values
    pairs.forEach(function (p) {
      var id = p[0], nid = p[1], header = p[2], f = S.F[id];
      var isHeaderYear = header && yearOfValue(S.v[id]) === END;
      if (f) {
        var t = f[0], flags = S.tplFlags[t], refs = f[1].map(function (ref, j) {
          var fl = flags[j] || '';
          if (fl.charAt(0) === 'P') fl = fl.slice(1);
          if (ref >= 0) return fl.indexOf('C') >= 0 ? ref : cellAt(S.sh[ref], S.r[ref], S.c[ref] + 1);
          var R = S.ranges[-ref - 1], fs = fl.slice(1).split(':');
          var nr = [R[0], R[1], (fs[0] || '').indexOf('C') >= 0 ? R[2] : R[2] + 1, R[3], (fs[1] || '').indexOf('C') >= 0 ? R[4] : R[4] + 1];
          return -(rangeAt(nr) + 1);
        });
        var kb = S.K.length, nl = S.tplNL[t];
        for (var q = 0; q < nl; q++) S.K.push(S.K[f[2] + q]);
        S.F[nid] = [t, refs, kb];
        S.v[nid] = isHeaderYear ? END + 1 : null;
      } else {
        var v = S.v[id];
        if (isHeaderYear) v = typeof v === 'number' ? END + 1 : String(v).replace(String(END), String(END + 1));
        else if (mode === 'trend' && isNum(v)) {
          // continue the last year-on-year growth for level variables; rates, shares and growth rows stay flat
          var lv = leftVal[id], txt = rowText[S.sh[id] + ':' + S.r[id]] || '';
          if (!RATE_RE.test(txt) && isNum(lv) && lv !== 0 && v !== 0 && (lv > 0) === (v > 0)) {
            var g = v / lv; if (g > 0.5 && g < 1.5) v = v * g;
          }
        }
        S.v[nid] = v;
      }
    });
    S.end = END + 1;
    return { ins: stepInfo, skipped: skipped };
  }
  function topoOrder(S) {
    var ids = Object.keys(S.F).map(Number), isF = new Uint8Array(S.N), indeg = new Int32Array(S.N), out = [];
    ids.forEach(function (id) { isF[id] = 1; });
    var succ = new Array(S.N);
    function members(R) { var a = [], m = S.smap[R[0]]; for (var r = R[1]; r <= R[3]; r++) for (var c = R[2]; c <= R[4]; c++) { var x = m.get(r * 20000 + c); if (x !== undefined) a.push(x); } return a; }
    var memo = {};
    ids.forEach(function (id) {
      var seen = {};
      S.F[id][1].forEach(function (ref) {
        var ps = ref >= 0 ? [ref] : (memo[ref] || (memo[ref] = members(S.ranges[-ref - 1])));
        ps.forEach(function (p) { if (isF[p] && !seen[p]) { seen[p] = 1; (succ[p] || (succ[p] = [])).push(id); indeg[id]++; } });
      });
    });
    var q = ids.filter(function (id) { return indeg[id] === 0; }), h = 0;
    while (h < q.length) { var x = q[h++]; out.push(x); (succ[x] || []).forEach(function (y) { if (--indeg[y] === 0) q.push(y); }); }
    if (out.length !== ids.length) {
      var bad = ids.filter(function (id) { return indeg[id] > 0; }).slice(0, 5).map(function (id) { var sm = S.sheets[S.sh[id]]; return S.core.books[sm.b].n + ' › ' + sm.n + '!' + colName(S.c[id]) + S.r[id]; });
      throw new Error('Dövri istinad yarandı: ' + bad.join(', '));
    }
    return out;
  }
  function unflatten(S, order) {
    var core = {}; for (var k in S.core) core[k] = S.core[k];
    core.sheets = S.sheets; core.ranges = S.ranges; core.K = S.K; core.ncells = S.N; core.order = order;
    var P = { ids: [], sh: S.sh.slice(0, S.N), r: S.r.slice(0, S.N), c: S.c.slice(0, S.N), v: S.v.slice(0, S.N), st: S.st.slice(0, S.N), f: S.F };
    for (var i = 0; i < S.N; i++) P.ids.push(i);
    return { core: core, parts: [P] };
  }
  function extend(core, parts, n, startEnd, modes) {
    var S = flatten(core, parts); S.end = startEnd; buildSmap(S);
    var steps = [];
    for (var i = 0; i < n; i++) steps.push(extendOne(S, (modes && modes[i]) || 'flat'));
    var order = topoOrder(S), res = unflatten(S, order);
    res.steps = steps; res.end = S.end;
    return res;
  }

  // ------------------------------------------------------------------ key migration of stored edits when columns are inserted/removed
  function mapKey(key, fn) {
    var p = key.split('|'); if (p.length < 3) return key;
    var hash = '', a = p[2], hi = a.indexOf('#'); if (hi >= 0) { hash = a.slice(hi); a = a.slice(0, hi); }
    var m = /^([A-Z]+)(\d+)$/.exec(a); if (!m) return key;
    var nc = fn(p[0] + '|' + p[1], colNum(m[1])); if (nc === null) return null;
    return p[0] + '|' + p[1] + '|' + colName(nc) + m[2] + hash;
  }
  function mapObj(o, fn) { var out = {}; Object.keys(o || {}).forEach(function (k) { var nk = mapKey(k, fn); if (nk !== null) out[nk] = o[k]; }); return out; }
  function migrateAll(fn) {
    var e = lsJSON(LS_EDITS, null);
    if (e) lsSet(LS_EDITS, JSON.stringify({ inputs: mapObj(e.inputs, fn), K: mapObj(e.K, fn), ov: mapObj(e.ov, fn) }));
    lsSet(LS_ACT, JSON.stringify(mapObj(lsJSON(LS_ACT, {}), fn)));
    var sc = lsJSON(LS_SCEN, {});
    Object.keys(sc).forEach(function (n) { var s = sc[n]; s.inputs = mapObj(s.inputs, fn); s.coefficients = mapObj(s.coefficients, fn); s.overrides = mapObj(s.overrides, fn); });
    lsSet(LS_SCEN, JSON.stringify(sc));
  }
  function insertFn(step) { return function (bs, c) { var X = step[bs]; if (!X) return c; var n = 0; X.forEach(function (x) { if (x <= c) n++; }); return c + n; }; }
  function removeFn(step) {
    return function (bs, c) {
      var X = step[bs]; if (!X) return c;
      var newCols = X.map(function (x, i) { return x + i; });
      if (newCols.indexOf(c) >= 0) return null;
      var n = 0; newCols.forEach(function (x) { if (x < c) n++; }); return c - n;
    };
  }

  // ------------------------------------------------------------------ public API
  var Store = {
    info: null,
    structure: function () { var s = lsJSON(LS_STRUCT, {}); return { extraYears: s.extraYears || 0, lastActual: s.lastActual || BASE_LAST_ACTUAL, modes: s.modes || [] }; },
    saveStructure: function (s) { lsSet(LS_STRUCT, JSON.stringify(s)); },
    actuals: function () { return lsJSON(LS_ACT, {}); },
    saveActuals: function (a) { return lsSet(LS_ACT, JSON.stringify(a)); },
    getCustom: function () { return idbGet('active').catch(function () { return null; }); },
    putCustom: function (rec) { return idbPut('active', rec); },
    clearCustom: function () { return idbDel('active'); },
    makeTpl: function (codes) {
      // eslint-disable-next-line no-new-func
      return new Function('V', 'K', 'G', 'X', 'return [' + codes.map(function (c) { return 'function(r,k){return ' + c + ';}'; }).join(',\n') + '];');
    },
    colName: colName, colNum: colNum,
    yearRowsOf: yearRowsOf, colYearIn: colYearIn,
    // add one forecast year (persisted as structure; stored edits migrated to new addresses)
    addYear: function (mode) {
      var st = this.structure(), base = this.base, modes = st.modes.slice(0, st.extraYears).concat([mode || 'trend']);
      var res = extend(base.core, base.parts, st.extraYears + 1, base.end, modes);
      migrateAll(insertFn(res.steps[res.steps.length - 1].ins));
      st.extraYears += 1; st.modes = modes; this.saveStructure(st);
      return res.end;
    },
    removeYear: function () {
      var st = this.structure(); if (!st.extraYears) return false;
      var base = this.base, res = extend(base.core, base.parts, st.extraYears, base.end, st.modes);
      migrateAll(removeFn(res.steps[res.steps.length - 1].ins));
      st.extraYears -= 1; st.modes = st.modes.slice(0, st.extraYears); if (st.lastActual >= res.end) st.lastActual = res.end - 2; this.saveStructure(st);
      return true;
    },
    previewAddYear: function (mode) {
      var st = this.structure(), base = this.base, modes = st.modes.slice(0, st.extraYears).concat([mode || 'trend']);
      var t0 = Date.now(), res = extend(base.core, base.parts, st.extraYears + 1, base.end, modes), last = res.steps[res.steps.length - 1];
      return { year: res.end, sheets: Object.keys(last.ins).length, skipped: last.skipped, cells: res.core.ncells - this.current.core.ncells, ms: Date.now() - t0 };
    },
    // boot: pick model, apply structure and actuals, build MakroModel, set baseline
    boot: function (cb) {
      var self = this;
      this.getCustom().then(function (rec) {
        var core = root.MODEL_CORE, parts = root.MODEL_PARTS, tplFactory = root.MODEL_TPL, custom = null, notes = [];
        if (rec && rec.core && rec.parts && rec.core.tplCode) {
          try { tplFactory = self.makeTpl(rec.core.tplCode); core = rec.core; parts = rec.parts; custom = rec.info || {}; if (rec.meta && root.MODEL_META) { for (var k in rec.meta) root.MODEL_META[k] = rec.meta[k]; } }
          catch (e) { notes.push('Yüklənmiş model açılmadı: ' + e.message); core = root.MODEL_CORE; parts = root.MODEL_PARTS; tplFactory = root.MODEL_TPL; custom = null; }
        }
        var baseEnd = (custom && custom.end) || BASE_END;
        self.base = { core: core, parts: parts, end: baseEnd };
        var st = self.structure(), end = baseEnd;
        if (st.extraYears > 0) {
          try { var res = extend(core, parts, st.extraYears, baseEnd, st.modes); core = res.core; parts = res.parts; end = res.end; }
          catch (e) { notes.push('Əlavə il tətbiq edilmədi: ' + e.message); st.extraYears = 0; }
        }
        self.current = { core: core, parts: parts, end: end };
        root.MODEL_CORE = core;
        var M = new root.MakroModel(core, parts, tplFactory);
        M.VX = M.V0.slice();
        var act = self.actuals(), nAct = 0, bad = 0;
        var sidx = {}; core.sheets.forEach(function (s, i) { sidx[core.books[s.b].n + '|' + s.n] = i; });
        Object.keys(act).forEach(function (k) {
          var p = k.split('|'), gs = sidx[p[0] + '|' + p[1]], m = /^([A-Z]+)(\d+)$/.exec(p[2] || '');
          if (gs === undefined || !m) { bad++; return; }
          var id = M.smap[gs].get(+m[2] * 20000 + colNum(m[1]));
          if (id === undefined) { bad++; return; }
          if (M.ft[id] >= 0) M.fixed.set(id, act[k]); else M.V[id] = act[k];
          nAct++;
        });
        var rebased = st.extraYears > 0 || nAct > 0;
        if (rebased) { M.recalcAll(); M.V0 = M.V.slice(); }
        var lastActual = Math.min(st.lastActual, end - 1);
        self.info = { custom: custom, extraYears: st.extraYears, endYear: end, lastActual: lastActual, actuals: nAct, badActuals: bad, rebased: rebased || !!custom,
          baseLabel: (rebased || custom) ? 'Baza model' : 'Excel bazası', baseShort: (rebased || custom) ? 'Bazada' : 'Excel-də', notes: notes };
        cb(M, self.info);
      });
    }
  };
  root.MakroStore = Store;
})(typeof window !== 'undefined' ? window : globalThis);
