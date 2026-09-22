/* CAEM — model store: active model (bundled or uploaded), realised data (actuals) and roll-forward of the first forecast year.
   Shared by index.html and panel.html. Requires engine.js, engine2.js and model.js.

   CAEM layout: the download sheets (External sector, Real sector, Prices, …) have fixed calendar-year columns; the 'Data' sheet
   and every analysis/projection sheet derive their year headers from INPUT!B9 (first forecast year), so their hand-typed values
   are tied to column positions. Rolling forward one year = move those hand-typed values one column to the left (the last
   column keeps its value) and add 1 to INPUT!B9 — exactly what an analyst does in Excel. */
(function (root) {
  'use strict';
  var P = 'caemModel.';
  var LS_STRUCT = P + 'structure.v1', LS_ACT = P + 'actuals.v1', LS_EDITS = P + 'edits.v1', LS_SCEN = P + 'scenarios.v1';
  var FIRST_YEAR_CELL = { sheet: 'INPUT', r: 9, c: 2 };
  function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); return true; } catch (e) { return false; } }
  function lsJSON(k, d) { try { var s = lsGet(k); return s ? JSON.parse(s) : d; } catch (e) { return d; } }

  // ------------------------------------------------------------------ IndexedDB (uploaded model)
  var DB = 'caemModel', STORE = 'models';
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

  // ------------------------------------------------------------------ helpers
  function colName(n) { var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function yearOfValue(v) {
    if (isNum(v) && v === Math.floor(v) && v >= 1980 && v <= 2060) return v;
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

  // ------------------------------------------------------------------ layout analysis (base model, Excel values)
  // positional sheet = its year header cells are formulas (driven by INPUT!B9); shift runs = >=3 consecutive constant cells in
  // year columns of one row (single parameters such as a smoothing weight next to the years are never moved)
  function analyseLayout(M, core) {
    var byS = core.sheets.map(function () { return []; });
    for (var i = 0; i < M.N; i++) byS[M.sh[i]].push(i);
    var L = { sheets: [], runs: [], b9: -1 };
    core.sheets.forEach(function (s, gs) {
      if (s.n === FIRST_YEAR_CELL.sheet && core.books[s.b].kind === 'model') { var b = M.smap[gs].get(FIRST_YEAR_CELL.r * 20000 + FIRST_YEAR_CELL.c); if (b !== undefined) L.b9 = b; }
    });
    // every formula that depends (transitively) on the first forecast year
    var dep = new Uint8Array(M.N);
    if (L.b9 >= 0) { var D = M.dependents(), st = [L.b9]; while (st.length) { var x = st.pop(), ds = D[x]; if (ds) for (var j = 0; j < ds.length; j++) if (!dep[ds[j]]) { dep[ds[j]] = 1; st.push(ds[j]); } } }
    L.dep = dep;
    core.sheets.forEach(function (s, gs) {
      var yr = yearRowsOf(byS[gs], M.r, M.c, M.V0), hdr = {}, positional = false;
      yr.forEach(function (y) { hdr[y.r] = 1; Object.keys(y.d).forEach(function (c) { var id = M.smap[gs].get(y.r * 20000 + +c); if (id !== undefined && dep[id]) positional = true; }); });
      L.sheets[gs] = { yr: yr, positional: positional && core.books[s.b].kind === 'model' && s.n !== '__massivlər__' };
      if (!L.sheets[gs].positional) return;
      var rows = {};
      byS[gs].forEach(function (id) { if (!hdr[M.r[id]]) (rows[M.r[id]] || (rows[M.r[id]] = [])).push(id); });
      Object.keys(rows).forEach(function (r) {
        var ids = rows[r].sort(function (a, b) { return M.c[a] - M.c[b]; }), run = [];
        function flush() {
          var n = run.filter(function (x) { return isNum(M.V0[x]); }).length;
          if (run.length >= 3 && n >= 2) L.runs.push({ gs: gs, r: +r, ids: run.slice() });
          run = [];
        }
        ids.forEach(function (id) {
          var y = colYearIn(yr, M.r[id], M.c[id]);
          var ok = y !== null && M.ft[id] < 0 && (M.V0[id] === null || isNum(M.V0[id]));
          if (ok && run.length && M.c[id] === M.c[run[run.length - 1]] + 1) run.push(id);
          else { flush(); if (ok) run.push(id); }
        });
        flush();
      });
    });
    // vertical layouts (years running down a column, e.g. the trend sheets B1a–B3 and C2): hand-typed values next to the
    // year column move up one row per roll
    L.vruns = [];
    var inRun = new Uint8Array(M.N); L.runs.forEach(function (r) { r.ids.forEach(function (x) { inRun[x] = 1; }); });
    core.sheets.forEach(function (s, gs) {
      if (core.books[s.b].kind !== 'model' || s.n === '__massivlər__') return;
      var cols = {};
      byS[gs].forEach(function (id) { var y = yearOfValue(M.V0[id]); if (y && M.ft[id] >= 0 && dep[id]) (cols[M.c[id]] || (cols[M.c[id]] = {}))[M.r[id]] = y; });
      Object.keys(cols).forEach(function (c) {
        var d = cols[c], rs = Object.keys(d).map(Number).sort(function (a, b) { return a - b; }), run = 0;
        for (var i = 1; i < rs.length; i++) if (rs[i] === rs[i - 1] + 1 && d[rs[i]] === d[rs[i - 1]] + 1) run++;
        if (run < 3) return;
        L.sheets[gs].vertical = L.sheets[gs].vertical || [];
        L.sheets[gs].vertical.push({ c: +c, rows: d });
        var byCol = {};
        byS[gs].forEach(function (id) { if (d[M.r[id]] !== undefined && M.c[id] !== +c && !inRun[id]) (byCol[M.c[id]] || (byCol[M.c[id]] = [])).push(id); });
        Object.keys(byCol).forEach(function (cc) {
          var ids = byCol[cc].sort(function (a, b) { return M.r[a] - M.r[b]; }), cur = [];
          function flushV() { var n = cur.filter(function (x) { return isNum(M.V0[x]); }).length; if (cur.length >= 3 && n >= 2) L.vruns.push({ gs: gs, rows: d, ids: cur.slice() }); cur = []; }
          ids.forEach(function (id) {
            var ok = M.ft[id] < 0 && (M.V0[id] === null || isNum(M.V0[id]));
            if (ok && cur.length && M.r[id] === M.r[cur[cur.length - 1]] + 1) cur.push(id); else { flushV(); if (ok) cur.push(id); }
          });
          flushV();
        });
      });
    });
    return L;
  }
  // one roll: values move one year column to the left; the last year column keeps its value
  function rollOnce(M, L) {
    var moves = 0;
    L.runs.forEach(function (run) {
      var gs = run.gs, yr = L.sheets[gs].yr, nv = [];
      run.ids.forEach(function (id) {
        var c = M.c[id], y = colYearIn(yr, M.r[id], c), y1 = colYearIn(yr, M.r[id], c + 1);
        if (y1 === null || y1 !== y + 1) { nv.push(M.V[id]); return; }      // last column of the year block: keep
        var nx = M.smap[gs].get(M.r[id] * 20000 + c + 1), v = nx === undefined ? null : M.V[nx];
        // no value to the right (row ends earlier) or an error: the cell keeps its value (flat, as 'fill right' would)
        if (v === null || v === undefined || v === '' || v instanceof root.X.XErr || typeof v === 'string') v = M.V[id];
        nv.push(v);
      });
      run.ids.forEach(function (id, i) { if (M.V[id] !== nv[i]) moves++; M.V[id] = nv[i]; });
    });
    (L.vruns || []).forEach(function (run) {
      var nv = run.ids.map(function (id) {
        var r = M.r[id], y = run.rows[r];
        if (run.rows[r + 1] !== y + 1) return M.V[id];
        var nx = M.smap[run.gs].get((r + 1) * 20000 + M.c[id]), v = nx === undefined ? null : M.V[nx];
        if (v === null || v === undefined || v === '' || v instanceof root.X.XErr || typeof v === 'string') v = M.V[id];
        return v;
      });
      run.ids.forEach(function (id, i) { if (M.V[id] !== nv[i]) moves++; M.V[id] = nv[i]; });
    });
    if (L.b9 >= 0 && isNum(M.V[L.b9])) M.V[L.b9] = M.V[L.b9] + 1;
    return moves;
  }

  // ------------------------------------------------------------------ key migration of stored scenario edits after a roll
  function mapKey(key, fn) {
    var p = key.split('|'); if (p.length < 3) return key;
    var hash = '', a = p[2], hi = a.indexOf('#'); if (hi >= 0) { hash = a.slice(hi); a = a.slice(0, hi); }
    var m = /^([A-Z]+)(\d+)$/.exec(a); if (!m) return key;
    var nc = fn(p[1], colNum(m[1]), +m[2]); if (nc === null) return null;
    if (typeof nc === 'object') return p[0] + '|' + p[1] + '|' + colName(nc.c) + nc.r + hash;
    return p[0] + '|' + p[1] + '|' + colName(nc) + m[2] + hash;
  }
  function mapObj(o, fn) { var out = {}; Object.keys(o || {}).forEach(function (k) { var nk = mapKey(k, fn); if (nk !== null) out[nk] = o[k]; }); return out; }
  function migrateEdits(fn) {
    var e = lsJSON(LS_EDITS, null);
    if (e) lsSet(LS_EDITS, JSON.stringify({ inputs: mapObj(e.inputs, fn), K: mapObj(e.K, fn), ov: mapObj(e.ov, fn) }));
    var sc = lsJSON(LS_SCEN, {});
    Object.keys(sc).forEach(function (n) { var s = sc[n]; s.inputs = mapObj(s.inputs, fn); s.coefficients = mapObj(s.coefficients, fn); s.overrides = mapObj(s.overrides, fn); });
    lsSet(LS_SCEN, JSON.stringify(sc));
  }

  // ------------------------------------------------------------------ public API
  var Store = {
    info: null, layout: null,
    keys: { edits: LS_EDITS, scen: LS_SCEN, act: LS_ACT, struct: LS_STRUCT, admin: P + 'admin.v1', toast: P + 'toast', panel: P + 'panel.v1' },
    structure: function () { var s = lsJSON(LS_STRUCT, {}); return { rolls: s.rolls || 0 }; },
    saveStructure: function (s) { lsSet(LS_STRUCT, JSON.stringify(s)); },
    actuals: function () { return lsJSON(LS_ACT, {}); },
    saveActuals: function (a) { return lsSet(LS_ACT, JSON.stringify(a)); },
    getCustom: function () { return idbGet('active').catch(function () { return null; }); },
    putCustom: function (rec) { return idbPut('active', rec); },
    clearCustom: function () { return idbDel('active'); },
    clearUserState: function () { [LS_EDITS, LS_ACT, LS_STRUCT].forEach(function (k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } }); },
    makeTpl: function (codes) {
      // eslint-disable-next-line no-new-func
      return new Function('V', 'K', 'G', 'X', 'return [' + codes.map(function (c) { return 'function(r,k){return ' + c + ';}'; }).join(',\n') + '];');
    },
    colName: colName, colNum: colNum, yearRowsOf: yearRowsOf, colYearIn: colYearIn,
    isPositional: function (gs) { return !!(this.layout && this.layout.sheets[gs] && this.layout.sheets[gs].positional); },
    // year of a row next to a vertical year column (null when the cell is not in such a block)
    verticalYear: function (gs, r, c) { var v = this.layout && this.layout.sheets[gs] && this.layout.sheets[gs].vertical; if (!v) return null; for (var i = 0; i < v.length; i++) if (v[i].c !== c && v[i].rows[r] !== undefined) return v[i].rows[r]; return null; },
    // realised data are stored in the layout of the base model (before any roll); on positional sheets the current column of a
    // base column is c - rolls
    baseKey: function (M, id) {
      var core = root.MODEL_CORE, s = core.sheets[M.sh[id]], c = M.c[id] + (this.isPositional(M.sh[id]) ? this.info.rolls : 0);
      return core.books[s.b].n + '|' + s.n + '|' + colName(c) + M.r[id];
    },
    idOfBaseKey: function (M, key) {
      var core = root.MODEL_CORE, p = key.split('|'), m = /^([A-Z]+)(\d+)$/.exec(p[2] || ''); if (!m) return -1;
      var gs = -1; core.sheets.forEach(function (s, i) { if (core.books[s.b].n === p[0] && s.n === p[1]) gs = i; });
      if (gs < 0) return -1;
      var c = colNum(m[1]) - (this.isPositional(gs) ? this.info.rolls : 0), x = M.smap[gs].get(+m[2] * 20000 + c);
      return x === undefined ? -1 : x;
    },
    // roll the first forecast year forward (persisted); scenario edits on positional sheets move one column left
    roll: function (dir) {
      var st = this.structure(), L = this.layout, self = this, core = root.MODEL_CORE;
      var sidx = {}; core.sheets.forEach(function (s, i) { sidx[s.n] = i; });
      // scenario edits follow their year: one column left on positional sheets, one row up next to a vertical year column
      var vertAt = function (gs, c, r) { var v = L.sheets[gs].vertical; return !!v && v.some(function (h) { return h.c !== c && h.rows[r] !== undefined; }); };
      var mig = function (d) {
        return function (sn, c, r) {
          var gs = sidx[sn]; if (gs === undefined) return c;
          if (L.sheets[gs].positional && !vertAt(gs, c, r)) return c - d >= 1 ? c - d : null;
          if (vertAt(gs, c, r)) return r - d >= 1 ? { c: c, r: r - d } : null;
          return c;
        };
      };
      if (dir > 0) { st.rolls += 1; migrateEdits(mig(1)); }
      else if (st.rolls > 0) { st.rolls -= 1; migrateEdits(mig(-1)); }
      this.saveStructure(st);
      return self.info.firstYear + (dir > 0 ? 1 : -1);
    },
    // dry run on a fresh model: apply stored actuals (+ staged ones) and the requested number of rolls, report new errors
    trial: function (extraActuals, rolls, baseActuals) {
      var base = this.base, M = new root.MakroModel(base.core, base.parts, base.tpl), L = this.layout;
      var n = this._applyActuals(M, Object.assign({}, baseActuals || this.actuals(), extraActuals || {}));
      M.recalcAll();
      var before = M.V.slice(), moved = 0;
      for (var k = 0; k < rolls; k++) { moved += rollOnce(M, L); M.recalcAll(); }
      var newErr = [], X = root.X;
      for (var i = 0; i < M.order.length; i++) { var id = M.order[i]; if (M.V[id] instanceof X.XErr && !(before[id] instanceof X.XErr) && !(M.VX0[id] instanceof X.XErr)) newErr.push(id); }
      return { model: M, actuals: n, moved: moved, newErrors: newErr };
    },
    _applyActuals: function (M, act) {
      var core = base0core(this), n = 0, bad = 0, sidx = {};
      core.sheets.forEach(function (s, i) { sidx[core.books[s.b].n + '|' + s.n] = i; });
      M.VX0 = M.V0;
      Object.keys(act).forEach(function (k) {
        var p = k.split('|'), gs = sidx[p[0] + '|' + p[1]], m = /^([A-Z]+)(\d+)$/.exec(p[2] || '');
        if (gs === undefined || !m) { bad++; return; }
        var id = M.smap[gs].get(+m[2] * 20000 + colNum(m[1]));
        if (id === undefined) { bad++; return; }
        if (M.ft[id] >= 0) M.fixed.set(id, act[k]); else M.V[id] = act[k];
        n++;
      });
      this._bad = bad;
      return n;
    },
    // boot: pick model, apply actuals and rolls, build MakroModel, set baseline
    boot: function (cb) {
      var self = this;
      this.getCustom().then(function (rec) {
        var core = root.MODEL_CORE, parts = root.MODEL_PARTS, tplFactory = root.MODEL_TPL, custom = null, notes = [];
        if (rec && rec.core && rec.parts && rec.core.tplCode) {
          try { tplFactory = self.makeTpl(rec.core.tplCode); core = rec.core; parts = rec.parts; custom = rec.info || {}; if (rec.meta && root.MODEL_META) { for (var k in rec.meta) root.MODEL_META[k] = rec.meta[k]; } }
          catch (e) { notes.push('Yüklənmiş model açılmadı: ' + e.message); core = root.MODEL_CORE; parts = root.MODEL_PARTS; tplFactory = root.MODEL_TPL; custom = null; }
        }
        root.MODEL_CORE = core;
        self.base = { core: core, parts: parts, tpl: tplFactory };
        var M = new root.MakroModel(core, parts, tplFactory);
        self.layout = analyseLayout(M, core);
        var baseFirst = self.layout.b9 >= 0 && isNum(M.V0[self.layout.b9]) ? M.V0[self.layout.b9] : 2025;
        M.VX = M.V0.slice();
        var st = self.structure(), act = self.actuals();
        var nAct = self._applyActuals(M, act), bad = self._bad;
        var rebased = nAct > 0 || st.rolls > 0;
        if (rebased) {
          M.recalcAll();
          for (var r = 0; r < st.rolls; r++) { rollOnce(M, self.layout); M.recalcAll(); }
          M.V0 = M.V.slice();
        }
        var first = baseFirst + st.rolls;
        self.info = { custom: custom, rolls: st.rolls, baseFirst: baseFirst, firstYear: first, lastActual: first - 1, endYear: first + 4, actuals: nAct, badActuals: bad,
          rebased: rebased || !!custom, baseLabel: (rebased || custom) ? 'Baza model' : 'Excel bazası', baseShort: (rebased || custom) ? 'Bazada' : 'Excel-də', notes: notes };
        cb(M, self.info);
      });
    }
  };
  function base0core(S) { return S.base.core; }
  root.MakroStore = Store;
})(typeof window !== 'undefined' ? window : globalThis);
