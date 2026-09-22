/* Makro model — model assembly and recalculation (browser + node). Requires engine.js (X). */
(function (root) {
  'use strict';
  var X = root.X;

  function Model(core, parts, tplFactory) {
    var self = this;
    this.core = core;
    var N = core.ncells;
    this.N = N;
    this.V = new Array(N);            // current values
    this.V0 = new Array(N);           // Excel cached values (original)
    this.sh = new Int32Array(N); this.r = new Int32Array(N); this.c = new Int32Array(N);
    this.st = new Int32Array(N);
    this.ft = new Int32Array(N).fill(-1);  // template index
    this.fr = new Array(N);                // refs per formula cell
    this.fk = new Int32Array(N);           // literal base index into K
    this.K0 = core.K.slice(); this.K = core.K.slice();
    for (var p = 0; p < parts.length; p++) {
      var P = parts[p];
      for (var i = 0; i < P.ids.length; i++) {
        var id = P.ids[i];
        this.sh[id] = P.sh[i]; this.r[id] = P.r[i]; this.c[id] = P.c[i]; this.st[id] = P.st[i];
        var v = P.v[i];
        if (v && typeof v === 'object' && v.e) v = X.E(v.e);
        this.V0[id] = v === undefined ? null : v;
      }
      for (var k in P.f) {
        var f = P.f[k], cid = +k;
        this.ft[cid] = f[0]; this.fr[cid] = f[1]; this.fk[cid] = f[2];
      }
    }
    for (i = 0; i < N; i++) this.V[i] = this.V0[i];
    // sheet cell maps
    this.smap = core.sheets.map(function () { return new Map(); });
    for (i = 0; i < N; i++) this.smap[this.sh[i]].set(this.r[i] * 20000 + this.c[i], i);
    // ranges (lazy dense id arrays)
    var rcache = new Array(core.ranges.length);
    this.range = function (gi) {
      var o = rcache[gi];
      if (o) return o;
      var R = core.ranges[gi], m = self.smap[R[0]], ids = [];
      for (var rr = R[1]; rr <= R[3]; rr++) for (var cc = R[2]; cc <= R[4]; cc++) { var x = m.get(rr * 20000 + cc); ids.push(x === undefined ? -1 : x); }
      return (rcache[gi] = { isRange: true, ids: ids, rows: R[3] - R[1] + 1, cols: R[4] - R[2] + 1, gi: gi, gs: R[0], r1: R[1], c1: R[2] });
    };
    function G(ref) { return self.range(-ref - 1); }
    this.T = tplFactory(this.V, this.K, G, X);
    X.bind(this.V, G);
    if (X.bindModel) X.bindModel(this);
    this.order = core.order;
    this.pos = new Int32Array(N).fill(-1);
    for (i = 0; i < this.order.length; i++) this.pos[this.order[i]] = i;
    this._deps = null;
    this.cyc = core.cyc || [];            // [start, end) blocks of this.order solved by sweeps
    this.cblk = new Int32Array(this.order.length).fill(-1);
    for (i = 0; i < this.cyc.length; i++) for (var q = this.cyc[i][0]; q < this.cyc[i][1]; q++) this.cblk[q] = i;
    this.cycInfo = [];                    // per block: {sweeps, converged}
    this.maxSweeps = 500;
    this.fixed = new Map();             // realised (actual) values entered over formula cells
  }

  Model.prototype.evalCell = function (id) {
    var v;
    X.CUR = id;
    try {
      v = this.T[this.ft[id]](this.fr[id], this.fk[id]);
      if (v === null || v === undefined) v = 0;
      else if (v && v.isGroup) { /* hidden array-group cell keeps its array */ }
      else if (v && (v.isArr || v.isRef || v.isRange)) { v = X.A1 ? X.A1(v) : X.E('#VALUE!'); if (v === null || v === undefined) v = 0; }
      if (v === 0) v = 0; // normalise -0
    } catch (e) {
      if (e instanceof X.XErr) v = e; else throw e;
    }
    return v;
  };

  // precedents of a formula cell (cell ids, ranges expanded)
  Model.prototype.precedents = function (id) {
    var out = [], refs = this.fr[id] || [];
    for (var i = 0; i < refs.length; i++) {
      var r = refs[i];
      if (r >= 0) out.push(r);
      else { var ids = this.range(-r - 1).ids; for (var j = 0; j < ids.length; j++) if (ids[j] >= 0) out.push(ids[j]); }
    }
    return out;
  };
  Model.prototype.dependents = function () {
    if (this._deps) return this._deps;
    var D = new Array(this.N);
    for (var i = 0; i < this.order.length; i++) {
      var id = this.order[i], pr = this.precedents(id);
      for (var j = 0; j < pr.length; j++) { var p = pr[j]; (D[p] || (D[p] = [])).push(id); }
    }
    for (i = 0; i < this.N; i++) if (D[i]) D[i] = Array.from(new Set(D[i]));
    return (this._deps = D);
  };

  function sameV(a, b) {
    if (a === b) return true;
    if (typeof a === 'number' && typeof b === 'number') { var d = Math.abs(a - b); return d <= 1e-12 * Math.max(1, Math.abs(a), Math.abs(b)); }
    if (a instanceof X.XErr && b instanceof X.XErr) return a.e === b.e;
    if (a && b && a.isGroup && b.isGroup && a.d.length === b.d.length) { for (var i = 0; i < a.d.length; i++) if (!sameV(a.d[i], b.d[i])) return false; return true; }
    return false;
  }
  Model.prototype.valueOf1 = function (id, overrides) {
    return (overrides && overrides.has(id)) ? overrides.get(id) : this.fixed.has(id) ? this.fixed.get(id) : this.evalCell(id);
  };
  // a cyclic block: sweep until no value changes (exact when a switch breaks the loop; Gauss-Seidel otherwise)
  Model.prototype.solveBlock = function (b, overrides) {
    var s = this.cyc[b][0], e = this.cyc[b][1], o = this.order, n;
    for (n = 1; n <= this.maxSweeps; n++) {
      var changed = false;
      for (var q = s; q < e; q++) { var id = o[q], v = this.valueOf1(id, overrides); if (!changed && !sameV(v, this.V[id])) changed = true; this.V[id] = v; }
      if (!changed) break;
    }
    this.cycInfo[b] = { sweeps: Math.min(n, this.maxSweeps), converged: n <= this.maxSweeps };
  };
  // full recalculation of every formula in topological order
  Model.prototype.recalcAll = function (overrides) {
    for (var i = 0; i < this.order.length; i++) {
      var b = this.cblk[i];
      if (b >= 0) { this.solveBlock(b, overrides); i = this.cyc[b][1] - 1; continue; }
      var id = this.order[i];
      this.V[id] = this.valueOf1(id, overrides);
    }
  };
  // incremental: recompute only the downstream cone of the changed cells (formula cells may also be overridden)
  Model.prototype.recalcFrom = function (changed, overrides) {
    var D = this.dependents(), seen = new Uint8Array(this.N), stack = [], dirty = [];
    for (var i = 0; i < changed.length; i++) { stack.push(changed[i]); }
    while (stack.length) {
      var x = stack.pop(), ds = D[x];
      if (this.ft[x] >= 0 && !seen[x]) { seen[x] = 1; dirty.push(x); }
      if (!ds) continue;
      for (var j = 0; j < ds.length; j++) if (!seen[ds[j]]) { seen[ds[j]] = 1; dirty.push(ds[j]); stack.push(ds[j]); }
    }
    var pos = this.pos;
    dirty.sort(function (a, b) { return pos[a] - pos[b]; });
    var lastB = -1;
    for (i = 0; i < dirty.length; i++) {
      var id = dirty[i], b = this.cblk[pos[id]];
      if (b >= 0) { if (b !== lastB) { this.solveBlock(b, overrides); lastB = b; } continue; }
      this.V[id] = this.valueOf1(id, overrides);
    }
    return dirty.length;
  };

  root.MakroModel = Model;
  if (typeof module !== 'undefined') module.exports = Model;
})(typeof window !== 'undefined' ? window : globalThis);
