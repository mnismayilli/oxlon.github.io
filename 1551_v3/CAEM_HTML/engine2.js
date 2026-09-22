/* Makro model — engine extension for the CAEM workbook: array values (broadcasting), references as values, lookup
   functions (INDEX/MATCH/VLOOKUP), error-catching functions, statistics (STDEV, SLOPE, INTERCEPT, LINEST), matrix algebra
   (TRANSPOSE, MMULT, MINVERSE) and the workbook's VBA user-defined functions HPF and HPP (Hodrick-Prescott filters). */
(function (root) {
  'use strict';
  var X = root.X, XErr = X.XErr, E = X.E;
  var V = null, G = null, M = null;
  var bind0 = X.bind;
  X.bind = function (values, g) { V = values; G = g; bind0(values, g); };
  X.bindModel = function (model) { M = model; };
  X.CUR = -1;
  function isErr(v) { return v instanceof XErr; }
  function isNum(v) { return typeof v === 'number'; }
  function Arr(r, c, d) { this.isArr = true; this.r = r; this.c = c; this.d = d; }
  X.Arr = Arr;

  // ---------------------------------------------------------------- conversions
  function rangeToArr(rg) { var d = new Array(rg.ids.length); for (var i = 0; i < d.length; i++) d[i] = rg.ids[i] < 0 ? null : V[rg.ids[i]]; return new Arr(rg.rows, rg.cols, d); }
  function refToRange(ref) {
    var m = M.smap[ref.gs], ids = [];
    for (var r = ref.r1; r <= ref.r2; r++) for (var c = ref.c1; c <= ref.c2; c++) { var x = m.get(r * 20000 + c); ids.push(x === undefined ? -1 : x); }
    return { isRange: true, ids: ids, rows: ref.r2 - ref.r1 + 1, cols: ref.c2 - ref.c1 + 1, gs: ref.gs, r1: ref.r1, c1: ref.c1 };
  }
  function isRefOne(v) { return v && typeof v === 'object' && v.id !== undefined && !v.isRange && !v.isArr && !v.isRef; }
  function sc(v) {                               // scalar argument: dereference a single-cell reference / 1x1 array
    if (isRefOne(v)) v = V[v.id];
    else if (v && (v.isArr || v.isRange || v.isRef)) { var A0 = toArr(v); v = A0.d[0]; }
    if (v instanceof XErr) throw v; return v;
  }
  X.sc = sc;
  function toArr(v) {
    if (v && v.isArr) return v;
    if (isRefOne(v)) return new Arr(1, 1, [V[v.id]]);
    if (v && v.isRange) return rangeToArr(v);
    if (v && v.isRef) return rangeToArr(refToRange(v));
    return null;
  }
  function asArr(v) { var a = toArr(v); return a || new Arr(1, 1, [v]); }
  X.toArr = toArr;
  function rangeRect(rg) {                       // static range object → rectangle (needs the range table)
    if (rg.gs !== undefined) return { gs: rg.gs, r1: rg.r1, c1: rg.c1, r2: rg.r1 + rg.rows - 1, c2: rg.c1 + rg.cols - 1 };
    return null;
  }

  // ---------------------------------------------------------------- element-wise (array mode) operators
  var SC = { add: X.add, sub: X.sub, mul: X.mul, div: X.div, pow: X.pow, cat: X.cat, eq: X.eq, ne: X.ne, lt: X.lt, gt: X.gt, le: X.le, ge: X.ge };
  function safe(fn, a, b) { try { return fn(a, b); } catch (e) { if (e instanceof XErr) return e; throw e; } }
  function bcast(a, b, fn) {
    var A = toArr(a), B = toArr(b);
    if (!A && !B) return fn(a, b);
    A = A || new Arr(1, 1, [a]); B = B || new Arr(1, 1, [b]);
    var r = Math.max(A.r, B.r), c = Math.max(A.c, B.c), d = new Array(r * c);
    for (var i = 0; i < r; i++) for (var j = 0; j < c; j++) {
      var ai = A.r === 1 ? 0 : i, aj = A.c === 1 ? 0 : j, bi = B.r === 1 ? 0 : i, bj = B.c === 1 ? 0 : j;
      if (ai >= A.r || aj >= A.c || bi >= B.r || bj >= B.c) { d[i * c + j] = E('#N/A'); continue; }
      d[i * c + j] = safe(fn, A.d[ai * A.c + aj], B.d[bi * B.c + bj]);
    }
    return new Arr(r, c, d);
  }
  function amap(a, fn) {
    var A = toArr(a); if (!A) return fn(a);
    var d = new Array(A.d.length); for (var i = 0; i < d.length; i++) d[i] = safe(fn, A.d[i]);
    return new Arr(A.r, A.c, d);
  }
  X.A = {};
  Object.keys(SC).forEach(function (k) { X.A[k] = function (a, b) { return bcast(a, b, SC[k]); }; });
  X.A.neg = function (a) { return amap(a, X.neg); };
  X.A.pct = function (a) { return amap(a, X.pct); };
  X.A.map = function (fn, a) { return amap(a, fn); };
  X.A.map2 = function (fn, a, b) { return bcast(a, b, fn); };
  X.A.IF = function (c, fa, fb) {
    var C = toArr(c);
    if (!C) { return X.b(c) ? fa() : fb(); }
    var a, b; try { a = fa(); } catch (e) { if (!(e instanceof XErr)) throw e; a = e; } try { b = fb(); } catch (e2) { if (!(e2 instanceof XErr)) throw e2; b = e2; }
    var t = bcast(C, 0, function (x) { return x; });
    return bcast(bcast(t, a, function (cv, av) { return [cv, av]; }), b, function (p, bv) {
      if (isErr(p[0])) return p[0]; var tb; try { tb = X.b(p[0]); } catch (e3) { return e3; } return tb ? p[1] : bv;
    });
  };
  X.A1 = function (v) {           // single-cell array formula: top-left element
    var A = toArr(v); if (!A) return v; var x = A.d[0]; if (x instanceof XErr) throw x; return x === undefined ? null : x;
  };
  X.GROUPV = function (v) { var A = toArr(v); if (A) { var o = new Arr(A.r, A.c, A.d); o.isGroup = true; return o; } var s1 = new Arr(1, 1, [v]); s1.isGroup = true; return s1; };
  X.AEL = function (vid) {        // member of a multi-cell array formula: vid = hidden group cell
    var g = V[vid]; if (g instanceof XErr) throw g;
    var ag = M.core.agroups[M.r[vid] - 1], A = toArr(g), row = M.r[X.CUR] - ag[1], col = M.c[X.CUR] - ag[2];
    if (!A) return g;
    if (A.r === 1 && A.c === 1) { row = 0; col = 0; }
    else if (A.r === 1) row = 0; else if (A.c === 1) col = 0;
    if (row >= A.r || col >= A.c) throw E('#N/A');
    var v = A.d[row * A.c + col];
    if (v instanceof XErr) throw v;
    return v === undefined || v === null ? 0 : v;
  };
  X.GROUP = function (v, ar, ac) { var A = toArr(v) || new Arr(1, 1, [v]); var o = new Arr(A.r, A.c, A.d); o.ar = ar; o.ac = ac; return o; };

  // ---------------------------------------------------------------- references
  X.REFOF = function (id) { return { isRef: true, gs: M.sh[id], r1: M.r[id], c1: M.c[id], r2: M.r[id], c2: M.c[id] }; };
  X.RANGEREF = function (rg) { var R = M.core.ranges[rg.gi]; return { isRef: true, gs: R[0], r1: R[1], c1: R[2], r2: R[3], c2: R[4] }; };
  function refOf(v) {
    if (v instanceof XErr) throw v;
    if (v && v.isRef) return v;
    if (v && v.isRange && v.gi !== undefined) return X.RANGEREF(v);
    throw E('#VALUE!');
  }
  X.RNG = function (a, b) {
    var A = refOf(a), B = refOf(b); if (A.gs !== B.gs) throw E('#VALUE!');
    return { isRef: true, gs: A.gs, r1: Math.min(A.r1, B.r1), c1: Math.min(A.c1, B.c1), r2: Math.max(A.r2, B.r2), c2: Math.max(A.c2, B.c2) };
  };
  function dims(v) { if (v && v.isRange) return [v.rows, v.cols]; if (v && v.isRef) return [v.r2 - v.r1 + 1, v.c2 - v.c1 + 1]; if (v && v.isArr) return [v.r, v.c]; return [1, 1]; }
  function idxArgs(src, r, c, argc) {
    var dm = dims(src), R = r === null || r === undefined ? 0 : Math.trunc(X.n(r)), C = c === null || c === undefined ? 0 : Math.trunc(X.n(c));
    if (argc < 3 && (dm[0] === 1 || dm[1] === 1)) { if (dm[0] === 1) { C = R; R = 1; } else C = 1; }  // 1-D: single index
    if (R < 0 || C < 0 || R > dm[0] || C > dm[1]) throw E('#REF!');
    return [R, C, dm];
  }
  X.INDEX = function (src, r, c, argc) {
    if (src instanceof XErr) throw src;
    var a = idxArgs(src, r, c, argc), R = a[0], C = a[1], dm = a[2], A = asArr(src);
    if (!R && dm[0] === 1) R = 1; if (!C && dm[1] === 1) C = 1;       // a single row/column collapses to one cell
    if (R && C) return A.d[(R - 1) * A.c + (C - 1)];
    if (!R && !C) return A;
    var d = [];
    if (!R) { for (var i = 0; i < dm[0]; i++) d.push(A.d[i * A.c + (C - 1)]); return new Arr(dm[0], 1, d); }
    for (var j = 0; j < dm[1]; j++) d.push(A.d[(R - 1) * A.c + j]); return new Arr(1, dm[1], d);
  };
  X.INDEXREF = function (src, r, c, argc) {
    var rf = refOf(src), a = idxArgs(src, r, c, argc), R = a[0], C = a[1];
    return { isRef: true, gs: rf.gs, r1: R ? rf.r1 + R - 1 : rf.r1, r2: R ? rf.r1 + R - 1 : rf.r2, c1: C ? rf.c1 + C - 1 : rf.c1, c2: C ? rf.c1 + C - 1 : rf.c2 };
  };
  function lower(v) { return typeof v === 'string' ? v.toLowerCase() : v; }
  function wildRe(s) { return new RegExp('^' + s.toLowerCase().replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$'); }
  function cmpType(a, b) { var ta = typeof a === 'number' ? 0 : typeof a === 'string' ? 1 : 2, tb = typeof b === 'number' ? 0 : typeof b === 'string' ? 1 : 2; return ta - tb; }
  function vecOf(src) { var A = asArr(src); if (A.r !== 1 && A.c !== 1) throw E('#N/A'); return A.d; }
  function matchExact(v, list) {
    if (v instanceof XErr) throw v; if (v === null) v = 0;
    var re = typeof v === 'string' && /[*?~]/.test(v) ? wildRe(v) : null, lv = lower(v);
    for (var i = 0; i < list.length; i++) {
      var x = list[i]; if (x === null || x === undefined) continue;
      if (re) { if (typeof x === 'string' && re.test(x.toLowerCase())) return i; }
      else if (typeof x === typeof v && lower(x) === lv) return i;
    }
    return -1;
  }
  function matchApprox(v, list, desc) {   // Excel binary-search semantics (largest value <= v; or smallest >= v when desc)
    if (v instanceof XErr) throw v;
    var lo = 0, hi = list.length - 1, best = -1;
    while (lo <= hi) {
      var mid = (lo + hi) >> 1, x = list[mid];
      if (x === null || x === undefined || cmpType(x, v) !== 0) {
        var k = mid; while (k <= hi && (list[k] === null || list[k] === undefined || cmpType(list[k], v) !== 0)) k++;
        if (k > hi) { hi = mid - 1; continue; }
        mid = k; x = list[k];
      }
      var c = lower(x) < lower(v) ? -1 : lower(x) > lower(v) ? 1 : 0;
      if (desc) c = -c;
      if (c <= 0) { best = mid; lo = mid + 1; } else hi = mid - 1;
    }
    return best;
  }
  X.MATCH = function (v, src, type) {
    if (v instanceof XErr) throw v; if (src instanceof XErr) throw src;
    var t = type === undefined || type === null ? 1 : X.n(type), list = vecOf(src), i;
    i = t === 0 ? matchExact(v, list) : matchApprox(v, list, t < 0);
    if (i < 0) throw E('#N/A'); return i + 1;
  };
  X.VLOOKUP = function (v, src, col, approx) {
    if (v instanceof XErr) throw v; if (src instanceof XErr) throw src;
    var A = asArr(src), k = Math.trunc(X.n(col)), exact = approx !== undefined && approx !== null && !X.b(approx);
    if (k < 1 || k > A.c) throw E(k < 1 ? '#VALUE!' : '#REF!');
    var first = []; for (var i = 0; i < A.r; i++) first.push(A.d[i * A.c]);
    var j = exact ? matchExact(v, first) : matchApprox(v, first, false);
    if (j < 0) throw E('#N/A'); var r = A.d[j * A.c + k - 1]; return r === undefined ? null : r;
  };
  X.HLOOKUP = function (v, src, row, approx) {
    if (v instanceof XErr) throw v; if (src instanceof XErr) throw src;
    var A = asArr(src), k = Math.trunc(X.n(row)), exact = approx !== undefined && approx !== null && !X.b(approx);
    if (k < 1 || k > A.r) throw E(k < 1 ? '#VALUE!' : '#REF!');
    var first = A.d.slice(0, A.c), j = exact ? matchExact(v, first) : matchApprox(v, first, false);
    if (j < 0) throw E('#N/A'); var r = A.d[(k - 1) * A.c + j]; return r === undefined ? null : r;
  };
  X.CHOOSE = function (i) { var k = Math.trunc(X.n(i)); if (k < 1 || k >= arguments.length) throw E('#VALUE!'); return arguments[k](); };

  // ---------------------------------------------------------------- error handling / information (thunks)
  function tryv(f) { try { return f(); } catch (e) { if (e instanceof XErr) return e; throw e; } }
  X.IFERROR = function (fa, fb) { var a = tryv(fa); var A = toArr(a); if (A) { var b; return amap(A, function (x) { if (x instanceof XErr) { if (b === undefined) b = tryv(fb); if (b instanceof XErr) throw b; return b; } return x === null || x === undefined ? 0 : x; }); } if (a instanceof XErr) { var bb = fb(); return bb; } return a; };
  X.IFNA = function (fa, fb) { var a = tryv(fa); var A = toArr(a); if (A) { var b; return amap(A, function (x) { if (x instanceof XErr && x.e === '#N/A') { if (b === undefined) b = tryv(fb); if (b instanceof XErr) throw b; return b; } if (x instanceof XErr) throw x; return x === null || x === undefined ? 0 : x; }); } if (a instanceof XErr && a.e === '#N/A') return fb(); if (a instanceof XErr) throw a; return a; };
  function isFn(test) { return function (fa) { var a = tryv(fa); var A = toArr(a); if (A && !(A.r === 1 && A.c === 1)) return amap(A, function (x) { return test(x); }); if (A) a = A.d[0]; return test(a); }; }
  X.ISNUMBER = isFn(function (x) { return typeof x === 'number'; });
  X.ISERROR = isFn(function (x) { return x instanceof XErr; });
  X.ISNA = isFn(function (x) { return x instanceof XErr && x.e === '#N/A'; });
  X.ISERR = isFn(function (x) { return x instanceof XErr && x.e !== '#N/A'; });
  X.ISTEXT = isFn(function (x) { return typeof x === 'string'; });
  X.ISBLANK = isFn(function (x) { return x === null || x === undefined; });
  X.ISLOGICAL = isFn(function (x) { return typeof x === 'boolean'; });
  X.NA = function () { throw E('#N/A'); };
  X.OR = function () {
    var any = false, res = false;
    for (var i = 0; i < arguments.length; i++) {
      var a = arguments[i], A = a && (a.isRange || a.isArr || a.isRef) ? toArr(a) : null, list = A ? A.d : (a && a.id !== undefined && a.constructor && a.constructor.name === 'RefOne') ? [V[a.id]] : null;
      if (A || list) { (list || A.d).forEach(function (v) { if (v instanceof XErr) throw v; if (typeof v === 'number' || typeof v === 'boolean') { any = true; if (v) res = true; } }); }
      else { any = true; if (X.b(a)) res = true; }
    }
    if (!any) throw E('#VALUE!'); return res;
  };
  X.NOT = function (a) { return !X.b(a); };

  // ---------------------------------------------------------------- numeric collection over ranges/arrays/scalars
  function nums(args, opts) {
    var out = [];
    for (var q = 0; q < args.length; q++) { var z = args[q]; if (z !== null && z !== undefined && typeof z !== 'object') X.n(z); else if (z instanceof XErr) throw z; }
    for (var i = 0; i < args.length; i++) {
      var a = args[i];
      if (a && a.id !== undefined && !a.isRange && !a.isArr && !a.isRef) { var v = V[a.id]; if (v instanceof XErr) throw v; if (typeof v === 'number') out.push(v); continue; }  // RefOne
      var A = toArr(a);
      if (A) { for (var j = 0; j < A.d.length; j++) { var w = A.d[j]; if (w instanceof XErr) throw w; if (typeof w === 'number') out.push(w); } continue; }
      if (a === null || a === undefined) { out.push(0); continue; }
      out.push(X.n(a));
    }
    return out;
  }
  function wrapAgg(name, f) { X[name] = function () { return f(nums(arguments)); }; }
  wrapAgg('SUM', function (a) { var t = 0; for (var i = 0; i < a.length; i++) t += a[i]; return t; });
  wrapAgg('AVERAGE', function (a) { if (!a.length) throw E('#DIV/0!'); var t = 0; for (var i = 0; i < a.length; i++) t += a[i]; return t / a.length; });
  wrapAgg('MIN', function (a) { return a.length ? Math.min.apply(null, a) : 0; });
  wrapAgg('MAX', function (a) { return a.length ? Math.max.apply(null, a) : 0; });
  wrapAgg('COUNT', function (a) { return a.length; });
  wrapAgg('PRODUCT', function (a) { var t = 1; for (var i = 0; i < a.length; i++) t *= a[i]; return a.length ? t : 0; });
  function variance(a, samp) { var n = a.length; if (n < (samp ? 2 : 1)) throw E('#DIV/0!'); var m = 0, i; for (i = 0; i < n; i++) m += a[i]; m /= n; var s = 0; for (i = 0; i < n; i++) s += (a[i] - m) * (a[i] - m); return s / (samp ? n - 1 : n); }
  wrapAgg('STDEV', function (a) { return Math.sqrt(variance(a, true)); });
  wrapAgg('STDEV.S', function (a) { return Math.sqrt(variance(a, true)); });
  wrapAgg('STDEVP', function (a) { return Math.sqrt(variance(a, false)); });
  wrapAgg('STDEV.P', function (a) { return Math.sqrt(variance(a, false)); });
  wrapAgg('VAR', function (a) { return variance(a, true); });
  wrapAgg('VARP', function (a) { return variance(a, false); });
  wrapAgg('MEDIAN', function (a) { if (!a.length) throw E('#NUM!'); a = a.slice().sort(function (x, y) { return x - y; }); var n = a.length; return n % 2 ? a[(n - 1) / 2] : (a[n / 2 - 1] + a[n / 2]) / 2; });
  X.COUNTA = function () { var n = 0; for (var i = 0; i < arguments.length; i++) { var a = arguments[i], A = toArr(a); if (A) A.d.forEach(function (v) { if (v !== null && v !== undefined) n++; }); else if (a && a.id !== undefined) { if (V[a.id] !== null && V[a.id] !== undefined) n++; } else if (a !== null && a !== undefined) n++; } return n; };
  X.SUMPRODUCT = function () {
    var arrs = Array.prototype.map.call(arguments, asArr), r0 = arrs[0].r, c0 = arrs[0].c, tot = 0;
    for (var i = 1; i < arrs.length; i++) if (arrs[i].r !== r0 || arrs[i].c !== c0) throw E('#VALUE!');
    for (var k = 0; k < r0 * c0; k++) { var p = 1; for (var j = 0; j < arrs.length; j++) { var v = arrs[j].d[k]; if (v instanceof XErr) throw v; p *= typeof v === 'number' ? v : 0; } tot += p; }
    return tot;
  };
  function pairs(ya, xa) {
    var Y = asArr(ya).d, Xv = asArr(xa).d; if (Y.length !== Xv.length) throw E('#N/A');
    var ys = [], xs = [];
    for (var i = 0; i < Y.length; i++) { if (Y[i] instanceof XErr) throw Y[i]; if (Xv[i] instanceof XErr) throw Xv[i]; if (typeof Y[i] === 'number' && typeof Xv[i] === 'number') { ys.push(Y[i]); xs.push(Xv[i]); } }
    return [ys, xs];
  }
  function slopeInt(ya, xa) {
    var p = pairs(ya, xa), ys = p[0], xs = p[1], n = ys.length; if (n < 1) throw E('#DIV/0!');
    var mx = 0, my = 0, i; for (i = 0; i < n; i++) { mx += xs[i]; my += ys[i]; } mx /= n; my /= n;
    var sxy = 0, sxx = 0; for (i = 0; i < n; i++) { sxy += (xs[i] - mx) * (ys[i] - my); sxx += (xs[i] - mx) * (xs[i] - mx); }
    if (sxx === 0) throw E('#DIV/0!'); var b = sxy / sxx; return [b, my - b * mx];
  }
  X.SLOPE = function (y, x) { return slopeInt(y, x)[0]; };
  X.INTERCEPT = function (y, x) { return slopeInt(y, x)[1]; };
  X.CORREL = function (a, b) {
    var p = pairs(a, b), xs = p[0], ys = p[1], m = xs.length; if (m < 2) throw E('#DIV/0!');
    var mx = 0, my = 0, i; for (i = 0; i < m; i++) { mx += xs[i]; my += ys[i]; } mx /= m; my /= m;
    var sxy = 0, sxx = 0, syy = 0; for (i = 0; i < m; i++) { var dx = xs[i] - mx, dy = ys[i] - my; sxy += dx * dy; sxx += dx * dx; syy += dy * dy; }
    if (sxx === 0 || syy === 0) throw E('#DIV/0!'); return sxy / Math.sqrt(sxx * syy);
  };

  // ---------------------------------------------------------------- scalar math (element-wise in array mode via X.A.map)
  X.SQRT = function (a) { var x = X.n(a); if (x < 0) throw E('#NUM!'); return Math.sqrt(x); };
  X.ABS = function (a) { return Math.abs(X.n(a)); };
  X.INT = function (a) { return Math.floor(X.n(a)); };
  X.MOD = function (a, b) { var x = X.n(a), y = X.n(b); if (y === 0) throw E('#DIV/0!'); return x - y * Math.floor(x / y); };
  X.POWER = function (a, b) { return X.pow(a, b); };
  X.LOG10 = function (a) { var x = X.n(a); if (x <= 0) throw E('#NUM!'); return Math.log10(x); };
  X.LOG = function (a, b) { var x = X.n(a), base = b === undefined || b === null ? 10 : X.n(b); if (x <= 0 || base <= 0 || base === 1) throw E('#NUM!'); return Math.log(x) / Math.log(base); };
  X.SIGN = function (a) { var x = X.n(a); return x > 0 ? 1 : x < 0 ? -1 : 0; };
  X.ROUNDUP = function (a, d) { var x = X.n(a), k = Math.trunc(X.n(d)), p = Math.pow(10, k); return Math.sign(x) * Math.ceil(Math.abs(x) * p - 1e-9) / p; };
  X.ROUNDDOWN = function (a, d) { var x = X.n(a), k = Math.trunc(X.n(d)), p = Math.pow(10, k); return Math.sign(x) * Math.floor(Math.abs(x) * p + 1e-9) / p; };
  X.CONCATENATE = function () { var s = ''; for (var i = 0; i < arguments.length; i++) s += X.s(arguments[i]); return s; };
  X.TRUE = function () { return true; }; X.FALSE = function () { return false; };

  // ---------------------------------------------------------------- matrices
  function mat(v) { var A = asArr(v), m = []; for (var i = 0; i < A.r; i++) { var row = []; for (var j = 0; j < A.c; j++) { var x = A.d[i * A.c + j]; if (x instanceof XErr) throw x; if (typeof x !== 'number') throw E('#VALUE!'); row.push(x); } m.push(row); } return m; }
  function fromMat(m) { var r = m.length, c = m[0].length, d = []; for (var i = 0; i < r; i++) for (var j = 0; j < c; j++) d.push(m[i][j]); return new Arr(r, c, d); }
  X.TRANSPOSE = function (v) { var A = toArr(v); if (!A) return v; var d = new Array(A.d.length); for (var i = 0; i < A.r; i++) for (var j = 0; j < A.c; j++) d[j * A.r + i] = A.d[i * A.c + j]; return new Arr(A.c, A.r, d); };
  function mmul(a, b) { if (a[0].length !== b.length) throw E('#VALUE!'); var r = a.length, c = b[0].length, n = b.length, o = []; for (var i = 0; i < r; i++) { var row = new Array(c).fill(0); for (var k = 0; k < n; k++) { var x = a[i][k]; for (var j = 0; j < c; j++) row[j] += x * b[k][j]; } o.push(row); } return o; }
  function minv(a) {
    var n = a.length; if (!n || a[0].length !== n) throw E('#VALUE!');
    var m = a.map(function (row, i) { var r = row.slice(); for (var j = 0; j < n; j++) r.push(i === j ? 1 : 0); return r; });
    for (var c = 0; c < n; c++) {
      var p = c, best = Math.abs(m[c][c]); for (var r = c + 1; r < n; r++) if (Math.abs(m[r][c]) > best) { best = Math.abs(m[r][c]); p = r; }
      if (best === 0) throw E('#NUM!');
      if (p !== c) { var t = m[p]; m[p] = m[c]; m[c] = t; }
      var piv = m[c][c]; for (var j2 = 0; j2 < 2 * n; j2++) m[c][j2] /= piv;
      for (r = 0; r < n; r++) if (r !== c && m[r][c] !== 0) { var f = m[r][c]; for (j2 = 0; j2 < 2 * n; j2++) m[r][j2] -= f * m[c][j2]; }
    }
    return m.map(function (row) { return row.slice(n); });
  }
  X.MMULT = function (a, b) { return fromMat(mmul(mat(a), mat(b))); };
  X.MINVERSE = function (a) { return fromMat(minv(mat(a))); };
  X.MDETERM = function (a) { var m = mat(a).map(function (r) { return r.slice(); }), n = m.length, det = 1; for (var c = 0; c < n; c++) { var p = c; for (var r = c + 1; r < n; r++) if (Math.abs(m[r][c]) > Math.abs(m[p][c])) p = r; if (m[p][c] === 0) return 0; if (p !== c) { var t = m[p]; m[p] = m[c]; m[c] = t; det = -det; } det *= m[c][c]; for (r = c + 1; r < n; r++) { var f = m[r][c] / m[c][c]; for (var j = c; j < n; j++) m[r][j] -= f * m[c][j]; } } return det; };
  X.MUNIT = function (n) { var k = Math.trunc(X.n(sc(n))), m = []; for (var i = 0; i < k; i++) { var r = new Array(k).fill(0); r[i] = 1; m.push(r); } return fromMat(m); };

  // LINEST via Householder QR (as Excel)
  X.LINEST = function (ky, kx, cst, stats) {
    var Y = asArr(ky), useC = cst === undefined || cst === null ? true : X.b(sc(cst)), st = stats === undefined || stats === null ? false : X.b(sc(stats));
    var y = Y.d.map(function (v) { if (v instanceof XErr) throw v; if (typeof v !== 'number') throw E('#VALUE!'); return v; }), n = y.length, cols = [];
    if (kx === undefined || kx === null) { var one = []; for (var i = 1; i <= n; i++) one.push(i); cols.push(one); }
    else {
      var Xa = asArr(kx);
      if ((Y.c === 1 && Xa.r === n) || (Y.r === 1 && Xa.c === n && Y.c === n && Xa.r !== n)) {
        var byCol = Y.c === 1 || Y.r !== 1;
        if (Y.c === 1) { for (var j = 0; j < Xa.c; j++) { var col = []; for (i = 0; i < Xa.r; i++) col.push(Xa.d[i * Xa.c + j]); cols.push(col); } }
        else { for (i = 0; i < Xa.r; i++) cols.push(Xa.d.slice(i * Xa.c, (i + 1) * Xa.c)); }
      } else if (Xa.d.length === n) cols.push(Xa.d.slice());
      else if (Y.r === 1 && Xa.c === n) { for (i = 0; i < Xa.r; i++) cols.push(Xa.d.slice(i * Xa.c, (i + 1) * Xa.c)); }
      else throw E('#REF!');
      cols.forEach(function (c) { c.forEach(function (v, k) { if (v instanceof XErr) throw v; if (typeof v !== 'number') throw E('#VALUE!'); }); });
    }
    var k = cols.length, p = k + (useC ? 1 : 0);
    // design matrix columns: [x1..xk, 1]
    var A = []; for (i = 0; i < n; i++) { var row = []; for (j = 0; j < k; j++) row.push(cols[j][i]); if (useC) row.push(1); A.push(row); }
    var ymean = 0; for (i = 0; i < n; i++) ymean += y[i]; ymean /= n;
    // QR (Householder) on A, apply to y
    var R = A.map(function (r) { return r.slice(); }), qy = y.slice(), m = n;
    for (var c = 0; c < p; c++) {
      var norm = 0; for (i = c; i < m; i++) norm += R[i][c] * R[i][c]; norm = Math.sqrt(norm);
      if (norm === 0) continue;
      var alpha = R[c][c] > 0 ? -norm : norm, v = []; for (i = 0; i < m; i++) v.push(i < c ? 0 : R[i][c]); v[c] -= alpha;
      var vv = 0; for (i = c; i < m; i++) vv += v[i] * v[i]; if (vv === 0) continue;
      for (var jj = c; jj < p; jj++) { var s = 0; for (i = c; i < m; i++) s += v[i] * R[i][jj]; s = 2 * s / vv; for (i = c; i < m; i++) R[i][jj] -= s * v[i]; }
      var sy = 0; for (i = c; i < m; i++) sy += v[i] * qy[i]; sy = 2 * sy / vv; for (i = c; i < m; i++) qy[i] -= sy * v[i];
    }
    var beta = new Array(p).fill(0);
    for (c = p - 1; c >= 0; c--) { var t = qy[c]; for (jj = c + 1; jj < p; jj++) t -= R[c][jj] * beta[jj]; if (Math.abs(R[c][c]) < 1e-300) throw E('#NUM!'); beta[c] = t / R[c][c]; }
    var coefs = []; for (j = k - 1; j >= 0; j--) coefs.push(beta[j]); coefs.push(useC ? beta[k] : 0);
    if (!st) return new Arr(1, k + 1, coefs);
    var ssr = 0, sst = 0; for (i = 0; i < n; i++) { var fit = 0; for (j = 0; j < p; j++) fit += A[i][j] * beta[j]; ssr += (y[i] - fit) * (y[i] - fit); sst += useC ? (y[i] - ymean) * (y[i] - ymean) : y[i] * y[i]; }
    var df = n - p, ssreg = sst - ssr, r2 = sst === 0 ? 1 : ssreg / sst, sey = df > 0 ? Math.sqrt(ssr / df) : E('#NUM!');
    // (R'R)^-1 for standard errors
    var Rp = []; for (i = 0; i < p; i++) Rp.push(R[i].slice(0, p));
    var Rinv = []; for (i = 0; i < p; i++) Rinv.push(new Array(p).fill(0));
    for (j = 0; j < p; j++) for (i = j; i >= 0; i--) { var sm = i === j ? 1 : 0; for (var l = i + 1; l <= j; l++) sm -= Rp[i][l] * Rinv[l][j]; Rinv[i][j] = sm / Rp[i][i]; }
    var se = []; for (i = 0; i < p; i++) { var q = 0; for (j = 0; j < p; j++) q += Rinv[i][j] * Rinv[i][j]; se.push(df > 0 ? Math.sqrt(q * ssr / df) : E('#NUM!')); }
    var ses = []; for (j = k - 1; j >= 0; j--) ses.push(se[j]); ses.push(useC ? se[k] : E('#N/A'));
    var F = df > 0 && ssr > 0 ? (ssreg / (useC ? p - 1 : p)) / (ssr / df) : E('#NUM!'), NA = E('#N/A'), w = k + 1, out = coefs.concat(ses);
    var r3 = [r2, sey], r4 = [F, df], r5 = [ssreg, ssr];
    [r3, r4, r5].forEach(function (rr) { for (var z = 0; z < w; z++) out.push(z < 2 ? rr[z] : NA); });
    return new Arr(5, w, out);
  };

  // ---------------------------------------------------------------- VBA user-defined functions (ported line by line)
  X.HPF = function (data, lambda) {                 // Kurt Annen's pentadiagonal HP filter (module HP_filter)
    var D = asArr(data), lam = X.n(sc(lambda)), nobs = D.r, ns = D.c, out = [], i, k;
    for (i = 0; i < nobs; i++) { out.push([]); for (k = 0; k < ns; k++) { var v = D.d[i * ns + k]; if (v instanceof XErr) throw v; out[i].push(typeof v === 'number' ? v : (v === null || v === undefined || v === '' ? 0 : X.n(v))); } }
    if (nobs <= 3) return fromMat(out);
    for (k = 0; k < ns; k++) {
      var A = new Array(nobs + 1).fill(0), B = new Array(nobs + 1).fill(0), C = new Array(nobs + 1).fill(0), y = [0];
      for (i = 1; i <= nobs; i++) y.push(out[i - 1][k]);
      A[1] = 1 + lam; B[1] = -2 * lam; C[1] = lam;
      for (i = 2; i <= nobs - 1; i++) { A[i] = 6 * lam + 1; B[i] = -4 * lam; C[i] = lam; }
      A[2] = 5 * lam + 1; A[nobs] = 1 + lam; A[nobs - 1] = 5 * lam + 1; B[1] = -2 * lam; B[nobs - 1] = -2 * lam; B[nobs] = 0; C[nobs - 1] = 0; C[nobs] = 0;
      var H1 = 0, H2 = 0, H3 = 0, H4 = 0, H5 = 0, HH1 = 0, HH2 = 0, HH3 = 0, HH5 = 0, z, HB, HC;
      for (i = 1; i <= nobs; i++) {
        z = A[i] - H4 * H1 - HH5 * HH2; HB = B[i]; HH1 = H1; H1 = (HB - H4 * H2) / z; B[i] = H1; HC = C[i]; HH2 = H2; H2 = HC / z; C[i] = H2;
        A[i] = (y[i] - HH3 * HH5 - H3 * H4) / z; HH3 = H3; H3 = A[i]; H4 = HB - H5 * HH1; HH5 = H5; H5 = HC;
      }
      H2 = 0; H1 = A[nobs]; y[nobs] = H1;
      for (i = nobs; i >= 1; i--) { y[i] = A[i] - B[i] * H1 - C[i] * H2; H2 = H1; H1 = y[i]; }
      for (i = 1; i <= nobs; i++) out[i - 1][k] = y[i];
    }
    return fromMat(out);
  };
  X.HPP = function (data, prior, lambda) {          // HP filter with priors (module HP_priors)
    var D = asArr(data), P = asArr(prior), lam = X.n(sc(lambda)), n = D.r, i, j;
    if (P.r !== n) throw E('#VALUE!');
    var y = []; for (i = 0; i < n; i++) { var v = D.d[i * D.c]; if (v instanceof XErr) throw v; y.push([typeof v === 'number' ? v : (v === null || v === '' ? 0 : X.n(v))]); }
    var Cm = []; for (i = 0; i < n - 2; i++) { var row = new Array(n).fill(0); row[i] = 1; row[i + 1] = -2; row[i + 2] = 1; Cm.push(row); }
    var CtC = mmul(tr(Cm), Cm), Am = []; for (i = 0; i < n; i++) { Am.push([]); for (j = 0; j < n; j++) Am[i].push((i === j ? 1 : 0) + lam * CtC[i][j]); }
    Am = minv(Am);
    var yhp = mmul(Am, y), idx = [], tau = [];
    for (j = 0; j < n; j++) { var pv = P.d[j * P.c]; if (pv !== null && pv !== undefined && pv !== '') { if (pv instanceof XErr) throw pv; idx.push(j); tau.push([X.n(pv)]); } }
    if (!idx.length) return fromMat(yhp);
    var Bm = idx.map(function (jj) { var r = new Array(n).fill(0); r[jj] = 1; return r; });
    var ABt = mmul(Am, tr(Bm)), BABt = mmul(mmul(Bm, Am), tr(Bm)), inv = BABt.length === 1 ? [[1 / BABt[0][0]]] : minv(BABt);
    var By = mmul(Bm, yhp), diff = tau.map(function (t, q) { return [t[0] - By[q][0]]; });
    var corr = mmul(mmul(ABt, inv), diff);
    return fromMat(yhp.map(function (r, q) { return [r[0] + corr[q][0]]; }));
  };
  function tr(m) { var r = m.length, c = m[0].length, o = []; for (var j = 0; j < c; j++) { var row = []; for (var i = 0; i < r; i++) row.push(m[i][j]); o.push(row); } return o; }
})(typeof window !== 'undefined' ? window : globalThis);
