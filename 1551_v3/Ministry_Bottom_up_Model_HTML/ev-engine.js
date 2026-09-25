/* Econometrics engine — EViews expression parser, annual-series evaluator and OLS with diagnostics.
   Runs in the browser (window.EvEngine) and in node. Data is {NAME: {year: value}}. */
(function (root) {
  'use strict';

  /* ------------------------------------------------------------------ parser */
  function lex(src) {
    var s = String(src), i = 0, out = [];
    while (i < s.length) {
      var ch = s[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"' || ch === "'") { var q = ch, j = i + 1; while (j < s.length && s[j] !== q) j++; out.push({ t: 'str', v: s.slice(i + 1, j) }); i = j + 1; continue; }
      if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(s[i + 1] || ''))) {
        var m = /^[0-9]*\.?[0-9]+(?:[eE][-+]?[0-9]+)?/.exec(s.slice(i));
        out.push({ t: 'num', v: +m[0] }); i += m[0].length; continue;
      }
      if (/[A-Za-z_@]/.test(ch)) { var k = /^@?[A-Za-z_][A-Za-z0-9_.]*/.exec(s.slice(i)); out.push({ t: 'id', v: k[0] }); i += k[0].length; continue; }
      if (ch === '>' || ch === '<' || ch === '=') {
        var two = s.substr(i, 2);
        if (two === '>=' || two === '<=' || two === '=<' || two === '=>' || two === '<>') { out.push({ t: 'op', v: two }); i += 2; continue; }
        out.push({ t: 'op', v: ch }); i++; continue;
      }
      if ('+-*/(),^'.indexOf(ch) >= 0) { out.push({ t: 'op', v: ch }); i++; continue; }
      i++;                                    // skip anything unexpected
    }
    return out;
  }
  var FUNCS = { DLOG: 1, LOG: 1, D: 1, EXP: 1, ABS: 1, DLOG_: 1 };
  function parse(src) {
    var tk = lex(src), p = 0;
    function peek() { return tk[p]; }
    function eat(v) { if (tk[p] && tk[p].v === v) { p++; return true; } return false; }
    function cmp() {
      var a = expr();
      var t = peek();
      if (t && t.t === 'op' && ['=', '>', '<', '>=', '<=', '=<', '=>'].indexOf(t.v) >= 0) {
        p++; var b = expr();
        var op = t.v === '=<' ? '<=' : t.v === '=>' ? '>=' : t.v;
        return { k: 'cmp', op: op, a: a, b: b };
      }
      return a;
    }
    function expr() {
      var a = term();
      for (; ;) { var t = peek(); if (t && t.t === 'op' && (t.v === '+' || t.v === '-')) { p++; a = { k: 'bin', op: t.v, a: a, b: term() }; } else return a; }
    }
    function term() {
      var a = unary();
      for (; ;) { var t = peek(); if (t && t.t === 'op' && (t.v === '*' || t.v === '/')) { p++; a = { k: 'bin', op: t.v, a: a, b: unary() }; } else return a; }
    }
    function unary() { var t = peek(); if (t && t.t === 'op' && t.v === '-') { p++; return { k: 'neg', a: unary() }; } if (t && t.t === 'op' && t.v === '+') { p++; return unary(); } return atom(); }
    function atom() {
      var t = peek();
      if (!t) return { k: 'num', v: 0 };
      if (t.t === 'num') { p++; return { k: 'num', v: t.v }; }
      if (t.t === 'str') { p++; return { k: 'num', v: 0 }; }
      if (t.t === 'op' && t.v === '(') { p++; var e = cmp(); eat(')'); return e; }
      if (t.t === 'id') {
        var name = t.v; p++;
        var up = name.toUpperCase();
        if (up === '@BEFORE' || up === '@AFTER' || up === '@DURING') {
          var arg = '';
          if (eat('(')) { var s2 = peek(); if (s2 && s2.t === 'str') { arg = s2.v; p++; } else if (s2 && s2.t === 'num') { arg = String(s2.v); p++; } eat(')'); }
          return { k: 'period', kind: up.slice(1), spec: arg };
        }
        if (up === '@SUM') {
          var inner = null, smp = '';
          if (eat('(')) { inner = expr(); if (eat(',')) { var s3 = peek(); if (s3 && s3.t === 'str') { smp = s3.v; p++; } } eat(')'); }
          return { k: 'sum', a: inner, spec: smp };
        }
        if (eat('(')) {
          // function call or lagged series
          var save = p, neg = eat('-');
          var n2 = peek();
          if (n2 && n2.t === 'num' && (function () { var q = tk[p + 1]; return q && q.v === ')'; })()) {
            p++; eat(')');
            return { k: 'var', name: up, lag: neg ? n2.v : -n2.v };
          }
          p = save;
          var e2 = cmp(); eat(')');
          if (FUNCS[up] || up === 'DLOG' || up === 'LOG' || up === 'D') return { k: 'fn', fn: up, a: e2 };
          return { k: 'fn', fn: up, a: e2 };
        }
        if (up === 'C') return { k: 'num', v: 1 };
        return { k: 'var', name: up, lag: 0 };
      }
      p++; return { k: 'num', v: 0 };
    }
    var out = cmp();
    return { ast: out, rest: p < tk.length };
  }

  /* --------------------------------------------------------------- evaluation */
  function yearsOf(spec) {
    // "2010 2012" | "2010_2012" | "2003+2009" | "2010"
    var s = String(spec).trim();
    if (s.indexOf('+') >= 0) return { set: s.split('+').map(function (x) { return +x.trim(); }) };
    var m = /^(\d{4})\s*[\s_\-]\s*(\d{4})$/.exec(s);
    if (m) return { lo: +m[1], hi: +m[2] };
    var one = /^(\d{4})$/.exec(s);
    if (one) return { lo: +one[1], hi: +one[1] };
    return { lo: -Infinity, hi: Infinity };
  }
  function Ctx(data, aliases) { this.d = data; this.al = aliases || {}; }
  Ctx.prototype.get = function (name, y) {
    var n = this.al[name] || name;
    var s = this.d[n];
    if (!s) return undefined;
    var v = s[y];
    return (v === undefined || v === null || !isFinite(v)) ? undefined : v;
  };
  function ev(node, y, ctx) {
    if (!node) return undefined;
    switch (node.k) {
      case 'num': return node.v;
      case 'var': {
        if (node.name === 'T') return y - (node.lag || 0);
        return ctx.get(node.name, y - (node.lag || 0));
      }
      case 'neg': { var a = ev(node.a, y, ctx); return a === undefined ? undefined : -a; }
      case 'bin': {
        var x = ev(node.a, y, ctx), z = ev(node.b, y, ctx);
        if (x === undefined || z === undefined) return undefined;
        if (node.op === '+') return x + z;
        if (node.op === '-') return x - z;
        if (node.op === '*') return x * z;
        if (node.op === '/') return z === 0 ? undefined : x / z;
        return undefined;
      }
      case 'cmp': {
        var l = ev(node.a, y, ctx), r = ev(node.b, y, ctx);
        if (l === undefined || r === undefined) return undefined;
        switch (node.op) {
          case '=': return l === r ? 1 : 0;
          case '>': return l > r ? 1 : 0;
          case '<': return l < r ? 1 : 0;
          case '>=': return l >= r ? 1 : 0;
          case '<=': return l <= r ? 1 : 0;
          case '<>': return l !== r ? 1 : 0;
        }
        return undefined;
      }
      case 'period': {
        var sp = yearsOf(node.spec);
        if (node.kind === 'BEFORE') return y < (sp.lo === -Infinity ? sp.hi : sp.lo) ? 1 : 0;
        if (node.kind === 'AFTER') return y >= (sp.hi === Infinity ? sp.lo : sp.lo) ? 1 : 0;
        if (sp.set) return sp.set.indexOf(y) >= 0 ? 1 : 0;
        return (y >= sp.lo && y <= sp.hi) ? 1 : 0;
      }
      case 'sum': {
        var sp2 = yearsOf(node.spec), tot = 0, lo = sp2.lo, hi = sp2.hi;
        if (sp2.set) { for (var i = 0; i < sp2.set.length; i++) { var vv = ev(node.a, sp2.set[i], ctx); if (vv !== undefined) tot += vv; } return tot; }
        if (!isFinite(lo) || !isFinite(hi)) { lo = 1990; hi = 2035; }
        for (var yy = lo; yy <= hi; yy++) { var v2 = ev(node.a, yy, ctx); if (v2 !== undefined) tot += v2; }
        return tot;
      }
      case 'fn': {
        var f = node.fn;
        if (f === 'DLOG') { var a1 = ev(node.a, y, ctx), a0 = ev(node.a, y - 1, ctx); if (a1 === undefined || a0 === undefined || a1 <= 0 || a0 <= 0) return undefined; return Math.log(a1) - Math.log(a0); }
        if (f === 'D') { var b1 = ev(node.a, y, ctx), b0 = ev(node.a, y - 1, ctx); if (b1 === undefined || b0 === undefined) return undefined; return b1 - b0; }
        if (f === 'LOG') { var c1 = ev(node.a, y, ctx); if (c1 === undefined || c1 <= 0) return undefined; return Math.log(c1); }
        if (f === 'EXP') { var e1 = ev(node.a, y, ctx); return e1 === undefined ? undefined : Math.exp(e1); }
        if (f === 'ABS') { var f1 = ev(node.a, y, ctx); return f1 === undefined ? undefined : Math.abs(f1); }
        var g = ev(node.a, y, ctx);
        return g;                                    // unknown function: pass through
      }
    }
    return undefined;
  }
  function series(expr, years, ctx) {
    var pr = typeof expr === 'string' ? parse(expr) : { ast: expr };
    return years.map(function (y) { var v = ev(pr.ast, y, ctx); return (v === undefined || !isFinite(v)) ? null : v; });
  }
  function varsIn(node, out) {
    out = out || [];
    if (!node || typeof node !== 'object') return out;
    if (node.k === 'var' && node.name !== 'T') { if (out.indexOf(node.name) < 0) out.push(node.name); }
    ['a', 'b'].forEach(function (k) { if (node[k]) varsIn(node[k], out); });
    return out;
  }

  /* ---------------------------------------------------------------------- OLS */
  function solve(A, b) {                               // Gauss-Jordan with partial pivoting; returns x and inverse
    var n = b.length, M = [], i, j, k;
    for (i = 0; i < n; i++) { M[i] = A[i].slice(); for (j = 0; j < n; j++) M[i].push(i === j ? 1 : 0); M[i].push(b[i]); }
    for (i = 0; i < n; i++) {
      var piv = i;
      for (k = i + 1; k < n; k++) if (Math.abs(M[k][i]) > Math.abs(M[piv][i])) piv = k;
      if (Math.abs(M[piv][i]) < 1e-12) return null;
      var t = M[i]; M[i] = M[piv]; M[piv] = t;
      var d = M[i][i];
      for (j = i; j < 2 * n + 1; j++) M[i][j] /= d;
      for (k = 0; k < n; k++) { if (k === i) continue; var f = M[k][i]; if (!f) continue; for (j = i; j < 2 * n + 1; j++) M[k][j] -= f * M[i][j]; }
    }
    var x = [], inv = [];
    for (i = 0; i < n; i++) { x.push(M[i][2 * n]); inv.push(M[i].slice(n, 2 * n)); }
    return { x: x, inv: inv };
  }
  // incomplete beta -> t and F p-values
  function betacf(a, b, x) {
    var MAXIT = 200, EPS = 3e-12, FPMIN = 1e-300;
    var qab = a + b, qap = a + 1, qam = a - 1, c = 1, d = 1 - qab * x / qap;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    d = 1 / d; var h = d;
    for (var m = 1; m <= MAXIT; m++) {
      var m2 = 2 * m, aa = m * (b - m) * x / ((qam + m2) * (a + m2));
      d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
      c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
      d = 1 / d; h *= d * c;
      aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
      d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
      c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
      d = 1 / d; var del = d * c; h *= del;
      if (Math.abs(del - 1) < EPS) break;
    }
    return h;
  }
  function gammaln(x) {
    var c = [76.18009172947146, -86.50532032941677, 24.01409824083091, -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
    var y = x, t = x + 5.5; t -= (x + 0.5) * Math.log(t);
    var s = 1.000000000190015;
    for (var j = 0; j < 6; j++) s += c[j] / ++y;
    return -t + Math.log(2.5066282746310005 * s / x);
  }
  function betai(a, b, x) {
    if (x <= 0) return 0; if (x >= 1) return 1;
    var bt = Math.exp(gammaln(a + b) - gammaln(a) - gammaln(b) + a * Math.log(x) + b * Math.log(1 - x));
    return x < (a + 1) / (a + b + 2) ? bt * betacf(a, b, x) / a : 1 - bt * betacf(b, a, 1 - x) / b;
  }
  function tProb(t, df) { if (!isFinite(t) || df <= 0) return NaN; return betai(df / 2, 0.5, df / (df + t * t)); }   // two-sided
  function fProb(f, d1, d2) { if (!isFinite(f) || f <= 0 || d1 <= 0 || d2 <= 0) return NaN; return betai(d2 / 2, d1 / 2, d2 / (d2 + d1 * f)); }

  function ols(y, X, names) {
    var n = y.length, k = X.length;                    // X: array of regressor columns
    if (!n || !k || n <= k) return { ok: false, why: 'müşahidə sayı azdır (n=' + n + ', k=' + k + ')' };
    var XtX = [], Xty = [], i, j, t;
    for (i = 0; i < k; i++) {
      XtX.push(new Array(k).fill(0)); var s = 0;
      for (t = 0; t < n; t++) s += X[i][t] * y[t];
      Xty.push(s);
    }
    for (i = 0; i < k; i++) for (j = 0; j < k; j++) { var s2 = 0; for (t = 0; t < n; t++) s2 += X[i][t] * X[j][t]; XtX[i][j] = s2; }
    var sol = solve(XtX, Xty);
    if (!sol) return { ok: false, why: 'reqressorlar xətti asılıdır (matris tərs çevrilmir)' };
    var b = sol.x, inv = sol.inv;
    var fit = [], res = [], ssr = 0, ybar = 0;
    for (t = 0; t < n; t++) ybar += y[t]; ybar /= n;
    var tss = 0;
    for (t = 0; t < n; t++) {
      var f2 = 0; for (i = 0; i < k; i++) f2 += b[i] * X[i][t];
      fit.push(f2); var e = y[t] - f2; res.push(e); ssr += e * e; tss += (y[t] - ybar) * (y[t] - ybar);
    }
    var df = n - k, s2e = ssr / df, ser = Math.sqrt(s2e);
    var se = [], tstat = [], pval = [];
    for (i = 0; i < k; i++) { var v = inv[i][i] * s2e; var sd = v > 0 ? Math.sqrt(v) : NaN; se.push(sd); var tt = b[i] / sd; tstat.push(tt); pval.push(tProb(tt, df)); }
    var dwn = 0; for (t = 1; t < n; t++) { var dd = res[t] - res[t - 1]; dwn += dd * dd; }
    var dw = ssr > 0 ? dwn / ssr : NaN;
    // Is a constant inside the column space? An explicit C column, or period dummies that partition the
    // sample (@BEFORE/@DURING/@AFTER), both mean the intercept is there — EViews then centres R2 and
    // drops one degree of freedom from the F test. R2 and F must use the SAME total sum of squares.
    var spansConst = false;
    (function () {
      var A = [], c2 = [];
      for (i = 0; i < k; i++) { A.push(new Array(k).fill(0)); var s3 = 0; for (t = 0; t < n; t++) s3 += X[i][t]; c2.push(s3); }
      for (i = 0; i < k; i++) for (j = 0; j < k; j++) { var s4 = 0; for (t = 0; t < n; t++) s4 += X[i][t] * X[j][t]; A[i][j] = s4; }
      var sc = solve(A, c2);
      if (!sc) return;
      var rss = 0;
      for (t = 0; t < n; t++) { var f3 = 0; for (i = 0; i < k; i++) f3 += sc.x[i] * X[i][t]; var e3 = 1 - f3; rss += e3 * e3; }
      spansConst = rss < 1e-9 * n;
    })();
    var hasConst = spansConst;
    var r2 = tss > 0 ? 1 - ssr / tss : NaN;                    // centred, as EViews reports it
    var adj = 1 - (1 - r2) * (n - (spansConst ? 1 : 0)) / df;
    var fdf1 = spansConst ? k - 1 : k;
    var fst = fdf1 > 0 && ssr > 0 ? ((tss - ssr) / fdf1) / s2e : NaN;
    return {
      ok: true, n: n, k: k, df: df, beta: b, se: se, t: tstat, p: pval, names: names || [],
      r2: r2, adjR2: adj, ser: ser, ssr: ssr, dw: dw, f: fst, fp: fProb(fst, fdf1, df),
      resid: res, fitted: fit, hasConst: hasConst
    };
  }

  root.EvEngine = { parse: parse, ev: ev, series: series, Ctx: Ctx, ols: ols, varsIn: varsIn, tProb: tProb, fProb: fProb, yearsOf: yearsOf };
  if (typeof module !== 'undefined') module.exports = root.EvEngine;
})(typeof window !== 'undefined' ? window : globalThis);
