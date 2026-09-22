/* Makro model — Excel formula runtime (shared by browser and verification harness).
   Values: number | string | boolean | null (empty) | XErr (error object).                     */
(function (root) {
  'use strict';
  function XErr(code) { this.e = code; }
  XErr.prototype.toString = function () { return this.e; };
  var ERR = {};
  ['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#N/A', '#NUM!', '#NULL!'].forEach(function (c) { ERR[c] = new XErr(c); });
  function E(code) { return ERR[code] || (ERR[code] = new XErr(code)); }
  function isErr(v) { return v instanceof XErr; }
  function RefOne(id) { this.id = id; }          // single-cell reference passed to a range function

  var X = { E: E, isErr: isErr, XErr: XErr };
  var V = null, RG = null;                         // bound by engine
  X.bind = function (values, rangeResolver) { V = values; RG = rangeResolver; };

  function n(v) {
    if (typeof v === 'number') return v;
    if (v === null || v === undefined) return 0;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (v instanceof XErr) throw v;
    if (typeof v === 'string') {
      var s = v.trim();
      if (s === '') throw ERR['#VALUE!'];
      var pct = false;
      if (s.charAt(s.length - 1) === '%') { pct = true; s = s.slice(0, -1); }
      if (/^[+-]?\d*,\d+$/.test(s)) s = s.replace(',', '.');           // az-AZ decimal comma, e.g. "0,7"
      else s = s.replace(/,(?=\d{3}(\D|$))/g, '');                     // thousands separators
      if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) throw ERR['#VALUE!'];
      var x = parseFloat(s); return pct ? x / 100 : x;
    }
    throw ERR['#VALUE!'];
  }
  function fin(x) { if (!isFinite(x)) throw ERR['#NUM!']; return x; }
  function gtext(x) {                              // Excel "General" number -> text (15 significant digits)
    if (x === 0) return '0';
    var s = Number(x.toPrecision(15)).toString();
    if (s.indexOf('e') >= 0) {
      var m = Number(x.toPrecision(15)).toExponential().split('e');
      var ex = parseInt(m[1], 10);
      s = m[0] + 'E' + (ex < 0 ? '-' : '+') + (Math.abs(ex) < 10 ? '0' : '') + Math.abs(ex);
    }
    return s;
  }
  function s(v) {
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return gtext(v);
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (v instanceof XErr) throw v;
    return String(v);
  }
  X.n = n; X.s = s; X.gtext = gtext;
  X.b = function (v) {
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (v === null || v === undefined) return false;
    if (v instanceof XErr) throw v;
    var u = String(v).toUpperCase();
    if (u === 'TRUE') return true; if (u === 'FALSE') return false;
    throw ERR['#VALUE!'];
  };
  X.neg = function (a) { return -n(a); };
  X.pct = function (a) { return n(a) / 100; };
  X.add = function (a, b) { return fin(n(a) + n(b)); };
  X.sub = function (a, b) { return fin(n(a) - n(b)); };
  X.mul = function (a, b) { return fin(n(a) * n(b)); };
  X.div = function (a, b) { var x = n(a), y = n(b); if (y === 0) throw ERR['#DIV/0!']; return fin(x / y); };
  X.pow = function (a, b) {
    var x = n(a), y = n(b);
    if (x === 0 && y === 0) throw ERR['#NUM!'];
    if (x === 0 && y < 0) throw ERR['#DIV/0!'];
    var r = Math.pow(x, y);
    if (isNaN(r) && x < 0) { var q = 1 / y, qi = Math.round(q); if (Math.abs(q - qi) < 1e-9 * Math.max(1, Math.abs(q)) && qi % 2) r = -Math.pow(-x, y); }  // Excel: odd roots of negatives
    if (isNaN(r)) throw ERR['#NUM!']; return fin(r);
  };
  X.cat = function (a, b) { return s(a) + s(b); };
  function typeRank(v) { return typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2; }
  function cmp(a, b) {
    if (a instanceof XErr) throw a; if (b instanceof XErr) throw b;
    if (a === null || a === undefined) a = (typeof b === 'string') ? '' : (typeof b === 'boolean') ? false : 0;
    if (b === null || b === undefined) b = (typeof a === 'string') ? '' : (typeof a === 'boolean') ? false : 0;
    var ta = typeRank(a), tb = typeRank(b);
    if (ta !== tb) return ta < tb ? -1 : 1;
    if (ta === 1) { a = a.toLowerCase(); b = b.toLowerCase(); }
    if (ta === 2) { a = a ? 1 : 0; b = b ? 1 : 0; }
    return a < b ? -1 : a > b ? 1 : 0;
  }
  X.eq = function (a, b) { return cmp(a, b) === 0; };
  X.ne = function (a, b) { return cmp(a, b) !== 0; };
  X.lt = function (a, b) { return cmp(a, b) < 0; };
  X.gt = function (a, b) { return cmp(a, b) > 0; };
  X.le = function (a, b) { return cmp(a, b) <= 0; };
  X.ge = function (a, b) { return cmp(a, b) >= 0; };

  X.R1 = function (id) { return new RefOne(id); };

  // iterate numeric content of arguments with Excel reference semantics
  function collect(args, onNum, opts) {
    for (var q = 0; q < args.length; q++) {                 // pass 1: value arguments (errors surface first)
      var z = args[q];
      if (!(z instanceof RefOne) && !(z && z.isRange) && z !== null && z !== undefined) n(z);
    }
    for (var i = 0; i < args.length; i++) {
      var a = args[i];
      if (a instanceof RefOne) {
        var v = V[a.id];
        if (v instanceof XErr) throw v;
        if (typeof v === 'number') onNum(v);
      } else if (a && a.isRange) {
        var ids = a.ids;
        for (var j = 0; j < ids.length; j++) {
          if (ids[j] < 0) continue;
          var w = V[ids[j]];
          if (w instanceof XErr) throw w;
          if (typeof w === 'number') onNum(w);
        }
      } else {
        if (a === null || a === undefined) { onNum(0); continue; }   // missing argument
        onNum(n(a));
      }
    }
  }
  X.SUM = function () { var t = 0; collect(arguments, function (x) { t += x; }); return fin(t); };
  X.AVERAGE = function () {
    var t = 0, c = 0; collect(arguments, function (x) { t += x; c++; });
    if (c === 0) throw ERR['#DIV/0!']; return fin(t / c);
  };
  X.MIN = function () { var m = Infinity; collect(arguments, function (x) { if (x < m) m = x; }); return m === Infinity ? 0 : m; };
  X.MAX = function () { var m = -Infinity; collect(arguments, function (x) { if (x > m) m = x; }); return m === -Infinity ? 0 : m; };
  X.AND = function () {
    var any = false, res = true;
    for (var i = 0; i < arguments.length; i++) {
      var a = arguments[i];
      var list = (a instanceof RefOne) ? [V[a.id]] : (a && a.isRange) ? a.ids.map(function (k) { return k < 0 ? null : V[k]; }) : null;
      if (list) {
        for (var j = 0; j < list.length; j++) {
          var v = list[j];
          if (v instanceof XErr) throw v;
          if (typeof v === 'number') { any = true; if (v === 0) res = false; }
          else if (typeof v === 'boolean') { any = true; if (!v) res = false; }
        }
      } else { any = true; if (!X.b(a)) res = false; }
    }
    if (!any) throw ERR['#VALUE!'];
    return res;
  };
  X.SUMPRODUCT = function () {
    var arrs = Array.prototype.slice.call(arguments);
    var r0 = arrs[0].rows, c0 = arrs[0].cols;
    for (var i = 1; i < arrs.length; i++) if (arrs[i].rows !== r0 || arrs[i].cols !== c0) throw ERR['#VALUE!'];
    var len = arrs[0].ids.length, tot = 0;
    for (var k = 0; k < len; k++) {
      var p = 1;
      for (var i2 = 0; i2 < arrs.length; i2++) {
        var id = arrs[i2].ids[k];
        var v = id < 0 ? null : V[id];
        if (v instanceof XErr) throw v;
        p *= (typeof v === 'number') ? v : 0;
      }
      tot += p;
    }
    return fin(tot);
  };
  X.CORREL = function (a, b) {
    if (a.ids.length !== b.ids.length) throw ERR['#N/A'];
    var xs = [], ys = [];
    for (var k = 0; k < a.ids.length; k++) {
      var u = a.ids[k] < 0 ? null : V[a.ids[k]], w = b.ids[k] < 0 ? null : V[b.ids[k]];
      if (u instanceof XErr) throw u; if (w instanceof XErr) throw w;
      if (typeof u === 'number' && typeof w === 'number') { xs.push(u); ys.push(w); }
    }
    var m = xs.length; if (m < 2) throw ERR['#DIV/0!'];
    var mx = 0, my = 0; for (var i = 0; i < m; i++) { mx += xs[i]; my += ys[i]; } mx /= m; my /= m;
    var sxy = 0, sxx = 0, syy = 0;
    for (i = 0; i < m; i++) { var dx = xs[i] - mx, dy = ys[i] - my; sxy += dx * dy; sxx += dx * dx; syy += dy * dy; }
    if (sxx === 0 || syy === 0) throw ERR['#DIV/0!'];
    return sxy / Math.sqrt(sxx * syy);
  };
  X.LN = function (a) { var x = n(a); if (x <= 0) throw ERR['#NUM!']; return Math.log(x); };
  X.EXP = function (a) { return fin(Math.exp(n(a))); };
  X.ROUND = function (a, d) {
    var x = n(a), k = Math.trunc(n(d));
    var sign = x < 0 ? -1 : 1, ax = Math.abs(x);
    var r = Number(Math.round(Number(ax + 'e' + k)) + 'e' + (-k));
    if (!isFinite(r)) { var p = Math.pow(10, k); r = Math.round(ax * p) / p; }
    return sign * r;
  };
  X.TEXT = function (a, f) {
    var fmt = s(f);
    if (typeof a === 'string') { var t = a.trim(); if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) return a; }
    return X.fmt(n(a), fmt);
  };

  // --------------------------------------------------------------- number formatting (display + TEXT)
  var fmtCache = {};
  function parseSection(sec) {
    var o = { pct: 0, dec: 0, decOpt: 0, thou: false, lead: '', trail: '', sci: false, intZero: 0, hasDot: false, scale: 0, text: false, red: false };
    var s2 = sec.replace(/\[(Red|Черный|Красный|Black|Blue|Green|Color\d+)\]/gi, function (m) { if (/red|крас/i.test(m)) o.red = true; return ''; })
                .replace(/\[[^\]]*\]/g, '');
    if (s2.indexOf('@') >= 0 && !/[0#?]/.test(s2)) { o.text = true; o.lead = s2.replace(/"/g, '').replace('@', ''); return o; }
    // literal strings / escapes -> collect lead & trail around numeric part
    var i = 0, part = 0, core = '';
    while (i < s2.length) {
      var ch = s2.charAt(i);
      if (ch === '"') { var j = s2.indexOf('"', i + 1); if (j < 0) j = s2.length; var lit = s2.slice(i + 1, j); if (part === 0) o.lead += lit; else o.trail += lit; i = j + 1; continue; }
      if (ch === '\\') { if (part === 0) o.lead += s2.charAt(i + 1); else o.trail += s2.charAt(i + 1); i += 2; continue; }
      if (ch === '_' ) { if (part === 0) o.lead += ' '; else o.trail += ' '; i += 2; continue; }
      if (ch === '*') { i += 2; continue; }
      if ('0#?.,Ee+-'.indexOf(ch) >= 0 && (part < 2) && !(ch === '-' && core === '') && !(ch === '+' && core === '')) {
        if ((ch === 'E' || ch === 'e') && /[0#?]/.test(core)) { o.sci = true; }
        if ((ch === '+' || ch === '-') && !o.sci) { if (part === 0) o.lead += ch; else o.trail += ch; i++; continue; }
        core += ch; part = 1; i++; continue;
      }
      if (ch === '%') { o.pct++; if (part === 0) o.lead += '%'; else o.trail += '%'; i++; continue; }
      if (part === 0) o.lead += ch; else { o.trail += ch; part = 2; }
      i++;
    }
    if (o.sci) {
      var mant = core.split(/[Ee]/)[0];
      o.dec = (mant.split('.')[1] || '').replace(/[^0#?]/g, '').length;
      o.sciDigits = (core.split(/[Ee][+-]?/)[1] || '00').length;
      o.intZero = /0/.test(mant.split('.')[0]) ? 1 : 0;
      return o;
    }
    var m = core.match(/,+$/); if (m) { o.scale = m[0].length; core = core.slice(0, -m[0].length); }
    o.hasDot = core.indexOf('.') >= 0;
    var dotAt = core.indexOf('.');
    var ip = dotAt >= 0 ? core.slice(0, dotAt) : core, dp = dotAt >= 0 ? core.slice(dotAt + 1) : '';
    o.dpat = dp;
    o.thou = ip.indexOf(',') >= 0;
    o.intZero = (ip.replace(/,/g, '').match(/0/g) || []).length;
    o.dec = dp.replace(/[^0#?]/g, '').length;
    o.decOpt = (dp.match(/[#?]/g) || []).length;
    o.decReq = (dp.match(/0/g) || []).length;
    o.empty = core === '';
    return o;
  }
  function getFmt(f) {
    if (fmtCache[f]) return fmtCache[f];
    var secs = [], cur = '', q = false;
    for (var i = 0; i < f.length; i++) {
      var ch = f.charAt(i);
      if (ch === '"') q = !q;
      if (ch === ';' && !q) { secs.push(cur); cur = ''; } else cur += ch;
    }
    secs.push(cur);
    return (fmtCache[f] = secs.map(parseSection));
  }
  function withThousands(s) { return s.replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
  X.fmt = function (x, f) {
    if (!f || f === 'General' || f === 'general') return gtext(x);
    if (/^(d|m|y|h)/i.test(f) && /[dmy]/i.test(f) && !/[0#]/.test(f)) return gtext(x);
    var secs = getFmt(f), sec, neg = x < 0;
    if (secs.length === 1) sec = secs[0];
    else if (x > 0 || (x === 0 && secs.length < 3)) sec = secs[0];
    else if (x < 0) sec = secs[1];
    else sec = secs[2] || secs[0];
    var useNegSign = neg && secs.length === 1;
    if (neg && secs.length > 1) x = -x;
    if (sec.text) return gtext(x);
    var v = Math.abs(x) * Math.pow(100, sec.pct) / Math.pow(1000, sec.scale);
    var body;
    if (sec.sci) {
      if (v === 0) body = (0).toFixed(sec.dec) + 'E+' + '0'.repeat(sec.sciDigits);
      else {
        var e = Math.floor(Math.log10(v)); var mnt = v / Math.pow(10, e);
        if (Number(mnt.toFixed(sec.dec)) >= 10) { mnt /= 10; e++; }
        body = mnt.toFixed(sec.dec) + 'E' + (e < 0 ? '-' : '+') + String(Math.abs(e)).padStart(sec.sciDigits, '0');
      }
    } else if (sec.empty) {
      body = '';
    } else {
      var r = X.ROUND(v, sec.dec);
      body = r.toFixed(sec.dec);
      var ip = body.split('.')[0], digits = body.split('.')[1] || '';
      var pat = sec.dpat || '', slots = [], dp = '';
      for (var pi = 0, di = 0; pi < pat.length; pi++) {
        var pc = pat.charAt(pi);
        if (pc === '0' || pc === '#' || pc === '?') slots.push({ ph: pc, d: digits.charAt(di++) });
        else slots.push({ lit: pc });
      }
      for (var si = slots.length - 1; si >= 0; si--) {               // drop trailing optional zeros
        var sl = slots[si]; if (sl.lit !== undefined) continue;
        if (sl.ph !== '0' && sl.d === '0') { sl.d = sl.ph === '?' ? ' ' : ''; } else break;
      }
      for (si = 0; si < slots.length; si++) dp += slots[si].lit !== undefined ? slots[si].lit : slots[si].d;
      if (ip === '0' && sec.intZero === 0) ip = '';
      if (sec.intZero > 1) ip = ip.padStart(sec.intZero, '0');
      if (sec.thou) ip = withThousands(ip);
      body = ip + (sec.hasDot ? '.' + dp : '');
      if (Number(r) === 0) useNegSign = false;
    }
    return (useNegSign ? '-' : '') + sec.lead + body + sec.trail;
  };
  X.fmtRed = function (x, f) {
    if (!f || f === 'General') return false;
    var secs = getFmt(f);
    return x < 0 && secs.length > 1 && secs[1].red;
  };

  root.X = X;
  if (typeof module !== 'undefined') module.exports = X;
})(typeof window !== 'undefined' ? window : globalThis);
