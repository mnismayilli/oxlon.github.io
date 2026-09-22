/* Makro model — reader for Excel binary workbooks (.xlsb, MS-XLSB / BIFF12). Decodes cells, parsed formulas (rgce → A1 text),
   shared/array formulas, defined names, external-link caches, styles and comments into the 'raw book' format used by
   MakroBuilder.compile. Works in the browser and in Node (needs Blob, DecompressionStream, Response). */
(function (root) {
  'use strict';
  var B = root.MakroBuilder;
  var ERR = { 0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A', 43: '#GETTING_DATA' };
  function colName(n) { var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

  // ------------------------------------------------------------------ binary helpers
  function Rd(u8, p) { this.u = u8; this.dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength); this.p = p || 0; }
  Rd.prototype.u8 = function () { return this.u[this.p++]; };
  Rd.prototype.u16 = function () { var v = this.dv.getUint16(this.p, true); this.p += 2; return v; };
  Rd.prototype.i16 = function () { var v = this.dv.getInt16(this.p, true); this.p += 2; return v; };
  Rd.prototype.u32 = function () { var v = this.dv.getUint32(this.p, true); this.p += 4; return v; };
  Rd.prototype.i32 = function () { var v = this.dv.getInt32(this.p, true); this.p += 4; return v; };
  Rd.prototype.f64 = function () { var v = this.dv.getFloat64(this.p, true); this.p += 8; return v; };
  Rd.prototype.skip = function (n) { this.p += n; };
  Rd.prototype.left = function () { return this.u.length - this.p; };
  var TD = new TextDecoder('utf-16le');
  Rd.prototype.wstr = function () { var n = this.u32(); if (n === 0xFFFFFFFF) return null; var s = TD.decode(this.u.subarray(this.p, this.p + 2 * n)); this.p += 2 * n; return s; };
  Rd.prototype.str16 = function () { var n = this.u16(); var s = TD.decode(this.u.subarray(this.p, this.p + 2 * n)); this.p += 2 * n; return s; };
  Rd.prototype.bytes = function (n) { var v = this.u.subarray(this.p, this.p + n); this.p += n; return v; };
  function records(u8, fn) {
    var i = 0, n = u8.length;
    while (i < n) {
      var b = u8[i++], rt = b & 0x7F;
      if (b & 0x80) { b = u8[i++]; rt |= (b & 0x7F) << 7; }
      var size = 0, shift = 0;
      for (var k = 0; k < 4; k++) { b = u8[i++]; size |= (b & 0x7F) << shift; shift += 7; if (!(b & 0x80)) break; }
      if (fn(rt, u8.subarray(i, i + size)) === false) return;
      i += size;
    }
  }
  function rk(v) {
    var x;
    if (v & 2) x = (v | 0) >> 2;
    else { var b = new DataView(new ArrayBuffer(8)); b.setUint32(4, (v & 0xFFFFFFFC) >>> 0, true); b.setUint32(0, 0, true); x = b.getFloat64(0, true); }
    return v & 1 ? x / 100 : x;
  }
  function readBin(zip, path) {
    var s = zip.stream(path); if (!s) return Promise.resolve(null);
    return new Response(s).arrayBuffer().then(function (ab) { return new Uint8Array(ab); });
  }

  // ------------------------------------------------------------------ function table (MS-XLSB Ftab)
  var FTAB = { 0: 'COUNT', 1: 'IF', 2: 'ISNA', 3: 'ISERROR', 4: 'SUM', 5: 'AVERAGE', 6: 'MIN', 7: 'MAX', 8: 'ROW', 9: 'COLUMN', 10: 'NA', 11: 'NPV', 12: 'STDEV',
    13: 'DOLLAR', 14: 'FIXED', 15: 'SIN', 16: 'COS', 17: 'TAN', 18: 'ATAN', 19: 'PI', 20: 'SQRT', 21: 'EXP', 22: 'LN', 23: 'LOG10', 24: 'ABS', 25: 'INT',
    26: 'SIGN', 27: 'ROUND', 28: 'LOOKUP', 29: 'INDEX', 30: 'REPT', 31: 'MID', 32: 'LEN', 33: 'VALUE', 34: 'TRUE', 35: 'FALSE', 36: 'AND', 37: 'OR',
    38: 'NOT', 39: 'MOD', 40: 'DCOUNT', 41: 'DSUM', 42: 'DAVERAGE', 43: 'DMIN', 44: 'DMAX', 45: 'DSTDEV', 46: 'VAR', 47: 'DVAR', 48: 'TEXT', 49: 'LINEST',
    50: 'TREND', 51: 'LOGEST', 52: 'GROWTH', 56: 'PV', 57: 'FV', 58: 'NPER', 59: 'PMT', 60: 'RATE', 61: 'MIRR', 62: 'IRR', 63: 'RAND', 64: 'MATCH',
    65: 'DATE', 66: 'TIME', 67: 'DAY', 68: 'MONTH', 69: 'YEAR', 70: 'WEEKDAY', 71: 'HOUR', 72: 'MINUTE', 73: 'SECOND', 74: 'NOW', 75: 'AREAS', 76: 'ROWS',
    77: 'COLUMNS', 78: 'OFFSET', 82: 'SEARCH', 83: 'TRANSPOSE', 86: 'TYPE', 97: 'ATAN2', 98: 'ASIN', 99: 'ACOS', 100: 'CHOOSE', 101: 'HLOOKUP',
    102: 'VLOOKUP', 105: 'ISREF', 109: 'LOG', 111: 'CHAR', 112: 'LOWER', 113: 'UPPER', 114: 'PROPER', 115: 'LEFT', 116: 'RIGHT', 117: 'EXACT',
    118: 'TRIM', 119: 'REPLACE', 120: 'SUBSTITUTE', 121: 'CODE', 124: 'FIND', 125: 'CELL', 126: 'ISERR', 127: 'ISTEXT', 128: 'ISNUMBER', 129: 'ISBLANK',
    130: 'T', 131: 'N', 140: 'DATEVALUE', 141: 'TIMEVALUE', 142: 'SLN', 143: 'SYD', 144: 'DDB', 148: 'INDIRECT', 162: 'CLEAN', 163: 'MDETERM',
    164: 'MINVERSE', 165: 'MMULT', 167: 'IPMT', 168: 'PPMT', 169: 'COUNTA', 183: 'PRODUCT', 184: 'FACT', 189: 'DPRODUCT', 190: 'ISNONTEXT',
    193: 'STDEVP', 194: 'VARP', 195: 'DSTDEVP', 196: 'DVARP', 197: 'TRUNC', 198: 'ISLOGICAL', 199: 'DCOUNTA', 212: 'ROUNDUP', 213: 'ROUNDDOWN',
    216: 'RANK', 219: 'ADDRESS', 220: 'DAYS360', 221: 'TODAY', 222: 'VDB', 227: 'MEDIAN', 228: 'SUMPRODUCT', 229: 'SINH', 230: 'COSH', 231: 'TANH',
    232: 'ASINH', 233: 'ACOSH', 234: 'ATANH', 235: 'DGET', 244: 'INFO', 247: 'DB', 252: 'FREQUENCY', 261: 'ERROR.TYPE', 269: 'AVEDEV', 270: 'BETADIST',
    271: 'GAMMALN', 272: 'BETAINV', 273: 'BINOMDIST', 274: 'CHIDIST', 275: 'CHIINV', 276: 'COMBIN', 277: 'CONFIDENCE', 278: 'CRITBINOM', 279: 'EVEN',
    280: 'EXPONDIST', 281: 'FDIST', 282: 'FINV', 283: 'FISHER', 284: 'FISHERINV', 285: 'FLOOR', 286: 'GAMMADIST', 287: 'GAMMAINV', 288: 'CEILING',
    289: 'HYPGEOMDIST', 290: 'LOGNORMDIST', 291: 'LOGINV', 292: 'NEGBINOMDIST', 293: 'NORMDIST', 294: 'NORMSDIST', 295: 'NORMINV', 296: 'NORMSINV',
    297: 'STANDARDIZE', 298: 'ODD', 299: 'PERMUT', 300: 'POISSON', 301: 'TDIST', 302: 'WEIBULL', 303: 'SUMXMY2', 304: 'SUMX2MY2', 305: 'SUMX2PY2',
    306: 'CHITEST', 307: 'CORREL', 308: 'COVAR', 309: 'FORECAST', 310: 'FTEST', 311: 'INTERCEPT', 312: 'PEARSON', 313: 'RSQ', 314: 'STEYX', 315: 'SLOPE',
    316: 'TTEST', 317: 'PROB', 318: 'DEVSQ', 319: 'GEOMEAN', 320: 'HARMEAN', 321: 'SUMSQ', 322: 'KURT', 323: 'SKEW', 324: 'ZTEST', 325: 'LARGE',
    326: 'SMALL', 327: 'QUARTILE', 328: 'PERCENTILE', 329: 'PERCENTRANK', 330: 'MODE', 331: 'TRIMMEAN', 332: 'TINV', 336: 'CONCATENATE', 337: 'POWER',
    342: 'RADIANS', 343: 'DEGREES', 344: 'SUBTOTAL', 345: 'SUMIF', 346: 'COUNTIF', 347: 'COUNTBLANK', 350: 'ISPMT', 351: 'DATEDIF', 354: 'ROMAN',
    358: 'GETPIVOTDATA', 359: 'HYPERLINK', 361: 'AVERAGEA', 362: 'MAXA', 363: 'MINA', 364: 'STDEVPA', 365: 'VARPA', 366: 'STDEVA', 367: 'VARA',
    480: 'IFERROR', 481: 'COUNTIFS', 482: 'SUMIFS', 483: 'AVERAGEIF', 484: 'AVERAGEIFS' };
  // fixed argument counts for PtgFunc (functions that never take a variable number of arguments)
  var FARGS = { 2: 1, 3: 1, 8: 1, 9: 1, 10: 0, 15: 1, 16: 1, 17: 1, 18: 1, 19: 0, 20: 1, 21: 1, 22: 1, 23: 1, 24: 1, 25: 1, 26: 1, 27: 2, 30: 2, 31: 3,
    32: 1, 33: 1, 34: 0, 35: 0, 38: 1, 39: 2, 48: 2, 63: 0, 65: 3, 66: 3, 67: 1, 68: 1, 69: 1, 71: 1, 72: 1, 73: 1, 74: 0, 75: 1, 76: 1, 77: 1, 83: 1,
    86: 1, 97: 2, 98: 1, 99: 1, 105: 1, 111: 1, 112: 1, 113: 1, 114: 1, 117: 2, 118: 1, 119: 4, 121: 1, 126: 1, 127: 1, 128: 1, 129: 1, 130: 1,
    131: 1, 140: 1, 141: 1, 142: 3, 143: 4, 162: 1, 163: 1, 164: 1, 165: 2, 184: 1, 190: 1, 198: 1, 212: 2, 213: 2, 221: 0, 229: 1, 230: 1, 231: 1,
    232: 1, 233: 1, 234: 1, 252: 2, 261: 1, 271: 1, 273: 4, 274: 2, 275: 2, 276: 2, 277: 3, 278: 3, 279: 1, 280: 3, 281: 3, 282: 3, 283: 1, 284: 1,
    285: 2, 286: 4, 287: 3, 288: 2, 289: 4, 290: 3, 291: 3, 292: 3, 293: 4, 294: 1, 295: 3, 296: 1, 297: 3, 298: 1, 299: 2, 300: 3, 301: 3, 302: 4,
    303: 2, 304: 2, 305: 2, 306: 2, 307: 2, 308: 2, 309: 3, 310: 2, 311: 2, 312: 2, 313: 2, 314: 2, 315: 2, 316: 4, 325: 2, 326: 2, 327: 2, 328: 2,
    331: 2, 332: 2, 337: 2, 342: 1, 343: 1, 346: 2, 347: 1, 350: 4, 351: 3, 480: 2 };

  // ------------------------------------------------------------------ formula decoding (rgce → Excel A1 text)
  // ctx: { sheetNames, xti:[{sb, first, last}], supbooks:[{kind:'self'|'ext'|'addin', idx, sheets:[]}], names:[{name, scope}], row, col, extra: Rd over rgcb }
  function cellRef(row, colw, rowRel, colRel) {
    return (colRel ? '' : '$') + colName((colw & 0x3FFF) + 1) + (rowRel ? '' : '$') + (row + 1);
  }
  function relCell(row, colw, ctx) {
    // PtgRefN / PtgAreaN: offsets relative to the formula cell when the relative flags are set
    var rowRel = !!(colw & 0x8000), colRel = !!(colw & 0x4000), c = colw & 0x3FFF, r = row;
    if (rowRel) r = (ctx.row + (row | 0)) ; // row is a signed offset
    if (colRel) { var off = c & 0x2000 ? c - 0x4000 : c; c = ctx.col + off; }
    r = ((r % 1048576) + 1048576) % 1048576; c = ((c % 16384) + 16384) % 16384;
    return (colRel ? '' : '$') + colName(c + 1) + (rowRel ? '' : '$') + (r + 1);
  }
  function quoteSheet(n) { return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) && !/^(R|C|RC)\d*$/i.test(n) ? n : "'" + n.replace(/'/g, "''") + "'"; }
  function xtiPrefix(ixti, ctx) {
    var x = ctx.xti[ixti]; if (!x) return { bad: true, p: '' };
    var sb = ctx.supbooks[x.sb]; if (!sb) return { bad: true, p: '' };
    if (sb.kind === 'self') {
      if (x.first < 0) return { bad: true, p: '' };
      var n1 = ctx.sheetNames[x.first], n2 = ctx.sheetNames[x.last];
      if (n1 === undefined) return { bad: true, p: '' };
      return { p: (x.first === x.last ? quoteSheet(n1) : quoteSheet(n1 + ':' + n2)) + '!' };
    }
    if (sb.kind === 'ext') {
      var sn = x.first >= 0 ? (sb.sheets[x.first] || '?') : '';
      return { p: "'[" + (sb.linkIndex + 1) + ']' + sn.replace(/'/g, "''") + "'!", ext: true };
    }
    return { p: '', addin: true };
  }
  function decodeFormula(rgce, ctx) {
    var rd = new Rd(rgce), st = [], extra = ctx.extra;
    function pop() { if (!st.length) throw new Error('rgce stack underflow'); return st.pop(); }
    function binop(op) { var b = pop(), a = pop(); st.push(a + op + b); }
    while (rd.left() > 0) {
      var ptg = rd.u8(), base = ptg & 0x1F, cls = ptg & 0x60;
      if (ptg === 0x18 || ptg === 0x19) base = ptg;
      if (!cls || ptg === 0x18 || ptg === 0x19) {
        switch (ptg) {
          case 0x01: throw new Error('PtgExp'); // handled by caller
          case 0x02: throw new Error('PtgTbl (data table)');
          case 0x03: binop('+'); break; case 0x04: binop('-'); break; case 0x05: binop('*'); break; case 0x06: binop('/'); break;
          case 0x07: binop('^'); break; case 0x08: binop('&'); break; case 0x09: binop('<'); break; case 0x0A: binop('<='); break;
          case 0x0B: binop('='); break; case 0x0C: binop('>='); break; case 0x0D: binop('>'); break; case 0x0E: binop('<>'); break;
          case 0x0F: binop(' '); break; case 0x10: binop(','); break; case 0x11: binop(':'); break;
          case 0x12: st.push('+' + pop()); break; case 0x13: st.push('-' + pop()); break; case 0x14: st.push(pop() + '%'); break;
          case 0x15: st.push('(' + pop() + ')'); break;
          case 0x16: st.push(''); break;
          case 0x17: st.push('"' + rd.str16().replace(/"/g, '""') + '"'); break;
          case 0x18: { var e = rd.u8(); throw new Error('PtgExtended ' + e); }
          case 0x19: {
            var t = rd.u8();
            if (t & 0x04) { var n = rd.u16(); rd.skip(2 * (n + 1)); }        // tAttrChoose
            else if (t & 0x10) { rd.u16(); st.push('SUM(' + pop() + ')'); } // tAttrSum
            else if (t & 0x40 || t & 0x41) { rd.u16(); }                     // tAttrSpace
            else rd.u16();                                                     // semi, if, goto, baxcel
            break;
          }
          case 0x1C: st.push(ERR[rd.u8()] || '#N/A'); break;
          case 0x1D: st.push(rd.u8() ? 'TRUE' : 'FALSE'); break;
          case 0x1E: st.push(String(rd.u16())); break;
          case 0x1F: { var x = rd.f64(); st.push(numText(x)); break; }
          default: throw new Error('ptg 0x' + ptg.toString(16));
        }
        continue;
      }
      switch (ptg & 0x1F | 0x20) {
        case 0x20: { // PtgArray: data in rgcb
          rd.skip(14);
          var rows = extra.u32(), cols = extra.u32(), rowsTxt = [];
          for (var i = 0; i < rows; i++) {
            var cells = [];
            for (var j = 0; j < cols; j++) {
              var ty = extra.u8();
              if (ty === 0x00) cells.push(numText(extra.f64()));
              else if (ty === 0x01) cells.push('"' + extra.str16().replace(/"/g, '""') + '"');
              else if (ty === 0x02) { cells.push(extra.u8() ? 'TRUE' : 'FALSE'); extra.skip(0); }
              else if (ty === 0x04) { cells.push(ERR[extra.u8()] || '#N/A'); extra.skip(3); }
              else throw new Error('SerAr type ' + ty);
            }
            rowsTxt.push(cells.join(','));
          }
          st.push('{' + rowsTxt.join(';') + '}'); break;
        }
        case 0x21: { var ift = rd.u16(), fn = FTAB[ift]; if (fn === undefined || FARGS[ift] === undefined) throw new Error('PtgFunc ' + ift); var na = FARGS[ift], args = []; for (var a = 0; a < na; a++) args.unshift(pop()); st.push(fn + '(' + args.join(',') + ')'); break; }
        case 0x22: {
          var np = rd.u8(), tab = rd.u16() & 0x7FFF, args2 = [];
          for (var a2 = 0; a2 < np; a2++) args2.unshift(pop());
          if (tab === 255) { var fname = args2.shift(); st.push(fname.replace(/^_xlfn\./, '') + '(' + args2.join(',') + ')'); }
          else { if (FTAB[tab] === undefined) throw new Error('PtgFuncVar ' + tab); st.push(FTAB[tab] + '(' + args2.join(',') + ')'); }
          break;
        }
        case 0x23: { var ni = rd.u32(), nm = ctx.names[ni - 1]; if (!nm) throw new Error('PtgName ' + ni); st.push(nm.name); break; }
        case 0x24: { var r1 = rd.u32(), c1 = rd.u16(); st.push(cellRef(r1, c1, !!(c1 & 0x8000), !!(c1 & 0x4000))); break; }
        case 0x25: { var ra = rd.u32(), rb = rd.u32(), ca = rd.u16(), cb = rd.u16(); st.push(areaText(ra, rb, ca, cb)); break; }
        case 0x26: { rd.u32(); rd.u16(); skipMem(extra); break; }            // PtgMemArea (+ PtgExtraMem in rgcb)
        case 0x27: { rd.u32(); rd.u16(); break; }                            // PtgMemErr
        case 0x28: { rd.u32(); rd.u16(); break; }                            // PtgMemNoMem
        case 0x29: { rd.u16(); break; }                                      // PtgMemFunc
        case 0x2A: { rd.u32(); rd.u16(); st.push('#REF!'); break; }
        case 0x2B: { rd.u32(); rd.u32(); rd.u16(); rd.u16(); st.push('#REF!'); break; }
        case 0x2C: { var rn = rd.i32(), cn = rd.u16(); st.push(relCell(rn, cn, ctx)); break; }
        case 0x2D: { var rn1 = rd.i32(), rn2 = rd.i32(), cn1 = rd.u16(), cn2 = rd.u16(); st.push(relCell(rn1, cn1, ctx) + ':' + relCell(rn2, cn2, ctx)); break; }
        case 0x39: {
          var ixn = rd.u16(), nidx = rd.u32(), xt = ctx.xti[ixn], sbk = xt ? ctx.supbooks[xt.sb] : null, nm2;
          if (sbk && sbk.kind === 'addin') nm2 = sbk.names[nidx - 1];
          else if (sbk && sbk.kind === 'ext') nm2 = "'[" + (sbk.linkIndex + 1) + "]'!" + (sbk.names[nidx - 1] || ('NAME' + nidx));
          else if (sbk && sbk.kind === 'self') nm2 = (ctx.names[nidx - 1] || {}).name;
          if (!nm2) throw new Error('PtgNameX ' + ixn + '/' + nidx);
          st.push(nm2); break;
        }
        case 0x3A: { var ix = rd.u16(), rr = rd.u32(), cc = rd.u16(), px = xtiPrefix(ix, ctx); st.push(px.bad ? '#REF!' : px.p + cellRef(rr, cc, !!(cc & 0x8000), !!(cc & 0x4000))); break; }
        case 0x3B: { var ix2 = rd.u16(), q1 = rd.u32(), q2 = rd.u32(), w1 = rd.u16(), w2 = rd.u16(), px2 = xtiPrefix(ix2, ctx); st.push(px2.bad ? '#REF!' : px2.p + areaText(q1, q2, w1, w2)); break; }
        case 0x3C: { var ix3 = rd.u16(); rd.u32(); rd.u16(); var p3 = xtiPrefix(ix3, ctx); st.push((p3.bad ? '' : p3.p) + '#REF!'); break; }
        case 0x3D: { var ix4 = rd.u16(); rd.u32(); rd.u32(); rd.u16(); rd.u16(); var p4 = xtiPrefix(ix4, ctx); st.push((p4.bad ? '' : p4.p) + '#REF!'); break; }
        default: throw new Error('ptg 0x' + ptg.toString(16));
      }
    }
    if (st.length !== 1) throw new Error('rgce stack ' + st.length);
    return st[0];
  }
  function areaText(r1, r2, c1, c2) {
    var whole = r1 === 0 && r2 === 0xFFFFF, wholeRow = (c1 & 0x3FFF) === 0 && (c2 & 0x3FFF) === 0x3FFF;
    if (whole) return (c1 & 0x4000 ? '' : '$') + colName((c1 & 0x3FFF) + 1) + ':' + (c2 & 0x4000 ? '' : '$') + colName((c2 & 0x3FFF) + 1);
    if (wholeRow) return (c1 & 0x8000 ? '' : '$') + (r1 + 1) + ':' + (c2 & 0x8000 ? '' : '$') + (r2 + 1);
    return cellRef(r1, c1, !!(c1 & 0x8000), !!(c1 & 0x4000)) + ':' + cellRef(r2, c2, !!(c2 & 0x8000), !!(c2 & 0x4000));
  }
  function skipMem(extra) { var n = extra.u32(); extra.skip(16 * n); }   // PtgExtraMem: count + UncheckedRfX[count]
  function numText(x) {
    if (!isFinite(x)) return '0';
    var s = String(x);
    return s.indexOf('e') >= 0 ? s.replace('e+', 'E').replace('e-', 'E-').replace('e', 'E') : s;
  }

  // ------------------------------------------------------------------ styles (styles.bin)
  var INDEXED = ['000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','800000','008000','000080','808000','800080','008080','C0C0C0','808080','9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF','000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF','00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99','3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696','003366','339966','003300','333300','993300','993366','333399','333333'];
  var BUILTIN_FMT = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 14: 'dd.mm.yyyy', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 49: '@' };
  function hex2(n) { var s = n.toString(16).toUpperCase(); return s.length < 2 ? '0' + s : s; }
  function brtColor(rd, theme) {
    var b0 = rd.u8(), idx = rd.u8(), tint = rd.i16(), r = rd.u8(), g = rd.u8(), b = rd.u8(); rd.u8();
    var type = b0 >> 1, v = null;
    if (type === 2) v = hex2(r) + hex2(g) + hex2(b);
    else if (type === 1) v = INDEXED[idx] || null;
    else if (type === 3 && theme) v = theme[idx] || null;
    if (v && tint) v = B.tintHex ? B.tintHex(v, tint / 32767) : v;
    return v;
  }
  function parseStyles(u8, theme) {
    var fmts = {}, fonts = [], fills = [], xfs = [], sec = null;
    for (var k in BUILTIN_FMT) fmts[k] = BUILTIN_FMT[k];
    records(u8, function (rt, d) {
      var rd = new Rd(d);
      if (rt === 0x2C) { var id = rd.u16(); fmts[id] = rd.wstr(); }
      else if (rt === 0x2B) { rd.u16(); var gr = rd.u16(), bls = rd.u16(); rd.skip(2 + 4); var col = brtColor(rd, theme); fonts.push({ b: bls >= 700, i: !!(gr & 2), c: col }); }
      else if (rt === 0x2D) { var fls = rd.u32(), fg = brtColor(rd, theme); fills.push(fls ? fg : null); }
      else if (rt === 0x269) sec = 'cell'; else if (rt === 0x26A) sec = null; else if (rt === 0x272) sec = 'style'; else if (rt === 0x273) sec = null;
      else if (rt === 0x2F && sec === 'cell') {
        rd.u16(); var ifmt = rd.u16(), ifont = rd.u16(), ifill = rd.u16(); rd.u16(); rd.u8(); var indent = rd.u8(), fl = rd.u16();
        var font = fonts[ifont] || { b: false, i: false, c: null }, fill = fills[ifill], alc = fl & 7;
        xfs.push({ fmt: fmts[ifmt] || 'General', b: font.b, i: font.i, fc: font.c && font.c !== '000000' ? font.c : null, bg: fill && fill !== 'FFFFFF' ? fill : null,
          ha: ({ 1: 'left', 2: 'center', 3: 'right', 6: 'centerContinuous' })[alc] || null, ind: indent, wrap: !!(fl & 0x40) });
      }
    });
    return xfs;
  }

  // ------------------------------------------------------------------ workbook
  function readXlsb(file, bookName, log) {
    var zip, wbRels = {}, book = { name: bookName, kind: 'model', source: 'xlsb', file: file.name, sheets: [], ext: [], xfs: [], names: [], calc: {}, arrays: [], placeholders: true };
    var sheetsIdx = [], supbooks = [], xti = [], sst = [], theme = null;
    var te = function (p) { return zip.text(p); };
    return B.openZip(file).then(function (z) {
      zip = z;
      return Promise.all([te('xl/_rels/workbook.bin.rels'), readBin(zip, 'xl/workbook.bin'), readBin(zip, 'xl/sharedStrings.bin'), te('xl/theme/theme1.xml'), readBin(zip, 'xl/styles.bin')]);
    }).then(function (r) {
      (r[0] || '').replace(/<Relationship\b[^>]*>/g, function (t) { var a = B.attrs(t); wbRels[a.Id] = a.Target; return ''; });
      theme = B.parseTheme ? B.parseTheme(r[3]) : null;
      book.xfs = r[4] ? parseStyles(r[4], theme) : [];
      if (r[2]) records(r[2], function (rt, d) { if (rt === 19) { var rd = new Rd(d); rd.u8(); sst.push(rd.wstr() || ''); } });
      var pendingNames = [];
      records(r[1], function (rt, d) {
        var rd = new Rd(d);
        if (rt === 156) { var state = rd.u32(); rd.u32(); var rel = rd.wstr(), nm = rd.wstr(); sheetsIdx.push({ name: nm, rel: rel, state: state === 0 ? 'visible' : state === 1 ? 'hidden' : 'veryHidden' }); }
        else if (rt === 357 || rt === 356) supbooks.push({ kind: 'self' });                                    // BrtSupSelf / BrtSupSame
        else if (rt === 355) { var rel2 = rd.wstr(); supbooks.push({ kind: 'ext', rel: rel2, sheets: [], names: [] }); }   // BrtSupBookSrc
        else if (rt === 358) supbooks.push({ kind: 'addin', names: [] });                                      // BrtSupAddin
        else if (rt === 362) { var n = rd.u32(); for (var i = 0; i < n; i++) xti.push({ sb: rd.u32(), first: rd.i32(), last: rd.i32() }); }
        else if (rt === 39) {
          var fl = rd.u32(); rd.u8(); var itab = rd.u32(), name = rd.wstr(), cce = rd.u32(), rgce = rd.bytes(cce), cb = rd.u32(), rgcb = rd.bytes(cb);
          pendingNames.push({ name: name, scope: itab === 0xFFFFFFFF ? null : itab, hidden: !!(fl & 1), func: !!(fl & 2), builtin: !!(fl & 0x20), rgce: rgce, rgcb: rgcb });
        }
        else if (rt === 157) { rd.u32(); book.calc.auto = rd.u32(); book.calc.count = rd.u32(); book.calc.delta = rd.f64(); rd.u32(); var cf = rd.u16(); book.calc.iterate = !!(cf & 4); }
      });
      // add-in (future function) names follow BrtSupAddin as BrtPlaceholderName records in the same stream
      var cur = -1;
      records(r[1], function (rt, d) {
        if (rt === 355 || rt === 356 || rt === 357 || rt === 358) cur++;
        else if (rt === 0x169 && cur >= 0 && supbooks[cur]) { var rd = new Rd(d); supbooks[cur].names.push(rd.wstr()); }
      });
      book.names = pendingNames.map(function (n) { return { name: n.name, scope: n.scope, hidden: n.hidden, builtin: n.builtin }; });
      book.sheetNames = sheetsIdx.map(function (s) { return s.name; });
      // external links
      var chain = Promise.resolve(), linkIndex = 0;
      supbooks.forEach(function (sb) {
        if (sb.kind !== 'ext') return;
        sb.linkIndex = linkIndex++;
        chain = chain.then(function () {
          var p = B.resolvePath('xl', wbRels[sb.rel]);
          return Promise.all([readBin(zip, p), te(p.replace('externalLinks/', 'externalLinks/_rels/') + '.rels')]).then(function (x) {
            var tg = null; (x[1] || '').replace(/<Relationship\b[^>]*>/g, function (t) { tg = B.attrs(t).Target; return ''; });
            var cache = {}, cursheet = null, currow = 0;
            if (x[0]) records(x[0], function (rt, d) {
              var rd = new Rd(d);
              if (rt === 0x167) { var n = rd.u32(); for (var i = 0; i < n; i++) sb.sheets.push(rd.wstr()); }            // BrtSupTabs
              else if (rt === 0x169) sb.names.push(new Rd(d).wstr());                                                      // BrtPlaceholderName
              else if (rt === 0x16B) { var itab = rd.u32(); cursheet = sb.sheets[itab]; if (cursheet !== undefined && !cache[cursheet]) cache[cursheet] = {}; }   // BrtExternTableStart
              else if (rt === 0x16E) currow = rd.u32();                                                                     // BrtExternRowHdr
              else if (rt >= 0x16F && rt <= 0x173 && cursheet !== undefined && cursheet !== null) {
                var col = rd.u32(), v = null;
                if (rt === 0x170) v = rd.f64(); else if (rt === 0x171) v = !!rd.u8(); else if (rt === 0x172) v = { e: ERR[rd.u8()] || '#N/A' }; else if (rt === 0x173) v = rd.wstr();
                if (rt !== 0x16F) cache[cursheet][colName(col + 1) + (currow + 1)] = v;
              }
            });
            book.ext[sb.linkIndex] = { target: tg, file: tg ? B.decodePath(tg) : null, sheets: sb.sheets, cache: cache };
          });
        });
      });
      return chain.then(function () {
        var ctxBase = { sheetNames: book.sheetNames, xti: xti, supbooks: supbooks, names: book.names };
        // decode defined-name formulas (absolute references; relative ones are evaluated from A1 of the scope sheet)
        pendingNames.forEach(function (n, i) {
          try { book.names[i].f = decodeFormula(n.rgce, Object.assign({}, ctxBase, { row: 0, col: 0, extra: new Rd(n.rgcb) })); }
          catch (e) { book.names[i].f = null; book.names[i].err = e.message; }
        });
        book.problems = [];
        var ch = Promise.resolve();
        sheetsIdx.forEach(function (si, idx) {
          ch = ch.then(function () {
            var path = B.resolvePath('xl', wbRels[si.rel]);
            if (log) log('  «' + si.name + '» vərəqi oxunur…');
            return readBin(zip, path).then(function (u8) { book.sheets.push(readSheet(u8, si, idx, ctxBase, sst, book, theme)); })
              .then(function () { return readComments(zip, path, book.sheets[book.sheets.length - 1]); });
          });
        });
        return ch;
      });
    }).then(function () { return book; });
  }
  function readComments(zip, sheetPath, sheet) {
    var parts = sheetPath.split('/'), fn = parts.pop();
    return zip.text(parts.join('/') + '/_rels/' + fn + '.rels').then(function (rx) {
      var cp = null; (rx || '').replace(/<Relationship\b[^>]*>/g, function (t) { var a = B.attrs(t); if (/comments\d*\.bin$/.test(a.Target)) cp = B.resolvePath(parts.join('/'), a.Target); return ''; });
      return cp ? readBin(zip, cp) : null;
    }).then(function (u8) {
      if (!u8) return;
      var ref = null;
      records(u8, function (rt, d) {
        var rd = new Rd(d);
        if (rt === 0x27B) { rd.u32(); var r1 = rd.u32(); rd.u32(); var c1 = rd.u32(); ref = colName(c1 + 1) + (r1 + 1); }     // BrtBeginComment
        else if (rt === 0x27D && ref) { rd.u8(); var t = rd.wstr(); sheet.cm[ref] = (t || '').trim(); ref = null; }          // BrtCommentText
      });
    });
  }
  function readSheet(u8, si, idx, ctxBase, sst, book, theme) {
    var sheet = { name: si.name, state: si.state, cells: [], merges: [], cols: [], hr: [], fz: null, cm: {}, tab: null, dw: 9.14, problems: [] };
    var row = 0, pend = [], shared = [], arrays = [];
    function cellHead(rd) { var col = rd.u32(), sf = rd.u32(); return { c: col, s: sf & 0xFFFFFF }; }
    records(u8, function (rt, d) {
      var rd = new Rd(d), h, v;
      switch (rt) {
        case 0: row = rd.u32(); rd.u32(); rd.u16(); var fl = rd.u16(); if (fl & 0x1000) sheet.hr.push(row + 1); break;   // BrtRowHdr (fDyZero = hidden)
        case 1: { h = cellHead(rd); var xs = book.xfs[h.s]; if (xs && (/^(FFFF00|FFFFCC)$/i.test(xs.bg || '') || /^FF0000$/i.test(xs.fc || ''))) sheet.cells.push({ r: row + 1, c: h.c + 1, v: null, f: null, s: h.s }); break; }   // blank input cell (yellow fill / red font)
        case 2: h = cellHead(rd); push(h, rk(rd.u32())); break;
        case 3: h = cellHead(rd); push(h, { e: ERR[rd.u8()] || '#N/A' }); break;
        case 4: h = cellHead(rd); push(h, !!rd.u8()); break;
        case 5: h = cellHead(rd); push(h, rd.f64()); break;
        case 6: h = cellHead(rd); push(h, rd.wstr()); break;
        case 7: h = cellHead(rd); push(h, sst[rd.u32()]); break;
        case 8: case 9: case 10: case 11: {
          h = cellHead(rd);
          if (rt === 8) v = rd.wstr(); else if (rt === 9) v = rd.f64(); else if (rt === 10) v = !!rd.u8(); else v = { e: ERR[rd.u8()] || '#N/A' };
          rd.u16();
          var cce = rd.u32(), rgce = rd.bytes(cce), cb = rd.u32(), rgcb = rd.bytes(cb);
          pend.push({ r: row, c: h.c, v: v, s: h.s, rgce: rgce, rgcb: rgcb });
          break;
        }
        case 426: { var ar = [rd.u32(), rd.u32(), rd.u32(), rd.u32()]; rd.u8(); var ce = rd.u32(), rg = rd.bytes(ce), cb2 = rd.u32(), rb = rd.bytes(cb2); arrays.push({ rf: ar, rgce: rg, rgcb: rb }); break; }  // BrtArrFmla
        case 427: { var sr = [rd.u32(), rd.u32(), rd.u32(), rd.u32()]; var ce2 = rd.u32(), rg2 = rd.bytes(ce2), cb3 = rd.u32(), rb2 = rd.bytes(cb3); shared.push({ rf: sr, rgce: rg2, rgcb: rb2 }); break; }  // BrtShrFmla
        case 60: { var c1 = rd.u32(), c2 = rd.u32(), w = rd.u32(); rd.u32(); var f2 = rd.u16(); sheet.cols.push({ min: c1 + 1, max: c2 + 1, w: w / 256, hidden: !!(f2 & 1) }); break; }
        case 151: { var xs = rd.f64(), ys = rd.f64(); rd.u32(); rd.u32(); rd.u32(); var pf = rd.u8(); if (pf & 1) sheet.fz = [Math.round(ys), Math.round(xs)]; break; }
        case 147: { rd.u16(); rd.u8(); var tc = brtColor(rd, theme); if (tc) sheet.tab = tc; break; }   // BrtWsProp: tab colour
        case 176: { var m1 = rd.u32(), m2 = rd.u32(), m3 = rd.u32(), m4 = rd.u32(); sheet.merges.push(colName(m3 + 1) + (m1 + 1) + ':' + colName(m4 + 1) + (m2 + 1)); break; }
        default: break;
      }
    });
    function push(h, val) { if (val === null || val === undefined || val === '') return; sheet.cells.push({ r: row + 1, c: h.c + 1, v: val, f: null, s: h.s }); }
    // formulas (after all shared/array definitions are known)
    pend.forEach(function (p) {
      var ctx = Object.assign({}, ctxBase, { row: p.r, col: p.c }), f = null, src = p;
      try {
        if (p.rgce.length && p.rgce[0] === 0x01) {
          var sh = shared.filter(function (s) { return p.r >= s.rf[0] && p.r <= s.rf[1] && p.c >= s.rf[2] && p.c <= s.rf[3]; })[0];
          var arr = arrays.filter(function (s) { return p.r >= s.rf[0] && p.r <= s.rf[1] && p.c >= s.rf[2] && p.c <= s.rf[3]; })[0];
          if (arr) {
            var single = arr.rf[0] === arr.rf[1] && arr.rf[2] === arr.rf[3];
            ctx.extra = new Rd(arr.rgcb);
            f = decodeFormula(arr.rgce, ctx);
            if (!single) {
              if (arr.gid === undefined) { arr.gid = book.arrays.length; book.arrays.push({ sheet: si.name, r1: arr.rf[0] + 1, r2: arr.rf[1] + 1, c1: arr.rf[2] + 1, c2: arr.rf[3] + 1, f: f }); }
              f = '{AGROUP:' + arr.gid + '}';
            }
            else f = '{ARRAY}' + f;
          } else if (sh) { ctx.extra = new Rd(sh.rgcb); f = decodeFormula(sh.rgce, ctx); }
          else throw new Error('paylaşılan düstur tapılmadı');
        } else { ctx.extra = new Rd(p.rgcb); f = decodeFormula(p.rgce, ctx); }
      } catch (e) { sheet.problems.push([colName(p.c + 1) + (p.r + 1), 'düstur oxunmadı (' + e.message + ') — Excel dəyəri saxlanıldı']); f = null; }
      var val = src.v;
      if ((val === null || val === '') && f === null) return;
      sheet.cells.push({ r: p.r + 1, c: p.c + 1, v: val, f: f, s: p.s });
    });
    return sheet;
  }

  root.MakroXlsb = { readXlsb: readXlsb, records: records, decodeFormula: decodeFormula, FTAB: FTAB };
})(typeof window !== 'undefined' ? window : globalThis);
