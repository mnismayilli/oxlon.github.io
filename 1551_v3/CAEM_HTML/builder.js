/* Makro model — in-browser model builder: reads .xlsx files, parses and compiles every formula, verifies the result
   against Excel's cached values and exports data files. Port of the verified Python pipeline (extract.py + compile.py). */
(function (root) {
  'use strict';
  var MAXROW = 1048576, MAXCOL = 16384;
  var ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  function unesc(s) {
    if (s.indexOf('\r') >= 0) s = s.replace(/\r\n?/g, '\n');     // XML end-of-line normalisation (as any XML parser)
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, function (m, e) {
      if (e.charAt(0) === '#') return String.fromCodePoint(e.charAt(1) === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
      return ENT[e];
    });
  }
  function attrs(s) { var o = {}, re = /([\w:]+)="([^"]*)"/g, m; while ((m = re.exec(s))) o[m[1]] = unesc(m[2]); return o; }
  // '>' is legal inside an XML attribute value, so tags are matched quote-aware (sheet names such as «Appendix =>>»)
  function tagRe(name) { return new RegExp('<' + name + '\\b(?:"[^"]*"|\'[^\']*\'|[^>"\'])*\\/?>', 'g'); }
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
  function colName(n) { var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function splitRef(a) { var m = /^([A-Z]+)(\d+)$/.exec(a); return m ? [+m[2], colNum(m[1])] : null; }
  function tick() { return new Promise(function (r) { setTimeout(r, 0); }); }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function norm(s) {
    return String(s).toLocaleLowerCase('az').replace(/\.(xlsx|xlsm)$/i, '').replace(/\s*\(\d+\)\s*$/, '').replace(/i̇/g, 'i').replace(/ı/g, 'i').replace(/ə/g, 'e')
      .replace(/ö/g, 'o').replace(/ü/g, 'u').replace(/ğ/g, 'g').replace(/ş/g, 's').replace(/ç/g, 'c').replace(/[\s_\-]+/g, ' ').trim();
  }
  function decodePath(t) {
    t = unesc(t);
    try { t = decodeURIComponent(t); } catch (e) { t = t.replace(/%20/g, ' '); }
    var parts = t.replace(/\\/g, '/').split('/');
    return parts[parts.length - 1];
  }

  // ------------------------------------------------------------------ ZIP reading (DecompressionStream)
  function openZip(blob) {
    return blob.arrayBuffer().then(function (ab) {
      var buf = new Uint8Array(ab), dv = new DataView(ab), eocd = -1;
      for (var i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
      if (eocd < 0) throw new Error('Fayl Excel (.xlsx) formatında deyil');
      var n = dv.getUint16(eocd + 10, true), off = dv.getUint32(eocd + 16, true), entries = {}, td = new TextDecoder();
      if (off === 0xffffffff) throw new Error('ZIP64 formatı dəstəklənmir');
      for (var k = 0; k < n; k++) {
        if (dv.getUint32(off, true) !== 0x02014b50) throw new Error('ZIP strukturu pozulub');
        var method = dv.getUint16(off + 10, true), csize = dv.getUint32(off + 20, true), nlen = dv.getUint16(off + 28, true), elen = dv.getUint16(off + 30, true), clen = dv.getUint16(off + 32, true), loff = dv.getUint32(off + 42, true);
        entries[td.decode(buf.subarray(off + 46, off + 46 + nlen)).replace(/^\//, '')] = { method: method, csize: csize, loff: loff };
        off += 46 + nlen + elen + clen;
      }
      function data(e) { var l = e.loff, s = l + 30 + dv.getUint16(l + 26, true) + dv.getUint16(l + 28, true); return buf.subarray(s, s + e.csize); }
      function stream(name) {
        var e = entries[name]; if (!e) return null;
        var src = new Blob([data(e)]).stream();
        if (e.method === 0) return src;
        if (e.method === 8) return src.pipeThrough(new DecompressionStream('deflate-raw'));
        throw new Error('Dəstəklənməyən sıxılma üsulu: ' + e.method);
      }
      function text(name) { var s = stream(name); if (!s) return Promise.resolve(null); return new Response(s).text(); }
      return { entries: entries, stream: stream, text: text, has: function (n) { return !!entries[n]; } };
    });
  }
  function resolvePath(base, target) {
    if (target.charAt(0) === '/') return target.slice(1);
    var parts = (base + '/' + target).split('/'), out = [];
    parts.forEach(function (p) { if (p === '..') out.pop(); else if (p && p !== '.') out.push(p); });
    return out.join('/');
  }
  function rels(xml) { var o = {}; if (!xml) return o; (xml.match(tagRe('Relationship')) || []).forEach(function (t) { var a = attrs(t); o[a.Id] = a.Target; }); return o; }

  // ------------------------------------------------------------------ styles
  var INDEXED = ['000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','800000','008000','000080','808000','800080','008080','C0C0C0','808080','9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF','000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF','00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99','3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696','003366','339966','003300','333300','993300','993366','333399','333333'];
  var BUILTIN_FMT = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??', 14: 'dd.mm.yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 48: '##0.0E+0', 49: '@' };
  function rgb2hls(r, g, b) {
    var mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2, h, s;
    if (mx === mn) return [0, l, 0];
    s = l <= 0.5 ? (mx - mn) / (mx + mn) : (mx - mn) / (2 - mx - mn);
    var rc = (mx - r) / (mx - mn), gc = (mx - g) / (mx - mn), bc = (mx - b) / (mx - mn);
    h = r === mx ? bc - gc : g === mx ? 2 + rc - bc : 4 + gc - rc;
    h = ((h / 6) % 1 + 1) % 1; return [h, l, s];
  }
  function hls2rgb(h, l, s) {
    if (s === 0) return [l, l, l];
    var m2 = l <= 0.5 ? l * (1 + s) : l + s - l * s, m1 = 2 * l - m2;
    function v(hh) { hh = ((hh % 1) + 1) % 1; if (hh < 1 / 6) return m1 + (m2 - m1) * hh * 6; if (hh < 0.5) return m2; if (hh < 2 / 3) return m1 + (m2 - m1) * (2 / 3 - hh) * 6; return m1; }
    return [v(h + 1 / 3), v(h), v(h - 1 / 3)];
  }
  function tint(hex, t) {
    var r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255, hls = rgb2hls(r, g, b), l = hls[1];
    l = t < 0 ? l * (1 + t) : l * (1 - t) + t;
    var o = hls2rgb(hls[0], l, hls[2]);
    return o.map(function (x) { var s = Math.round(x * 255).toString(16).toUpperCase(); return s.length < 2 ? '0' + s : s; }).join('');
  }
  function colorOf(tag, theme) {
    if (!tag) return null;
    var a = attrs(tag), v;
    if (a.rgb) v = a.rgb.slice(-6);
    else if (a.theme !== undefined && theme) { v = theme[+a.theme]; if (!v) return null; }
    else if (a.indexed !== undefined) { v = INDEXED[+a.indexed]; if (!v) return null; }
    else return null;
    if (a.tint) v = tint(v, parseFloat(a.tint));
    return v;
  }
  function parseTheme(xml) {
    if (!xml) return null;
    var cs = /<a:clrScheme\b[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(xml); if (!cs) return null;
    var cols = [], re = /<a:(\w+)>\s*<a:(?:sysClr|srgbClr)\b([^>]*)>/g, m;
    while ((m = re.exec(cs[1]))) { var a = attrs(m[2]); cols.push(a.lastClr || a.val); }
    return [cols[1], cols[0], cols[3], cols[2]].concat(cols.slice(4));
  }
  function parseStyles(xml, theme) {
    var fmts = {}; for (var k in BUILTIN_FMT) fmts[k] = BUILTIN_FMT[k];
    (xml.match(/<numFmt\b[^>]*\/?>/g) || []).forEach(function (t) { var a = attrs(t); fmts[+a.numFmtId] = a.formatCode; });
    function section(name) { var m = new RegExp('<' + name + '\\b[^>]*>([\\s\\S]*?)</' + name + '>').exec(xml); return m ? m[1] : ''; }
    var fonts = (section('fonts').match(/<font\b[^>]*?(?:\/>|>[\s\S]*?<\/font>)/g) || []).map(function (f) {
      var b = /<b\b([^>]*)\/?>/.exec(f), i = /<i\b([^>]*)\/?>/.exec(f), c = /<color\b[^>]*\/?>/.exec(f);
      var off = function (m) { if (!m) return false; var a = attrs(m[1]); return !(a.val === '0' || a.val === 'false'); };
      return { b: off(b), i: off(i), c: c ? colorOf(c[0], theme) : null };
    });
    var fills = (section('fills').match(/<fill\b[^>]*?(?:\/>|>[\s\S]*?<\/fill>)/g) || []).map(function (f) {
      var pf = /<patternFill\b([^>]*)(?:\/>|>([\s\S]*?)<\/patternFill>)/.exec(f);
      if (!pf) return null; var a = attrs(pf[1]); if (!a.patternType || a.patternType === 'none') return null;
      var fg = pf[2] ? /<fgColor\b[^>]*\/?>/.exec(pf[2]) : null; return fg ? colorOf(fg[0], theme) : null;
    });
    return (section('cellXfs').match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g) || []).map(function (x) {
      var a = attrs(/<xf\b[^>]*/.exec(x)[0]), al = /<alignment\b[^>]*\/?>/.exec(x), aa = al ? attrs(al[0]) : null;
      var font = fonts[+(a.fontId || 0)] || { b: false, i: false, c: null }, fill = fills[+(a.fillId || 0)];
      return { fmt: fmts[+(a.numFmtId || 0)] || 'General', b: font.b, i: font.i, fc: font.c && font.c !== '000000' ? font.c : null,
        bg: fill && fill !== 'FFFFFF' ? fill : null, ha: aa ? aa.horizontal || null : null, ind: aa ? +(aa.indent || 0) : 0, wrap: !!(aa && aa.wrapText === '1') };
    });
  }

  // ------------------------------------------------------------------ formula tokenizer (shared-formula shifting + parser)
  var KINDS = ['str', 'qsheet', 'xsheet', 'sheet', 'err', 'colrng', 'rowrng', 'cell', 'num', 'func', 'bool', 'name', 'ws', 'op'];
  var TOK = /("(?:[^"]|"")*")|('(?:[^']|'')+'!)|(\[\d+\][^!'\s(),;+\-*\/^&=<>:]*!)|([A-Za-z_À-￿0-9][A-Za-z0-9_.À-￿&]*!)|(#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!))|(\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![A-Za-z0-9_(]))|(\$?\d+:\$?\d+(?![0-9.]))|(\$?[A-Z]{1,3}\$?\d+(?![A-Za-z0-9_(]))|((?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?)|([A-Za-z_][A-Za-z0-9_.]*\()|(TRUE|FALSE)(?![A-Za-z0-9_.])|([A-Za-z_\\À-￿][A-Za-z0-9_.\\?À-￿]*)|(\s+)|(<>|<=|>=|[-+*\/^&=<>%(),:;{}])/y;
  function tokenize(f) {
    var out = [], pos = 0, m;
    while (pos < f.length) {
      TOK.lastIndex = pos; m = TOK.exec(f);
      if (!m) throw new Error('Düstur oxunmadı (mövqe ' + pos + ')');
      for (var g = 1; g <= 14; g++) if (m[g] !== undefined) { out.push([KINDS[g - 1], m[g]]); break; }
      pos = TOK.lastIndex;
    }
    return out;
  }
  function shiftCell(t, dr, dc) {
    var m = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/.exec(t);
    return m[1] + (m[1] ? m[2] : colName(colNum(m[2]) + dc)) + m[3] + (m[3] ? m[4] : String(+m[4] + dr));
  }
  function shiftFormula(f, dr, dc) {
    return tokenize(f).map(function (tk) {
      var k = tk[0], t = tk[1];
      if (k === 'cell') return shiftCell(t, dr, dc);
      if (k === 'colrng') return t.split(':').map(function (x) { return x.charAt(0) === '$' ? x : colName(colNum(x) + dc); }).join(':');
      if (k === 'rowrng') return t.split(':').map(function (x) { return x.charAt(0) === '$' ? x : String(+x + dr); }).join(':');
      return t;
    }).join('');
  }

  // ------------------------------------------------------------------ streaming worksheet scanner
  function scanSheet(zip, path, onCell, onTag) {
    var st = zip.stream(path); if (!st) return Promise.resolve();
    var reader = st.pipeThrough(new TextDecoderStream()).getReader(), buf = '';
    var re = /<(c|row|col|pane|sheetFormatPr|tabColor|mergeCell)\b([^>]*?)(\/?)>/g;
    function step() {
      return reader.read().then(function (res) {
        if (res.value) buf += res.value;
        var done = res.done, pos = 0, keep = -1, m;
        re.lastIndex = 0;
        while ((m = re.exec(buf))) {
          if (m[1] === 'c' && !m[3]) {
            var end = buf.indexOf('</c>', re.lastIndex);
            if (end < 0) { keep = m.index; break; }
            onCell(m[2], buf.slice(re.lastIndex, end)); re.lastIndex = end + 4; pos = re.lastIndex; continue;
          }
          if (m[1] === 'c') onCell(m[2], ''); else onTag(m[1], m[2]);
          pos = re.lastIndex;
        }
        if (done) return;
        if (keep >= 0) buf = buf.slice(keep);
        else { var lt = buf.lastIndexOf('<'); buf = buf.slice(lt >= pos ? lt : pos); }
        return step();
      });
    }
    return step();
  }

  // ------------------------------------------------------------------ read one .xlsx into the raw-book format
  function readXlsx(file, bookName, log) {
    var zip, wbRels, ss = [], theme, xfs;
    return openZip(file).then(function (z) {
      zip = z;
      return Promise.all([z.text('xl/workbook.xml'), z.text('xl/_rels/workbook.xml.rels'), z.text('xl/sharedStrings.xml'), z.text('xl/theme/theme1.xml'), z.text('xl/styles.xml')]);
    }).then(function (r) {
      var wb = r[0]; if (!wb) throw new Error('xl/workbook.xml tapılmadı');
      wbRels = rels(r[1]);
      if (r[2]) (r[2].match(/<si\b[^>]*>[\s\S]*?<\/si>|<si\/>/g) || []).forEach(function (si) {
        si = si.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
        var t = '', re = /<t\b[^>]*>([\s\S]*?)<\/t>|<t\/>/g, m; while ((m = re.exec(si))) t += m[1] ? unesc(m[1]) : '';
        ss.push(t);
      });
      theme = parseTheme(r[3]); xfs = r[4] ? parseStyles(r[4], theme) : [{ fmt: 'General', b: false, i: false, fc: null, bg: null, ha: null, ind: 0, wrap: false }];
      var sheetTags = wb.match(tagRe('sheet')) || [], extTags = wb.match(tagRe('externalReference')) || [];
      var book = { name: bookName, kind: 'model', source: 'xlsx', file: file.name, sheets: [], ext: [], xfs: xfs, names: [], arrays: [], placeholders: true,
        sheetNames: sheetTags.map(function (t) { return attrs(t).name; }) };
      // defined names (workbook and sheet scope); Excel's future-function prefixes are removed everywhere
      (wb.match(/<definedName\b[^>]*>[\s\S]*?<\/definedName>/g) || []).forEach(function (d) {
        var da = attrs(/<definedName\b[^>]*>/.exec(d)[0]), txt = unesc(d.replace(/^<definedName\b[^>]*>/, '').replace(/<\/definedName>$/, ''));
        if (!da.name || /^_xlnm\./.test(da.name)) return;
        book.names.push({ name: da.name, scope: da.localSheetId !== undefined ? +da.localSheetId : null, hidden: da.hidden === '1', f: txt.replace(/_xl(?:fn|ws)\./g, '') });
      });
      var chain = Promise.resolve();
      extTags.forEach(function (t) {
        chain = chain.then(function () {
          var p = resolvePath('xl', wbRels[attrs(t)['r:id']]);
          return Promise.all([zip.text(p), zip.text(p.replace('externalLinks/', 'externalLinks/_rels/') + '.rels')]).then(function (x) {
            var lr = rels(x[1]), tg = Object.keys(lr).length ? lr[Object.keys(lr)[0]] : null, xml = x[0] || '';
            var names = (xml.match(tagRe('sheetName')) || []).map(function (s) { return attrs(s).val; }), cache = {};
            var sre = /<sheetData\b([^>]*?)(?:\/>|>([\s\S]*?)<\/sheetData>)/g, sm;
            while ((sm = sre.exec(xml))) {
              var sn = names[+attrs(sm[1]).sheetId], cells = {};
              (sm[2] || '').replace(/<cell\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cell>)/g, function (all, ca, inner) {
                var a = attrs(ca), vm = inner ? /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner) : null, t2 = a.t || 'n', val = null;
                if (vm) { var s = unesc(vm[1]); if (t2 === 's' || t2 === 'str') val = s; else if (t2 === 'b') val = s === '1'; else if (t2 === 'e') val = { e: s }; else { var fv = parseFloat(s); val = isNaN(fv) ? s : fv; } }
                cells[a.r] = val; return '';
              });
              cache[sn] = cells;
            }
            book.ext.push({ target: tg, file: tg ? decodePath(tg) : null, sheets: names, cache: cache });
          });
        });
      });
      sheetTags.forEach(function (t, si) {
        chain = chain.then(function () {
          var arrays = [], a = attrs(t), path = resolvePath('xl', wbRels[a['r:id']]), cells = [], shared = {}, sheet = { name: a.name, state: a.state || 'visible', cells: cells, merges: [], cols: [], hr: [], fz: null, cm: {}, tab: null, dw: 9.14, problems: [] };
          var dflt = 9.14;
          if (log) log('  «' + a.name + '» vərəqi oxunur…');
          return scanSheet(zip, path, function (at, content) {
            var ca = attrs(at), r = ca.r; if (!r) return;
            var fm = content ? /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(content) : null, form = null;
            if (fm) {
              var fa = attrs(fm[1]), ftext = fm[2] !== undefined ? unesc(fm[2]) : '';
              if (fa.t === 'shared') {
                if (ftext) { shared[fa.si] = [r, ftext]; form = ftext; }
                else if (shared[fa.si]) { var m0 = splitRef(shared[fa.si][0]), m1 = splitRef(r); try { form = shiftFormula(shared[fa.si][1], m1[0] - m0[0], m1[1] - m0[1]); } catch (e) { form = null; sheet.problems.push([r, 'paylaşılan düstur açılmadı']); } }
              } else if (fa.t === 'array') {
                var ar = fa.ref ? fa.ref.split(':') : [r];
                if (ar.length < 2 || ar[0] === ar[1]) form = '{ARRAY}' + ftext;
                else { var p1 = splitRef(ar[0]), p2 = splitRef(ar[1]), gid = book.arrays.length; book.arrays.push({ sheet: a.name, r1: p1[0], c1: p1[1], r2: p2[0], c2: p2[1], f: ftext.replace(/_xl(?:fn|ws)\./g, '') }); arrays.push(gid); form = '{AGROUP:' + gid + '}'; }
              } else if (fa.t === 'dataTable') { sheet.problems.push([r, 'məlumat cədvəli (sabit kimi saxlanıldı)']); form = null; }
              else form = ftext;
              if (form) form = form.replace(/_xl(?:fn|ws)\./g, '');
            }
            var t2 = ca.t || 'n', vm = content ? /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(content) : null, val = null;
            if (t2 === 'inlineStr') { var im = /<is>([\s\S]*?)<\/is>/.exec(content || ''); val = ''; if (im) im[1].replace(/<t\b[^>]*>([\s\S]*?)<\/t>/g, function (x, y) { val += unesc(y); return ''; }); }
            else if (vm) {
              var s = unesc(vm[1]);
              if (t2 === 's') val = ss[+s]; else if (t2 === 'str') val = s; else if (t2 === 'b') val = s === '1'; else if (t2 === 'e') val = { e: s };
              else { var fv = parseFloat(s); val = isNaN(fv) ? s : fv; }
            } else if (t2 === 'str') val = '';
            if (val === null && form === null) {   // keep blank input cells (yellow fill / red font) so they can be edited
              var xs0 = xfs[+(ca.s || 0)]; if (!(xs0 && (/^(FFFF00|FFFFCC)$/i.test(xs0.bg || '') || /^FF0000$/i.test(xs0.fc || '')))) return;
              var rc0 = splitRef(r); cells.push({ r: rc0[0], c: rc0[1], v: null, f: null, s: +(ca.s || 0) }); return;
            }
            if (val === '' && form === null) return;
            var rc = splitRef(r);
            cells.push({ r: rc[0], c: rc[1], v: val === undefined ? null : val, f: form, s: +(ca.s || 0) });
          }, function (tag, at) {
            var ta = attrs(at);
            if (tag === 'row') { if (ta.hidden === '1') sheet.hr.push(+ta.r); }
            else if (tag === 'col') sheet.cols.push({ min: +ta.min, max: +ta.max, w: ta.width !== undefined ? parseFloat(ta.width) : dflt, hidden: ta.hidden === '1' });
            else if (tag === 'pane') { if (ta.state === 'frozen' || ta.state === 'frozenSplit') sheet.fz = [Math.floor(parseFloat(ta.ySplit || 0)), Math.floor(parseFloat(ta.xSplit || 0))]; }
            else if (tag === 'sheetFormatPr') { dflt = parseFloat(ta.defaultColWidth || ta.baseColWidth || 8) + 0.7; if (!isFinite(dflt)) dflt = 8.7; sheet.dw = dflt; }
            else if (tag === 'mergeCell') sheet.merges.push(ta.ref);
            else if (tag === 'tabColor') sheet.tab = colorOf('<x ' + at + '/>', theme);
          }).then(function () {
            // members of multi-cell array formulas carry only a cached value: link them to their group
            arrays.forEach(function (gid) {
              var g = book.arrays[gid];
              cells.forEach(function (cd) { if (cd.r >= g.r1 && cd.r <= g.r2 && cd.c >= g.c1 && cd.c <= g.c2 && !(cd.r === g.r1 && cd.c === g.c1) && !cd.f) cd.f = '{AGROUP:' + gid + '}'; });
            });
            var rp = path.split('/'), fn = rp.pop();
            return zip.text(rp.join('/') + '/_rels/' + fn + '.rels').then(function (rx) {
              var rr = rels(rx), cp = null;
              Object.keys(rr).forEach(function (k) { if (/comments/.test(rr[k])) cp = resolvePath(rp.join('/'), rr[k]); });
              return cp ? zip.text(cp) : null;
            }).then(function (cx) {
              if (cx) (cx.match(/<comment\b[^>]*>[\s\S]*?<\/comment>/g) || []).forEach(function (c) {
                var ref = attrs(/<comment\b[^>]*>/.exec(c)[0]).ref, txt = '';
                c.replace(/<t\b[^>]*>([\s\S]*?)<\/t>/g, function (x, y) { txt += unesc(y); return ''; });
                sheet.cm[ref] = txt.trim();
              });
              book.sheets.push(sheet);
              return tick();
            });
          });
        });
      });
      return chain.then(function () { return book; });
    });
  }

  // ------------------------------------------------------------------ raw book reconstructed from the current (compiled) model
  function rawFromCore(core, parts, bi) {
    var N = core.ncells, sh = [], r = [], c = [], v = [], st = [], F = {};
    parts.forEach(function (P) { for (var i = 0; i < P.ids.length; i++) { var id = P.ids[i]; sh[id] = P.sh[i]; r[id] = P.r[i]; c[id] = P.c[i]; v[id] = P.v[i]; st[id] = P.st[i]; } for (var k in P.f) F[k] = P.f[k]; });
    var bk = core.books[bi], sidx = {};
    bk.sheets.forEach(function (gs, i) { sidx[gs] = i; });
    var book = { name: bk.n, kind: bk.kind, source: 'core', sheets: bk.sheets.map(function (gs) { var m = core.sheets[gs]; return { name: m.n, state: m.st || 'visible', cells: [], merges: m.merges || [], cols: m.cols || [], hr: m.hr || [], fz: m.fz || null, cm: m.cm || {}, tab: m.tab || null, dw: m.dw || 9.14, problems: [] }; }), ext: [], xfs: core.styles };
    function q(n) { return "'" + String(n).replace(/'/g, "''") + "'"; }
    function a1(rr, cc, fl) { return (fl.indexOf('C') >= 0 ? '$' : '') + colName(cc) + (fl.indexOf('R') >= 0 ? '$' : '') + rr; }
    function prefix(gs2, explicit, own) {
      var b2 = core.books[core.sheets[gs2].b].n;
      if (b2 !== own) return q('[' + b2 + ']' + core.sheets[gs2].n) + '!';
      return explicit || true ? q(core.sheets[gs2].n) + '!' : '';
    }
    for (var id = 0; id < N; id++) {
      if (sidx[sh[id]] === undefined) continue;
      var f = F[id], text = null;
      if (f) {
        var d = core.disp[f[0]], refs = f[1], kb = f[2], li = 0, ri = 0, out = '', i2 = 0;
        while (i2 < d.length) {
          var ch = d.charAt(i2);
          if (ch === '\x01') { out += String(core.K[kb + li]); li++; i2++; continue; }
          if (ch === '\x02') {
            var j = d.indexOf('\x03', i2), fl = d.slice(i2 + 1, j), ref = refs[ri++];
            if (fl.charAt(0) === 'P') fl = fl.slice(1);
            if (ref >= 0) out += prefix(sh[ref], true, bk.n) + a1(r[ref], c[ref], fl);
            else { var R = core.ranges[-ref - 1], fs = fl.slice(1).split(':'); out += prefix(R[0], true, bk.n) + a1(R[1], R[2], fs[0] || '') + ':' + a1(R[3], R[4], fs[1] || ''); }
            i2 = j + 1; continue;
          }
          var k2 = i2; while (k2 < d.length && d.charAt(k2) !== '\x01' && d.charAt(k2) !== '\x02') k2++;
          out += d.slice(i2, k2); i2 = k2;
        }
        text = out;
      }
      if (!f && (v[id] === null || v[id] === undefined)) continue;   // virtual empty cell
      book.sheets[sidx[sh[id]]].cells.push({ r: r[id], c: c[id], v: v[id] === undefined ? null : v[id], f: text, s: st[id] });
    }
    return book;
  }

  // ------------------------------------------------------------------ parser
  function Parser(toks) { this.t = toks.filter(function (x) { return x[0] !== 'ws'; }); this.i = 0; }
  Parser.prototype.peek = function (k) { var j = this.i + (k || 0); return j < this.t.length ? this.t[j] : [null, null]; };
  Parser.prototype.next = function () { return this.t[this.i++]; };
  Parser.prototype.parse = function () { var e = this.comparison(); if (this.i !== this.t.length) throw new Error('artıq simvollar'); return e; };
  Parser.prototype.comparison = function () { var a = this.concat(); while (['=', '<>', '<', '>', '<=', '>='].indexOf(this.peek()[1]) >= 0) { var op = this.next()[1]; a = ['bin', op, a, this.concat()]; } return a; };
  Parser.prototype.concat = function () { var a = this.additive(); while (this.peek()[1] === '&') { this.next(); a = ['bin', '&', a, this.additive()]; } return a; };
  Parser.prototype.additive = function () { var a = this.mult(); while ((this.peek()[1] === '+' || this.peek()[1] === '-') && this.peek()[0] === 'op') { var op = this.next()[1]; a = ['bin', op, a, this.mult()]; } return a; };
  Parser.prototype.mult = function () { var a = this.power(); while (this.peek()[1] === '*' || this.peek()[1] === '/') { var op = this.next()[1]; a = ['bin', op, a, this.power()]; } return a; };
  Parser.prototype.power = function () { var a = this.unary(); while (this.peek()[1] === '^') { this.next(); a = ['bin', '^', a, this.unary()]; } return a; };
  Parser.prototype.unary = function () {
    if (this.peek()[1] === '-' && this.peek()[0] === 'op') { this.next(); return ['neg', this.unary()]; }
    if (this.peek()[1] === '+' && this.peek()[0] === 'op') { this.next(); return this.unary(); }
    return this.postfix();
  };
  Parser.prototype.postfix = function () {
    var a = this.primary();
    while (this.peek()[1] === ':' && this.peek()[0] === 'op') { this.next(); a = ['rngop', a, this.primary()]; }
    while (this.peek()[1] === '%') { this.next(); a = ['pct', a]; }
    return a;
  };
  Parser.prototype.primary = function () {
    var p = this.peek(), k = p[0], t = p[1];
    if (k === 'num') { this.next(); return ['num', parseFloat(t)]; }
    if (k === 'str') { this.next(); return ['str', t.slice(1, -1).replace(/""/g, '"')]; }
    if (k === 'bool') { this.next(); return ['bool', t === 'TRUE']; }
    if (k === 'err') { this.next(); return ['err', t, t]; }
    if (k === 'func') {
      this.next(); var name = t.slice(0, -1).toUpperCase(), args = [];
      if (this.peek()[1] === ')') { this.next(); return ['call', name, args]; }
      for (;;) {
        if (this.peek()[1] === ',' || this.peek()[1] === ')') args.push(['missing']); else args.push(this.comparison());
        var t2 = this.next(); if (!t2) throw new Error('mötərizə bağlanmayıb');
        if (t2[1] === ')') break; if (t2[1] !== ',') throw new Error('arqument ayırıcısı səhvdir');
      }
      return ['call', name, args];
    }
    if (t === '(') { this.next(); var e = this.comparison(); var cl = this.next(); if (!cl || cl[1] !== ')') throw new Error('mötərizə bağlanmayıb'); return ['paren', e]; }
    if (t === '{' && k === 'op') {
      this.next(); var rows = [[]];
      for (;;) {
        var q = this.peek(), sign = 1;
        if (q[1] === '-') { this.next(); sign = -1; q = this.peek(); }
        this.next();
        if (q[0] === 'num') rows[rows.length - 1].push(sign * parseFloat(q[1]));
        else if (q[0] === 'str') rows[rows.length - 1].push(q[1].slice(1, -1).replace(/""/g, '"'));
        else if (q[0] === 'bool') rows[rows.length - 1].push(q[1] === 'TRUE');
        else if (q[0] === 'err') rows[rows.length - 1].push({ e: q[1] });
        else throw new Error('massiv sabiti');
        var sep = this.next()[1];
        if (sep === '}') break; if (sep === ';') rows.push([]); else if (sep !== ',') throw new Error('massiv sabiti');
      }
      return ['aconst', rows];
    }
    var prefix = null;
    if (k === 'qsheet' || k === 'xsheet' || k === 'sheet') {
      this.next(); prefix = t; p = this.peek(); k = p[0]; t = p[1];
      if (k === 'err') { this.next(); return ['err', t, prefix + t]; }
    }
    if (k === 'cell') {
      this.next();
      if (this.peek()[1] === ':' && this.peek(1)[0] === 'cell') { this.next(); var t3 = this.next()[1]; return ['rng', prefix, t, t3]; }
      return ['ref', prefix, t];
    }
    if (k === 'name') { this.next(); return ['nm', prefix, t]; }
    if (k === 'colrng') { this.next(); var ab = t.split(':'); return ['rng', prefix, ab[0] + '1', ab[1] + MAXROW, 'col', t]; }
    if (k === 'rowrng') { this.next(); return ['rng', prefix, null, null, 'row', t]; }
    throw new Error('gözlənilməyən simvol ' + t);
  };

  // ------------------------------------------------------------------ compile raw books into engine data
  var RANGE_FUNCS = { SUM: 1, AVERAGE: 1, MIN: 1, MAX: 1, AND: 1, OR: 1, SUMPRODUCT: 1, CORREL: 1, COUNT: 1, COUNTA: 1, PRODUCT: 1, STDEV: 1, 'STDEV.S': 1,
    'STDEV.P': 1, STDEVP: 1, VAR: 1, VARP: 1, MEDIAN: 1, SLOPE: 1, INTERCEPT: 1, TRANSPOSE: 1, MMULT: 1, MINVERSE: 1, MDETERM: 1, LINEST: 1, HPF: 1, HPP: 1 };
  var SCALAR_FUNCS = { IF: 1, LN: 1, EXP: 1, ROUND: 1, TEXT: 1, SQRT: 1, ABS: 1, INT: 1, MOD: 1, POWER: 1, LOG: 1, LOG10: 1, SIGN: 1, ROUNDUP: 1, ROUNDDOWN: 1,
    NOT: 1, CONCATENATE: 1, NA: 1, TRUE: 1, FALSE: 1, MUNIT: 1 };
  var ELEM1 = { LN: 1, EXP: 1, SQRT: 1, ABS: 1, INT: 1, SIGN: 1, NOT: 1, LOG10: 1 }, ELEM2 = { ROUND: 1, MOD: 1, POWER: 1, LOG: 1, TEXT: 1, ROUNDUP: 1, ROUNDDOWN: 1 };
  var THUNK_FUNCS = { IFERROR: 1, IFNA: 1, ISNUMBER: 1, ISERROR: 1, ISNA: 1, ISERR: 1, ISTEXT: 1, ISBLANK: 1, ISLOGICAL: 1, CHOOSE: 1 };
  var LOOKUP_FUNCS = { INDEX: 1, MATCH: 1, VLOOKUP: 1, HLOOKUP: 1 };
  var BINFN = { '+': 'add', '-': 'sub', '*': 'mul', '/': 'div', '^': 'pow', '&': 'cat', '=': 'eq', '<>': 'ne', '<': 'lt', '>': 'gt', '<=': 'le', '>=': 'ge' };
  function flagsOf(t) { var m = /^(\$?)[A-Z]{1,3}(\$?)\d+$/.exec(t); return (m[1] ? 'C' : '') + (m[2] ? 'R' : ''); }
  function cellParts(t) { var m = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/.exec(t); return [+m[4], colNum(m[2])]; }

  function compile(raws, log) {
    var books = [], bookIdx = {}, sheets = [], sheetIdx = {}, SH = [], R = [], C = [], V = [], FT = [], ST = [], VIRT = [];
    var styles = [], styleKey = {}, problems = [], rawOf = {};
    function styleOf(x) {
      var o = x || { fmt: 'General', b: false, i: false, fc: null, bg: null, ha: null, ind: 0, wrap: false };
      var k = [o.fmt, o.b, o.i, o.fc, o.bg, o.ha, o.ind, o.wrap].join('');
      if (styleKey[k] === undefined) { styleKey[k] = styles.length; styles.push({ fmt: o.fmt, b: !!o.b, i: !!o.i, fc: o.fc || null, bg: o.bg || null, ha: o.ha || null, ind: o.ind || 0, wrap: !!o.wrap }); }
      return styleKey[k];
    }
    styleOf(null);
    function addBook(name, kind) { bookIdx[name] = books.length; books.push({ n: name, kind: kind, sheets: [] }); }
    function addSheet(bname, sm) {
      var gs = sheets.length; sheetIdx[bname + '' + sm.name] = gs; books[bookIdx[bname]].sheets.push(gs);
      sheets.push({ b: bookIdx[bname], n: sm.name, st: sm.state, merges: sm.merges, cols: sm.cols, hr: sm.hr, fz: sm.fz, cm: sm.cm, tab: sm.tab, dw: sm.dw, map: new Map(), maxR: 0, maxC: 0 });
      return gs;
    }
    function addCell(gs, r, c, v, f, s, virt) {
      var id = SH.length; SH.push(gs); R.push(r); C.push(c); V.push(v); FT.push(f); ST.push(s); VIRT.push(!!virt);
      var sm = sheets[gs]; sm.map.set(r * 20000 + c, id); if (r > sm.maxR) sm.maxR = r; if (c > sm.maxC) sm.maxC = c;
      return id;
    }
    raws.forEach(function (rb) {
      addBook(rb.name, rb.kind); rawOf[rb.name] = rb;
      rb.sheets.forEach(function (sm) {
        var gs = addSheet(rb.name, sm);
        (sm.problems || []).forEach(function (p) { problems.push({ where: rb.name + ' › ' + sm.name + '!' + p[0], what: p[1] }); });
        sm.cells.forEach(function (cd) { addCell(gs, cd.r, cd.c, cd.v, cd.f, styleOf(rb.xfs[cd.s]), false); });
      });
    });
    // download-style sheets (constant calendar years in a header row): create empty cells for the missing years of every
    // indicator row, so realised data can later be typed into a year that has no value yet
    raws.forEach(function (rb) {
      if (rb.kind !== 'model' || !rb.placeholders) return;
      rb.sheets.forEach(function (sm) {
        var gs = sheetIdx[rb.name + '\x01' + sm.name], m = sheets[gs].map, byRow = {};
        sm.cells.forEach(function (cd) { (byRow[cd.r] || (byRow[cd.r] = [])).push(cd); });
        var hdr = null;
        Object.keys(byRow).map(Number).sort(function (a, b) { return a - b; }).some(function (r) {
          var ys = byRow[r].filter(function (cd) { return (!cd.f || sm.name === 'Data') && typeof cd.v === 'number' && cd.v === Math.floor(cd.v) && cd.v >= 1980 && cd.v <= 2060; }).sort(function (a, b) { return a.c - b.c; });
          var run = 0; for (var i = 1; i < ys.length; i++) if (ys[i].c === ys[i - 1].c + 1 && ys[i].v === ys[i - 1].v + 1) run++;
          if (run >= 10) { hdr = { r: r, c1: ys[0].c, c2: ys[ys.length - 1].c }; return true; }
          return false;
        });
        if (!hdr) return;
        Object.keys(byRow).map(Number).forEach(function (r) {
          if (r <= hdr.r) return;
          var cs = byRow[r], lab = cs.some(function (cd) { return cd.c < hdr.c1 && typeof cd.v === 'string' && cd.v.trim(); });
          var nums = cs.filter(function (cd) { return cd.c >= hdr.c1 && cd.c <= hdr.c2 && (typeof cd.v === 'number' || cd.f); });
          if (!lab || !nums.length) return;
          // the 'Data' sheet (year columns follow INPUT!B9): only rows typed by hand, never formula rows
          if (sm.name === 'Data' && nums.some(function (cd) { return cd.f; })) return;
          var st = nums[nums.length - 1].s;
          for (var c = hdr.c1; c <= hdr.c2; c++) if (!m.has(r * 20000 + c)) addCell(gs, r, c, null, null, styleOf(rb.xfs[st]), false);
        });
      });
    });
    // input cells known from the inputs catalogue that are blank in the file (e.g. empty shock paths): create them empty
    raws.forEach(function (rb) {
      (rb.extraCells || []).forEach(function (x) {
        var gs = sheetIdx[rb.name + '\x01' + x[0]]; if (gs === undefined) return;
        if (!sheets[gs].map.has(x[1] * 20000 + x[2])) addCell(gs, x[1], x[2], null, null, 0, false);
      });
    });
    var modelNames = raws.filter(function (b) { return b.kind === 'model'; }).map(function (b) { return b.name; });
    var FILE2BOOK = {};
    modelNames.forEach(function (n) { FILE2BOOK[norm(n)] = n; });
    raws.forEach(function (b) { if (b.file) FILE2BOOK[norm(b.file)] = b.name; });
    function extBookFor(L) {
      var fname = L.file || '?', mb = FILE2BOOK[norm(fname)];
      if (mb) return [mb, 'model'];
      var bn = 'EXT: ' + fname; if (bookIdx[bn] === undefined) addBook(bn, 'ext'); return [bn, 'ext'];
    }
    function getOrVirtual(bname, sname, r, c, L) {
      var key = bname + '' + sname, kind = books[bookIdx[bname]].kind;
      if (sheetIdx[key] === undefined) {
        if (kind === 'model') return null;
        addSheet(bname, { name: sname, state: 'visible', merges: [], cols: [], hr: [], fz: null, cm: {}, tab: null, dw: 9.14 });
      }
      var gs = sheetIdx[key], x = sheets[gs].map.get(r * 20000 + c);
      if (kind === 'ext' && L) {
        var cv = (L.cache[sname] || {})[colName(c) + r];
        if (x !== undefined) { if (cv !== undefined && !VIRT[x]) V[x] = cv; return x; }
        return addCell(gs, r, c, cv === undefined ? null : cv, null, 0, false);
      }
      if (x !== undefined) return x;
      return addCell(gs, r, c, null, null, 0, kind === 'model');
    }
    function parsePrefix(prefix) {
      var p = prefix.slice(0, -1);
      if (p.charAt(0) === "'") p = p.slice(1, -1).replace(/''/g, "'");
      var m = /^\[([^\]]+)\](.*)$/.exec(p);
      if (m) return [m[1], m[2]];
      return [null, p];
    }
    var ranges = [], rangeIdx = {}, templates = {}, codes = [], disp = [], K = [], cellF = {}, deps = {}, xdeps = {};
    var nModelCells = SH.length;
    function targetOf(ctx, prefix) {
      if (!prefix) return [ctx.book, ctx.sheet, null];
      var pp = parsePrefix(prefix), link = pp[0], sname = pp[1];
      if (link === null) return [ctx.book, sname, null];
      if (/^\d+$/.test(link)) {
        var L = rawOf[ctx.book].ext ? rawOf[ctx.book].ext[+link - 1] : null;
        if (!L) return [null, sname, null];
        var eb = extBookFor(L); return [eb[0], sname, eb[1] === 'ext' ? L : null, L];
      }
      if (bookIdx[link] === undefined) return [null, sname, null];
      return [link, sname, null];
    }
    function resolveCell(ctx, prefix, t) {
      var tg = targetOf(ctx, prefix); if (!tg[0]) { ctx.problem = 'naməlum fayl istinadı'; return null; }
      var rc = cellParts(t);
      if (sheetIdx[tg[0] + '' + tg[1]] === undefined && books[bookIdx[tg[0]]].kind === 'model') { problems.push({ where: ctx.where, what: 'mövcud olmayan vərəqə istinad: ' + tg[0] + ' › ' + tg[1] }); return null; }
      var id = getOrVirtual(tg[0], tg[1], rc[0], rc[1], tg[2]);
      if (id !== null && tg[3] && books[bookIdx[tg[0]]].kind === 'model') ctx.xlinks.push([id, tg[3], tg[1], colName(rc[1]) + rc[0]]);
      return id;
    }
    function resolveRange(ctx, prefix, a, b, kind, rawT) {
      var tg = targetOf(ctx, prefix); if (!tg[0]) return null;
      var key = tg[0] + '' + tg[1];
      if (sheetIdx[key] === undefined) {
        if (books[bookIdx[tg[0]]].kind === 'model') { problems.push({ where: ctx.where, what: 'mövcud olmayan vərəqə istinad: ' + tg[0] + ' › ' + tg[1] }); return null; }
        getOrVirtual(tg[0], tg[1], 1, 1, tg[2]);
      }
      var gs = sheetIdx[key], r1, c1, r2, c2;
      if (kind === 'row') { var xy = rawT.replace(/\$/g, '').split(':'); r1 = +xy[0]; r2 = +xy[1]; c1 = 1; c2 = MAXCOL; }
      else { var p1 = cellParts(a), p2 = cellParts(b); r1 = Math.min(p1[0], p2[0]); r2 = Math.max(p1[0], p2[0]); c1 = Math.min(p1[1], p2[1]); c2 = Math.max(p1[1], p2[1]); }
      if (kind) { r2 = Math.min(r2, Math.max(sheets[gs].maxR, r1)); c2 = Math.min(c2, Math.max(sheets[gs].maxC, c1)); }
      if (tg[2]) Object.keys(tg[2].cache[tg[1]] || {}).forEach(function (addr) { var rc = splitRef(addr); if (rc && rc[0] >= r1 && rc[0] <= r2 && rc[1] >= c1 && rc[1] <= c2) getOrVirtual(tg[0], tg[1], rc[0], rc[1], tg[2]); });
      var k = [gs, r1, c1, r2, c2].join(',');
      if (rangeIdx[k] === undefined) { rangeIdx[k] = ranges.length; ranges.push([gs, r1, c1, r2, c2]); }
      if (tg[3] && books[bookIdx[tg[0]]].kind === 'model') { var m = sheets[gs].map; for (var rr = r1; rr <= r2; rr++) for (var cc = c1; cc <= c2; cc++) { var x = m.get(rr * 20000 + cc); if (x !== undefined) ctx.xlinks.push([x, tg[3], tg[1], colName(cc) + rr]); } }
      return rangeIdx[k];
    }
    function rangeMembers(gi) {
      var Rg = ranges[gi], m = sheets[Rg[0]].map, out = [];
      if ((Rg[3] - Rg[1] + 1) * (Rg[4] - Rg[2] + 1) <= m.size) { for (var r = Rg[1]; r <= Rg[3]; r++) for (var c = Rg[2]; c <= Rg[4]; c++) { var x = m.get(r * 20000 + c); if (x !== undefined) out.push(x); } }
      else m.forEach(function (x, key) { var r = Math.floor(key / 20000), c = key % 20000; if (r >= Rg[1] && r <= Rg[3] && c >= Rg[2] && c <= Rg[4]) out.push(x); });
      return out;
    }
    // names: workbook-level and sheet-level defined names (inlined; the display keeps the name)
    var nameDefs = {}, nameAst = {};
    raws.forEach(function (rb) {
      (rb.names || []).forEach(function (nm) {
        var scope = nm.scope === null || nm.scope === undefined ? '' : (rb.sheetNames || [])[nm.scope];
        if (scope === undefined) return;
        nameDefs[rb.name + '|' + scope + '|' + nm.name.toLowerCase()] = nm;
      });
    });
    function findName(ctx, prefix, text) {
      var lc = text.toLowerCase();
      if (prefix) { var pp = parsePrefix(prefix); if (pp[0] === null) return nameDefs[ctx.book + '|' + pp[1] + '|' + lc] || null; return null; }
      return nameDefs[ctx.book + '|' + ctx.sheet + '|' + lc] || nameDefs[ctx.book + '||' + lc] || null;
    }
    function nameTree(nm) {
      var key = nm.name + '|' + nm.scope;
      if (nameAst[key] === undefined) nameAst[key] = nm.f === null || nm.f === undefined ? null : new Parser(tokenize(nm.f)).parse();
      return nameAst[key];
    }
    var MARK = String.fromCharCode(5);
    function pushRef(ctx, x) {
      if (ctx.inName) { ctx.nrefs.push(x); return MARK + 'R' + (ctx.nrefs.length - 1) + MARK; }
      ctx.refs.push(x); return String(ctx.refs.length - 1);
    }
    function pushLit(ctx, v) {
      if (ctx.inName) { ctx.nlits.push(v); return MARK + 'K' + (ctx.nlits.length - 1) + MARK; }
      ctx.lits.push(v); return String(ctx.lits.length - 1);
    }
    function thunk(code) { return 'function(){return ' + code + ';}'; }
    function kindOf(n, ctx) {                       // syntactic kind, looking through names and parentheses
      var guard = 0;
      while (n && (n[0] === 'paren' || n[0] === 'nm') && guard++ < 20) {
        if (n[0] === 'paren') n = n[1];
        else { var nm = findName(ctx, n[1], n[2]); n = nm ? nameTree(nm) : null; }
      }
      return n ? n[0] : 'err';
    }
    var P1 = String.fromCharCode(1), P2 = String.fromCharCode(2), P3 = String.fromCharCode(3);
    function gen(n, ctx, tpl, am, asRef) {
      var k = n[0], c, a, b, idx;
      if (k === 'num') { idx = pushLit(ctx, n[1]); tpl.push(P1); return 'K[k+' + idx + ']'; }
      if (k === 'str') { tpl.push('"' + n[1].replace(/"/g, '""') + '"'); return JSON.stringify(n[1]); }
      if (k === 'bool') { tpl.push(n[1] ? 'TRUE' : 'FALSE'); return n[1] ? 'true' : 'false'; }
      if (k === 'err') { tpl.push(n[2]); return 'X.E(' + JSON.stringify(n[1]) + ')'; }
      if (k === 'missing') return 'null';
      if (k === 'aconst') {
        tpl.push('{' + n[1].map(function (r) { return r.map(function (v) { return typeof v === 'string' ? '"' + v + '"' : v && v.e ? v.e : String(v).toUpperCase(); }).join(','); }).join(';') + '}');
        var flat = [].concat.apply([], n[1]).map(function (v) { return v && v.e ? 'X.E(' + JSON.stringify(v.e) + ')' : JSON.stringify(v); });
        return 'new X.Arr(' + n[1].length + ',' + n[1][0].length + ',[' + flat.join(',') + '])';
      }
      if (k === 'paren') { tpl.push('('); c = gen(n[1], ctx, tpl, am, asRef); tpl.push(')'); return '(' + c + ')'; }
      if (k === 'neg') { tpl.push('-'); return (am ? 'X.A.neg(' : 'X.neg(') + gen(n[1], ctx, tpl, am) + ')'; }
      if (k === 'pct') { c = gen(n[1], ctx, tpl, am); tpl.push('%'); return (am ? 'X.A.pct(' : 'X.pct(') + c + ')'; }
      if (k === 'bin') { a = gen(n[2], ctx, tpl, am); tpl.push(n[1]); b = gen(n[3], ctx, tpl, am); return (am ? 'X.A.' : 'X.') + BINFN[n[1]] + '(' + a + ',' + b + ')'; }
      if (k === 'nm') {
        var nm = findName(ctx, n[1], n[2]);
        if (!nm) throw new Error('naməlum ad ' + n[2]);
        tpl.push((n[1] || '') + n[2]);
        var tree = nameTree(nm);
        if (!tree) return 'X.E("#NAME?")';
        if ((ctx.nameDepth = (ctx.nameDepth || 0) + 1) > 20) throw new Error('ad zənciri çox dərindir');
        var wasIn = ctx.inName; ctx.inName = true;
        var code = gen(tree, ctx, [], am, asRef);
        ctx.inName = wasIn; ctx.nameDepth--;
        return '(' + code + ')';
      }
      if (k === 'ref') {
        var cid = resolveCell(ctx, n[1], n[2]);
        if (cid === null) { tpl.push((n[1] || '') + '#REF!'); return 'X.E("#REF!")'; }
        idx = pushRef(ctx, cid); ctx.deps.add(cid);
        tpl.push(P2 + (n[1] ? 'P' : '') + flagsOf(n[2]) + P3);
        return asRef ? 'X.REFOF(r[' + idx + '])' : 'V[r[' + idx + ']]';
      }
      if (k === 'rng') {
        var gi, f1 = '', f2 = '';
        if (n.length > 4 && (n[4] === 'col' || n[4] === 'row')) gi = resolveRange(ctx, n[1], n[2], n[3], n[4], n[5]);
        else { gi = resolveRange(ctx, n[1], n[2], n[3]); f1 = flagsOf(n[2]); f2 = flagsOf(n[3]); }
        if (gi === null) { tpl.push((n[1] || '') + '#REF!'); return 'X.E("#REF!")'; }
        idx = pushRef(ctx, -(gi + 1));
        rangeMembers(gi).forEach(function (m) { ctx.deps.add(m); });
        tpl.push(P2 + (n[1] ? 'P' : '') + 'G' + f1 + ':' + f2 + P3);
        return 'G(r[' + idx + '])';
      }
      if (k === 'rngop') {
        a = gen(n[1], ctx, tpl, am, true); tpl.push(':'); b = gen(n[2], ctx, tpl, am, true);
        return 'X.RNG(' + a + ',' + b + ')';
      }
      if (k === 'call') {
        var name = n[1], args = n[2], cs = [];
        if (!RANGE_FUNCS[name] && !SCALAR_FUNCS[name] && !THUNK_FUNCS[name] && !LOOKUP_FUNCS[name]) throw new Error('dəstəklənməyən funksiya ' + name);
        tpl.push(name + '(');
        args.forEach(function (arg, i) {
          if (i) tpl.push(',');
          var ak = kindOf(arg, ctx), rangeArg = RANGE_FUNCS[name] || (LOOKUP_FUNCS[name] && ((name === 'INDEX' && i === 0) || (name !== 'INDEX' && i === 1)));
          if (rangeArg && (ak === 'ref' || ak === 'rng' || ak === 'rngop')) {
            var useRef = ak === 'rngop' || (ak === 'ref' && !!LOOKUP_FUNCS[name]) || (ak === 'ref' && asRef);
            var cc = gen(arg, ctx, tpl, am, useRef);
            if (ak === 'ref' && RANGE_FUNCS[name] && !useRef && cc.indexOf('X.E') < 0) {
              var ri = ctx.inName ? MARK + 'R' + (ctx.nrefs.length - 1) + MARK : String(ctx.refs.length - 1);
              cc = 'X.R1(r[' + ri + '])';
            }
            cs.push(cc);
          } else {
            var code2 = gen(arg, ctx, tpl, am || (!!RANGE_FUNCS[name] && name !== 'AND' && name !== 'OR'));
            cs.push(THUNK_FUNCS[name] ? (name === 'CHOOSE' && i === 0 ? code2 : thunk(code2)) : code2);
          }
        });
        tpl.push(')');
        if (name === 'IF') {
          while (cs.length < 3) cs.push(cs.length === 2 ? 'false' : 'null');
          if (args.length === 3 && args[2][0] === 'missing') cs[2] = '0';
          if (args.length >= 2 && args[1][0] === 'missing') cs[1] = '0';
          if (am) return 'X.A.IF(' + cs[0] + ',' + thunk(cs[1]) + ',' + thunk(cs[2]) + ')';
          return '(X.b(' + cs[0] + ')?' + cs[1] + ':' + cs[2] + ')';
        }
        if (name === 'INDEX') return (asRef ? 'X.INDEXREF(' : 'X.INDEX(') + [cs[0], cs[1] || 'null', cs[2] || 'null', args.length].join(',') + ')';
        if (am && ELEM1[name]) return 'X.A.map(X.' + name + ',' + cs[0] + ')';
        if (am && ELEM2[name]) return 'X.A.map2(X.' + name + ',' + cs[0] + ',' + (cs[1] || 'null') + ')';
        return 'X[' + JSON.stringify(name) + '](' + cs.join(',') + ')';
      }
      throw new Error('node ' + k);
    }
    function finishCode(code, ctx) {
      var nr = ctx.refs.length, nl = ctx.lits.length, re = new RegExp(MARK + '([RK])(\\d+)' + MARK, 'g');
      code = code.replace(re, function (m, t, i) { return String((t === 'R' ? nr : nl) + +i); });
      ctx.refs = ctx.refs.concat(ctx.nrefs); ctx.lits = ctx.lits.concat(ctx.nlits);
      return code;
    }
    // multi-cell array formulas: one hidden group cell per array (on a hidden sheet), members read their element
    var agroups = [];
    raws.forEach(function (rb) {
      if (!rb.arrays || !rb.arrays.length) return;
      var base = agroups.length;
      var gsA = addSheet(rb.name, { name: '__massivlər__', state: 'veryHidden', merges: [], cols: [], hr: [], fz: null, cm: {}, tab: null, dw: 9.14 });
      rb.arrays.forEach(function (ag) {
        var vid = addCell(gsA, agroups.length + 1, 1, null, { group: agroups.length, f: ag.f, sheet: ag.sheet }, 0, false);
        var gsS = sheetIdx[rb.name + '\u0001' + ag.sheet];
        agroups.push([gsS, ag.r1, ag.c1, ag.r2, ag.c2, vid]);
      });
      rb.sheets.forEach(function (sm) {
        var gs2 = sheetIdx[rb.name + '\u0001' + sm.name];
        sm.cells.forEach(function (cd) {
          if (typeof cd.f === 'string' && cd.f.indexOf('{AGROUP:') === 0) {
            var id = sheets[gs2].map.get(cd.r * 20000 + cd.c), gidx = +cd.f.slice(8, -1);
            FT[id] = { member: base + gidx };
          }
        });
      });
    });
    var nForm = 0;
    for (var cid = 0; cid < SH.length; cid++) {
      if (FT[cid] === null || FT[cid] === undefined) continue;
      var sm = sheets[SH[cid]], bname = books[sm.b].n, f0 = FT[cid];
      var ctx = { book: bname, sheet: sm.n, refs: [], lits: [], nrefs: [], nlits: [], deps: new Set(), xlinks: [], where: bname + ' › ' + sm.n + '!' + colName(C[cid]) + R[cid] };
      var tpl = [], code;
      try {
        if (typeof f0 === 'object' && f0.member !== undefined) {
          var vid2 = agroups[f0.member][5]; ctx.refs.push(vid2); ctx.deps.add(vid2); tpl.push('{AEL}'); code = 'X.AEL(r[0])';
        } else if (typeof f0 === 'object' && f0.group !== undefined) {
          ctx.sheet = f0.sheet; ctx.where = bname + ' › ' + f0.sheet + ' (massiv ' + (f0.group + 1) + ')';
          tpl.push('{'); code = 'X.GROUPV(' + gen(new Parser(tokenize(f0.f)).parse(), ctx, tpl, true) + ')'; tpl.push('}');
        } else if (f0.indexOf('{ARRAY}') === 0) {
          tpl.push('{'); code = 'X.A1(' + gen(new Parser(tokenize(f0.slice(7))).parse(), ctx, tpl, true) + ')'; tpl.push('}');
        } else code = gen(new Parser(tokenize(f0)).parse(), ctx, tpl, false);
        code = finishCode(code, ctx);
      }
      catch (e) {
        var ftxt = typeof f0 === 'object' ? (f0.f || '') : f0;
        problems.push({ where: ctx.where, what: 'düstur hesablanmır (' + e.message + '), Excel dəyəri saxlanıldı: =' + String(ftxt).slice(0, 120) });
        FT[cid] = null; continue;
      }
      var dkey = tpl.join(''), key = dkey + '\n' + code;
      if (templates[key] === undefined) { templates[key] = codes.length; codes.push(code); disp.push(dkey); }
      var kb = K.length; Array.prototype.push.apply(K, ctx.lits);
      cellF[cid] = [templates[key], ctx.refs, kb];
      deps[cid] = Array.from(ctx.deps);
      if (ctx.xlinks.length) xdeps[cid] = ctx.xlinks;
      nForm++;
    }
    // evaluation order: strongly connected components (Tarjan, iterative) in dependency order.
    // Blocks with a cycle (e.g. a regime switch that breaks the loop only dynamically) are solved by repeated sweeps.
    var N = SH.length, isF = new Uint8Array(N), order = [], cyc = [];
    var fids = Object.keys(cellF).map(Number);
    fids.forEach(function (id) { isF[id] = 1; });
    var fdeps = {}; fids.forEach(function (id) { fdeps[id] = deps[id].filter(function (p) { return isF[p]; }); });
    var index = new Int32Array(N).fill(-1), low = new Int32Array(N), onSt = new Uint8Array(N), st = [], cnt = 0;
    fids.forEach(function (s0) {
      if (index[s0] >= 0) return;
      var work = [[s0, 0]]; index[s0] = low[s0] = cnt++; st.push(s0); onSt[s0] = 1;
      while (work.length) {
        var top = work[work.length - 1], v = top[0], ds = fdeps[v];
        if (top[1] < ds.length) {
          var w = ds[top[1]++];
          if (index[w] < 0) { index[w] = low[w] = cnt++; st.push(w); onSt[w] = 1; work.push([w, 0]); }
          else if (onSt[w] && index[w] < low[v]) low[v] = index[w];
          continue;
        }
        work.pop();
        if (work.length) { var u = work[work.length - 1][0]; if (low[v] < low[u]) low[u] = low[v]; }
        if (low[v] === index[v]) {
          var comp = [], x2;
          do { x2 = st.pop(); onSt[x2] = 0; comp.push(x2); } while (x2 !== v);
          if (comp.length === 1 && fdeps[v].indexOf(v) < 0) { order.push(v); continue; }
          // inside the block: DFS post-order over precedents, so most values are fresh in the first sweep
          var inC = new Set(comp), done = new Set(), start = order.length;
          comp.slice().reverse().forEach(function (c0) {
            if (done.has(c0)) return; var wk = [[c0, 0]]; done.add(c0);
            while (wk.length) {
              var t2 = wk[wk.length - 1], d2 = fdeps[t2[0]];
              if (t2[1] < d2.length) { var y = d2[t2[1]++]; if (inC.has(y) && !done.has(y)) { done.add(y); wk.push([y, 0]); } continue; }
              wk.pop(); order.push(t2[0]);
            }
          });
          cyc.push([start, order.length]);
        }
      }
    });
    if (cyc.length) {
      var big = cyc.map(function (b) { return b[1] - b[0]; }).reduce(function (a, b) { return a + b; }, 0);
      problems.push({ where: 'model', what: 'dövri asılılıq blokları: ' + cyc.length + ' blok, ' + big + ' xana — ardıcıl yaxınlaşma ilə hesablanır (keçid açarları dövrü dinamik qırır)' });
    }
    var sheetOut = sheets.map(function (s) { return { b: s.b, n: s.n, st: s.st, merges: s.merges, cols: s.cols, hr: s.hr, fz: s.fz, cm: s.cm, tab: s.tab, dw: s.dw }; });
    var namesOut = [];
    // only names that resolve to a cell or range on a real sheet (workbooks carry many legacy junk names: #REF!, array constants)
    var NREF = /^(?:'(?:[^'\[\]]|'')+'|[^'!\[\]]+)!\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?$/;
    raws.forEach(function (rb) { (rb.names || []).forEach(function (nm) { if (nm.f && (nm.scope === null || nm.scope === undefined) && NREF.test(nm.f)) namesOut.push({ b: rb.name, n: nm.name, f: nm.f }); }); });
    var core = { books: books, sheets: sheetOut, styles: styles, disp: disp, ranges: ranges, K: K, order: order, ncells: N, files: [], tplCode: codes, agroups: agroups, names: namesOut, cyc: cyc };
    var P = { ids: [], sh: SH, r: R, c: C, v: V.map(function (v) { return v === undefined ? null : v; }), st: ST, f: {} };
    for (var i = 0; i < N; i++) P.ids.push(i);
    Object.keys(cellF).forEach(function (k2) { P.f[k2] = cellF[k2]; });
    return { core: core, parts: [P], problems: problems, deps: deps, xdeps: xdeps, raws: raws, nForm: nForm };
  }

  // ------------------------------------------------------------------ verification (mirrors the Node harness)
  function same(a, b) {
    if (a && typeof a === 'object' && a.e) a = { e: a.e }; if (b && typeof b === 'object' && b.e) b = { e: b.e };
    if (isNum(a) && isNum(b)) { var d = Math.abs(a - b); return d <= 1e-9 || d <= 1e-9 * Math.max(Math.abs(a), Math.abs(b)); }
    if (a && b && typeof a === 'object' && typeof b === 'object') return a.e === b.e;
    if ((a === null || a === '' || a === undefined) && (b === null || b === '' || b === undefined)) return true;
    return a === b;
  }
  function show(v) { if (v && typeof v === 'object' && v.e) return v.e; if (v === null || v === undefined) return '(boş)'; return typeof v === 'number' ? String(+v.toPrecision(12)) : JSON.stringify(v); }
  function verify(res) {
    var core = res.core, M = new root.MakroModel(core, res.parts, root.MakroStore.makeTpl(core.tplCode)), X = root.X;
    function addr(id) { var s = core.sheets[M.sh[id]]; return core.books[s.b].n + ' › ' + s.n + '!' + colName(M.c[id]) + M.r[id]; }
    var srcOf = {}; res.raws.forEach(function (b) { srcOf[b.name] = b.source; });
    var local = { ok: 0, bad: [] }, glob = { ok: 0, bad: [] }, link = { ok: 0, bad: [] };
    var hiddenSheet = {}; core.sheets.forEach(function (s, gs) { if (s.n === '__massivlər__') hiddenSheet[gs] = 1; });
    M.order.forEach(function (id) {
      if (hiddenSheet[M.sh[id]]) { M.V[id] = M.evalCell(id); return; }
      var xd = res.xdeps[id], saved = [];
      if (xd) xd.forEach(function (x) { var cv = ((x[1].cache || {})[x[2]] || {})[x[3]]; saved.push([x[0], M.V[x[0]]]); M.V[x[0]] = cv && typeof cv === 'object' && cv.e ? X.E(cv.e) : (cv === undefined ? null : cv); });
      var v = M.evalCell(id);
      for (var i = saved.length - 1; i >= 0; i--) M.V[saved[i][0]] = saved[i][1];
      if (same(v, M.V0[id])) local.ok++; else if (local.bad.length < 20000) local.bad.push([addr(id), show(v), show(M.V0[id])]); else local.bad.push(null);
    });
    for (var i = 0; i < M.N; i++) M.V[i] = M.V0[i];
    M.recalcAll();
    M.order.forEach(function (id) { if (hiddenSheet[M.sh[id]]) return; if (same(M.V[id], M.V0[id])) glob.ok++; else if (glob.bad.length < 20000) glob.bad.push([addr(id), show(M.V[id]), show(M.V0[id])]); else glob.bad.push(null); });
    // link caches of uploaded workbooks vs the source workbook values
    var sidx = {}; core.sheets.forEach(function (s, gs) { sidx[core.books[s.b].n + '' + s.n] = gs; });
    res.raws.forEach(function (b) {
      if (b.source !== 'xlsx') return;
      (b.ext || []).forEach(function (L) {
        var tb = null; core.books.forEach(function (bk) { if (bk.kind === 'model' && norm(bk.n) === norm(L.file || '')) tb = bk.n; });
        res.raws.forEach(function (rb) { if (rb.file && norm(rb.file) === norm(L.file || '')) tb = rb.name; });
        if (!tb) return;
        Object.keys(L.cache).forEach(function (sn) {
          var gs = sidx[tb + '' + sn]; if (gs === undefined) return;
          Object.keys(L.cache[sn]).forEach(function (a) {
            var rc = splitRef(a); if (!rc) return;
            var id = M.smap[gs].get(rc[0] * 20000 + rc[1]), sv = id === undefined ? null : M.V0[id];
            if (sv && sv instanceof X.XErr) sv = { e: sv.e };
            if (same(L.cache[sn][a], sv)) link.ok++; else if (link.bad.length < 200) link.bad.push([b.name + ' → ' + tb + ' › ' + sn + '!' + a, show(L.cache[sn][a]), show(sv)]); else link.bad.push(null);
          });
        });
      });
    });
    return { formulas: M.order.filter(function (id) { return !hiddenSheet[M.sh[id]]; }).length, local: local, global: glob, link: link, model: M };
  }
  function linkTables(res) {
    var core = res.core, bl = {}, sl = {};
    Object.keys(res.deps).forEach(function (k) {
      var cid = +k, s = core.sheets[res.parts[0].sh[cid]], b = core.books[s.b].n;
      res.deps[k].forEach(function (d) {
        var ts = core.sheets[res.parts[0].sh[d]], tb = core.books[ts.b].n;
        if (tb === b) return;
        var e = bl[tb + '' + b] || (bl[tb + '' + b] = { src: tb, dst: b, n: 0, cells: new Set(), formulas: new Set() }); e.n++; e.cells.add(d); e.formulas.add(cid);
        var key = [tb, ts.n, b, s.n].join(''), e2 = sl[key] || (sl[key] = { sb: tb, ss: ts.n, db: b, ds: s.n, n: 0, cells: new Set(), formulas: new Set() }); e2.n++; e2.cells.add(d); e2.formulas.add(cid);
      });
    });
    function fin(o) { return Object.keys(o).map(function (k) { var e = o[k], r = {}; for (var x in e) r[x] = e[x] instanceof Set ? e[x].size : e[x]; return r; }).sort(function (a, b) { return b.n - a.n; }); }
    return { bookLinks: fin(bl), sheetLinks: fin(sl) };
  }

  // ------------------------------------------------------------------ export (data folder as ZIP, stored)
  var CRC = (function () { var t = new Uint32Array(256); for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(u8) { var c = 0xFFFFFFFF; for (var i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function zipStore(files) {
    var te = new TextEncoder(), chunks = [], central = [], off = 0;
    files.forEach(function (f) {
      var name = te.encode(f.name), data = typeof f.data === 'string' ? te.encode(f.data) : f.data, crc = crc32(data);
      var lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true); lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, name.length, true);
      chunks.push(new Uint8Array(lh.buffer), name, data);
      var ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, name.length, true); ch.setUint32(42, off, true);
      central.push(new Uint8Array(ch.buffer), name);
      off += 30 + name.length + data.length;
    });
    var csize = central.reduce(function (s, x) { return s + x.length; }, 0), e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, csize, true); e.setUint32(16, off, true);
    return new Blob(chunks.concat(central, [new Uint8Array(e.buffer)]), { type: 'application/zip' });
  }
  function dataFiles(core, parts, meta) {
    var c2 = {}; for (var k in core) if (k !== 'tplCode') c2[k] = core[k];
    c2.files = []; for (var i = 0; i < 8; i++) c2.files.push('b' + (i < 10 ? '0' : '') + i + '.js');
    var files = [{ name: 'data/core.js', data: 'window.MODEL_CORE=' + JSON.stringify(c2) + ';\nwindow.MODEL_TPL=function(V,K,G,X){return [\n' + core.tplCode.map(function (c) { return 'function(r,k){return ' + c + ';}'; }).join(',\n') + '\n];};\n' }];
    var empty = JSON.stringify({ ids: [], sh: [], r: [], c: [], v: [], st: [], f: {} });
    for (i = 0; i < 8; i++) files.push({ name: 'data/' + c2.files[i], data: 'window.MODEL_PARTS=window.MODEL_PARTS||[];window.MODEL_PARTS.push(' + (i === 0 ? JSON.stringify(parts[0]) : empty) + ');\n' });
    files.push({ name: 'data/meta.js', data: 'window.MODEL_META=' + JSON.stringify(meta) + ';\n' });
    return files;
  }

  root.MakroBuilder = { openZip: openZip, readXlsx: readXlsx, rawFromCore: rawFromCore, compile: compile, verify: verify, linkTables: linkTables,
    zipStore: zipStore, dataFiles: dataFiles, norm: norm, tokenize: tokenize, shiftFormula: shiftFormula,
    attrs: attrs, resolvePath: resolvePath, decodePath: decodePath, parseTheme: parseTheme, tintHex: tint, unesc: unesc };
})(typeof window !== 'undefined' ? window : globalThis);
