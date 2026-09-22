/* Makro model — write an .xlsx workbook from the calculated model (values, number formats, fonts, fills, alignment,
   column widths, merged cells, frozen panes, hidden rows/columns). Used to hand the ministry an updated Excel file.
   Formulas are not written: the values are the result of the current scenario, and keeping formulas would make Excel
   recalculate them from the other (unchanged) workbooks on disk and overwrite those results. */
(function (root) {
  'use strict';
  var NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  function esc(s) {
    return String(s).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '').replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c];
    });
  }
  function colName(n) { var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function num(v) { var s = String(v); return s.indexOf('e') >= 0 ? s.replace('e', 'E') : s; }
  var BUILTIN = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11, '@': 49 };

  // style table: core.styles -> numFmts / fonts / fills / cellXfs
  function styles(used, coreStyles) {
    var fmts = {}, fmtId = 164, fonts = [], fontKey = {}, fills = [], fillKey = {}, xfs = [], xfKey = {}, map = {};
    function font(st) {
      var k = (st.b ? 'b' : '') + (st.i ? 'i' : '') + '|' + (st.fc || '');
      if (fontKey[k] === undefined) { fontKey[k] = fonts.length; fonts.push('<font><sz val="11"/><name val="Calibri"/><family val="2"/>' + (st.b ? '<b/>' : '') + (st.i ? '<i/>' : '') + (st.fc ? '<color rgb="FF' + st.fc + '"/>' : '') + '</font>'); }
      return fontKey[k];
    }
    function fill(st) {
      if (!st.bg) return 0;
      if (fillKey[st.bg] === undefined) { fillKey[st.bg] = fills.length + 2; fills.push('<fill><patternFill patternType="solid"><fgColor rgb="FF' + st.bg + '"/><bgColor indexed="64"/></patternFill></fill>'); }
      return fillKey[st.bg];
    }
    function fmt(st) {
      var f = st.fmt || 'General';
      if (BUILTIN[f] !== undefined) return BUILTIN[f];
      if (fmts[f] === undefined) fmts[f] = fmtId++;
      return fmts[f];
    }
    used.forEach(function (si) {
      var st = coreStyles[si] || {}, nf = fmt(st), fo = font(st), fi = fill(st);
      var al = '', apply = '';
      if (st.ha || st.ind || st.wrap) {
        al = '<alignment' + (st.ha ? ' horizontal="' + (st.ha === 'centerContinuous' ? 'centerContinuous' : st.ha) + '"' : '') + (st.wrap ? ' wrapText="1"' : '') + (st.ind ? ' indent="' + st.ind + '"' : '') + '/>';
        apply = ' applyAlignment="1"';
      }
      var xf = '<xf numFmtId="' + nf + '" fontId="' + fo + '" fillId="' + fi + '" borderId="0" xfId="0"' +
        (nf ? ' applyNumberFormat="1"' : '') + (fo ? ' applyFont="1"' : '') + (fi ? ' applyFill="1"' : '') + apply + (al ? '>' + al + '</xf>' : '/>');
      if (xfKey[xf] === undefined) { xfKey[xf] = xfs.length; xfs.push(xf); }
      map[si] = xfKey[xf];
    });
    var nfs = Object.keys(fmts);
    var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="' + NS + '">' +
      (nfs.length ? '<numFmts count="' + nfs.length + '">' + nfs.map(function (f) { return '<numFmt numFmtId="' + fmts[f] + '" formatCode="' + esc(f) + '"/>'; }).join('') + '</numFmts>' : '') +
      '<fonts count="' + Math.max(1, fonts.length) + '">' + (fonts.length ? fonts.join('') : '<font><sz val="11"/><name val="Calibri"/></font>') + '</fonts>' +
      '<fills count="' + (fills.length + 2) + '"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' + fills.join('') + '</fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + Math.max(1, xfs.length) + '">' + (xfs.length ? xfs.join('') : '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>') + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    return { xml: xml, map: map };
  }

  function sheetXML(S, xfMap) {
    var hidden = {}, all = {};
    (S.hr || []).forEach(function (r) { hidden[r] = 1; all[r] = 1; });          // hidden rows survive even when empty
    Object.keys(S.rows).forEach(function (r) { all[r] = 1; });
    var rows = Object.keys(all).map(Number).sort(function (a, b) { return a - b; }), maxR = 0, maxC = 1;
    var body = '';
    rows.forEach(function (r) {
      var cs = (S.rows[r] || []).sort(function (a, b) { return a.c - b.c; }), cells = '';
      cs.forEach(function (cd) {
        var s = xfMap[cd.s] || 0, ref = colName(cd.c) + r, v = cd.v;
        if (cd.c > maxC) maxC = cd.c;
        var sa = s ? ' s="' + s + '"' : '';
        if (v === null || v === undefined || v === '') { if (s) cells += '<c r="' + ref + '"' + sa + '/>'; return; }
        if (typeof v === 'number') { if (!isFinite(v)) return; cells += '<c r="' + ref + '"' + sa + '><v>' + num(v) + '</v></c>'; }
        else if (typeof v === 'boolean') cells += '<c r="' + ref + '"' + sa + ' t="b"><v>' + (v ? 1 : 0) + '</v></c>';
        else if (typeof v === 'object' && v.e) cells += '<c r="' + ref + '"' + sa + ' t="e"><v>' + esc(v.e) + '</v></c>';
        else cells += '<c r="' + ref + '"' + sa + ' t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
      });
      if (!cells && !hidden[r]) return;
      if (r > maxR) maxR = r;
      body += '<row r="' + r + '"' + (hidden[r] ? ' hidden="1"' : '') + '>' + cells + '</row>';
    });
    var pane = '';
    if (S.fz && (S.fz[0] || S.fz[1])) {
      var top = colName((S.fz[1] || 0) + 1) + ((S.fz[0] || 0) + 1);
      pane = '<pane' + (S.fz[1] ? ' xSplit="' + S.fz[1] + '"' : '') + (S.fz[0] ? ' ySplit="' + S.fz[0] + '"' : '') + ' topLeftCell="' + top + '" activePane="bottomRight" state="frozen"/>';
    }
    var cols = (S.cols || []).filter(function (cd) { return cd.min >= 1 && cd.max >= cd.min; }).map(function (cd) {
      return '<col min="' + cd.min + '" max="' + Math.min(cd.max, 16384) + '" width="' + (cd.w > 0 ? cd.w.toFixed(2) : 8.43) + '" customWidth="1"' + (cd.hidden || cd.w === 0 ? ' hidden="1"' : '') + '/>';
    }).join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="' + NS + '" xmlns:r="' + NSR + '">' +
      (S.tab ? '<sheetPr><tabColor rgb="FF' + S.tab + '"/></sheetPr>' : '') +
      '<dimension ref="A1:' + colName(maxC) + Math.max(1, maxR) + '"/>' +
      '<sheetViews><sheetView' + (S.first ? ' tabSelected="1"' : '') + ' workbookViewId="0">' + pane + '</sheetView></sheetViews>' +
      '<sheetFormatPr defaultColWidth="' + (S.dw || 9.14).toFixed(2) + '" defaultRowHeight="15"/>' +
      (cols ? '<cols>' + cols + '</cols>' : '') +
      '<sheetData>' + body + '</sheetData>' +
      ((S.merges && S.merges.length) ? '<mergeCells count="' + S.merges.length + '">' + S.merges.map(function (m) { return '<mergeCell ref="' + esc(m) + '"/>'; }).join('') + '</mergeCells>' : '') +
      '</worksheet>';
  }

  // sheets: [{ name, state, cols, merges, hr, fz, tab, dw, rows: {r: [{c, v, s}]} }]
  function pack(sheets) {
    var used = {};
    sheets.forEach(function (S) { Object.keys(S.rows).forEach(function (r) { S.rows[r].forEach(function (cd) { used[cd.s] = 1; }); }); });
    var st = styles(Object.keys(used).map(Number), root.MODEL_CORE.styles);
    var files = [
      { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        sheets.map(function (S, i) { return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'; }).join('') +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
      { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="' + NSR + '/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="' + NS + '" xmlns:r="' + NSR + '"><sheets>' +
        sheets.map(function (S, i) { return '<sheet name="' + esc(S.name.slice(0, 31)) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"' + (S.state && S.state !== 'visible' ? ' state="' + (S.state === 'veryHidden' ? 'veryHidden' : 'hidden') + '"' : '') + '/>'; }).join('') +
        '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        sheets.map(function (S, i) { return '<Relationship Id="rId' + (i + 1) + '" Type="' + NSR + '/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>'; }).join('') +
        '<Relationship Id="rId' + (sheets.length + 1) + '" Type="' + NSR + '/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/styles.xml', data: st.xml }
    ];
    sheets.forEach(function (S, i) { files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: sheetXML(S, st.map) }); });
    return root.MakroBuilder.zipStore(files);
  }

  // build one workbook of the model as .xlsx; opts.base = true exports the baseline instead of the current scenario
  function book(core, M, bi, opts) {
    opts = opts || {};
    var V = opts.base ? M.V0 : M.V, out = [];
    core.books[bi].sheets.forEach(function (gs, i) {
      var sm = core.sheets[gs], S = { name: sm.n, state: sm.st, cols: sm.cols, merges: sm.merges, hr: sm.hr, fz: sm.fz, tab: sm.tab, dw: sm.dw, rows: {}, first: i === 0 };
      out.push(S);
    });
    var idx = {}; core.books[bi].sheets.forEach(function (gs, i) { idx[gs] = i; });
    for (var id = 0; id < M.N; id++) {
      var k = idx[M.sh[id]]; if (k === undefined) continue;
      var v = V[id];
      if (v instanceof root.X.XErr) v = { e: v.e };
      if ((v === null || v === undefined || v === '') && !M.st[id]) continue;
      var S2 = out[k], r = M.r[id];
      (S2.rows[r] || (S2.rows[r] = [])).push({ c: M.c[id], v: v === undefined ? null : v, s: M.st[id] });
    }
    return pack(out);
  }
  function fileName(core, bi) { return core.books[bi].n.replace(/^EXT: /, '') + '.xlsx'; }

  root.MakroXlsx = { book: book, pack: pack, fileName: fileName };
})(typeof window !== 'undefined' ? window : globalThis);
