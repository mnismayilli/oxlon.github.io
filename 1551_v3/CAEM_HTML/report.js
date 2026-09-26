/* Makro model — hesabat qurucusu.
   Builds a print-ready document from the live model: the ministry can pick indicators and years,
   title the report, preview it, then export .xlsx or print to PDF. The same module serves both
   models; everything model-specific arrives through the cfg hooks passed to init(). */
(function (root) {
  'use strict';

  var CFG = null, ST = null, MODE = 'custom';
  var $ = function (s, el) { return (el || document).querySelector(s); };
  var $$ = function (s, el) { return Array.prototype.slice.call((el || document).querySelectorAll(s)); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function nf(v, d) { return CFG.nf ? CFG.nf(v, d) : (isNum(v) ? v.toFixed(d == null ? 1 : d) : '—'); }
  // Azerbaijani-friendly folding for search: Ə->e, ı->i, and İ's combining dot removed
  function fold(s) {
    var t = String(s == null ? '' : s).toLowerCase();
    try { t = t.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { }
    return t.replace(/\u0259/g, 'e').replace(/\u0131/g, 'i');
  }
  function today() { var d = new Date(); return d.getDate() + '.' + ('0' + (d.getMonth() + 1)).slice(-2) + '.' + d.getFullYear(); }

  // ---------------------------------------------------------------- state
  function defState() {
    var Y = CFG.years(), all = Y.all;
    return {
      title: CFG.defaultTitle || 'Makroiqtisadi proqnoz',
      sub: '', note: '',
      y0: all[0], y1: all[all.length - 1],
      items: (CFG.defaultItems || []).slice(),
      scen: 1, base: 0, diff: 0, spark: 1,
      pages: { xulase: 1, esas: 1, ssenari: 1 },
      land: 1
    };
  }
  function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) { } }
  function load() {
    var d = defState();
    try {
      var s = JSON.parse(lsGet(CFG.ns + '.report') || 'null');
      if (s && typeof s === 'object') { for (var k in d) if (s[k] !== undefined) d[k] = s[k]; }
    } catch (e) { }
    var all = CFG.years().all;
    if (all.indexOf(d.y0) < 0) d.y0 = all[0];
    if (all.indexOf(d.y1) < 0) d.y1 = all[all.length - 1];
    if (d.y1 < d.y0) d.y1 = d.y0;
    d.items = (d.items || []).filter(function (id) { return !!item(id); });
    return d;
  }
  function save() { lsSet(CFG.ns + '.report', JSON.stringify(ST)); }

  // saved report templates
  function tplAll() { try { return JSON.parse(lsGet(CFG.ns + '.reportTpl') || '[]') || []; } catch (e) { return []; } }
  function tplSave(list) { lsSet(CFG.ns + '.reportTpl', JSON.stringify(list)); }

  // ---------------------------------------------------------------- catalogue
  var CAT = null, BYID = null;
  function cat() {
    if (!CAT) {
      CAT = CFG.sections();
      BYID = {};
      CAT.forEach(function (s) { s.items.forEach(function (it) { it.sec = s.title; BYID[it.id] = it; }); });
    }
    return CAT;
  }
  function item(id) { cat(); return BYID[id] || null; }
  function years() { var all = CFG.years().all, out = []; all.forEach(function (y) { if (y >= ST.y0 && y <= ST.y1) out.push(y); }); return out; }
  function fcFrom() { return CFG.years().forecastFrom; }
  function val(it, y, base) { return CFG.value(it, y, base); }
  function decOf(it, v) { if (it.dec != null) return it.dec; if (/%/.test(it.unit || '')) return 1; return isNum(v) && Math.abs(v) >= 1000 ? 0 : 1; }

  // ---------------------------------------------------------------- print + page styles (injected once)
  var SHEET = [
    '.rep-wrap{display:grid;grid-template-columns:330px minmax(0,1fr);gap:18px;align-items:start}',
    '@media (max-width:980px){.rep-wrap{grid-template-columns:minmax(0,1fr)}}',
    '.rep-main{min-width:0}',
    '.rep-side{min-width:0}',
    '.rep-side{display:flex;flex-direction:column;gap:14px;position:sticky;top:78px;min-width:0}',
    '@media (max-width:980px){.rep-side{position:static}}',
    '.rep-f{display:flex;flex-direction:column;gap:5px;margin-bottom:11px}',
    '.rep-f>label{font-size:12px;font-weight:600;color:var(--muted,#71808D)}',
    '.rep-f input[type=text],.rep-f textarea,.rep-f select{font:inherit;font-size:13px;padding:7px 9px;border:1px solid var(--line,#E7EBEF);border-radius:7px;background:var(--bg,#fff);color:inherit;width:100%}',
    '.rep-f textarea{min-height:56px;resize:vertical}',
    '.rep-row{display:flex;gap:8px}.rep-row>*{flex:1;min-width:0}',
    '.rep-ck{display:flex;align-items:center;gap:7px;font-size:13px;padding:3px 0;cursor:pointer}',
    '.rep-ck input{margin:0;flex:none}',
    '.rep-pick{max-height:330px;overflow:auto;border:1px solid var(--line,#E7EBEF);border-radius:8px;padding:8px}',
    '.rep-pick h5{margin:9px 0 3px;font-size:11px;letter-spacing:.04em;text-transform:uppercase;color:var(--muted,#71808D)}',
    '.rep-pick h5:first-child{margin-top:0}',
    '.rep-pick label{display:flex;gap:7px;align-items:flex-start;font-size:12.5px;padding:2.5px 3px;border-radius:5px;cursor:pointer;line-height:1.35}',
    '.rep-pick label:hover{background:var(--bg2,#F5F7FC)}',
    '.rep-pick label.sub{padding-left:14px}',
    '.rep-pick input{margin:2px 0 0;flex:none}',
    '.rep-pick .u{color:var(--muted,#71808D);font-size:11px;margin-left:3px}',
    '.rep-sel{margin-top:9px;display:flex;flex-direction:column;gap:3px;max-height:230px;overflow:auto}',
    '.rep-sel .it{display:flex;align-items:center;gap:6px;font-size:12.5px;padding:4px 6px;border:1px solid var(--line,#E7EBEF);border-radius:6px;background:var(--bg,#fff)}',
    '.rep-sel .it span{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
    '.rep-sel .it button{border:0;background:none;cursor:pointer;color:var(--muted,#71808D);font-size:13px;line-height:1;padding:2px 4px;border-radius:4px}',
    '.rep-sel .it button:hover{background:var(--bg2,#F5F7FC);color:var(--ink,#15202B)}',
    '.rep-bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:12px}',
    '.rep-bar .sp{margin-left:auto}',
    /* the document */
    '.rep-doc{background:#fff;color:#15202B;border:1px solid var(--line,#E7EBEF);border-radius:10px;padding:26px 28px;font-size:12.5px}',
    '.rep-doc .org{font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:#71808D}',
    '.rep-doc h1{font-size:20px;margin:6px 0 2px;font-weight:700;line-height:1.25}',
    '.rep-doc .sub{font-size:13px;color:#4A5764;margin:0 0 4px}',
    '.rep-doc .meta{font-size:11.5px;color:#71808D;border-bottom:1.5px solid #15202B;padding-bottom:9px;margin-bottom:14px;display:flex;gap:14px;flex-wrap:wrap}',
    '.rep-doc h2{font-size:13px;margin:18px 0 7px;font-weight:700}',
    '.rep-tw{overflow-x:auto;max-width:100%}',
    '.rep-tw>table{width:auto;min-width:100%}',
    '.rep-doc table{border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums}',
    '.rep-doc th.l,.rep-doc td.l{min-width:165px}',
    '.rep-doc.w2 th,.rep-doc.w2 td{padding:3px 5px;font-size:12px}',
    '.rep-doc.w3 th,.rep-doc.w3 td{padding:2.5px 4px;font-size:11.5px}',
    '.rep-doc.w3 th.l,.rep-doc.w3 td.l{min-width:150px}',
    '.rep-doc th,.rep-doc td{padding:4px 7px;text-align:right;border-bottom:1px solid #E7EBEF;white-space:nowrap}',
    '.rep-doc th.l,.rep-doc td.l{text-align:left;white-space:normal}',
    '.rep-doc thead th{font-size:10.5px;text-transform:uppercase;letter-spacing:.03em;color:#4A5764;border-bottom:1.5px solid #B8C2CC;vertical-align:bottom}',
    '.rep-doc thead th.fc{color:#1F6FB2}',
    '.rep-doc tbody tr.head td{font-weight:700;background:#F5F7FC;font-size:11px;text-transform:uppercase;letter-spacing:.03em;padding-top:7px;padding-bottom:5px}',
    '.rep-doc td.l.sub{padding-left:18px}',
    '.rep-doc td.u{color:#71808D;font-size:11px;text-align:left}',
    '.rep-doc tbody tr:hover{background:#FAFBFD}',
    '.rep-doc .spark{width:62px;height:17px;display:block}',
    '.rep-doc .nil{color:#B8C2CC}',
    '.rep-doc .dif.up{color:#1E7B4F}.rep-doc .dif.dn{color:#B3261E}',
    '.rep-doc .note{margin-top:14px;font-size:11.5px;color:#4A5764;border-top:1px solid #E7EBEF;padding-top:9px;white-space:pre-wrap}',
    '.rep-doc .foot{margin-top:9px;font-size:10.5px;color:#71808D}',
    '.rep-empty{padding:34px;text-align:center;color:var(--muted,#71808D);font-size:13px}',
    '.rep-pg+.rep-pg{margin-top:26px;border-top:1px dashed #E7EBEF;padding-top:20px}',
    /* print */
    '@media print{',
    '  body{background:#fff}',
    '  body>*{display:none!important}',
    '  body>#rep-print{display:block!important;position:static}',
    '  #rep-print .rep-doc{border:0;border-radius:0;padding:0;font-size:9.4pt}',
    '  #rep-print .rep-tw{overflow:visible}',
    '  #rep-print .rep-doc.w2{font-size:8pt}',
    '  #rep-print .rep-doc.w3{font-size:6.8pt}',
    '  #rep-print .rep-doc.w2 th.l,#rep-print .rep-doc.w2 td.l{min-width:0}',
    '  #rep-print .rep-doc.w3 th.l,#rep-print .rep-doc.w3 td.l{min-width:0}',
    '  #rep-print .rep-doc thead{display:table-header-group}',
    '  #rep-print .rep-doc tr{page-break-inside:avoid}',
    '  #rep-print .rep-pg{page-break-before:always}',
    '  #rep-print .rep-pg:first-child{page-break-before:auto}',
    '  #rep-print .rep-doc tbody tr:hover{background:none}',
    '}'
  ].join('\n');
  function injectCSS() {
    if ($('#rep-css')) return;
    var s = document.createElement('style'); s.id = 'rep-css'; s.textContent = SHEET;
    document.head.appendChild(s);
  }
  function pageRule(land) {
    var id = 'rep-page-css', el = $('#' + id);
    if (!el) { el = document.createElement('style'); el.id = id; document.head.appendChild(el); }
    el.textContent = '@page{size:A4 ' + (land ? 'landscape' : 'portrait') + ';margin:14mm 12mm}';
  }

  // ---------------------------------------------------------------- sparkline (self-contained)
  function spark(vals) {
    var pts = vals.filter(isNum);
    if (pts.length < 2) return '';
    var lo = Math.min.apply(null, pts), hi = Math.max.apply(null, pts), W = 62, H = 17, P = 2;
    var span = hi - lo || 1, n = vals.length, d = '', started = false, last = null;
    for (var i = 0; i < n; i++) {
      if (!isNum(vals[i])) continue;
      var x = P + (n === 1 ? 0 : (W - 2 * P) * i / (n - 1)), y = H - P - (H - 2 * P) * (vals[i] - lo) / span;
      d += (started ? 'L' : 'M') + x.toFixed(1) + ' ' + y.toFixed(1);
      started = true; last = [x, y];
    }
    if (!started) return '';
    return '<svg class="spark" viewBox="0 0 ' + W + ' ' + H + '" aria-hidden="true">' +
      '<path d="' + d + '" fill="none" stroke="#1F6FB2" stroke-width="1.4" stroke-linejoin="round" stroke-linecap="round"/>' +
      '<circle cx="' + last[0].toFixed(1) + '" cy="' + last[1].toFixed(1) + '" r="1.9" fill="#1F6FB2"/></svg>';
  }

  // ---------------------------------------------------------------- the document
  function docHead(title, sub) {
    var Y = years(), scen = CFG.scenarioName ? CFG.scenarioName() : '';
    return '<div class="org">' + esc(CFG.org || '') + '</div>' +
      '<h1>' + esc(title || '—') + '</h1>' +
      (sub ? '<p class="sub">' + esc(sub) + '</p>' : '') +
      '<div class="meta"><span>' + esc(CFG.model || '') + '</span>' +
      (Y.length ? '<span>' + Y[0] + '–' + Y[Y.length - 1] + '</span>' : '') +
      (scen ? '<span>' + esc(scen) + '</span>' : '') +
      '<span>' + today() + '</span></div>';
  }
  function dens(n) { return n <= 10 ? '' : n <= 18 ? ' w2' : ' w3'; }
  function colCount() { return (ST.scen ? 1 : 0) + (ST.base ? 1 : 0) + (ST.diff ? 1 : 0); }
  function valCellsHTML(it, y) {
    var out = '', v = val(it, y, false), b = val(it, y, true), d = decOf(it, isNum(v) ? v : b);
    if (ST.scen) out += '<td class="n">' + (isNum(v) ? nf(v, d) : '<span class="nil">—</span>') + '</td>';
    if (ST.base) out += '<td class="n">' + (isNum(b) ? nf(b, d) : '<span class="nil">—</span>') + '</td>';
    if (ST.diff) {
      var ok = isNum(v) && isNum(b), df = ok ? v - b : null;
      out += '<td class="n">' + (!ok ? '<span class="nil">—</span>' : Math.abs(df) < Math.pow(10, -d) / 2 ? '<span class="nil">0</span>' :
        '<span class="dif ' + (df > 0 ? 'up' : 'dn') + '">' + (df > 0 ? '+' : '−') + nf(Math.abs(df), d) + '</span>') + '</td>';
    }
    return out;
  }
  function customDoc() {
    var Y = years(), fc = fcFrom(), cc = colCount();
    if (!ST.items.length) return '<div class="rep-doc"><div class="rep-empty">Hesabat üçün ən azı bir göstərici seçin.</div></div>';
    if (!cc) return '<div class="rep-doc"><div class="rep-empty">Ən azı bir sütun növü seçin (ssenari, baza və ya fərq).</div></div>';
    var sub = [];
    if (ST.scen) sub.push('Ssenari'); if (ST.base) sub.push('Baza'); if (ST.diff) sub.push('Fərq');
    var h = '<div class="rep-doc' + dens(Y.length * cc) + '">' + docHead(ST.title, ST.sub);
    h += '<div class="rep-tw"><table><thead><tr><th class="l" rowspan="' + (cc > 1 ? 2 : 1) + '">Göstərici</th><th class="l" rowspan="' + (cc > 1 ? 2 : 1) + '">Ölçü vahidi</th>';
    Y.forEach(function (y) { h += '<th colspan="' + cc + '" class="' + (y >= fc ? 'fc' : '') + '">' + y + (y >= fc ? '' : '') + '</th>'; });
    if (ST.spark) h += '<th rowspan="' + (cc > 1 ? 2 : 1) + '">Trend</th>';
    h += '</tr>';
    if (cc > 1) { h += '<tr>'; Y.forEach(function () { sub.forEach(function (s) { h += '<th>' + s + '</th>'; }); }); h += '</tr>'; }
    h += '</thead><tbody>';
    var lastSec = null;
    ST.items.forEach(function (id) {
      var it = item(id); if (!it) return;
      if (it.sec !== lastSec) { lastSec = it.sec; h += '<tr class="head"><td class="l" colspan="' + (2 + Y.length * cc + (ST.spark ? 1 : 0)) + '">' + esc(it.sec) + '</td></tr>'; }
      h += '<tr><td class="l' + (it.sub ? ' sub' : '') + '">' + esc(it.name) + '</td><td class="u">' + esc(it.unit || '') + '</td>';
      Y.forEach(function (y) { h += valCellsHTML(it, y); });
      if (ST.spark) h += '<td>' + spark(Y.map(function (y) { return val(it, y, false); })) + '</td>';
      h += '</tr>';
    });
    h += '</tbody></table></div>';
    if (ST.note) h += '<div class="note">' + esc(ST.note) + '</div>';
    h += '<div class="foot">' + esc(CFG.footer || '') + (fc ? ' · ' + fc + '-ci ildən etibarən proqnoz.' : '') + '</div>';
    return h + '</div>';
  }

  // ---------------------------------------------------------------- the ministry's 3-page report
  function p3Doc() {
    var P = CFG.pages3(), on = P.filter(function (p) { return ST.pages[p.id]; });
    if (!on.length) return '<div class="rep-doc"><div class="rep-empty">Ən azı bir səhifə seçin.</div></div>';
    var wide = 0;
    on.forEach(function (p) { var g = p.grid(ST.y0, ST.y1); if (g.cols.length > wide) wide = g.cols.length; });
    var h = '<div class="rep-doc' + dens(wide) + '">' + docHead(ST.title, ST.sub);
    on.forEach(function (p, i) {
      var g = p.grid(ST.y0, ST.y1);
      h += '<div class="rep-pg">';
      h += '<h2>' + esc(p.name) + (g.note ? ' <span class="u" style="font-weight:400">' + esc(g.note) + '</span>' : '') + '</h2>';
      h += '<div class="rep-tw"><table><thead>';
      g.head.forEach(function (hr) {
        var cells = hr.cells || hr;
        h += '<tr>' + cells.map(function (c) {
          return '<th' + (c.l ? ' class="l"' : (c.fc ? ' class="fc"' : '')) + (c.span > 1 ? ' colspan="' + c.span + '"' : '') + (c.rs > 1 ? ' rowspan="' + c.rs + '"' : '') + '>' + esc(c.t) + '</th>';
        }).join('') + '</tr>';
      });
      h += '</thead><tbody>';
      g.rows.forEach(function (r) {
        if (r.head) { h += '<tr class="head"><td class="l" colspan="' + (2 + g.cols.length) + '">' + esc(r.lab) + '</td></tr>'; return; }
        h += '<tr><td class="l' + (r.ind ? ' sub' : '') + '"' + (r.bold ? ' style="font-weight:600"' : '') + '>' + esc(r.lab) + '</td><td class="u">' + esc(r.unit || '') + '</td>' +
          r.cells.map(function (c) { return '<td class="n">' + (isNum(c) ? nf(c, r.dec) : '<span class="nil">—</span>') + '</td>'; }).join('') + '</tr>';
      });
      h += '</tbody></table></div></div>';
    });
    if (ST.note) h += '<div class="note">' + esc(ST.note) + '</div>';
    h += '<div class="foot">' + esc(CFG.footer || '') + '</div>';
    return h + '</div>';
  }
  function docHTML() { return MODE === 'p3' ? p3Doc() : customDoc(); }

  // ---------------------------------------------------------------- Excel export
  var XSTY = {
    ttl: { b: 1, fc: '15202B' },
    org: { fc: '71808D' },
    meta: { fc: '4A5764' },
    hdr: { b: 1, bg: 'EDF2F9', fc: '15202B', ha: 'center', wrap: 1 },
    hdrf: { b: 1, bg: 'E3EDF7', fc: '1F6FB2', ha: 'center', wrap: 1 },
    lab: { wrap: 1 },
    labb: { b: 1, wrap: 1 },
    lsub: { ind: 1, wrap: 1 },
    sec: { b: 1, bg: 'F2F5F9', wrap: 1 },
    unit: { fc: '71808D' },
    note: { fc: '4A5764', wrap: 1 },
    n0: { fmt: '#,##0' }, n1: { fmt: '#,##0.0' }, n2: { fmt: '#,##0.00' },
    nil: { fc: 'B8C2CC', ha: 'center' }
  };
  function nsty(d) { return d <= 0 ? 'n0' : d === 1 ? 'n1' : 'n2'; }
  function put(S, r, c, v, s) { (S.rows[r] || (S.rows[r] = [])).push({ c: c, v: v, s: s || 'lab' }); }
  function sheetHead(S, title, sub) {
    var r = 1;
    put(S, r++, 1, CFG.org || '', 'org');
    put(S, r++, 1, title || '', 'ttl');
    if (sub) put(S, r++, 1, sub, 'meta');
    var scen = CFG.scenarioName ? CFG.scenarioName() : '';
    put(S, r++, 1, (CFG.model || '') + (scen ? ' · ' + scen : '') + ' · ' + today(), 'meta');
    return r + 1;
  }
  function customSheets() {
    var Y = years(), cc = colCount(), S = { name: 'Hesabat', rows: {}, cols: [], merges: [], fz: null };
    var r = sheetHead(S, ST.title, ST.sub), top = r;
    var sub = []; if (ST.scen) sub.push('Ssenari'); if (ST.base) sub.push('Baza'); if (ST.diff) sub.push('Fərq');
    put(S, r, 1, 'Göstərici', 'hdr'); put(S, r, 2, 'Ölçü vahidi', 'hdr');
    var fc = fcFrom(), c = 3;
    Y.forEach(function (y) {
      put(S, r, c, y, y >= fc ? 'hdrf' : 'hdr');
      if (cc > 1) { S.merges.push(colL(c) + r + ':' + colL(c + cc - 1) + r); for (var k = 1; k < cc; k++) put(S, r, c + k, null, y >= fc ? 'hdrf' : 'hdr'); }
      c += cc;
    });
    if (cc > 1) {
      S.merges.push('A' + r + ':A' + (r + 1)); S.merges.push('B' + r + ':B' + (r + 1));
      var c2 = 3;
      Y.forEach(function (y) { sub.forEach(function (t) { put(S, r + 1, c2++, t, y >= fc ? 'hdrf' : 'hdr'); }); });
      r++;
    }
    S.fz = [r, 2];
    r++;
    var lastSec = null;
    ST.items.forEach(function (id) {
      var it = item(id); if (!it) return;
      if (it.sec !== lastSec) {
        lastSec = it.sec;
        put(S, r, 1, it.sec, 'sec');
        for (var k = 2; k <= 2 + Y.length * cc; k++) put(S, r, k, null, 'sec');
        r++;
      }
      put(S, r, 1, it.name, it.sub ? 'lsub' : 'lab');
      put(S, r, 2, it.unit || '', 'unit');
      var cx = 3;
      Y.forEach(function (y) {
        var v = val(it, y, false), b = val(it, y, true), d = decOf(it, isNum(v) ? v : b);
        if (ST.scen) put(S, r, cx++, isNum(v) ? v : null, nsty(d));
        if (ST.base) put(S, r, cx++, isNum(b) ? b : null, nsty(d));
        if (ST.diff) put(S, r, cx++, (isNum(v) && isNum(b)) ? v - b : null, nsty(d));
      });
      r++;
    });
    if (ST.note) { r++; put(S, r, 1, ST.note, 'note'); }
    r++; put(S, r, 1, CFG.footer || '', 'org');
    S.cols = [{ min: 1, max: 1, w: 42 }, { min: 2, max: 2, w: 13 }, { min: 3, max: 2 + Y.length * cc, w: cc > 1 ? 10 : 11.5 }];
    return [S];
  }
  function colL(n) { var s = ''; while (n > 0) { var k = (n - 1) % 26; s = String.fromCharCode(65 + k) + s; n = Math.floor((n - 1) / 26); } return s; }
  function p3Sheets() {
    var P = CFG.pages3(), out = [];
    P.forEach(function (p) {
      if (!ST.pages[p.id]) return;
      var g = p.grid(ST.y0, ST.y1), S = { name: p.name.slice(0, 31), rows: {}, cols: [], merges: [], fz: null };
      var r = sheetHead(S, ST.title || p.name, ST.sub);
      var top = r;
      g.head.forEach(function (hr) {
        var cells = hr.cells || hr, c = hr.from || 1;
        cells.forEach(function (cell) {
          put(S, r, c, cell.t, cell.fc ? 'hdrf' : 'hdr');
          if (cell.span > 1) { S.merges.push(colL(c) + r + ':' + colL(c + cell.span - 1) + r); for (var k = 1; k < cell.span; k++) put(S, r, c + k, null, cell.fc ? 'hdrf' : 'hdr'); }
          if (cell.rs > 1) S.merges.push(colL(c) + r + ':' + colL(c) + (r + cell.rs - 1));
          c += cell.span || 1;
        });
        r++;
      });
      S.fz = [r - 1, 2];
      if (g.colw) S.colw = g.colw;
      g.rows.forEach(function (row) {
        if (row.head) {
          put(S, r, 1, row.lab, 'sec');
          for (var k = 2; k <= 2 + g.cols.length; k++) put(S, r, k, null, 'sec');
          r++; return;
        }
        put(S, r, 1, row.lab, row.bold ? 'labb' : (row.ind ? 'lsub' : 'lab'));
        put(S, r, 2, row.unit || '', 'unit');
        row.cells.forEach(function (v, i) { put(S, r, 3 + i, isNum(v) ? v : null, nsty(row.dec)); });
        r++;
      });
      if (ST.note) { r++; put(S, r, 1, ST.note, 'note'); }
      S.cols = [{ min: 1, max: 1, w: 40 }, { min: 2, max: 2, w: 13 }, { min: 3, max: 2 + g.cols.length, w: 11 }];
      out.push(S);
    });
    return out;
  }
  function fileBase() {
    var t = (ST.title || 'hesabat').replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
    return (t || 'hesabat') + ' ' + new Date().toISOString().slice(0, 10);
  }
  function download(name, data) {
    var blob = data instanceof Blob ? data : new Blob([data], { type: 'text/plain;charset=utf-8' });
    var url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 400);
  }
  function exportXlsx() {
    var sheets = MODE === 'p3' ? p3Sheets() : customSheets();
    if (!sheets.length) { CFG.toast('Əvvəlcə hesabatın məzmununu seçin.'); return; }
    sheets[0].first = 1;
    try {
      var blob = root.MakroXlsx.pack(sheets, XSTY);
      download(fileBase() + '.xlsx', blob);
      CFG.toast('Excel faylı yükləndi (' + Math.max(1, Math.round(blob.size / 1024)) + ' KB)');
    } catch (e) { CFG.toast('Excel faylı yaradılmadı: ' + e.message); }
  }

  // ---------------------------------------------------------------- print / PDF
  function printDoc() {
    pageRule(ST.land);
    var old = $('#rep-print'); if (old) old.remove();
    var d = document.createElement('div'); d.id = 'rep-print'; d.innerHTML = docHTML();
    document.body.appendChild(d);
    var clean = function () { var n = $('#rep-print'); if (n) n.remove(); window.removeEventListener('afterprint', clean); };
    window.addEventListener('afterprint', clean);
    setTimeout(function () { window.print(); setTimeout(clean, 1500); }, 60);
  }

  // ---------------------------------------------------------------- UI
  function yearOpts(sel) { return CFG.years().all.map(function (y) { return '<option value="' + y + '"' + (y === sel ? ' selected' : '') + '>' + y + '</option>'; }).join(''); }
  function pickerHTML(q) {
    q = fold(q);
    var h = '';
    cat().forEach(function (s) {
      var rows = s.items.filter(function (it) { return !q || fold(it.name).indexOf(q) >= 0 || fold(it.unit).indexOf(q) >= 0 || fold(s.title).indexOf(q) >= 0; });
      if (!rows.length) return;
      h += '<h5>' + esc(s.title) + '</h5>';
      rows.forEach(function (it) {
        h += '<label class="' + (it.sub ? 'sub' : '') + '"><input type="checkbox" data-pick="' + esc(it.id) + '"' + (ST.items.indexOf(it.id) >= 0 ? ' checked' : '') + '><span>' + esc(it.name) + (it.unit ? '<span class="u">' + esc(it.unit) + '</span>' : '') + '</span></label>';
      });
    });
    return h || '<div class="rep-empty" style="padding:16px">Uyğun göstərici tapılmadı.</div>';
  }
  function selHTML() {
    if (!ST.items.length) return '<div class="small muted" style="padding:6px 2px">Heç nə seçilməyib.</div>';
    return ST.items.map(function (id, i) {
      var it = item(id); if (!it) return '';
      return '<div class="it" data-i="' + i + '"><span title="' + esc(it.name) + '">' + esc(it.name) + '</span>' +
        '<button data-up="' + i + '" title="Yuxarı"' + (i === 0 ? ' disabled style="opacity:.3"' : '') + '>↑</button>' +
        '<button data-dn="' + i + '" title="Aşağı"' + (i === ST.items.length - 1 ? ' disabled style="opacity:.3"' : '') + '>↓</button>' +
        '<button data-rm="' + i + '" title="Sil">✕</button></div>';
    }).join('');
  }
  function tplHTML() {
    var t = tplAll();
    return '<select id="rep-tpl"><option value="">— saxlanmış şablon —</option>' +
      t.map(function (x, i) { return '<option value="' + i + '">' + esc(x.name) + '</option>'; }).join('') + '</select>';
  }
  function sideHTML() {
    var h = '<div class="card"><h3 style="margin:0 0 11px;font-size:14px">Sənəd</h3>';
    h += '<div class="rep-f"><label for="rep-title">Başlıq</label><input type="text" id="rep-title" value="' + esc(ST.title) + '" placeholder="Hesabatın adı"></div>';
    h += '<div class="rep-f"><label for="rep-sub">Alt başlıq</label><input type="text" id="rep-sub" value="' + esc(ST.sub) + '" placeholder="isteğe bağlı"></div>';
    h += '<div class="rep-row"><div class="rep-f"><label for="rep-y0">İlk il</label><select id="rep-y0">' + yearOpts(ST.y0) + '</select></div>' +
      '<div class="rep-f"><label for="rep-y1">Son il</label><select id="rep-y1">' + yearOpts(ST.y1) + '</select></div></div>';
    if (MODE === 'custom') {
      h += '<div class="rep-f"><label>Sütunlar</label>' +
        '<label class="rep-ck"><input type="checkbox" id="rep-scen"' + (ST.scen ? ' checked' : '') + '> Ssenari dəyəri</label>' +
        '<label class="rep-ck"><input type="checkbox" id="rep-base"' + (ST.base ? ' checked' : '') + '> Baza dəyəri</label>' +
        '<label class="rep-ck"><input type="checkbox" id="rep-diff"' + (ST.diff ? ' checked' : '') + '> Fərq (ssenari − baza)</label>' +
        '<label class="rep-ck"><input type="checkbox" id="rep-spark"' + (ST.spark ? ' checked' : '') + '> Trend qrafiki</label></div>';
    } else {
      h += '<div class="rep-f"><label>Səhifələr</label>' + CFG.pages3().map(function (p) {
        return '<label class="rep-ck"><input type="checkbox" data-pg="' + p.id + '"' + (ST.pages[p.id] ? ' checked' : '') + '> ' + esc(p.name) + '</label>';
      }).join('') + '</div>';
    }
    h += '<div class="rep-f"><label for="rep-note">Qeyd</label><textarea id="rep-note" placeholder="Hesabatın sonunda görünür">' + esc(ST.note) + '</textarea></div>';
    h += '<div class="rep-f"><label class="rep-ck"><input type="checkbox" id="rep-land"' + (ST.land ? ' checked' : '') + '> Çapda albom (landscape)</label></div>';
    h += '<div class="rep-f"><label>Şablon</label>' + tplHTML() +
      '<div style="display:flex;gap:6px;margin-top:6px"><button class="btn sm" id="rep-tsave">Saxla</button><button class="btn sm" id="rep-tdel">Sil</button></div></div>';
    h += '</div>';
    if (MODE === 'custom') {
      h += '<div class="card"><h3 style="margin:0 0 9px;font-size:14px">Göstəricilər <span class="small muted" id="rep-cnt">' + ST.items.length + ' seçilib</span></h3>' +
        '<div class="rep-f"><input type="text" id="rep-q" placeholder="Göstərici axtar…"></div>' +
        '<div class="rep-pick" id="rep-pick">' + pickerHTML('') + '</div>' +
        '<div style="display:flex;gap:6px;margin-top:8px"><button class="btn sm" id="rep-none">Hamısını təmizlə</button></div>' +
        '<div class="rep-sel" id="rep-sel">' + selHTML() + '</div></div>';
    }
    return h;
  }
  function uncheck(id) { $$('[data-pick]').forEach(function (c) { if (c.getAttribute('data-pick') === id) c.checked = false; }); }
  function refreshPreview() { var p = $('#rep-prev'); if (p) p.innerHTML = docHTML(); }
  function refreshSel() {
    var s = $('#rep-sel'); if (s) s.innerHTML = selHTML();
    var c = $('#rep-cnt'); if (c) c.textContent = ST.items.length + ' seçilib';
  }
  function page(v) {
    injectCSS(); cat();
    if (!ST) ST = load();
    if (MODE === 'p3' && !CFG.pages3) MODE = 'custom';
    var modes = CFG.pages3 ? '<div class="seg" id="rep-mode" role="group" aria-label="Hesabat növü">' +
      '<button data-m="custom" class="' + (MODE === 'custom' ? 'on' : '') + '">Öz hesabatım</button>' +
      '<button data-m="p3" class="' + (MODE === 'p3' ? 'on' : '') + '">Nazirliyin 3 səhifəlik hesabatı</button></div>' : '';
    v.innerHTML = '<div class="eyebrow">Hesabat</div><h1 class="h1">' + esc(CFG.pageTitle || 'Hesabat hazırla') + '</h1>' +
      '<p class="lead">' + esc(MODE === 'p3' ? (CFG.leadP3 || '') : (CFG.leadCustom || '')) + '</p>' +
      (modes ? '<div style="margin:14px 0 16px">' + modes + '</div>' : '<div style="height:10px"></div>') +
      '<div class="rep-wrap"><div class="rep-side">' + sideHTML() + '</div>' +
      '<div class="rep-main"><div class="rep-bar"><button class="btn pri" id="rep-pdf">PDF / çap</button><button class="btn" id="rep-xlsx">Excel (.xlsx)</button>' +
      '<span class="sp small muted">Önizləmə çap ediləcək sənədin eynisidir</span></div>' +
      '<div id="rep-prev">' + docHTML() + '</div></div></div>';
    wire(v);
  }
  function wire(v) {
    function bindText(id, key) { var el = $('#' + id); if (!el) return; el.addEventListener('input', function () { ST[key] = el.value; save(); refreshPreview(); }); }
    bindText('rep-title', 'title'); bindText('rep-sub', 'sub'); bindText('rep-note', 'note');
    function bindCk(id, key) { var el = $('#' + id); if (!el) return; el.addEventListener('change', function () { ST[key] = el.checked ? 1 : 0; save(); refreshPreview(); }); }
    bindCk('rep-scen', 'scen'); bindCk('rep-base', 'base'); bindCk('rep-diff', 'diff'); bindCk('rep-spark', 'spark');
    var land = $('#rep-land'); if (land) land.addEventListener('change', function () { ST.land = land.checked ? 1 : 0; save(); });
    ['rep-y0', 'rep-y1'].forEach(function (id, i) {
      var el = $('#' + id); if (!el) return;
      el.addEventListener('change', function () {
        ST[i ? 'y1' : 'y0'] = +el.value;
        if (ST.y1 < ST.y0) { if (i) ST.y0 = ST.y1; else ST.y1 = ST.y0; var o = $('#rep-y' + (i ? '0' : '1')); if (o) o.value = i ? ST.y0 : ST.y1; }
        save(); refreshPreview();
      });
    });
    var q = $('#rep-q');
    if (q) q.addEventListener('input', function () { $('#rep-pick').innerHTML = pickerHTML(q.value); });
    var pick = $('#rep-pick');
    if (pick) pick.addEventListener('change', function (e) {
      var cb = e.target.closest('[data-pick]'); if (!cb) return;
      var id = cb.getAttribute('data-pick'), i = ST.items.indexOf(id);
      if (cb.checked && i < 0) ST.items.push(id); else if (!cb.checked && i >= 0) ST.items.splice(i, 1);
      save(); refreshSel(); refreshPreview();
    });
    var sel = $('#rep-sel');
    if (sel) sel.addEventListener('click', function (e) {
      var b = e.target.closest('button'); if (!b) return;
      var up = b.getAttribute('data-up'), dn = b.getAttribute('data-dn'), rm = b.getAttribute('data-rm'), i;
      if (up !== null) { i = +up; if (i > 0) { var t = ST.items[i - 1]; ST.items[i - 1] = ST.items[i]; ST.items[i] = t; } }
      else if (dn !== null) { i = +dn; if (i < ST.items.length - 1) { var t2 = ST.items[i + 1]; ST.items[i + 1] = ST.items[i]; ST.items[i] = t2; } }
      else if (rm !== null) { i = +rm; var gone = ST.items.splice(i, 1)[0]; uncheck(gone); }
      else return;
      save(); refreshSel(); refreshPreview();
    });
    var none = $('#rep-none');
    if (none) none.addEventListener('click', function () { ST.items = []; save(); $$('[data-pick]').forEach(function (c) { c.checked = false; }); refreshSel(); refreshPreview(); });
    $$('[data-pg]').forEach(function (cb) {
      cb.addEventListener('change', function () { ST.pages[cb.getAttribute('data-pg')] = cb.checked ? 1 : 0; save(); refreshPreview(); });
    });
    var mode = $('#rep-mode');
    if (mode) mode.addEventListener('click', function (e) {
      var b = e.target.closest('[data-m]'); if (!b) return;
      MODE = b.getAttribute('data-m'); page($('#view') || v);
    });
    var tpl = $('#rep-tpl');
    if (tpl) tpl.addEventListener('change', function () {
      var t = tplAll()[+tpl.value];
      if (!t) return;
      for (var k in t.st) if (ST[k] !== undefined) ST[k] = t.st[k];
      ST.items = (t.st.items || []).filter(function (id) { return !!item(id); });
      save(); page($('#view') || v); CFG.toast('Şablon tətbiq olundu: ' + t.name);
    });
    var ts = $('#rep-tsave');
    if (ts) ts.addEventListener('click', function () {
      var n = prompt('Şablonun adı:', ST.title || 'Hesabat');
      if (!n) return;
      var list = tplAll(), copy = JSON.parse(JSON.stringify(ST));
      var at = -1; list.forEach(function (x, i) { if (x.name === n) at = i; });
      if (at >= 0) list[at] = { name: n, st: copy }; else list.push({ name: n, st: copy });
      tplSave(list); page($('#view') || v); CFG.toast('Şablon saxlanıldı: ' + n);
    });
    var td = $('#rep-tdel');
    if (td) td.addEventListener('click', function () {
      var s = $('#rep-tpl'); if (!s || s.value === '') { CFG.toast('Əvvəlcə şablonu seçin.'); return; }
      var list = tplAll(), t = list[+s.value];
      if (!t || !confirm('«' + t.name + '» şablonu silinsin?')) return;
      list.splice(+s.value, 1); tplSave(list); page($('#view') || v); CFG.toast('Şablon silindi.');
    });
    var pdf = $('#rep-pdf'); if (pdf) pdf.addEventListener('click', printDoc);
    var xl = $('#rep-xlsx'); if (xl) xl.addEventListener('click', exportXlsx);
  }

  root.MakroReport = {
    init: function (cfg) { CFG = cfg; CAT = null; BYID = null; ST = null; },
    page: page,
    reset: function () { CAT = null; BYID = null; ST = null; },
    setMode: function (m) { MODE = m; }
  };
})(typeof window !== 'undefined' ? window : globalThis);
