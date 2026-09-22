/* CAEM proqnoz modeli — klassik görünüş (hash-routed single page). */
(function () {
  'use strict';
  var CORE = window.MODEL_CORE, META = window.MODEL_META || { books: {}, methodology: [], issues: [], groups: [] };
  var X = window.X;
  var M, BOOKS = CORE.books, SHEETS = CORE.sheets, STY = CORE.styles, INFO = { firstYear: 2025, lastActual: 2024, endYear: 2029, baseShort: 'Excel-də', baseLabel: 'Excel bazası' };
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var main, inspector, sidebar;

  // ------------------------------------------------------------------ constants / labels
  var BOOK = 'CAEM';
  var STATUS_LABEL = { 'uyğun': 'Uyğundur', 'qismən': 'Qismən uyğun', 'fərqli': 'Fərqlidir', 'tapılmadı': 'Tapılmadı' };
  var TRIVIAL = { 0: 1, 1: 1, 2: 1, 3: 1, 4: 1, 10: 1, 12: 1, 100: 1, 1000: 1, 1000000: 1 };
  var RH = 22, HDR = 22, GUT = 46;
  var KEYS = window.MakroStore.keys, LS_EDITS = KEYS.edits, LS_SCEN = KEYS.scen, LS_ADMIN = KEYS.admin;
  var ARR_SHEET = '__massivlər__';

  function bookTitle() { var m = META.books && META.books[BOOK]; return (m && m.title_az) || 'CAEM proqnoz modeli'; }
  function sheetMeta(gs) { var m = META.books && META.books[BOOK]; return (m && m.sheets && m.sheets[SHEETS[gs].n]) || null; }
  function sheetTitle(gs) { var m = sheetMeta(gs); return (m && m.title_az) || SHEETS[gs].n.trim(); }
  function bookOf(gs) { return BOOKS[SHEETS[gs].b].n; }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function colName(n) { var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
  function a1(r, c) { return colName(c) + r; }
  function parseA1(s) { var m = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(String(s).toUpperCase()); return m ? { r: +m[2], c: colNum(m[1]) } : null; }
  function fmtInt(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ' '); }
  function toast(msg) { var t = $('#toast'); t.textContent = msg; t.hidden = false; clearTimeout(toast._t); toast._t = setTimeout(function () { t.hidden = true; }, 2800); }
  function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* storage unavailable */ } }
  function fold(s) { return String(s).toLocaleLowerCase('az').replace(/i̇/g, 'i').replace(/[ıİ]/g, 'i').replace(/ə/g, 'e').replace(/ö/g, 'o').replace(/ü/g, 'u').replace(/ğ/g, 'g').replace(/ş/g, 's').replace(/ç/g, 'c'); }

  // ------------------------------------------------------------------ sections (documentation tab-colour groups)
  var GROUPS = [], GROUP_OF = {}, ARR_GS = -1, SIDX_N = {};
  function buildGroups() {
    SHEETS.forEach(function (s, gs) { SIDX_N[s.n] = gs; if (s.n === ARR_SHEET) ARR_GS = gs; });
    var seen = {};
    (META.groups || []).forEach(function (g) {
      var list = (g.sheets || []).map(function (n) { return SIDX_N[n]; }).filter(function (gs) { return gs !== undefined && !seen[gs]; });
      list.forEach(function (gs) { seen[gs] = 1; });
      if (list.length) GROUPS.push({ id: g.id, title: g.title_az, color: g.color || '#8A96A3', desc: g.desc_az || '', sheets: list });
    });
    var rest = [];
    SHEETS.forEach(function (s, gs) { if (!seen[gs] && gs !== ARR_GS && BOOKS[s.b].kind === 'model') rest.push(gs); });
    if (rest.length) GROUPS.push({ id: 'gx', title: 'Digər vərəqlər', color: '#8A96A3', desc: '', sheets: rest });
    GROUPS.forEach(function (g, gi) { g.sheets.forEach(function (gs) { GROUP_OF[gs] = gi; }); });
  }
  function groupColor(gs) { var gi = GROUP_OF[gs]; return gi === undefined ? '#98A2AE' : GROUPS[gi].color; }
  function tabColor(gs) { return SHEETS[gs].tab ? '#' + SHEETS[gs].tab : groupColor(gs); }
  function isShown(gs) { return gs !== ARR_GS && BOOKS[SHEETS[gs].b].kind === 'model'; }

  // ------------------------------------------------------------------ derived indexes
  var SC, TPL_NL, KOWNER, NAMES_OF = {}, NAME_CELL = {}, COEF = null, SWITCH = {}, LOCKED = {};
  function buildIndexes() {
    SC = SHEETS.map(function () { return []; });
    for (var i = 0; i < M.N; i++) SC[M.sh[i]].push(i);
    SC.forEach(function (a) { a.sort(function (x, y) { return M.r[x] - M.r[y] || M.c[x] - M.c[y]; }); });
    TPL_NL = CORE.disp.map(function (d) { return (d.match(/\x01/g) || []).length; });
    KOWNER = new Int32Array(M.K.length).fill(-1);
    for (var j = 0; j < M.order.length; j++) {
      var id = M.order[j], n = TPL_NL[M.ft[id]], kb = M.fk[id];
      for (var q = 0; q < n; q++) KOWNER[kb + q] = id;
    }
    // defined names that point to one cell (coefficients, switches, macro anchors)
    (CORE.names || []).forEach(function (nm) {
      var m = /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]{1,3})\$?(\d+)$/.exec(nm.f || '');
      if (!m) return;
      var gs = SIDX_N[(m[1] || m[2]).replace(/''/g, "'")]; if (gs === undefined) return;
      var x = cellId(gs, +m[4], colNum(m[3])); if (x < 0) return;
      (NAMES_OF[x] || (NAMES_OF[x] = [])).push(nm.n); NAME_CELL[nm.n.toLowerCase()] = x;
    });
    // switches from the inputs catalogue (text cells with a fixed list of options)
    ((META.inputs && META.inputs.groups) || []).forEach(function (g) {
      (g.items || []).forEach(function (it) {
        if (it.kind !== 'switch' || !it.options || !it.options.length) return;
        (it.cells || []).forEach(function (a) { var gs = SIDX_N[it.sheet], rc = parseA1(a); if (gs === undefined || !rc) return; var x = cellId(gs, rc.r, rc.c); if (x >= 0) SWITCH[x] = { options: it.options, label: it.label_az }; });
      });
    });
    var b9 = window.MakroStore.layout ? window.MakroStore.layout.b9 : -1;
    if (b9 >= 0) LOCKED[b9] = 'Proqnozun ilk ili «Model yeniləmə» səhifəsində dəyişdirilir (faktiki məlumat daxil edildikdən sonra).';
  }
  // coefficient cells: constants referenced through a defined name, constants on the Parametrization sheet and the
  // coefficient/add-factor items of the inputs catalogue
  function coefSet() {
    if (COEF) return COEF;
    COEF = new Set(); var D = M.dependents();
    Object.keys(NAMES_OF).forEach(function (k) { var x = +k; if (isConstNum(x) && D[x]) COEF.add(x); });
    var gp = SIDX_N['Parametrization']; if (gp !== undefined) SC[gp].forEach(function (x) { if (isConstNum(x) && D[x]) COEF.add(x); });
    ((META.inputs && META.inputs.groups) || []).forEach(function (g) {
      (g.items || []).forEach(function (it) {
        if (it.kind !== 'coef') return;
        (it.cells || []).forEach(function (a) { var gs = SIDX_N[it.sheet], rc = parseA1(a); if (gs === undefined || !rc) return; var x = cellId(gs, rc.r, rc.c); if (x >= 0 && isConstNum(x)) COEF.add(x); });
      });
    });
    return COEF;
  }
  function isCoefCell(id) { return coefSet().has(id); }
  function isConstNum(id) { return M.ft[id] < 0 && typeof M.V0[id] === 'number'; }
  function isConstText(id) { return M.ft[id] < 0 && typeof M.V0[id] === 'string' && M.V0[id].trim() !== ''; }
  function isEditable(id) { var D = M.dependents(); return M.ft[id] < 0 && !!D[id] && !LOCKED[id] && (typeof M.V0[id] === 'number' || !!SWITCH[id] || isConstText(id)); }
  // array formulas: a member cell reads its element from the hidden group cell
  function isMember(id) { return M.ft[id] >= 0 && CORE.disp[M.ft[id]] === '{AEL}'; }
  function groupOfMember(id) { return M.fr[id][0]; }
  function groupRange(vid) { var g = CORE.agroups[M.r[vid] - 1]; return { gs: g[0], r1: g[1], c1: g[2], r2: g[3], c2: g[4] }; }
  function groupAnchor(vid) { var g = groupRange(vid); return cellId(g.gs, g.r1, g.c1); }
  function realCell(id) { return M.sh[id] === ARR_GS ? groupAnchor(id) : id; }

  // year header detection (per sheet, lazy)
  var YR = {};
  function yearRows(gs) {
    if (YR[gs]) return YR[gs];
    var rows = {};
    SC[gs].forEach(function (id) {
      var v = M.V0[id], y = null;
      if (typeof v === 'number' && v === Math.floor(v) && v >= 1980 && v <= 2060) y = v;
      else if (typeof v === 'string') { var m = /^\s*((?:19|20)\d\d)\s*\D{0,12}$/.exec(v); if (m) y = +m[1]; }
      if (y) (rows[M.r[id]] || (rows[M.r[id]] = {}))[M.c[id]] = y;
    });
    var out = [];
    Object.keys(rows).forEach(function (r) {
      var d = rows[r], cs = Object.keys(d).map(Number).sort(function (a, b) { return a - b; }), run = 0;
      for (var i = 1; i < cs.length; i++) if (cs[i] === cs[i - 1] + 1 && d[cs[i]] === d[cs[i - 1]] + 1) run++;
      if (run >= 3) out.push({ r: +r, d: d });
    });
    out.sort(function (a, b) { return a.r - b.r; });
    return (YR[gs] = out);
  }
  function colYear(gs, r, c) {
    var yr = yearRows(gs), best = null, first = null;
    for (var i = 0; i < yr.length; i++) {
      if (yr[i].d[c] === undefined) continue;
      if (first === null) first = yr[i].d[c];
      if (yr[i].r <= r) best = yr[i].d[c];
    }
    return best !== null ? best : first;
  }
  function rowLabel(gs, r) {
    var ids = rowCells(gs, r), firstNum = 1e9, parts = [];
    for (var i = 0; i < ids.length; i++) if (typeof M.V[ids[i]] === 'number' && M.c[ids[i]] > 2) { firstNum = M.c[ids[i]]; break; }
    for (i = 0; i < ids.length && parts.length < 3; i++) {
      var v = M.V[ids[i]];
      if (M.c[ids[i]] >= firstNum) break;
      if (typeof v === 'string' && v.trim() && !/^\s*[\d.,\-%]+\s*$/.test(v)) parts.push(v.trim().slice(0, 90));
    }
    return parts.join(' · ');
  }
  var ROWIDX = {};
  function rowCells(gs, r) {
    var idx = ROWIDX[gs];
    if (!idx) { idx = ROWIDX[gs] = {}; SC[gs].forEach(function (id) { (idx[M.r[id]] || (idx[M.r[id]] = [])).push(id); }); }
    return idx[r] || [];
  }
  function cellId(gs, r, c) { var x = M.smap[gs].get(r * 20000 + c); return x === undefined ? -1 : x; }
  function cellKey(id) { return bookOf(M.sh[id]) + '|' + SHEETS[M.sh[id]].n + '|' + a1(M.r[id], M.c[id]); }
  function idFromKey(k) {
    var p = k.split('|'); if (p.length < 3) return -1;
    var gs = sheetIndex(p[0], p[1]); if (gs < 0) return -1;
    var rc = parseA1(p[2]); return rc ? cellId(gs, rc.r, rc.c) : -1;
  }
  var SIDX = null;
  function sheetIndex(b, s) {
    if (!SIDX) { SIDX = {}; SHEETS.forEach(function (sh, i) { SIDX[BOOKS[sh.b].n + '|' + sh.n] = i; }); }
    var v = SIDX[b + '|' + s]; return v === undefined ? -1 : v;
  }
  function fullAddr(id) { id = realCell(id); return SHEETS[M.sh[id]].n + '!' + a1(M.r[id], M.c[id]); }

  // ------------------------------------------------------------------ value formatting
  function fmtGeneral(v) {
    if (v === 0) return '0';
    var a = Math.abs(v);
    if (a >= 1e11 || a < 1e-6) return v.toExponential(4).replace('e', 'E');
    var intD = Math.max(1, Math.floor(Math.log10(a)) + 1), dec = Math.max(0, Math.min(10 - intD, 9));
    return String(Number(v.toFixed(dec)));
  }
  function fmtNum(v, fmt) {
    if (!fmt || fmt === 'General' || fmt === '@') return fmtGeneral(v);
    try { var s = X.fmt(v, fmt); return s === '' && v !== 0 ? fmtGeneral(v) : s; } catch (e) { return fmtGeneral(v); }
  }
  function dispVal(v, st) {
    if (v === null || v === undefined) return '';
    if (v instanceof X.XErr) return v.e;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'string') return v;
    if (v && v.isArr) return '{massiv ' + v.r + '×' + v.c + '}';
    return fmtNum(v, st ? st.fmt : null);
  }
  function fullVal(v) {
    if (v === null || v === undefined) return '(boş)';
    if (v instanceof X.XErr) return v.e;
    if (typeof v === 'number') return String(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (v && v.isArr) return '{massiv ' + v.r + '×' + v.c + '}';
    return '"' + v + '"';
  }
  function shortVal(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof X.XErr) return v.e;
    if (typeof v === 'number') return fmtGeneral(v);
    if (v && v.isArr) return '{' + v.r + '×' + v.c + '}';
    return String(v).slice(0, 40);
  }
  function sameVal(a, b) {
    if (typeof a === 'number' && typeof b === 'number') { var d = Math.abs(a - b); return d <= 1e-9 || d <= 1e-9 * Math.max(Math.abs(a), Math.abs(b)); }
    if (a instanceof X.XErr && b instanceof X.XErr) return a.e === b.e;
    if ((a === null || a === '') && (b === null || b === '')) return true;
    if (a && b && a.isArr && b.isArr) { if (a.d.length !== b.d.length) return false; for (var i = 0; i < a.d.length; i++) if (!sameVal(a.d[i], b.d[i])) return false; return true; }
    return a === b;
  }
  function parseUserNumber(s) {
    s = String(s).trim().replace(/[\s ]/g, '').replace('−', '-');
    if (s === '') return null;
    if (/^[+-]?\d*,\d+$/.test(s)) s = s.replace(',', '.');
    var pct = /%$/.test(s); if (pct) s = s.slice(0, -1);
    if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) return undefined;
    var x = parseFloat(s); return pct ? x / 100 : x;
  }

  // ------------------------------------------------------------------ edits (admin): inputs, coefficients (literals), overrides
  var EDITS = { inputs: new Map(), K: new Map(), ov: new Map() };   // id -> value ; kIndex -> value ; id -> value
  function persist() {
    var o = { inputs: {}, K: {}, ov: {} };
    EDITS.inputs.forEach(function (v, id) { o.inputs[cellKey(id)] = v; });
    EDITS.K.forEach(function (v, k) { var id = KOWNER[k]; o.K[cellKey(id) + '#' + (k - M.fk[id])] = v; });
    EDITS.ov.forEach(function (v, id) { o.ov[cellKey(id)] = v; });
    lsSet(LS_EDITS, JSON.stringify(o));
    updateChangeCount();
  }
  function serialize() {
    var o = { format: 'caem-model-edits', version: 1, saved: new Date().toISOString(), firstYear: INFO.firstYear, inputs: {}, coefficients: {}, overrides: {} };
    EDITS.inputs.forEach(function (v, id) { o.inputs[cellKey(id)] = v; });
    EDITS.K.forEach(function (v, k) { var id = KOWNER[k]; o.coefficients[cellKey(id) + '#' + (k - M.fk[id])] = v; });
    EDITS.ov.forEach(function (v, id) { o.overrides[cellKey(id)] = v; });
    return o;
  }
  function applySerialized(o, silent) {
    var changed = [], bad = 0;
    clearAllEdits(true);
    var inputs = o.inputs || {}, coefs = o.K || o.coefficients || {}, ovs = o.ov || o.overrides || {};
    Object.keys(inputs).forEach(function (k) { var id = idFromKey(k); if (id < 0 || M.ft[id] >= 0 || LOCKED[id]) { bad++; return; } EDITS.inputs.set(id, inputs[k]); M.V[id] = inputs[k]; changed.push(id); });
    Object.keys(coefs).forEach(function (k) {
      var p = k.lastIndexOf('#'), id = idFromKey(k.slice(0, p)), i = +k.slice(p + 1);
      if (id < 0 || M.ft[id] < 0 || i >= TPL_NL[M.ft[id]]) { bad++; return; }
      var ki = M.fk[id] + i; EDITS.K.set(ki, coefs[k]); M.K[ki] = coefs[k]; changed.push(id);
    });
    Object.keys(ovs).forEach(function (k) { var id = idFromKey(k); if (id < 0 || M.ft[id] < 0) { bad++; return; } EDITS.ov.set(id, ovs[k]); changed.push(id); });
    if (changed.length) M.recalcFrom(changed, EDITS.ov);
    persist();
    if (!silent) toast('Tətbiq edildi: ' + changed.length + ' dəyişiklik' + (bad ? ' · ' + bad + ' tanınmayan açar ötürüldü' : ''));
    return bad;
  }
  function clearAllEdits(noRecalc) {
    var changed = [];
    EDITS.inputs.forEach(function (v, id) { M.V[id] = M.V0[id]; changed.push(id); });
    EDITS.K.forEach(function (v, k) { M.K[k] = M.K0[k]; changed.push(KOWNER[k]); });
    EDITS.ov.forEach(function (v, id) { changed.push(id); });
    EDITS.inputs.clear(); EDITS.K.clear(); EDITS.ov.clear();
    if (!noRecalc && changed.length) M.recalcFrom(changed, EDITS.ov);
    return changed.length;
  }
  function editCount() { return EDITS.inputs.size + EDITS.K.size + EDITS.ov.size; }
  function updateChangeCount() {
    var el = $('#change-count'), n = editCount();
    el.hidden = n === 0; el.textContent = n + ' dəyişiklik';
    el.title = 'Admin panelinə keçid'; el.style.cursor = 'pointer'; el.onclick = function () { location.hash = '#/admin'; };
  }
  function setInput(id, v) {
    if (LOCKED[id]) { toast(LOCKED[id]); return; }
    if (sameVal(v, M.V0[id])) EDITS.inputs.delete(id); else EDITS.inputs.set(id, v);
    M.V[id] = v; var n = M.recalcFrom([id], EDITS.ov); persist(); afterChange(n);
  }
  function setCoef(kIdxs, v) {
    var owners = [];
    kIdxs.forEach(function (k) { M.K[k] = v; if (sameVal(v, M.K0[k])) EDITS.K.delete(k); else EDITS.K.set(k, v); owners.push(KOWNER[k]); });
    var n = M.recalcFrom(owners, EDITS.ov); persist(); afterChange(n);
  }
  function setOverride(id, v) {
    if (v === undefined) EDITS.ov.delete(id); else EDITS.ov.set(id, v);
    var n = M.recalcFrom([id], EDITS.ov); persist(); afterChange(n);
  }
  function afterChange(n) {
    toast('Yenidən hesablandı: ' + fmtInt(n) + ' düstur' + cycNote());
    if (GRID) GRID.render(true);
    if (SEL) renderInspector();
    refreshEditorValues();
  }
  function cycNote() {
    var bad = (M.cycInfo || []).filter(function (c) { return c && !c.converged; }).length;
    return bad ? ' · diqqət: ' + bad + ' dövri blok yaxınlaşmadı' : '';
  }

  // ------------------------------------------------------------------ formula rendering
  function quoteSheet(n) { return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) && !/^[A-Z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'"; }
  function refA1(r, c, flags) { return (flags.indexOf('C') >= 0 ? '$' : '') + colName(c) + (flags.indexOf('R') >= 0 ? '$' : '') + r; }
  function prefixFor(curGs, gs2, explicit) {
    if (bookOf(gs2) !== bookOf(curGs)) return "'[" + bookOf(gs2).replace(/^EXT: /, '') + ']' + SHEETS[gs2].n.replace(/'/g, "''") + "'!";
    return explicit ? quoteSheet(SHEETS[gs2].n) + '!' : '';
  }
  // returns HTML; opts.hlLit = literal index to highlight; opts.plain = text only
  function formulaHTML(id, opts) {
    opts = opts || {};
    if (isMember(id)) {                       // element of a multi-cell array formula: show the array formula
      var vid = groupOfMember(id), g = groupRange(vid);
      var inner = formulaHTML(vid, opts), rng = a1(g.r1, g.c1) + ':' + a1(g.r2, g.c2);
      return opts.plain ? inner + '  [' + rng + ']' : inner + ' <span class="muted small">— ' + esc(rng) + ' massivinin elementi</span>';
    }
    var d = CORE.disp[M.ft[id]], refs = M.fr[id], kb = M.fk[id], cur = M.sh[id] === ARR_GS ? groupRange(id).gs : M.sh[id];
    var out = '', li = 0, ri = 0, i = 0, txt = '';
    while (i < d.length) {
      var ch = d.charAt(i);
      if (ch === '\x01') {
        var k = kb + li, v = M.K[k], mod = !sameVal(v, M.K0[k]);
        var s = String(v);
        if (opts.plain) txt += s;
        else if (opts.hlLit !== undefined) out += li === opts.hlLit ? '<span class="hl">' + esc(s) + '</span>' : esc(s);
        else out += '<span class="lit' + (mod ? ' mod' : '') + '" data-k="' + k + '" title="' + (mod ? INFO.baseShort + ': ' + M.K0[k] + ' · ' : '') + 'Düsturdakı əmsal' + (ADMIN ? ' — dəyişmək üçün klikləyin' : '') + '">' + esc(s) + '</span>';
        li++; i++; continue;
      }
      if (ch === '\x02') {
        var j = d.indexOf('\x03', i), flags = d.slice(i + 1, j), ref = refs[ri++], explicit = flags.charAt(0) === 'P';
        if (explicit) flags = flags.slice(1);
        var text, target, gs2;
        if (ref >= 0) {
          gs2 = M.sh[ref]; text = prefixFor(cur, gs2, explicit) + refA1(M.r[ref], M.c[ref], flags); target = ref;
        } else {
          var R = CORE.ranges[-ref - 1], fl = flags.slice(1).split(':');
          gs2 = R[0]; text = prefixFor(cur, gs2, explicit) + refA1(R[1], R[2], fl[0] || '') + ':' + refA1(R[3], R[4], fl[1] || '');
          if (opts.plain) { txt += text; i = j + 1; continue; }
          out += '<a class="ref' + (gs2 === cur ? ' local' : '') + '" data-gs="' + gs2 + '" data-r="' + R[1] + '" data-c="' + R[2] + '" title="Diapazon ' + esc(SHEETS[gs2].n) + '">' + esc(text) + '</a>';
          i = j + 1; continue;
        }
        if (opts.plain) txt += text;
        else if (opts.hlLit !== undefined) out += esc(text);
        else {
          var lab = rowLabel(gs2, M.r[target]), yr = colYear(gs2, M.r[target], M.c[target]);
          out += '<a class="ref' + (gs2 === cur ? ' local' : '') + '" data-gs="' + gs2 + '" data-r="' + M.r[target] + '" data-c="' + M.c[target] + '" title="' +
            esc((lab ? lab + ' · ' : '') + (yr ? yr + ' · ' : '') + '= ' + shortVal(M.V[target])) + '">' + esc(text) + '</a>';
        }
        i = j + 1; continue;
      }
      var k2 = i; while (k2 < d.length && d.charAt(k2) !== '\x01' && d.charAt(k2) !== '\x02') k2++;
      var chunk = d.slice(i, k2);
      if (opts.plain) txt += chunk; else out += opts.hlLit !== undefined ? esc(chunk) : nameLinks(chunk);
      i = k2;
    }
    var body = opts.plain ? txt : out;
    // array formulas are displayed as Excel does: {=…}
    if (d.charAt(0) === '{' && d.charAt(d.length - 1) === '}') return '{=' + body.slice(opts.plain ? 1 : 1, body.length - 1) + '}';
    return '=' + body;
  }
  // defined names inside formula text become links to the cell they name
  function nameLinks(chunk) {
    return esc(chunk).replace(/[A-Za-z_][A-Za-z0-9_.]*/g, function (w) {
      var x = NAME_CELL[w.toLowerCase()];
      if (x === undefined) return w;
      return '<a class="ref nm" data-gs="' + M.sh[x] + '" data-r="' + M.r[x] + '" data-c="' + M.c[x] + '" title="Ad: ' + esc(w) + ' → ' + esc(fullAddr(x)) + ' = ' + esc(shortVal(M.V[x])) + '">' + w + '</a>';
    });
  }

  // ------------------------------------------------------------------ routing
  var ROUTE = {}, GRID = null, SEL = null, ADMIN = false;
  function go(h) { if (location.hash !== h) location.hash = h; else route(); }
  function route() {
    var h = decodeURIComponent(location.hash.replace(/^#\/?/, '')), p = h.split('/');
    ROUTE = { name: p[0] || 'overview', a: p[1], b: p[2], c: p[3] };
    document.body.classList.remove('nav-open');
    if (GRID) { GRID.destroy(); GRID = null; }
    main.classList.remove('grid-mode');
    main.scrollTop = 0;
    if (ROUTE.name !== 's' || (ROUTE.b && ROUTE.b !== 'cedvel')) closeInspector();
    switch (ROUTE.name) {
      case 'g': pageGroup(+ROUTE.a, ROUTE.b || 'tesvir'); break;
      case 's': if (SHEETS[+ROUTE.a] && isShown(+ROUTE.a)) pageSheet(+ROUTE.a, ROUTE.b || 'cedvel', ROUTE.c); else pageOverview(); break;
      case 'metod': pageMethod(); break;
      case 'elaqeler': pageLinks(); break;
      case 'admin': pageAdmin(); break;
      case 'yenile': pageUpdate(); break;
      default: pageOverview();
    }
    renderNav();
  }

  // ------------------------------------------------------------------ sidebar
  var OPEN = {};
  function renderNav() {
    var h = '<div class="nav-sec">Model</div>';
    h += navLink('#/', 'Ümumi baxış və metodologiya', ROUTE.name === 'overview', '#3B4F9A');
    h += navLink('#/elaqeler', 'Vərəqlər arası əlaqələr', ROUTE.name === 'elaqeler', '#1F7A4D');
    h += navLink('#/metod', 'Sənədlə yoxlama', ROUTE.name === 'metod', '#6A3FB5');
    h += navLink('#/admin', 'Admin paneli və makroslar', ROUTE.name === 'admin', '#B3261E');
    h += navLink('#/yenile', 'Model yeniləmə (il, fayl)', ROUTE.name === 'yenile', '#8A5300');
    h += '<a class="nav-link" href="panel.html"><span class="dot" style="background:#C8A13A"></span>İş paneli (sadə görünüş) ↗</a>';
    h += '<div class="nav-sec">Bölmələr (vərəq qrupları)</div>';
    var curG = ROUTE.name === 'g' ? +ROUTE.a : ROUTE.name === 's' ? GROUP_OF[+ROUTE.a] : -1;
    GROUPS.forEach(function (g, gi) {
      var open = OPEN[gi] || curG === gi;
      h += '<a class="nav-book' + (open ? ' open' : '') + (ROUTE.name === 'g' && +ROUTE.a === gi ? ' active' : '') + '" href="#/g/' + gi + '" data-book="' + gi + '">' +
        '<span class="dot" style="background:' + g.color + '"></span><span>' + esc(g.title) + '<small>' + g.sheets.length + ' vərəq</small></span><span class="caret" data-toggle="' + gi + '">▶</span></a>';
      if (open) {
        h += '<div class="nav-sheets">';
        g.sheets.forEach(function (gs) {
          var act = ROUTE.name === 's' && +ROUTE.a === gs, t = sheetTitle(gs);
          h += '<a class="nav-sheet' + (act ? ' active' : '') + '" href="#/s/' + gs + '/cedvel"><span class="tabdot" style="background:' + tabColor(gs) + '"></span>' + esc(t) + (t !== SHEETS[gs].n.trim() ? '<small>' + esc(SHEETS[gs].n) + '</small>' : '') + (SHEETS[gs].st !== 'visible' ? ' <small>(gizli)</small>' : '') + '</a>';
        });
        h += '</div>';
      }
    });
    sidebar.innerHTML = h;
    var act = sidebar.querySelector('.nav-sheet.active');
    if (act && act.scrollIntoView) { var rb = act.getBoundingClientRect(), sb = sidebar.getBoundingClientRect(); if (rb.top < sb.top || rb.bottom > sb.bottom) act.scrollIntoView({ block: 'center' }); }
  }
  function navLink(href, label, active, color) { return '<a class="nav-link' + (active ? ' active' : '') + '" href="' + href + '"><span class="dot" style="background:' + color + '"></span>' + esc(label) + '</a>'; }

  // ------------------------------------------------------------------ shared bits
  function statusChip(s) { return '<span class="chip status-' + esc(s) + '">' + esc(STATUS_LABEL[s] || s) + '</span>'; }
  function sheetStats(gs) {
    var f = 0, c = 0; SC[gs].forEach(function (id) { if (M.ft[id] >= 0) f++; else if (M.V0[id] !== null) c++; });
    return { cells: SC[gs].length, formulas: f, consts: c };
  }
  function tabsHTML(base, tabs, cur) {
    return '<div class="tabs" role="tablist">' + tabs.map(function (t) {
      return '<a class="tab' + (t[0] === cur ? ' active' : '') + '" role="tab" href="' + base + t[0] + '">' + esc(t[1]) + (t[2] !== undefined ? '<span class="cnt">' + t[2] + '</span>' : '') + '</a>';
    }).join('') + '</div>';
  }
  // "Sheet name!A1" / "'Sheet name'!A1:B2" (sheet names from this workbook) -> links
  var SHEET_RE = null;
  function sheetRe() {
    if (SHEET_RE) return SHEET_RE;
    var names = SHEETS.filter(function (s, gs) { return isShown(gs); }).map(function (s) { return s.n; }).sort(function (a, b) { return b.length - a.length; });
    var alt = names.map(function (n) { return esc(n).replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/ +/g, ' +'); }).join('|');
    SHEET_RE = new RegExp("(?:&#39;)?(" + alt + ")(?:&#39;)?!(\\$?[A-Z]{1,3}\\$?\\d+(?::\\$?[A-Z]{1,3}\\$?\\d+)?)", 'g');
    return SHEET_RE;
  }
  function sheetByLooseName(s) {
    s = s.replace(/&amp;/g, '&').replace(/&#39;/g, "'");
    if (SIDX_N[s] !== undefined) return SIDX_N[s];
    var k = s.replace(/\s+/g, ' ').trim(), hit = -1;
    SHEETS.forEach(function (sh, gs) { if (hit < 0 && isShown(gs) && sh.n.replace(/\s+/g, ' ').trim() === k) hit = gs; });
    return hit;
  }
  function linkify(text) {
    var re = sheetRe(); re.lastIndex = 0;
    return esc(text).replace(re, function (all, s, a) {
      var gs = sheetByLooseName(s); if (gs < 0) return all;
      var first = a.split(':')[0].replace(/\$/g, '');
      return '<a href="#/s/' + gs + '/cedvel/' + first + '"><code class="f">' + all + '</code></a>';
    });
  }
  function addrLink(s) { var h = linkify(s); return h.indexOf('<a ') >= 0 ? h : '<code class="f">' + h + '</code>'; }

  // ------------------------------------------------------------------ OVERVIEW
  function pageOverview() {
    var O = META.overview || {}, V = META.verify || {};
    var h = '<div class="eyebrow">Ümumi baxış</div><div class="page-h"><div><h1>CAEM proqnoz modeli: quruluş, əlaqələr və metodologiya</h1>' +
      '<p class="lead">' + linkify(O.lead || '') + '</p></div></div>';
    if (INFO.custom) h += '<div class="note note-warn" style="margin-top:14px"><b>Yüklənmiş model istifadə olunur</b> (' + esc(window.CaemLocalTime ? window.CaemLocalTime(INFO.custom.date) : '') + ': ' + esc((INFO.custom.files || []).map(function (f) { return f.name; }).join(', ')) + '). Yoxlama nəticələri və əlaqə cədvəlləri yeni fayl üzrə yenidən hesablanıb; vərəq təsvirləri, sənədlə tutuşdurma və məsələlər siyahısı isə orijinal fayla aiddir. <a href="#/yenile">Model yeniləmə</a></div>';
    if (INFO.rolls || INFO.actuals) h += '<div class="note" style="margin-top:10px">Baza model: ' + (INFO.rolls ? 'proqnozun ilk ili ' + INFO.baseFirst + ' → <b>' + INFO.firstYear + '</b> (' + INFO.rolls + ' il irəli çəkilib); ' : '') + (INFO.actuals ? INFO.actuals + ' faktiki/baza dəyəri daxil edilib; ' : '') + 'proqnoz üfüqü ' + INFO.firstYear + '–' + INFO.endYear + '. <a href="#/yenile">Model yeniləmə</a></div>';
    h += '<div class="stats">' +
      stat(fmtInt(V.visible || 0) + ' / ' + fmtInt(V.sheets || 0), 'vərəq (görünən / cəmi)') + stat(fmtInt(V.formulas || 0), 'düstur') +
      stat(fmtInt(V.arrays || 0), 'massiv düsturu (' + fmtInt(V.arrayCells || 0) + ' xana)') + stat(fmtInt(V.names || 0), 'adlandırılmış xana/diapazon') +
      stat(INFO.firstYear + '–' + INFO.endYear, 'proqnoz üfüqü') +
      stat((V.local_ok !== undefined ? (100 * V.local_ok / V.formulas).toFixed(3) + '%' : '—'), 'düstur Excel ilə üst-üstə düşür') + '</div>';
    if (O.sections) O.sections.forEach(function (s) { h += '<section class="block"><h2>' + esc(s.h) + '</h2>' + (s.html || '') + '</section>'; });
    h += '<section class="block"><h2>Bölmələr arasında məlumat axını <small>ox istiqaməti: mənbə → istifadəçi · qalınlıq istinad sayına mütənasibdir</small></h2><div class="diagram-wrap">' + groupDiagram() + '</div>' +
      '<p class="small muted">Qovşağın üzərinə gəlin — yalnız həmin bölmənin əlaqələri qalır. Qovşağa klikləyin — bölmənin səhifəsinə keçid.</p></section>';
    h += '<section class="block"><h2>Əlaqə matrisi <small>sətir: mənbə bölmə · sütun: istifadə edən bölmə · xanada düsturlardakı istinad sayı</small></h2>' + groupMatrix() + '</section>';
    h += '<section class="block"><h2>Yoxlama nəticələri</h2>' + verifyHTML() + '</section>';
    if (META.issues && META.issues.length) {
      var hi = META.issues.filter(function (x) { return x.severity === 'yüksək'; });
      h += '<section class="block"><h2>Yüksək əhəmiyyətli məsələlər <small>' + hi.length + ' / cəmi ' + META.issues.length + '</small></h2>' + issuesHTML(hi) +
        '<p class="small" style="margin-top:8px"><a href="#/metod#issues">Bütün ' + META.issues.length + ' məsələ — «Sənədlə yoxlama» səhifəsində</a></p></section>';
    }
    h += '<section class="block"><h2>Bölmələr</h2><div class="cards">' + GROUPS.map(function (g, gi) {
      var f = 0; g.sheets.forEach(function (gs) { f += sheetStats(gs).formulas; });
      return '<a class="card" href="#/g/' + gi + '"><div class="meta"><span class="dot" style="background:' + g.color + '"></span>' + g.sheets.length + ' vərəq · ' + fmtInt(f) + ' düstur</div><h3>' + esc(g.title) + '</h3><p>' + esc((g.desc || '').split(/(?<=\.)\s/)[0]) + '</p></a>';
    }).join('') + '</div></section>';
    main.innerHTML = h;
    wireDiagram();
  }
  function stat(v, l) { return '<div class="stat"><b>' + v + '</b><span>' + esc(l) + '</span></div>'; }
  function verifyHTML() {
    var V = META.verify || {};
    if (!V.formulas) return '<p class="muted">Yoxlama məlumatı yoxdur.</p>';
    var cyc = (V.cyc || []).reduce(function (a, b) { return a + b; }, 0);
    var h = '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Yoxlama</th><th class="n">Uyğun</th><th class="n">Fərqli</th><th>İzah</th></tr></thead><tbody>';
    h += '<tr><td><b>1. Düstur-düstur yoxlama</b></td><td class="n">' + fmtInt(V.local_ok) + '</td><td class="n">' + fmtInt(V.formulas - V.local_ok) + '</td><td>Hər düstur (' + fmtInt(V.arrayCells || 0) + ' massiv elementi daxil) Excel-in saxladığı giriş dəyərləri ilə bu saytın hesablama mühərriki tərəfindən yenidən hesablanıb və Excel-in saxladığı nəticə ilə müqayisə edilib (nisbi dəqiqlik 10⁻⁹).</td></tr>';
    h += '<tr><td><b>2. Tam yenidən hesablama</b></td><td class="n">' + fmtInt(V.global_ok) + '</td><td class="n">' + fmtInt(V.formulas - V.global_ok) + '</td><td>Bütün model yalnız sabit giriş məlumatlarından başlayaraq asılılıq ardıcıllığı ilə hesablanıb: HP filtrləri (VBA funksiyaları HPF/HPP JavaScript-ə köçürülüb), LINEST/MINVERSE/MMULT reqressiya və matris hesabları, INDEX:INDEX dinamik diapazonları daxil.</td></tr>';
    h += '<tr><td><b>3. Dövri asılılıqlar</b></td><td class="n">' + (V.cyc ? V.cyc.length : 0) + ' blok</td><td class="n">—</td><td>' + fmtInt(cyc) + ' xana (inflyasiya — faiz dərəcəsi — məzənnə) yalnız «üzən məzənnə» rejimində həqiqi qarşılıqlı asılılıq yaradır; «sabit məzənnə» rejimində (INPUT!C9) keçid açarı dövrü qırır və Excel-dəki kimi dəqiq hesablanır. Üzən rejimdə bloklar ardıcıl yaxınlaşma ilə (Excel-in iterativ hesablaması kimi) həll olunur.</td></tr>';
    h += '</tbody></table></div>';
    if (V.diffs && V.diffs.length) h += '<div class="note note-warn" style="margin-top:10px"><b>Fərqlənən xanalar.</b> ' + V.diffs.map(function (d) { return linkify(d); }).join('<br>') + '</div>';
    return h;
  }
  function issuesHTML(list) {
    var order = { 'yüksək': 0, 'orta': 1, 'aşağı': 2 };
    list = list.slice().sort(function (a, b) { return (order[a.severity] === undefined ? 3 : order[a.severity]) - (order[b.severity] === undefined ? 3 : order[b.severity]); });
    return '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Dərəcə</th><th>Yer</th><th>Məsələ</th></tr></thead><tbody>' + list.map(function (x) {
      return '<tr><td><span class="chip sev-' + esc(x.severity) + '">' + esc(x.severity) + '</span></td><td>' + linkify(x.where) + '</td><td>' + linkify(x.issue_az) + '</td></tr>';
    }).join('') + '</tbody></table></div>';
  }
  function groupLinks() {
    if (groupLinks.c) return groupLinks.c;
    var m = {};
    (META.sheetLinks || []).forEach(function (e) {
      var a = GROUP_OF[SIDX_N[e.ss]], b = GROUP_OF[SIDX_N[e.ds]]; if (a === undefined || b === undefined || a === b) return;
      var k = a + '>' + b; (m[k] || (m[k] = { a: a, b: b, n: 0, formulas: 0 })); m[k].n += e.n; m[k].formulas += e.formulas;
    });
    return (groupLinks.c = Object.keys(m).map(function (k) { return m[k]; }));
  }
  function groupDiagram() {
    var L = groupLinks(), n = GROUPS.length, W = 820, H = 600, cx = W / 2, cy = H / 2 + 4, R = 225;
    var pos = GROUPS.map(function (g, i) { var a = -Math.PI / 2 + 2 * Math.PI * i / n; return [cx + R * Math.cos(a), cy + R * Math.sin(a), a]; });
    var mx = 1; L.forEach(function (e) { if (e.n > mx) mx = e.n; });
    var s = '<svg id="bookdiag" viewBox="0 0 ' + W + ' ' + H + '" width="' + W + '" role="img" aria-label="Bölmələr arasında əlaqə diaqramı"><defs>';
    GROUPS.forEach(function (g, i) { s += '<marker id="ar' + i + '" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="9" markerHeight="9" markerUnits="userSpaceOnUse" orient="auto-start-reverse"><path d="M0,0L10,5L0,10z" fill="' + g.color + '"/></marker>'; });
    s += '</defs><g class="edges">';
    L.forEach(function (e) {
      var a = pos[e.a], b = pos[e.b], dx = b[0] - a[0], dy = b[1] - a[1], len = Math.sqrt(dx * dx + dy * dy), ux = dx / len, uy = dy / len;
      var x1 = a[0] + ux * 30, y1 = a[1] + uy * 30, x2 = b[0] - ux * 34, y2 = b[1] - uy * 34, mxp = (x1 + x2) / 2 - uy * 24, myp = (y1 + y2) / 2 + ux * 24;
      var w = 0.8 + 5.2 * Math.sqrt(e.n / mx);
      s += '<path d="M' + x1.toFixed(1) + ',' + y1.toFixed(1) + ' Q' + mxp.toFixed(1) + ',' + myp.toFixed(1) + ' ' + x2.toFixed(1) + ',' + y2.toFixed(1) + '" fill="none" stroke="' + GROUPS[e.a].color + '" stroke-opacity=".55" stroke-width="' + w.toFixed(2) + '" marker-end="url(#ar' + e.a + ')" data-s="' + e.a + '" data-d="' + e.b + '"><title>' + esc(GROUPS[e.a].title + ' → ' + GROUPS[e.b].title + ': ' + fmtInt(e.n) + ' istinad, ' + fmtInt(e.formulas) + ' düstur') + '</title></path>';
    });
    s += '</g><g class="nodes">';
    GROUPS.forEach(function (g, i) {
      var p = pos[i], tx = p[0] + Math.cos(p[2]) * 38, ty = p[1] + Math.sin(p[2]) * 38 + 4, anchor = Math.abs(Math.cos(p[2])) < 0.3 ? 'middle' : Math.cos(p[2]) > 0 ? 'start' : 'end';
      var words = g.title.split(' '), l1 = '', l2 = '';
      words.forEach(function (w) { if ((l1 + ' ' + w).trim().length <= 24 && !l2) l1 = (l1 + ' ' + w).trim(); else l2 = (l2 + ' ' + w).trim(); });
      s += '<g class="node" data-b="' + i + '" data-bi="' + i + '" style="cursor:pointer"><circle cx="' + p[0].toFixed(1) + '" cy="' + p[1].toFixed(1) + '" r="24" fill="#FFFFFF" stroke="' + g.color + '" stroke-width="3"/>' +
        '<text x="' + p[0].toFixed(1) + '" y="' + (p[1] + 4).toFixed(1) + '" text-anchor="middle" font-size="12" font-weight="700" fill="#44525F">' + (i + 1) + '</text>' +
        '<text x="' + tx.toFixed(1) + '" y="' + ty.toFixed(1) + '" text-anchor="' + anchor + '" font-size="12" fill="#44525F">' + esc(l1) + (l2 ? '<tspan x="' + tx.toFixed(1) + '" dy="14">' + esc(l2.length > 30 ? l2.slice(0, 29) + '…' : l2) + '</tspan>' : '') + '</text></g>';
    });
    return s + '</g></svg>';
  }
  function wireDiagram() {
    var svg = document.getElementById('bookdiag'); if (!svg) return;
    svg.addEventListener('mouseover', function (e) {
      var g = e.target.closest('.node'); if (!g) return; var b = g.getAttribute('data-b');
      svg.querySelectorAll('.edges path').forEach(function (p) { var on = p.getAttribute('data-s') === b || p.getAttribute('data-d') === b; p.style.opacity = on ? 1 : 0.06; });
    });
    svg.addEventListener('mouseout', function (e) { if (e.target.closest('.node')) svg.querySelectorAll('.edges path').forEach(function (p) { p.style.opacity = ''; }); });
    svg.addEventListener('click', function (e) { var g = e.target.closest('.node'); if (g) go('#/g/' + g.getAttribute('data-bi')); });
  }
  function groupMatrix() {
    var L = groupLinks(), m = {}, mx = 1;
    L.forEach(function (e) { m[e.a + '>' + e.b] = e; if (e.n > mx) mx = e.n; });
    var h = '<div class="tbl-wrap"><table class="tbl mx"><thead><tr><th>Mənbə ↓ / İstifadəçi →</th>' + GROUPS.map(function (g, i) { return '<th class="n" title="' + esc(g.title) + '">' + (i + 1) + '</th>'; }).join('') + '</tr></thead><tbody>';
    GROUPS.forEach(function (g, a) {
      h += '<tr><th style="position:static;text-transform:none;letter-spacing:0;font-size:12.5px;color:var(--ink)"><span class="dot" style="display:inline-block;background:' + g.color + ';margin-right:6px"></span>' + (a + 1) + '. ' + esc(g.title) + '</th>';
      GROUPS.forEach(function (g2, b) {
        if (a === b) { h += '<td class="cellv muted">—</td>'; return; }
        var e = m[a + '>' + b]; if (!e) { h += '<td class="cellv"></td>'; return; }
        var al = Math.max(0.08, Math.sqrt(e.n / mx));
        h += '<td class="cellv" style="background:rgba(59,79,154,' + (al * 0.55).toFixed(3) + ');' + (al > 0.6 ? 'color:#fff' : '') + '" title="' + esc(g.title + ' → ' + g2.title + ': ' + e.n + ' istinad · ' + e.formulas + ' düstur') + '">' + fmtInt(e.n) + '</td>';
      });
      h += '</tr>';
    });
    return h + '</tbody></table></div>';
  }

  // ------------------------------------------------------------------ LINKS page
  function pageLinks() {
    var L = META.sheetLinks || [];
    var h = '<div class="eyebrow">Əlaqələr</div><div class="page-h"><div><h1>Vərəqlər arası əlaqələr</h1><p class="lead">Hər sətir bir vərəqdən digər vərəqə gedən istinadları göstərir: neçə düstur həmin mənbədən oxuyur və mənbədə neçə xana istifadə olunur (adlandırılmış xanalar və massiv düsturları daxil). Vərəq adına klikləyin — cədvələ keçid.</p></div></div>';
    h += '<section class="block"><h2>Bölmələr üzrə əlaqə matrisi</h2>' + groupMatrix() + '</section>';
    h += '<section class="block"><h2>Vərəq səviyyəsində keçidlər <small>' + L.length + ' əlaqə</small></h2><div class="filters"><label for="lf">Süzgəc</label><input id="lf" type="search" placeholder="vərəq adı" style="height:30px;border:1px solid var(--line);border-radius:6px;padding:0 8px;width:260px"></div>';
    h += '<div class="tbl-wrap" style="max-height:70vh;overflow:auto"><table class="tbl" id="ltbl"><thead><tr><th>Mənbə vərəq</th><th></th><th>İstifadə edən vərəq</th><th class="n">İstinad</th><th class="n">Mənbə xana</th><th class="n">Düstur</th></tr></thead><tbody>';
    L.forEach(function (e) {
      var s1 = SIDX_N[e.ss], s2 = SIDX_N[e.ds]; if (s1 === undefined || s2 === undefined || !isShown(s1) || !isShown(s2)) return;
      h += '<tr data-f="' + esc(fold(e.ss + ' ' + e.ds + ' ' + sheetTitle(s1) + ' ' + sheetTitle(s2))) + '"><td><span class="dot" style="display:inline-block;background:' + tabColor(s1) + ';margin-right:6px"></span><a href="#/s/' + s1 + '/cedvel">' + esc(sheetTitle(s1)) + '</a> <code class="muted small">' + esc(e.ss) + '</code></td><td class="muted">→</td><td><span class="dot" style="display:inline-block;background:' + tabColor(s2) + ';margin-right:6px"></span><a href="#/s/' + s2 + '/elaqe">' + esc(sheetTitle(s2)) + '</a> <code class="muted small">' + esc(e.ds) + '</code></td><td class="n">' + fmtInt(e.n) + '</td><td class="n">' + fmtInt(e.cells) + '</td><td class="n">' + fmtInt(e.formulas) + '</td></tr>';
    });
    h += '</tbody></table></div></section>';
    main.innerHTML = h;
    $('#lf').addEventListener('input', function () {
      var q = fold(this.value.trim());
      document.querySelectorAll('#ltbl tbody tr').forEach(function (tr) { tr.hidden = q && tr.getAttribute('data-f').indexOf(q) < 0; });
    });
  }

  // ------------------------------------------------------------------ METHOD page
  function pageMethod() {
    var ML = META.methodology || [], counts = {};
    ML.forEach(function (x) { counts[x.status] = (counts[x.status] || 0) + 1; });
    var h = '<div class="eyebrow">Metodologiya</div><div class="page-h"><div><h1>Sənədlə yoxlama: «AZE MOE CAEM — Documentation, v0»</h1><p class="lead">Sənəddəki hər tənlik, identiklik və qayda modeldə onu həyata keçirən xanalarla tutuşdurulub. Xana ünvanına klikləyin — düstur, adlandırılmış əmsallar və əlaqələr həmin vərəqdə açılır.</p></div></div>';
    h += '<div class="row-gap" style="margin-top:14px">' + ['uyğun', 'qismən', 'fərqli', 'tapılmadı'].map(function (s) { return counts[s] ? statusChip(s) + '<b style="margin-right:12px">' + counts[s] + '</b>' : ''; }).join('') + '</div>';
    if (META.method_intro) h += '<div class="prose" style="margin-top:14px">' + META.method_intro + '</div>';
    var secs = [];
    ML.forEach(function (x) { if (secs.indexOf(x.sec) < 0) secs.push(x.sec); });
    secs.forEach(function (sec) {
      h += '<section class="block"><h2>' + esc(sec) + '</h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th style="width:190px">Sənəddə</th><th>Tənlik / qayda</th><th style="width:120px">Status</th><th>Modeldə həyata keçirilməsi</th></tr></thead><tbody>';
      ML.filter(function (x) { return x.sec === sec; }).forEach(function (x) {
        h += '<tr id="' + esc(x.id) + '"><td><b>' + esc(x.id) + '</b> ' + esc(x.title) + (x.desc ? '<div class="small muted" style="margin-top:4px">' + esc(x.desc) + '</div>' : '') + '</td><td><code class="f">' + esc(x.eq) + '</code></td><td>' + statusChip(x.status) + '</td><td>' +
          '<div>' + linkify(x.finding_az || '') + '</div>' +
          (x.cells && x.cells.length ? '<div class="row-gap small" style="margin-top:6px">' + x.cells.slice(0, 10).map(addrLink).join(' ') + '</div>' : '') +
          (x.formula ? '<div class="formula-box" style="margin-top:6px">' + esc(x.formula) + '</div>' : '') +
          (x.coefficients && x.coefficients.length ? '<div class="small muted" style="margin-top:6px">Əmsallar: ' + x.coefficients.map(function (k) { return esc(k.name) + ' = ' + esc(k.value) + (k.cell ? ' (' + linkify(k.cell) + ')' : ''); }).join('; ') + '</div>' : '') +
          '</td></tr>';
      });
      h += '</tbody></table></div></section>';
    });
    var IS = META.issues || [];
    if (IS.length) {
      h += '<section class="block" id="issues"><h2>Modeldə aşkar edilmiş məsələlər <small>' + IS.length + '</small></h2><div class="filters"><label for="isb">Bölmə</label><select id="isb" style="height:30px;border:1px solid var(--line);border-radius:6px"><option value="">hamısı</option>' +
        GROUPS.map(function (g, gi) { return '<option value="' + gi + '">' + esc(g.title) + '</option>'; }).join('') +
        '</select><label for="iss">Dərəcə</label><select id="iss" style="height:30px;border:1px solid var(--line);border-radius:6px"><option value="">hamısı</option><option>yüksək</option><option>orta</option><option>aşağı</option></select></div><div id="isl"></div></section>';
    }
    main.innerHTML = h;
    if (IS.length) {
      var drawIs = function () {
        var b = $('#isb').value, sv = $('#iss').value;
        $('#isl').innerHTML = issuesHTML(IS.filter(function (x) { return (b === '' || issueGroups(x).indexOf(+b) >= 0) && (!sv || x.severity === sv); }));
      };
      $('#isb').onchange = drawIs; $('#iss').onchange = drawIs; drawIs();
      if (/#issues$/.test(location.hash) || ROUTE.a === 'issues') setTimeout(function () { var el = document.getElementById('issues'); if (el) el.scrollIntoView(); }, 30);
    }
  }
  function issueSheets(x) {
    var out = [], re = sheetRe(); re.lastIndex = 0; var t = esc(String(x.where || '')), m;
    while ((m = re.exec(t))) { var gs = sheetByLooseName(m[1]); if (gs >= 0 && out.indexOf(gs) < 0) out.push(gs); }
    return out;
  }
  function issueGroups(x) { return issueSheets(x).map(function (gs) { return GROUP_OF[gs]; }); }
  function issuesFor(gsList) { return (META.issues || []).filter(function (x) { return issueSheets(x).some(function (gs) { return gsList.indexOf(gs) >= 0; }); }); }

  // ------------------------------------------------------------------ GROUP page (section of sheets)
  function pageGroup(gi, tab) {
    var g = GROUPS[gi]; if (!g) { pageOverview(); return; }
    var base = '#/g/' + gi + '/';
    var h = '<div class="eyebrow">Bölmə ' + (gi + 1) + ' / ' + GROUPS.length + '</div><div class="page-h"><div><h1><span class="dot" style="display:inline-block;width:14px;height:14px;margin-right:8px;background:' + g.color + '"></span>' + esc(g.title) + '</h1><div class="sub">' + g.sheets.length + ' vərəq</div></div></div>';
    h += tabsHTML(base, [['tesvir', 'Təsvir'], ['giris', 'Giriş məlumatları (Admin)'], ['emsal', 'Əmsallar'], ['elaqe', 'Əlaqələr']], tab);
    h += '<div id="tabbody"></div>';
    main.innerHTML = h;
    var tb = $('#tabbody');
    if (tab === 'tesvir') {
      var t = '';
      if (g.desc) t += '<section class="block"><p class="lead">' + linkify(g.desc) + '</p></section>';
      t += '<section class="block"><h2>Vərəqlər (alt səhifələr)</h2><div class="cards">' + g.sheets.map(function (gs) {
        var sm = sheetMeta(gs) || {}, st = sheetStats(gs);
        return '<a class="card" href="#/s/' + gs + '/cedvel"><div class="meta"><span class="dot" style="background:' + tabColor(gs) + '"></span><code>' + esc(SHEETS[gs].n) + '</code>' + (SHEETS[gs].st !== 'visible' ? '<span class="chip">gizli</span>' : '') + '</div><h3>' + esc(sheetTitle(gs)) + '</h3>' +
          (sm.desc_az ? '<p>' + esc(sm.desc_az.length > 420 ? sm.desc_az.slice(0, 420).replace(/\s\S*$/, '') + ' …' : sm.desc_az) + '</p>' : '') + '<div class="meta">' + fmtInt(st.formulas) + ' düstur · ' + fmtInt(st.consts) + ' sabit' + (sm.doc_ref ? ' · ' + esc(sm.doc_ref) : '') + '</div></a>';
      }).join('') + '</div></section>';
      var bis = issuesFor(g.sheets);
      if (bis.length) t += '<section class="block"><h2>Bu bölmədə aşkar edilmiş məsələlər <small>' + bis.length + '</small></h2>' + issuesHTML(bis) + '</section>';
      tb.innerHTML = t;
    } else if (tab === 'giris') {
      tb.innerHTML = adminIntro() + '<div id="inputs"></div>';
      renderInputsForSheets(g.sheets, $('#inputs'), true);
    } else if (tab === 'emsal') {
      tb.innerHTML = coefIntro() + '<div id="coefs"></div>';
      renderCoefsForSheets(g.sheets, $('#coefs'), true);
    } else if (tab === 'elaqe') {
      tb.innerHTML = linksForSheets(g.sheets, true);
    }
  }

  // ------------------------------------------------------------------ SHEET page
  function pageSheet(gs, tab, cellA1) {
    var sm = sheetMeta(gs) || {}, gi = GROUP_OF[gs], g = GROUPS[gi] || { title: '', color: '#98A2AE' };
    var base = '#/s/' + gs + '/';
    var st = sheetStats(gs), nIs = issuesFor([gs]).length;
    var h = '<div class="sheet-head"><div class="eyebrow"><a href="#/g/' + gi + '">' + esc(g.title) + '</a> › vərəq</div><div class="page-h"><div><h1><span class="tabdot big" style="background:' + tabColor(gs) + '"></span>' + esc(sheetTitle(gs)) + '</h1><div class="sub"><span class="orig">' + esc(SHEETS[gs].n) + '</span>' +
      ' <span class="muted small">' + fmtInt(st.formulas) + ' düstur · ' + fmtInt(st.consts) + ' sabit</span>' + (nIs ? ' <a class="chip chip-warn" href="#/s/' + gs + '/haqqinda">' + nIs + ' məsələ</a>' : '') + '</div>' + (sm.desc_az && tab !== 'cedvel' ? '<p class="lead" style="margin-top:8px">' + linkify(sm.desc_az) + '</p>' : '') + '</div></div>';
    h += tabsHTML(base, [['cedvel', 'Cədvəl'], ['giris', 'Giriş məlumatları (Admin)'], ['emsal', 'Əmsallar'], ['elaqe', 'Əlaqələr'], ['haqqinda', 'Təsvir']], tab) + '</div>';
    if (tab === 'cedvel') {
      main.classList.add('grid-mode');
      h += '<div class="grid-toolbar"><div class="legend"><span><i style="background:var(--input)"></i>giriş (sabit)</span><span><i style="background:var(--ink)"></i>düstur</span><span><i style="background:var(--link)"></i>massiv düsturu</span><span><i style="background:var(--coef)"></i>əmsal</span><span><i style="background:var(--edit);box-shadow:inset 0 0 0 1px #E8C95A"></i>dəyişdirilib</span><span><i style="background:var(--chg)"></i>nəticə dəyişib</span></div>' +
        macroButtons(gs) +
        '<label class="small" style="margin-left:auto;display:inline-flex;gap:6px;align-items:center"><input type="checkbox" id="showhidden"> gizli sətir/sütunlar</label>' +
        '<button class="btn btn-sm" id="csv">CSV ixrac</button></div><div class="gridwrap" id="gridwrap" tabindex="0" aria-label="Cədvəl"></div><div class="grid-status" id="gstatus"></div>';
      main.innerHTML = h;
      if (!SC[gs].length) { $('#gridwrap').innerHTML = '<div class="note" style="margin:16px">Bu vərəqdə xana yoxdur (Excel-də yalnız şəkil, diaqram və ya naviqasiya düymələri var). ' + (sm.desc_az ? esc(sm.desc_az) : '') + '</div>'; return; }
      GRID = new Grid($('#gridwrap'), gs);
      $('#showhidden').checked = !!GRID.showHidden;
      $('#showhidden').addEventListener('change', function () { GRID.showHidden = this.checked; GRID.layout(); GRID.render(true); });
      $('#csv').addEventListener('click', function () { exportCSV(gs); });
      wireMacroButtons();
      if (cellA1) { var rc = parseA1(cellA1); if (rc) { GRID.select(rc.r, rc.c, true); } }
      else if (SEL && SEL.gs === gs) GRID.select(SEL.r, SEL.c, true);
      else { closeInspector(); updateStatus(); }
      return;
    }
    h += '<div id="tabbody"></div>';
    main.innerHTML = h;
    var tb = $('#tabbody');
    if (tab === 'giris') { tb.innerHTML = adminIntro(); var d = document.createElement('div'); tb.appendChild(d); renderInputsForSheets([gs], d, false); }
    else if (tab === 'emsal') { tb.innerHTML = coefIntro(); var d2 = document.createElement('div'); tb.appendChild(d2); renderCoefsForSheets([gs], d2, false); }
    else if (tab === 'elaqe') tb.innerHTML = linksForSheets([gs], false);
    else if (tab === 'haqqinda') {
      var t = '<section class="block">' + (sm.desc_az ? (sm.doc_ref ? '<p class="small muted">Sənəddə: ' + esc(sm.doc_ref) + '</p>' : '') : '<p class="muted">Təsvir yoxdur.</p>') + '</section>';
      var nm = namesOnSheet(gs);
      if (nm.length) t += '<section class="block"><h2>Bu vərəqdəki adlandırılmış xanalar <small>' + nm.length + '</small></h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th>Ad</th><th>Xana</th><th>Göstərici</th><th class="n">Dəyər</th></tr></thead><tbody>' + nm.map(function (x) {
        return '<tr><td><code class="f">' + esc(x.n) + '</code></td><td><a href="#/s/' + gs + '/cedvel/' + a1(M.r[x.id], M.c[x.id]) + '">' + a1(M.r[x.id], M.c[x.id]) + '</a></td><td class="small">' + esc(rowLabel(gs, M.r[x.id])) + '</td><td class="n">' + esc(dispVal(M.V[x.id], STY[M.st[x.id]])) + '</td></tr>';
      }).join('') + '</tbody></table></div></section>';
      var sis = issuesFor([gs]);
      if (sis.length) t += '<section class="block"><h2>Bu vərəqdə aşkar edilmiş məsələlər <small>' + sis.length + '</small></h2>' + issuesHTML(sis) + '</section>';
      tb.innerHTML = t;
    }
  }
  function namesOnSheet(gs) {
    var out = []; Object.keys(NAMES_OF).forEach(function (k) { var x = +k; if (M.sh[x] === gs) NAMES_OF[x].forEach(function (n) { out.push({ n: n, id: x }); }); });
    return out.sort(function (a, b) { return M.r[a.id] - M.r[b.id] || M.c[a.id] - M.c[b.id]; });
  }
  function adminIntro() {
    return '<div class="note ' + (ADMIN ? 'note-coef' : 'note-warn') + '" style="margin-top:16px">' + (ADMIN
      ? '<b>Admin rejimi aktivdir.</b> Dəyərləri dəyişin və Enter basın — bütün asılı düsturlar dərhal yenidən hesablanır. Mətn açarları (məs. «Model determined» / manual, məzənnə rejimi) siyahıdan seçilir. Dəyişikliklər bu brauzerdə saxlanılır; <a href="#/admin">Admin panelində</a> ixrac etmək, ssenari kimi saxlamaq və ya sıfırlamaq olar.'
      : '<b>Baxış rejimi.</b> Giriş məlumatlarını dəyişmək üçün yuxarıdakı <b>Admin rejimi</b> açarını aktivləşdirin.') +
      ' Burada ən azı bir düstura təsir edən sabit xanalar göstərilir (sarı/qırmızı işarəli əl ilə daxil edilən xanalar ' + INFO.firstYear + '–' + INFO.endYear + ' üçün); düstur xanaları boz rəngdə, yalnız oxumaq üçündür.</div>';
  }
  function coefIntro() {
    return '<div class="note note-coef" style="margin-top:16px"><b>Əmsallar iki mənbədən gəlir.</b> (1) Adlandırılmış xanalarda saxlanan davranış tənliklərinin əmsalları (məs. <code>CPI_coeff_1</code>, <code>Cons_coeff_1</code>, <code>C_lambda</code>) — əsasən «Parametrization» vərəqində; düsturlar onlara adla istinad edir. (2) Düsturların içinə birbaşa yazılmış ədədi sabitlər. Eyni sətirdə eyni düstur quruluşunda təkrarlanan sabit bir qrup kimi göstərilir — qrupu dəyişmək bütün illərə tətbiq olunur.' + (ADMIN ? '' : ' <b>Dəyişmək üçün Admin rejimini aktivləşdirin.</b>') + '</div>';
  }

  // ------------------------------------------------------------------ input editor
  var FORECAST_ONLY = true;
  function renderInputsForSheets(list, host, grouped) {
    var D = M.dependents();
    var h = '<div class="filters"><label for="fc"><input type="checkbox" id="fc"' + (FORECAST_ONLY ? ' checked' : '') + '> yalnız ' + INFO.lastActual + '–' + INFO.endYear + ' illəri (ilsiz sütunlar həmişə göstərilir)</label><span class="muted small" id="inpcount"></span></div>';
    host.innerHTML = h + '<div id="inphost"></div>';
    var inner = $('#inphost', host);
    function draw() {
      var total = 0, out = '';
      list.forEach(function (gs) {
        var rows = {}, cols = {};
        SC[gs].forEach(function (id) {
          if (!D[id] || M.ft[id] >= 0) return;
          if (!(isConstNum(id) || SWITCH[id] || isConstText(id))) return;
          if (isConstText(id) && !SWITCH[id] && !(STY[M.st[id]] && (STY[M.st[id]].bg === 'FFFF00' || STY[M.st[id]].fc === 'FF0000'))) return;
          var y = colYear(gs, M.r[id], M.c[id]);
          if (FORECAST_ONLY && y !== null && y < INFO.lastActual) return;
          (rows[M.r[id]] || (rows[M.r[id]] = [])).push(id); cols[M.c[id]] = 1;
        });
        var rks = Object.keys(rows).map(Number).sort(function (a, b) { return a - b; });
        if (!rks.length) { if (!grouped) out += '<p class="muted">Bu vərəqdə düsturlara təsir edən sabit giriş xanası yoxdur' + (FORECAST_ONLY ? ' (seçilmiş illər üzrə)' : '') + '.</p>'; return; }
        var cks = Object.keys(cols).map(Number).sort(function (a, b) { return a - b; });
        var n = 0; rks.forEach(function (r) { n += rows[r].length; }); total += n;
        var tbl = '<div class="ed-wrap"><table class="ed-tbl"><thead><tr><th class="lab">Göstərici</th>' + cks.map(function (c) {
          var y = colYear(gs, rks[0], c); return '<th class="n" title="Sütun ' + colName(c) + '">' + (y ? y : colName(c)) + '</th>';
        }).join('') + '</tr></thead><tbody>';
        rks.forEach(function (r) {
          tbl += '<tr><td class="lab"><a class="rn" href="#/s/' + gs + '/cedvel/' + a1(r, rows[r][0] ? M.c[rows[r][0]] : 1) + '">' + r + '</a>' + esc(rowLabel(gs, r) || '(etiketsiz)') + '</td>';
          cks.forEach(function (c) {
            var id = cellId(gs, r, c);
            if (id < 0) { tbl += '<td></td>'; return; }
            if (M.ft[id] < 0 && D[id] && (isConstNum(id) || SWITCH[id] || isConstText(id))) tbl += '<td>' + inputControl(id) + '</td>';
            else if (M.ft[id] >= 0) {
              var ch = !sameVal(M.V[id], M.V0[id]);
              tbl += '<td class="fv' + (ch ? ' chg' : '') + '" data-fid="' + id + '" title="' + esc(fullAddr(id) + ' · düstur') + '">' + esc(dispVal(M.V[id], STY[M.st[id]])) + '</td>';
            } else tbl += '<td class="fv">' + esc(dispVal(M.V[id], STY[M.st[id]])) + '</td>';
          });
          tbl += '</tr>';
        });
        tbl += '</tbody></table></div>';
        if (grouped) out += '<details class="grp"' + (list.length <= 3 ? ' open' : '') + '><summary><span class="tabdot" style="background:' + tabColor(gs) + '"></span>' + esc(sheetTitle(gs)) + ' <code class="muted small">' + esc(SHEETS[gs].n) + '</code><span class="chip chip-info">' + n + ' giriş</span><a class="small" style="margin-left:auto;font-weight:400" href="#/s/' + gs + '/giris">vərəqdə aç</a></summary><div class="grp-body">' + tbl + '</div></details>';
        else out += tbl;
      });
      inner.innerHTML = out || '<p class="muted">Giriş xanası tapılmadı.</p>';
      $('#inpcount', host).textContent = fmtInt(total) + ' giriş xanası';
    }
    draw();
    $('#fc', host).addEventListener('change', function () { FORECAST_ONLY = this.checked; draw(); });
    wireValueInputs(host, function (inp, v) { setInput(+inp.getAttribute('data-id'), v); });
  }
  // numeric field, option list (switch) or free text; locked cells are read-only
  function inputControl(id) {
    var mod = EDITS.inputs.has(id), dis = ADMIN && !LOCKED[id] ? '' : ' disabled', t = esc(fullAddr(id) + (mod ? ' · ' + INFO.baseShort + ': ' + fullVal(M.V0[id]) : '') + (LOCKED[id] ? ' · ' + LOCKED[id] : ''));
    if (SWITCH[id]) {
      var opts = SWITCH[id].options.slice(); if (opts.indexOf(M.V[id]) < 0) opts.unshift(M.V[id]);
      return '<select class="v sw' + (mod ? ' mod' : '') + '" data-id="' + id + '" title="' + t + '"' + dis + '>' + opts.map(function (o) { return '<option' + (o === M.V[id] ? ' selected' : '') + '>' + esc(o) + '</option>'; }).join('') + '</select>';
    }
    if (typeof M.V0[id] === 'string') return '<input class="v txt' + (mod ? ' mod' : '') + '" data-id="' + id + '" data-kind="text" value="' + esc(M.V[id] == null ? '' : M.V[id]) + '" title="' + t + '"' + dis + '>';
    return '<input class="v' + (mod ? ' mod' : '') + (LOCKED[id] ? ' locked' : '') + '" data-id="' + id + '" value="' + esc(fmtEditVal(M.V[id])) + '" title="' + t + '"' + dis + ' inputmode="decimal">';
  }
  function fmtEditVal(v) { if (typeof v === 'number') return String(Number(v.toPrecision(15))); return v === null || v === undefined ? '' : String(v); }
  function wireValueInputs(host, apply) {
    host.addEventListener('keydown', function (e) {
      var inp = e.target; if (!inp.classList || !inp.classList.contains('v') || inp.tagName === 'SELECT') return;
      if (e.key === 'Enter') { e.preventDefault(); commitInput(inp, apply); }
      if (e.key === 'Escape') { inp.value = inp.getAttribute('data-prev') || inp.value; inp.blur(); }
    });
    host.addEventListener('focusin', function (e) { if (e.target.classList && e.target.classList.contains('v')) e.target.setAttribute('data-prev', e.target.value); });
    host.addEventListener('focusout', function (e) {
      var inp = e.target; if (!inp.classList || !inp.classList.contains('v') || inp.tagName === 'SELECT') return;
      if (inp.value !== inp.getAttribute('data-prev')) commitInput(inp, apply);
    });
    host.addEventListener('change', function (e) {
      var s = e.target; if (s.tagName !== 'SELECT' || !s.classList.contains('v')) return;
      if (!ADMIN) { toast('Dəyişmək üçün Admin rejimini aktivləşdirin'); return; }
      apply(s, s.value);
    });
  }
  function commitInput(inp, apply) {
    if (!ADMIN) { toast('Dəyişmək üçün Admin rejimini aktivləşdirin'); return; }
    if (inp.getAttribute('data-kind') === 'text') { inp.setAttribute('data-prev', inp.value); apply(inp, inp.value); return; }
    var v = parseUserNumber(inp.value);
    if (v === undefined) { toast('Rəqəm daxil edin (məs. 12.5 və ya 12,5)'); inp.value = inp.getAttribute('data-prev'); return; }
    if (v === null) { toast('Boş dəyər qəbul edilmir — sıfırlamaq üçün Admin panelindən istifadə edin'); inp.value = inp.getAttribute('data-prev'); return; }
    inp.setAttribute('data-prev', inp.value);
    apply(inp, v);
  }
  function refreshEditorValues() {
    document.querySelectorAll('.v[data-id]').forEach(function (inp) {
      var id = +inp.getAttribute('data-id'); if (document.activeElement === inp) return;
      if (inp.tagName === 'SELECT') inp.value = M.V[id]; else inp.value = typeof M.V[id] === 'number' ? fmtEditVal(M.V[id]) : (M.V[id] == null ? '' : M.V[id]);
      inp.classList.toggle('mod', EDITS.inputs.has(id));
    });
    document.querySelectorAll('input.v[data-k]').forEach(function (inp) {
      if (document.activeElement === inp) return;
      var ks = inp.getAttribute('data-k').split(',').map(Number), v = M.K[ks[0]], mixed = ks.some(function (k) { return !sameVal(M.K[k], v); });
      inp.value = mixed ? '' : String(v); inp.placeholder = mixed ? 'müxtəlif' : '';
      inp.classList.toggle('mod', ks.some(function (k) { return EDITS.K.has(k); }));
    });
    document.querySelectorAll('td.fv[data-fid]').forEach(function (td) {
      var id = +td.getAttribute('data-fid'); td.textContent = dispVal(M.V[id], STY[M.st[id]]); td.classList.toggle('chg', !sameVal(M.V[id], M.V0[id]));
    });
  }

  // ------------------------------------------------------------------ coefficient editor
  var COEF_ALL = false;
  function coefGroups(gs) {
    var groups = [], open = {};
    SC[gs].forEach(function (id) {
      var t = M.ft[id]; if (t < 0) return;
      var n = TPL_NL[t]; if (!n) return;
      for (var li = 0; li < n; li++) {
        var k = M.fk[id] + li, key = M.r[id] + '|' + t + '|' + li, g = open[key];
        if (g && g.lastC === M.c[id] - 1 && sameVal(M.K0[g.ks[0]], M.K0[k])) { g.ids.push(id); g.ks.push(k); g.lastC = M.c[id]; }
        else { g = open[key] = { r: M.r[id], t: t, li: li, ids: [id], ks: [k], lastC: M.c[id] }; groups.push(g); }
      }
    });
    return groups;
  }
  // named coefficient cells used by the formulas of these sheets (and the ones that live on them)
  function coefRefs(gs) {
    var set = new Set(), C = coefSet();
    SC[gs].forEach(function (id) {
      if (C.has(id)) set.add(id);
      if (M.ft[id] < 0) return;
      M.precedents(id).forEach(function (p) { if (C.has(p)) set.add(p); });
    });
    return Array.from(set).sort(function (a, b) { return M.sh[a] - M.sh[b] || M.r[a] - M.r[b] || M.c[a] - M.c[b]; });
  }
  function isTrivial(v) { return typeof v === 'number' && TRIVIAL[Math.abs(v)] === 1; }
  function renderCoefsForSheets(list, host, grouped) {
    host.innerHTML = '<div class="filters"><label for="ca"><input type="checkbox" id="ca"' + (COEF_ALL ? ' checked' : '') + '> texniki sabitləri də göstər (0, 1, 2, 3, 4, 10, 12, 100, 1000 …)</label><span class="muted small" id="coefcount"></span></div><div id="coefhost"></div>';
    var inner = $('#coefhost', host);
    function draw() {
      var out = '', total = 0, D = M.dependents();
      list.forEach(function (gs) {
        var eqs = coefRefs(gs), grs = coefGroups(gs).filter(function (g) { return COEF_ALL || !isTrivial(M.K0[g.ks[0]]); });
        if (!eqs.length && !grs.length) { if (!grouped) out += '<p class="muted">Bu vərəqdə əmsal yoxdur.</p>'; return; }
        total += eqs.length + grs.length;
        var t = '';
        if (eqs.length) {
          t += '<h3 style="font-size:14px;margin:14px 0 6px">Adlandırılmış əmsallar (davranış tənlikləri, parametrlər) <span class="chip chip-coef">' + eqs.length + '</span></h3>';
          t += '<div class="ed-wrap"><table class="ed-tbl coef-tbl"><thead><tr><th class="lab">Ad / göstərici</th><th>Xana</th><th class="n">Dəyər</th><th class="n">' + INFO.baseShort + '</th><th class="n">İstifadə</th></tr></thead><tbody>';
          eqs.forEach(function (id) {
            var mod = EDITS.inputs.has(id), nm = NAMES_OF[id] ? NAMES_OF[id].join(', ') : '';
            t += '<tr><td class="lab">' + (nm ? '<code class="f">' + esc(nm) + '</code> ' : '') + '<span class="muted">' + esc(rowLabel(M.sh[id], M.r[id]) || '') + '</span></td><td><a href="#/s/' + M.sh[id] + '/cedvel/' + a1(M.r[id], M.c[id]) + '"><code class="f">' + esc(fullAddr(id)) + '</code></a></td>' +
              '<td><input class="v' + (mod ? ' mod' : '') + '" data-id="' + id + '" value="' + esc(fmtEditVal(M.V[id])) + '"' + (ADMIN ? '' : ' disabled') + ' inputmode="decimal"></td><td class="fv">' + esc(fmtEditVal(M.V0[id])) + '</td><td class="fv">' + (D[id] ? D[id].length : 0) + ' düstur</td></tr>';
          });
          t += '</tbody></table></div>';
        }
        if (grs.length) {
          t += '<h3 style="font-size:14px;margin:18px 0 6px">Düsturların içindəki ədədi sabitlər <span class="chip chip-coef">' + grs.length + ' qrup</span></h3>';
          t += '<div class="ed-wrap"><table class="ed-tbl coef-tbl"><thead><tr><th class="lab">Göstərici (sətir)</th><th>Xanalar</th><th>Düstur (vurğulanan sabit)</th><th class="n">Dəyər</th><th class="n">' + INFO.baseShort + '</th></tr></thead><tbody>';
          grs.forEach(function (g) {
            var first = g.ids[0], last = g.ids[g.ids.length - 1], v = M.K[g.ks[0]], mixed = g.ks.some(function (k) { return !sameVal(M.K[k], v); });
            var mod = g.ks.some(function (k) { return EDITS.K.has(k); });
            var yr1 = colYear(gs, g.r, M.c[first]), yr2 = colYear(gs, g.r, M.c[last]);
            t += '<tr><td class="lab"><a class="rn" href="#/s/' + gs + '/cedvel/' + a1(g.r, M.c[first]) + '">' + g.r + '</a>' + esc(rowLabel(gs, g.r) || '(etiketsiz)') + '</td>' +
              '<td><code class="f">' + a1(g.r, M.c[first]) + (g.ids.length > 1 ? ':' + a1(g.r, M.c[last]) : '') + '</code><div class="small muted">' + (yr1 ? yr1 + (yr2 && yr2 !== yr1 ? '–' + yr2 : '') : '') + ' · ' + g.ids.length + ' xana</div></td>' +
              '<td class="snip">' + formulaHTML(first, { hlLit: g.li }) + '</td>' +
              '<td><input class="v' + (mod ? ' mod' : '') + '" data-k="' + g.ks.join(',') + '" value="' + (mixed ? '' : esc(String(v))) + '" placeholder="' + (mixed ? 'müxtəlif' : '') + '"' + (ADMIN ? '' : ' disabled') + ' inputmode="decimal"></td><td class="fv">' + esc(String(M.K0[g.ks[0]])) + '</td></tr>';
          });
          t += '</tbody></table></div>';
        }
        if (grouped) out += '<details class="grp"' + (list.length <= 3 ? ' open' : '') + '><summary><span class="tabdot" style="background:' + tabColor(gs) + '"></span>' + esc(sheetTitle(gs)) + ' <code class="muted small">' + esc(SHEETS[gs].n) + '</code><span class="chip chip-coef">' + (eqs.length + grs.length) + '</span><a class="small" style="margin-left:auto;font-weight:400" href="#/s/' + gs + '/emsal">vərəqdə aç</a></summary><div class="grp-body">' + t + '</div></details>';
        else out += t;
      });
      inner.innerHTML = out || '<p class="muted">Əmsal tapılmadı.</p>';
      $('#coefcount', host).textContent = fmtInt(total) + ' əmsal / qrup';
    }
    draw();
    $('#ca', host).addEventListener('change', function () { COEF_ALL = this.checked; draw(); });
    wireValueInputs(host, function (inp, v) {
      if (inp.hasAttribute('data-k')) setCoef(inp.getAttribute('data-k').split(',').map(Number), v);
      else setInput(+inp.getAttribute('data-id'), v);
    });
  }

  // ------------------------------------------------------------------ links for sheets
  function linksForSheets(list, grouped) {
    var D = M.dependents(), inc = {}, out = {}, set = {};
    list.forEach(function (gs) { set[gs] = 1; });
    list.forEach(function (gs) {
      SC[gs].forEach(function (id) {
        var srcId = isMember(id) ? groupOfMember(id) : id;
        if (M.ft[id] >= 0) M.precedents(srcId).forEach(function (p) { p = realCell(p); var s = M.sh[p]; if (!set[s] && isShown(s)) { inc[s] = inc[s] || { n: 0, cells: new Set() }; inc[s].n++; inc[s].cells.add(p); } });
        var ds = D[id]; if (ds) ds.forEach(function (d) { d = realCell(d); var s = M.sh[d]; if (!set[s] && isShown(s)) { out[s] = out[s] || { n: 0, cells: new Set() }; out[s].n++; out[s].cells.add(id); } });
      });
    });
    function tbl(obj, title, dir) {
      var ks = Object.keys(obj).map(Number).sort(function (a, b) { return obj[b].n - obj[a].n; });
      var h = '<section class="block"><h2>' + title + ' <small>' + ks.length + ' vərəq</small></h2>';
      if (!ks.length) return h + '<p class="muted">Yoxdur.</p></section>';
      h += '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Bölmə</th><th>Vərəq</th><th class="n">İstinad</th><th class="n">' + (dir === 'in' ? 'Mənbə xana' : 'Bu vərəqdən oxunan xana') + '</th></tr></thead><tbody>';
      ks.forEach(function (s) {
        var g = GROUPS[GROUP_OF[s]];
        h += '<tr><td><span class="dot" style="display:inline-block;background:' + (g ? g.color : '#98A2AE') + ';margin-right:6px"></span>' + esc(g ? g.title : '') + '</td><td><a href="#/s/' + s + '/cedvel">' + esc(sheetTitle(s)) + '</a> <code class="muted small">' + esc(SHEETS[s].n) + '</code></td><td class="n">' + fmtInt(obj[s].n) + '</td><td class="n">' + fmtInt(obj[s].cells.size) + '</td></tr>';
      });
      return h + '</tbody></table></div></section>';
    }
    return tbl(inc, grouped ? 'Bu bölmə məlumatı haradan alır' : 'Bu vərəq məlumatı haradan alır', 'in') + tbl(out, grouped ? 'Bu bölmənin nəticələrindən kim istifadə edir' : 'Bu vərəqin nəticələrindən kim istifadə edir', 'out');
  }

  // ------------------------------------------------------------------ macros (VBA ported): reconciliation and calibration
  var MACRO_SHEETS = { '1d. Real GDP - Expenditure': 1, '1a. Real GDP - Production': 1, '1b. Infl and Ex rates': 1, '1c. Interest rates': 1 };
  function macroButtons(gs) {
    if (!MACRO_SHEETS[SHEETS[gs].n]) return '';
    return '<span class="row-gap" style="margin-left:12px"><button class="btn btn-sm" data-macro="YP" title="ResidualYP: təklif tərəfi ÜDM artımını dəyişməklə təklif–tələb və deflyator fərqlərini bağlayır">Uzlaşdır (YP)</button><button class="btn btn-sm" data-macro="CP" title="ResidualCP: tələb tərəfində düzəliş (WedgeSD) ilə fərqləri bağlayır">Uzlaşdır (CP)</button><button class="btn btn-sm" data-macro="CALI" title="CaliManual: 6 davranış tənliyinin əlavə amillərini son proqnozlara uyğunlaşdırır">Kalibrlə</button></span>';
  }
  function wireMacroButtons() { document.querySelectorAll('[data-macro]').forEach(function (b) { b.onclick = function () { runMacro(b.getAttribute('data-macro')); }; }); }
  function wedges() {
    try {
      var C = window.CaemMacros, g = C.nameCell(M, 'GDP_diff'), d = C.nameCell(M, 'deflator_diff'), out = [];
      for (var j = 0; j <= 4; j++) out.push([M.V[C.at(M, g, j)], M.V[C.at(M, d, j)]]);
      return out;
    } catch (e) { return null; }
  }
  function runMacro(kind) {
    if (!ADMIN) { toast('Makroslar giriş dəyərlərini dəyişir — əvvəlcə Admin rejimini aktivləşdirin'); return; }
    var C = window.CaemMacros, t0 = performance.now(), res;
    var set = function (id, v) { M.V[id] = v; M.recalcFrom([id], EDITS.ov); };
    try { res = kind === 'CALI' ? C.calibrate(M, set) : C.reconcile(M, set, kind); }
    catch (e) { toast('Makro dayandı: ' + e.message); return; }
    res.changed.forEach(function (v, id) { if (sameVal(v, M.V0[id])) EDITS.inputs.delete(id); else EDITS.inputs.set(id, v); });
    persist();
    var ms = Math.round(performance.now() - t0), msg;
    if (kind === 'CALI') msg = 'Kalibrləmə tamamlandı: ' + res.changed.size + ' əlavə amil yeniləndi (' + ms + ' ms).';
    else if (res.status === 'done') msg = 'Hazırdır! 5 illik təklif–tələb və deflyator fərqləri bağlandı. Maksimal fərq: ' + res.maxWedge.toExponential(2) + ' (tolerantlıq ' + res.tol + ').';
    else if (res.status === 'notclosed') msg = 'Xəbərdarlıq: t+' + res.year + ' ilinin fərqləri 20 iterasiyadan sonra bağlanmadı (maks. ' + res.maxWedge.toExponential(2) + '). Düyməni yenidən basın və ya ilkin təxminləri dəyişin.';
    else msg = 'Xəta: t+' + res.year + ' ilinin fərqlərini bağlamaq əvvəlki illərin fərqlərini dəyişir — sistemdə irəliyə baxan dəyişən var.';
    alert(msg);
    afterChange(res.recalcs);
    if (ROUTE.name === 'admin') pageAdmin();
  }

  // ------------------------------------------------------------------ ADMIN page
  function pageAdmin() {
    var h = '<div class="eyebrow">İdarəetmə</div><div class="page-h"><div><h1>Admin paneli</h1><p class="lead">Giriş məlumatları, mətn açarları, əmsallar və düstur üzərinə yazılmış dəyərlər burada idarə olunur. Dəyişikliklər avtomatik olaraq bu brauzerdə saxlanılır və bütün səhifələrdə (iş paneli daxil) dərhal əks olunur.</p></div></div>';
    h += '<div class="note ' + (ADMIN ? 'note-coef' : 'note-warn') + '" style="margin-top:14px">' + (ADMIN ? '<b>Admin rejimi aktivdir.</b> Hər səhifənin <b>Giriş məlumatları</b> və <b>Əmsallar</b> tablarında, həmçinin cədvəldə (xananı iki dəfə klikləyin) dəyərləri dəyişə bilərsiniz.' : '<b>Admin rejimi söndürülüb.</b> Dəyişiklik etmək üçün yuxarıdakı açarı aktivləşdirin.') + ' <button class="btn btn-sm" id="adm-toggle" style="margin-left:8px">' + (ADMIN ? 'Söndür' : 'Aktivləşdir') + '</button></div>';
    var w = wedges();
    h += '<section class="block"><h2>Makroslar <small>Excel-dəki VBA düymələri (Əlavə 6) — JavaScript-də eyni alqoritmlə</small></h2>';
    if (w) h += '<div class="tbl-wrap" style="max-width:720px"><table class="tbl"><thead><tr><th>Fərq</th>' + w.map(function (x, j) { return '<th class="n">' + (INFO.firstYear + j) + '</th>'; }).join('') + '</tr></thead><tbody>' +
      '<tr><td>Təklif–tələb fərqi (GDP_diff)</td>' + w.map(function (x) { return '<td class="n">' + esc(shortVal(x[0])) + '</td>'; }).join('') + '</tr>' +
      '<tr><td>ÜDM deflyatoru fərqi (deflator_diff)</td>' + w.map(function (x) { return '<td class="n">' + esc(shortVal(x[1])) + '</td>'; }).join('') + '</tr></tbody></table></div>';
    h += '<div class="row-gap" style="margin-top:10px"><button class="btn" data-macro="YP">Uzlaşdır: təklif tərəfi (ResidualYP)</button><button class="btn" data-macro="CP">Uzlaşdır: tələb tərəfi düzəlişi (ResidualCP)</button><button class="btn" data-macro="CALI">Əlavə amilləri kalibrlə (CaliManual)</button></div>' +
      '<p class="small muted">Uzlaşdırma hər proqnoz ili üçün ardıcıl olaraq iki tənliyi (ÜDM və deflyator fərqi = 0) «Goal Seek» ilə həll edir (tolerantlıq 0,001, cəmi 20 dövrə). Kalibrləmə C, I, M, X üçün 100·(ln Final − ln Model), π və faiz üçün (Final − Model) qədər əlavə amili dəyişir. Nəticələr giriş dəyişikliyi kimi qeyd olunur və geri qaytarıla bilər.</p></section>';
    h += '<section class="block"><h2>Əməliyyatlar</h2><div class="row-gap">' +
      '<button class="btn" id="adm-export">Dəyişiklikləri JSON kimi ixrac et</button>' +
      '<label class="btn" for="adm-import" style="display:inline-block">JSON idxal et<input type="file" id="adm-import" accept=".json,application/json" hidden></label>' +
      '<button class="btn" id="adm-full">Bütün modeli yenidən hesabla</button>' +
      '<button class="btn btn-danger" id="adm-reset">Bütün dəyişiklikləri sıfırla</button></div>' +
      '<p class="small muted">«Bütün modeli yenidən hesabla» ' + fmtInt(M.order.length) + ' düsturun hamısını giriş məlumatlarından başlayaraq yenidən hesablayır (adi rejimdə yalnız dəyişikliyin təsir etdiyi düsturlar hesablanır).</p></section>';
    h += '<div class="note" style="margin-top:14px"><b>Faktiki məlumat, proqnozun ilk ilini irəli çəkmək və ya yeni Excel faylı?</b> <a href="#/yenile">Model yeniləmə</a> səhifəsinə keçin.</div>';
    h += '<section class="block"><h2>Ssenarilər <small>adlandırılmış dəyişiklik dəstləri</small></h2><div class="row-gap"><input id="scn-name" placeholder="Ssenari adı, məs. Neft 60$ + üzən məzənnə" style="height:32px;border:1px solid var(--line);border-radius:6px;padding:0 10px;min-width:260px"><button class="btn btn-primary" id="scn-save">Cari dəyişiklikləri saxla</button></div><div id="scn-list" style="margin-top:10px"></div></section>';
    h += '<section class="block"><h2>Cari dəyişikliklər <small>' + editCount() + '</small></h2><div id="adm-changes"></div></section>';
    h += '<section class="block"><h2>Təsir</h2><div id="adm-impact"></div></section>';
    main.innerHTML = h;
    wireMacroButtons();
    $('#adm-toggle').onclick = function () { setAdmin(!ADMIN); pageAdmin(); };
    $('#adm-export').onclick = function () { download('caem-deyisiklikler.json', JSON.stringify(serialize(), null, 1), 'application/json'); };
    $('#adm-import').onchange = function () {
      var f = this.files[0]; if (!f) return; var rd = new FileReader();
      rd.onload = function () { try { applySerialized(JSON.parse(rd.result)); pageAdmin(); } catch (e) { toast('Fayl oxunmadı: JSON formatı səhvdir'); } };
      rd.readAsText(f);
    };
    $('#adm-full').onclick = function () { var t = performance.now(); M.recalcAll(EDITS.ov); toast('Tam hesablama: ' + fmtInt(M.order.length) + ' düstur, ' + Math.round(performance.now() - t) + ' ms' + cycNote()); pageAdmin(); };
    $('#adm-reset').onclick = function () {
      if (!editCount()) { toast('Dəyişiklik yoxdur'); return; }
      if (!confirm('Bütün ' + editCount() + ' dəyişiklik silinsin və model baza dəyərlərinə qaytarılsın?')) return;
      clearAllEdits(); for (var i = 0; i < M.N; i++) M.V[i] = M.V0[i]; M.recalcAll(); persist(); toast('Model baza dəyərlərinə qaytarıldı'); pageAdmin();
    };
    $('#scn-save').onclick = function () {
      var name = $('#scn-name').value.trim(); if (!name) { toast('Ssenari adını yazın'); return; }
      var all = scenarios(); all[name] = serialize(); lsSet(LS_SCEN, JSON.stringify(all)); toast('Saxlanıldı: ' + name); drawScen();
    };
    drawScen(); drawChanges(); drawImpact();
  }
  function pageUpdate() {
    var h = '<div class="eyebrow">İdarəetmə</div><div class="page-h"><div><h1>Model yeniləmə</h1><p class="lead">İl başa çatanda həmin ilin faktiki (realizə olunmuş) məlumatını daxil edin və proqnozun ilk ilini bir il irəli çəkin — CAEM-in öz qaydası ilə (INPUT vərəqində «First Forecast Year»). Və ya Nazirliyin yenilənmiş CAEM Excel faylını yükləyin. Hər əməliyyatdan sonra bütün model avtomatik yenidən hesablanır; nəticələr həm bu görünüşdə, həm də iş panelində görünür.</p></div></div>';
    h += '<section class="block" id="mu-host"></section>';
    main.innerHTML = h;
    window.CaemUpdate.render($('#mu-host'), M, INFO);
  }
  function scenarios() { try { return JSON.parse(lsGet(LS_SCEN) || '{}'); } catch (e) { return {}; } }
  function drawScen() {
    var all = scenarios(), ks = Object.keys(all), el = $('#scn-list'); if (!el) return;
    if (!ks.length) { el.innerHTML = '<p class="muted small">Saxlanmış ssenari yoxdur. Ssenarilər yalnız bu brauzerdə saxlanılır; paylaşmaq üçün JSON ixracından istifadə edin.</p>'; return; }
    el.innerHTML = '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Ad</th><th>Saxlanıb</th><th class="n">Dəyişiklik</th><th></th></tr></thead><tbody>' + ks.map(function (k) {
      var s = all[k], n = Object.keys(s.inputs || {}).length + Object.keys(s.coefficients || {}).length + Object.keys(s.overrides || {}).length;
      return '<tr><td><b>' + esc(k) + '</b></td><td class="small">' + esc((s.saved || '').replace('T', ' ').slice(0, 16)) + '</td><td class="n">' + n + '</td><td><div class="row-gap"><button class="btn btn-sm" data-load="' + esc(k) + '">Yüklə</button><button class="btn btn-sm btn-danger" data-del="' + esc(k) + '">Sil</button></div></td></tr>';
    }).join('') + '</tbody></table></div>';
    el.onclick = function (e) {
      var l = e.target.getAttribute('data-load'), d = e.target.getAttribute('data-del');
      if (l) { applySerialized(all[l]); pageAdmin(); }
      if (d && confirm('«' + d + '» ssenarisi silinsin?')) { delete all[d]; lsSet(LS_SCEN, JSON.stringify(all)); drawScen(); }
    };
  }
  function drawChanges() {
    var el = $('#adm-changes'), rows = [];
    EDITS.inputs.forEach(function (v, id) { rows.push({ kind: 'Giriş', id: id, v: v, v0: M.V0[id], undo: function () { setInput(id, M.V0[id]); } }); });
    EDITS.K.forEach(function (v, k) { var id = KOWNER[k]; rows.push({ kind: 'Əmsal', id: id, v: v, v0: M.K0[k], lit: k - M.fk[id], undo: function () { setCoef([k], M.K0[k]); } }); });
    EDITS.ov.forEach(function (v, id) { rows.push({ kind: 'Üzərinə yazılıb', id: id, v: v, v0: M.V0[id], undo: function () { setOverride(id, undefined); } }); });
    if (!rows.length) { el.innerHTML = '<p class="muted">Dəyişiklik yoxdur — model ' + (INFO.rebased ? 'baza model' : 'Excel faylındakı') + ' dəyərlərlə eynidir.</p>'; return; }
    el.innerHTML = '<div class="tbl-wrap" style="max-height:60vh;overflow:auto"><table class="tbl"><thead><tr><th>Növ</th><th>Vərəq › xana</th><th>Göstərici</th><th class="n">' + INFO.baseShort + '</th><th class="n">Yeni</th><th></th></tr></thead><tbody>' + rows.map(function (x, i) {
      return '<tr><td><span class="chip ' + (x.kind === 'Əmsal' ? 'chip-coef' : x.kind === 'Giriş' ? 'chip-info' : 'chip-warn') + '">' + x.kind + '</span></td><td><a href="#/s/' + M.sh[x.id] + '/cedvel/' + a1(M.r[x.id], M.c[x.id]) + '"><code class="f">' + esc(fullAddr(x.id)) + '</code></a>' + (NAMES_OF[x.id] ? ' <code class="muted small">' + esc(NAMES_OF[x.id][0]) + '</code>' : '') + (x.lit !== undefined ? ' <span class="muted small">(' + (x.lit + 1) + '-ci sabit)</span>' : '') + '</td><td class="small">' + esc(rowLabel(M.sh[x.id], M.r[x.id])) + ' <span class="muted">' + (colYear(M.sh[x.id], M.r[x.id], M.c[x.id]) || '') + '</span></td><td class="n">' + esc(shortVal(x.v0)) + '</td><td class="n"><b>' + esc(shortVal(x.v)) + '</b></td><td><button class="btn btn-sm" data-undo="' + i + '">Geri qaytar</button></td></tr>';
    }).join('') + '</tbody></table></div>';
    el.onclick = function (e) { var i = e.target.getAttribute('data-undo'); if (i !== null) { rows[+i].undo(); pageAdmin(); } };
  }
  function drawImpact() {
    var el = $('#adm-impact'); if (!el) return;
    var n = 0, byG = {};
    for (var j = 0; j < M.order.length; j++) { var id = M.order[j]; if (M.sh[id] === ARR_GS) continue; if (!sameVal(M.V[id], M.V0[id])) { n++; var g = GROUP_OF[M.sh[id]]; byG[g] = (byG[g] || 0) + 1; } }
    if (!n) { el.innerHTML = '<p class="muted">Bütün düsturların nəticəsi ' + (INFO.rebased ? 'baza modelin' : 'Excel faylındakı') + ' dəyərlərlə eynidir.</p>'; return; }
    el.innerHTML = '<p><b>' + fmtInt(n) + '</b> düsturun nəticəsi ' + (INFO.rebased ? 'baza model' : 'Excel') + ' dəyərindən fərqlənir.</p><div class="row-gap">' + Object.keys(byG).map(function (g) { var G = GROUPS[g]; return G ? '<span class="chip" style="background:#fff;border:1px solid var(--line)"><span class="dot" style="background:' + G.color + '"></span>' + esc(G.title) + ': ' + fmtInt(byG[g]) + '</span>' : ''; }).join('') + '</div>';
  }
  function download(name, text, type) {
    try {
      var blob = text instanceof Blob ? text : new Blob([text], { type: type }), a = document.createElement('a');
      a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    } catch (e) { toast('Yükləmə bu mühitdə mümkün olmadı'); }
  }
  function exportCSV(gs) {
    var ids = SC[gs]; if (!ids.length) return;
    var maxR = 0, maxC = 0; ids.forEach(function (id) { if (M.V[id] === null) return; if (M.r[id] > maxR) maxR = M.r[id]; if (M.c[id] > maxC) maxC = M.c[id]; });
    var lines = [];
    for (var r = 1; r <= maxR; r++) {
      var row = [];
      for (var c = 1; c <= maxC; c++) {
        var id = cellId(gs, r, c), v = id < 0 ? '' : M.V[id];
        var s = v === null ? '' : v instanceof X.XErr ? v.e : String(v);
        row.push(/[",;\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
      }
      lines.push(row.join(';'));
    }
    download(('CAEM - ' + SHEETS[gs].n.trim()).replace(/[\\/:*?"<>|]/g, '_') + '.csv', '﻿' + lines.join('\r\n'), 'text/csv');
  }

  // ------------------------------------------------------------------ GRID (2-D virtualised, Excel-like frozen panes)
  function Grid(host, gs) {
    this.host = host; this.gs = gs; this.showHidden = false;
    var meta = SHEETS[gs], ids = SC[gs], maxR = 0, maxC = 0;
    ids.forEach(function (id) { if (M.r[id] > maxR) maxR = M.r[id]; if (M.c[id] > maxC) maxC = M.c[id]; });
    this.maxR = Math.max(maxR + 3, 30); this.maxC = Math.max(maxC + 2, 12);
    this.hiddenRows = new Set(meta.hr || []);
    this.colW = new Array(this.maxC + 1); this.hiddenCols = new Set();
    var dw = Math.round((meta.dw || 9.14) * 7 + 5);
    for (var c = 1; c <= this.maxC; c++) this.colW[c] = dw;
    (meta.cols || []).forEach(function (cd) {
      for (var c2 = cd.min; c2 <= Math.min(cd.max, this.maxC); c2++) { this.colW[c2] = Math.max(18, Math.round(cd.w * 7 + 5)); if (cd.hidden || cd.w === 0) this.hiddenCols.add(c2); }
    }, this);
    this.fz = (meta.fz || [0, 0]).slice();
    // no freeze defined in Excel: freeze the label columns left of the first year column (and the year header row)
    var yr0 = yearRows(gs)[0];
    if (yr0) {
      var firstYearCol = Math.min.apply(null, Object.keys(yr0.d).map(Number));
      // Excel freezes nothing useful (e.g. only a hidden column A): keep the label columns left of the first year visible
      var visW = 0; for (var c3 = 1; c3 <= (this.fz[1] || 0); c3++) if (!this.hiddenCols.has(c3)) visW += this.colW[c3];
      if ((!this.fz[1] || visW < 40) && firstYearCol > 1 && firstYearCol <= 8 && firstYearCol - 1 > (this.fz[1] || 0)) this.fz[1] = firstYearCol - 1;
      if (!this.fz[0] && yr0.r <= 6) this.fz[0] = yr0.r;
    }
    this.comments = meta.cm || {};
    var self = this;
    host.innerHTML = '<div class="gtop"><div class="gcorner"></div><div class="gtopmain"></div></div><div class="gbody"><div class="gleft"></div><div class="gmain"></div></div>';
    this.L = { corner: $('.gcorner', host), top: $('.gtopmain', host), left: $('.gleft', host), main: $('.gmain', host) };
    this.layout();
    this._raf = 0;
    this.onScroll = function () { if (!self._raf) self._raf = requestAnimationFrame(function () { self._raf = 0; self.render(); }); };
    host.addEventListener('scroll', this.onScroll, { passive: true });
    this.onResize = function () { self.render(true); };
    window.addEventListener('resize', this.onResize);
    host.addEventListener('mousedown', function (e) { self.onMouse(e); });
    host.addEventListener('dblclick', function (e) { var p = self.hit(e); if (p) self.startEdit(p.r, p.c); });
    host.addEventListener('keydown', function (e) { self.onKey(e); });
    this.render(true);
  }
  Grid.prototype.destroy = function () { this.host.removeEventListener('scroll', this.onScroll); window.removeEventListener('resize', this.onResize); };
  Grid.prototype.layout = function () {
    var fzR = this.fz[0], fzC = this.fz[1];
    this.rows = []; this.frows = [];
    for (var r = 1; r <= this.maxR; r++) { if (!this.showHidden && this.hiddenRows.has(r)) continue; (r <= fzR ? this.frows : this.rows).push(r); }
    this.cols = []; this.fcols = [];
    for (var c = 1; c <= this.maxC; c++) { if (!this.showHidden && this.hiddenCols.has(c)) continue; (c <= fzC ? this.fcols : this.cols).push(c); }
    // guard: frozen label area wider than 45% of the view → shrink those columns proportionally (min 36px)
    var fw = 0; this.fcols.forEach(function (c2) { fw += this.colW[c2]; }, this);
    var maxFw = (this.host.clientWidth || 900) * 0.45;
    if (fw > maxFw) {
      var k = maxFw / fw; fw = 0;
      this.fcols.forEach(function (c2) { this.colW[c2] = Math.max(36, Math.round(this.colW[c2] * k)); fw += this.colW[c2]; }, this);
      if (fw > maxFw * 1.3) { this.cols = this.fcols.concat(this.cols); this.fcols = []; fw = 0; }
    }
    if (this.frows.length * RH > (this.host.clientHeight || 500) * 0.4) { this.rows = this.frows.concat(this.rows); this.frows = []; }
    this.FW = GUT + fw; this.FH = HDR + this.frows.length * RH;
    this.fx = {}; var x = GUT; this.fcols.forEach(function (c2) { this.fx[c2] = x; x += this.colW[c2]; }, this);
    this.mx = []; this.mxc = {}; x = 0; this.cols.forEach(function (c2) { this.mx.push(x); this.mxc[c2] = x; x += this.colW[c2]; }, this);
    this.MW = x; this.MH = this.rows.length * RH;
    this.rowIdx = {}; this.rows.forEach(function (r2, i) { this.rowIdx[r2] = i; }, this);
    this.frowIdx = {}; this.frows.forEach(function (r2, i) { this.frowIdx[r2] = i; }, this);
    var L = this.L;
    L.corner.style.cssText = 'width:' + this.FW + 'px;height:' + this.FH + 'px';
    L.top.style.cssText = 'width:' + this.MW + 'px;height:' + this.FH + 'px';
    L.left.style.cssText = 'width:' + this.FW + 'px;height:' + this.MH + 'px';
    L.main.style.cssText = 'width:' + this.MW + 'px;height:' + this.MH + 'px';
    // sticky panes only stick inside their parent row, so the rows must span the full content width
    L.corner.parentNode.style.width = (this.FW + this.MW) + 'px';
    L.left.parentNode.style.width = (this.FW + this.MW) + 'px';
    this.yrs = yearRows(this.gs);
  };
  Grid.prototype.colHeaderText = function (c) {
    // show year above column letter when the sheet's first year row defines it
    var y = null; for (var i = 0; i < this.yrs.length; i++) if (this.yrs[i].d[c] !== undefined) { y = this.yrs[i].d[c]; break; }
    return { t: colName(c), y: y };
  };
  Grid.prototype.cellHTML = function (id, x, y, w, rowCellsMap, colsList, colXs) {
    var st = STY[M.st[id]] || {}, v = M.V[id], txt = dispVal(v, st), cls = 'c', css = '';
    var isNum = typeof v === 'number', isErr = v instanceof X.XErr;
    if (M.ft[id] >= 0) { cls += isMember(id) || CORE.disp[M.ft[id]].charAt(0) === '{' ? ' xb' : ' f'; if (EDITS.ov.has(id)) cls += ' ed'; else if (!sameVal(v, M.V0[id])) cls += ' chg'; }
    else if (typeof M.V0[id] === 'number') { if (isCoefCell(id)) cls += ' coefc'; else if (M.dependents()[id]) cls += ' inp'; if (EDITS.inputs.has(id)) cls += ' ed'; }
    else if (SWITCH[id] || (typeof M.V0[id] === 'string' && M.dependents()[id] && EDITS.inputs.has(id))) { cls += ' sw'; if (EDITS.inputs.has(id)) cls += ' ed'; }
    if (isErr) cls += ' err';
    if (isNum || isErr) cls += ' num'; else cls += ' t';
    if (this.comments[a1(M.r[id], M.c[id])]) cls += ' cm';
    if (st.b) css += 'font-weight:700;';
    if (st.i) css += 'font-style:italic;';
    if (st.fc && !isErr && M.ft[id] < 0 && typeof M.V0[id] !== 'number') css += 'color:#' + st.fc + ';';
    if (st.bg) css += 'background:' + soften(st.bg) + ';';
    if (st.ha === 'center' || st.ha === 'centerContinuous') css += 'text-align:center;';
    else if (st.ha === 'right') css += 'text-align:right;';
    else if (st.ha === 'left' && isNum) css += 'text-align:left;';
    if (st.ind) css += 'padding-left:' + (5 + st.ind * 9) + 'px;';
    if (!isNum && !isErr && txt && rowCellsMap) {
      // let text spill into following empty columns (Excel behaviour)
      var ci = colsList.indexOf(M.c[id]), extra = 0;
      for (var k = ci + 1; k < colsList.length && k <= ci + 12; k++) { if (rowCellsMap[colsList[k]] !== undefined) break; extra += this.colW[colsList[k]]; }
      if (extra) { w += extra; css += 'border-right-color:transparent;'; }
      if (!st.bg && extra) css += 'z-index:1;';
    }
    return '<div class="' + cls + '" style="left:' + x + 'px;top:' + y + 'px;width:' + w + 'px;' + css + '">' + esc(txt) + '</div>';
  };
  Grid.prototype.render = function (force) {
    var H = this.host, st = H.scrollTop, sl = H.scrollLeft, vw = H.clientWidth, vh = H.clientHeight;
    var r0 = Math.max(0, Math.floor(st / RH) - 4), r1 = Math.min(this.rows.length, Math.ceil((st + vh) / RH) + 4);
    // visible main columns
    var xs = this.mx, lo = 0, hi = xs.length - 1, left = sl - 200, right = sl + vw;
    while (lo < hi) { var mid = (lo + hi + 1) >> 1; if (xs[mid] <= left) lo = mid; else hi = mid - 1; }
    var c0 = lo, c1 = c0; while (c1 < xs.length && xs[c1] < right) c1++;
    var key = r0 + ':' + r1 + ':' + c0 + ':' + c1;
    if (!force && key === this._key) return;
    this._key = key;
    var visCols = this.cols.slice(c0, c1), allCols = this.cols, self = this;
    // --- main
    var h = '';
    for (var i = r0; i < r1; i++) {
      var r = this.rows[i], y = i * RH, rc = rowCells(this.gs, r), rmap = {};
      rc.forEach(function (id) { rmap[M.c[id]] = id; });
      visCols.forEach(function (c) { var id = rmap[c]; if (id !== undefined) h += self.cellHTML(id, self.mxc[c], y, self.colW[c], rmap, allCols); });
    }
    h += this.selHTML('main');
    this.L.main.innerHTML = h;
    // --- left (row numbers + frozen columns)
    h = '';
    for (i = r0; i < r1; i++) {
      r = this.rows[i]; y = i * RH;
      h += '<div class="hd' + (SEL && SEL.gs === this.gs && SEL.r === r ? ' selhd' : '') + '" style="left:0;top:' + y + 'px;width:' + GUT + 'px">' + r + '</div>';
      if (this.fcols.length) {
        var rm2 = {}; rowCells(this.gs, r).forEach(function (id) { rm2[M.c[id]] = id; });
        this.fcols.forEach(function (c) { var id = rm2[c]; if (id !== undefined) h += self.cellHTML(id, self.fx[c], y, self.colW[c], rm2, self.fcols); });
      }
    }
    h += this.selHTML('left');
    this.L.left.innerHTML = h;
    // --- top (column letters + frozen rows)
    h = '';
    visCols.forEach(function (c) {
      var ht = self.colHeaderText(c);
      h += '<div class="hd' + (ht.y ? ' yr' : '') + (SEL && SEL.gs === self.gs && SEL.c === c ? ' selhd' : '') + '" style="left:' + self.mxc[c] + 'px;top:0;width:' + self.colW[c] + 'px" title="Sütun ' + ht.t + (ht.y ? ' · ' + ht.y : '') + '">' + (ht.y ? ht.t + ' · ' + ht.y : ht.t) + '</div>';
    });
    this.frows.forEach(function (r2, fi) {
      var rm3 = {}; rowCells(self.gs, r2).forEach(function (id) { rm3[M.c[id]] = id; });
      visCols.forEach(function (c) { var id = rm3[c]; if (id !== undefined) h += self.cellHTML(id, self.mxc[c], HDR + fi * RH, self.colW[c], rm3, allCols); });
    });
    h += this.selHTML('top');
    this.L.top.innerHTML = h;
    // --- corner
    h = '<div class="hd" style="left:0;top:0;width:' + GUT + 'px"></div>';
    this.fcols.forEach(function (c) { var ht = self.colHeaderText(c); h += '<div class="hd' + (SEL && SEL.gs === self.gs && SEL.c === c ? ' selhd' : '') + '" style="left:' + self.fx[c] + 'px;top:0;width:' + self.colW[c] + 'px">' + ht.t + '</div>'; });
    this.frows.forEach(function (r2, fi) {
      h += '<div class="hd" style="left:0;top:' + (HDR + fi * RH) + 'px;width:' + GUT + 'px">' + r2 + '</div>';
      var rm4 = {}; rowCells(self.gs, r2).forEach(function (id) { rm4[M.c[id]] = id; });
      self.fcols.forEach(function (c) { var id = rm4[c]; if (id !== undefined) h += self.cellHTML(id, self.fx[c], HDR + fi * RH, self.colW[c], rm4, self.fcols); });
    });
    h += this.selHTML('corner');
    this.L.corner.innerHTML = h;
  };
  Grid.prototype.where = function (r, c) {
    var inFr = this.frowIdx[r] !== undefined, inFc = this.fx[c] !== undefined;
    var y = inFr ? HDR + this.frowIdx[r] * RH : (this.rowIdx[r] !== undefined ? this.rowIdx[r] * RH : null);
    var x = inFc ? this.fx[c] : (this.mxc[c] !== undefined ? this.mxc[c] : null);
    if (y === null || x === null) return null;
    return { layer: inFr ? (inFc ? 'corner' : 'top') : (inFc ? 'left' : 'main'), x: x, y: y, w: this.colW[c] };
  };
  Grid.prototype.selHTML = function (layer) {
    if (!SEL || SEL.gs !== this.gs) return '';
    var p = this.where(SEL.r, SEL.c); if (!p || p.layer !== layer) return '';
    return '<div class="selbox" style="left:' + (p.x - 1) + 'px;top:' + (p.y - 1) + 'px;width:' + (p.w + 1) + 'px;height:' + (RH + 1) + 'px"></div>';
  };
  Grid.prototype.hit = function (e) {
    var t = e.target, layer = null;
    ['corner', 'top', 'left', 'main'].forEach(function (k) { if (this.L[k].contains(t)) layer = k; }, this);
    if (!layer) return null;
    var rect = this.L[layer].getBoundingClientRect(), x = e.clientX - rect.left, y = e.clientY - rect.top, r = null, c = null, i;
    if (layer === 'main' || layer === 'left') { i = Math.floor(y / RH); r = this.rows[i]; }
    else { if (y < HDR) return null; i = Math.floor((y - HDR) / RH); r = this.frows[i]; }
    if (layer === 'main' || layer === 'top') {
      var xs = this.mx, lo = 0, hi = xs.length - 1; while (lo < hi) { var mid = (lo + hi + 1) >> 1; if (xs[mid] <= x) lo = mid; else hi = mid - 1; } c = this.cols[lo];
    } else {
      if (x < GUT) return null;
      for (var k = 0; k < this.fcols.length; k++) { var cc = this.fcols[k]; if (x >= this.fx[cc] && x < this.fx[cc] + this.colW[cc]) c = cc; }
    }
    return r && c ? { r: r, c: c } : null;
  };
  Grid.prototype.onMouse = function (e) {
    if (e.target.classList && e.target.classList.contains('cell-editor')) return;
    var p = this.hit(e); if (!p) return;
    this.host.focus({ preventScroll: true });
    this.select(p.r, p.c, false);
  };
  Grid.prototype.select = function (r, c, scroll) {
    SEL = { gs: this.gs, r: r, c: c };
    if (scroll) this.scrollTo(r, c);
    this.render(true);
    renderInspector(); updateStatus();
    var h = '#/s/' + this.gs + '/cedvel/' + a1(r, c);
    if (location.hash !== h) history.replaceState(null, '', h);
  };
  Grid.prototype.scrollTo = function (r, c) {
    var H = this.host;
    if (this.rowIdx[r] !== undefined) { var y = this.rowIdx[r] * RH; if (y < H.scrollTop || y > H.scrollTop + H.clientHeight - this.FH - RH * 2) H.scrollTop = Math.max(0, y - (H.clientHeight - this.FH) / 3); }
    if (this.mxc[c] !== undefined) { var x = this.mxc[c]; if (x < H.scrollLeft || x + this.colW[c] > H.scrollLeft + H.clientWidth - this.FW) H.scrollLeft = Math.max(0, x - (H.clientWidth - this.FW) / 3); }
  };
  Grid.prototype.onKey = function (e) {
    if (!SEL || SEL.gs !== this.gs || e.target !== this.host) return;
    var rows = this.frows.concat(this.rows), cols = this.fcols.concat(this.cols), ri = rows.indexOf(SEL.r), ci = cols.indexOf(SEL.c);
    var mv = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Tab: [0, e.shiftKey ? -1 : 1] }[e.key];
    if (mv) {
      e.preventDefault();
      var nr = rows[Math.max(0, Math.min(rows.length - 1, ri + mv[0]))], nc = cols[Math.max(0, Math.min(cols.length - 1, ci + mv[1]))];
      this.select(nr, nc, true); return;
    }
    if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); this.startEdit(SEL.r, SEL.c); return; }
    if (ADMIN && e.key.length === 1 && /[0-9.,\-]/.test(e.key) && !e.ctrlKey && !e.metaKey) { e.preventDefault(); this.startEdit(SEL.r, SEL.c, e.key); }
  };
  Grid.prototype.startEdit = function (r, c, initial) {
    var id = cellId(this.gs, r, c);
    if (id < 0) return;
    if (M.ft[id] >= 0) { toast('Bu düstur xanasıdır — dəyəri müfəttiş panelində üzərinə yazmaq olar (Admin)'); return; }
    if (LOCKED[id]) { toast(LOCKED[id]); return; }
    if (typeof M.V0[id] !== 'number') { if (SWITCH[id] || M.dependents()[id]) { var s0 = $('#ins-sw', inspector) || $('#ins-in', inspector); if (s0) s0.focus(); else toast('Mətn açarını sağ paneldə dəyişin (Admin)'); } return; }
    if (!ADMIN) { toast('Dəyişmək üçün yuxarıdan Admin rejimini aktivləşdirin'); return; }
    var p = this.where(r, c); if (!p) return;
    var ed = document.createElement('input'), self = this, done = false;
    ed.className = 'cell-editor'; ed.setAttribute('inputmode', 'decimal');
    ed.style.cssText = 'left:' + (p.x - 1) + 'px;top:' + (p.y - 1) + 'px;width:' + Math.max(p.w + 2, 90) + 'px';
    ed.value = initial !== undefined ? initial : fmtEditVal(M.V[id]);
    this.L[p.layer].appendChild(ed); ed.focus(); if (initial === undefined) ed.select();
    function finish(commit) {
      if (done) return; done = true;
      var v = parseUserNumber(ed.value); ed.remove(); self.host.focus({ preventScroll: true });
      if (!commit) return;
      if (v === undefined || v === null) { toast('Rəqəm daxil edin'); return; }
      if (!sameVal(v, M.V[id])) setInput(id, v);
    }
    ed.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); finish(true); var rows = self.frows.concat(self.rows), i = rows.indexOf(r); if (i < rows.length - 1) self.select(rows[i + 1], c, true); }
      else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
      e.stopPropagation();
    });
    ed.addEventListener('blur', function () { finish(true); });
  };
  var SOFT = {};
  function soften(hex) {
    if (SOFT[hex]) return SOFT[hex];
    var r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16), a = 0.55;
    var f = function (x) { return Math.round(x * a + 255 * (1 - a)); };
    return (SOFT[hex] = 'rgb(' + f(r) + ',' + f(g) + ',' + f(b) + ')');
  }
  function updateStatus() {
    var el = $('#gstatus'); if (!el) return;
    if (!SEL) { el.innerHTML = '<span class="muted">Xananı seçin — düstur, əlaqələr və əmsallar sağ paneldə göstəriləcək. Klaviatura: oxlar, Enter (redaktə, Admin).</span>'; return; }
    var id = cellId(SEL.gs, SEL.r, SEL.c);
    el.innerHTML = '<b>' + a1(SEL.r, SEL.c) + '</b><span class="fx">' + (id < 0 ? '' : M.ft[id] >= 0 ? esc(formulaHTML(id, { plain: true })) : esc(fullVal(M.V[id]))) + '</span>';
  }


  // ------------------------------------------------------------------ INSPECTOR
  function closeInspector() { inspector.hidden = true; }
  function renderInspector() {
    if (!SEL) { closeInspector(); return; }
    var gs = SEL.gs, r = SEL.r, c = SEL.c, id = cellId(gs, r, c);
    var lab = rowLabel(gs, r), yr = colYear(gs, r, c);
    var h = '<div class="ins-h"><div><div class="eyebrow">' + esc(sheetTitle(gs)) + '</div><h3>' + esc(lab || 'Xana ' + a1(r, c)) + '</h3><div class="ins-addr">' + esc(SHEETS[gs].n) + '!' + a1(r, c) + (yr ? ' · <b>' + yr + '</b>' : '') + '</div></div><button class="icon-btn" id="ins-close" aria-label="Paneli bağla">✕</button></div>';
    if (id < 0) { h += '<p class="muted" style="margin-top:14px">Boş xana.</p>'; inspector.innerHTML = h; inspector.hidden = false; wireInspector(); return; }
    var v = M.V[id], v0 = M.V0[id], isF = M.ft[id] >= 0, D = M.dependents(), mem = isMember(id);
    var kind = isF ? (mem ? '<span class="chip chip-link">Massiv düsturunun elementi</span>' : CORE.disp[M.ft[id]].charAt(0) === '{' ? '<span class="chip chip-link">Massiv düsturu</span>' : '<span class="chip">Düstur</span>')
      : (typeof v0 === 'number' ? (isCoefCell(id) ? '<span class="chip chip-coef">Əmsal</span>' : D[id] ? '<span class="chip chip-info">Giriş məlumatı</span>' : '<span class="chip">Sabit (heç bir düstur istifadə etmir)</span>') : (SWITCH[id] ? '<span class="chip chip-info">Açar (seçim)</span>' : D[id] ? '<span class="chip chip-info">Mətn girişi</span>' : '<span class="chip">Mətn</span>'));
    h += '<div class="ins-sec"><div class="row-gap">' + kind + (NAMES_OF[id] ? NAMES_OF[id].map(function (n) { return '<span class="chip chip-coef" title="Adlandırılmış xana">' + esc(n) + '</span>'; }).join('') : '') + (EDITS.inputs.has(id) || EDITS.ov.has(id) ? '<span class="chip chip-warn">dəyişdirilib</span>' : (!sameVal(v, v0) ? '<span class="chip chip-warn">nəticə dəyişib</span>' : '')) + '</div>';
    h += '<div class="ins-val" style="margin-top:8px">' + esc(dispVal(v, STY[M.st[id]]) || '—') + '</div><div class="small muted">Tam dəyər: ' + esc(fullVal(v)) + (!sameVal(v, v0) ? ' · ' + INFO.baseShort + ': ' + esc(fullVal(v0)) : '') + '</div></div>';
    if (LOCKED[id]) h += '<div class="ins-sec"><div class="note small">' + esc(LOCKED[id]) + ' <a href="#/yenile">Model yeniləmə</a></div></div>';
    if (isF) {
      h += '<div class="ins-sec"><h4>Düstur</h4><div class="formula-box" id="ins-formula">' + formulaHTML(id) + '</div>';
      var nl = TPL_NL[M.ft[mem ? groupOfMember(id) : id]];
      if (nl && !mem) h += '<div class="small muted" style="margin-top:4px">Bənövşəyi rəqəmlər düsturun içindəki əmsallardır' + (ADMIN ? ' — klikləyin və yeni dəyər yazın (yalnız bu xana üçün). Bütün illər üçün: <a href="#/s/' + gs + '/emsal">Əmsallar</a>.' : '. Dəyişmək üçün Admin rejimi.') + '</div>';
      if (mem) { var g = groupRange(groupOfMember(id)); h += '<div class="small muted" style="margin-top:4px">Excel-də bir massiv düsturu ' + a1(g.r1, g.c1) + ':' + a1(g.r2, g.c2) + ' diapazonunu bütövlükdə doldurur; bu xana həmin massivin ' + (r - g.r1 + 1) + '-ci sətir, ' + (c - g.c1 + 1) + '-ci sütun elementidir.</div>'; }
      h += '</div>';
    }
    if (SHEETS[gs].cm && SHEETS[gs].cm[a1(r, c)]) h += '<div class="ins-sec"><h4>Şərh (Excel)</h4><div class="note small">' + esc(SHEETS[gs].cm[a1(r, c)]) + '</div></div>';
    if (ADMIN && !isF && !LOCKED[id] && (typeof v0 === 'number' || SWITCH[id] || (D[id] && typeof v0 === 'string'))) {
      var ctl = SWITCH[id] ? '<select id="ins-sw" class="v">' + (function () { var o = SWITCH[id].options.slice(); if (o.indexOf(v) < 0) o.unshift(v); return o.map(function (x) { return '<option' + (x === v ? ' selected' : '') + '>' + esc(x) + '</option>'; }).join(''); })() + '</select>'
        : typeof v0 === 'string' ? '<input id="ins-in" class="v" data-kind="text" value="' + esc(v == null ? '' : v) + '">' : '<input id="ins-in" class="v" data-id="' + id + '" value="' + esc(fmtEditVal(v)) + '" inputmode="decimal">';
      h += '<div class="ins-sec"><h4>Dəyəri dəyiş (Admin)</h4><div class="edit-row">' + ctl + (SWITCH[id] ? '' : '<button class="btn btn-primary btn-sm" id="ins-apply">Tətbiq et</button>') + (EDITS.inputs.has(id) ? '<button class="btn btn-sm" id="ins-reset">' + esc(INFO.baseShort.replace('-də', '')) + ' dəyəri</button>' : '') + '</div>' + (SWITCH[id] ? '<div class="small muted" style="margin-top:4px">' + esc(SWITCH[id].label || '') + '</div>' : '') + '</div>';
    }
    if (ADMIN && isF && !mem) {
      h += '<div class="ins-sec"><h4>Düsturun üzərinə yaz (Admin)</h4><div class="edit-row"><input id="ins-ov" value="' + (EDITS.ov.has(id) ? esc(fmtEditVal(EDITS.ov.get(id))) : '') + '" placeholder="sabit dəyər" inputmode="decimal"><button class="btn btn-sm" id="ins-ov-apply">Yaz</button>' + (EDITS.ov.has(id) ? '<button class="btn btn-sm" id="ins-ov-clear">Düsturu bərpa et</button>' : '') + '</div><div class="small muted" style="margin-top:4px">Düstur hesablanmır, yazılan dəyər asılı xanalara ötürülür.</div></div>';
    }
    if (isF) {
      var pr = uniq(M.precedents(mem ? groupOfMember(id) : id).map(realCell));
      h += '<div class="ins-sec"><h4>Mənbələr (' + pr.length + ')</h4>' + depList(pr, 40) + '</div>';
    }
    var ds = uniq((D[id] || []).map(realCell));
    h += '<div class="ins-sec"><h4>Asılı xanalar (' + ds.length + ')</h4>' + (ds.length ? depList(ds, 40) : '<p class="small muted">Bu xananı heç bir düstur istifadə etmir.</p>') + '</div>';
    inspector.innerHTML = h; inspector.hidden = false;
    wireInspector(id);
  }
  function uniq(a) { return Array.from(new Set(a.filter(function (x) { return x >= 0; }))); }
  function depList(ids, max) {
    ids = ids.slice().sort(function (a, b) { return M.sh[a] - M.sh[b] || M.r[a] - M.r[b] || M.c[a] - M.c[b]; });
    var h = '<div class="dep-list">';
    ids.slice(0, max).forEach(function (p) {
      var g = M.sh[p], y = colYear(g, M.r[p], M.c[p]);
      h += '<a class="dep" data-gs="' + g + '" data-r="' + M.r[p] + '" data-c="' + M.c[p] + '"><span><span class="a">' + (g !== SEL.gs ? esc(SHEETS[g].n) + '!' : '') + a1(M.r[p], M.c[p]) + (y ? ' · ' + y : '') + (NAMES_OF[p] ? ' · ' + esc(NAMES_OF[p][0]) : '') + '</span><span class="l">' + esc(rowLabel(g, M.r[p]) || '') + '</span></span><span class="v">' + esc(dispVal(M.V[p], STY[M.st[p]])) + '</span></a>';
    });
    if (ids.length > max) h += '<div class="small muted" style="padding:4px 6px">… və daha ' + (ids.length - max) + '</div>';
    return h + '</div>';
  }
  function wireInspector(id) {
    $('#ins-close', inspector).onclick = closeInspector;
    inspector.onclick = function (e) {
      var a = e.target.closest('[data-gs]');
      if (a) { e.preventDefault(); gotoCell(+a.getAttribute('data-gs'), +a.getAttribute('data-r'), +a.getAttribute('data-c')); return; }
      var lit = e.target.closest('.lit');
      if (lit && ADMIN) editLiteral(lit);
    };
    var inp = $('#ins-in', inspector);
    if (inp) {
      var apply = function () {
        if (inp.getAttribute('data-kind') === 'text') { setInput(id, inp.value); return; }
        var v = parseUserNumber(inp.value); if (v === undefined || v === null) { toast('Rəqəm daxil edin'); return; } setInput(id, v);
      };
      $('#ins-apply', inspector).onclick = apply;
      inp.onkeydown = function (e) { if (e.key === 'Enter') apply(); };
    }
    var sw = $('#ins-sw', inspector); if (sw) sw.onchange = function () { setInput(id, sw.value); };
    var rs = $('#ins-reset', inspector); if (rs) rs.onclick = function () { setInput(id, M.V0[id]); };
    var ov = $('#ins-ov', inspector);
    if (ov) {
      var applyOv = function () { var v = parseUserNumber(ov.value); if (v === undefined || v === null) { toast('Rəqəm daxil edin'); return; } setOverride(id, v); };
      $('#ins-ov-apply', inspector).onclick = applyOv;
      ov.onkeydown = function (e) { if (e.key === 'Enter') applyOv(); };
      var oc = $('#ins-ov-clear', inspector); if (oc) oc.onclick = function () { setOverride(id, undefined); };
    }
  }
  function editLiteral(span) {
    var k = +span.getAttribute('data-k');
    var inp = document.createElement('input');
    inp.value = String(M.K[k]); inp.style.cssText = 'width:' + Math.max(80, span.offsetWidth + 30) + 'px;font:inherit;border:1px solid var(--coef);border-radius:3px;padding:0 3px';
    span.replaceWith(inp); inp.focus(); inp.select();
    var done = false;
    function fin(ok) {
      if (done) return; done = true;
      var v = parseUserNumber(inp.value);
      if (ok && v !== undefined && v !== null && !sameVal(v, M.K[k])) setCoef([k], v); else renderInspector();
    }
    inp.onkeydown = function (e) { if (e.key === 'Enter') fin(true); if (e.key === 'Escape') fin(false); };
    inp.onblur = function () { fin(true); };
  }
  function gotoCell(gs, r, c) {
    if (gs === ARR_GS) { var a = groupAnchor(cellId(gs, r, c)); if (a < 0) return; gs = M.sh[a]; r = M.r[a]; c = M.c[a]; }
    SEL = { gs: gs, r: r, c: c };
    if (GRID && GRID.gs === gs) { GRID.select(r, c, true); return; }
    go('#/s/' + gs + '/cedvel/' + a1(r, c));
  }

  // ------------------------------------------------------------------ search
  var SEARCH_IDX = null;
  function buildSearch() {
    SEARCH_IDX = [];
    SHEETS.forEach(function (s, gs) { if (isShown(gs)) SEARCH_IDX.push({ t: fold(sheetTitle(gs) + ' ' + s.n), gs: gs, sheet: true }); });
    Object.keys(NAMES_OF).forEach(function (k) { var x = +k; if (isShown(M.sh[x])) NAMES_OF[x].forEach(function (n) { SEARCH_IDX.push({ t: fold(n + ' ' + rowLabel(M.sh[x], M.r[x])), id: x, name: n }); }); });
    for (var id = 0; id < M.N; id++) {
      var v = M.V0[id];
      if (typeof v === 'string' && v.length > 2 && M.ft[id] < 0 && isShown(M.sh[id]) && /[A-Za-zƏəÜüÖöĞğİıŞşÇç]/.test(v)) SEARCH_IDX.push({ t: fold(v), id: id });
    }
  }
  function doSearch(q) {
    var box = $('#search-results');
    q = fold(q.trim());
    if (q.length < 2) { box.hidden = true; return; }
    if (!SEARCH_IDX) buildSearch();
    var terms = q.split(/\s+/), res = [];
    for (var i = 0; i < SEARCH_IDX.length && res.length < 60; i++) {
      var e = SEARCH_IDX[i], ok = true;
      for (var j = 0; j < terms.length; j++) if (e.t.indexOf(terms[j]) < 0) { ok = false; break; }
      if (ok) res.push(e);
    }
    if (!res.length) { box.innerHTML = '<div class="sr-empty">Nəticə tapılmadı</div>'; box.hidden = false; return; }
    box.innerHTML = res.map(function (e) {
      if (e.sheet) return '<a class="sr-item" href="#/s/' + e.gs + '/cedvel"><b>' + esc(sheetTitle(e.gs)) + '</b><small>Vərəq · ' + esc(SHEETS[e.gs].n) + '</small></a>';
      var gs = M.sh[e.id];
      if (e.name) return '<a class="sr-item" href="#/s/' + gs + '/cedvel/' + a1(M.r[e.id], M.c[e.id]) + '"><b>' + esc(e.name) + '</b> = ' + esc(shortVal(M.V[e.id])) + '<small>Ad · ' + esc(SHEETS[gs].n) + '!' + a1(M.r[e.id], M.c[e.id]) + '</small></a>';
      return '<a class="sr-item" href="#/s/' + gs + '/cedvel/' + a1(M.r[e.id], M.c[e.id]) + '">' + esc(M.V0[e.id].slice(0, 110)) + '<small>' + esc(SHEETS[gs].n) + '!' + a1(M.r[e.id], M.c[e.id]) + '</small></a>';
    }).join('');
    box.hidden = false;
  }

  // ------------------------------------------------------------------ admin toggle
  function setAdmin(on) {
    ADMIN = !!on; document.body.classList.toggle('admin', ADMIN); $('#admin-toggle').checked = ADMIN; lsSet(LS_ADMIN, ADMIN ? '1' : '0');
    if (GRID) GRID.render(true);
    if (SEL) renderInspector();
    document.querySelectorAll('.v').forEach(function (i) { var id = i.getAttribute('data-id'); i.disabled = !ADMIN || (id !== null && !!LOCKED[+id]); });
    if (ROUTE.name === 'g' || (ROUTE.name === 's' && ROUTE.b !== 'cedvel')) route();
  }

  // ------------------------------------------------------------------ boot
  function boot() {
    main = $('#main'); inspector = $('#inspector'); sidebar = $('#sidebar');
    var t0 = performance.now();
    window.MakroStore.boot(function (model, info) {
      try { boot2(model, info, t0); }
      catch (e) { $('#boot-msg').textContent = 'Model yüklənmədi: ' + e.message; console.error(e); }
    });
  }
  function boot2(model, info, t0) {
    M = model; INFO = info;
    CORE = window.MODEL_CORE; META = window.MODEL_META || META; BOOKS = CORE.books; SHEETS = CORE.sheets; STY = CORE.styles;
    buildGroups();
    buildIndexes();
    M.dependents();
    try { var saved = JSON.parse(lsGet(LS_EDITS) || 'null'); if (saved) applySerialized(saved, true); } catch (e) { /* ignore corrupt storage */ }
    updateChangeCount();
    $('#brand-sub').textContent = 'CAEM · proqnoz ' + INFO.firstYear + '–' + INFO.endYear;
    ADMIN = lsGet(LS_ADMIN) === '1'; document.body.classList.toggle('admin', ADMIN); $('#admin-toggle').checked = ADMIN;
    $('#admin-toggle').addEventListener('change', function () { setAdmin(this.checked); toast(ADMIN ? 'Admin rejimi aktivdir — giriş, açarlar və əmsallar dəyişdirilə bilər' : 'Baxış rejimi'); });
    $('#nav-toggle').addEventListener('click', function () { document.body.classList.toggle('nav-open'); });
    sidebar.addEventListener('click', function (e) {
      var t = e.target.closest('[data-toggle]');
      if (t) { e.preventDefault(); var gi = +t.getAttribute('data-toggle'); OPEN[gi] = !OPEN[gi]; renderNav(); }
    });
    main.addEventListener('click', function (e) {
      var a = e.target.closest('.ref[data-gs]');
      if (a) { e.preventDefault(); gotoCell(+a.getAttribute('data-gs'), +a.getAttribute('data-r'), +a.getAttribute('data-c')); }
    });
    var s = $('#search'), st;
    s.addEventListener('input', function () { clearTimeout(st); st = setTimeout(function () { doSearch(s.value); }, 120); });
    s.addEventListener('keydown', function (e) { if (e.key === 'Enter') { var f = $('#search-results .sr-item'); if (f) { location.hash = f.getAttribute('href'); $('#search-results').hidden = true; s.blur(); } } if (e.key === 'Escape') { $('#search-results').hidden = true; s.blur(); } });
    document.addEventListener('click', function (e) { if (!e.target.closest('.search')) $('#search-results').hidden = true; else if (e.target.closest('.sr-item')) $('#search-results').hidden = true; });
    window.addEventListener('hashchange', function () {
      // same-sheet cell selection updates do not rebuild the page
      var m = /^#\/s\/(\d+)\/cedvel\/([A-Z]+\d+)$/.exec(location.hash);
      if (m && GRID && GRID.gs === +m[1]) { var rc = parseA1(m[2]); GRID.select(rc.r, rc.c, true); return; }
      route();
    });
    route();
    $('#boot').remove();
    if (window.CaemUpdate) window.CaemUpdate.pendingToast();
    console.log('Model hazırdır:', Math.round(performance.now() - t0), 'ms');
  }
  window.addEventListener('DOMContentLoaded', function () {
    try { boot(); }
    catch (e) { $('#boot-msg').textContent = 'Model yüklənmədi: ' + e.message; console.error(e); }
  });
})();
