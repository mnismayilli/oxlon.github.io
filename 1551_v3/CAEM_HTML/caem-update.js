/* CAEM — model update UI (shared by index.html and panel.html):
   1) realised data for the current first forecast year + roll the first forecast year forward (CAEM's own INPUT!B9 mechanism),
   2) uploading a new CAEM workbook (.xlsb / .xlsm / .xlsx). */
(function (root) {
  'use strict';
  var Store = root.MakroStore, B = root.MakroBuilder, X = root.X;
  var CSS = '.mu{--mu-line:var(--line,#DCE3E9);--mu-line2:var(--line-2,#EDF1F4);--mu-bg2:var(--surface-2,#F8FAFB);--mu-ink2:var(--ink-2,#455463);--mu-muted:var(--muted,#738190);--mu-acc:var(--accent,#3B4F9A);--mu-ok:#2E7D32;--mu-oks:#E7F4E8;--mu-warn:#8A5300;--mu-warns:#FFF2D9;--mu-err:#B3261E;--mu-errs:#FBE9E7;font-size:14px}' +
    '.mu h3{font-size:16px;margin:0 0 6px}.mu p{margin:0 0 8px;color:var(--mu-ink2)}.mu .mu-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:14px}' +
    '.mu .mu-card{background:var(--surface,#fff);border:1px solid var(--mu-line);border-radius:10px;padding:16px 18px;display:flex;flex-direction:column;gap:8px}' +
    '.mu .mu-kv{display:grid;grid-template-columns:auto 1fr;gap:4px 14px;font-size:13.5px;margin:0}.mu .mu-kv dt{color:var(--mu-muted)}.mu .mu-kv dd{margin:0;font-weight:600}' +
    '.mu .mu-row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}.mu .mu-btn{display:inline-flex;align-items:center;gap:6px;height:34px;padding:0 14px;border-radius:8px;border:1px solid var(--mu-line);background:var(--surface,#fff);font-weight:600;cursor:pointer;font-size:13.5px;color:inherit;white-space:nowrap;font-family:inherit}' +
    '.mu .mu-btn:hover{border-color:var(--mu-acc);color:var(--mu-acc)}.mu .mu-btn.pri{background:var(--mu-acc);border-color:var(--mu-acc);color:#fff}.mu .mu-btn.pri:hover{filter:brightness(.92);color:#fff}.mu .mu-btn.dang:hover{border-color:var(--mu-err);color:var(--mu-err)}.mu .mu-btn:disabled{opacity:.5;cursor:not-allowed}' +
    '.mu .mu-btn.sm{height:28px;padding:0 10px;font-size:12.5px}.mu .mu-note{border-radius:8px;padding:10px 12px;font-size:13px;background:var(--mu-bg2);border:1px solid var(--mu-line2)}.mu .mu-note.warn{background:var(--mu-warns);border-color:#F1D9A6}.mu .mu-note.ok{background:var(--mu-oks);border-color:#BFE3C4}.mu .mu-note.err{background:var(--mu-errs);border-color:#F2C1BC}' +
    '.mu .mu-work{margin-top:18px}.mu .mu-tabs{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0}.mu .mu-tab{border:1px solid var(--mu-line);background:var(--surface,#fff);border-radius:999px;padding:5px 12px;font-weight:700;font-size:13px;cursor:pointer;color:var(--mu-ink2);font-family:inherit}.mu .mu-tab.on{background:var(--mu-acc);border-color:var(--mu-acc);color:#fff}' +
    '.mu input[type=search],.mu select{height:32px;border:1px solid var(--mu-line);border-radius:7px;padding:0 9px;background:var(--surface,#fff);font-family:inherit}.mu input.mu-v{width:120px;height:28px;border:1px solid var(--mu-line);border-radius:6px;padding:0 7px;text-align:right;font-variant-numeric:tabular-nums;font-family:inherit}' +
    '.mu input.mu-v.set{background:#FFF3BF;border-color:#E5C24F;font-weight:700}.mu input.mu-v.saved{background:#E6EFF8;border-color:#B7CFEA;font-weight:700}' +
    '.mu table.mu-t{border-collapse:collapse;width:100%;font-size:13px}.mu .mu-t th,.mu .mu-t td{padding:6px 8px;border-bottom:1px solid var(--mu-line2);text-align:left;vertical-align:middle}.mu .mu-t th{font-size:11.5px;color:var(--mu-muted);background:var(--mu-bg2);position:sticky;top:0;z-index:1}.mu .mu-t td.n,.mu .mu-t th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}' +
    '.mu .mu-wrap{border:1px solid var(--mu-line);border-radius:10px;background:var(--surface,#fff);max-height:62vh;overflow:auto}.mu details.mu-g{border-top:1px solid var(--mu-line2)}.mu details.mu-g>summary{padding:8px 12px;cursor:pointer;font-weight:700;list-style:none}.mu details.mu-g>summary::-webkit-details-marker{display:none}.mu details.mu-g>summary::before{content:"▸ ";color:var(--mu-muted)}.mu details.mu-g[open]>summary::before{content:"▾ "}' +
    '.mu .mu-src{font:11px "SF Mono",Menlo,Consolas,monospace;color:var(--mu-muted)}.mu .mu-log{font:12px "SF Mono",Menlo,Consolas,monospace;background:#0F1720;color:#D6E2EC;border-radius:8px;padding:10px 12px;max-height:220px;overflow:auto;white-space:pre-wrap}' +
    '.mu .mu-stats{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px}.mu .mu-stat{border:1px solid var(--mu-line);border-radius:8px;padding:10px 12px;background:var(--surface,#fff)}.mu .mu-stat b{display:block;font-size:20px;font-variant-numeric:tabular-nums}.mu .mu-stat span{font-size:12px;color:var(--mu-muted)}' +
    '.mu .mu-chip{display:inline-block;padding:1px 8px;border-radius:999px;font-size:11.5px;font-weight:700;background:var(--mu-line2)}.mu .mu-chip.ok{background:var(--mu-oks);color:var(--mu-ok)}.mu .mu-chip.warn{background:var(--mu-warns);color:var(--mu-warn)}.mu .mu-chip.err{background:var(--mu-errs);color:var(--mu-err)}' +
    '.mu .mu-steps{margin:0;padding-left:20px;color:var(--mu-ink2);font-size:13px}.mu .mu-steps li{margin:3px 0}';
  function css() { if (document.getElementById('mu-css')) return; var s = document.createElement('style'); s.id = 'mu-css'; s.textContent = CSS; document.head.appendChild(s); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function fold(s) { return String(s).toLocaleLowerCase('az').replace(/i̇/g, 'i').replace(/[ıİ]/g, 'i').replace(/ə/g, 'e').replace(/ö/g, 'o').replace(/ü/g, 'u').replace(/ğ/g, 'g').replace(/ş/g, 's').replace(/ç/g, 'c'); }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function nf(v, dec) {
    if (v instanceof X.XErr) return v.e; if (!isNum(v)) return v == null ? '' : String(v);
    if (dec === undefined) { var a = Math.abs(v); dec = a >= 10000 ? 0 : a >= 100 ? 1 : a >= 1 ? 2 : 4; }
    var s = Math.abs(v).toFixed(dec), p = s.split('.'); p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return (v < 0 && Number(s) !== 0 ? '−' : '') + p[0] + (p[1] ? ',' + p[1] : '');
  }
  function raw(v) { return isNum(v) ? String(+v.toPrecision(15)).replace('.', ',') : ''; }
  function parseNum(s) {
    s = String(s).trim().replace(/[\s ]/g, '').replace('−', '-');
    if (s === '') return null;
    if (/^[+-]?\d*,\d+$/.test(s)) s = s.replace(',', '.');
    else if (/^[+-]?\d{1,3}(\.\d{3})+,\d+$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) return undefined;
    return parseFloat(s);
  }
  function download(name, data, type) {
    try { var blob = data instanceof Blob ? data : new Blob([data], { type: type || 'text/plain' }), a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500); return true; }
    catch (e) { return false; }
  }
  function colName(n) { return Store.colName(n); }
  function localTime(iso) { try { var d = new Date(iso); if (isNaN(d)) return ''; var z = function (n) { return (n < 10 ? '0' : '') + n; }; return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate()) + ' ' + z(d.getHours()) + ':' + z(d.getMinutes()); } catch (e) { return ''; } }
  root.CaemLocalTime = localTime;

  // ------------------------------------------------------------------ model context
  function Ctx(M, info) {
    var core = root.MODEL_CORE, self = this;
    this.M = M; this.info = info; this.core = core; this.L = Store.layout;
    this.sidx = {}; core.sheets.forEach(function (s, i) { if (core.books[s.b].kind === 'model') self.sidx[s.n] = i; });
    // year headers of the current baseline (after any roll the positional sheets show later years than the Excel file)
    var YR = {};
    this.yrs = function (gs) {
      if (YR[gs]) return YR[gs];
      var ids = []; M.smap[gs].forEach(function (id) { ids.push(id); });
      return (YR[gs] = Store.yearRowsOf(ids, M.r, M.c, M.V0));
    };
    this.colOf = function (gs, y) { var yr = self.yrs(gs); for (var i = 0; i < yr.length; i++) for (var c in yr[i].d) if (yr[i].d[c] === y) return +c; return null; };
    this.cell = function (gs, r, c) { var x = M.smap[gs].get(r * 20000 + c); return x === undefined ? -1 : x; };
    this.addr = function (id) { return core.sheets[M.sh[id]].n + '!' + colName(M.c[id]) + M.r[id]; };
    this.rowText = function (gs, r, maxC) { var t = []; for (var c = 1; c <= (maxC || 9); c++) { var x = self.cell(gs, r, c); if (x >= 0 && typeof M.V0[x] === 'string' && M.V0[x].trim()) t.push(M.V0[x].trim()); } return t; };
  }
  // series of the 'Data' sheet that need a value in year Y (a value exists in year Yprev but not in Y); each resolves to the
  // cell where the value must be typed: the calendar-year cell of the download sheet the Data formula reads, or the Data cell
  Ctx.prototype.needs = function (Y, Yprev, kind) {
    var M = this.M, core = this.core, gsD = this.sidx.Data; if (gsD === undefined) return [];
    var cY = this.colOf(gsD, Y), cP = this.colOf(gsD, Yprev); if (cY === null || cP === null) return [];
    var out = [], self = this, rows = {};
    M.smap[gsD].forEach(function (id) { rows[M.r[id]] = 1; });
    Object.keys(rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
      var a = self.cell(gsD, r, cP), b = self.cell(gsD, r, cY); if (a < 0 || !isNum(M.V[a])) return;
      if (b >= 0 && isNum(M.V[b])) return;
      var codeId = self.cell(gsD, r, 4), code = codeId >= 0 ? M.V0[codeId] : null, nameId = self.cell(gsD, r, 5), unitId = self.cell(gsD, r, 6);
      var fcell = (b >= 0 && M.ft[b] >= 0) ? b : (M.ft[a] >= 0 ? a : -1), target = -1;
      if (fcell >= 0) {
        var rr = M.fr[fcell].filter(function (x) { return x < 0; })[0];
        if (rr !== undefined && code !== null) {
          var gs2 = core.ranges[-rr - 1][0], c2 = self.colOf(gs2, Y), row = null;
          if (gs2 !== gsD && !self.L.sheets[gs2].positional) M.smap[gs2].forEach(function (id) { if (row === null && M.c[id] <= 9 && M.V0[id] === code) row = M.r[id]; });
          if (row !== null && c2 !== null) target = self.cell(gs2, row, c2);
        }
      } else if (b >= 0) target = b;
      if (target < 0) return;
      var c2p = self.colOf(gsD, Yprev - 1), a2 = c2p === null ? -1 : self.cell(gsD, r, c2p);
      out.push({ id: target, kind: kind, year: Y, code: code, label: (nameId >= 0 ? String(M.V0[nameId]).trim() : '') || self.rowText(gsD, r).join(' · '), unit: unitId >= 0 && typeof M.V0[unitId] === 'string' ? M.V0[unitId] : '', prev: M.V[a], prev2: a2 >= 0 ? M.V[a2] : null, prevYear: Yprev, dataRow: r, where: self.addr(target) });
    });
    return out;
  };
  // every indicator row of the calendar-year (download) sheets for year Y
  Ctx.prototype.allRows = function (Y) {
    var M = this.M, core = this.core, self = this, out = [];
    core.sheets.forEach(function (s, gs) {
      if (core.books[s.b].kind !== 'model' || self.L.sheets[gs].positional || !self.yrs(gs).length) return;
      var c = self.colOf(gs, Y); if (c === null) return;
      var rows = {}; M.smap[gs].forEach(function (id) { rows[M.r[id]] = 1; });
      Object.keys(rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
        var id = self.cell(gs, r, c); if (id < 0 || M.ft[id] >= 0) return;
        if (self.yrs(gs).some(function (y) { return y.r === r; })) return;
        var lab = self.rowText(gs, r, Math.min(c - 1, 9)).filter(function (t) { return !/^(Annual|Unit|Million|Billion|Thousand)$/i.test(t); }).slice(0, 3).join(' · '); if (!lab) return;
        var p = self.cell(gs, r, c - 1);
        out.push({ id: id, kind: 'all', year: Y, label: lab, prev: p >= 0 ? M.V[p] : null, prevYear: Y - 1, sheet: s.n, gs: gs, where: self.addr(id) });
      });
    });
    return out;
  };

  // ------------------------------------------------------------------ main render
  var C = null, host = null, STATE = { view: null, tab: 'need', q: '', onlySet: false, staged: new Map(), open: new Set(), upload: null, check: null };
  function render(el, M, info) {
    css(); host = el;
    if (!C || C.M !== M) { C = new Ctx(M, info); STATE.need = null; STATE.all = null; }
    var cust = info.custom, Y = info.firstYear;
    var h = '<div class="mu"><div class="mu-grid">';
    h += '<div class="mu-card"><h3>Modelin vəziyyəti</h3><dl class="mu-kv">' +
      '<dt>Mənbə</dt><dd>' + (cust ? 'Yüklənmiş Excel faylı <span class="mu-chip warn">' + esc(localTime(cust.date)) + '</span>' : 'Orijinal ' + esc((root.MODEL_META && root.MODEL_META.verify && root.MODEL_META.verify.source) || 'CAEM.xlsx') + ' (qovluqdakı <code>data</code>)') + '</dd>' +
      '<dt>Proqnozun ilk ili</dt><dd>' + Y + (info.rolls ? ' <span class="mu-chip ok">Excel-də ' + info.baseFirst + ', +' + info.rolls + ' il</span>' : '') + '</dd>' +
      '<dt>Faktiki illər</dt><dd>' + info.lastActual + '-cü ilədək</dd><dt>Proqnoz üfüqü</dt><dd>' + Y + '–' + info.endYear + '</dd>' +
      '<dt>Faktiki/baza dəyərləri</dt><dd>' + info.actuals + ' dəyər daxil edilib' + (info.badActuals ? ' <span class="mu-chip err">' + info.badActuals + ' tanınmadı</span>' : '') + '</dd></dl>' +
      (info.notes && info.notes.length ? '<div class="mu-note err">' + info.notes.map(esc).join('<br>') + '</div>' : '') + '</div>';
    h += '<div class="mu-card"><h3>1. Faktiki məlumat və yeni il</h3><p>İl başa çatanda: (a) ' + Y + '-ci ilin faktiki (realizə olunmuş) rəqəmlərini ilkin məlumat vərəqlərinə daxil edin; (b) proqnoz üfüqünə əlavə olunacaq ' + (info.endYear + 1) + '-ci il üçün xarici mühit fərziyyələrini verin; (c) proqnozun ilk ilini ' + (Y + 1) + '-ə çəkin. CAEM-in öz qaydası ilə bütün vərəqlər yeni illərə uyğunlaşır.</p>' +
      '<div class="mu-row"><button class="mu-btn pri" id="mu-act">' + Y + ' faktiki məlumatını daxil et</button>' + (info.rolls ? '<button class="mu-btn dang" id="mu-back">Proqnozun ilk ilini ' + (Y - 1) + '-ə qaytar</button>' : '') + '</div>' +
      '<p class="mu-src" style="font-family:inherit">İrəli çəkmədən əvvəl sistem sınaq hesablaması aparır və yeni yaranacaq xətaları səbəbləri ilə göstərir.</p></div>';
    h += '<div class="mu-card"><h3>2. Yeni Excel faylı yüklə</h3><p>Nazirliyin yenilənmiş CAEM iş kitabını seçin (.xlsb; həmçinin .xlsm/.xlsx). Sistem bütün düsturları, adları və massiv düsturlarını oxuyur, Excel-in saxladığı nəticələrlə yoxlayır və yalnız sonra aktivləşdirməyə imkan verir.</p>' +
      '<div class="mu-row"><label class="mu-btn pri" for="mu-files">Faylı seç<input type="file" id="mu-files" accept=".xlsb,.xlsm,.xlsx" hidden></label>' + (cust ? '<button class="mu-btn dang" id="mu-orig">Orijinal modelə qayıt</button>' : '') + '</div>' +
      '<p class="mu-src" style="font-family:inherit">Fayl kompüterdən çıxmır — hər şey brauzerdə emal olunur.</p></div>';
    h += '</div><div class="mu-work" id="mu-work"></div></div>';
    el.innerHTML = h;
    q('#mu-act').onclick = function () { STATE.view = 'act'; STATE.staged = new Map(); STATE.check = null; drawWork(); q('#mu-work').scrollIntoView({ behavior: 'smooth' }); };
    var back = q('#mu-back'); if (back) back.onclick = function () { if (!confirm('Proqnozun ilk ili ' + (Y - 1) + '-ə qaytarılsın? Daxil edilmiş faktiki dəyərlər saxlanılır.')) return; Store.roll(-1); reload('Proqnozun ilk ili ' + (Y - 1) + '-ə qaytarıldı'); };
    q('#mu-files').onchange = function () { startUpload(Array.prototype.slice.call(this.files)); this.value = ''; };
    var og = q('#mu-orig'); if (og) og.onclick = function () {
      if (!confirm('Yüklənmiş model silinsin və orijinal CAEM faylına qayıdılsın? Ssenari dəyişiklikləri, faktiki məlumat və il sürüşdürmələri də sıfırlanacaq.')) return;
      Store.clearCustom().then(function () { Store.clearUserState(); reload('Orijinal model bərpa edildi'); });
    };
    drawWork();
  }
  function q(s) { return host.querySelector(s); }
  function reload(msg) { try { sessionStorage.setItem(Store.keys.toast, msg || ''); } catch (e) { /* ignore */ } location.reload(); }
  function drawWork() {
    var w = q('#mu-work'); if (!w) return;
    if (STATE.view === 'act') return drawActuals(w);
    if (STATE.view === 'upload') return drawUpload(w);
    w.innerHTML = '';
  }

  // ------------------------------------------------------------------ realised data + roll forward
  function lists() {
    var Y = C.info.firstYear;
    if (!STATE.need) STATE.need = C.needs(Y, Y - 1, 'actual').concat(C.needs(Y + 5, Y + 4, 'horizon'));
    return STATE.need;
  }
  function savedMap() { var act = Store.actuals(), m = {}; Object.keys(act).forEach(function (k) { var id = Store.idOfBaseKey(C.M, k); if (id >= 0) m[id] = act[k]; }); return m; }
  function drawActuals(w) {
    var info = C.info, Y = info.firstYear;
    var h = '<div class="mu-card"><div class="mu-row" style="justify-content:space-between"><h3 style="margin:0">Faktiki məlumat ' + Y + ' və ' + (info.endYear + 1) + ' fərziyyələri</h3><button class="mu-btn sm" id="mu-close">Bağla</button></div>' +
      '<ol class="mu-steps"><li><b>Tələb olunanlar</b>: «Data» vərəqinin ' + (Y - 1) + '-də dəyəri olan, lakin ' + Y + '-də olmayan bütün sıraları (faktiki məlumat) və ' + (info.endYear) + '-də olan, lakin ' + (info.endYear + 1) + '-də olmayan xarici mühit fərziyyələri. Hər sıra dəyərin yazılacağı ilkin məlumat xanasına bağlanıb.</li>' +
      '<li>Çox rəqəm üçün: <b>CSV şablonunu</b> yükləyin, Excel-də «Dəyər» sütununu doldurun və <b>CSV idxal</b> edin.</li>' +
      '<li><b>Yoxla</b> — sistem ' + (Y + 1) + '-ə irəli çəkilmiş modeli sınaq olaraq hesablayır və yeni yaranacaq xətaları göstərir. <b>Saxla</b> — dəyərlər baza modelə yazılır. <b>Saxla və ' + (Y + 1) + '-ə irəli çək</b> — həm də proqnozun ilk ili dəyişir.</li></ol>' +
      '<div class="mu-row"><div class="mu-tabs" style="margin:0"><button class="mu-tab' + (STATE.tab === 'need' ? ' on' : '') + '" data-t="need">Tələb olunanlar</button><button class="mu-tab' + (STATE.tab === 'all' ? ' on' : '') + '" data-t="all">Bütün ilkin məlumat sətirləri (' + Y + ')</button></div>' +
      '<input type="search" id="mu-q" placeholder="Axtar (məs. ÜDM, CPI, neft)" value="' + esc(STATE.q) + '" style="min-width:220px"><label style="font-size:13px"><input type="checkbox" id="mu-only"' + (STATE.onlySet ? ' checked' : '') + '> yalnız daxil edilənlər</label></div>' +
      '<div id="mu-list"></div>' +
      '<div class="mu-row" style="margin-top:6px"><button class="mu-btn" id="mu-tpl">CSV şablonu</button><label class="mu-btn" for="mu-csv">CSV idxal et<input type="file" id="mu-csv" accept=".csv,.txt" hidden></label><button class="mu-btn" id="mu-fill" title="Faktiki rəqəmlər hələ yoxdursa: səviyyə göstəriciləri son artım tempi ilə, faiz və nisbətlər son dəyərlə müvəqqəti doldurulur">Boşları müvəqqəti doldur</button></div>' +
      '<div id="mu-check"></div>' +
      '<div class="mu-row" style="margin-top:6px;padding-top:10px;border-top:1px solid var(--mu-line2)"><span class="mu-src" id="mu-cnt"></span><span style="margin-left:auto"></span><button class="mu-btn" id="mu-test">Yoxla</button><button class="mu-btn" id="mu-save">Saxla</button><button class="mu-btn pri" id="mu-roll">Saxla və ' + (Y + 1) + '-ə irəli çək</button></div></div>';
    w.innerHTML = h;
    w.querySelector('#mu-close').onclick = function () { STATE.view = null; w.innerHTML = ''; };
    Array.prototype.forEach.call(w.querySelectorAll('[data-t]'), function (b) { b.onclick = function () { STATE.tab = b.getAttribute('data-t'); drawActuals(w); }; });
    var qi = w.querySelector('#mu-q'); qi.oninput = function () { STATE.q = qi.value; clearTimeout(drawActuals.t); drawActuals.t = setTimeout(drawList, 200); };
    w.querySelector('#mu-only').onchange = function () { STATE.onlySet = this.checked; drawList(); };
    w.querySelector('#mu-tpl').onclick = function () { csvTemplate(STATE.tab === 'all' ? C.allRows(Y) : lists()); };
    w.querySelector('#mu-csv').onchange = function () { var f = this.files[0]; if (!f) return; var rd = new FileReader(); rd.onload = function () { importCSV(String(rd.result)); drawList(); }; rd.readAsText(f, 'utf-8'); this.value = ''; };
    w.querySelector('#mu-fill').onclick = function () {
      var sv = savedMap(), n = 0;
      lists().forEach(function (e) { if (sv[e.id] === undefined && !STATE.staged.has(e.id) && isNum(e.prev)) { STATE.staged.set(e.id, placeholder(e)); n++; } });
      toastMsg(n ? n + ' boş sıra müvəqqəti dəyərlə dolduruldu (səviyyələr — son artım tempi ilə, faizlər və nisbətlər — son dəyərlə; sarı) — faktiki rəqəmlərlə əvəz edin' : 'Boş sıra yoxdur'); drawList();
    };
    w.querySelector('#mu-test').onclick = function () { runCheck(false); };
    w.querySelector('#mu-save').onclick = function () { saveStaged(false); };
    w.querySelector('#mu-roll').onclick = function () { runCheck(true); };
    w.onchange = function (e) {
      var t = e.target; if (!t.classList || !t.classList.contains('mu-v')) return;
      var id = +t.getAttribute('data-id'), v = parseNum(t.value);
      if (v === undefined) { toastMsg('Rəqəm daxil edin (məs. 3,1)'); t.value = ''; return; }
      STATE.staged.set(id, v); STATE.check = null; t.classList.toggle('set', v !== null); updCount();
    };
    drawList(); drawCheck();
  }
  // provisional value: level series continue their last growth rate, percentages / rates / ratios stay flat
  var RATE_RE = /%|percent|rate|ratio|share|growth|change|index|deflator|faiz|nisbət|pay|artım|dərəcə/i;
  function placeholder(e) {
    if (isNum(e.prev2) && e.prev2 !== 0 && (e.prev > 0) === (e.prev2 > 0) && !RATE_RE.test(e.label + ' ' + (e.unit || ''))) {
      var g = e.prev / e.prev2; if (g > 0.5 && g < 1.5) return e.prev * g;
    }
    return e.prev;
  }
  function updCount() {
    var el = host.querySelector('#mu-cnt'); if (!el) return;
    var sv = savedMap(), L = lists(), filled = L.filter(function (e) { return (STATE.staged.has(e.id) ? STATE.staged.get(e.id) : sv[e.id]) !== undefined && (STATE.staged.has(e.id) ? STATE.staged.get(e.id) : sv[e.id]) !== null; }).length;
    var a = 0; STATE.staged.forEach(function (v) { if (v !== null) a++; });
    el.textContent = 'Tələb olunanlar: ' + filled + ' / ' + L.length + ' doldurulub' + (a ? ' · saxlanmamış: ' + a : '');
  }
  function drawList() {
    var el = host.querySelector('#mu-list'); if (!el) return;
    var M = C.M, sv = savedMap(), qq = fold(STATE.q.trim()), Y = C.info.firstYear;
    var list = STATE.tab === 'all' ? (STATE.all || (STATE.all = C.allRows(Y))) : lists();
    list = list.filter(function (e) { return (!qq || fold(e.label + ' ' + e.where + ' ' + (e.code || '')).indexOf(qq) >= 0) && (!STATE.onlySet || sv[e.id] !== undefined || STATE.staged.has(e.id)); });
    function row(e) {
      var s = sv[e.id], st = STATE.staged.has(e.id) ? STATE.staged.get(e.id) : undefined, val = st !== undefined ? st : s;
      return '<tr><td>' + esc(e.label) + (e.unit ? ' <span class="mu-src" style="font-family:inherit">' + esc(e.unit) + '</span>' : '') + '<div class="mu-src">' + esc(e.where) + (e.code ? ' · kod ' + esc(e.code) : '') + '</div></td><td class="n">' + e.year + '</td><td class="n">' + nf(e.prev) + ' <span class="mu-src">(' + e.prevYear + ')</span></td>' +
        '<td class="n"><input class="mu-v' + (st !== undefined && st !== null ? ' set' : s !== undefined && st === undefined ? ' saved' : '') + '" data-id="' + e.id + '" value="' + esc(val === null || val === undefined ? '' : raw(val)) + '" inputmode="decimal" aria-label="Dəyər"></td></tr>';
    }
    var head = '<table class="mu-t"><thead><tr><th>Göstərici</th><th class="n">İl</th><th class="n">Əvvəlki dəyər</th><th class="n">Dəyər</th></tr></thead><tbody>';
    if (!list.length) { el.innerHTML = '<div class="mu-note">Uyğun sətir tapılmadı.</div>'; updCount(); return; }
    if (STATE.tab === 'need') {
      var a = list.filter(function (e) { return e.kind === 'actual'; }), b = list.filter(function (e) { return e.kind === 'horizon'; });
      el.innerHTML = '<div class="mu-wrap">' + (a.length ? '<details class="mu-g" open><summary>Faktiki məlumat ' + Y + ' <span class="mu-chip">' + a.length + '</span></summary>' + head + a.map(row).join('') + '</tbody></table></details>' : '') +
        (b.length ? '<details class="mu-g" open><summary>Yeni üfüq ili ' + (Y + 5) + ': ekzogen fərziyyələr <span class="mu-chip">' + b.length + '</span></summary>' + head + b.map(row).join('') + '</tbody></table></details>' : '') + '</div>';
    } else {
      var groups = [], gi = {};
      list.forEach(function (e) { if (gi[e.gs] === undefined) { gi[e.gs] = groups.length; groups.push({ gs: e.gs, sheet: e.sheet, rows: [] }); } groups[gi[e.gs]].rows.push(e); });
      el.innerHTML = '<div class="mu-wrap">' + groups.map(function (g) { var open = STATE.open.has(g.gs) || qq; return '<details class="mu-g" data-g="' + g.gs + '"' + (open ? ' open' : '') + '><summary>' + esc(g.sheet) + ' <span class="mu-chip">' + g.rows.length + '</span></summary><div class="mu-gb">' + (open ? head + g.rows.slice(0, 400).map(row).join('') + '</tbody></table>' : '') + '</div></details>'; }).join('') + '</div>';
      Array.prototype.forEach.call(el.querySelectorAll('details.mu-g'), function (d) {
        d.addEventListener('toggle', function () { var gs = +d.getAttribute('data-g'); if (d.open) { STATE.open.add(gs); var bx = d.querySelector('.mu-gb'); if (!bx.innerHTML) bx.innerHTML = head + groups[gi[gs]].rows.map(row).join('') + '</tbody></table>'; } else STATE.open.delete(gs); });
      });
    }
    updCount();
  }
  function stagedActuals() { var o = {}; STATE.staged.forEach(function (v, id) { o[Store.baseKey(C.M, id)] = v; }); return o; }
  // dry run: stored + staged values, first forecast year moved by one; report new errors and their root causes
  function runCheck(thenRoll) {
    var el = host.querySelector('#mu-check'); el.innerHTML = '<div class="mu-note">Sınaq hesablaması aparılır…</div>';
    setTimeout(function () {
      var extra = stagedActuals(), del = []; Object.keys(extra).forEach(function (k) { if (extra[k] === null) { delete extra[k]; del.push(k); } });
      var res;
      try { var saved = Store.actuals(); del.forEach(function (k) { delete saved[k]; }); res = Store.trial(Object.assign({}, extra), C.info.rolls + 1, saved); }
      catch (e) { el.innerHTML = '<div class="mu-note err">Sınaq alınmadı: ' + esc(e.message) + '</div>'; return; }
      STATE.check = diagnose(res);
      drawCheck();
      if (thenRoll) {
        if (STATE.check.roots.length && !confirm('Sınaq ' + STATE.check.total + ' yeni xəta göstərir (' + STATE.check.roots.length + ' əsas səbəb). Buna baxmayaraq ' + (C.info.firstYear + 1) + '-ə irəli çəkilsin?')) return;
        saveStaged(true);
      }
    }, 30);
  }
  function diagnose(res) {
    var T = res.model, core = root.MODEL_CORE, errs = res.newErrors, isErr = function (id) { return T.V[id] instanceof X.XErr; };
    var roots = errs.filter(function (id) { return !T.precedents(id).some(isErr); });
    var arr = -1; core.sheets.forEach(function (s, i) { if (s.n === '__massivlər__') arr = i; });
    var out = [], fixes = [];
    roots.slice(0, 60).forEach(function (id) {
      var real = id, note = '';
      if (T.sh[id] === arr) { var g = core.agroups[T.r[id] - 1]; real = T.smap[g[0]].get(g[1] * 20000 + g[2]); note = 'massiv düsturu'; }
      var blanks = Array.from(new Set(T.precedents(id))).filter(function (p) { var v = T.V[p]; return v === null || v === '' || v === ' ' || (typeof v === 'string' && !v.trim()); }).slice(0, 3);
      out.push({ id: real, err: T.V[id].e, note: note, blanks: blanks.map(function (p) { return core.sheets[T.sh[p]].n + '!' + colName(T.c[p]) + T.r[p]; }) });
      // estimation-sample start years ("Start") that now fall on a year without data
      if (note) T.precedents(id).forEach(function (p) {
        if (T.ft[p] >= 0 || !isNum(T.V[p]) || T.V[p] < 1980 || T.V[p] > 2060 || T.V[p] !== Math.floor(T.V[p])) return;
        var lab = T.smap[T.sh[p]].get(T.r[p] * 20000 + T.c[p] - 1); if (lab === undefined || !/start/i.test(String(T.V[lab]))) return;
        if (!fixes.some(function (f) { return f.id === p; })) fixes.push({ id: p, from: T.V[p], to: T.V[p] + 1, where: core.sheets[T.sh[p]].n + '!' + colName(T.c[p]) + T.r[p] });
      });
    });
    return { total: errs.length, roots: out, fixes: fixes, actuals: res.actuals };
  }
  function drawCheck() {
    var el = host.querySelector('#mu-check'); if (!el) return;
    var ck = STATE.check; if (!ck) { el.innerHTML = ''; return; }
    if (!ck.total) { el.innerHTML = '<div class="mu-note ok"><b>Sınaq uğurludur.</b> ' + (C.info.firstYear + 1) + '-ə irəli çəkilmiş modeldə yeni xəta yaranmır (' + ck.actuals + ' faktiki/baza dəyəri tətbiq olundu).</div>'; return; }
    var M = C.M;
    var h = '<div class="mu-note warn"><b>Sınaq: ' + ck.total + ' yeni xəta, ' + ck.roots.length + ' əsas səbəb.</b> Adətən səbəb boş qalmış faktiki dəyər və ya fərziyyədir — aşağıdakı xanaların mənbələrini doldurun.</div>';
    if (ck.fixes.length) h += '<div class="mu-note" style="margin-top:6px"><b>Təklif:</b> qiymətləndirmə (OLS) nümunəsinin başlanğıc ili artıq məlumatı olmayan ilə düşür — ' + ck.fixes.map(function (f) { return '<code>' + esc(f.where) + '</code> ' + f.from + ' → ' + f.to; }).join(', ') + ' <button class="mu-btn sm" id="mu-fix">Tətbiq et</button></div>';
    h += '<div class="mu-wrap" style="max-height:260px;margin-top:6px"><table class="mu-t"><thead><tr><th>Xana (əsas səbəb)</th><th>Xəta</th><th>Boş mənbələr</th></tr></thead><tbody>' + ck.roots.map(function (r) {
      return '<tr><td class="mu-src">' + esc(C.addr(r.id)) + (r.note ? ' (' + r.note + ')' : '') + '</td><td>' + esc(r.err) + '</td><td class="mu-src">' + esc(r.blanks.join(', ')) + '</td></tr>';
    }).join('') + '</tbody></table></div>';
    el.innerHTML = h;
    var fx = el.querySelector('#mu-fix'); if (fx) fx.onclick = function () {
      var act = Store.actuals(); ck.fixes.forEach(function (f) { act[root.MODEL_CORE.books[0].n + '|' + f.where.replace('!', '|')] = f.to; });
      Store.saveActuals(act); toastMsg('Başlanğıc il(lər) yeniləndi — yenidən yoxlanılır'); runCheck(false);
    };
  }
  function saveStaged(roll) {
    var act = Store.actuals(), n = 0;
    STATE.staged.forEach(function (v, id) { var k = Store.baseKey(C.M, id); if (v === null) { if (act[k] !== undefined) { delete act[k]; n++; } } else { act[k] = v; n++; } });
    if (!n && !roll) { toastMsg('Saxlanılacaq dəyişiklik yoxdur'); return; }
    if (!Store.saveActuals(act)) { alert('Brauzer yaddaşına yazmaq mümkün olmadı.'); return; }
    if (roll) { var y = Store.roll(1); reload(n + ' dəyər saxlanıldı · proqnozun ilk ili: ' + y); }
    else reload(n + ' faktiki/baza dəyəri saxlanıldı və model yenidən hesablandı');
  }
  function csvTemplate(list) {
    var M = C.M, lines = ['Vərəq;Xana;Kod;Göstərici;İl;Əvvəlki dəyər;Dəyər'];
    list.forEach(function (e) { var s = C.core.sheets[M.sh[e.id]]; lines.push([s.n, colName(M.c[e.id]) + M.r[e.id], e.code || '', e.label, e.year, raw(e.prev), ''].map(function (x) { x = String(x); return /[;"\n]/.test(x) ? '"' + x.replace(/"/g, '""') + '"' : x; }).join(';')); });
    download('caem-faktiki-' + C.info.firstYear + '.csv', '﻿' + lines.join('\r\n'), 'text/csv');
  }
  function parseCSV(text) {
    text = text.replace(/^﻿/, '');
    var first = text.split(/\r?\n/)[0], sep = [';', '\t', ','].map(function (s) { return [s, first.split(s).length]; }).sort(function (a, b) { return b[1] - a[1]; })[0][0];
    var rows = [], row = [], cur = '', inQ = false;
    for (var i = 0; i < text.length; i++) {
      var ch = text.charAt(i);
      if (inQ) { if (ch === '"') { if (text.charAt(i + 1) === '"') { cur += '"'; i++; } else inQ = false; } else cur += ch; continue; }
      if (ch === '"') inQ = true; else if (ch === sep) { row.push(cur); cur = ''; } else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text.charAt(i + 1) === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; } else cur += ch;
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return rows;
  }
  function importCSV(text) {
    var rows = parseCSV(text); if (rows.length < 2) { toastMsg('CSV faylında məlumat yoxdur'); return; }
    var h = rows[0].map(function (x) { return fold(x.trim()); });
    function col(names) { for (var i = 0; i < h.length; i++) for (var j = 0; j < names.length; j++) if (h[i] === names[j]) return i; return -1; }
    var cs = col(['vereq', 'sheet']), cx = col(['xana', 'cell']), cv = col(['deyer', 'faktiki deyer', 'value']);
    if (cs < 0 || cx < 0 || cv < 0) { toastMsg('Başlıqlar tanınmadı: «Vərəq; Xana; Dəyər» sütunları lazımdır'); return; }
    var ok = 0, bad = 0, badList = [];
    rows.slice(1).forEach(function (r, i) {
      if (!r.length || r.every(function (x) { return !String(x).trim(); })) return;
      var sval = (r[cv] || '').trim(); if (sval === '') return;
      var v = parseNum(sval), gs = C.sidx[(r[cs] || '').trim()], m = /^([A-Z]+)(\d+)$/.exec(String(r[cx] || '').trim().toUpperCase()), id = -1;
      if (gs !== undefined && m) id = C.cell(gs, +m[2], Store.colNum(m[1]));
      if (id < 0 || C.M.ft[id] >= 0 || v === undefined || v === null) { bad++; if (badList.length < 5) badList.push('sətir ' + (i + 2)); return; }
      STATE.staged.set(id, v); ok++;
    });
    STATE.onlySet = true; STATE.check = null; var only = host.querySelector('#mu-only'); if (only) only.checked = true;
    toastMsg('İdxal: ' + ok + ' dəyər qəbul edildi' + (bad ? ', ' + bad + ' sətir tanınmadı (' + badList.join(', ') + ')' : '') + ' — yoxlayıb «Saxla» basın');
  }

  // ------------------------------------------------------------------ upload a new workbook
  function startUpload(files) {
    if (!files.length) return;
    STATE.view = 'upload';
    STATE.upload = { file: files[0], log: [], result: null, busy: false };
    drawWork(); q('#mu-work').scrollIntoView({ behavior: 'smooth' });
  }
  function drawUpload(w) {
    var U = STATE.upload, f = U.file, h = '<div class="mu-card"><div class="mu-row" style="justify-content:space-between"><h3 style="margin:0">Yeni CAEM faylı</h3><button class="mu-btn sm" id="mu-up-close">Bağla</button></div>';
    h += '<dl class="mu-kv"><dt>Fayl</dt><dd>' + esc(f.name) + '</dd><dt>Ölçü</dt><dd>' + (f.size / 1048576).toFixed(1) + ' MB</dd></dl>';
    h += '<div class="mu-row"><button class="mu-btn pri" id="mu-run"' + (U.busy ? ' disabled' : '') + '>Oxu, yoxla və hesabla</button></div>';
    if (U.log.length) h += '<div class="mu-log" id="mu-log">' + esc(U.log.join('\n')) + '</div>';
    if (U.result) h += reportHTML(U.result);
    h += '</div>';
    w.innerHTML = h;
    w.querySelector('#mu-up-close').onclick = function () { if (U.busy) return; STATE.view = null; STATE.upload = null; w.innerHTML = ''; };
    var run = w.querySelector('#mu-run'); if (run) run.onclick = function () { runBuild(w); };
    var act = w.querySelector('#mu-activate'); if (act) act.onclick = activate;
    var zip = w.querySelector('#mu-zip'); if (zip) zip.onclick = exportZip;
    var lg = w.querySelector('#mu-log'); if (lg) lg.scrollTop = lg.scrollHeight;
  }
  function log(msg) { var U = STATE.upload; U.log.push(msg); var el = host.querySelector('#mu-log'); if (el) { el.textContent = U.log.join('\n'); el.scrollTop = el.scrollHeight; } else drawUpload(host.querySelector('#mu-work')); }
  function runBuild(w) {
    var U = STATE.upload, f = U.file, isB = /\.xlsb$/i.test(f.name);
    if (typeof DecompressionStream === 'undefined') { alert('Bu brauzer Excel fayllarını aça bilmir (DecompressionStream yoxdur). Chrome, Edge, Safari 16.4+ və ya Firefox 113+ istifadə edin.'); return; }
    U.busy = true; U.log = []; U.result = null; drawUpload(w);
    var t0 = Date.now(), bookName = root.MODEL_CORE.books[0].n;
    log('▸ ' + f.name + ' oxunur…');
    (isB ? root.MakroXlsb.readXlsb(f, bookName, function (m) { log(m); }) : B.readXlsx(f, bookName, function (m) { log(m); })).then(function (rb) {
      var nc = 0, nf2 = 0; rb.sheets.forEach(function (s) { nc += s.cells.length; s.cells.forEach(function (c) { if (c.f) nf2++; }); });
      log('  ✓ ' + rb.sheets.length + ' vərəq, ' + nc.toLocaleString('az') + ' xana, ' + nf2.toLocaleString('az') + ' düstur, ' + (rb.names || []).length + ' ad, ' + (rb.arrays || []).length + ' massiv düsturu');
      log('▸ Düsturlar təhlil edilir…');
      return new Promise(function (r) { setTimeout(function () { r(rb); }, 30); });
    }).then(function (rb) {
      // blank input cells of the current catalogue (shock paths etc.) exist in the model so they stay editable
      var extra = []; ((root.MODEL_META && root.MODEL_META.inputs && root.MODEL_META.inputs.groups) || []).forEach(function (g) { g.items.forEach(function (it) { it.cells.forEach(function (a) { var m = /^([A-Z]+)(\d+)$/.exec(a); if (m) extra.push([it.sheet, +m[2], Store.colNum(m[1])]); }); }); });
      rb.extraCells = extra;
      var res = B.compile([rb]);
      log('  ✓ ' + res.nForm.toLocaleString('az') + ' düstur, ' + res.core.tplCode.length + ' şablon, ' + res.core.ncells.toLocaleString('az') + ' xana; qeyd: ' + res.problems.length);
      log('▸ Excel-in saxladığı nəticələrlə yoxlama…');
      return new Promise(function (r) { setTimeout(function () { r({ res: res, rb: rb }); }, 30); });
    }).then(function (o) {
      var v = B.verify(o.res);
      log('  ✓ düstur-düstur: ' + v.local.ok + ' / ' + v.formulas + ' · tam hesablama: ' + v.global.ok + ' / ' + v.formulas);
      log('Hazırdır — ' + ((Date.now() - t0) / 1000).toFixed(1) + ' san.');
      U.result = { res: o.res, v: v, rb: o.rb, files: [{ name: f.name, size: f.size, modified: f.lastModified }] };
      U.busy = false; drawUpload(w);
    }).catch(function (e) { U.busy = false; log('✗ Xəta: ' + (e && e.message ? e.message : e)); drawUpload(w); console.error(e); });
  }
  function pct(a, b) { return b ? (100 * a / b).toFixed(3).replace('.', ',') + '%' : '—'; }
  function diffList(title, d) {
    var items = d.bad.filter(Boolean); if (!d.bad.length) return '';
    return '<details class="mu-g"><summary>' + esc(title) + ' <span class="mu-chip warn">' + d.bad.length + '</span></summary><div style="padding:0 12px 10px"><table class="mu-t"><thead><tr><th>Xana</th><th class="n">Hesablanan</th><th class="n">Excel-də</th></tr></thead><tbody>' +
      items.slice(0, 100).map(function (x) { return '<tr><td class="mu-src">' + esc(x[0]) + '</td><td class="n">' + esc(x[1]) + '</td><td class="n">' + esc(x[2]) + '</td></tr>'; }).join('') + '</tbody></table></div></details>';
  }
  function reportHTML(R) {
    var v = R.v, res = R.res, good = v.local.ok / v.formulas >= 0.999 && v.global.ok / v.formulas >= 0.999;
    var probs = res.problems.filter(function (p) { return !/dövri asılılıq/.test(p.what); });
    var h = '<div class="mu-stats" style="margin-top:12px"><div class="mu-stat"><b>' + res.core.sheets.filter(function (x) { return x.n !== '__massivlər__'; }).length + '</b><span>vərəq</span></div><div class="mu-stat"><b>' + v.formulas.toLocaleString('az') + '</b><span>düstur</span></div>' +
      '<div class="mu-stat"><b>' + pct(v.local.ok, v.formulas) + '</b><span>düstur-düstur yoxlama</span></div><div class="mu-stat"><b>' + pct(v.global.ok, v.formulas) + '</b><span>tam yenidən hesablama</span></div><div class="mu-stat"><b>' + probs.length + '</b><span>hesablanmayan düstur</span></div></div>';
    h += '<div class="mu-note ' + (good && !probs.length ? 'ok' : good ? 'warn' : 'err') + '" style="margin-top:10px">' + (good
      ? '<b>Yoxlama uğurludur.</b> Sistemin hesabladığı nəticələr Excel-in saxladığı nəticələrlə üst-üstə düşür.'
      : '<b>Diqqət:</b> nəticələrin bir hissəsi Excel-in saxladığı dəyərlərdən fərqlənir. Adətən bu, faylın Excel-də yenidən hesablanmadan saxlanması deməkdir: faylı Excel-də açıb «Hesabla» (F9) edin və yenidən saxlayın.') + '</div>';
    if (probs.length) h += '<details class="mu-g" open><summary>Qeydlər <span class="mu-chip warn">' + probs.length + '</span></summary><div style="padding:0 12px 10px"><table class="mu-t"><tbody>' + probs.slice(0, 100).map(function (p) { return '<tr><td class="mu-src" style="white-space:nowrap">' + esc(p.where) + '</td><td>' + esc(p.what) + '</td></tr>'; }).join('') + '</tbody></table></div></details>';
    h += diffList('Düstur-düstur yoxlamada fərqlənən xanalar', v.local) + diffList('Tam hesablamada fərqlənən xanalar', v.global);
    h += '<div class="mu-row" style="margin-top:12px"><button class="mu-btn pri" id="mu-activate">Bu modeli aktivləşdir</button><button class="mu-btn" id="mu-zip">«data» qovluğunu ZIP kimi yüklə</button></div>' +
      '<p class="mu-src" style="font-family:inherit">«Aktivləşdir» — model bu brauzerdə yadda saxlanılır və hər iki səhifədə (klassik görünüş və iş paneli) istifadə olunur. Hamı üçün daimi etmək: ZIP-i açıb içindəki <code>data</code> qovluğunu sayt qovluğundakı <code>data</code> qovluğunun yerinə köçürün.</p>';
    return h;
  }
  function metaPatch(R) {
    var res = R.res, v = R.v, core = res.core, rb = R.rb, sl = {};
    Object.keys(res.deps).forEach(function (k) {
      var cid = +k, s = res.parts[0].sh[cid];
      res.deps[k].forEach(function (d) { var t = res.parts[0].sh[d]; if (t === s) return; var key = t + '>' + s, e = sl[key] || (sl[key] = { ss: core.sheets[t].n, ds: core.sheets[s].n, n: 0, cells: {}, formulas: {} }); e.n++; e.cells[d] = 1; e.formulas[cid] = 1; });
    });
    var links = Object.keys(sl).map(function (k) { var e = sl[k]; return { ss: e.ss, ds: e.ds, n: e.n, cells: Object.keys(e.cells).length, formulas: Object.keys(e.formulas).length }; }).sort(function (a, b) { return b.n - a.n; });
    return {
      verify: { sheets: core.sheets.filter(function (s) { return s.n !== '__massivlər__'; }).length, visible: rb.sheets.filter(function (s) { return s.state === 'visible'; }).length, cells: core.ncells, formulas: v.formulas,
        arrays: core.agroups.length, arrayCells: core.agroups.reduce(function (a, g) { return a + (g[3] - g[1] + 1) * (g[4] - g[2] + 1); }, 0), literals: core.K.length, names: (core.names || []).length, cyc: core.cyc.map(function (b) { return b[1] - b[0]; }),
        local_ok: v.local.ok, global_ok: v.global.ok, diffs: v.local.bad.filter(Boolean).slice(0, 20).map(function (d) { return d[0] + ': Excel-də ' + d[2] + ', hesablanan ' + d[1]; }) },
      sheetLinks: links,
      stale: { date: new Date().toISOString(), files: R.files.map(function (f) { return f.name; }) }
    };
  }
  function activate() {
    var R = STATE.upload.result; if (!R) return;
    if (!confirm('Yeni model aktivləşdirilsin? Mövcud ssenari dəyişiklikləri, faktiki məlumat və il sürüşdürmələri sıfırlanacaq (lazımdırsa, əvvəlcə ixrac edin).')) return;
    var rec = { core: R.res.core, parts: R.res.parts, meta: metaPatch(R), info: { date: new Date().toISOString(), files: R.files } };
    Store.putCustom(rec).then(function () { Store.clearUserState(); reload('Yeni model aktivləşdirildi: ' + R.files.map(function (f) { return f.name; }).join(', ')); })
      .catch(function (e) { alert('Brauzer yaddaşına yazmaq mümkün olmadı (' + (e && e.message) + '). «data» qovluğunu ZIP kimi yükləyib sayt qovluğuna köçürün.'); });
  }
  function exportZip() {
    var R = STATE.upload.result; if (!R) return;
    var meta = {}, base = root.MODEL_META || {}; for (var k in base) meta[k] = base[k];
    var p = metaPatch(R); for (k in p) meta[k] = p[k];
    var blob = B.zipStore(B.dataFiles(R.res.core, R.res.parts, meta));
    if (!download('caem-data-' + new Date().toISOString().slice(0, 10) + '.zip', blob)) alert('Yükləmə bu mühitdə mümkün olmadı.');
  }
  var TT = null;
  function toastMsg(m) {
    var t = document.getElementById('toast');
    if (t) { t.textContent = m; t.hidden = false; clearTimeout(TT); TT = setTimeout(function () { t.hidden = true; }, 3600); } else alert(m);
  }
  function pendingToast() { try { var m = sessionStorage.getItem(Store.keys.toast); if (m) { sessionStorage.removeItem(Store.keys.toast); setTimeout(function () { toastMsg(m); }, 300); } } catch (e) { /* ignore */ } }

  root.CaemUpdate = { render: render, pendingToast: pendingToast, Ctx: Ctx };
})(typeof window !== 'undefined' ? window : globalThis);
