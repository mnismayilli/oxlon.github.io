/* Makro model — shared admin UI for model updates (used by index.html and panel.html):
   1) realised (actual) data for a year + adding a forecast year, 2) uploading new Excel files. */
(function (root) {
  'use strict';
  var Store = root.MakroStore, B = root.MakroBuilder, X = root.X;
  var CSS = '.mu{--mu-line:var(--line,#DCE3E9);--mu-line2:var(--line-2,#EDF1F4);--mu-bg2:var(--surface-2,#F8FAFB);--mu-ink2:var(--ink-2,#455463);--mu-muted:var(--muted,#738190);--mu-acc:var(--accent,#0E6F7C);--mu-accs:var(--accent-soft,#E1F0F2);--mu-ok:#2E7D32;--mu-oks:#E7F4E8;--mu-warn:#8A5300;--mu-warns:#FFF2D9;--mu-err:#B3261E;--mu-errs:#FBE9E7;font-size:14px}' +
    '.mu h3{font-size:16px;margin:0 0 6px}.mu p{margin:0 0 8px;color:var(--mu-ink2)}.mu .mu-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:14px}' +
    '.mu .mu-card{background:var(--surface,#fff);border:1px solid var(--mu-line);border-radius:10px;padding:16px 18px;display:flex;flex-direction:column;gap:8px}' +
    '.mu .mu-kv{display:grid;grid-template-columns:auto 1fr;gap:4px 14px;font-size:13.5px}.mu .mu-kv dt{color:var(--mu-muted)}.mu .mu-kv dd{margin:0;font-weight:600}' +
    '.mu .mu-row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}.mu .mu-btn{display:inline-flex;align-items:center;gap:6px;height:34px;padding:0 14px;border-radius:8px;border:1px solid var(--mu-line);background:var(--surface,#fff);font-weight:600;cursor:pointer;font-size:13.5px;color:inherit;white-space:nowrap}' +
    '.mu .mu-btn:hover{border-color:var(--mu-acc);color:var(--mu-acc)}.mu .mu-btn.pri{background:var(--mu-acc);border-color:var(--mu-acc);color:#fff}.mu .mu-btn.pri:hover{filter:brightness(.92);color:#fff}.mu .mu-btn.dang:hover{border-color:var(--mu-err);color:var(--mu-err)}.mu .mu-btn:disabled{opacity:.5;cursor:not-allowed}' +
    '.mu .mu-btn.sm{height:28px;padding:0 10px;font-size:12.5px}.mu .mu-note{border-radius:8px;padding:10px 12px;font-size:13px;background:var(--mu-bg2);border:1px solid var(--mu-line2)}.mu .mu-note.warn{background:var(--mu-warns);border-color:#F1D9A6}.mu .mu-note.ok{background:var(--mu-oks);border-color:#BFE3C4}.mu .mu-note.err{background:var(--mu-errs);border-color:#F2C1BC}' +
    '.mu .mu-work{margin-top:18px}.mu .mu-tabs{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0}.mu .mu-tab{border:1px solid var(--mu-line);background:var(--surface,#fff);border-radius:999px;padding:5px 12px;font-weight:700;font-size:13px;cursor:pointer;color:var(--mu-ink2)}.mu .mu-tab.on{background:var(--mu-acc);border-color:var(--mu-acc);color:#fff}' +
    '.mu input[type=search],.mu select,.mu input.mu-in{height:32px;border:1px solid var(--mu-line);border-radius:7px;padding:0 9px;background:var(--surface,#fff)}.mu input.mu-v{width:120px;height:28px;border:1px solid var(--mu-line);border-radius:6px;padding:0 7px;text-align:right;font-variant-numeric:tabular-nums}' +
    '.mu input.mu-v.set{background:#FFF3BF;border-color:#E5C24F;font-weight:700}.mu input.mu-v.saved{background:#E6EFF8;border-color:#B7CFEA;font-weight:700}' +
    '.mu table.mu-t{border-collapse:collapse;width:100%;font-size:13px}.mu .mu-t th,.mu .mu-t td{padding:6px 8px;border-bottom:1px solid var(--mu-line2);text-align:left;vertical-align:middle}.mu .mu-t th{font-size:11.5px;color:var(--mu-muted);background:var(--mu-bg2);position:sticky;top:0;z-index:1}.mu .mu-t td.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}' +
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
    var s = Math.abs(v).toFixed(dec), p = s.split('.'); p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return (v < 0 && Number(s) !== 0 ? '−' : '') + p[0] + (p[1] ? ',' + p[1] : '');
  }
  function raw(v) { return isNum(v) ? String(+v.toPrecision(15)).replace('.', ',') : ''; }
  function parseNum(s) {
    s = String(s).trim().replace(/[\s ]/g, '').replace('−', '-');
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

  // ------------------------------------------------------------------ model helpers
  function Ctx(M, info) {
    var core = root.MODEL_CORE, self = this;
    this.M = M; this.info = info; this.core = core;
    this.rows = core.sheets.map(function () { return {}; });
    for (var i = 0; i < M.N; i++) { var t = this.rows[M.sh[i]]; (t[M.r[i]] || (t[M.r[i]] = [])).push(i); }
    this.rows.forEach(function (t) { for (var k in t) t[k].sort(function (a, b) { return M.c[a] - M.c[b]; }); });
    this.yr = {};
    this.book = function (gs) { return core.books[core.sheets[gs].b].n; };
    this.yearRows = function (gs) {
      if (self.yr[gs]) return self.yr[gs];
      var ids = []; for (var r in self.rows[gs]) Array.prototype.push.apply(ids, self.rows[gs][r]);
      return (self.yr[gs] = Store.yearRowsOf(ids, M.r, M.c, M.V0));
    };
    this.colFor = function (gs, r, y) { var yr = self.yearRows(gs), best = null; for (var i = 0; i < yr.length; i++) { if (yr[i].r > r && best) break; for (var c in yr[i].d) if (yr[i].d[c] === y) best = +c; } return best; };
    this.label = function (gs, r) {
      var ids = self.rows[gs][r] || [], firstNum = 1e9, parts = [];
      for (var i = 0; i < ids.length; i++) if (isNum(M.V0[ids[i]]) && M.c[ids[i]] > 4 || M.ft[ids[i]] >= 0 && isNum(M.V0[ids[i]])) { firstNum = M.c[ids[i]]; break; }
      for (i = 0; i < ids.length && parts.length < 3; i++) { var v = M.V0[ids[i]]; if (M.c[ids[i]] >= firstNum) break; if (typeof v === 'string' && v.trim() && !/^[\s\d.,\-%!_]+$/.test(v)) parts.push(v.trim()); }
      return parts.join(' · ');
    };
    this.key = function (id) { var s = core.sheets[M.sh[id]]; return core.books[s.b].n + '|' + s.n + '|' + Store.colName(M.c[id]) + M.r[id]; };
    this.addr = function (id) { var s = core.sheets[M.sh[id]]; return core.books[s.b].n.replace(/^MOE /, '') + ' › ' + s.n + '!' + Store.colName(M.c[id]) + M.r[id]; };
    this.sidx = {}; core.sheets.forEach(function (s, i) { self.sidx[core.books[s.b].n + '|' + s.n] = i; });
    this.idOf = function (b, s, a1) { var gs = self.sidx[b + '|' + s], m = /^\$?([A-Z]+)\$?(\d+)$/.exec(String(a1).trim().toUpperCase()); if (gs === undefined || !m) return -1; var x = M.smap[gs].get(+m[2] * 20000 + Store.colNum(m[1])); return x === undefined ? -1 : x; };
    this.source = function (id) { var g = 0; while (M.ft[id] >= 0 && g++ < 8) { var d = core.disp[M.ft[id]]; if (!/^\x02[^\x03]*\x03$/.test(d) || M.fr[id][0] < 0) break; id = M.fr[id][0]; } return id; };
  }

  // ------------------------------------------------------------------ main render
  var OUTPUT = { 'MOE REPORT 3 PAGES': 1, '8 vərəq': 1 };
  var C = null, host = null, STATE = { view: null, year: null, tab: 'key', q: '', onlySet: false, staged: new Map(), open: new Set(), upload: null };
  function render(el, M, info) {
    css(); host = el;
    if (!C || C.M !== M) C = new Ctx(M, info);
    if (STATE.year === null) STATE.year = info.lastActual + 1;
    var cust = info.custom, st = Store.structure();
    var yrsSel = []; for (var y = 2015; y <= info.endYear; y++) yrsSel.push(y);
    var h = '<div class="mu"><div class="mu-grid">';
    h += '<div class="mu-card"><h3>Modelin vəziyyəti</h3><dl class="mu-kv">' +
      '<dt>Mənbə</dt><dd>' + (cust ? 'Yüklənmiş Excel faylları <span class="mu-chip warn">' + esc((cust.date || '').replace('T', ' ').slice(0, 16)) + '</span>' : 'Orijinal Excel faylları (qovluqdakı <code>data</code>)') + '</dd>' +
      '<dt>Faktiki illər</dt><dd>' + info.lastActual + '-cü ilədək</dd><dt>Gözlənilən il</dt><dd>' + (info.lastActual + 1) + '</dd>' +
      '<dt>Proqnoz üfüqü</dt><dd>' + (info.lastActual + 2) + '–' + info.endYear + (info.extraYears ? ' <span class="mu-chip ok">+' + info.extraYears + ' əlavə il</span>' : '') + '</dd>' +
      '<dt>Faktiki məlumat</dt><dd>' + info.actuals + ' dəyər daxil edilib' + (info.badActuals ? ' <span class="mu-chip err">' + info.badActuals + ' tanınmadı</span>' : '') + '</dd></dl>' +
      (info.notes && info.notes.length ? '<div class="mu-note err">' + info.notes.map(esc).join('<br>') + '</div>' : '') +
      '<div class="mu-row"><label for="mu-la" class="mu-src" style="font-family:inherit">Son faktiki il</label><select id="mu-la">' + yrsSel.map(function (y2) { return '<option' + (y2 === info.lastActual ? ' selected' : '') + '>' + y2 + '</option>'; }).join('') + '</select><button class="mu-btn sm" id="mu-la-set">Təyin et</button></div></div>';
    h += '<div class="mu-card"><h3>1. Faktiki məlumat və yeni il</h3><p>İl başa çatanda: (a) həmin ilin faktiki (realizə olunmuş) rəqəmlərini daxil edin — model onları proqnozun yerinə qoyur və sonrakı illəri onlardan başlayaraq yenidən hesablayır; (b) proqnoz üfüqünə növbəti ili əlavə edin.</p>' +
      '<div class="mu-row"><button class="mu-btn pri" id="mu-act">' + (info.lastActual + 1) + ' faktiki məlumatlarını daxil et</button></div>' +
      '<div class="mu-row"><button class="mu-btn" id="mu-add">' + (info.endYear + 1) + ' ilini proqnoza əlavə et</button>' + (info.extraYears ? '<button class="mu-btn dang" id="mu-rem">' + info.endYear + ' ilini sil</button>' : '') + '</div>' +
      '<p class="mu-src" style="font-family:inherit">Yeni il Excel-dəki kimi yaradılır: hər il sütunlu vərəqdə ' + info.endYear + ' sütunundan sonra yeni sütun əlavə olunur, düsturlar sağa köçürülür (nisbi istinadlar bir il sürüşür), giriş dəyərləri əvvəlki ildən götürülür.</p></div>';
    h += '<div class="mu-card"><h3>2. Yeni Excel faylları yüklə</h3><p>Nazirliyin yenilənmiş .xlsx fayllarını seçin (bir, bir neçə və ya 11-in hamısı). Sistem bütün düsturları oxuyur, fayllar arası əlaqələri qurur, Excel-in saxladığı nəticələrlə yoxlayır və yalnız sonra aktivləşdirməyə imkan verir.</p>' +
      '<div class="mu-row"><label class="mu-btn pri" for="mu-files">Faylları seç<input type="file" id="mu-files" accept=".xlsx,.xlsm" multiple hidden></label>' + (cust ? '<button class="mu-btn dang" id="mu-orig">Orijinal modelə qayıt</button>' : '') + '</div>' +
      '<p class="mu-src" style="font-family:inherit">Fayllar kompüterdən çıxmır — hər şey brauzerdə emal olunur.</p></div>';
    h += '</div><div class="mu-work" id="mu-work"></div></div>';
    el.innerHTML = h;
    q('#mu-la-set').onclick = function () { var s2 = Store.structure(); s2.lastActual = +q('#mu-la').value; Store.saveStructure(s2); reload('Son faktiki il: ' + s2.lastActual); };
    q('#mu-act').onclick = function () { STATE.view = 'act'; STATE.year = info.lastActual + 1; STATE.staged = new Map(); drawWork(); q('#mu-work').scrollIntoView({ behavior: 'smooth' }); };
    q('#mu-add').onclick = addYear;
    var rem = q('#mu-rem'); if (rem) rem.onclick = function () { if (!confirm(info.endYear + ' ili və ona aid bütün daxil edilmiş dəyərlər silinsin?')) return; Store.removeYear(); reload(info.endYear + ' ili silindi'); };
    q('#mu-files').onchange = function () { startUpload(Array.prototype.slice.call(this.files)); this.value = ''; };
    var og = q('#mu-orig'); if (og) og.onclick = function () {
      if (!confirm('Yüklənmiş model silinsin və orijinal Excel fayllarına qayıdılsın? Ssenari dəyişiklikləri, faktiki məlumat və əlavə illər də sıfırlanacaq.')) return;
      Store.clearCustom().then(function () { clearUserState(); reload('Orijinal model bərpa edildi'); });
    };
    drawWork();
  }
  function q(s) { return host.querySelector(s); }
  function reload(msg) { try { sessionStorage.setItem('makroModel.toast', msg || ''); } catch (e) { /* ignore */ } location.reload(); }
  function clearUserState() { ['makroModel.edits.v1', 'makroModel.actuals.v1', 'makroModel.structure.v1'].forEach(function (k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } }); }

  function addYear() { STATE.view = 'add'; drawWork(); q('#mu-work').scrollIntoView({ behavior: 'smooth' }); }
  function drawAdd(w) {
    var pv;
    try { pv = Store.previewAddYear(STATE.addMode || 'trend'); } catch (e) { w.innerHTML = '<div class="mu-card"><div class="mu-note err">Yeni il əlavə edilə bilmir: ' + esc(e.message) + '</div></div>'; return; }
    var m = STATE.addMode || 'trend', last = C.info.endYear;
    w.innerHTML = '<div class="mu-card"><h3>' + pv.year + ' ilini proqnoza əlavə et</h3>' +
      '<p>Hər il sütunlu vərəqdə ' + last + ' sütunundan sonra yeni sütun yaradılır (Excel-in «sütun əlavə et + sağa doldur» əməliyyatı kimi): düsturlar köçürülür, nisbi istinadlar bir il sürüşür, ' + last + '-dən sağdakı köməkçi bloklar bir sütun sağa keçir və onlara olan istinadlar avtomatik yenilənir.</p>' +
      '<dl class="mu-kv"><dt>Vərəq</dt><dd>' + pv.sheets + '</dd><dt>Yeni xana</dt><dd>' + pv.cells.toLocaleString('az') + '</dd>' + (pv.skipped.length ? '<dt>Ötürülür</dt><dd>' + pv.skipped.map(esc).join(', ') + ' <span class="mu-src" style="font-family:inherit">(eyni il bir neçə sütunda — proqnoz versiyalarının müqayisəsi; heç bir vərəq ondan istifadə etmir)</span></dd>' : '') + '</dl>' +
      '<p style="margin-top:8px"><b>Yeni ilin giriş (ekzogen) dəyərləri:</b></p>' +
      '<label style="display:block;margin:4px 0"><input type="radio" name="mu-mode" value="trend"' + (m === 'trend' ? ' checked' : '') + '> <b>Son ilin artım tempi ilə davam etdirilsin</b> (tövsiyə olunur) — səviyyə göstəriciləri (məzənnələr, əhali, hasilat, məbləğlər) ' + (last - 1) + '→' + last + ' artım tempini davam etdirir; faiz, artım tempi, nisbət və əmsal sətirləri ' + last + ' dəyərində qalır.</label>' +
      '<label style="display:block;margin:4px 0"><input type="radio" name="mu-mode" value="flat"' + (m === 'flat' ? ' checked' : '') + '> <b>' + last + ' dəyəri saxlanılsın</b> — Excel-də sütunu sağa köçürmək kimi; səviyyə göstəriciləri sabit qalır (məs. tərəfdaş ölkələrin məzənnələri dəyişmir, bu da effektiv məzənnəyə və inflyasiyaya təsir edir).</label>' +
      '<div class="mu-note">Hər iki halda yeni ilin bütün fərziyyələrini sonra «Ssenari» bölməsində (iş paneli) və ya «Giriş məlumatları» tablarında (klassik görünüş) dəqiqləşdirə bilərsiniz. Əməliyyat geri qaytarıla bilər («' + pv.year + ' ilini sil»).</div>' +
      '<div class="mu-row"><button class="mu-btn pri" id="mu-add-go">' + pv.year + ' ilini əlavə et</button><button class="mu-btn" id="mu-add-cancel">Ləğv et</button></div></div>';
    Array.prototype.forEach.call(w.querySelectorAll('input[name=mu-mode]'), function (r) { r.onchange = function () { STATE.addMode = r.value; }; });
    w.querySelector('#mu-add-cancel').onclick = function () { STATE.view = null; w.innerHTML = ''; };
    w.querySelector('#mu-add-go').onclick = function () {
      var mode = STATE.addMode || 'trend';
      try { Store.addYear(mode); } catch (e) { alert('Xəta: ' + e.message); return; }
      reload(pv.year + ' ili əlavə edildi — ' + pv.sheets + ' vərəq, ' + pv.cells + ' xana (' + (mode === 'trend' ? 'artım tempi davam etdirildi' : 'son dəyər saxlanıldı') + ')');
    };
  }

  // ------------------------------------------------------------------ realised data editor
  function keyIndicators(Y) {
    var gs = C.sidx['MOE REPORT 3 PAGES|base 60']; if (gs === undefined) return [];
    var M = C.M, out = [], seen = {}, main = '';
    Object.keys(C.rows[gs]).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
      if (r < 4 || r > 142) return;
      var ids = C.rows[gs][r], a = '', u = '';
      ids.forEach(function (id) { if (M.c[id] === 1 && typeof M.V0[id] === 'string') a = M.V0[id].trim(); if (M.c[id] === 2 && typeof M.V0[id] === 'string') u = M.V0[id].trim(); });
      var c = C.colFor(gs, r, Y); if (!c) return;
      var id = M.smap[gs].get(r * 20000 + c); if (id === undefined || !isNum(M.V0[id])) { if (a && /^[A-ZƏÜÖĞİŞÇ ]{4,}/.test(a)) main = ''; return; }
      var sub = !a || /^[a-zəüöğışç]/.test(a);
      if (!sub) main = a;
      var src = C.source(id); if (seen[src]) return; seen[src] = 1;
      if (OUTPUT[C.book(M.sh[src])]) return;   // computed inside the report itself — realised data must go to the model sheets
      out.push({ id: src, label: (sub && main ? main + ' — ' + (a || u) : a) + (a && u ? ' · ' + u : ''), where: C.addr(src), sheet: 'Əsas göstəricilər' });
    });
    return out;
  }
  function allRows(Y) {
    if (STATE.allY === Y && STATE.all) return STATE.all;
    var M = C.M, out = [], core = C.core;
    core.sheets.forEach(function (s, gs) {
      if (core.books[s.b].kind !== 'model' || OUTPUT[core.books[s.b].n] || !C.yearRows(gs).length) return;
      Object.keys(C.rows[gs]).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
        var c = C.colFor(gs, r, Y); if (!c) return;
        var id = M.smap[gs].get(r * 20000 + c); if (id === undefined || !isNum(M.V0[id])) return;
        if (C.yearRows(gs).some(function (y) { return y.r === r; })) return;
        var lab = C.label(gs, r); if (!lab) return;
        out.push({ id: id, label: lab, where: C.addr(id), gs: gs, book: core.books[s.b].n, sheet: s.n, f: fold(lab + ' ' + s.n + ' ' + core.books[s.b].n) });
      });
    });
    STATE.all = out; STATE.allY = Y;
    return out;
  }
  function drawWork() {
    var w = q('#mu-work'); if (!w) return;
    if (STATE.view === 'act') return drawActuals(w);
    if (STATE.view === 'upload') return drawUpload(w);
    if (STATE.view === 'add') return drawAdd(w);
    w.innerHTML = '';
  }
  function drawActuals(w) {
    var Y = STATE.year, info = C.info, act = Store.actuals(), M = C.M, nSaved = 0;
    Object.keys(act).forEach(function (k) { if (/\|/.test(k)) nSaved++; });
    var yrs = []; for (var y = Math.max(2015, info.lastActual - 3); y <= info.endYear; y++) yrs.push(y);
    var h = '<div class="mu-card"><div class="mu-row" style="justify-content:space-between"><h3 style="margin:0">Faktiki məlumat: ' + Y + '</h3><div class="mu-row"><label class="mu-src" style="font-family:inherit" for="mu-y">İl</label><select id="mu-y">' + yrs.map(function (y2) { return '<option' + (y2 === Y ? ' selected' : '') + '>' + y2 + '</option>'; }).join('') + '</select><button class="mu-btn sm" id="mu-close">Bağla</button></div></div>' +
      '<ol class="mu-steps"><li>Göstəricini tapın: «Əsas göstəricilər» nəticə cədvəlindəki göstəricilərin mənbə xanalarıdır; «Bütün sətirlər» — modelin istənilən sətri.</li><li>«Faktiki» sütununa rəqəmi yazın (məs. 3,1). Mümkünsə ən detallı səviyyədə daxil edin — məcmu göstəricilər (ÜDM və s.) komponentlərdən avtomatik yığılır.</li>' +
      '<li>Çox rəqəm üçün: <b>CSV şablonunu</b> yükləyin, Excel-də «Faktiki dəyər» sütununu doldurun və <b>CSV idxal</b> edin.</li><li>«Tətbiq et» — model yenidən hesablanır; ' + Y + '-dən sonrakı proqnozlar faktiki rəqəmlərdən başlayır.</li></ol>' +
      '<div class="mu-row"><div class="mu-tabs" style="margin:0"><button class="mu-tab' + (STATE.tab === 'key' ? ' on' : '') + '" data-t="key">Əsas göstəricilər</button><button class="mu-tab' + (STATE.tab === 'all' ? ' on' : '') + '" data-t="all">Bütün sətirlər</button></div>' +
      '<input type="search" id="mu-q" placeholder="Axtar (məs. ÜDM, inflyasiya, əhali)" value="' + esc(STATE.q) + '" style="min-width:240px"><label style="font-size:13px"><input type="checkbox" id="mu-only"' + (STATE.onlySet ? ' checked' : '') + '> yalnız daxil edilənlər</label></div>' +
      '<div id="mu-list"></div>' +
      '<div class="mu-row" style="margin-top:6px"><button class="mu-btn" id="mu-tpl-key">CSV şablonu: əsas göstəricilər</button><button class="mu-btn" id="mu-tpl-all">CSV şablonu: bütün sətirlər</button><label class="mu-btn" for="mu-csv">CSV idxal et<input type="file" id="mu-csv" accept=".csv,.txt" hidden></label></div>' +
      '<div class="mu-row" style="margin-top:6px;padding-top:10px;border-top:1px solid var(--mu-line2)"><label style="font-size:13.5px"><input type="checkbox" id="mu-mark"' + (Y === info.lastActual + 1 ? ' checked' : '') + '> ' + Y + '-i faktiki il kimi təsdiqlə (gözlənilən il ' + (Y + 1) + ' olacaq)</label>' +
      '<span style="margin-left:auto" class="mu-src" id="mu-cnt"></span><button class="mu-btn" id="mu-clear-y">' + Y + ' üzrə daxil edilənləri sil</button><button class="mu-btn pri" id="mu-apply">Tətbiq et və yenidən hesabla</button></div></div>';
    w.innerHTML = h;
    w.querySelector('#mu-y').onchange = function () { STATE.year = +this.value; STATE.staged = new Map(); drawActuals(w); };
    w.querySelector('#mu-close').onclick = function () { STATE.view = null; w.innerHTML = ''; };
    Array.prototype.forEach.call(w.querySelectorAll('[data-t]'), function (b) { b.onclick = function () { STATE.tab = b.getAttribute('data-t'); drawActuals(w); }; });
    var qi = w.querySelector('#mu-q'); qi.oninput = function () { STATE.q = qi.value; clearTimeout(drawActuals.t); drawActuals.t = setTimeout(function () { drawList(); }, 200); };
    w.querySelector('#mu-only').onchange = function () { STATE.onlySet = this.checked; drawList(); };
    w.querySelector('#mu-tpl-key').onclick = function () { csvTemplate(keyIndicators(Y), Y, 'esas'); };
    w.querySelector('#mu-tpl-all').onclick = function () { csvTemplate(allRows(Y), Y, 'butun'); };
    w.querySelector('#mu-csv').onchange = function () { var f = this.files[0]; if (!f) return; var rd = new FileReader(); rd.onload = function () { importCSV(String(rd.result), Y); drawList(); }; rd.readAsText(f, 'utf-8'); this.value = ''; };
    w.querySelector('#mu-clear-y').onclick = function () {
      var n = 0; Object.keys(act).forEach(function (k) { var id = keyToId(k); if (id >= 0 && yearOfId(id) === Y) { STATE.staged.set(id, null); n++; } });
      toastMsg(n ? n + ' dəyər silinmək üçün qeyd edildi — «Tətbiq et» basın' : Y + ' üzrə daxil edilmiş dəyər yoxdur'); drawList();
    };
    w.querySelector('#mu-apply').onclick = function () { applyActuals(Y, w.querySelector('#mu-mark').checked); };
    w.onchange = function (e) {
      var t = e.target; if (!t.classList || !t.classList.contains('mu-v')) return;
      var id = +t.getAttribute('data-id'), v = parseNum(t.value);
      if (v === undefined) { toastMsg('Rəqəm daxil edin (məs. 3,1)'); t.value = ''; return; }
      STATE.staged.set(id, v); t.classList.toggle('set', v !== null); updCount();
    };
    drawList();
  }
  function keyToId(k) { var p = k.split('|'); return p.length < 3 ? -1 : C.idOf(p[0], p[1], p[2]); }
  function yearOfId(id) { var M = C.M, yr = C.yearRows(M.sh[id]); return Store.colYearIn(yr, M.r[id], M.c[id]); }
  function updCount() { var el = host.querySelector('#mu-cnt'); if (!el) return; var a = 0, d = 0; STATE.staged.forEach(function (v) { if (v === null) d++; else a++; }); el.textContent = (a || d) ? 'Tətbiq olunmamış: ' + a + ' yeni' + (d ? ', ' + d + ' silinəcək' : '') : ''; }
  function drawList() {
    var Y = STATE.year, el = host.querySelector('#mu-list'); if (!el) return;
    var act = Store.actuals(), saved = {}; Object.keys(act).forEach(function (k) { var id = keyToId(k); if (id >= 0) saved[id] = act[k]; });
    var M = C.M, qq = fold(STATE.q.trim()), list = STATE.tab === 'key' ? keyIndicators(Y) : allRows(Y);
    list = list.filter(function (e) { return (!qq || (e.f || fold(e.label + ' ' + e.where)).indexOf(qq) >= 0) && (!STATE.onlySet || saved[e.id] !== undefined || STATE.staged.has(e.id)); });
    function row(e) {
      var sv = saved[e.id], st = STATE.staged.has(e.id) ? STATE.staged.get(e.id) : undefined, val = st !== undefined ? st : sv;
      var excel = M.VX ? M.VX[e.id] : M.V0[e.id], base = M.V0[e.id];
      return '<tr><td>' + esc(e.label) + '<div class="mu-src">' + esc(e.where) + (M.ft[e.id] >= 0 ? ' · düstur' : ' · giriş') + '</div></td><td class="n">' + nf(isNum(excel) ? excel : base) + '</td><td class="n">' + (sv !== undefined ? nf(base) : '') + '</td>' +
        '<td class="n"><input class="mu-v' + (st !== undefined && st !== null ? ' set' : sv !== undefined && st === undefined ? ' saved' : '') + '" data-id="' + e.id + '" value="' + esc(val === null || val === undefined ? '' : raw(val)) + '" inputmode="decimal" aria-label="Faktiki dəyər"></td></tr>';
    }
    var head = '<table class="mu-t"><thead><tr><th>Göstərici</th><th class="n">Model (Excel)</th><th class="n">Hazırda (faktiki ilə)</th><th class="n">Faktiki ' + Y + '</th></tr></thead><tbody>';
    if (STATE.tab === 'key' || qq || STATE.onlySet) {
      var shown = list.slice(0, 400);
      el.innerHTML = list.length ? '<div class="mu-wrap">' + head + shown.map(row).join('') + '</tbody></table></div>' + (list.length > 400 ? '<p class="mu-src">İlk 400 sətir göstərilir — axtarışı dəqiqləşdirin.</p>' : '') : '<div class="mu-note">Uyğun sətir tapılmadı' + (STATE.tab === 'key' ? ' (yekun hesabat vərəqi «base 60» bu modeldə tapılmadı və ya ' + Y + ' sütunu yoxdur)' : '') + '.</div>';
    } else {
      var groups = [], gi = {};
      list.forEach(function (e) { var k = e.gs; if (gi[k] === undefined) { gi[k] = groups.length; groups.push({ gs: k, book: e.book, sheet: e.sheet, rows: [] }); } groups[gi[k]].rows.push(e); });
      el.innerHTML = '<div class="mu-wrap">' + groups.map(function (g) { var open = STATE.open.has(g.gs); return '<details class="mu-g" data-g="' + g.gs + '"' + (open ? ' open' : '') + '><summary>' + esc(g.book.replace(/^MOE /, '') + ' › ' + g.sheet) + ' <span class="mu-chip">' + g.rows.length + '</span></summary><div class="mu-gb">' + (open ? head + g.rows.map(row).join('') + '</tbody></table>' : '') + '</div></details>'; }).join('') + '</div>';
      Array.prototype.forEach.call(el.querySelectorAll('details.mu-g'), function (d) {
        d.addEventListener('toggle', function () { var gs = +d.getAttribute('data-g'); if (d.open) { STATE.open.add(gs); var b = d.querySelector('.mu-gb'); if (!b.innerHTML) b.innerHTML = head + groups[gi[gs]].rows.map(row).join('') + '</tbody></table>'; } else STATE.open.delete(gs); });
      });
    }
    updCount();
  }
  function csvTemplate(list, Y, tag) {
    var M = C.M, lines = ['İş kitabı;Vərəq;Xana;Göstərici;İl;Model dəyəri;Faktiki dəyər'];
    list.forEach(function (e) { var s = C.core.sheets[M.sh[e.id]]; lines.push([C.core.books[s.b].n, s.n, Store.colName(M.c[e.id]) + M.r[e.id], e.label, Y, raw(M.V0[e.id]), ''].map(function (x) { x = String(x); return /[;"\n]/.test(x) ? '"' + x.replace(/"/g, '""') + '"' : x; }).join(';')); });
    download('faktiki-' + Y + '-' + tag + '.csv', '﻿' + lines.join('\r\n'), 'text/csv');
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
  function importCSV(text, Y) {
    var rows = parseCSV(text); if (rows.length < 2) { toastMsg('CSV faylında məlumat yoxdur'); return; }
    var h = rows[0].map(function (x) { return fold(x.trim()); });
    function col(names) { for (var i = 0; i < h.length; i++) for (var j = 0; j < names.length; j++) if (h[i] === names[j]) return i; return -1; }
    var cb = col(['is kitabi', 'kitab', 'fayl']), cs = col(['vereq', 'sheet']), cx = col(['xana', 'cell', 'unvan']), cv = col(['faktiki deyer', 'faktiki', 'deyer', 'value']);
    var cr = col(['setir', 'row']), cy = col(['il', 'year']);
    if (cb < 0 || cs < 0 || cv < 0 || (cx < 0 && (cr < 0 || cy < 0))) { toastMsg('Başlıqlar tanınmadı: «İş kitabı; Vərəq; Xana; Faktiki dəyər» sütunları lazımdır'); return; }
    var ok = 0, bad = 0, badList = [];
    rows.slice(1).forEach(function (r, i) {
      if (!r.length || r.every(function (x) { return !String(x).trim(); })) return;
      var sval = (r[cv] || '').trim(); if (sval === '') return;
      var v = parseNum(sval), id;
      if (cx >= 0 && (r[cx] || '').trim()) id = C.idOf((r[cb] || '').trim(), (r[cs] || '').trim(), r[cx]);
      else { var gs = C.sidx[(r[cb] || '').trim() + '|' + (r[cs] || '').trim()], c = gs === undefined ? null : C.colFor(gs, +r[cr], +r[cy] || Y); id = c ? (C.M.smap[gs].get(+r[cr] * 20000 + c)) : -1; if (id === undefined) id = -1; }
      if (id < 0 || v === undefined || v === null) { bad++; if (badList.length < 5) badList.push('sətir ' + (i + 2)); return; }
      STATE.staged.set(id, v); ok++;
    });
    STATE.onlySet = true; var only = host.querySelector('#mu-only'); if (only) only.checked = true;
    toastMsg('İdxal: ' + ok + ' dəyər qəbul edildi' + (bad ? ', ' + bad + ' sətir tanınmadı (' + badList.join(', ') + ')' : '') + ' — yoxlayıb «Tətbiq et» basın');
  }
  function applyActuals(Y, mark) {
    var act = Store.actuals(), n = 0;
    STATE.staged.forEach(function (v, id) { var k = C.key(id); if (v === null) { if (act[k] !== undefined) { delete act[k]; n++; } } else { act[k] = v; n++; } });
    if (!n && !mark) { toastMsg('Tətbiq ediləcək dəyişiklik yoxdur'); return; }
    if (!Store.saveActuals(act)) { alert('Brauzer yaddaşına yazmaq mümkün olmadı.'); return; }
    if (mark) { var st = Store.structure(); if (st.lastActual < Y) { st.lastActual = Y; Store.saveStructure(st); } }
    reload(n + ' faktiki dəyər tətbiq edildi' + (mark ? ' · ' + Y + ' faktiki il kimi təsdiqləndi' : ''));
  }

  // ------------------------------------------------------------------ upload new Excel files
  function startUpload(files) {
    if (!files.length) return;
    var books = C.core.books.filter(function (b) { return b.kind === 'model'; }).map(function (b) { return b.n; });
    STATE.view = 'upload';
    STATE.upload = { files: files.map(function (f) { var n = B.norm(f.name), m = null; books.forEach(function (b) { if (B.norm(b) === n) m = b; }); return { file: f, book: m }; }), books: books, log: [], result: null, busy: false };
    drawWork(); q('#mu-work').scrollIntoView({ behavior: 'smooth' });
  }
  function drawUpload(w) {
    var U = STATE.upload, h = '<div class="mu-card"><div class="mu-row" style="justify-content:space-between"><h3 style="margin:0">Yeni Excel faylları</h3><button class="mu-btn sm" id="mu-up-close">Bağla</button></div>';
    h += '<table class="mu-t"><thead><tr><th>Fayl</th><th class="n">Ölçü</th><th>Modeldə hansı faylı əvəz edir</th></tr></thead><tbody>' + U.files.map(function (f, i) {
      return '<tr><td>' + esc(f.file.name) + '</td><td class="n">' + (f.file.size / 1048576).toFixed(1) + ' MB</td><td><select data-map="' + i + '"' + (U.busy ? ' disabled' : '') + '><option value="">— istifadə etmə —</option>' + U.books.map(function (b) { return '<option' + (b === f.book ? ' selected' : '') + '>' + esc(b) + '</option>'; }).join('') + '</select>' + (!f.book ? ' <span class="mu-chip warn">ad tanınmadı — seçin</span>' : '') + '</td></tr>';
    }).join('') + '</tbody></table>';
    var mapped = U.files.filter(function (f) { return f.book; }).length;
    h += '<p class="mu-src" style="font-family:inherit">Yüklənməyən ' + (U.books.length - mapped) + ' fayl mövcud modeldən götürüləcək. Fayllar arası əlaqələr xana ünvanı ilə qurulur — faylların birində sətir/sütun əlavə edilibsə, ona istinad edən faylları da birlikdə yükləyin.</p>';
    h += '<div class="mu-row"><button class="mu-btn pri" id="mu-run"' + (U.busy || !mapped ? ' disabled' : '') + '>Yoxla və hesabla</button></div>';
    if (U.log.length) h += '<div class="mu-log" id="mu-log">' + esc(U.log.join('\n')) + '</div>';
    if (U.result) h += reportHTML(U.result);
    h += '</div>';
    w.innerHTML = h;
    w.querySelector('#mu-up-close').onclick = function () { if (U.busy) return; STATE.view = null; STATE.upload = null; w.innerHTML = ''; };
    Array.prototype.forEach.call(w.querySelectorAll('[data-map]'), function (s) { s.onchange = function () { U.files[+s.getAttribute('data-map')].book = s.value || null; U.result = null; drawUpload(w); }; });
    var run = w.querySelector('#mu-run'); if (run) run.onclick = function () { runBuild(w); };
    var act = w.querySelector('#mu-activate'); if (act) act.onclick = function () { activate(); };
    var zip = w.querySelector('#mu-zip'); if (zip) zip.onclick = function () { exportZip(); };
    var lg = w.querySelector('#mu-log'); if (lg) lg.scrollTop = lg.scrollHeight;
  }
  function log(msg) { var U = STATE.upload; U.log.push(msg); var el = host.querySelector('#mu-log'); if (el) { el.textContent = U.log.join('\n'); el.scrollTop = el.scrollHeight; } else drawUpload(host.querySelector('#mu-work')); }
  function runBuild(w) {
    var U = STATE.upload, dup = {};
    for (var i = 0; i < U.files.length; i++) { var b = U.files[i].book; if (b) { if (dup[b]) { toastMsg('«' + b + '» üçün iki fayl seçilib'); return; } dup[b] = 1; } }
    if (typeof DecompressionStream === 'undefined') { alert('Bu brauzer Excel fayllarını aça bilmir (DecompressionStream yoxdur). Chrome, Edge, Safari 16.4+ və ya Firefox 113+ istifadə edin.'); return; }
    U.busy = true; U.log = []; U.result = null; drawUpload(w);
    var t0 = Date.now(), base = Store.base, core = base.core, raws = [], byBook = {};
    U.files.forEach(function (f) { if (f.book) byBook[f.book] = f.file; });
    var chain = Promise.resolve();
    core.books.forEach(function (bk, bi) {
      if (bk.kind !== 'model') return;
      if (byBook[bk.n]) chain = chain.then(function () {
        log('▸ ' + byBook[bk.n].name + ' oxunur (' + bk.n + ')…');
        return B.readXlsx(byBook[bk.n], bk.n, function (m) { log(m); }).then(function (rb) {
          var nc = 0, nf2 = 0; rb.sheets.forEach(function (s) { nc += s.cells.length; s.cells.forEach(function (c) { if (c.f) nf2++; }); });
          log('  ✓ ' + rb.sheets.length + ' vərəq, ' + nc.toLocaleString('az') + ' xana, ' + nf2.toLocaleString('az') + ' düstur, ' + rb.ext.length + ' xarici keçid');
          raws.push(rb);
        });
      });
      else chain = chain.then(function () { raws.push(B.rawFromCore(core, base.parts, bi)); });
    });
    chain.then(function () {
      core.books.forEach(function (bk, bi) { if (bk.kind === 'ext') raws.push(B.rawFromCore(core, base.parts, bi)); });
      log('▸ Düsturlar təhlil edilir və fayllar arası əlaqələr qurulur…');
      return new Promise(function (r) { setTimeout(r, 30); });
    }).then(function () {
      var res = B.compile(raws);
      log('  ✓ ' + res.nForm.toLocaleString('az') + ' düstur, ' + res.core.tplCode.length + ' şablon, ' + res.core.ncells.toLocaleString('az') + ' xana; problem: ' + res.problems.length);
      log('▸ Excel-in saxladığı nəticələrlə yoxlama…');
      return new Promise(function (r) { setTimeout(function () { r(res); }, 30); });
    }).then(function (res) {
      var v = B.verify(res);
      log('  ✓ düstur-düstur: ' + v.local.ok + ' / ' + v.formulas + ' · tam hesablama: ' + v.global.ok + ' / ' + v.formulas + ' · keçid keşləri: ' + v.link.ok + ' / ' + (v.link.ok + v.link.bad.length));
      log('Hazırdır — ' + ((Date.now() - t0) / 1000).toFixed(1) + ' san.');
      U.result = { res: res, v: v, files: U.files.filter(function (f) { return f.book; }).map(function (f) { return { name: f.file.name, book: f.book, size: f.file.size, modified: f.file.lastModified }; }) };
      U.busy = false; drawUpload(w);
    }).catch(function (e) {
      U.busy = false; log('✗ Xəta: ' + (e && e.message ? e.message : e)); drawUpload(w); console.error(e);
    });
  }
  function pct(a, b) { return b ? (100 * a / b).toFixed(3).replace('.', ',') + '%' : '—'; }
  function diffList(title, d) {
    var items = d.bad.filter(Boolean); if (!d.bad.length) return '';
    return '<details class="mu-g"><summary>' + esc(title) + ' <span class="mu-chip warn">' + d.bad.length + '</span></summary><div style="padding:0 12px 10px"><table class="mu-t"><thead><tr><th>Xana</th><th class="n">Hesablanan</th><th class="n">Excel-də</th></tr></thead><tbody>' +
      items.slice(0, 100).map(function (x) { return '<tr><td class="mu-src">' + esc(x[0]) + '</td><td class="n">' + esc(x[1]) + '</td><td class="n">' + esc(x[2]) + '</td></tr>'; }).join('') + '</tbody></table>' + (d.bad.length > items.length ? '<p class="mu-src">… və daha ' + (d.bad.length - items.length) + '</p>' : '') + '</div></details>';
  }
  function reportHTML(R) {
    var v = R.v, res = R.res, lb = v.link.ok + v.link.bad.length;
    var good = v.local.ok / v.formulas >= 0.999 && v.global.ok / v.formulas >= 0.999;
    var h = '<div class="mu-stats" style="margin-top:12px">' +
      '<div class="mu-stat"><b>' + res.core.sheets.length + '</b><span>vərəq</span></div><div class="mu-stat"><b>' + v.formulas.toLocaleString('az') + '</b><span>düstur</span></div>' +
      '<div class="mu-stat"><b>' + pct(v.local.ok, v.formulas) + '</b><span>düstur-düstur yoxlama</span></div><div class="mu-stat"><b>' + pct(v.global.ok, v.formulas) + '</b><span>tam yenidən hesablama</span></div>' +
      '<div class="mu-stat"><b>' + (lb ? pct(v.link.ok, lb) : '—') + '</b><span>fayllararası keçid keşləri</span></div><div class="mu-stat"><b>' + res.problems.length + '</b><span>problem</span></div></div>';
    h += '<div class="mu-note ' + (good && !res.problems.length ? 'ok' : good ? 'warn' : 'err') + '" style="margin-top:10px">' + (good
      ? '<b>Yoxlama uğurludur.</b> Sistemin hesabladığı nəticələr Excel-in saxladığı nəticələrlə üst-üstə düşür. Fərqlər (varsa) aşağıda göstərilib.'
      : '<b>Diqqət:</b> nəticələrin bir hissəsi Excel-in saxladığı dəyərlərdən fərqlənir. Adətən bu, faylların Excel-də yenidən hesablanmadan saxlanması və ya fayllar arası keçidlərin köhnə qalması deməkdir. Faylları Excel-də birlikdə açıb, «Hesabla» (F9) və «Keçidləri yenilə» edib yenidən saxlayın.') + '</div>';
    if (res.problems.length) h += '<details class="mu-g" open><summary>Problemlər <span class="mu-chip warn">' + res.problems.length + '</span></summary><div style="padding:0 12px 10px"><table class="mu-t"><tbody>' + res.problems.slice(0, 100).map(function (p) { return '<tr><td class="mu-src" style="white-space:nowrap">' + esc(p.where) + '</td><td>' + esc(p.what) + '</td></tr>'; }).join('') + '</tbody></table></div></details>';
    h += diffList('Düstur-düstur yoxlamada fərqlənən xanalar', v.local) + diffList('Tam hesablamada fərqlənən xanalar', v.global) + diffList('Köhnə qalmış fayllararası keçidlər', v.link);
    h += compareHTML(v.model, res.core);
    h += '<div class="mu-row" style="margin-top:12px"><button class="mu-btn pri" id="mu-activate">Bu modeli aktivləşdir</button><button class="mu-btn" id="mu-zip">«data» qovluğunu ZIP kimi yüklə</button></div>' +
      '<p class="mu-src" style="font-family:inherit">«Aktivləşdir» — model bu brauzerdə yadda saxlanılır və hər iki səhifədə (klassik görünüş və iş paneli) istifadə olunur. Hamı üçün daimi etmək: ZIP-i açıb içindəki <code>data</code> qovluğunu sayt qovluğundakı <code>data</code> qovluğunun yerinə köçürün.</p>';
    return h;
  }
  function compareHTML(Mn, coreN) {
    var Mo = C.M, co = C.core;
    function find(core, b, s) { for (var i = 0; i < core.sheets.length; i++) if (core.books[core.sheets[i].b].n === b && core.sheets[i].n === s) return i; return -1; }
    var go = find(co, 'MOE REPORT 3 PAGES', 'base 60'), gn = find(coreN, 'MOE REPORT 3 PAGES', 'base 60'); if (go < 0 || gn < 0) return '';
    function cell(M, core, gs, r, y) {
      var ids = []; for (var i = 0; i < M.N; i++) if (M.sh[i] === gs && M.r[i] <= 4) ids.push(i);
      var yr = Store.yearRowsOf(ids, M.r, M.c, M.V0), c = null; yr.forEach(function (x) { for (var k in x.d) if (x.d[k] === y) c = +k; });
      var id = c ? M.smap[gs].get(r * 20000 + c) : undefined; return id === undefined ? null : M.V0[id];
    }
    var rows = [[4, 'ÜDM, mln AZN'], [6, 'Real ÜDM artımı, %'], [13, 'Qeyri-neft-qaz ÜDM artımı, %'], [124, 'İnflyasiya, %'], [126, 'Cari hesab balansı, mln USD'], [92, 'Büdcə kəsiri, mln AZN']];
    var yrs = []; for (var y = C.info.lastActual + 1; y <= Math.min(C.info.endYear, C.info.lastActual + 6); y++) yrs.push(y);
    return '<details class="mu-g"><summary>Əsas göstəricilər: mövcud model → yeni fayllar</summary><div style="padding:0 12px 10px;overflow-x:auto"><table class="mu-t"><thead><tr><th>Göstərici</th>' + yrs.map(function (y2) { return '<th class="n">' + y2 + '</th>'; }).join('') + '</tr></thead><tbody>' +
      rows.map(function (r) { return '<tr><td>' + r[1] + '</td>' + yrs.map(function (y2) { var a = cell(Mo, co, go, r[0], y2), b = cell(Mn, coreN, gn, r[0], y2); return '<td class="n">' + nf(a) + ' → <b>' + nf(b) + '</b></td>'; }).join('') + '</tr>'; }).join('') + '</tbody></table></div></details>';
  }
  function metaPatch(R) {
    var res = R.res, v = R.v, lt = B.linkTables(res), xrefs = 0;
    lt.bookLinks.forEach(function (e) { if (!/^EXT: /.test(e.src)) xrefs += e.n; });
    var models = res.core.books.filter(function (b) { return b.kind === 'model'; });
    return {
      verify: { books: models.length, sheets: models.reduce(function (s, b) { return s + b.sheets.length; }, 0), formulas: v.formulas, literals: res.core.K.length, xrefs: xrefs,
        local_ok: v.local.ok, global_ok: v.global.ok, link_total: v.link.ok + v.link.bad.length, link_ok: v.link.ok,
        diffs: v.local.bad.filter(Boolean).slice(0, 20).map(function (d) { return d[0] + ': Excel-də ' + d[2] + ', hesablanan ' + d[1]; }) },
      bookLinks: lt.bookLinks, sheetLinks: lt.sheetLinks,
      stale: { date: new Date().toISOString(), files: R.files.map(function (f) { return f.name; }) }
    };
  }
  function endYearOf(core, parts) {
    var P = parts[0], gs = -1; core.sheets.forEach(function (s, i) { if (core.books[s.b].n === 'MOE SNA' && s.n === 'IN') gs = i; });
    var best = 0;
    for (var i = 0; i < P.ids.length; i++) if ((gs < 0 || P.sh[i] === gs) && P.r[i] <= 4 && typeof P.v[i] === 'number' && P.v[i] >= 2020 && P.v[i] <= 2060 && P.v[i] === Math.floor(P.v[i])) best = Math.max(best, P.v[i]);
    return best || 2030;
  }
  function activate() {
    var R = STATE.upload.result; if (!R) return;
    if (!confirm('Yeni model aktivləşdirilsin? Mövcud ssenari dəyişiklikləri, faktiki məlumat və əlavə illər sıfırlanacaq (lazımdırsa, əvvəlcə ixrac edin).')) return;
    var res = R.res, meta = metaPatch(R), rec = { core: res.core, parts: res.parts, meta: meta, info: { date: new Date().toISOString(), files: R.files, end: endYearOf(res.core, res.parts) } };
    Store.putCustom(rec).then(function () { clearUserState(); reload('Yeni model aktivləşdirildi: ' + R.files.map(function (f) { return f.name; }).join(', ')); })
      .catch(function (e) { alert('Brauzer yaddaşına yazmaq mümkün olmadı (' + (e && e.message) + '). «data» qovluğunu ZIP kimi yükləyib sayt qovluğuna köçürün.'); });
  }
  function exportZip() {
    var R = STATE.upload.result; if (!R) return;
    var meta = {}, base = root.MODEL_META || {}; for (var k in base) meta[k] = base[k];
    var p = metaPatch(R); for (k in p) meta[k] = p[k];
    var blob = B.zipStore(B.dataFiles(R.res.core, R.res.parts, meta));
    if (!download('makro-model-data-' + new Date().toISOString().slice(0, 10) + '.zip', blob)) alert('Yükləmə bu mühitdə mümkün olmadı.');
  }
  var TT = null;
  function toastMsg(m) {
    var t = document.getElementById('toast');
    if (t) { t.textContent = m; t.hidden = false; clearTimeout(TT); TT = setTimeout(function () { t.hidden = true; }, 3200); } else alert(m);
  }
  function pendingToast() { try { var m = sessionStorage.getItem('makroModel.toast'); if (m) { sessionStorage.removeItem('makroModel.toast'); setTimeout(function () { toastMsg(m); }, 300); } } catch (e) { /* ignore */ } }

  root.MakroUpdate = { render: render, pendingToast: pendingToast };
})(typeof window !== 'undefined' ? window : globalThis);
