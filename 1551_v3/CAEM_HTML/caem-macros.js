/* CAEM — the workbook's VBA macros ported to JavaScript (modules CloseWedge and Calibration).
   Excel runs them with automatic calculation, so every assignment below is followed by a recalculation of its dependents.
   The macros address cells through defined names plus a column offset j = 0..4 (the five projection years). */
(function (root) {
  'use strict';
  var X = root.X;

  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
  // defined name -> cell id (single-cell names such as GDP_diff = '1d. Real GDP - Expenditure'!$P$384)
  function nameCell(M, name) {
    var core = M.core, list = core.names || [], f = null;
    for (var i = 0; i < list.length; i++) if (list[i].n.toLowerCase() === name.toLowerCase()) { f = list[i].f; break; }
    if (!f) throw new Error('Ad tapılmadı: ' + name);
    var m = /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]{1,3})\$?(\d+)$/.exec(f);
    if (!m) throw new Error('Ad tək xanaya istinad etmir: ' + name + ' = ' + f);
    var sn = (m[1] || m[2]).replace(/''/g, "'"), gs = -1;
    core.sheets.forEach(function (s, k) { if (s.n === sn && core.books[s.b].kind === 'model') gs = k; });
    if (gs < 0) throw new Error('Vərəq tapılmadı: ' + sn);
    return { gs: gs, r: +m[4], c: colNum(m[3]) };
  }
  function at(M, nc, j) {
    var id = M.smap[nc.gs].get(nc.r * 20000 + nc.c + j);
    if (id === undefined) throw new Error('Xana yoxdur: ' + M.core.sheets[nc.gs].n + ' sütun ' + (nc.c + j) + ', sətir ' + nc.r);
    return id;
  }
  function num(v) { if (typeof v === 'number') return v; if (v instanceof X.XErr) throw new Error('Xəta dəyəri: ' + v.e); if (v === null || v === '') return 0; var x = +v; if (isNaN(x)) throw new Error('Rəqəm deyil: ' + v); return x; }

  // Context: set(id, v) assigns an input and recalculates dependents (the host decides how edits are recorded)
  function Ctx(M, set) { this.M = M; this.set = set; this.changed = new Map(); this.recalcs = 0; }
  Ctx.prototype.put = function (id, v) { this.set(id, v); this.changed.set(id, v); this.recalcs++; };
  Ctx.prototype.val = function (id) { return num(this.M.V[id]); };

  // Excel Goal Seek: change one input cell until target = goal (secant iterations with a Newton start, as Excel)
  function goalSeek(cx, target, change, goal) {
    var x0 = cx.val(change), f0 = cx.val(target) - goal;
    if (Math.abs(f0) < 1e-9) return true;
    var h = Math.max(1e-6, Math.abs(x0) * 1e-6), x1 = x0 + h;
    cx.put(change, x1); var f1 = cx.val(target) - goal;
    for (var it = 0; it < 100; it++) {
      if (Math.abs(f1) < 1e-9) return true;
      var d = f1 - f0;
      if (d === 0 || !isFinite(d)) { x1 = x1 + h * 10; cx.put(change, x1); f1 = cx.val(target) - goal; continue; }
      var x2 = x1 - f1 * (x1 - x0) / d;
      if (!isFinite(x2)) break;
      x0 = x1; f0 = f1; x1 = x2;
      cx.put(change, x1); f1 = cx.val(target) - goal;
      if (Math.abs(x1 - x0) < 1e-12 * Math.max(1, Math.abs(x1))) break;
    }
    return Math.abs(f1) < 1e-6;
  }

  // Sub ResidualYP / ResidualCP (module CloseWedge): close the supply–demand wedge (GDP_diff) and the GDP-deflator wedge
  // (deflator_diff) for 5 projection years, year by year; YP moves the supply-side growth (GDP_percent_change),
  // CP moves the adjustment that closes the wedge on the demand side (WedgeSD); both move the initial deflator guess.
  function reconcile(M, set, mode) {
    var cx = new Ctx(M, set), ii = 20, jj = 4, tol = 0.001, i = 1;
    var gdpDiff = nameCell(M, 'GDP_diff'), defDiff = nameCell(M, 'deflator_diff');
    var ch1 = nameCell(M, mode === 'CP' ? 'WedgeSD' : 'GDP_percent_change'), ch2 = nameCell(M, 'deflator_percent_change');
    var maxWedge = 0, flag1 = false, flag2 = false, j, cur = 0;
    for (j = 0; j <= jj; j++) {
      // (the VBA counter i is not reset between years: at most 20 Goal-Seek rounds in total)
      while (i <= ii && (Math.abs(cx.val(at(M, gdpDiff, j))) > tol || Math.abs(cx.val(at(M, defDiff, j))) > tol)) {
        goalSeek(cx, at(M, gdpDiff, j), at(M, ch1, j), 0);
        goalSeek(cx, at(M, defDiff, j), at(M, ch2, j), 0);
        i++;
      }
      cur = Math.abs(cx.val(at(M, gdpDiff, j)));          // VBA checks GDP_diff twice (as written in the macro)
      maxWedge = Math.max(cur, maxWedge);
      if (cur > tol) { flag1 = true; break; }
      var dbl = 0; for (var k = 0; k <= j; k++) dbl = Math.max(dbl, Math.abs(cx.val(at(M, gdpDiff, k))));
      if (dbl > tol) { flag2 = true; break; }
    }
    var status = flag1 ? 'notclosed' : flag2 ? 'forward' : 'done';
    return { status: status, year: j + 1, maxWedge: maxWedge, tol: tol, rounds: i - 1, changed: cx.changed, recalcs: cx.recalcs };
  }

  // Sub CaliManual (module Calibration): back out the add factors so that the six behavioural equations reproduce the final
  // (manual or model-determined) projections
  function calibrate(M, set) {
    var cx = new Ctx(M, set), jj = 4, j;
    var N = function (n) { return nameCell(M, n); };
    var logEq = [['Add_factor_C', 'Final_C', 'Model_C'], ['Add_factor_I', 'Final_I', 'Model_I'], ['Add_factor_M', 'Final_M', 'Model_M'], ['Add_factor_X', 'Final_X', 'Model_X']];
    var linEq = [['Add_factor_Pi', 'Final_Pi', 'Model_Pi'], ['Add_factor_intr', 'Final_intr', 'Model_intr']];
    var cells = logEq.concat(linEq).map(function (t) { return [N(t[0]), N(t[1]), N(t[2])]; });
    for (j = 0; j <= jj; j++) {
      cells.forEach(function (t, q) {
        var a = at(M, t[0], j), fin = cx.val(at(M, t[1], j)), mod = cx.val(at(M, t[2], j)), nv;
        if (q < 4) { if (fin <= 0 || mod <= 0) throw new Error('Loqarifm üçün müsbət dəyər lazımdır (' + ['C', 'I', 'M', 'X'][q] + ', il ' + (j + 1) + ')'); nv = cx.val(a) + 100 * (Math.log(fin) - Math.log(mod)); }
        else nv = cx.val(a) + fin - mod;
        cx.put(a, nv);
      });
    }
    return { status: 'done', changed: cx.changed, recalcs: cx.recalcs };
  }

  root.CaemMacros = { reconcile: reconcile, calibrate: calibrate, nameCell: nameCell, at: at };
})(typeof window !== 'undefined' ? window : globalThis);
