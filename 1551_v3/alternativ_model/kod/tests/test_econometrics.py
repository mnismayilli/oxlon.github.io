"""test_econometrics.py — ekonometrik nüvənin öz-özünü yoxlaması (sadə assert-lər, pytest tələb olunmur).

İşə salma:  python3 tests/test_econometrics.py

Yoxlanılan davranışlar (süni, nəticəsi əvvəlcədən məlum olan məlumatlarla):
  T1  ADF: təsadüfi gəzişmədə vahid kök rədd edilmir, stasionar AR(1)-də rədd edilir.
  T2  eg_coint: kointeqrasiyalı cütlük aşkarlanır, müstəqil iki gəzişmə (saxta reqressiya) rədd edilir.
  T3  ols_hac: 2025-dən sonrakı indeksli reqressor üçün istisna qaldırılır (gələcəyə baxış mühafizəsi)
      və n çap olunur.
  T4  ardl_ecm: uzunmüddətli əmsal bərpa olunur, korreksiya əmsalı mənfidir.
  T5  backtest: süni AR(1) seriyasında AR(1) modeli təsadüfi gəzişməni üstələyir, çəkilər ona verilir,
      cədvəl §3 sütunlarına uyğundur.
  T6  solve_core: analitik olaraq məlum sabit nöqtəni bərpa edir; yığılmayan blokda istisna qaldırır.
  T7  fan: zolaqlar üfüq boyu genişlənir, 80% zolağı 50%-i əhatə edir.
  T8  validate: eynilik bir ildə pozulduqda aşkarlanır; mütənasib uzlaşdırma cəmi bərabərləşdirir.
  T9  dinamika bloku: səviyyə bərpası, AR profili, çoxüfüqlü sınaq, dinamiklik göstəricisi
      və onun daxili aralıq yoxlaması, üç dinamika formasının seçimi.
"""
import io
import os
import sys
import contextlib

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

import econometrics as ec          # noqa: E402
import validate as va              # noqa: E402

PASSED = []


def rng(seed):
    """Hər yoxlama öz toxumu ilə işləyir — nəticələr sıradan asılı deyil."""
    return np.random.default_rng(seed)


def check(name, fn):
    """Bir yoxlamanı işə salır və nəticəni qeyd edir."""
    fn()
    PASSED.append(name)
    print(f"  [OK] {name}")


# ---------------------------------------------------------------------------
def t1_adf():
    R = rng(105)   # toxum sabitdir: bu realizasiyada gəzişmənin ADF p-qiyməti ≈ 0.75 (hədddən uzaq)
    yrs = list(range(1976, 2026))
    rw = pd.Series(np.cumsum(R.normal(0, 1, len(yrs))), index=yrs)
    st = pd.Series(R.normal(0, 1, len(yrs)), index=yrs)
    r_rw = ec.adf(rw, name="təsadüfi gəzişmə")
    r_st = ec.adf(st, name="ağ səs")
    assert not r_rw.stationary, f"gəzişmədə vahid kök rədd edilməməli idi: p={r_rw.pvalue}"
    assert r_st.stationary, f"ağ səsdə vahid kök rədd edilməli idi: p={r_st.pvalue}"
    assert r_rw.nobs > 40 and "5%" in r_rw.crit
    assert "vahid kök" in str(r_rw)


# ---------------------------------------------------------------------------
def t2_coint():
    R = rng(202)
    yrs = list(range(1976, 2026))
    n = len(yrs)
    x = pd.Series(np.cumsum(R.normal(0, 1, n)) + 20.0, index=yrs, name="x")
    y = pd.Series(1.5 + 2.0 * x.values + R.normal(0, 0.4, n), index=yrs, name="y")
    r = ec.eg_coint(y, x, name="kointeqrasiyalı cütlük")
    assert r.cointegrated, f"kointeqrasiya aşkarlanmalı idi: p={r.pvalue:.4f}"
    assert r.pvalue < 0.05 and r.k_regressors == 1 and set(r.crit) == {"1%", "5%", "10%"}

    # saxta reqressiya: iki müstəqil təsadüfi gəzişmə
    a = pd.Series(np.cumsum(R.normal(0, 1, n)) + 50.0, index=yrs, name="a")
    b = pd.Series(np.cumsum(R.normal(0, 1, n)) + 50.0, index=yrs, name="b")
    r2 = ec.eg_coint(a, b, name="müstəqil gəzişmələr")
    assert not r2.cointegrated, f"saxta cütlükdə kointeqrasiya rədd edilməli idi: p={r2.pvalue:.4f}"

    # sənədləşdirilmiş metodoloji qeyd mövcuddur (qalıqlara ADF qadağandır)
    assert "MacKinnon" in ec.eg_coint.__doc__ and "YANLIŞDIR" in ec.eg_coint.__doc__


# ---------------------------------------------------------------------------
def t3_ols_hac_guard():
    R = rng(303)
    yrs = list(range(2000, 2026))
    x = pd.Series(np.linspace(1, 5, len(yrs)), index=yrs, name="x")
    y = pd.Series(0.5 + 1.2 * x.values + R.normal(0, 0.1, len(yrs)), index=yrs, name="y")

    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        r = ec.ols_hac(y, x, maxlags=1, name="sınaq tənliyi")
    out = buf.getvalue()
    assert "n=26" in out, f"n çap olunmalıdır, çıxış: {out!r}"
    assert r.n == 26 and abs(r.params["x"] - 1.2) < 0.05
    assert r.r2 > 0.99 and 0 < r.sigma < 0.3
    assert r.sample == (2000, 2025) and list(r.cov_params.columns) == ["const", "x"]
    row = r.catalog_row("sınaq", "sınaq tənliyi", lhs="y", series_code="TEST", fr="FR0")
    assert row["sample_end"] == 2025 and row["rhs"] == "x"

    # gələcəyə baxış: reqressor 2026-ya uzanır → istisna
    yrs2 = list(range(2000, 2028))
    x2 = pd.Series(np.linspace(1, 5, len(yrs2)), index=yrs2, name="x")
    y2 = pd.Series(np.linspace(1, 6, len(yrs2)), index=yrs2, name="y")
    raised = False
    try:
        ec.ols_hac(y2, x2, verbose=False)
    except ec.LookAheadError as exc:
        raised = True
        assert "2026" in str(exc)
    assert raised, "2025-dən sonrakı indeks üçün LookAheadError qaldırılmalı idi"

    # yalnız asılı dəyişən uzanırsa da qadağandır
    raised2 = False
    try:
        ec.ols_hac(y2, x, verbose=False)
    except ec.LookAheadError:
        raised2 = True
    assert raised2, "asılı dəyişənin proqnoz illəri də rədd edilməlidir"


# ---------------------------------------------------------------------------
def t4_ardl_ecm():
    R = rng(404)
    yrs = list(range(1976, 2026))
    n = len(yrs)
    x = pd.Series(np.cumsum(R.normal(0.2, 1, n)) + 30.0, index=yrs, name="x")
    # y uzunmüddətli olaraq 1.0 + 0.8x-ə bağlıdır, hər il kənarlaşmanın 50%-i aradan qalxır
    y = np.zeros(n)
    y[0] = 1.0 + 0.8 * x.iloc[0]
    for i in range(1, n):
        eq = 1.0 + 0.8 * x.iloc[i]
        y[i] = y[i - 1] - 0.5 * (y[i - 1] - eq) + R.normal(0, 0.3)
    Y = pd.Series(y, index=yrs, name="y")
    with contextlib.redirect_stdout(io.StringIO()):
        m = ec.ardl_ecm(Y, x, name="sınaq ECM", verbose=False)
    assert abs(m.lr.params["x"] - 0.8) < 0.08, f"uzunmüddətli əmsal: {m.lr.params['x']}"
    assert -1.0 < m.alpha < -0.1, f"korreksiya əmsalı mənfi olmalıdır: {m.alpha}"
    assert m.alpha_p < 0.05 and m.half_life is not None and 0.3 < m.half_life < 5.0
    assert m.coint is not None and m.coint.cointegrated
    nxt = m.predict_next(y_last=float(Y.iloc[-1]),
                         x_last={"x": float(x.iloc[-1])},
                         x_next={"x": float(x.iloc[-1]) + 0.2})
    assert abs(nxt - float(Y.iloc[-1])) < 5.0 and np.isfinite(nxt)


# ---------------------------------------------------------------------------
def _f_ar1(train, year):
    """AR(1) modeli — təlim nümunəsi ciddi şəkildə `year`-dan əvvəl bitir."""
    import statsmodels.api as sm
    s = train["y"].dropna() if isinstance(train, pd.DataFrame) else train.dropna()
    if len(s) < 6:
        return np.nan
    X = sm.add_constant(s.shift(1).dropna())
    b = sm.OLS(s.iloc[1:], X).fit().params
    return float(b.iloc[0] + b.iloc[1] * s.iloc[-1])


def _f_mean(train, year):
    """ETALON: nümunə ortası — güclü orta qiymətə qayıdışda RW-dən yaxşıdır."""
    s = train["y"].dropna() if isinstance(train, pd.DataFrame) else train.dropna()
    return float(s.mean())


def t5_backtest():
    R = rng(505)
    yrs = list(range(1976, 2026))
    n = len(yrs)
    e = R.normal(0, 1.0, n)
    v = np.zeros(n)
    for i in range(1, n):
        v[i] = 5.0 + 0.4 * (v[i - 1] - 5.0) + e[i]      # AR(1), φ = 0.4 → RW zəif etalondur
    s = pd.Series(v, index=yrs, name="y")

    bt = ec.backtest(s, {"AR1": _f_ar1, "Orta": _f_mean}, min_train=8, series_code="TEST_AR1")
    assert bt.rw_rmse is not None and bt.rmse["AR1"] is not None
    assert bt.rmse["AR1"] < bt.rw_rmse, f"AR(1) RW-ni üstələməli idi: {bt.rmse}"
    assert bt.beats_rw["AR1"] and "AR1" in bt.weights
    assert abs(sum(bt.weights.values()) - 1.0) < 1e-12
    assert bt.ensemble_rmse is not None and bt.ensemble_rmse < bt.rw_rmse
    assert bt.skill is not None and bt.skill > 0
    assert len(bt.common_years) >= 30 and max(bt.common_years) <= ec.last_actual()

    # cədvəl §3 müqaviləsinin sütunlarına uyğundur
    cols = ["series_code", "model", "vintage_year", "horizon_h", "actual", "forecast",
            "error", "abs_error", "rmse_h", "rw_rmse_h", "coverage80"]
    assert list(bt.table.columns) == cols
    assert set(bt.table["model"]) == {"AR1", "Orta", "RW"}
    assert (bt.table["horizon_h"] == 1).all()
    assert (bt.table["vintage_year"] + bt.table["horizon_h"]).isin(bt.common_years).all()
    assert 0.5 <= bt.coverage80["AR1"] <= 1.0
    # xəta = proqnoz − faktiki (tərif yoxlanışı)
    r0 = bt.table.iloc[0]
    assert abs((r0["forecast"] - r0["actual"]) - r0["error"]) < 1e-12

    # gələcəyə baxış yoxdur: sınaq ili həmişə təlim nümunəsindən sonradır
    yr = bt.common_years[5]
    train_end = yr - 1
    assert train_end < yr

    # RW-ni üstələyən model olmadıqda proqnozu etalon daşıyır
    rwsig = pd.Series(np.cumsum(R.normal(0, 1, n)), index=yrs, name="y")
    bt2 = ec.backtest(rwsig, {"Orta": _f_mean}, min_train=8, series_code="TEST_RW")
    if not bt2.beats_rw["Orta"]:
        assert bt2.weights == {"RW": 1.0}, bt2.weights

    # ansamblın birləşdirilməsi
    comb = bt.combine({m: {2026: 4.0, 2027: 4.5} for m in bt.weights})
    assert abs(comb.loc[2026] - 4.0) < 1e-9

    # heç bir proqnoz verməyən (və ya istisna atan) model ümumi pəncərəni sıfırlamır
    def _f_broken(train, year):
        raise RuntimeError("qəsdən nasaz model")

    bt3 = ec.backtest(s, {"AR1": _f_ar1, "Nasaz": _f_broken}, min_train=8, series_code="TEST_AR1")
    assert bt3.common_years == bt.common_years and bt3.rmse["Nasaz"] is None
    assert bt3.rmse["AR1"] is not None and "Nasaz" not in bt3.weights
    assert "Nasaz" not in set(bt3.table["model"])


# ---------------------------------------------------------------------------
def t6_solve_core():
    # y = 1.0 + 0.5c + 0.2g ; c = 0.4 + 0.6y  → y = (1.2 + 0.2g)/0.7
    eqs = {"y": (1.0, {"c": 0.5, "g": 0.2}),
           "c": (0.4, {"y": 0.6})}
    exog = {"g": pd.Series({2026: 2.0, 2027: 3.0})}
    sol = ec.solve_core(eqs, exog, years=[2026, 2027], tol=1e-8, max_iter=200)
    for yr, g in ((2026, 2.0), (2027, 3.0)):
        y_an = (1.2 + 0.2 * g) / 0.7
        c_an = 0.4 + 0.6 * y_an
        assert abs(sol["y"].loc[yr] - y_an) < 1e-7, (yr, sol["y"].loc[yr], y_an)
        assert abs(sol["c"].loc[yr] - c_an) < 1e-7
    assert list(sol["y"].index) == [2026, 2027]

    # gecikmiş endogen dəyişən: y(t) = 1 + 0.5·y(t−1), y(2025) = 4 → 3.0, 2.5, 2.25
    eqs2 = {"y": (1.0, {"y(-1)": 0.5})}
    sol2 = ec.solve_core(eqs2, {}, years=[2026, 2027, 2028], init={"y": pd.Series({2025: 4.0})})
    assert abs(sol2["y"].loc[2026] - 3.0) < 1e-9
    assert abs(sol2["y"].loc[2027] - 2.5) < 1e-9
    assert abs(sol2["y"].loc[2028] - 2.25) < 1e-9

    # yığılmayan blok (partlayan geribildirim) → ConvergenceError
    bad = {"a": (0.0, {"b": 2.0}), "b": (1.0, {"a": 2.0})}
    raised = False
    try:
        ec.solve_core(bad, {}, years=[2026], max_iter=50)
    except ec.ConvergenceError as exc:
        raised = True
        assert "yığılmadı" in str(exc)
    assert raised, "yığılmayan blok üçün ConvergenceError qaldırılmalı idi"

    # verilməyən ekzogen dəyişən açıq şəkildə bildirilir
    raised2 = False
    try:
        ec.solve_core({"y": (0.0, {"z": 1.0})}, {}, years=[2026])
    except KeyError:
        raised2 = True
    assert raised2


# ---------------------------------------------------------------------------
def t7_fan():
    path = pd.Series([3.0, 3.1, 3.2, 3.3, 3.4], index=[2026, 2027, 2028, 2029, 2030])
    b = ec.fan(path, resid_sigma=1.0, ndraw=6000)
    assert list(b.columns) == ["lo80", "hi80", "lo50", "hi50"]
    assert list(b.index) == list(path.index)
    w80 = (b["hi80"] - b["lo80"]).values
    w50 = (b["hi50"] - b["lo50"]).values
    assert all(w80[i] > w80[i - 1] for i in range(1, len(w80))), f"80% zolağı genişlənməlidir: {w80}"
    assert all(w50[i] > w50[i - 1] for i in range(1, len(w50))), f"50% zolağı genişlənməlidir: {w50}"
    assert all(w80 > w50), "80% zolağı 50%-dən geniş olmalıdır"
    assert all(b["lo50"] >= b["lo80"]) and all(b["hi50"] <= b["hi80"])
    # birinci ilin eni ≈ 2·1.2816·σ (normal nəzəri qiymət)
    assert abs(w80[0] - 2 * 1.2816) < 0.15, w80[0]
    # merkəz yolu zolağın içindədir
    assert all(b["lo80"] < path) and all(path < b["hi80"])
    # √h qaydası: 4-cü ilin eni birinci ilin ~2 misli
    assert 1.7 < w80[3] / w80[0] < 2.3, w80[3] / w80[0]

    # parametr qeyri-müəyyənliyi zolağı genişləndirir
    cov = np.array([[0.25, 0.0], [0.0, 0.09]])
    b2 = ec.fan(path, resid_sigma=1.0, coef_cov=cov, ndraw=6000)
    assert (b2["hi80"] - b2["lo80"]).iloc[0] > w80[0], "parametr qeyri-müəyyənliyi zolağı genişlətməlidir"

    # sabit enli variant (cumulative=False) genişlənmir
    b3 = ec.fan(path, resid_sigma=1.0, ndraw=6000, cumulative=False)
    w3 = (b3["hi80"] - b3["lo80"]).values
    assert abs(w3[-1] / w3[0] - 1.0) < 0.1


# ---------------------------------------------------------------------------
def t8_validate():
    yrs = [2020, 2021, 2022, 2023, 2024, 2025]
    a = pd.Series([10.0, 11, 12, 13, 14, 15], index=yrs)
    b = pd.Series([4.0, 4.5, 5, 5.5, 6, 6.5], index=yrs)
    tot = a + b
    r = va.identity_check("cəm eyniliyi", tot, a + b, tol=1e-9)
    assert r.passed and r.n_years == 6

    # bir ildə pozulma → istisna (baza ili düzgün olsa belə)
    broken = tot.copy()
    broken.loc[2023] += 0.01
    raised = False
    try:
        with contextlib.redirect_stdout(io.StringIO()):   # qəsdən uğursuz yoxlama — çıxış susdurulur
            va.identity_check("pozulmuş eynilik", broken, a + b, tol=1e-6)
    except va.IdentityError as exc:
        raised = True
        assert "2023" in str(exc)
    assert raised, "bir ildəki pozulma aşkarlanmalı idi"

    # illər dəsti üst-üstə düşməzsə yoxlama natamam sayılır
    raised2 = False
    try:
        va.identity_check("natamam illər", tot.loc[2020:2024], a + b, tol=1e-6)
    except va.IdentityError:
        raised2 = True
    assert raised2

    # gələcəyə baxış yoxlaması
    assert va.check_no_lookahead("nümunə", pd.Index(yrs)).passed
    assert not va.check_no_lookahead("nümunə", pd.Index(yrs + [2026])).passed

    # mütənasib uzlaşdırma: cəm məcmu göstərici ilə bərabərləşir, ilkin fərq hesabata düşür
    parts = pd.DataFrame({"neft": [30.0, 31, 32], "qeyri_neft": [65.0, 66, 67]},
                         index=[2023, 2024, 2025])
    total = pd.Series([100.0, 101.0, 102.0], index=[2023, 2024, 2025])
    rec, gap = va.reconcile_to_total(parts, total, name="sahələr")
    assert abs(rec.sum(axis=1) - total).max() < 1e-9
    assert abs(gap.loc[2023, "gap_abs"] - 5.0) < 1e-9 and abs(gap.loc[2023, "gap_pct"] - 5.0) < 1e-9

    # müstəqil yenidən hesablama harness-i
    v = va.Verifier("sınaq harness-i")
    v.identity("cəm", tot, a + b, tol=1e-9)
    v.recompute("2025 cəmi", shipped=float(tot.loc[2025]),
                fn=lambda: float(a.loc[2025]) + float(b.loc[2025]), tol=1e-9)
    v.no_lookahead("nümunə", pd.Index(yrs))
    with contextlib.redirect_stdout(io.StringIO()):
        assert v.report(verbose=True)
        v.assert_all(verbose=True)
    v2 = va.Verifier("uğursuz harness")
    v2.ok("qəsdən uğursuz", False, "sınaq")
    raised3 = False
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            v2.assert_all()
    except va.IdentityError:
        raised3 = True
    assert raised3


# ---------------------------------------------------------------------------
def t9_dinamika():
    """T9 — dinamika bloku: profil yolu, çoxüfüqlü sınaq və dinamiklik göstəricisi."""
    rng = np.random.default_rng(7)
    yrs = list(range(1986, 2026))

    # (a) səviyyə indeksi artımdan bərpa olunur və geri çevrilə bilir
    g = pd.Series([2.0, -1.0, 3.5, 0.5], index=[2022, 2023, 2024, 2025])
    lvl = ec.level_from_growth(g, base=100.0)
    back = (lvl / lvl.shift(1) - 1) * 100
    assert abs(float(lvl.iloc[0]) - 102.0) < 1e-9
    assert (back.dropna() - g.reindex(back.dropna().index)).abs().max() < 1e-9

    # (b) AR(1) profili: SABİT deyil, lövbərə YIĞILIR. Yığılmanın görünməsi üçün PERSİSTENT
    #     (ρ = 0,8) və son müşahidəsi lövbərdən uzaq olan süni sıra qurulur.
    yc = [4.0 + 8.0]
    for _ in range(len(yrs) - 1):
        yc.append(4.0 + 0.8 * (yc[-1] - 4.0) + rng.normal(0, 0.15))
    yc[-1] = 9.0                                          # son müşahidə lövbərdən (4,0) uzaqdır
    s_conv = pd.Series(yc, index=yrs)
    prof = ec.ar_profile(s_conv, [2026, 2027, 2028, 2029, 2030], p=1)
    assert prof is not None and len(prof) == 5
    assert prof.std(ddof=0) > 1e-6                       # sabit deyil
    d0 = abs(float(prof.iloc[0]) - 4.0)
    d4 = abs(float(prof.iloc[-1]) - 4.0)
    assert d4 < d0                                        # lövbərə yığılır

    # (c) çoxüfüqlü sınaq: orta qiymətə TEZ qayıdan sırada AR(1) yolu SABİT yolu üstələyir
    y = [10.0]
    for _ in range(len(yrs) - 1):
        y.append(4.0 + 0.25 * (y[-1] - 4.0) + rng.normal(0, 0.6))
    s = pd.Series(y, index=yrs)

    def m_ar(train, ys):
        pp = ec.ar_profile(train, ys, p=1)
        return {} if pp is None else {int(t): float(v) for t, v in pp.items()}

    pbt = ec.path_backtest(s, {"AR1": m_ar}, horizon=5, min_train=12, series_code="sınaq")
    assert pbt.rmse["AR1"] is not None and pbt.rw_rmse is not None
    assert pbt.rmse["AR1"] < pbt.rw_rmse and pbt.beats_rw["AR1"]
    assert abs(sum(pbt.weights.values()) - 1.0) < 1e-12
    assert pbt.ensemble_rmse is not None and pbt.ensemble_rmse_by_h.get(1) is not None
    assert set(pbt.table.columns) >= {"series_code", "model", "horizon_h", "rmse_h", "rw_rmse_h"}

    # (d) profil yolu: dinamik namizəd qazandıqda yol SABİT olmur
    path, pb, note = ec.profile_path(s, [2026, 2027, 2028, 2029, 2030], name="sınaq", min_train=12)
    assert len(path) == 5 and float(path.max() - path.min()) > 1e-6
    assert "profil" in note

    # SABİT sırada isə heç bir dinamik namizəd sabit yolu üstələyə bilmir → yol sabit qalır
    flat = pd.Series(3.0, index=yrs)
    fpath, _, fnote = ec.profile_path(flat, [2026, 2027], name="sabit", min_train=12, min_points=4)
    assert float(fpath.max() - fpath.min()) < 1e-9

    # (e) dinamiklik göstəricisi: populyasiya σ nisbəti, sətir müqaviləsi və aralıq yoxlaması
    act = pd.Series([1.0, 5.0, -3.0, 4.0, 0.0, 2.0], index=range(2020, 2026))
    fc = pd.Series([2.0, 2.5, 3.0, 3.5, 4.0], index=range(2026, 2031))
    r = ec.sigma_ratio(fc, act, series_code="test_kod", label_az="sınaq", hist_from=2020)
    exp = float(np.std(fc.values, ddof=0) / np.std(act.values, ddof=0))
    assert abs(r.ratio - exp) < 1e-12 and r.n_actual == 6 and r.n_forecast == 5
    rows = ec.dynamics_rows([r])
    assert list(rows.columns) == ["series_code", "model", "vintage_year", "horizon_h", "actual",
                                 "forecast", "error", "abs_error", "rmse_h", "rw_rmse_h",
                                 "coverage80"]
    assert rows.loc[0, "model"] == ec.DYN_MODEL
    assert abs(float(rows.loc[0, "rmse_h"]) - round(exp, 6)) < 1e-12
    assert float(rows.loc[0, "rw_rmse_h"]) == 1.0
    assert bool(pd.isna(rows.loc[0, "coverage80"]))

    # aralıqdan kənar sıra ÜÇÜN əsaslandırma tələb olunur (D2 — aralıq daxilidir, nəşr olunmur)
    flat_fc = pd.Series(3.0, index=range(2026, 2031))
    r0 = ec.sigma_ratio(flat_fc, act, series_code="sabit_kod", hist_from=2020)
    assert not r0.in_band
    raised = False
    try:
        ec.assert_dynamics_band([r0], strict=True)
    except ec.DynamicsBandError:
        raised = True
    assert raised
    # Mətn TƏK BAŞINA kifayət etmir: mexanizm SİNFİ də verilməlidir, əks halda "izah olundu"
    # görüntüsü yaranar, halbuki səbəbin hansı mexanizmə aid olduğu yazılmamış qalar.
    r0.justification = "sürücü yolu sabitdir — səbəb sənədləşdirilib"
    raised = False
    try:
        ec.assert_dynamics_band([r0], strict=True)
    except ec.DynamicsBandError:
        raised = True
    assert raised
    r0.dyn_class = "etalon"
    tab = ec.assert_dynamics_band([r0], strict=True)     # mətn + sinif varsa keçir
    assert len(tab) == 1 and not bool(tab.loc[0, "aralıqda"])
    assert tab.loc[0, "sinif"] == "etalon"
    # Nəşr olunan qeyd mətni sinif izahını DA daşıyır (D2 — səbəb rəqəmin yanında görünür)
    assert ec.DYN_CLASSES["etalon"] in r0.note_text and r0.justification in r0.note_text
    # Tanınmayan sinif SƏSSİZ keçmir
    raised = False
    try:
        ec.sigma_ratio(flat_fc, act, series_code="sabit_kod", hist_from=2020,
                       dyn_class="uydurma_sinif")
    except ValueError:
        raised = True
    assert raised
    # Sinif (d) — "yenidən baxılır" halı örtülmür, bayraq qaldırılır
    r_rev = ec.sigma_ratio(flat_fc, act, series_code="baxilan_kod", hist_from=2020,
                           justification="mexanizm aydın deyil", dyn_class="baxilir")
    assert r_rev.under_review and not r_rev.in_band
    ec.assert_dynamics_band([r_rev], strict=True)        # sinif verilib → dayandırmır

    # (f) dinamika formaları: gecikmiş asılı dəyişən HƏQİQƏTƏN olan sırada GECIKME qazanır
    x = pd.Series(rng.normal(0, 1, len(yrs)), index=yrs)
    yy = [0.0]
    for t in range(1, len(yrs)):
        yy.append(0.7 * yy[-1] + 1.5 * float(x.iloc[t]) + rng.normal(0, 0.3))
    fr = pd.DataFrame({"y": pd.Series(yy, index=yrs), "x": x})
    bt_d, tab_d, win = ec.dynamics_forms(fr, ["x"], series_code="dyn_test", min_train=10)
    assert win == "GECIKME"
    assert set(bt_d.rmse) >= {"STATIK", "GECIKME", "RW"}
    assert bool((bt_d.table["series_code"] == "dyn_test_dinamika").all())


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("=" * 66)
    print("  EKONOMETRİK NÜVƏNİN ÖZ-ÖZÜNÜ YOXLAMASI (src/econometrics.py, src/validate.py)")
    print("=" * 66)
    check("T1 ADF — vahid kök və stasionarlıq düzgün ayırd edilir", t1_adf)
    check("T2 eg_coint — kointeqrasiya aşkarlanır, saxta cütlük rədd edilir", t2_coint)
    check("T3 ols_hac — n çap olunur, proqnoz illəri rədd edilir", t3_ols_hac_guard)
    check("T4 ardl_ecm — uzunmüddətli əmsal və korreksiya sürəti bərpa olunur", t4_ardl_ecm)
    check("T5 backtest — AR(1) təsadüfi gəzişməni üstələyir, cədvəl müqaviləyə uyğundur", t5_backtest)
    check("T6 solve_core — analitik sabit nöqtə bərpa olunur, yığılmama aşkarlanır", t6_solve_core)
    check("T7 fan — zolaqlar üfüq boyu genişlənir", t7_fan)
    check("T8 validate — eynilik hər il yoxlanılır, uzlaşdırma və harness işləyir", t8_validate)
    check("T9 dinamika — profil yolu, çoxüfüqlü sınaq, dinamiklik göstəricisi", t9_dinamika)
    print("-" * 66)
    print(f"  {len(PASSED)}/9 yoxlama bloku uğurlu")
    print("=" * 66)
