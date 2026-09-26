"""econometrics.py — MİİS §15.5.1 paketinin ekonometrik nüvəsi.

Bu modul bütün FR notebook-larının istifadə etdiyi YEGANƏ qiymətləndirmə qatıdır: vahid kök
sınağı (ADF), kointeqrasiya (Engle–Granger, MacKinnon kritik qiymətləri), HAC (Newey–West)
standart xətaları ilə OLS, iki addımlı xəta korreksiyası modeli (ECM), genişlənən pəncərəli
nümunədən kənar (out-of-sample) geriyə doğru sınaq, yelpik zolaqları və simultan nüvənin
sabit nöqtə həlli.

Metodoloji qərarlar (BUILD_CONTEXT.md §D4–D5-ə uyğun, dəyişdirilmir):
  1. Nümunə **LAST_ACTUAL** ilində bitir; `ols_hac` proqnoz üfüqünə (2026–2030) aid indeksi
     olan hər hansı reqressoru RƏDD EDİR — bu, gələcəyə baxışın (look-ahead) struktur mühafizəsidir.
  2. Kointeqrasiya yalnız `statsmodels.tsa.stattools.coint` vasitəsilə yoxlanılır; qalıqlar
     üzərində adi ADF p-qiyməti heç bir yerdə istifadə olunmur (səbəb `eg_coint` sənədində).
  3. Hər davranış tənliyi təsadüfi gəzişmə (RW) etalonuna qarşı sınaqdan keçirilir; RW-ni
     üstələməyən tənlik ya ekzogen/etalon yanaşma ilə əvəz olunur, ya da nişanlanır və
     proqnozu ansambl daşıyır.
  4. Yelpik zolaqları qalıq dispersiyası İLƏ BİRLİKDƏ parametr qeyri-müəyyənliyindən
     (əmsal kovariasiyasından çoxölçülü normal seçmə) qurulur və üfüq boyu genişlənir.
"""
from __future__ import annotations

import math
import re
import warnings
from dataclasses import dataclass, field
from typing import Callable, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple, Union

import numpy as np
import pandas as pd
import statsmodels.api as sm
from statsmodels.tsa.stattools import adfuller, coint

__all__ = [
    "LookAheadError", "ConvergenceError", "DynamicsBandError",
    "ADFResult", "CointResult", "RegResult", "ECMResult", "BacktestResult",
    "PathBacktestResult", "DynamicsResult",
    "adf", "eg_coint", "ols_hac", "ardl_ecm", "backtest", "fan", "solve_core",
    "assert_no_lookahead", "last_actual", "years_of", "SEED",
    "level_from_growth", "ar_profile", "path_backtest", "profile_path",
    "sigma_ratio", "dynamics_rows", "assert_dynamics_band",
    "DYN_BAND_LO", "DYN_BAND_HI", "DYN_MODEL", "DYN_HIST_START", "DYN_CLASSES",
    "SPEC_MODEL", "NON_ENSEMBLE_MODELS",
]

# Bütün təsadüfi seçmələr üçün vahid toxum — nəticələr təkrar istehsal olunandır.
SEED = 20260605

_Z80 = 1.2815515655446004   # standart normalın 90-cı faizliyi → 80% iki tərəfli zolaq
_Z50 = 0.6744897501960817   # 75-ci faizlik → 50% iki tərəfli zolaq

_FALLBACK_LAST_ACTUAL = 2025   # config.py hələ yoxdursa müvəqqəti dəyər (§7-də bərkidilib)
_LAST_ACTUAL_CACHE: Optional[int] = None


class LookAheadError(ValueError):
    """Nümunəyə proqnoz üfüqünə aid il düşdükdə qaldırılır (gələcəyə baxış mühafizəsi)."""


class ConvergenceError(RuntimeError):
    """Sabit nöqtə iterasiyası verilmiş dəqiqliklə yığılmadıqda qaldırılır."""


class DynamicsBandError(AssertionError):
    """Dinamiklik göstəricisi DAXİLİ qəbul aralığından kənara çıxdıqda və sıra üçün yazılı
    əsaslandırma verilmədikdə qaldırılır (bax `assert_dynamics_band`)."""


# ---------------------------------------------------------------------------
# Köməkçi funksiyalar
# ---------------------------------------------------------------------------
def last_actual() -> int:
    """Son faktiki il. config.py mövcud olduqda ondan oxunur (vahid mənbə prinsipi)."""
    global _LAST_ACTUAL_CACHE
    if _LAST_ACTUAL_CACHE is not None:
        return _LAST_ACTUAL_CACHE
    for _imp in ("relative", "flat"):
        try:
            if _imp == "relative":
                from . import config as _cfg      # type: ignore
            else:
                import config as _cfg             # type: ignore
            _LAST_ACTUAL_CACHE = int(_cfg.LAST_ACTUAL)
            return _LAST_ACTUAL_CACHE
        except Exception:
            continue
    return _FALLBACK_LAST_ACTUAL   # config yaranana qədər §7-də bərkidilmiş dəyər


def years_of(index) -> np.ndarray:
    """İndeksdən il nömrələrini çıxarır (tam ədəd, tarix və ya dövr indeksi qəbul olunur)."""
    idx = pd.Index(index)
    if isinstance(idx, pd.DatetimeIndex) or isinstance(idx, pd.PeriodIndex):
        return np.asarray(idx.year, dtype=int)
    if pd.api.types.is_integer_dtype(idx) or pd.api.types.is_float_dtype(idx):
        return np.asarray(idx, dtype=float).astype(int)
    try:
        return np.asarray(pd.to_datetime(idx).year, dtype=int)
    except Exception as exc:   # indeks il kimi şərh oluna bilmir
        raise ValueError(
            "İndeks il kimi şərh oluna bilmir; illik indeks (int və ya tarix) tələb olunur."
        ) from exc


def assert_no_lookahead(obj, name: str = "nümunə", limit: Optional[int] = None) -> None:
    """Struktur mühafizə: indeksdə LAST_ACTUAL-dan sonrakı il varsa, qiymətləndirməni dayandırır."""
    lim = last_actual() if limit is None else int(limit)
    yrs = years_of(getattr(obj, "index", obj))
    if yrs.size and int(yrs.max()) > lim:
        bad = sorted(int(y) for y in np.unique(yrs[yrs > lim]))
        raise LookAheadError(
            f"{name}: nümunə {lim}-dən sonrakı illəri əhatə edir ({bad}). "
            f"Qiymətləndirmə nümunəsi {lim} ilində bitməlidir — proqnoz dəyərləri reqressiyaya "
            f"heç vaxt daxil edilmir (BUILD_CONTEXT §D5)."
        )


def _as_frame(x, prefix: str = "x") -> pd.DataFrame:
    """Seriya və ya cədvəli DataFrame-ə çevirir, adsız sütunlara ad verir."""
    if isinstance(x, pd.DataFrame):
        return x.copy()
    if isinstance(x, pd.Series):
        s = x.copy()
        if s.name is None:
            s.name = prefix
        return s.to_frame()
    arr = np.asarray(x)
    if arr.ndim == 1:
        return pd.DataFrame({prefix: arr})
    return pd.DataFrame(arr, columns=[f"{prefix}{i+1}" for i in range(arr.shape[1])])


def _fmt(v, nd: int = 4) -> str:
    return "n/a" if v is None or (isinstance(v, float) and not np.isfinite(v)) else f"{v:.{nd}f}"


# ---------------------------------------------------------------------------
# 1. ADF — vahid kök sınağı
# ---------------------------------------------------------------------------
@dataclass
class ADFResult:
    """ADF sınağının yığcam nəticəsi."""
    name: str
    stat: float
    pvalue: float
    usedlag: int
    nobs: int
    crit: Dict[str, float]
    regression: str
    alpha: float

    @property
    def stationary(self) -> bool:
        """Verilmiş əhəmiyyət səviyyəsində vahid kök hipotezi rədd olunursa True."""
        return bool(self.pvalue < self.alpha)

    @property
    def note_az(self) -> str:
        """Notebook mətninə birbaşa yazıla bilən bir cümləlik AZ şərh."""
        if self.stationary:
            return (f"ADF ({self.name}): vahid kök rədd edilir (p = {self.pvalue:.3f}, "
                    f"n = {self.nobs}, gecikmə = {self.usedlag}) → stasionar.")
        return (f"ADF ({self.name}): vahid kök rədd edilmir (p = {self.pvalue:.3f}, "
                f"n = {self.nobs}, gecikmə = {self.usedlag}) → səviyyədə qeyri-stasionar.")

    def __str__(self) -> str:   # yığcam çap
        crit = ", ".join(f"{k}: {v:.2f}" for k, v in sorted(self.crit.items()))
        return (f"ADF[{self.name}] stat={self.stat:.3f} p={self.pvalue:.4f} "
                f"lag={self.usedlag} n={self.nobs} kritik({crit}) "
                f"→ {'stasionar' if self.stationary else 'vahid kök'}")


def adf(series, name: Optional[str] = None, regression: str = "c",
        maxlag: Optional[int] = None, alpha: float = 0.05) -> ADFResult:
    """Augmented Dickey–Fuller sınağı (gecikmə seçimi AIC ilə avtomatikdir).

    Səviyyə seriyası üçün `regression="c"` (sabitli), trendli seriya üçün `"ct"` tətbiq olunur.
    Nəticə obyekti həm p-qiymətini, həm istifadə olunan gecikməni, həm də müşahidə sayını
    daşıyır — hesabatda n həmişə göstərilir.
    """
    s = pd.Series(series).dropna().astype(float)
    if len(s) < 6:
        raise ValueError(f"ADF üçün müşahidə sayı kifayət deyil (n = {len(s)}).")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        stat, pval, usedlag, nobs, crit, _icbest = adfuller(
            s.values, maxlag=maxlag, regression=regression, autolag="AIC"
        )
    return ADFResult(
        name=name or (s.name if s.name is not None else "seriya"),
        stat=float(stat), pvalue=float(pval), usedlag=int(usedlag), nobs=int(nobs),
        crit={k: float(v) for k, v in crit.items()}, regression=regression, alpha=float(alpha),
    )


# ---------------------------------------------------------------------------
# 2. Engle–Granger kointeqrasiya
# ---------------------------------------------------------------------------
@dataclass
class CointResult:
    """Engle–Granger kointeqrasiya sınağının nəticəsi (MacKinnon kritik qiymətləri)."""
    name: str
    stat: float
    pvalue: float
    crit: Dict[str, float]
    nobs: int
    k_regressors: int
    trend: str
    alpha: float

    @property
    def cointegrated(self) -> bool:
        return bool(self.pvalue < self.alpha)

    @property
    def note_az(self) -> str:
        if self.cointegrated:
            return (f"Kointeqrasiya ({self.name}): EG statistikası {self.stat:.3f}, "
                    f"p = {self.pvalue:.3f} (n = {self.nobs}) → uzunmüddətli əlaqə təsdiqlənir, "
                    f"ECM tətbiq olunur.")
        return (f"Kointeqrasiya ({self.name}): EG statistikası {self.stat:.3f}, "
                f"p = {self.pvalue:.3f} (n = {self.nobs}) → uzunmüddətli əlaqə təsdiqlənmir, "
                f"tənlik fərqlər (Δ) səviyyəsində qurulur.")

    def __str__(self) -> str:
        crit = ", ".join(f"{k}: {v:.2f}" for k, v in sorted(self.crit.items()))
        return (f"EG-coint[{self.name}] stat={self.stat:.3f} p={self.pvalue:.4f} "
                f"n={self.nobs} k={self.k_regressors} kritik({crit}) "
                f"→ {'kointeqrasiya var' if self.cointegrated else 'kointeqrasiya yoxdur'}")


def eg_coint(y, x, name: Optional[str] = None, trend: str = "c",
             maxlag: Optional[int] = None, alpha: float = 0.05) -> CointResult:
    """Engle–Granger kointeqrasiya sınağı — YALNIZ `statsmodels.tsa.stattools.coint` vasitəsilə.

    **Metodoloji qeyd.**
    Kointeqrasiya sınağını "əvvəlcə səviyyələr üzrə OLS qur, sonra qalıqlara adi ADF tətbiq et və
    ADF-in p-qiymətini oxu" ardıcıllığı ilə aparmaq YANLIŞDIR. Səbəbi budur: qalıqlar müşahidə
    olunmuş kəmiyyət deyil, OLS tərəfindən **qiymətləndirilmiş** kointeqrasiya vektorunun
    nəticəsidir; OLS isə məhz qalıq dispersiyasını minimuma endirir. Buna görə qalıqlar süni
    şəkildə "stasionara oxşar" görünür və adi Dickey–Fuller cədvəlləri (Hi hipotezini) həddindən
    artıq tez-tez rədd edir. Düzgün sınaq üçün kritik qiymətlər reqressorların sayından və
    deterministik hədlərdən asılı olan **MacKinnon (1991, 2010) kointeqrasiya cədvəllərindən**
    götürülməlidir. `coint()` funksiyası genişləndirilmiş Engle–Granger (aeg) prosedurunu məhz bu
    cədvəllərlə həyata keçirir; buna görə paketdə kointeqrasiya yalnız bu funksiya ilə yoxlanılır və
    qalıqlar üzərində adi ADF p-qiyməti heç bir yerdə hesabat dəyəri kimi istifadə edilmir.

    Parametrlər: `y` — asılı seriya, `x` — bir və ya bir neçə reqressor (Series/DataFrame).
    Sıfır hipotezi: kointeqrasiya YOXDUR. p < alpha olduqda uzunmüddətli əlaqə qəbul edilir.
    """
    Y = pd.Series(y).astype(float)
    X = _as_frame(x, prefix="x").astype(float)
    df = pd.concat([Y.rename("y"), X], axis=1).dropna()
    if len(df) < 12:
        raise ValueError(f"Kointeqrasiya sınağı üçün müşahidə sayı kifayət deyil (n = {len(df)}).")
    assert_no_lookahead(df, name or "kointeqrasiya nümunəsi")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        stat, pval, crit = coint(
            df["y"].values, df.drop(columns=["y"]).values,
            trend=trend, method="aeg", maxlag=maxlag, autolag="aic",
        )
    labels = ["1%", "5%", "10%"]
    return CointResult(
        name=name or f"{Y.name or 'y'} ~ {', '.join(map(str, X.columns))}",
        stat=float(stat), pvalue=float(pval),
        crit={labels[i]: float(c) for i, c in enumerate(np.atleast_1d(crit))},
        nobs=int(len(df)), k_regressors=int(df.shape[1] - 1), trend=trend, alpha=float(alpha),
    )


# ---------------------------------------------------------------------------
# 3. OLS + HAC (Newey–West)
# ---------------------------------------------------------------------------
@dataclass
class RegResult:
    """OLS + HAC nəticəsi: n, əmsallar, HAC standart xətaları, t, p, R², düzəldilmiş R²."""
    name: str
    n: int
    params: pd.Series
    se: pd.Series
    tvalues: pd.Series
    pvalues: pd.Series
    r2: float
    adj_r2: float
    sigma: float                 # reqressiyanın standart xətası (se_regression)
    resid: pd.Series
    fitted: pd.Series
    cov_params: pd.DataFrame     # HAC kovariasiya matrisi — yelpik zolaqları üçün
    exog_names: List[str]
    sample: Tuple[int, int]
    maxlags: int
    fit: object = field(repr=False, default=None)

    def predict(self, Xnew) -> pd.Series:
        """Yeni dizayn matrisi üzrə proqnoz (sabit avtomatik əlavə olunur, əgər modeldə varsa)."""
        Xn = _as_frame(Xnew, prefix="x").astype(float)
        if "const" in self.exog_names and "const" not in Xn.columns:
            Xn = sm.add_constant(Xn, has_constant="add")
        missing = [c for c in self.exog_names if c not in Xn.columns]
        if missing:
            raise ValueError(f"Proqnoz üçün dəyişənlər çatışmır: {missing}")
        Xn = Xn.reindex(columns=self.exog_names)
        return pd.Series(Xn.values @ self.params.reindex(self.exog_names).values, index=Xn.index)

    @property
    def table(self) -> pd.DataFrame:
        """Əmsal cədvəli — notebook-da birbaşa göstərilə bilər."""
        return pd.DataFrame({
            "əmsal": self.params, "HAC s.x.": self.se,
            "t": self.tvalues, "p": self.pvalues,
        })

    def catalog_row(self, eq_name: str, description: str, lhs: str, series_code: str = "",
                    fr: str = "", workbook: str = "", sheet: str = "",
                    rhs: Optional[str] = None) -> Dict[str, object]:
        """`outputs/equations_catalog.csv` üçün bir sətir (§3 müqaviləsinin sütunları).

        `rhs` verilmədikdə reqressorların adlarından qurulur. Adın özü izahlı olmadıqda
        (məsələn `absorbsiya` kimi mürəkkəb göstərici) tənliyi oxuyanın dəyişənin nədən
        ibarət olduğunu görməsi üçün AÇIQ sətir verilir.
        """
        if rhs is None:
            rhs = " + ".join(n for n in self.exog_names if n != "const")
        return {
            "workbook": workbook, "sheet": sheet, "eq_name": eq_name,
            "description": description, "lhs": lhs, "rhs": rhs,
            "series_code": series_code, "fr": fr,
            "sample_start": self.sample[0], "sample_end": self.sample[1],
            "adj_r2": round(self.adj_r2, 4), "se_regression": round(self.sigma, 4),
        }

    def __str__(self) -> str:
        head = (f"OLS+HAC[{self.name}]  n={self.n}  nümunə={self.sample[0]}–{self.sample[1]}  "
                f"HAC gecikmə={self.maxlags}\n"
                f"  R²={_fmt(self.r2,3)}  düzəldilmiş R²={_fmt(self.adj_r2,3)}  "
                f"reqressiyanın s.x.={_fmt(self.sigma,4)}")
        lines = [head, f"  {'dəyişən':<16}{'əmsal':>12}{'HAC s.x.':>12}{'t':>9}{'p':>9}"]
        for nm in self.exog_names:
            lines.append(f"  {nm:<16}{self.params[nm]:>12.4f}{self.se[nm]:>12.4f}"
                         f"{self.tvalues[nm]:>9.2f}{self.pvalues[nm]:>9.3f}")
        return "\n".join(lines)


def ols_hac(y, X, maxlags: int = 1, add_const: bool = True, name: Optional[str] = None,
            verbose: bool = True) -> RegResult:
    """OLS + Newey–West (HAC) standart xətaları. Müşahidə sayı n HƏMİŞƏ çap olunur.

    **Gələcəyə baxış mühafizəsi (struktur).** Əgər `y` və ya hər hansı reqressorun indeksi
    `config.LAST_ACTUAL`-dan sonrakı ili əhatə edirsə, funksiya `LookAheadError` qaldırır və
    qiymətləndirmə aparılmır. Beləliklə proqnoz üfüqünün (2026–2030) dəyərləri heç bir
    reqressiyaya daxil ola bilmir — bu, notebook səviyyəsində kəsmə əvəzinə mərkəzləşdirilmiş
    və pozula bilməyən qaydadır.

    Parametrlər: `maxlags` — HAC pəncərəsinin gecikmə sayı (illik məlumatda adətən 1–2).
    """
    Y = pd.Series(y).astype(float)
    Xf = _as_frame(X, prefix="x").astype(float)
    nm = name or (Y.name if Y.name is not None else "tənlik")

    # struktur mühafizə — həm asılı dəyişən, həm reqressorlar üçün
    assert_no_lookahead(Xf, f"{nm} (reqressorlar)")
    assert_no_lookahead(Y, f"{nm} (asılı dəyişən)")

    df = pd.concat([Y.rename("__y__"), Xf], axis=1).dropna()
    if df.empty:
        raise ValueError(f"{nm}: kəsişən müşahidə yoxdur.")
    Yc = df["__y__"]
    Xc = df.drop(columns=["__y__"])
    if add_const and "const" not in Xc.columns:
        Xc = sm.add_constant(Xc, has_constant="add")
        Xc = Xc[["const"] + [c for c in Xc.columns if c != "const"]]
    if len(df) <= Xc.shape[1]:
        raise ValueError(f"{nm}: sərbəstlik dərəcəsi yoxdur (n = {len(df)}, k = {Xc.shape[1]}).")

    fit = sm.OLS(Yc, Xc).fit(cov_type="HAC",
                             cov_kwds={"maxlags": int(maxlags), "use_correction": True})
    yrs = years_of(df.index)
    res = RegResult(
        name=nm, n=int(fit.nobs), params=fit.params, se=fit.bse,
        tvalues=fit.tvalues, pvalues=fit.pvalues,
        r2=float(fit.rsquared), adj_r2=float(fit.rsquared_adj),
        sigma=float(np.sqrt(fit.mse_resid)),
        resid=pd.Series(fit.resid, index=df.index),
        fitted=pd.Series(fit.fittedvalues, index=df.index),
        cov_params=pd.DataFrame(np.asarray(fit.cov_params()),
                                index=list(Xc.columns), columns=list(Xc.columns)),
        exog_names=list(Xc.columns), sample=(int(yrs.min()), int(yrs.max())),
        maxlags=int(maxlags), fit=fit,
    )
    if verbose:   # n həmişə çap olunur — hesabat intizamı (BUILD_CONTEXT §D5)
        print(f"[OLS+HAC] {nm}: n={res.n}, nümunə={res.sample[0]}–{res.sample[1]}, "
              f"R²={res.r2:.3f}, düzəldilmiş R²={res.adj_r2:.3f}, s.x.={res.sigma:.4f}")
    return res


# ---------------------------------------------------------------------------
# 4. İki addımlı Engle–Granger / ARDL-ECM
# ---------------------------------------------------------------------------
@dataclass
class ECMResult:
    """İki addımlı Engle–Granger nəticəsi: uzunmüddətli əlaqə + qısamüddətli korreksiya."""
    name: str
    lr: RegResult                 # 1-ci addım: səviyyələr üzrə uzunmüddətli tənlik
    sr: RegResult                 # 2-ci addım: fərqlər üzrə ECM
    alpha: float                  # korreksiya əmsalı (mənfi olmalıdır)
    alpha_t: float
    alpha_p: float
    ec: pd.Series                 # tarazlıqdan kənarlaşma (error-correction termi)
    coint: Optional[CointResult]
    x_names: List[str]

    @property
    def half_life(self) -> Optional[float]:
        """Tarazlığa qayıdışın yarım-ömrü (il). Korreksiya əmsalı düzgün işarədə deyilsə None."""
        if -1.0 < self.alpha < 0.0:
            return float(math.log(0.5) / math.log(1.0 + self.alpha))
        return None

    @property
    def note_az(self) -> str:
        hl = self.half_life
        hl_txt = f", yarım-ömür ≈ {hl:.1f} il" if hl is not None else ""
        return (f"ECM ({self.name}): korreksiya əmsalı α = {self.alpha:.3f} "
                f"(t = {self.alpha_t:.2f}, p = {self.alpha_p:.3f}){hl_txt}; "
                f"uzunmüddətli əmsallar: "
                + ", ".join(f"{k} = {self.lr.params[k]:.3f}" for k in self.x_names) + ".")

    def predict_next(self, y_last: float, x_last: Mapping[str, float],
                     x_next: Mapping[str, float], dy_last: Optional[float] = None) -> float:
        """Bir addımlıq proqnoz: ECM-in bərpa etdiyi Δy-i son səviyyəyə əlavə edir."""
        xl = pd.Series({k: float(x_last[k]) for k in self.x_names})
        xn = pd.Series({k: float(x_next[k]) for k in self.x_names})
        lr_hat = self.lr.params.get("const", 0.0) + float((self.lr.params.reindex(self.x_names) * xl).sum())
        ec_last = float(y_last - lr_hat)
        b = self.sr.params
        dy = float(b.get("const", 0.0)) + float(b.get("ec_l1", 0.0)) * ec_last
        for k in self.x_names:
            dy += float(b.get(f"d_{k}", 0.0)) * float(xn[k] - xl[k])
        if "dy_l1" in b.index and dy_last is not None:
            dy += float(b["dy_l1"]) * float(dy_last)
        return float(y_last + dy)


def ardl_ecm(y, x, name: Optional[str] = None, maxlags: int = 1, add_dy_lag: bool = False,
             check_coint: bool = True, verbose: bool = True) -> ECMResult:
    """İki addımlı Engle–Granger prosedurası (uzunmüddətli səviyyə tənliyi + ECM).

    1-ci addım — səviyyələr üzrə uzunmüddətli əlaqə: y = β₀ + β′x + u.
    2-ci addım — fərqlər üzrə qısamüddətli tənlik: Δy = c + α·u₍t−1₎ + γ′Δx (+ δ·Δy₍t−1₎).

    α korreksiya əmsalı mənfi və statistik əhəmiyyətli olmalıdır: tarazlıqdan kənarlaşmanın
    hər il hansı payının aradan qalxdığını göstərir. Bu spesifikasiya YALNIZ `eg_coint`
    kointeqrasiyanı təsdiqlədikdə istifadə edilir; əks halda tənlik saf fərqlər səviyyəsində
    qurulmalıdır (`check_coint=True` olduqda funksiya bunu yoxlayır və nəticəni daşıyır).
    """
    Y = pd.Series(y).astype(float)
    X = _as_frame(x, prefix="x").astype(float)
    nm = name or f"{Y.name or 'y'} ~ {', '.join(map(str, X.columns))}"
    assert_no_lookahead(X, f"{nm} (reqressorlar)")
    assert_no_lookahead(Y, f"{nm} (asılı dəyişən)")

    cr = None
    if check_coint:
        cr = eg_coint(Y, X, name=nm)
        if verbose:
            print(f"[ECM] {cr.note_az}")

    lr = ols_hac(Y, X, maxlags=maxlags, name=f"{nm} — uzunmüddətli", verbose=verbose)
    ec = (Y - lr.fitted).dropna().rename("ec")

    d = pd.DataFrame({"dy": Y.diff()})
    for c in X.columns:
        d[f"d_{c}"] = X[c].diff()
    d["ec_l1"] = ec.shift(1)
    if add_dy_lag:
        d["dy_l1"] = Y.diff().shift(1)
    d = d.dropna()
    rhs_cols = [c for c in d.columns if c != "dy"]
    sr = ols_hac(d["dy"], d[rhs_cols], maxlags=maxlags, name=f"{nm} — ECM", verbose=verbose)

    return ECMResult(
        name=nm, lr=lr, sr=sr,
        alpha=float(sr.params["ec_l1"]), alpha_t=float(sr.tvalues["ec_l1"]),
        alpha_p=float(sr.pvalues["ec_l1"]), ec=ec, coint=cr, x_names=list(X.columns),
    )


# ---------------------------------------------------------------------------
# 5. Geriyə doğru sınaq (genişlənən pəncərə, bir addımlıq, nümunədən kənar)
# ---------------------------------------------------------------------------
@dataclass
class BacktestResult:
    """Genişlənən pəncərəli OOS sınağın nəticəsi (`az_forecast_model/src/build_ensemble.py`-dan portlanıb)."""
    series_code: str
    rmse: Dict[str, Optional[float]]
    rw_rmse: Optional[float]
    weights: Dict[str, float]
    table: pd.DataFrame
    errors: Dict[str, Dict[int, float]]
    coverage80: Dict[str, Optional[float]]
    common_years: List[int]
    ensemble_rmse: Optional[float]
    beats_rw: Dict[str, bool]
    horizon_h: int = 1

    @property
    def skill(self) -> Optional[float]:
        """Ansamblın RW-yə nisbətdə bacarığı: 1 − RMSE(ansambl)/RMSE(RW). Müsbət = yaxşı."""
        if self.ensemble_rmse is None or not self.rw_rmse:
            return None
        return float(1.0 - self.ensemble_rmse / self.rw_rmse)

    @property
    def note_az(self) -> str:
        if not self.common_years:
            return (f"Geriyə doğru sınaq ({self.series_code}): ümumi sınaq pəncərəsi boşdur — "
                    f"nümunə qısadır, nəticə hesabat dəyəri daşımır.")
        kept = ", ".join(f"{m} (çəki {w:.2f})" for m, w in self.weights.items()) or "yoxdur"
        sk = self.skill
        sk_txt = f"; bacarıq = {sk*100:.0f}%" if sk is not None else ""
        return (f"Geriyə doğru sınaq ({self.series_code}): ümumi pəncərə "
                f"{self.common_years[0]}–{self.common_years[-1]} ({len(self.common_years)} il), "
                f"RW RMSE = {_fmt(self.rw_rmse, 3)}, ansambl RMSE = {_fmt(self.ensemble_rmse, 3)}"
                f"{sk_txt}. Saxlanılan modellər: {kept}.")

    def combine(self, paths: Mapping[str, Mapping[int, float]]) -> pd.Series:
        """Saxlanılan modellərin proqnoz yollarını tərs-RMSE² çəkiləri ilə birləşdirir."""
        missing = [m for m in self.weights if m not in paths]
        if missing:
            raise KeyError(f"Ansambl üçün yol verilməyən modellər: {missing}")
        years = sorted(set.intersection(*[set(paths[m].keys()) for m in self.weights]))
        return pd.Series({t: sum(self.weights[m] * float(paths[m][t]) for m in self.weights)
                          for t in years})

    def __str__(self) -> str:
        return self.note_az


def _rw_forecast(train, target_year):   # ETALON: təsadüfi gəzişmə — sonuncu müşahidə
    s = train["y"] if isinstance(train, pd.DataFrame) else train
    return float(s.dropna().iloc[-1])


def backtest(y, models: Mapping[str, Callable], min_train: int = 8,
             series_code: str = "", test_years: Optional[Iterable[int]] = None,
             min_points: int = 4, target_col: str = "y",
             verbose: bool = False) -> BacktestResult:
    """Genişlənən pəncərəli, bir addımlıq, nümunədən kənar geriyə doğru sınaq.

    Hər sınaq ili t üçün model YALNIZ t-dən əvvəlki illərin məlumatı ilə qiymətləndirilir
    (gələcəyə baxış yoxdur), t ili proqnozlaşdırılır və xəta qeyd olunur. Bütün modellər EYNİ
    (ümumi) sınaq pəncərəsində — hamısının proqnoz verdiyi illərdə — qiymətləndirilir ki, RMSE
    müqayisəsi ədalətli olsun.

    `models` — {ad: funksiya} lüğəti; funksiya `fn(train, target_year)` imzasına malikdir və
    `train` t-dən əvvəl bitən seriya (və ya çoxdəyişənli cədvəl) olur. "RW" adlı etalon avtomatik
    əlavə edilir (verilməyibsə).

    Modellərin seçimi: RW-ni üstələməyən model kənarlaşdırılır; sağ qalanlar tərs-RMSE² (dəqiqlik)
    çəkiləri ilə birləşdirilir. Heç bir model RW-ni üstələmirsə, çəkilərin hamısı RW-yə verilir —
    yəni proqnozu ekzogen/etalon yanaşma daşıyır (BUILD_CONTEXT §D4).

    `.table` — `outputs/validation_backtest.csv` sətirləri. `vintage_year` təlim nümunəsinin
    bitdiyi (informasiya) ilidir; proqnozlaşdırılan il = vintage_year + horizon_h.
    `coverage80` modelin öz OOS RMSE-sindən qurulmuş ±1.2816·RMSE zolağının faktiki əhatəsidir.
    """
    frame = y.copy() if isinstance(y, pd.DataFrame) else pd.DataFrame({"y": pd.Series(y).astype(float)})
    if target_col != "y":
        frame = frame.rename(columns={target_col: "y"})
    if "y" not in frame.columns:
        raise KeyError(f"Hədəf sütunu tapılmadı: '{target_col}'.")
    frame = frame.copy()
    frame.index = years_of(frame.index)
    frame = frame.sort_index()

    mdl: Dict[str, Callable] = dict(models)
    mdl.setdefault("RW", _rw_forecast)

    lim = last_actual()
    yv = frame["y"].dropna()
    if test_years is None:
        cand = [int(t) for t in yv.index if t <= lim]
        cand = cand[min_train:]           # ilk min_train il yalnız təlim üçün istifadə olunur
    else:
        cand = [int(t) for t in test_years if t <= lim]

    errs: Dict[str, Dict[int, float]] = {m: {} for m in mdl}
    preds: Dict[str, Dict[int, float]] = {m: {} for m in mdl}
    for t in cand:
        if t not in frame.index:
            continue
        train = frame.loc[: t - 1]        # ciddi şəkildə t-dən əvvəl — gələcəyə baxış yoxdur
        if len(train["y"].dropna()) < min_train:
            continue
        actual = frame.loc[t, "y"]
        if pd.isna(actual):
            continue
        for nm, fn in mdl.items():
            try:
                p = fn(train, t)
            except Exception:
                p = np.nan
            if p is None or pd.isna(p):
                continue
            preds[nm][t] = float(p)
            errs[nm][t] = float(p) - float(actual)

    # Ümumi (common) pəncərə YALNIZ proqnoz verə bilən modellər üzrə kəsişmədən alınır; heç bir
    # proqnoz verməyən model (məsələn nümunə tələbini ödəməyən) pəncərəni sıfırlamır, sadəcə
    # qiymətləndirmədən kənarda qalır və RMSE-si None kimi qeyd olunur.
    active = [m for m in mdl if len(errs[m]) >= max(int(min_points), 1)]
    common = sorted(set.intersection(*[set(errs[m].keys()) for m in active])) if active else []
    n_common = len(common)

    def _rmse(m: str) -> Optional[float]:
        if m not in active or n_common < min_points:
            return None
        return float(np.sqrt(np.mean([errs[m][t] ** 2 for t in common])))

    rmse = {m: _rmse(m) for m in mdl}
    rw = rmse.get("RW")

    def _cov(m: str) -> Optional[float]:
        r = rmse.get(m)
        if r is None:
            return None
        if r <= 0:
            return 1.0
        return float(np.mean([abs(errs[m][t]) <= _Z80 * r for t in common]))

    cover = {m: _cov(m) for m in mdl}
    beats = {m: bool(rmse.get(m) is not None and rw is not None and rmse[m] < rw)
             for m in mdl if m != "RW"}

    kept = {m: rmse[m] for m in mdl if m != "RW" and beats.get(m)}
    if kept:
        inv = {m: 1.0 / (max(r, 1e-12) ** 2) for m, r in kept.items()}
        Z = sum(inv.values())
        weights = {m: inv[m] / Z for m in inv}
    else:
        # heç bir struktur model RW-ni üstələmir → proqnozu etalon daşıyır (§D4)
        weights = {"RW": 1.0} if rw is not None else {}

    if weights and common:
        ens_e = [sum(weights[m] * errs[m][t] for m in weights) for t in common]
        ens_rmse = float(np.sqrt(np.mean([e ** 2 for e in ens_e])))
    else:
        ens_rmse = None

    rows: List[Dict[str, object]] = []
    for m in mdl:
        for t in common:
            if t not in preds[m]:
                continue
            rows.append({
                "series_code": series_code, "model": m,
                "vintage_year": int(t - 1),          # informasiya nümunəsinin bitdiyi il
                "horizon_h": 1,
                "actual": float(frame.loc[t, "y"]), "forecast": preds[m][t],
                "error": errs[m][t], "abs_error": abs(errs[m][t]),
                "rmse_h": rmse[m], "rw_rmse_h": rw, "coverage80": cover[m],
            })
    table = pd.DataFrame(rows, columns=[
        "series_code", "model", "vintage_year", "horizon_h", "actual", "forecast",
        "error", "abs_error", "rmse_h", "rw_rmse_h", "coverage80"])

    res = BacktestResult(
        series_code=series_code, rmse=rmse, rw_rmse=rw, weights=weights, table=table,
        errors=errs, coverage80=cover, common_years=[int(t) for t in common],
        ensemble_rmse=ens_rmse, beats_rw=beats,
    )
    if verbose:
        print(f"[SINAQ] {res.note_az}")
    return res


# ---------------------------------------------------------------------------
# 6. Yelpik zolaqları
# ---------------------------------------------------------------------------
def fan(path, resid_sigma: float, coef_cov=None, ndraw: int = 2000,
        X_future=None, cumulative: bool = True, seed: int = SEED) -> pd.DataFrame:
    """Yelpik zolaqları: qalıq dispersiyası + parametr qeyri-müəyyənliyi, üfüq boyu genişlənən.

    Simulyasiya: hər çəkiliş üçün (a) hər il müstəqil qalıq şoku ε ~ N(0, σ²) yaradılır və
    `cumulative=True` olduqda üfüq boyu YIĞILIR (səviyyə/loq-səviyyə yollarında xəta √h qaydası
    ilə genişlənir); (b) `coef_cov` verildikdə əmsal vektoru üçün çoxölçülü normal paylanmadan
    sapma çəkilir və gələcək dizayn matrisi vasitəsilə yola ötürülür. Zolaqlar `ndraw` yolun
    empirik faizliklərindən hesablanır.

    `X_future` verilmədikdə parametr sapması vahid sətir vektoru ilə ötürülür — bu, **konservativ
    yaxınlaşmadır** (parametr qeyri-müəyyənliyi səviyyə sürüşməsi kimi daxil olur); dəqiq nəticə
    üçün gələcək dizayn matrisi verilməlidir.

    Nəticə: `path` indeksi ilə `lo80, hi80, lo50, hi50` sütunları (§3 müqaviləsinin adları).
    """
    p = pd.Series(path).astype(float)
    H = len(p)
    if H == 0:
        return pd.DataFrame(columns=["lo80", "hi80", "lo50", "hi50"])
    sigma = float(abs(resid_sigma))
    rng = np.random.default_rng(seed)

    eps = rng.normal(0.0, sigma, size=(int(ndraw), H)) if sigma > 0 else np.zeros((int(ndraw), H))
    dev = np.cumsum(eps, axis=1) if cumulative else eps

    if coef_cov is not None:
        C = np.asarray(coef_cov, dtype=float)
        if C.ndim != 2 or C.shape[0] != C.shape[1]:
            raise ValueError("coef_cov kvadrat kovariasiya matrisi olmalıdır.")
        k = C.shape[0]
        Xf = np.ones((H, k)) if X_future is None else np.asarray(X_future, dtype=float)
        if Xf.shape != (H, k):
            raise ValueError(f"X_future ölçüsü ({H}, {k}) olmalıdır, verilən: {Xf.shape}.")
        db = rng.multivariate_normal(np.zeros(k), C, size=int(ndraw), method="svd")
        dev = dev + db @ Xf.T

    sims = p.values[None, :] + dev
    q = np.percentile(sims, [10.0, 90.0, 25.0, 75.0], axis=0)
    return pd.DataFrame({"lo80": q[0], "hi80": q[1], "lo50": q[2], "hi50": q[3]}, index=p.index)


# ---------------------------------------------------------------------------
# 7. Simultan nüvənin həlli (Gauss–Seidel sabit nöqtə)
# ---------------------------------------------------------------------------
_LAG_RE = re.compile(r"^\s*([A-Za-z_][\w\.]*)\s*\(\s*-\s*(\d+)\s*\)\s*$")


def _parse_driver(key: str) -> Tuple[str, int]:
    """"x" → (x, 0); "x(-1)" → (x, 1). Gecikmə göstərilməzsə cari il nəzərdə tutulur."""
    m = _LAG_RE.match(key)
    if m:
        return m.group(1), int(m.group(2))
    return key.strip(), 0


def _series_lookup(obj, year: int) -> Optional[float]:
    if obj is None:
        return None
    if isinstance(obj, (int, float, np.floating)):
        return float(obj)
    if isinstance(obj, pd.Series):
        return float(obj.loc[year]) if year in obj.index else None
    if isinstance(obj, Mapping):
        return float(obj[year]) if year in obj else None
    return None


def solve_core(eqs: Mapping[str, Tuple[float, Mapping[str, float]]],
               exog: Mapping[str, object], years: Sequence[int],
               tol: float = 1e-8, max_iter: int = 200,
               init: Optional[Mapping[str, object]] = None,
               verbose: bool = False) -> Dict[str, pd.Series]:
    """Simultan xətti blokun il-bə-il sabit nöqtə (Gauss–Seidel) həlli.

    `eqs` — {hədəf: (sabit, {sürücü: əmsal})}. Sürücü adı endogen (eyni lüğətdə hədəf olan) və ya
    ekzogen ola bilər; "ad(-1)" yazılışı bir il gecikməni bildirir. Gecikmiş endogen dəyər əvvəlki
    illərin həllindən, üfüqdən əvvəlki illər üçün isə `init`-dən götürülür.

    Alqoritm hər il üçün dəyərləri ardıcıl yeniləyir (Gauss–Seidel: yenilənmiş dəyər dərhal
    növbəti tənlikdə istifadə olunur) və maksimal dəyişmə `tol`-dan kiçik olana qədər təkrarlanır.
    Blok xətti və diaqonal üstünlüklü olduğundan yığılma sürətlidir; `max_iter` daxilində
    yığılma baş verməzsə `ConvergenceError` qaldırılır (səssiz yaxınlaşmaya yol verilmir).
    """
    targets = list(eqs.keys())
    solved: Dict[str, Dict[int, float]] = {t: {} for t in targets}
    init = init or {}

    def lookup(base: str, yr: int, cur: Dict[str, float], this_year: int) -> float:
        if base in solved:
            if yr == this_year:
                return cur[base]
            if yr in solved[base]:
                return solved[base][yr]
            v = _series_lookup(init.get(base), yr)
            if v is not None:
                return v
            raise KeyError(f"'{base}' dəyişəninin {yr} ili üçün dəyəri yoxdur "
                           f"(nə həll olunmuş, nə də `init` daxilində).")
        v = _series_lookup(exog.get(base), yr)
        if v is None:
            v = _series_lookup(init.get(base), yr)
        if v is None:
            raise KeyError(f"Ekzogen '{base}' dəyişəninin {yr} ili üçün dəyəri verilməyib.")
        return v

    for yr in [int(y) for y in years]:
        # başlanğıc təxmini: əvvəlki ilin həlli (varsa), əks halda sıfır
        cur = {t: (solved[t][yr - 1] if (yr - 1) in solved[t] else
                   (_series_lookup(init.get(t), yr - 1) or 0.0)) for t in targets}
        converged = False
        for it in range(1, int(max_iter) + 1):
            maxdiff = 0.0
            for t in targets:
                const, coefs = eqs[t]
                val = float(const)
                for drv, c in coefs.items():
                    base, lag = _parse_driver(drv)
                    val += float(c) * lookup(base, yr - lag, cur, yr)
                maxdiff = max(maxdiff, abs(val - cur[t]))
                cur[t] = val          # Gauss–Seidel: dərhal yenilənir
            if maxdiff < tol:
                converged = True
                break
        if not converged:
            raise ConvergenceError(
                f"Simultan nüvə {yr} ilində {max_iter} iterasiyada yığılmadı "
                f"(son maksimal dəyişmə {maxdiff:.3e}, tələb olunan {tol:.1e}). "
                f"Blokun əmsalları yoxlanılmalıdır."
            )
        if verbose:
            print(f"[NÜVƏ] {yr}: {it} iterasiyada yığıldı.")
        for t in targets:
            solved[t][yr] = cur[t]

    return {t: pd.Series(solved[t]).sort_index() for t in targets}


# ---------------------------------------------------------------------------
# 8. Dinamika bloku — profil yolları, çoxüfüqlü sınaq və dinamiklik göstəricisi
# ---------------------------------------------------------------------------
# Bu bölmə dinamika tələblərinin üç düzəlişindən ikisini paketin YEGANƏ qatına yığır:
#   (ii) sabit fərziyyə açarlarının yerinə **profil** yolları — lövbər səviyyəsi, yığılma sürəti
#        və sıranın öz avtokorrelyasiyasına kalibrlənmiş dövri komponent;
#   (metrik) σ(proqnoz)/σ(faktiki) — **dinamiklik göstəricisi**.
# Birinci düzəliş (gecikmiş asılı dəyişən / ECM formaları) tənlik səviyyəsindədir və dəftərlərdə
# `backtest` ilə qurulur; burada yalnız onun üçün lazım olan səviyyə köməkçisi (`level_from_growth`)
# saxlanılır.
#
# METODOLOJİ QEYD — nə üçün YOL seçimi bir addımlıq sınaqla aparılmır.
# Bir addımlıq (h = 1) sınaq tənliyin FORMASINI seçmək üçün doğru ölçüdür: hər il üçün model
# yalnız ondan əvvəlki məlumatla qiymətləndirilir və növbəti ilin dəyəri proqnozlaşdırılır.
# Lakin NƏŞR OLUNAN obyekt bir illik proqnoz deyil — beş illik YOLDUR (2026-2030). Təsadüfi
# gəzişmənin (RW) yolu "son müşahidə hər il təkrarlanır" deməkdir: h = 1-də çox vaxt üstün olan
# bu qayda h ≥ 2-də stasionar artım templəri üçün orta qiymətə qayıdışdan geri qalır. Ona görə
# yolun modelləri (və ya onların çəkiləri) EYNİ genişlənən pəncərə məntiqi ilə, lakin h = 1…H
# üfüqlərinin hamısında ölçülür (`path_backtest`). Hər iki nəticə nəşr olunur: forma qərarı
# h = 1 cədvəlindən, yol qərarı çoxüfüqlü cədvəldən oxunur.

DYN_MODEL = "DINAMIKLIK"          # `validation_backtest.csv`-də göstəricinin model adı
# «Nazirlik spesifikasiyası» ssenarisinin (FR13) sınaq sətirləri EYNİ faylda, eyni müqavilə
# sütunları ilə saxlanılır, lakin onlar AYRI spesifikasiyanın nəticəsidir: bu paketin mərkəzi
# yolunun ansambl çəkisinə və "ən yaxşı model" seçiminə DAXİL EDİLMİR. Oxuyan hər dəftər bu adı
# `DYN_MODEL` ilə birlikdə kənarlaşdırır.
SPEC_MODEL = "NAZIRLIK_SPES"
NON_ENSEMBLE_MODELS = (DYN_MODEL, SPEC_MODEL)
DYN_HIST_START = 2013             # faktiki σ-nın hesablandığı pəncərənin başlanğıcı
# DAXİLİ qəbul aralığı (D2). Nəşr olunan YALNIZ göstəricinin özüdür; aralıq müqavilə sənədinə
# çevrilməsin deyə heç bir çıxış faylına, hesabata və ya ekrana yazılmır — burada, kodda qalır.
DYN_BAND_LO = 0.4
DYN_BAND_HI = 1.2

# ---------------------------------------------------------------------------
# Dinamiklik göstəricisinin SİNİF bölgüsü (D2). Aralıqdan kənara düşən sıranın əsaslandırması
# iki hissədən ibarətdir: (1) sıranın hansı MEXANİZMƏ görə az dəyişkən olduğu — bu, aşağıdakı
# dörd sinifdən biridir və hər sinif üçün mətn burada, bir yerdə saxlanılır; (2) sıranın ÖZ
# ölçülmüş diaqnostikası (ansambl çəkiləri, sürücü yolunun eni, tarixi σ) — onu hər dəftər özü
# verir. Sinif mətni nəşr olunur, aralığın özü isə nəşr olunmur (D2 dəyişməyib).
#
# Yeni ALL-CAPS nişan İCAD EDİLMİR (§5.3): icazə verilən nişanlar yalnız STRUKTUR, ETALON,
# DATA-GAP-dır; ona görə sinif mətnləri adi AZ nəsri ilə yazılır və yalnız `etalon` sinfi
# mövcud ETALON nişanını işlədir.
DYN_CLASSES: Dict[str, str] = {
    # (a) ekzogen/plan yolu: dəyişkənlik modeldən deyil, NƏŞR OLUNMUŞ xarici girişdən gəlir
    "ekzogen": ("ekzogen yolla şərtlənən sıra — dəyişkənlik `assumptions.csv`-də nəşr olunan "
                "xarici girişdən (hasilat/büdcə planı, layihə bazası və ya sürücü açarı) gəlir, "
                "modelin özündən deyil; şərti orta proqnoz onu şərtləndirən yoldan daha dəyişkən "
                "ola bilməz"),
    # (b) qiymət bloku: deflyator/inflyasiya sırası sabit qiymət prosesinə bağlıdır
    "qiymet": ("qiymət blokunun sırası — yol İQİ hədəf yoluna və/və ya Brent-in şərti orta yoluna "
               "bağlıdır; sabit qiymət prosesi üzərində şərti ortanın az dəyişkən olması "
               "qüsur deyil, məhz DÜZGÜN davranışdır"),
    # (c) etalonla daşınan yol: struktur namizədlər sınaqda uduzdu (D4 qaydası (a))
    # Uduzan yarışın KONKRET ünvanı sıranın öz mətnində verilir, ona görə burada təkrarlanmır.
    "etalon": ("ETALON ilə daşınan yol — struktur namizədlər nümunədən kənar sınaqda təsadüfi "
               "gəzişməni üstələmədi, ona görə D4 qaydasına əsasən mərkəzi yolu etalon daşıyır və "
               "yol konstruksiya etibarilə hamardır"),
    # (d) izah olunmayan hal: səmimi şəkildə AÇIQ saxlanılır, örtülmür
    "baxilir": ("səbəb tam izah olunmur — sıra YENİDƏN BAXILIR: aşağı dəyişkənliyin mexanizmi "
                "yuxarıdakı üç sinifdən heç biri ilə tam əsaslandırıla bilmir və növbəti "
                "buraxılışda ayrıca araşdırılacaq"),
}


def level_from_growth(g, base: float = 100.0, base_year: Optional[int] = None) -> pd.Series:
    """Faiz artım tempi sırasından səviyyə indeksi qurur: L_t = L_{t-1}(1 + g_t/100).

    Səviyyənin miqyası ixtiyaridir (loq-səviyyə tənliyində sabitə hopur), ona görə kointeqrasiya
    və ECM formaları artım tempi kimi nəşr olunan sıralar üçün də qurula bilir.
    """
    s = pd.Series(g).astype(float).dropna().sort_index()
    if s.empty:
        return pd.Series(dtype=float)
    lvl = base * (1.0 + s / 100.0).cumprod()
    if base_year is not None and base_year in lvl.index:
        lvl = lvl * (base / float(lvl.loc[base_year]))
    return lvl.rename(getattr(s, "name", None))


# --- AR profilləri ---------------------------------------------------------
def _ar_fit(s: pd.Series, p: int) -> Optional[Tuple[float, List[float]]]:
    """AR(p) sabit hədli OLS. (const, [ρ1..ρp]) qaytarır; nümunə çatmazsa None."""
    y = pd.Series(s).astype(float).dropna().sort_index()
    if len(y) < max(6, 3 * p + 2):
        return None
    d = pd.DataFrame({"y": y})
    for k in range(1, p + 1):
        d[f"l{k}"] = y.shift(k)
    d = d.dropna()
    if len(d) < max(5, p + 3):
        return None
    cols = [f"l{k}" for k in range(1, p + 1)]
    fit = sm.OLS(d["y"], sm.add_constant(d[cols], has_constant="add")).fit()
    return float(fit.params["const"]), [float(fit.params[c]) for c in cols]


def ar_profile(hist, years: Sequence[int], p: int = 1) -> Optional[pd.Series]:
    """AR(p) şərti orta yolu: lövbər (uzunmüddətli orta) + son müşahidədən yığılma (+ dövrilik).

    Bu, uydurulmuş dövri komponent DEYİL — sıranın öz avtokorrelyasiyasına qiymətləndirilmiş
    AR(p) prosesinin şərti gözləmə yoludur: y_{T+h} = c + Σ ρ_k y_{T+h−k}. p = 1 monoton
    yığılma verir; p = 2 köklər kompleks olduqda **sönən dövr** verir — tələb olunan "sıranın
    tarixi avtokorrelyasiyasına kalibrlənmiş dövri komponent" tələbi məhz budur.
    """
    h = pd.Series(hist).astype(float).dropna().sort_index()
    fitres = _ar_fit(h, p)
    if fitres is None:
        return None
    c0, rhos = fitres
    if abs(sum(rhos)) >= 0.999:        # qeyri-stasionar qiymətləndirmə — profil qurulmur
        return None
    hist_vals = [float(v) for v in h.iloc[-p:]]
    out: Dict[int, float] = {}
    lags = list(reversed(hist_vals))    # lags[0] = y_{t-1}, lags[1] = y_{t-2}, ...
    for yr in [int(v) for v in years]:
        nxt = c0 + sum(rhos[k] * lags[k] for k in range(p))
        out[yr] = float(nxt)
        lags = [float(nxt)] + lags[:-1]
    return pd.Series(out).sort_index()


def _ar_cycle_note(hist, p: int = 2) -> str:
    """AR(2) köklərinin kompleks olub-olmadığını (yəni sönən dövrün mövcudluğunu) yazır."""
    fitres = _ar_fit(pd.Series(hist), p)
    if fitres is None or p != 2:
        return ""
    _, (r1, r2) = fitres
    disc = r1 ** 2 + 4 * r2
    if disc < 0 and r2 != 0:
        period = 2 * math.pi / math.acos(min(1.0, max(-1.0, r1 / (2 * math.sqrt(-r2))))) if -r2 > 0 else None
        return f"kompleks köklər → sönən dövr, təxmini dövr uzunluğu {period:.1f} il" if period else "kompleks köklər → sönən dövr"
    return "həqiqi köklər → monoton yığılma"


# --- Çoxüfüqlü (h = 1…H) genişlənən pəncərəli sınaq ------------------------
@dataclass
class PathBacktestResult:
    """h = 1…H üfüqlərinin hamısında ölçülmüş genişlənən pəncərəli sınaq.

    `rmse` — üfüqlər üzrə BİRLƏŞDİRİLMİŞ (pooled) RMSE; `rmse_by_h` — üfüq-üfüq RMSE.
    Çəkilər yalnız RW-nin birləşdirilmiş RMSE-sini üstələyən modellər arasında paylanır (§D4).
    """
    series_code: str
    horizon: int
    rmse: Dict[str, Optional[float]]
    rmse_by_h: Dict[str, Dict[int, Optional[float]]]
    rw_rmse: Optional[float]
    weights: Dict[str, float]
    n_points: Dict[str, int]
    beats_rw: Dict[str, bool]
    table: pd.DataFrame
    ensemble_rmse: Optional[float] = None
    ensemble_rmse_by_h: Dict[int, Optional[float]] = field(default_factory=dict)

    @property
    def skill(self) -> Dict[str, Optional[float]]:
        return {m: (None if (r is None or not self.rw_rmse) else float(1.0 - r / self.rw_rmse))
                for m, r in self.rmse.items()}

    @property
    def note_az(self) -> str:
        kept = ", ".join(f"{m} {w:.2f}" for m, w in self.weights.items()) or "yoxdur"
        return (f"Çoxüfüqlü sınaq ({self.series_code}, h = 1-{self.horizon}): "
                f"RW birləşdirilmiş RMSE = {_fmt(self.rw_rmse, 3)}; yol çəkiləri: {kept}.")

    def combine(self, paths: Mapping[str, Mapping[int, float]]) -> pd.Series:
        missing = [m for m in self.weights if m not in paths]
        if missing:
            raise KeyError(f"Yol ansamblı üçün verilməyən modellər: {missing}")
        years = sorted(set.intersection(*[set(paths[m].keys()) for m in self.weights]))
        return pd.Series({t: sum(self.weights[m] * float(paths[m][t]) for m in self.weights)
                          for t in years})


def path_backtest(y, path_models: Mapping[str, Callable], horizon: int = 5, min_train: int = 8,
                  series_code: str = "", min_points: int = 6,
                  verbose: bool = False) -> PathBacktestResult:
    """Genişlənən pəncərəli, ÇOXÜFÜQLÜ (h = 1…H), nümunədən kənar sınaq.

    Hər vintaj ili v üçün model YALNIZ v-yə qədərki məlumatla qurulur və v+1 … v+H illəri üçün
    BÜTÖV YOL verir; hər üfüqün xətası ayrıca yığılır. Modelin funksiya imzası
    `fn(train: pd.Series, years: List[int]) -> Mapping[int, float]`-dır.

    "RW" (sabit yol — son müşahidə hər il təkrarlanır) etalonu avtomatik əlavə olunur; onu
    üstələməyən model çəki almır (§D4). Ədalət üçün bütün modellər EYNİ (vintaj, üfüq) cütlərində
    ölçülür.
    """
    s = pd.Series(y).astype(float).dropna()
    s.index = years_of(s.index)
    s = s.sort_index()
    lim = last_actual()
    s = s[s.index <= lim]

    mdl: Dict[str, Callable] = dict(path_models)
    mdl.setdefault("RW", lambda train, yrs: {t: float(train.iloc[-1]) for t in yrs})

    idx = [int(t) for t in s.index]
    errs: Dict[str, Dict[Tuple[int, int], float]] = {m: {} for m in mdl}
    for i, v in enumerate(idx):
        if i + 1 < min_train:
            continue
        train = s.loc[:v]
        tgt = [t for t in idx if v < t <= v + horizon]
        if not tgt:
            continue
        for nm, fn in mdl.items():
            try:
                pth = fn(train, tgt)
            except Exception:
                pth = None
            if not pth:
                continue
            for t in tgt:
                if t not in pth or pd.isna(pth[t]):
                    continue
                errs[nm][(v, int(t - v))] = float(pth[t]) - float(s.loc[t])

    active = [m for m in mdl if len(errs[m]) >= max(int(min_points), 1)]
    common = sorted(set.intersection(*[set(errs[m].keys()) for m in active])) if active else []

    def _rm(m: str, keys) -> Optional[float]:
        if m not in active or not keys:
            return None
        vals = [errs[m][k] ** 2 for k in keys if k in errs[m]]
        return float(np.sqrt(np.mean(vals))) if vals else None

    rmse = {m: _rm(m, common) for m in mdl}
    by_h = {m: {h: _rm(m, [k for k in common if k[1] == h]) for h in range(1, horizon + 1)}
            for m in mdl}
    rw = rmse.get("RW")
    beats = {m: bool(rmse.get(m) is not None and rw is not None and rmse[m] < rw)
             for m in mdl if m != "RW"}
    kept = {m: rmse[m] for m in mdl if m != "RW" and beats.get(m)}
    if kept:
        inv = {m: 1.0 / (max(r, 1e-12) ** 2) for m, r in kept.items()}
        Z = sum(inv.values())
        weights = {m: inv[m] / Z for m in inv}
    else:
        weights = {"RW": 1.0} if rw is not None else {}

    rows = []
    for m in mdl:
        for h in range(1, horizon + 1):
            r = by_h[m][h]
            if r is None:
                continue
            rows.append({"series_code": series_code, "model": m, "horizon_h": h,
                         "n": len([k for k in common if k[1] == h]),
                         "rmse_h": r, "rw_rmse_h": by_h["RW"][h],
                         "bacarıq %": (None if not by_h["RW"][h] else
                                       round((1 - r / by_h["RW"][h]) * 100, 1))})
    # ansamblın (çəkilərlə birləşdirilmiş yolun) öz OOS xətası — zolağın σ₁-i buradan gəlir
    ens_by_h: Dict[int, Optional[float]] = {}
    ens_all: Optional[float] = None
    if weights and common:
        ens_e = {k: sum(weights[m] * errs[m][k] for m in weights) for k in common
                 if all(k in errs[m] for m in weights)}
        if ens_e:
            ens_all = float(np.sqrt(np.mean([e ** 2 for e in ens_e.values()])))
            for h in range(1, horizon + 1):
                vals = [e ** 2 for k, e in ens_e.items() if k[1] == h]
                ens_by_h[h] = float(np.sqrt(np.mean(vals))) if vals else None

    res = PathBacktestResult(
        series_code=series_code, horizon=int(horizon), rmse=rmse, rmse_by_h=by_h, rw_rmse=rw,
        weights=weights, n_points={m: len(errs[m]) for m in mdl}, beats_rw=beats,
        table=pd.DataFrame(rows), ensemble_rmse=ens_all, ensemble_rmse_by_h=ens_by_h)
    if verbose:
        print(f"[YOL SINAĞI] {res.note_az}")
    return res


def profile_path(hist, years: Sequence[int], name: str = "", min_train: int = 8,
                 horizon: Optional[int] = None, min_points: int = 6,
                 extra_models: Optional[Mapping[str, Callable]] = None,
                 verbose: bool = False) -> Tuple[pd.Series, PathBacktestResult, str]:
    """Bir sıranı **profil** kimi proqnoz üfüqünə uzadır və qərarı çoxüfüqlü sınağa buraxır.

    Namizəd yollar: `SABIT` (son müşahidə — RW), `ORTA5` (son beş ilin ortası), `AR1` (monoton
    yığılma), `AR2` (avtokorrelyasiyaya kalibrlənmiş sönən dövr) və çağıran tərəfin əlavə etdiyi
    yollar. Qazanan(lar) tərs-RMSE² çəkiləri ilə birləşdirilir; heç biri sabit yolu üstələmirsə
    yol SABİT qalır və bu, qeydə AÇIQ yazılır (uydurma dinamika əlavə edilmir).

    Qaytarır: (yol, çoxüfüqlü sınaq nəticəsi, AZ qeyd).
    """
    h = pd.Series(hist).astype(float).dropna().sort_index()
    h.index = years_of(h.index)
    h = h[h.index <= last_actual()]
    yrs = [int(v) for v in years]
    H = int(horizon or len(yrs))

    def _m_flat(train, ys):
        return {t: float(train.iloc[-1]) for t in ys}

    def _m_mean5(train, ys):
        if len(train) < 5:
            return {}
        v = float(train.iloc[-5:].mean())
        return {t: v for t in ys}

    def _mk_ar(p):
        def f(train, ys):
            pth = ar_profile(train, ys, p=p)
            return {} if pth is None else {int(t): float(v) for t, v in pth.items()}
        return f

    models: Dict[str, Callable] = {"ORTA5": _m_mean5, "AR1": _mk_ar(1), "AR2": _mk_ar(2)}
    if extra_models:
        models.update(dict(extra_models))
    bt = path_backtest(h, models, horizon=H, min_train=min_train, series_code=name,
                       min_points=min_points, verbose=False)

    last = float(h.iloc[-1])
    paths: Dict[str, Dict[int, float]] = {
        "RW": {t: last for t in yrs},
        "ORTA5": {t: float(h.iloc[-5:].mean()) for t in yrs} if len(h) >= 5 else {t: last for t in yrs},
    }
    for p, nm in ((1, "AR1"), (2, "AR2")):
        pp = ar_profile(h, yrs, p=p)
        paths[nm] = {t: last for t in yrs} if pp is None else {int(t): float(v) for t, v in pp.items()}
    if extra_models:
        for nm, fn in extra_models.items():
            try:
                pp = fn(h, yrs)
            except Exception:
                pp = None
            paths[nm] = {t: last for t in yrs} if not pp else {int(t): float(v) for t, v in pp.items()}

    w = {m: v for m, v in bt.weights.items() if m in paths}
    if not w:
        w = {"RW": 1.0}
    tot = sum(w.values())
    w = {m: v / tot for m, v in w.items()}
    path = pd.Series({t: sum(w[m] * paths[m][t] for m in w) for t in yrs}).sort_index()

    flat = (max(path) - min(path)) < 1e-9
    note = (f"profil: {', '.join(f'{m} {v:.2f}' for m, v in w.items())}"
            f"; çoxüfüqlü (h=1-{H}) RW RMSE = {_fmt(bt.rw_rmse, 3)}"
            f"; AR(2): {_ar_cycle_note(h)}"
            f"; tarixi əhatə {int(h.index.min())}-{int(h.index.max())}"
            + ("; SABİT yol — heç bir dinamik namizəd sabit yolu üstələmədi" if flat else ""))
    if verbose:
        print(f"[PROFİL] {name}: {note}")
    return path, bt, note


# --- Dinamiklik göstəricisi ----------------------------------
@dataclass
class DynamicsResult:
    """σ(proqnoz)/σ(faktiki) — dinamiklik göstəricisi (populyasiya σ-sı)."""
    series_code: str
    label_az: str
    sd_forecast: float
    sd_actual: float
    ratio: float
    n_actual: int
    n_forecast: int
    hist_from: int
    hist_to: int
    fc_from: int
    fc_to: int
    justification: str = ""
    dyn_class: str = ""

    @property
    def in_band(self) -> bool:
        return bool(DYN_BAND_LO <= self.ratio <= DYN_BAND_HI)

    @property
    def under_review(self) -> bool:
        """Sinif (d): mexanizm tam izah olunmur — qeyd bunu AÇIQ yazır, örtmür."""
        return self.dyn_class == "baxilir"

    @property
    def note_text(self) -> str:
        """Nəşr olunan qeydin MƏTNİ: sinif izahı + sıranın öz ölçülmüş diaqnostikası.

        Sinif verilməyibsə yalnız diaqnostika qaytarılır (geriyə uyğunluq): köhnə çağırış
        yerləri qeydsiz qalmır.
        """
        cls = DYN_CLASSES.get(self.dyn_class, "")
        just = str(self.justification or "").strip()
        if cls and just:
            return f"{cls}; {just}"
        return cls or just

    def __str__(self) -> str:
        return (f"DINAMIKLIK[{self.series_code}] σ(proqnoz)={self.sd_forecast:.3f} "
                f"σ(faktiki)={self.sd_actual:.3f} nisbət={self.ratio:.3f}")


def sigma_ratio(forecast, actual, series_code: str = "", label_az: str = "",
                hist_from: int = DYN_HIST_START, hist_to: Optional[int] = None,
                justification: str = "", dyn_class: str = "") -> DynamicsResult:
    """Dinamiklik göstəricisi: proqnoz yolunun standart kənarlaşması ÷ faktiki sıranın standart
    kənarlaşması. Hər ikisi **populyasiya** σ-sıdır (ddof = 0).

    Göstərici sıranın nə qədər "canlı" olduğunu ölçür: sıfıra yaxın dəyər proqnozun praktik olaraq
    sabit olduğunu bildirir. Şərti orta proqnoz tərifinə görə faktikidən az dəyişkən olmalıdır
    (proqnozlaşdırıla bilməyən şoklar onun tərkibində yoxdur), ona görə göstərici 1-dən kiçik
    gözlənilir; qiymətləndirmə aralığı DAXİLİdir və nəşr olunmur (D2).

    `dyn_class` — aşağı dəyişkənliyin MEXANİZM sinfi (`DYN_CLASSES` açarlarından biri:
    `ekzogen` / `qiymet` / `etalon` / `baxilir`). Aralıqdan kənara düşən sıra üçün sinif
    verilməlidir; tanınmayan açar səhvin səssiz keçməməsi üçün DƏRHAL rədd edilir.
    """
    if dyn_class and dyn_class not in DYN_CLASSES:
        raise ValueError(f"{series_code}: tanınmayan dinamiklik sinfi {dyn_class!r}; "
                         f"icazə verilənlər: {sorted(DYN_CLASSES)}")
    f = pd.Series(forecast).astype(float).dropna().sort_index()
    a = pd.Series(actual).astype(float).dropna().sort_index()
    a.index = years_of(a.index)
    f.index = years_of(f.index)
    hi = last_actual() if hist_to is None else int(hist_to)
    a = a[(a.index >= int(hist_from)) & (a.index <= hi)]
    if len(a) < 3 or len(f) < 2:
        raise ValueError(f"{series_code}: dinamiklik göstəricisi üçün müşahidə azdır "
                         f"(faktiki n = {len(a)}, proqnoz n = {len(f)}).")
    sd_a = float(np.std(a.values, ddof=0))
    sd_f = float(np.std(f.values, ddof=0))
    if sd_a <= 0:
        raise ValueError(f"{series_code}: faktiki sıranın σ-sı sıfırdır, nisbət təyin olunmur.")
    return DynamicsResult(
        series_code=series_code, label_az=label_az or series_code,
        sd_forecast=sd_f, sd_actual=sd_a, ratio=float(sd_f / sd_a),
        n_actual=int(len(a)), n_forecast=int(len(f)),
        hist_from=int(a.index.min()), hist_to=int(a.index.max()),
        fc_from=int(f.index.min()), fc_to=int(f.index.max()),
        justification=justification, dyn_class=dyn_class)


def dynamics_rows(results: Iterable[DynamicsResult]) -> pd.DataFrame:
    """`validation_backtest.csv` üçün DINAMIKLIK sətirləri (§3 müqaviləsinin 11 sütunu).

    Sütunların bu cədvəldəki oxunuşu (dəftərlərdə və seriya lüğətində eyni şəkildə izah olunur):
      · `actual`     — σ(faktiki), tarixi pəncərə;
      · `forecast`   — σ(proqnoz), 2026-2030;
      · `error`      — σ(proqnoz) − σ(faktiki);
      · `rmse_h`     — GÖSTƏRİCİNİN ÖZÜ: σ(proqnoz)/σ(faktiki);
      · `rw_rmse_h`  — istinad dəyəri 1,0 (faktiki qədər dəyişkən proqnoz);
      · `horizon_h`  — proqnoz pəncərəsinin uzunluğu (il);
      · `vintage_year` — son faktiki il.
    """
    rows = []
    for r in results:
        rows.append({
            "series_code": r.series_code, "model": DYN_MODEL,
            "vintage_year": int(r.hist_to), "horizon_h": int(r.n_forecast),
            "actual": round(r.sd_actual, 6), "forecast": round(r.sd_forecast, 6),
            "error": round(r.sd_forecast - r.sd_actual, 6),
            "abs_error": round(abs(r.sd_forecast - r.sd_actual), 6),
            "rmse_h": round(r.ratio, 6), "rw_rmse_h": 1.0, "coverage80": None,
        })
    return pd.DataFrame(rows, columns=[
        "series_code", "model", "vintage_year", "horizon_h", "actual", "forecast",
        "error", "abs_error", "rmse_h", "rw_rmse_h", "coverage80"])


def assert_dynamics_band(results: Iterable[DynamicsResult], strict: bool = True) -> pd.DataFrame:
    """DAXİLİ qəbul yoxlaması (D2): göstərici aralıqdan kənardadırsa, sıra üçün YAZILI
    əsaslandırma tələb olunur.

    Aralığın özü nəşr olunmur — nə hesabatda, nə çıxış fayllarında, nə də ekranda: R13-ün
    təsnifat məntiqi ilə nəşr olunmuş hədəf müqavilə öhdəliyinə çevrilərdi (D2/D4 eyni məntiq).
    Buna görə funksiya yalnız DƏFTƏR daxilində çağırılır və pozuntu halında ya yazılı
    əsaslandırma tələb edir (strict), ya da xəbərdarlıq verir.
    """
    res = list(results)
    rows = [{"series_code": r.series_code, "nisbət": round(r.ratio, 3),
             "aralıqda": r.in_band, "sinif": r.dyn_class,
             "əsaslandırma": r.justification} for r in res]
    tab = pd.DataFrame(rows)
    # Sinif (d) — mexanizmi izah olunmayan sıralar. Bunlar ÖRTÜLMÜR: hər icrada ayrıca, görünən
    # sətirlə çap olunur ki, təhvildə "izah olundu" görüntüsü yaranmasın.
    review = [r for r in res if r.under_review]
    if review:
        print("[YENİDƏN BAXILIR] dinamiklik göstəricisinin mexanizmi tam izah olunmayan sıra(lar): "
              + "; ".join(f"{r.series_code}={r.ratio:.3f}" for r in review))
    bad = [r for r in res
           if not r.in_band and not (str(r.justification).strip() and str(r.dyn_class).strip())]
    if bad:
        detail = "; ".join(f"{r.series_code}={r.ratio:.3f}" for r in bad[:8])
        msg = (f"dinamiklik göstəricisi daxili qəbul aralığından kənardadır və yazılı "
               f"əsaslandırma (mətn + mexanizm sinfi) tam verilməyib: {detail}")
        if strict:
            raise DynamicsBandError(msg)
        # `warnings.warn(msg)` əvəzinə birbaşa çap: Python-un standart xəbərdarlıq formatlayıcısı
        # çağırış yerinin MÜTLƏQ fayl yolunu ("<yol>/src/econometrics.py:NNNN: UserWarning: ...")
        # mesaja əlavə edir, bu isə icra edilmiş dəftərin saxlanılan çıxışına düşərək təhvil
        # sızma yoxlamasını pozur (DELIVERY_SPEC §4.2). Davranış eynidir — icranı DAYANDIRMIR,
        # yalnız görünən xəbərdarlıq verir (D2/D4).
        print(f"[XƏBƏRDARLIQ] {msg}")
    return tab


# ---------------------------------------------------------------------------
# 9. Dinamika formaları — statik / gecikməli / ECM, bir yerdə
# ---------------------------------------------------------------------------
# Artım tempi kimi nəşr olunan sıra üçün ECM forması: səviyyə indeksi artımdan bərpa olunur
# (`level_from_growth`), uzunmüddətli əlaqə loq-səviyyələr üzərində qurulur, qısamüddətli tənlik
# isə fərqlər üzərində. Miqyas ixtiyaridir və uzunmüddətli tənliyin sabitinə hopur.
ECM_COINT_ALPHA = 0.10       # ECM formasının kointeqrasiya qapısı
ECM_MIN_OBS = 12             # `eg_coint` minimumu


def ecm_growth_columns(frame: pd.DataFrame, ycol: str = "y", xcol: str = "x",
                       suffix: str = "", min_obs: int = ECM_MIN_OBS) -> Optional[Dict[str, str]]:
    """Çərçivəyə ECM sütunlarını əlavə edir (loq-səviyyələr və onların birinci fərqləri).

    Çərçivə YERİNDƏ dəyişdirilir; qaytarılan lüğət sütun adlarını daşıyır. Müşahidə sayı
    `min_obs`-dan azdırsa None qaytarılır (ECM forması qurulmur, səbəb hesabatda yazılır).
    """
    sub = frame[[ycol, xcol]].dropna()
    if len(sub) < int(min_obs):
        return None
    frame[f"ly{suffix}"] = np.log(level_from_growth(sub[ycol]))
    frame[f"lx{suffix}"] = np.log(level_from_growth(sub[xcol]))
    frame[f"dly{suffix}"] = np.log1p(sub[ycol] / 100.0)
    frame[f"dlx{suffix}"] = np.log1p(sub[xcol] / 100.0)
    return {"ly": f"ly{suffix}", "lx": f"lx{suffix}",
            "dly": f"dly{suffix}", "dlx": f"dlx{suffix}"}


def ecm_growth_fit(ly, lx, dly, dlx) -> Optional[Dict[str, float]]:
    """İki addımlı Engle–Granger: uzunmüddətli səviyyə tənliyi + fərqlər üzrə korreksiya."""
    ly = pd.Series(ly).astype(float)
    lx = pd.Series(lx).astype(float)
    if len(ly.dropna()) < 8:
        return None
    X1 = sm.add_constant(pd.DataFrame({"lx": lx}), has_constant="add")
    lr = sm.OLS(ly, X1).fit()
    ec = ly - lr.predict(X1)
    d2 = pd.DataFrame({"dly": dly, "dlx": dlx, "ec1": ec.shift(1)}).dropna()
    if len(d2) < 6:
        return None
    sr = sm.OLS(d2["dly"], sm.add_constant(d2[["ec1", "dlx"]], has_constant="add")).fit()
    return {"c": float(lr.params["const"]), "beta": float(lr.params["lx"]),
            "a": float(sr.params["const"]), "alpha": float(sr.params["ec1"]),
            "gamma": float(sr.params["dlx"]), "n_lr": int(lr.nobs), "n_sr": int(sr.nobs),
            "alpha_p": float(sr.pvalues["ec1"])}


def ecm_growth_model(frame: pd.DataFrame, cols: Mapping[str, str], min_train: int = 10) -> Callable:
    """`backtest` üçün h = 1 modeli: hər vintajda ECM yenidən qiymətləndirilir; ec_{t-1} faktiki
    səviyyələrdən, Δln x_t isə sürücünün t ilindəki faktiki dəyərindən götürülür."""
    def f(train, t):
        dd = train[[cols["ly"], cols["lx"], cols["dly"], cols["dlx"]]].dropna()
        if len(dd) < min_train or t not in frame.index or pd.isna(frame.loc[t, cols["dlx"]]):
            return np.nan
        p = ecm_growth_fit(dd[cols["ly"]], dd[cols["lx"]], dd[cols["dly"]], dd[cols["dlx"]])
        if p is None or (t - 1) not in dd.index:
            return np.nan
        ec_last = float(dd.loc[t - 1, cols["ly"]] - p["c"] - p["beta"] * dd.loc[t - 1, cols["lx"]])
        dly = p["a"] + p["alpha"] * ec_last + p["gamma"] * float(frame.loc[t, cols["dlx"]])
        return float((np.exp(dly) - 1.0) * 100.0)
    return f


def ecm_growth_path(frame: pd.DataFrame, cols: Mapping[str, str], gx_future, years) -> Tuple[
        Optional[pd.Series], Optional[Dict[str, float]]]:
    """ECM formasının proqnoz yolu: səviyyə və korreksiya termi rekursiv aparılır."""
    dd = frame[[cols["ly"], cols["lx"], cols["dly"], cols["dlx"]]].dropna()
    p = ecm_growth_fit(dd[cols["ly"]], dd[cols["lx"]], dd[cols["dly"]], dd[cols["dlx"]])
    if p is None:
        return None, None
    ly_prev, lx_prev, out = float(dd[cols["ly"]].iloc[-1]), float(dd[cols["lx"]].iloc[-1]), {}
    gx = pd.Series(gx_future)
    for t in [int(v) for v in years]:
        if t not in gx.index or pd.isna(gx.loc[t]):
            break
        dlx = float(np.log1p(float(gx.loc[t]) / 100.0))
        dly = p["a"] + p["alpha"] * (ly_prev - p["c"] - p["beta"] * lx_prev) + p["gamma"] * dlx
        out[t] = float((np.exp(dly) - 1.0) * 100.0)
        ly_prev, lx_prev = ly_prev + dly, lx_prev + dlx
    return pd.Series(out), p


def dynamics_forms(frame: pd.DataFrame, xcols: Sequence[str], series_code: str,
                   ycol: str = "y", ecm_x: Optional[str] = None, min_train: int = 8,
                   min_points: int = 4, extra_models: Optional[Mapping[str, Callable]] = None,
                   zero_at_forecast: Optional[Sequence[str]] = None
                   ) -> Tuple[BacktestResult, pd.DataFrame, str]:
    """Bir davranış tənliyinin ÜÇ forması eyni pəncərədə (dinamika yarışı):

    `STATIK` — mövcud spesifikasiya; `GECIKME` — üzərinə gecikmiş asılı dəyişən; `ECM` — loq
    səviyyələr üzərində Engle–Granger (yalnız `ecm_x` verildikdə və kointeqrasiya təsdiqləndikdə
    mərkəzi yola buraxılır). Qərarı genişlənən pəncərəli, bir addımlıq sınaq verir; UDUZAN forma
    silinmir — cədvəldə və `validation_backtest.csv`-də `<kod>_dinamika` adı ilə qalır (D4).

    `zero_at_forecast` — sınıq nöqtə (dummy) sütunlarının siyahısı: onlar TƏLİM nümunəsində
    normal reqressor kimi işləyir, lakin PROQNOZ addımında sıfır götürülür. Bu, real vaxt
    intizamıdır — model sınaq anında gələcək devalvasiyanı və ya rejim sınığını bilə bilməz.
    Verilmədikdə davranış dəyişmir (dummy sütunu proqnoz addımında da öz dəyəri ilə daxil olur).

    Qaytarır: (sınaq nəticəsi, qərar cədvəli, qazanan formanın adı).
    """
    zero_fc = set(zero_at_forecast or ())
    fr = frame.copy()
    fr["__lag__"] = fr[ycol].shift(1)
    cols_st = list(xcols)
    cols_lg = list(xcols) + ["__lag__"]

    def _mk(cs):
        need = [q for q in cs if q not in zero_fc]      # sıfırlanan sütun NaN-a görə bloklamamalıdır
        def f(train, t):
            dd = train[[ycol] + cs].dropna()
            if len(dd) < min_train or t not in fr.index or pd.isna(fr.loc[t, need]).any():
                return np.nan
            fit = sm.OLS(dd[ycol], sm.add_constant(dd[cs], has_constant="add")).fit()
            v = float(fit.params.get("const", 0.0))
            for q in cs:
                xq = 0.0 if q in zero_fc else float(fr.loc[t, q])
                v += float(fit.params[q]) * xq
            return v
        return f

    models: Dict[str, Callable] = {"STATIK": _mk(cols_st), "GECIKME": _mk(cols_lg)}
    if extra_models:
        models.update(dict(extra_models))
    coint_p, ecm_ok = None, False
    if ecm_x is not None:
        cols = ecm_growth_columns(fr, ycol, ecm_x, suffix="_dyn")
        if cols is not None:
            sub = fr[[cols["ly"], cols["lx"]]].dropna()
            try:
                coint_p = float(eg_coint(sub[cols["ly"]], sub[cols["lx"]],
                                         name=f"ECM {series_code}").pvalue)
            except Exception:
                coint_p = None
            models["ECM"] = ecm_growth_model(fr, cols)
    bt = backtest(fr.rename(columns={ycol: "y"}) if ycol != "y" else fr, models,
                  min_train=min_train, min_points=min_points,
                  series_code=f"{series_code}_dinamika", verbose=False)

    def _sk(m):
        r = bt.rmse.get(m)
        return None if (r is None or not bt.rw_rmse) else round((1 - r / bt.rw_rmse) * 100, 1)

    ecm_ok = (coint_p is not None and coint_p < ECM_COINT_ALPHA and _sk("ECM") is not None)
    scores = {m: _sk(m) for m in ("STATIK", "GECIKME") if _sk(m) is not None}
    if ecm_ok:
        scores["ECM"] = _sk("ECM")
    win = max(scores, key=lambda m: (scores[m], m == "STATIK")) if scores else "STATIK"
    tab = pd.DataFrame([{
        "sıra": series_code,
        "pəncərə": (f"{bt.common_years[0]}-{bt.common_years[-1]}" if bt.common_years else "boş"),
        "STATIK %": _sk("STATIK"), "GECIKME %": _sk("GECIKME"), "ECM %": _sk("ECM"),
        "kointeqrasiya p": (None if coint_p is None else round(coint_p, 4)),
        "qazanan forma": win + ("" if (ecm_x is None or ecm_ok)
                                else "; ECM kointeqrasiya qapısından keçmədi"),
    }])
    return bt, tab, win
