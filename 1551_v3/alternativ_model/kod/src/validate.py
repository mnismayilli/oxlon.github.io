"""validate.py — balans eyniliklərinin və müstəqil yenidən hesablamanın yoxlanılması.

İki vəzifə:
  1. **Eynilik yoxlaması** (`identity_check`) — mühasibat eyniliyi BÜTÜN illər üçün, yalnız baza
     ilində deyil, yoxlanılır; hər hansı ildə fərq dözümdən böyükdürsə istisna qaldırılır və
     yoxlama uğursuz sayılır. Eynilik "təxminən qapanmalıdır" prinsipi qəbul edilmir.
  2. **Müstəqil yenidən hesablama** (`Verifier`) — təhvil verilən rəqəm ilkin girişlərdən
     asılı olmayan qısa kodla yenidən törədilir və nəşr olunmuş dəyərlə tutuşdurulur. Bu nümunə
     `az_forecast_model/src/oxlon/verify_oxlon.py` harness-inin davamıdır: hər yoxlama bir sətir
     PASS/FAIL kimi çap olunur, sonda ümumi nəticə qaytarılır.

Əlavə olaraq `reconcile_to_total` FR2/FR3 sahə cəmlərinin FR1 ÜDM-i ilə mütənasib
uzlaşdırılmasını həyata keçirir və **uzlaşdırmadan əvvəlki fərqi** açıq şəkildə hesabata verir.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple, Union

import numpy as np
import pandas as pd

try:
    from .econometrics import last_actual, years_of   # type: ignore
except Exception:   # modul birbaşa (paketsiz) idxal edildikdə
    from econometrics import last_actual, years_of    # type: ignore

__all__ = [
    "IdentityError", "IdentityResult", "identity_check", "check_no_lookahead",
    "reconcile_to_total", "Verifier",
]


class IdentityError(AssertionError):
    """Balans eyniliyi ən azı bir ildə qapanmadıqda qaldırılır."""


@dataclass
class IdentityResult:
    """Bir eynilik yoxlamasının nəticəsi."""
    name: str
    n_years: int
    max_abs_diff: float
    worst_year: Optional[int]
    tol: float
    passed: bool
    detail: str = ""

    def __str__(self) -> str:
        return (f"[{'PASS' if self.passed else 'FAIL'}] eynilik «{self.name}»: "
                f"{self.n_years} il, maksimal mütləq fərq {self.max_abs_diff:.3e} "
                f"({self.worst_year}), dözüm {self.tol:.1e}")


def _aligned(lhs, rhs, years: Optional[Iterable[int]] = None) -> pd.DataFrame:
    """İki seriyanı illər üzrə uzlaşdırır; illər dəsti üst-üstə düşməlidir."""
    L = pd.Series(lhs).astype(float)
    R = pd.Series(rhs).astype(float)
    L.index = years_of(L.index)
    R.index = years_of(R.index)
    if years is not None:
        yrs = [int(y) for y in years]
        miss_l = [y for y in yrs if y not in L.index]
        miss_r = [y for y in yrs if y not in R.index]
        if miss_l or miss_r:
            raise IdentityError(f"Yoxlama üçün tələb olunan illər çatışmır: "
                                f"sol tərəf {miss_l}, sağ tərəf {miss_r}.")
    else:
        only_l = sorted(set(L.index) - set(R.index))
        only_r = sorted(set(R.index) - set(L.index))
        if only_l or only_r:
            raise IdentityError(f"Eynilik BÜTÜN illər üçün yoxlanılmalıdır, lakin illər dəsti "
                                f"üst-üstə düşmür: yalnız solda {only_l}, yalnız sağda {only_r}.")
        yrs = sorted(set(L.index))
    df = pd.DataFrame({"lhs": L.reindex(yrs), "rhs": R.reindex(yrs)})
    nan_years = sorted(df.index[df.isna().any(axis=1)].tolist())
    if nan_years:
        raise IdentityError(f"Eynilik yoxlanışında boş (NaN) dəyərlər var: illər {nan_years}.")
    return df


def identity_check(name: str, lhs, rhs, tol: float = 1e-6,
                   years: Optional[Iterable[int]] = None, rel_tol: Optional[float] = None,
                   raise_on_fail: bool = True, verbose: bool = False) -> IdentityResult:
    """Balans eyniliyini HƏR il üçün yoxlayır: |sol − sağ| ≤ tol.

    Böyük miqyaslı göstəricilərdə (məsələn mln AZN ilə ÜDM) mütləq dözümlə yanaşı nisbi dözüm
    (`rel_tol`) verilə bilər: il ya mütləq, ya da nisbi şərti ödəməlidir. Uğursuzluq halında
    pozulan illər sadalanır və `IdentityError` qaldırılır (`raise_on_fail=False` olduqda yalnız
    nəticə obyekti qaytarılır).
    """
    df = _aligned(lhs, rhs, years)
    diff = (df["lhs"] - df["rhs"]).abs()
    ok_abs = diff <= float(tol)
    if rel_tol is not None:
        denom = df["rhs"].abs().replace(0.0, np.nan)
        ok_rel = (diff / denom) <= float(rel_tol)
        ok = ok_abs | ok_rel.fillna(False)
    else:
        ok = ok_abs
    bad = sorted(df.index[~ok].tolist())
    worst = int(diff.idxmax()) if len(diff) else None
    detail = ""
    if bad:
        detail = ("pozulan illər: " + ", ".join(
            f"{y}: sol={df.loc[y,'lhs']:.6g}, sağ={df.loc[y,'rhs']:.6g}, "
            f"fərq={diff.loc[y]:.6g}" for y in bad[:6]))
        if len(bad) > 6:
            detail += f" … (ümumilikdə {len(bad)} il)"
    res = IdentityResult(name=name, n_years=int(len(df)), max_abs_diff=float(diff.max()),
                         worst_year=worst, tol=float(tol), passed=not bad, detail=detail)
    if verbose or bad:
        print(str(res) + (f"\n        {detail}" if detail else ""))
    if bad and raise_on_fail:
        raise IdentityError(f"Eynilik «{name}» qapanmır — {detail}")
    return res


def check_no_lookahead(name: str, index, limit: Optional[int] = None) -> IdentityResult:
    """Nümunənin son faktiki ildən kənara çıxmadığını yoxlayır (qiymətləndirmə intizamı)."""
    lim = last_actual() if limit is None else int(limit)
    yrs = years_of(getattr(index, "index", index))
    bad = sorted(int(y) for y in np.unique(yrs[yrs > lim]))
    return IdentityResult(name=f"gələcəyə baxış yoxdur: {name}", n_years=int(yrs.size),
                          max_abs_diff=float(max(bad) - lim if bad else 0.0),
                          worst_year=(bad[-1] if bad else None), tol=0.0, passed=not bad,
                          detail=(f"nümunəyə düşən proqnoz illəri: {bad}" if bad else
                                  f"nümunə {lim} ilində bitir"))


def reconcile_to_total(parts: pd.DataFrame, total: pd.Series, name: str = "sahə cəmi",
                       report_only: bool = False) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Hissələri (sahələr/komponentlər) məcmu göstəriciyə MÜTƏNASİB uzlaşdırır.

    Qaytarılır: (uzlaşdırılmış cədvəl, fərq hesabatı). Fərq hesabatı uzlaşdırmadan ƏVVƏLKİ
    fərqi — həm mütləq, həm faizlə — hər il üçün saxlayır; bu rəqəm hesabatda açıq göstərilir,
    çünki uzlaşdırma özü fərqi gizlədir. `report_only=True` olduqda düzəliş aparılmır.
    """
    P = parts.copy().astype(float)
    P.index = years_of(P.index)
    T = pd.Series(total).astype(float)
    T.index = years_of(T.index)
    yrs = sorted(set(P.index) & set(T.index))
    if not yrs:
        raise IdentityError(f"«{name}»: hissələr və məcmu göstərici üçün ortaq il yoxdur.")
    P = P.loc[yrs]
    T = T.loc[yrs]
    s = P.sum(axis=1)
    gap = pd.DataFrame({
        "year": yrs, "sum_parts": s.values, "total": T.values,
        "gap_abs": (T - s).values,
        "gap_pct": np.where(T.values != 0, (T.values - s.values) / T.values * 100.0, np.nan),
    }).set_index("year")
    if report_only:
        return P, gap
    factor = (T / s.replace(0.0, np.nan))
    rec = P.mul(factor, axis=0)
    identity_check(f"{name} → məcmu göstərici", rec.sum(axis=1), T, tol=1e-6, rel_tol=1e-9)
    return rec, gap


@dataclass
class Verifier:
    """Müstəqil yenidən hesablama harness-i: yoxlamaları toplayır, PASS/FAIL hesabatı çap edir.

    İstifadə nümunəsi:
        v = Verifier("FR1 — ÜDM nüvəsi")
        v.identity("ÜDM = neft + qeyri-neft + xalis vergilər", lhs, rhs, tol=1e-6)
        v.recompute("2026 ÜDM real artımı", shipped=2.34, fn=lambda: _yenidən_hesabla(), tol=0.01)
        v.no_lookahead("qeyri-neft tələb tənliyi", nümunə.index)
        v.assert_all()
    """
    title: str
    checks: List[Tuple[str, bool, str]] = field(default_factory=list)

    # --- əsas qeydiyyat ---
    def ok(self, name: str, cond: bool, detail: str = "") -> bool:
        self.checks.append((name, bool(cond), detail))
        return bool(cond)

    def close(self, name: str, got, want, tol: float = 1e-6, detail: str = "") -> bool:
        """Skalyar və ya seriya müqayisəsi verilmiş dözümlə."""
        try:
            if np.isscalar(got) and np.isscalar(want):
                d = np.array([abs(float(got) - float(want))])
            else:
                g = pd.Series(got).astype(float)
                w = pd.Series(want).astype(float)
                yrs = sorted(set(g.index) & set(w.index))
                if not yrs:
                    raise ValueError("müqayisə üçün ortaq indeks yoxdur")
                d = np.abs(g.reindex(yrs).values - w.reindex(yrs).values)
            worst = float(np.nanmax(d)) if d.size else float("nan")
            return self.ok(name, bool(worst <= tol),
                           detail or f"maksimal fərq {worst:.3e} ≤ {tol:.1e}")
        except Exception as exc:
            return self.ok(name, False, f"müqayisə mümkün olmadı: {exc}")

    def identity(self, name: str, lhs, rhs, tol: float = 1e-6,
                 rel_tol: Optional[float] = None) -> bool:
        try:
            r = identity_check(name, lhs, rhs, tol=tol, rel_tol=rel_tol, raise_on_fail=False)
            return self.ok(f"eynilik: {name}", r.passed,
                           r.detail or f"{r.n_years} il, maksimal fərq {r.max_abs_diff:.3e}")
        except IdentityError as exc:
            return self.ok(f"eynilik: {name}", False, str(exc))

    def recompute(self, name: str, shipped, fn: Callable[[], object], tol: float = 1e-6,
                  detail: str = "") -> bool:
        """Nəşr olunmuş dəyəri ilkin girişlərdən müstəqil şəkildə yenidən törədib tutuşdurur."""
        try:
            again = fn()
        except Exception as exc:
            return self.ok(f"yenidən hesablama: {name}", False, f"hesablama xətası: {exc}")
        return self.close(f"yenidən hesablama: {name}", again, shipped, tol=tol, detail=detail)

    def no_lookahead(self, name: str, index, limit: Optional[int] = None) -> bool:
        r = check_no_lookahead(name, index, limit=limit)
        return self.ok(r.name, r.passed, r.detail)

    # --- hesabat ---
    @property
    def failed(self) -> List[Tuple[str, bool, str]]:
        return [c for c in self.checks if not c[1]]

    def report(self, verbose: bool = True) -> bool:
        npass = sum(1 for _, c, _ in self.checks if c)
        if verbose:
            line = "=" * 62
            print(line)
            print(f"  YOXLAMA: {self.title}")
            print(line)
            for nm, c, det in self.checks:
                print(f"  [{'PASS' if c else 'FAIL'}] {nm}")
                if det:
                    print(f"         {det}")
            print("-" * 62)
            print(f"  {npass}/{len(self.checks)} yoxlama uğurlu")
        return npass == len(self.checks)

    def assert_all(self, verbose: bool = True) -> None:
        """Bütün yoxlamalar uğurlu deyilsə istisna qaldırır — boru xətti səssiz davam etmir."""
        ok_all = self.report(verbose=verbose)
        if not ok_all:
            names = "; ".join(nm for nm, c, _ in self.checks if not c)
            raise IdentityError(f"«{self.title}»: uğursuz yoxlamalar — {names}")
