# -*- coding: utf-8 -*-
"""ministry_spec.py — «Nazirlik spesifikasiyası» ssenarisi (FR13).

Nazirliyin öz makroekonometrik nümunə-modelinin **92 sətirlik tənlik kataloqu** bu paketin
içində İKİNCİ, İŞLƏK spesifikasiya kimi həll olunur: onun ÖZ əmsalları, onun ÖZ dəyişən
adları, birləşdirilmiş məlumat qatı üzərində. Bu, modelin əvəzlənməsi deyil — ayrıca
**ssenaridir** və eyni geriyə doğru sınaq cədvəlində, eyni təsadüfi gəzişmə (RW) etalonuna
qarşı qiymətləndirilir.

METODOLOJİ MƏNŞƏ (nə haradan gəlir):
  · tənlik sintaksisinin analizatoru və qiymətləndiricisi — EViews-tipli tənlik
    oxuyucusu metodikası əsasında, burada Azərbaycan dilində sənədləşdirilmiş şəkildə
    yazılıb; əlavə olaraq `@SUM`-un tək-illik forması və sıfıra bölmə mühafizəsi
    dəqiqləşdirilib;
  · blok-rekursiv Gauss–Seidel həlledicisi və partlayan yol mühafizəsi — standart
    sahə-həlledici metodika əsasında; ekzogen yolların mənbəyi kimi əvvəlki yanaşmalardan
    fərqli olaraq sabit uzatmalar (m2 6.644, brent ~69) deyil, bu modul
    **bu paketin öz nəşr olunmuş proqnoz yollarını** sürücü kimi götürür;
  · tənlik kataloqu — `data/ministry_equations_catalog.csv`; Nazirliyin öz tənlik
    kataloqunun DƏYİŞDİRİLMƏMİŞ nüsxəsidir;
  · sıralar — `data_layer` §13 (`moe_spec_panel`), Nazirliyin nümunə-model iş kitablarından
    YALNIZ faktiki (≤2024) sütunlar, etiket yoxlaması və provenans qeydi ilə.

TƏSNİFAT (kataloqun sətir-sətir təsnifatı — bağlayıcı): 92 sətirdən **67-si** təqdim olunduğu kimi işləyir
(A), **14-ü** sənədləşdirilmiş əvəzedici və ya qısa nümunə ilə (B), **8-i** kataloqda tərifi
verilməyən xəta-korreksiyası səviyyəsini gözləyir (C), **3-ü** dörd tədiyə balansı
dəyişəninin tərifini gözləyir (D). D3 qərarına görə `ECM_MS_TP1` yeganə istisnadır: o,
`ECM_MS_TP` ilə eyni qəbul edilir və **BAYRAQLANIR**; qalan 7 səviyyə üçün sətir
«tərif gözlənilir» statusu ilə göstərilir. Kataloqun 49-cu sətri sintaksis deyil, mətndir —
statusu «mətn dəqiqləşdirilməlidir».
"""
from __future__ import annotations

import ast
import math
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Callable, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

import numpy as np
import pandas as pd

_SRC_DIR = os.path.dirname(os.path.abspath(__file__))
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

from config import (  # noqa: E402
    FIRST_FORECAST, LAST_ACTUAL, LAST_FORECAST, ASSUMPTIONS, OUT,
    P_MINISTRY_CATALOG,
)

# ==========================================================================================
# 1. EViews ifadə analizatoru — port (`oxlon/eviews.py`)
# ==========================================================================================
# Kataloqda faktiki işlənən sintaksis: DLOG(x) · LOG(x) · D(x) · x(-1) gecikmələri ·
# @BEFORE("Y") · @DURING("Y1 Y2") · @DURING("Y1+Y2") · @AFTER("Y") · (T=Y)/(T>Y)/(T<Y) ·
# @SUM(ifadə,"Y1 Y2") · ECM_* terminləri · sabit `c`.
# Üsul: EViews sintaksisi düzgün Python ifadəsinə NORMALLAŞDIRILIR, `ast` ilə təhlil edilir,
# sonra ağac ZAMAN SÜRÜŞMƏSİNƏ həssas qiymətləndirici ilə gəzilir — beləliklə DLOG/D
# gecikmiş ALT-İFADƏLƏR üzərində düzgün işləyir (məsələn DLOG(MGNO(-1)*ER(-1)/CPI(-1))).

FUNCS = {"DLOG", "LOG", "D", "BEFORE", "DURING", "AFTER", "SUM", "EXP"}
RESERVED = FUNCS | {"T", "t", "c"}


def normalise(rhs: str) -> str:
    """EViews ifadəsini Python-un təhlil edə biləcəyi sətrə çevirir."""
    s = str(rhs).strip()
    s = s.replace("@", "")                       # @BEFORE -> BEFORE
    s = s.replace("=<", "<=").replace("=>", ">=")
    # (T=2015) / T=2015 / t=2015 -> T==2015 (yalnız dördrəqəmli il; <=, >=, == toxunulmur)
    s = re.sub(r'(?<![<>=!])\b[Tt]\s*=\s*(\d{4})', r'T==\1', s)
    s = s.replace("====", "==")
    # MÜHÜM DƏQİQLƏŞDİRMƏ (mənbə portda YOX İDİ): kataloqda dummi bəzən MÖTƏRİZƏSİZ yazılır
    # (`… + -0.151702 * T=1998 + 0.150852 * (T=2011) + …`). Python-da müqayisə operatoru
    # toplama və vurmadan ZƏİF bağlandığı üçün belə ifadə BÜTÜN sağ tərəfi tək bir müqayisəyə
    # çevirir və nəticə 0/1 olur — 40, 42, 17-ci sətirlərdə tənlik səssizcə «sıfır artım»
    # verirdi. Ona görə hər `T <op> il` ifadəsi mötərizəyə alınır.
    s = re.sub(r'\b[Tt]\s*(==|<=|>=|<|>)\s*(\d{4})', r'(T\1\2)', s)
    return s


_STRING_LITERAL_RE = re.compile(r'"[^"]*"|\'[^\']*\'')
_RESERVED_UPPER = {r.upper() for r in RESERVED}


def variables_in(rhs: str) -> set:
    """İfadədə istinad olunan dəyişən adları (funksiya adları və `T`/`c` istisna olmaqla).

    Mənbə koddan İKİ DƏQİQLƏŞDİRMƏ (hər ikisi kataloqda faktiki rast gəlinən hallara görə):
      (1) sətir sabitləri (`@DURING("2004_2008")`, `@SUM(…,"1995_2022")`) əvvəlcə çıxarılır —
          əks halda `_2008` kimi parçalar dəyişən adı sayılırdı;
      (2) funksiya adları BÖYÜK-KİÇİK HƏRFDƏN ASILI OLMADAN tanınır — kataloqda `dLOG` yazılışı
          da işlənir.
    """
    s = _STRING_LITERAL_RE.sub(" ", normalise(rhs))
    out = set()
    for m in re.finditer(r'[A-Za-z_][A-Za-z0-9_]*', s):
        name = m.group(0)
        if name.upper() in _RESERVED_UPPER:
            continue
        out.add(name)
    return out


def _during_set(arg: str) -> set:
    """`@DURING("2005 2007")` → {2005,2006,2007}; `@DURING("2003+2009")` → {2003,2009}."""
    nums = [int(x) for x in re.findall(r'\d{4}', arg)]
    if not nums:
        return set()
    if "+" in arg:
        return set(nums)
    if len(nums) == 2:
        return set(range(nums[0], nums[1] + 1))
    return set(nums)


def _str_of(node) -> str:
    if isinstance(node, ast.Constant):
        return str(node.value)
    return ""


class _Ev:
    """Normallaşdırılmış ağacı `y0` ilində, `shift` gecikməsi (≤0) ilə qiymətləndirir."""

    def __init__(self, data: Mapping[str, Mapping[int, float]], y0: int,
                 scale: Optional[Mapping[str, float]] = None):
        self.D, self.y0 = data, int(y0)
        self.scale = dict(scale or {})

    def _resolve(self, name: str):
        """Adı BÖYÜK-KİÇİK HƏRFDƏN ASILI OLMADAN tapır.

        Kataloq eyni sıranı iki yazılışda çağırır (`ecm_rxgno` təyin edilir, `ECM_RXGNO`
        istinad olunur; `log(rxgno)` və `LOG(RXGNO)`). Mənbə mühərrik bu sətirlərdə səssizcə
        uğursuz olurdu; port adları normallaşdırır və bunu SƏNƏDLƏŞDİRİR.
        """
        s = self.D.get(name)
        if s is not None:
            return s
        up = name.upper()
        for k in self.D:
            if k.upper() == up:
                return self.D[k]
        return None

    def get(self, name: str, year: int) -> float:
        s = self._resolve(name)
        if s is None or year not in s or s[year] is None:
            raise KeyError(f"{name}@{year}")
        v = float(s[year])
        if not math.isfinite(v):
            raise KeyError(f"{name}@{year}")
        k = self.scale.get(name)
        if k is None:
            k = self.scale.get(name.upper(), 1.0)
        return v * float(k)

    def _span(self) -> Tuple[int, int]:
        yrs = [y for s in self.D.values() for y in s]
        return (min(yrs), max(yrs)) if yrs else (self.y0, self.y0)

    def ev(self, node, shift: int = 0) -> float:
        if isinstance(node, ast.Expression):
            return self.ev(node.body, shift)
        if isinstance(node, ast.BinOp):
            a = self.ev(node.left, shift)
            b = self.ev(node.right, shift)
            if isinstance(node.op, ast.Add):
                return a + b
            if isinstance(node.op, ast.Sub):
                return a - b
            if isinstance(node.op, ast.Mult):
                return a * b
            if isinstance(node.op, ast.Div):
                if b == 0:
                    raise KeyError("sıfıra bölmə")
                return a / b
            if isinstance(node.op, ast.Pow):
                return a ** b
            raise ValueError(f"dəstəklənməyən əməliyyat: {node.op}")
        if isinstance(node, ast.UnaryOp):
            v = self.ev(node.operand, shift)
            return -v if isinstance(node.op, ast.USub) else v
        if isinstance(node, ast.Constant):
            return node.value
        if isinstance(node, ast.Compare):
            left = self.ev(node.left, shift)
            right = self.ev(node.comparators[0], shift)
            op = node.ops[0]
            r = (left == right if isinstance(op, ast.Eq) else
                 left > right if isinstance(op, ast.Gt) else
                 left < right if isinstance(op, ast.Lt) else
                 left >= right if isinstance(op, ast.GtE) else
                 left <= right if isinstance(op, ast.LtE) else
                 left != right)
            return 1.0 if r else 0.0
        if isinstance(node, ast.Name):
            if node.id in ("T", "t"):
                return float(self.y0 + shift)
            if node.id == "c":
                return 1.0
            return self.get(node.id, self.y0 + shift)
        if isinstance(node, ast.Call):
            fid = node.func.id if isinstance(node.func, ast.Name) else None
            f = (fid or "").upper()
            if f == "DLOG":
                a = self.ev(node.args[0], shift)
                b = self.ev(node.args[0], shift - 1)
                if a <= 0 or b <= 0:
                    raise KeyError("DLOG: müsbət olmayan arqument")
                return math.log(a) - math.log(b)
            if f == "LOG":
                a = self.ev(node.args[0], shift)
                if a <= 0:
                    raise KeyError("LOG: müsbət olmayan arqument")
                return math.log(a)
            if f == "EXP":
                return math.exp(self.ev(node.args[0], shift))
            if f == "D":
                return self.ev(node.args[0], shift) - self.ev(node.args[0], shift - 1)
            if f == "BEFORE":
                return 1.0 if (self.y0 + shift) < int(re.findall(r'\d{4}', _str_of(node.args[0]))[0]) else 0.0
            if f == "AFTER":
                return 1.0 if (self.y0 + shift) >= int(re.findall(r'\d{4}', _str_of(node.args[0]))[0]) else 0.0
            if f == "DURING":
                return 1.0 if (self.y0 + shift) in _during_set(_str_of(node.args[0])) else 0.0
            if f == "SUM":
                # `@SUM(ifadə,"Y1 Y2")` — verilmiş dövr üzrə cəm. Dövr VERİLMƏYİBSƏ (kataloqun
                # 34-cü sətri: `@SUM(OPWTI*(T=2015))`) EViews bütün nümunə üzrə cəmləyir, yəni
                # nəticə 2015-ci ilin dəyəridir. Mənbə port bunu cari illə məhdudlaşdırırdı və
                # ifadə sıfıra bölmə ilə dağılırdı; burada bütün panel dövrü götürülür.
                yrs = [int(x) for x in re.findall(r'\d{4}', _str_of(node.args[1]))] if len(node.args) > 1 else []
                if len(yrs) >= 2:
                    rng = range(yrs[0], yrs[-1] + 1)
                else:
                    lo, hi = self._span()
                    rng = range(lo, hi + 1)
                tot = 0.0
                for yy in rng:
                    try:
                        tot += _Ev(self.D, yy, self.scale).ev(node.args[0], 0)
                    except KeyError:
                        pass
                return tot
            if fid is not None:           # gecikmə ilə dəyişən istinadı, məsələn CPI(-1)
                lag = int(self.ev(node.args[0], 0)) if node.args else 0
                return self.get(fid, self.y0 + shift + lag)
        raise ValueError(f"tanınmayan qovşaq: {ast.dump(node)[:60]}")


def repair(lhs: str, rhs: str, name: str = "") -> Tuple[str, str]:
    """Kataloq çıxarılışının MƏLUM yazı qüsurlarını düzəldir (mənbədə olduğu kimi sənədlənir).

    Üç hal: `eq_dmgno_fin`-də mötərizənin yeri; `W_WATE` tənliyində `LOG CPI` (mötərizəsiz);
    `TIKINTI` sətrində ikiqat `= DLOG(...) =` ön şəkilçisi. Başqa heç bir düzəliş edilmir —
    əmsallar toxunulmazdır.
    """
    n = name or ""
    if "dmgno_fin" in n:
        rhs = rhs.replace("DLOG((INO/ER(-1))*(T=2009)", "DLOG(INO/ER(-1))*(T=2009)")
    if "W_WATE" in lhs:
        rhs = rhs.replace("LOG CPI", "LOG(CPI)")
    if "TIKINTI" in lhs and " = " in rhs:
        rhs = rhs.split(" = ")[-1]
    return lhs, rhs


class EViewsEq:
    """Bir kataloq sətri: sol tərəfin çevirməsi + hədəf adı + sağ tərəfin ağacı."""

    def __init__(self, lhs: str, rhs: str, name: str = ""):
        lhs, rhs = repair(lhs, rhs, name)
        # Bəzi sətirlərdə çıxarılış zamanı sol tərəfə tam tənlik yapışıb
        # (`DLOG(SERVICES_PAID) = -0.0055 + …`) — sol tərəf birinci `=`-ə qədərdir.
        lhs = str(lhs).split("=")[0]
        self.lhs_raw, self.rhs_raw, self.name = lhs, rhs, name
        self.rhs_ast = ast.parse(normalise(rhs), mode="eval")
        m = re.match(r'\s*(DLOG|LOG|D)\s*\(\s*(.+?)\s*\)\s*$', lhs.strip(), re.I)
        if m:
            self.transform, inner = m.group(1).upper(), m.group(2).strip()
        else:
            self.transform, inner = "LEVEL", lhs.strip()
        self.lhs_inner = inner
        # Sol tərəf sadə ad ola bilər (`RVA_CONST`) və ya İFADƏ ola bilər
        # (`MGNO/CPICMTP`, `MGNO/(CPIMTP/NERMTP)`, `M1/MB`). İkinci halda tənlik həmin
        # nisbəti proqnozlaşdırır; həll edilən DƏYİŞƏN isə ifadədəki BİRİNCİ addır və o,
        # nisbətdən geri çıxarılır (ifadə həmin ad üzrə xəttidir).
        if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', inner):
            self.composite, self.lhs_ast, self.target = False, None, inner
        else:
            self.composite = True
            self.lhs_ast = ast.parse(normalise(inner), mode="eval")
            names = [n for n in re.findall(r'[A-Za-z_][A-Za-z0-9_]*', normalise(inner))
                     if n.upper() not in _RESERVED_UPPER]
            self.target = names[0] if names else inner
            self.lhs_expr = inner

    # ---- qiymətləndirmə ----------------------------------------------------------------
    #: Bu tənliyə XÜSUSİ vahid uyğunlaşdırması (bax `EQUATION_UNIT_FIX`).
    scale: Dict[str, float] = {}

    def eval_rhs(self, year: int, data) -> float:
        return _Ev(data, year, self.scale).ev(self.rhs_ast, 0)

    def _lhs_value(self, year: int, data) -> float:
        """Sol tərəfin (ad və ya ifadə) həmin ildəki dəyəri."""
        if not self.composite:
            return _Ev(data, year, self.scale).get(self.target, year)
        return _Ev(data, year, self.scale).ev(self.lhs_ast, 0)

    def level(self, year: int, data) -> Optional[float]:
        """Sağ tərəfi qiymətləndirib sol tərəfin çevirməsini GERİ açır (səviyyə qaytarır).

        Mürəkkəb sol tərəfdə (nisbət) əvvəlcə nisbətin səviyyəsi tapılır, sonra hədəf
        dəyişən nisbətdən geri çıxarılır: ifadə hədəf üzrə xətti olduğuna görə miqyas əmsalı
        hədəfi müvəqqəti 1-ə bərabər tutmaqla hesablanır.
        """
        val = self.eval_rhs(year, data)
        try:
            prev = self._lhs_value(year - 1, data)
        except Exception:
            prev = None
        if self.transform == "DLOG":
            if prev is None or prev <= 0:
                return None
            cur = float(prev) * math.exp(val)
        elif self.transform == "LOG":
            cur = math.exp(val)
        elif self.transform == "D":
            if prev is None:
                return None
            cur = float(prev) + val
        else:
            cur = val
        if not self.composite:
            return cur
        store = data.get(self.target)
        if store is None:
            return None
        keep, store[year] = store.get(year), 1.0
        try:
            k = _Ev(data, year, self.scale).ev(self.lhs_ast, 0)
        except Exception:
            k = None
        finally:
            if keep is None:
                store.pop(year, None)
            else:
                store[year] = keep
        if not k:
            return None
        return cur / k

    @property
    def rhs_vars(self) -> set:
        return variables_in(self.rhs_raw)


# ==========================================================================================
# 2. Kataloq və sətir-sətir təsnifat
# ==========================================================================================
# Təsnifatın sətir siyahıları BAĞLAYICIDIR — burada yenidən törədilmir, köçürülür.
BUCKET_B_ROWS = (1, 2, 25, 26, 27, 33, 34, 43, 47, 53, 61, 73, 76, 78)
BUCKET_C_ROWS = (3, 9, 13, 23, 30, 35, 75, 81)
BUCKET_D_ROWS = (8, 15, 16)
PROSE_ROWS = (49,)                       # sintaksis deyil, mətn — «mətn dəqiqləşdirilməlidir»

# D3 / R17: `ECM_MS_TP1` `ECM_MS_TP` ilə eyni qəbul edilir və BAYRAQLANIR; qalan 7 ad gözləyir.
ECM_ASSUMED_EQUIVALENT = {"ECM_MS_TP1": "ECM_MS_TP"}
DANGLING_ECM = ("ECM_MGNO_NEW", "ECM_MGO12", "ECM_RIDNS2020", "ECM_RIEANEW2020",
                "ECM_RIEA_NEW_2020", "ECM_RVA_ACCOM2020", "ECM_W_STMANAG2020")
# Bucket D-nin dörd tərifsiz tədiyə balansı adı (təsnifatın D səbəti).
UNDEFINED_BOP_VARS = ("OI_L_OS", "MSC", "IFOC", "RACBAR")

# Vahid uyğunlaşdırması — MƏLUMATLA ƏSASLANDIRILMIŞ, əmsala toxunmadan.
# Kataloqun 43-cü sətri (`DLOG(HC)`) faiz dərəcəsini ONLUQ kəsr kimi işlədir, halbuki iş
# kitabının `Mon` vərəqindəki `nird` sırası FAİZ vahidindədir. Fərq nümunədaxili uyğunluqla
# birmənalı ayrılır: faiz vahidində orta mütləq xəta 202.1 %, onluq kəsrdə 3.1 % (2005–2024,
# n = 20). Eyni yoxlama 19, 20 və 58-ci sətirlərdə ƏKS nəticə verir (faiz vahidi daha yaxşıdır,
# 3.97/9.71/8.65 % qarşı 6.43/12.01/12.60 %), ona görə uyğunlaşdırma YALNIZ 43-cü sətrə tətbiq
# olunur və hesabatda cədvəl şəklində göstərilir. Əmsallar dəyişdirilmir.
EQUATION_UNIT_FIX: Dict[int, Dict[str, float]] = {43: {"NIRD": 0.01}}
EQUATION_UNIT_FIX_NOTE = (
    "43-cü sətir: `NIRD` onluq kəsrə çevrilir (÷100) — nümunədaxili uyğunluq bunu birmənalı "
    "göstərir (OMFX 202.1 % → 3.1 %); əmsallara toxunulmur")

STATUS_A = "işə salınır"
STATUS_B = "işə salınır — əvəzedici/qısa nümunə (bayraqlı)"
STATUS_C_FLAG = "işə salınır — ECM ekvivalentliyi FƏRZ EDİLİB (bayraqlı)"
STATUS_C = "tərif gözlənilir"
STATUS_D = "bloklanıb — dəyişən tərifi yoxdur"
STATUS_PROSE = "mətn dəqiqləşdirilməlidir"

BUCKET_NOTE = {
    "A": "birləşdirilmiş məlumat qatında təqdim olunduğu kimi işləyir",
    "B": "sənədləşdirilmiş əvəzedici və ya qısa/birləşdirilmiş nümunə ilə işləyir",
    "C": "kataloqda tərifi verilməyən xəta-korreksiyası səviyyəsini gözləyir",
    "D": "tədiyə balansı dəyişənlərinin tərifini gözləyir",
}


def _bucket(no: int) -> str:
    if no in BUCKET_D_ROWS:
        return "D"
    if no in BUCKET_C_ROWS:
        return "C"
    if no in BUCKET_B_ROWS:
        return "B"
    return "A"


def _base_target(lhs: str) -> str:
    s = str(lhs or "").split("=")[0].strip()
    m = re.match(r'^(DLOG|LOG|D)\s*\(\s*(.+?)\s*\)\s*$', s, re.I)
    return _base_target(m.group(2)) if m else s


@dataclass
class SpecRow:
    """Kataloqun bir sətri + onun statusu."""
    no: int
    workbook: str
    sheet: str
    eq_name: str
    description: str
    lhs: str
    rhs: str
    target: str
    transform: str
    bucket: str
    status: str
    note_az: str
    eq: Optional[EViewsEq] = None
    parses: bool = True
    is_ecm: bool = False
    dangling: Tuple[str, ...] = ()


def load_catalog(path: str = P_MINISTRY_CATALOG) -> List[SpecRow]:
    """Nazirliyin 92 sətirlik kataloqunu oxuyur və hər sətri sətir-sətir təsnifata görə təsnif edir."""
    raw = pd.read_csv(path).fillna("")
    rows: List[SpecRow] = []
    for i, r in enumerate(raw.itertuples(index=False), start=1):
        lhs, rhs = str(r.lhs), str(r.rhs)
        tgt = _base_target(lhs)
        bucket = _bucket(i)
        eq, parses = None, True
        try:
            eq = EViewsEq(lhs, rhs, str(r.eq_name))
            if i in EQUATION_UNIT_FIX:
                eq.scale = dict(EQUATION_UNIT_FIX[i])
            tgt = eq.target
        except Exception:
            parses = False
        dang = tuple(sorted({v for v in variables_in(rhs) if v.upper().startswith("ECM")
                             and v.upper() in {d.upper() for d in DANGLING_ECM}}))
        is_ecm = tgt.upper().startswith("ECM")
        if i in PROSE_ROWS or not parses:
            status, note = STATUS_PROSE, (
                "kataloq sətri EViews sintaksisi deyil, MƏTN kimi çıxarılıb "
                "(«before 2005 * …»); mətnin dəqiqləşdirilməsi tələb olunur")
        elif bucket == "D":
            miss = sorted({v for v in (variables_in(rhs) | {tgt})
                           if v.upper() in {u.upper() for u in UNDEFINED_BOP_VARS}})
            status, note = STATUS_D, (
                "tərifi verilməyən dəyişən(lər): " + ", ".join(miss) +
                " — dördü də MOE BOP `eq14` vərəqindədir, məlumat sətri yoxdur")
        elif bucket == "C":
            if dang:
                status, note = STATUS_C, "gözlənilən səviyyə tərifi: " + ", ".join(dang)
            else:
                status, note = STATUS_C, "kataloqda tərifi verilməyən ECM səviyyəsi"
        elif bucket == "B":
            status, note = STATUS_B, BUCKET_NOTE["B"]
        else:
            status, note = STATUS_A, BUCKET_NOTE["A"]
        if i in EQUATION_UNIT_FIX:
            note = (note + "; " if note else "") + EQUATION_UNIT_FIX_NOTE
        if eq is not None and eq.target.lower() == "eq01":
            note = ("sol tərəfdə sıra adı deyil, TƏNLİK ADI (`eq01`) durur — forması 48-ci "
                    "sətirlə eynidir (tikintinin real əlavə dəyəri); hədəf adı "
                    "dəqiqləşdirilməlidir, ona görə ssenaridə həll edilmir")
        # D3 istisnası: ECM_MS_TP1 ekvivalent qəbul edilir → sətir İŞLƏYİR, amma bayraqlanır.
        if any(v.upper() in {k.upper() for k in ECM_ASSUMED_EQUIVALENT} for v in variables_in(rhs)):
            status = STATUS_C_FLAG
            note = ("`ECM_MS_TP1` kataloqda təyin olunmayıb; D3/R17 qərarına görə `ECM_MS_TP` "
                    "ilə EYNİ qəbul edilir və nəticə BAYRAQLANIR")
        rows.append(SpecRow(no=i, workbook=str(r.workbook), sheet=str(r.sheet),
                            eq_name=str(r.eq_name), description=str(r.description),
                            lhs=lhs, rhs=rhs, target=tgt,
                            transform=(eq.transform if eq else "—"),
                            bucket=bucket, status=status, note_az=note,
                            eq=eq, parses=parses, is_ecm=is_ecm, dangling=dang))
    return rows


def catalog_frame(rows: Optional[List[SpecRow]] = None) -> pd.DataFrame:
    """Kataloq + status cədvəli (AZ sütun adları ilə hesabat üçün)."""
    rows = rows or load_catalog()
    return pd.DataFrame([{
        "№": r.no, "iş kitabı": r.workbook, "vərəq": r.sheet, "tənlik": r.eq_name.split(",")[0],
        "təsvir": r.description, "hədəf": r.target, "çevirmə": r.transform,
        "səbət": r.bucket, "status": r.status, "qeyd": r.note_az,
    } for r in rows])


# Rejim dummilərinin AKTİVLİK yoxlaması. Sıralar iş kitablarında 1995/1999/2001-dən başladığı
# üçün bəzi dummilər mövcud nümunədə HEÇ VAXT 1 olmur (həmişə 0), bəziləri isə HƏMİŞƏ 1 olur və
# faktiki olaraq sabitə çevrilir. Sabit əmsallı ssenaridə hər ikisi sənədləşdirilməlidir.
DUMMY_RE = re.compile(
    r'BEFORE\("(\d{4})"\)|AFTER\("(\d{4})"\)|DURING\("([^"]*)"\)'
    r'|[Tt]\s*(==|<=|>=|<|>)\s*(\d{4})')


def _dummy_state(kind: str, arg, first_year: int, last_year: int) -> Optional[str]:
    """Dummi verilmiş nümunə dövründə həmişə 0-dırsa «həmişə 0», həmişə 1-dirsə «həmişə 1»."""
    yrs = range(first_year, last_year + 1)
    if kind == "BEFORE":
        vals = {1 if y < arg else 0 for y in yrs}
    elif kind == "AFTER":
        vals = {1 if y >= arg else 0 for y in yrs}
    elif kind == "DURING":
        vals = {1 if y in arg else 0 for y in yrs}
    else:
        op, yy = kind, arg
        vals = {1 if ((y == yy and op == "==") or (y <= yy and op == "<=") or
                      (y >= yy and op == ">=") or (y < yy and op == "<") or
                      (y > yy and op == ">")) else 0 for y in yrs}
    if vals == {0}:
        return "həmişə 0"
    if vals == {1}:
        return "həmişə 1"
    return None


def inactive_pre_sample_dummies(rows: Optional[List[SpecRow]] = None,
                                panel: Optional[Mapping[str, Mapping[int, float]]] = None,
                                last_year: int = 2024) -> pd.DataFrame:
    """Mövcud nümunədə heç vaxt dəyişməyən rejim dummiləri — sətir-sətir.

    Hər tənlik üçün nümunənin başlanğıcı onun ÖZ dəyişənlərindən hesablanır (ən gec başlayan
    sıra), ona görə nəticə fərziyyə deyil, panelin faktı ilə müəyyən olunur.
    """
    rows = rows or load_catalog()
    panel = panel if panel is not None else history_panel()
    out = []
    for r in rows:
        if not r.parses:
            continue
        starts = []
        for v in (r.eq.rhs_vars | {r.target}):
            s = _lookup(panel, v)
            if s:
                starts.append(min(s))
        if not starts:
            continue
        first_year = max(starts)
        zero, one = [], []
        for m in DUMMY_RE.finditer(normalise(r.rhs)):
            if m.group(1):
                st = _dummy_state("BEFORE", int(m.group(1)), first_year, last_year)
                txt = f'@BEFORE("{m.group(1)}")'
            elif m.group(2):
                st = _dummy_state("AFTER", int(m.group(2)), first_year, last_year)
                txt = f'@AFTER("{m.group(2)}")'
            elif m.group(3) is not None:
                st = _dummy_state("DURING", _during_set(m.group(3)), first_year, last_year)
                txt = f'@DURING("{m.group(3)}")'
            else:
                st = _dummy_state(m.group(4), int(m.group(5)), first_year, last_year)
                txt = f"(T{m.group(4)}{m.group(5)})"
            if st == "həmişə 0":
                zero.append(txt)
            elif st == "həmişə 1":
                one.append(txt)
        if zero or one:
            out.append({"№": r.no, "hədəf": r.target, "nümunənin başlanğıcı": first_year,
                        "həmişə 0 (qeyri-aktiv)": ", ".join(sorted(set(zero))) or "—",
                        "həmişə 1 (sabitə çevrilir)": ", ".join(sorted(set(one))) or "—"})
    return pd.DataFrame(out)


def branch_dummy_table(rows: Optional[List[SpecRow]] = None,
                       panel: Optional[Mapping[str, Mapping[int, float]]] = None,
                       first_row: int = 61, last_row: int = 90) -> pd.DataFrame:
    """Fəaliyyət növləri bloku: hər tənliyin nümunə başlanğıcı və ən erkən rejim həddi.

    Təsnifat bu blok üçün nümunənin 1999/2001-dən başladığını qeyd edir; cədvəl həmin
    faktı tənlik-tənlik göstərir və ən erkən `@BEFORE` həddinin neçə nümunə ilini əhatə etdiyini
    açır.
    """
    rows = rows or load_catalog()
    panel = panel if panel is not None else history_panel()
    out = []
    for r in rows:
        if not (first_row <= r.no <= last_row and r.parses):
            continue
        starts = [min(s) for s in (_lookup(panel, v) for v in (r.eq.rhs_vars | {r.target})) if s]
        ys = [int(x) for x in re.findall(r'BEFORE\("(\d{4})"\)', normalise(r.rhs))]
        y0 = max(starts) if starts else None
        y1 = min(ys) if ys else None
        out.append({"№": r.no, "hədəf": r.target, "nümunənin başlanğıcı": y0,
                    "ən erkən rejim həddi": y1,
                    "həmin rejimin əhatə etdiyi il": (None if (y0 is None or y1 is None)
                                                     else max(0, y1 - y0))})
    return pd.DataFrame(out)


# ==========================================================================================
# 3. Birləşdirilmiş dəyişən xəritəsi və panel
# ==========================================================================================
# Kataloqun bəzi adları iş kitabında ayrıca sıra deyil: ya eyni sıranın ikinci adıdır (ALIAS),
# ya da bir sətirlik AÇIQ törəmədir (DERIVED). Hər ikisi `data_layer` §13-də sənədlənib və
# burada hər həll addımından sonra yenidən hesablanır ki, endogen dəyişənlərlə uyğun qalsın.

def _dl():
    from importlib import import_module
    return import_module("data_layer")


def _apply_alias_and_derived(panel: Dict[str, Dict[int, float]], years: Iterable[int],
                             skip: Iterable[str] = ()) -> None:
    """ALIAS və DERIVED sıralarını verilmiş illər üçün yenidən hesablayır.

    `skip` — kataloqun ÖZ tənliyi ilə həll olunan adlar (məsələn `TIKINTI`, `DEF_REST`,
    `DEF_FINANCE`): onlar törəmə/alias qaydası ilə ƏVƏZLƏNMİR, əks halda həll nəticəsi
    səssizcə üzərinə yazılardı.
    """
    dl = _dl()
    skip = {str(x).upper() for x in skip}
    for dst, (src, _note) in dl.MOE_SPEC_ALIAS.items():
        if dst.upper() in skip:
            continue
        if src in panel:
            for y in years:
                if y in panel[src]:
                    panel.setdefault(dst, {})[y] = panel[src][y]

    def _div(a, b, y):
        va, vb = panel.get(a, {}).get(y), panel.get(b, {}).get(y)
        if va is None or vb in (None, 0):
            return None
        return va / vb

    for y in years:
        for dst, (a, b) in (("RW", ("W", "CPI")), ("DEF_FINANCE", ("VA_FINANCE", "RVA_FINANCE")),
                            ("DEF_REST", ("VA_REST", "RVA_REST")), ("RIO", ("IO", "DEFID")),
                            ("RINO", ("INO", "DEFID"))):
            if dst.upper() in skip:
                continue
            v = _div(a, b, y)
            if v is not None:
                panel.setdefault(dst, {})[y] = v
        # REALFINALC = nominal son istehlak ÷ ev təsərrüfatlarının deflyatoru (HC/RHC)
        defl = _div("HC", "RHC", y)
        fc = panel.get("FINALC_NOM", {}).get(y)
        if "REALFINALC" not in skip and defl not in (None, 0) and fc is not None:
            panel.setdefault("REALFINALC", {})[y] = fc / defl


def history_panel() -> Dict[str, Dict[int, float]]:
    """Nazirliyin iş kitablarından gələn TARİX paneli (≤2024) + alias/törəmələr."""
    dl = _dl()
    df = dl.moe_spec_panel()
    panel: Dict[str, Dict[int, float]] = {}
    for var, g in df.groupby("var"):
        panel[str(var)] = {int(y): float(v) for y, v in zip(g["year"], g["value"])}
    yrs = sorted({y for s in panel.values() for y in s})
    _apply_alias_and_derived(panel, yrs)
    # `ECM_MS_TP1` ≡ `ECM_MS_TP` (D3, bayraqlı) — səviyyə hesablandıqdan sonra kopyalanır.
    return panel


# --- ekzogen sürücülərin bu paketin proqnozları ilə uzadılması ---------------------------
# Hər sətir: Nazirliyin dəyişəni → (mənbə növü, açar, tətbiq qaydası, AZ qeyd).
#   mənbə növü: "a" = `data/assumptions.csv` açarı; "f" = `outputs/forecast_long.csv` (source=ours)
#   qayda: "g"       — mənbə faiz artımıdır, səviyyə həmin templə uzadılır;
#          "lvl"     — mənbə eyni tərifli səviyyədir, onun artım tempi tətbiq olunur;
#          "direct"  — mənbənin dəyəri birbaşa səviyyə kimi götürülür;
#          "hold"    — son faktiki səviyyə saxlanılır.
EXOG_DRIVER_MAP: Dict[str, Tuple[str, str, str, str]] = {
    "CPI":       ("a", "cpi_infl", "g", "İQİ indeksi FR9 inflyasiya yolu ilə uzadılır"),
    "INFLATION": ("a", "cpi_infl", "direct", "inflyasiya faizi birbaşa FR9 yoludur"),
    "ER":        ("a", "usd_azn", "direct", "AZN/USD məzənnəsi FR10 rejim qaydasından"),
    "NER":       ("a", "usd_azn", "lvl", "nominal məzənnə indeksi məzənnənin artım tempi ilə"),
    "OPIWTI":    ("a", "brent_usd", "lvl", "WTI qiyməti FR5 Brent yolunun ARTIM TEMPİ ilə "
                  "(səviyyə fərqi saxlanılır: Brent ≠ WTI)"),
    "OPWTI":     ("a", "brent_usd", "lvl", "neft qiyməti FR5 Brent yolunun artım tempi ilə"),
    "W":         ("f", "wage_avg", "lvl", "orta aylıq əmək haqqı FR8 yolundan"),
    "MW":        ("a", "cpi_infl", "g", "minimum əmək haqqı inflyasiya templə (fərziyyə)"),
    "L":         ("f", "employment_g", "g", "məşğulluq FR8 artım tempindən"),
    "POP":       ("a", "pop_avg", "lvl", "əhalinin sayı FR8 demoqrafiya yolundan"),
    "RGDPMTP":   ("a", "partner_gdp_realg", "g", "tərəfdaş ölkələrin real ÜDM indeksi"),
    "CPIMTP":    ("a", "import_price_infl", "g", "tərəfdaş İQİ idxal qiymət yolu ilə"),
    "DEFID":     ("f", "inv_defl_g", "g", "investisiya deflyatoru FR6 yolundan"),
    "RID":       ("f", "inv_total_g", "g", "daxili investisiya, real — FR6 artım tempi"),
    "RIF":       ("f", "inv_foreign_g", "g", "xarici investisiya, real — FR6 artım tempi"),
    "RIDNS":     ("f", "inv_nonoil_nonstate_g", "g", "qeyri-dövlət investisiyası — FR6"),
    "RGDPO":     ("a", "oil_realg", "g", "neft-qaz ÜDM-i FR05b real artımı ilə"),
    "AGRI":      ("f", "g_agri", "gd:d_agri", "kənd təsərrüfatı əlavə dəyəri (real × deflyator)"),
    "VA_CONST":  ("f", "g_constr", "gd:d_constr", "tikinti əlavə dəyəri (real × deflyator)"),
    "TRADE":     ("f", "g_trade", "gd:d_trade", "ticarət əlavə dəyəri (real × deflyator)"),
    "VA_INFORM": ("f", "g_ict", "gd:d_ict", "rabitə əlavə dəyəri (real × deflyator)"),
    "VA_MANU":   ("f", "g_manuf", "gd:d_manuf", "emal sənayesi əlavə dəyəri (real × deflyator)"),
    "RVA_MANU":  ("f", "g_manuf", "g", "emal sənayesi, real"),
    "RVA_MINE":  ("f", "g_mining", "g", "mədənçıxarma, real"),
    "RVA_TRANS": ("f", "g_transport", "g", "nəqliyyat, real"),
    "RVA_INFORM": ("f", "g_ict", "g", "rabitə, real"),
    "FOREIGNTURIST": ("f", "g_tourism", "g", "turist sayı turizm sahəsinin real artımı ilə"),
    "HI":        ("f", "income_realg", "gd:cpi_infl", "əhalinin gəlirləri (real × inflyasiya)"),
    "POPINCOME": ("f", "income_realg", "g", "əhalinin gəlirləri, sabit qiymətlərlə"),
    "RIEA":      ("f", "income_realg", "g", "sahibkarlıq gəlirləri, sabit qiymətlərlə"),
    "IEA":       ("f", "income_realg", "gd:cpi_infl", "sahibkarlıq gəlirləri, nominal"),
}
# Xüsusi qayda ilə uzadılan, lakin cədvələ sığmayan sıralar (aşağıda `_extend_special`).
SPECIAL_EXOG = ("LCPI_HP", "NEER", "REER", "NERMTP")


class Drivers:
    """Sürücü dəyərlərinin VAHİD oxuyucusu: əvvəlcə nəşr olunmuş proqnoz sıraları
    (`forecast_long.csv`, `source=ours` — faktiki və proqnoz illəri birlikdə), sonra
    `assumptions.csv`.

    2025-ci il xüsusi haldır: fərziyyə faylı yalnız 2026–2030-u əhatə edir, çünki 2025 bizim
    müqavilədə FAKTİKİ ildir. Ona görə həmin il üçün dəyər nəşr olunmuş faktiki sıradan
    götürülür; yalnız fərziyyə kimi mövcud olan açarlar üçün isə (`partner_gdp_realg`,
    `import_price_infl`) yolun İLK ili istifadə olunur və bu, jurnalda yazılır.
    """

    def __init__(self, assumptions: pd.DataFrame, ours: pd.DataFrame):
        self.a = {(str(k), int(y)): float(v) for k, y, v in
                  zip(assumptions["assumption_key"], assumptions["year"], assumptions["value"])}
        self.f = {(str(k), int(y)): float(v) for k, y, v in
                  zip(ours["series_code"], ours["year"], ours["value"])}
        self.a_first = {}
        for (k, y) in self.a:
            if k not in self.a_first or y < self.a_first[k]:
                self.a_first[k] = y
        self.fallback = {"usd_azn": "fx_usd_azn_eop"}

    def value(self, key: str, year: int) -> Optional[float]:
        for k in (key, self.fallback.get(key)):
            if not k:
                continue
            if (k, year) in self.f:
                return self.f[(k, year)]
            if (k, year) in self.a:
                return self.a[(k, year)]
        first = self.a_first.get(key)
        if first is not None and year < first:
            return self.a[(key, first)]
        return None


def _rule_growth(rule: str, key: str, drivers: "Drivers", year: int) -> Optional[float]:
    """Bir ildə tətbiq olunacaq faiz artımı (`gd:` ikinci açarla birləşdirir: real × deflyator)."""
    g1 = drivers.value(key, year)
    if g1 is None:
        return None
    if rule.startswith("gd:"):
        g2 = drivers.value(rule.split(":", 1)[1], year)
        if g2 is None:
            return None
        return ((1 + g1 / 100.0) * (1 + g2 / 100.0) - 1) * 100.0
    return g1


def _default_growth(unit: str, year: int, drivers: "Drivers") -> Optional[float]:
    """Xəritədə olmayan (və ya həmin ildə mənbəsi olmayan) ekzogen sıralar üçün susma qaydası.

    Qayda AÇIQDIR və hesabatda çap olunur: faiz/nisbət sıraları SAXLANILIR; indekslər
    inflyasiya ilə; sabit qiymətli sıralar qeyri-neft real artımı ilə; pul ifadəsindəki
    sıralar qeyri-neft nominal artımı ilə; natural göstəricilər qeyri-neft real artımı ilə.
    """
    u = (unit or "").lower()
    infl = drivers.value("cpi_infl", year)
    realg = drivers.value("nonoil_realg", year)
    defl = drivers.value("nonoil_defl", year)
    if "%" in u or "nisbət" in u or "loq" in u:
        return 0.0
    if "indeks" in u:
        return infl
    if "sabit" in u:
        return realg
    if "azn" in u or "usd" in u:
        if infl is None or realg is None:
            return None
        base = defl if defl is not None else infl
        return ((1 + realg / 100.0) * (1 + base / 100.0) - 1) * 100.0
    return realg


def extend_exogenous(panel: Dict[str, Dict[int, float]], exog: Iterable[str],
                     years: Sequence[int], drivers: "Drivers",
                     units: Mapping[str, str]) -> pd.DataFrame:
    """Ekzogen sıraları 2025–2030 üçün BU PAKETİN yolları ilə uzadır; qeyd cədvəli qaytarır."""
    log = []
    for var in sorted(exog):
        s = panel.setdefault(var, {})
        hist = [y for y in s if y <= 2024]
        if not hist:
            log.append({"dəyişən": var, "qayda": "—", "mənbə": "—",
                        "qeyd": "tarixi dəyər yoxdur, uzadılmadı"})
            continue
        prev = s[max(hist)]
        unit = units.get(var, "")
        kind, key, rule, note = EXOG_DRIVER_MAP.get(var, ("", "", "susma", ""))
        used_default, stopped = 0, None
        for y in years:
            if rule == "direct":
                v = drivers.value(key, y)
                if v is not None:
                    s[y] = float(v)
                    continue
                g = _default_growth(unit, y, drivers)
                used_default += 1
            elif rule == "lvl":
                cur, pre = drivers.value(key, y), drivers.value(key, y - 1)
                if cur is not None and pre not in (None, 0):
                    prev = prev * cur / pre
                    s[y] = prev
                    continue
                g = _default_growth(unit, y, drivers)
                used_default += 1
            elif rule == "susma":
                g = _default_growth(unit, y, drivers)
            else:
                g = _rule_growth(rule, key, drivers, y)
                if g is None:
                    g = _default_growth(unit, y, drivers)
                    used_default += 1
            if g is None:
                stopped = y
                break
            prev = prev * (1 + g / 100.0)
            s[y] = prev
        src = "susma qaydası" if rule == "susma" else f"{'fərziyyə' if kind == 'a' else 'proqnoz'}:{key}"
        q = note or ("susma qaydası: faiz → saxlanılır; indeks → inflyasiya; sabit → qeyri-neft "
                     "real artımı; pul ifadəsi → qeyri-neft nominal artımı; natural → qeyri-neft "
                     f"real artımı (vahid: {unit or '—'})")
        if used_default:
            q += f"; {used_default} ildə susma qaydası tətbiq olunub"
        if stopped:
            q += f"; {stopped}-ci ildən UZADILMADI (mənbə yoxdur)"
        log.append({"dəyişən": var, "qayda": rule, "mənbə": src, "qeyd": q})
    return pd.DataFrame(log)


def _extend_special(panel: Dict[str, Dict[int, float]], years: Sequence[int],
                    drivers: "Drivers") -> List[dict]:
    """NEER/REER/NERMTP/LCPI_HP — açıq, sənədləşdirilmiş qaydalarla."""
    def _a(k, y):
        return drivers.value(k, y)
    log = []
    for var, note in (("NEER", "de-fakto bağlılıq: nominal effektiv məzənnə SAXLANILIR"),
                      ("NERMTP", "tərəfdaş məzənnə indeksi SAXLANILIR — tərif təsdiqi gözlənilir")):
        s = panel.get(var, {})
        h = [y for y in s if y <= 2024]
        if h:
            for y in years:
                s[y] = s[max(h)]
            log.append({"dəyişən": var, "qayda": "saxlanılır", "mənbə": "—", "qeyd": note})
    s = panel.get("REER", {})
    h = [y for y in s if y <= 2024]
    if h:
        prev = s[max(h)]
        for y in years:
            di = _a("cpi_infl", y)
            dp = _a("import_price_infl", y)
            if di is None or dp is None:
                break
            prev = prev * (1 + (di - dp) / 100.0)
            s[y] = prev
        log.append({"dəyişən": "REER", "qayda": "inflyasiya fərqi", "mənbə": "fərziyyə:cpi_infl−import_price_infl",
                    "qeyd": "real effektiv məzənnə inflyasiya fərqi ilə sürüşür"})
    s = panel.get("LCPI_HP", {})
    h = [y for y in s if y <= 2024]
    if h:
        prev = s[max(h)]
        for y in years:
            di = _a("cpi_infl", y)
            if di is None:
                break
            prev = prev + math.log1p(di / 100.0)
            s[y] = prev
        log.append({"dəyişən": "LCPI_HP", "qayda": "loq-artım", "mənbə": "fərziyyə:cpi_infl",
                    "qeyd": "loq-İQİ trendi inflyasiya yolu ilə uzadılır"})
    return log


# ==========================================================================================
# 4. Blok-rekursiv Gauss–Seidel həlledicisi (port: `engine_sectoral.py`)
# ==========================================================================================
EXPLOSIVE_DLOG = 0.30       # |Δln| > 0.30 (≈ 35 % illik) → yol partlayır, dəyər qəbul edilmir
GS_PASSES = 3               # eyni il daxilində eyni-zamanlı əlaqələr üçün keçid sayı


def runnable_rows(rows: List[SpecRow]) -> List[SpecRow]:
    """Statusu «işə salınır» olan (A, B və bayraqlı ECM) davranış tənlikləri."""
    return [r for r in rows if r.parses and not r.is_ecm
            and r.status in (STATUS_A, STATUS_B, STATUS_C_FLAG)]


def ecm_rows(rows: List[SpecRow]) -> List[SpecRow]:
    """Kataloqda TƏYİN OLUNMUŞ xəta-korreksiyası səviyyələri (15 sətir)."""
    return [r for r in rows if r.parses and r.is_ecm]


def refresh_ecm(panel: Dict[str, Dict[int, float]], rows: List[SpecRow], year: int) -> None:
    """Bir il üçün bütün təyin olunmuş ECM səviyyələrini yenidən hesablayır."""
    for e in ecm_rows(rows):
        try:
            panel.setdefault(e.target, {})[year] = e.eq.eval_rhs(year, panel)
        except Exception:
            pass
    # D3: təyin olunmayan `ECM_MS_TP1` təyin olunmuş `ECM_MS_TP`-dən götürülür (BAYRAQLI).
    for dst, src in ECM_ASSUMED_EQUIVALENT.items():
        v = panel.get(src, {}).get(year)
        if v is not None:
            panel.setdefault(dst, {})[year] = v


def _lookup(panel: Mapping[str, Mapping[int, float]], name: str):
    s = panel.get(name)
    if s is not None:
        return s
    up = name.upper()
    for k in panel:
        if k.upper() == up:
            return panel[k]
    return None


def available(eq: EViewsEq, panel: Mapping[str, Mapping[int, float]], year: int) -> bool:
    """Sağ tərəfin bütün dəyişənləri `year` və `year−1` üçün mövcuddurmu?"""
    for v in eq.rhs_vars | ({eq.target} if eq.transform != "LOG" else set()):
        if v in ("T", "t", "c"):
            continue
        s = _lookup(panel, v)
        if not s or (year - 1) not in s:
            return False
    return True


def solve(panel: Dict[str, Dict[int, float]], rows: List[SpecRow],
          years: Sequence[int], exog: Iterable[str] = ()) -> Tuple[set, pd.DataFrame]:
    """Blok-rekursiv Gauss–Seidel: hər il üçün `GS_PASSES` keçid, sonra ECM səviyyələri.

    Ekzogen (`exog`) dəyişənlər YENİDƏN proqnozlaşdırılmır — onlar bu paketin öz yollarından
    gəlir. Partlayan yol mühafizəsi: |Δln| > 0.30 olan həll qəbul edilmir və sətir jurnalda
    «partlayan» kimi qeyd olunur.
    """
    exog = set(exog)
    endo = {r.target for r in runnable_rows(rows) if r.target not in exog}
    solved, log = set(), []
    for y in years:
        for _ in range(GS_PASSES):
            for r in runnable_rows(rows):
                if r.target in exog:
                    continue
                if not available(r.eq, panel, y):
                    continue
                try:
                    val = r.eq.level(y, panel)
                except Exception:
                    continue
                if val is None or not math.isfinite(val):
                    continue
                prev = panel.get(r.target, {}).get(y - 1)
                if prev and prev > 0 and val > 0 and abs(math.log(val / prev)) > EXPLOSIVE_DLOG:
                    log.append({"il": y, "№": r.no, "hədəf": r.target, "hal": "partlayan yol — rədd edildi"})
                    continue
                panel.setdefault(r.target, {})[y] = val
                solved.add(r.target)
        refresh_ecm(panel, rows, y - 1)
        refresh_ecm(panel, rows, y)
        _apply_alias_and_derived(panel, [y], skip=endo)
    return solved, pd.DataFrame(log)


# ==========================================================================================
# 5. Geriyə doğru sınaq — Nazirliyin SABİT əmsalları, bizim genişlənən pəncərəmiz
# ==========================================================================================
# Metodoloji qeyd (hesabatda da yazılır): Nazirliyin əmsalları TAM nümunə üzərində
# qiymətləndirilib, ona görə onların bir addımlıq proqnozu **nümunədaxili əmsallarla**
# hesablanır — bu, müqayisəni ONLARIN XEYRİNƏ meyilləndirir. Bizim tənliklərimiz isə hər
# vintajda yenidən qiymətləndirilir (gələcəyə baxış yoxdur). Fərq açıq göstərilir; onların
# spesifikasiyası bu güzəştə baxmayaraq uduzursa, nəticə daha da möhkəmdir.

# Nazirliyin hədəfi ↔ bu paketin sıra kodu. `çevirmə`: hədəfin səviyyəsindən bizim
# göstəriciyə keçid. Bütün cütlər səviyyə üzrə YOXLANILIB (1995–2024 faktiki dəyərlər üst-üstə
# düşür — eyni rəsmi statistika), buna görə müqayisə tərif fərqi daşımır.
TARGET_MAP: Dict[str, dict] = {
    "CPI":        dict(code="cpi_infl", conv="pct_change", unit="%",
                       note="İQİ indeksindən orta illik inflyasiyaya"),
    "MGNO":       dict(code="imp_goods_nonoil", conv="level", unit="mln USD",
                       note="qeyri neft-qaz malların idxalı — səviyyə eynidir"),
    "RXGNO":      dict(code="exp_goods_nonoil", conv="level", unit="mln USD",
                       note="qeyri neft-qaz malların ixracı — səviyyə eynidir"),
    "XS_TP":      dict(code="exp_serv_transport", conv="level", unit="mln USD",
                       note="nəqliyyat xidmətlərinin ixracı"),
    "MS_TP":      dict(code="imp_serv_transport", conv="level", unit="mln USD",
                       note="nəqliyyat xidmətlərinin idxalı"),
    "XS_IFORM":   dict(code="exp_serv_telecom", conv="level", unit="mln USD",
                       note="rabitə xidmətlərinin ixracı"),
    "MS_TRAVP":   dict(code="imp_serv_travel", conv="level_plus", plus="MS_TRAVB", unit="mln USD",
                       note="turizm xidmətlərinin idxalı = şəxsi + işgüzar səfərlər"),
    "RVA_CONST":  dict(code="g_constr", conv="pct_change", unit="%",
                       note="tikinti əlavə dəyərinin real artım tempi"),
    "RVA_TRADE":  dict(code="g_trade", conv="pct_change", unit="%",
                       note="ticarət əlavə dəyərinin real artım tempi"),
    "RVA_ACCOM":  dict(code="g_tourism", conv="pct_change", unit="%",
                       note="turizm və ictimai iaşənin real artım tempi"),
    "DEF_CONSTR": dict(code="d_constr", conv="pct_change", unit="%",
                       note="tikinti deflyatorunun artımı"),
    "DEF_ACCOM":  dict(code="d_tourism", conv="pct_change", unit="%",
                       note="turizm deflyatorunun artımı"),
    "DEF_TRANSP": dict(code="d_transport", conv="pct_change", unit="%",
                       note="nəqliyyat deflyatorunun artımı"),
    "FPI_AZ":     dict(code="d_agri", conv="pct_change", unit="%",
                       note="kənd təsərrüfatı deflyatorunun artımı (kataloqda `FPI_AZ`)"),
    "DEFID":      dict(code="inv_defl_g", conv="pct_change", unit="%",
                       note="investisiya deflyatorunun artımı"),
    "RHC":        dict(code="final_consumption", conv="pct_change", unit="%",
                       note="ev təsərrüfatlarının son istehlakının real artım tempi "
                            "(bizim sınağımız bu kod altında məhz real artımı saxlayır)"),
    "NONOILWAGE": dict(code="wage_nonoil", conv="level", unit="AZN",
                       note="qeyri neft-qaz sektorunda orta aylıq əmək haqqı"),
}


def _convert(target: str, level: float, prev_level: Optional[float], spec: dict,
             panel: Mapping[str, Mapping[int, float]], year: int) -> Optional[float]:
    """Nazirliyin hədəf səviyyəsini bizim göstəricinin vahidinə çevirir."""
    conv = spec["conv"]
    if conv == "level":
        return float(level)
    if conv == "level_plus":
        other = panel.get(spec["plus"], {}).get(year)
        return None if other is None else float(level) + float(other)
    if conv == "pct_change":
        if prev_level in (None, 0):
            return None
        return (float(level) / float(prev_level) - 1.0) * 100.0
    return None


def spec_predictions(panel: Mapping[str, Mapping[int, float]], rows: List[SpecRow],
                     years: Sequence[int]) -> Dict[Tuple[int, str], Dict[int, float]]:
    """Hər işlək tənlik üçün bir addımlıq (şərti) proqnozlar: {(№, hədəf): {il: dəyər}}.

    «Şərti» — sağ tərəfin FAKTİKİ dəyərləri həmin ildə məlum sayılır; bu, bizim öz
    `dynamics_forms` sınağımızın konvensiyasının eynisidir (sürücü faktiki, əmsal keçmişdən).
    """
    out: Dict[Tuple[int, str], Dict[int, float]] = {}
    for r in runnable_rows(rows):
        if r.target not in TARGET_MAP:
            continue
        spec = TARGET_MAP[r.target]
        preds: Dict[int, float] = {}
        for y in years:
            if not available(r.eq, panel, y):
                continue
            try:
                lvl = r.eq.level(y, panel)
            except Exception:
                continue
            if lvl is None or not math.isfinite(lvl):
                continue
            v = _convert(r.target, lvl, panel.get(r.target, {}).get(y - 1), spec, panel, y)
            if v is not None and math.isfinite(v):
                preds[y] = float(v)
        if preds:
            out[(r.no, r.target)] = preds
    return out


def implied_actual(target: str, spec: dict, panel: Mapping[str, Mapping[int, float]],
                   year: int) -> Optional[float]:
    """Nazirliyin FAKTİKİ sırasından bizim göstəricinin vahidində «faktiki» dəyər.

    Müqayisənin ÖN ŞƏRTİ budur: eyni ildə eyni anlayış. Bu funksiya onların sırasını bizim
    çevirməmizlə hesablayır və `side_by_side` onu bizim sınaq faylındakı `actual` sütunu ilə
    tutuşdurur; uyğunsuzluq böyükdürsə cüt müqayisədən ÇIXARILIR (tərif fərqi).
    """
    s = panel.get(target, {})
    if year not in s:
        return None
    return _convert(target, s[year], s.get(year - 1), spec, panel, year)


#: Bu ssenarinin öz sətirlərinin model adı — `validation_backtest.csv`-də ikinci sütun.
SPEC_MODEL = "NAZIRLIK_SPES"
#: Müqayisədə BİZİM model kimi sayılmayan adlar: σ nisbəti göstəricisi və bu ssenarinin özü
#: (əks halda əvvəlki icradan qalmış `NAZIRLIK_SPES` sətirləri "bizim ən yaxşı formamız" kimi
#: seçilər və cədvəl özünü özü ilə müqayisə edərdi).
NOT_OUR_MODELS = (SPEC_MODEL, "DINAMIKLIK")


def _our_backtest_table(path: Optional[str] = None) -> pd.DataFrame:
    p = path or os.path.join(OUT, "validation_backtest.csv")
    bt = pd.read_csv(p, float_precision="round_trip")
    return bt[~bt["model"].isin(NOT_OUR_MODELS)].copy()


def side_by_side(panel: Mapping[str, Mapping[int, float]], rows: List[SpecRow],
                 backtest_path: Optional[str] = None,
                 min_points: int = 4) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Onların spesifikasiyası ↔ bizim tənliklərimiz ↔ RW, EYNİ pəncərədə.

    Pəncərə = bizim sınağımızın həmin sıra üzrə illəri ∩ onların tənliyinin proqnoz verə
    bildiyi illər. Hər üç RMSE məhz bu ortaq pəncərədə yenidən hesablanır — ona görə
    diskdəki `rw_rmse_h` ilə fərqlənə bilər (fərqli pəncərə), bu, hesabatda yazılır.

    Qaytarır: (bacarıq cədvəli, `validation_backtest.csv` üçün sətirlər).
    """
    bt = _our_backtest_table(backtest_path)
    years_all = list(range(1996, LAST_ACTUAL + 1))
    preds = spec_predictions(panel, rows, years_all)
    skill_rows, out_rows = [], []
    by_code: Dict[str, List[Tuple[int, str, Dict[int, float]]]] = {}
    for (no, tgt), p in preds.items():
        by_code.setdefault(TARGET_MAP[tgt]["code"], []).append((no, tgt, p))

    for code, cands in sorted(by_code.items()):
        ours = bt[bt["series_code"] == code]
        if ours.empty:
            for no, tgt, p in cands:
                skill_rows.append({"sıra kodu": code, "Nazirlik №": no, "hədəf": tgt,
                                   "pəncərə": "—", "n": 0, "qeyd": "bizim sınağımızda bu sıra yoxdur"})
            continue
        actual = ours.drop_duplicates("vintage_year").set_index(
            ours.drop_duplicates("vintage_year")["vintage_year"] + 1)["actual"].to_dict()
        our_years = set(actual)
        rw = ours[ours["model"] == "RW"].copy()
        rw_pred = dict(zip(rw["vintage_year"] + 1, rw["forecast"]))
        # tərif yoxlaması: onların FAKTİKİ sırası bizim `actual` sütunu ilə üst-üstə düşməlidir
        tgt0 = cands[0][1]
        spec0 = TARGET_MAP[tgt0]
        chk_years = [y for y in sorted(our_years) if implied_actual(tgt0, spec0, panel, y) is not None]
        mismatch = None
        if len(chk_years) >= 4:
            a = np.array([actual[y] for y in chk_years], dtype=float)
            b = np.array([implied_actual(tgt0, spec0, panel, y) for y in chk_years], dtype=float)
            sd = float(np.std(a)) or 1.0
            mismatch = float(np.mean(np.abs(a - b)) / sd)
        if mismatch is not None and mismatch > 1.0:
            for no, tgt, p in cands:
                skill_rows.append({"sıra kodu": code, "Nazirlik №": no, "hədəf": tgt,
                                   "pəncərə": "—", "n": 0,
                                   "qeyd": f"tərif uyğunsuzluğu (fərq/std = {mismatch:.2f}) — "
                                           "müqayisə aparılmır"})
            continue
        best = None
        for no, tgt, p in cands:
            common = sorted(our_years & set(p) & set(rw_pred))
            if len(common) < min_points:
                skill_rows.append({"sıra kodu": code, "Nazirlik №": no, "hədəf": tgt,
                                   "pəncərə": (f"{min(common)}–{max(common)}" if common else "—"),
                                   "n": len(common),
                                   "qeyd": f"ortaq pəncərə {min_points} ildən qısadır"})
                continue
            e_th = np.array([p[y] - actual[y] for y in common], dtype=float)
            e_rw = np.array([rw_pred[y] - actual[y] for y in common], dtype=float)
            rmse_th = float(np.sqrt(np.mean(e_th ** 2)))
            rmse_rw = float(np.sqrt(np.mean(e_rw ** 2)))
            # bizim ƏN YAXŞI formamız eyni pəncərədə (müqayisə qəsdən bizim əleyhimizə qurulur)
            our_best, our_name = None, None
            for m, g in ours[ours["model"] != "RW"].groupby("model"):
                pr = dict(zip(g["vintage_year"] + 1, g["forecast"]))
                if not set(common) <= set(pr):
                    continue
                e = np.array([pr[y] - actual[y] for y in common], dtype=float)
                r = float(np.sqrt(np.mean(e ** 2)))
                if our_best is None or r < our_best:
                    our_best, our_name = r, str(m)
            rec = {"sıra kodu": code, "Nazirlik №": no, "hədəf": tgt,
                   "pəncərə": f"{min(common)}–{max(common)}", "n": len(common),
                   "RMSE Nazirlik": rmse_th, "RMSE bizim": our_best, "RMSE RW": rmse_rw,
                   "bizim model": our_name,
                   "bacarıq Nazirlik %": (None if not rmse_rw else round((1 - rmse_th / rmse_rw) * 100, 1)),
                   "bacarıq bizim %": (None if (not rmse_rw or our_best is None)
                                       else round((1 - our_best / rmse_rw) * 100, 1)),
                   "qeyd": ("" if mismatch is None else f"tərif uyğunluğu: fərq/std = {mismatch:.2f}")}
            skill_rows.append(rec)
            if best is None or rmse_th < best[0]:
                best = (rmse_th, no, tgt, p, common, rmse_rw, actual)
        if best is None:
            continue
        rmse_th, no, tgt, p, common, rmse_rw, actual = best
        cov = float(np.mean([abs(p[y] - actual[y]) <= 1.2816 * rmse_th for y in common]))
        for y in common:
            err = p[y] - actual[y]
            out_rows.append({
                "series_code": code, "model": SPEC_MODEL, "vintage_year": int(y - 1),
                "horizon_h": 1, "actual": float(actual[y]), "forecast": float(p[y]),
                "error": float(err), "abs_error": abs(float(err)),
                "rmse_h": rmse_th, "rw_rmse_h": rmse_rw, "coverage80": cov})
    skill = pd.DataFrame(skill_rows)
    if not skill.empty:
        skill = skill.sort_values(["sıra kodu", "Nazirlik №"], kind="stable").reset_index(drop=True)
    return skill, pd.DataFrame(out_rows, columns=[
        "series_code", "model", "vintage_year", "horizon_h", "actual", "forecast",
        "error", "abs_error", "rmse_h", "rw_rmse_h", "coverage80"])


# ==========================================================================================
# 6. Tam ssenari: tarix + ekzogen uzatma + həll + proqnoz sətirləri
# ==========================================================================================
@dataclass
class ScenarioResult:
    panel: Dict[str, Dict[int, float]]
    rows: List[SpecRow]
    exog: set
    endo: set
    solved: set
    exog_log: pd.DataFrame
    solve_log: pd.DataFrame
    status: pd.DataFrame


def run_scenario(assumptions_path: str = ASSUMPTIONS,
                 forecast_path: Optional[str] = None,
                 years: Optional[Sequence[int]] = None) -> ScenarioResult:
    """Ssenarini uçdan-uca işə salır: tarix paneli → ekzogen uzatma → Gauss–Seidel həll."""
    dl = _dl()
    rows = load_catalog()
    panel = history_panel()
    prov = dl.moe_spec_provenance()
    units = dict(zip(prov["var"], prov["unit"]))
    for dst, (src, _n) in dl.MOE_SPEC_ALIAS.items():
        units.setdefault(dst, units.get(src, ""))
    for dst in dl.MOE_SPEC_DERIVED:
        units.setdefault(dst, "")

    assumptions = pd.read_csv(assumptions_path)
    fpath = forecast_path or os.path.join(OUT, "forecast_long.csv")
    fl = pd.read_csv(fpath, float_precision="round_trip")
    ours = fl[fl["source"] == "ours"][["series_code", "year", "value"]].drop_duplicates(
        ["series_code", "year"], keep="last")

    run = runnable_rows(rows)
    endo = {r.target for r in run}
    exog = {v for v in panel if v not in endo}
    # ECM səviyyələri panelin bir hissəsidir, ekzogen kimi uzadılmır
    exog = {v for v in exog if not v.upper().startswith("ECM")}
    yrs = list(years or range(LAST_ACTUAL, LAST_FORECAST + 1))

    drivers = Drivers(assumptions, ours)
    log = _extend_special(panel, yrs, drivers)
    exog_tab = extend_exogenous(panel, exog - set(SPECIAL_EXOG), yrs, drivers, units)
    exog_log = pd.concat([exog_tab, pd.DataFrame(log)], ignore_index=True).sort_values(
        "dəyişən", kind="stable").reset_index(drop=True)
    _apply_alias_and_derived(panel, yrs)

    for y in range(1996, 2025):
        refresh_ecm(panel, rows, y)

    # PBCRHC körpüsü: iş kitabında bu sıra 2020-dən 2030-a qədər −0.0047-də
    # DONDURULUB, yəni həmin illər tarix deyil. Ona görə 2020–2024 aralığı kataloqun ÖZ AR(6)
    # tənliyi (27-ci sətir) ilə doldurulur və bu, BAYRAQLANIR: təsnifat həmin tənliyin sıranı
    # tam təkrar etmədiyini qeyd edir (qalıq s.k. 0.0085, sıranın s.k. 0.0097).
    bridge = [r for r in runnable_rows(rows) if r.target == "PBCRHC"]
    bridge_years = [y for y in range(2020, LAST_ACTUAL) if y not in panel.get("PBCRHC", {})]
    if bridge and bridge_years:
        solve(panel, rows[:0] + bridge, bridge_years, exog=set())
        for y in bridge_years:
            refresh_ecm(panel, rows, y)

    solved, solve_log = solve(panel, rows, yrs, exog=exog)

    return ScenarioResult(panel=panel, rows=rows, exog=exog, endo=endo, solved=solved,
                          exog_log=exog_log, solve_log=solve_log, status=catalog_frame(rows))


def forecast_rows(res: ScenarioResult, fr: str = "FR13") -> pd.DataFrame:
    """Ssenarinin 2026–2030 yollarını `forecast_long.csv` müqaviləsinə uyğun sətirlərə çevirir."""
    dl = _dl()
    dct = dl.series_dictionary() if hasattr(dl, "series_dictionary") else None
    names = ({} if dct is None else dict(zip(dct["series_code"], dct["name_az"])))
    unitm = ({} if dct is None else dict(zip(dct["series_code"], dct["unit"])))
    out = []
    for tgt, spec in TARGET_MAP.items():
        if tgt not in res.solved:
            continue
        code = spec["code"]
        s = res.panel.get(tgt, {})
        for y in range(FIRST_FORECAST, LAST_FORECAST + 1):
            if y not in s:
                continue
            v = _convert(tgt, s[y], s.get(y - 1), spec, res.panel, y)
            if v is None or not math.isfinite(v):
                continue
            out.append({"fr": fr, "series_code": code,
                        "series_name_az": names.get(code, code), "unit": unitm.get(code, spec["unit"]),
                        "year": int(y), "kind": "forecast", "source": "ministry_spec",
                        "value": float(v), "lo80": None, "hi80": None, "lo50": None, "hi50": None})
    return pd.DataFrame(out, columns=["fr", "series_code", "series_name_az", "unit", "year",
                                      "kind", "source", "value", "lo80", "hi80", "lo50", "hi50"])


def bucket_summary(rows: Optional[List[SpecRow]] = None) -> pd.DataFrame:
    """Səbət sayımı — kodun özündən yenidən hesablanmış şəkildə."""
    rows = rows or load_catalog()
    tab = []
    for b in ("A", "B", "C", "D"):
        sel = [r for r in rows if r.bucket == b]
        tab.append({"səbət": b, "izah": BUCKET_NOTE[b], "sətir": len(sel),
                    "davranış tənliyi": len([r for r in sel if not r.is_ecm]),
                    "işə salınır": len([r for r in sel if r.status in
                                        (STATUS_A, STATUS_B, STATUS_C_FLAG)])})
    return pd.DataFrame(tab)
