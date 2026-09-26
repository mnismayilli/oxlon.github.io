# -*- coding: utf-8 -*-
"""Məlumat qatı — MİİS §15.5.1 (D6).

Bu modul paketin YEGANƏ məlumat girişidir. Üç mənbəni oxuyur:

  1. `Statistik data dinamika 05.06.2026 +.xlsx` (113 MB) — Nazirliyin operativ
     statistika bazası: etiketə əsaslanan başlıq aşkarlanması, axın/qalıq/nisbət
     təsnifatı, ilin əvvəlindən artan (YTD) sıraların de-kumulyasiyası;
  2. `8_vereq_original.xlsx` — Nazirliyin statutar model şablonu (Nazirlər
     Kabinetinin 24.05.2004 tarixli 75 nömrəli qərarı, bənd 2.4.1);
  3. beş CSV panel faylı (illik makro, model sıraları, xarici blok, aylıq panel,
     DSK-nın rəsmi 2025 nəticələri).

Memarlıq qaydası (D1): faktiki məlumat və Nazirliyin Proqnoz sütunları AYRI
obyektlərdə saxlanılır. `actuals()` heç bir halda 2025-dən sonrakı ili qaytarmır;
Proqnoz dəyərləri yalnız `ministry_baseline()` vasitəsilə, müqayisə məqsədi ilə
əlçatandır. Beləliklə heç bir reqressiya Proqnoz dəyərini görə bilmir.

Excel-ə hər müraciət etiketlə yoxlanılır: sətir nömrəsi göstərilsə belə, həmin
sətrin mətn etiketi gözlənilən etiketlə tutuşdurulur və uyğunsuzluq halında
istisna atılır. Çılpaq mövqe indeksi ilə oxu yoxdur.

Statutar şablonun buraxılışı barədə mühüm qeyd: onun Hesabat bloku 2024-cü ildə
bitir, yəni 2025 sütunu Nazirliyin proqnozudur. Sərhəd fərz edilmir, başlıq
mətnindən aşkarlanır və hər vərəqdə eyni olduğu yoxlanılır. Faktiki 2025 dəyərləri
iş kitabından və DSK-nın rəsmi nəticələrindən gəlir; şablonun 2025 proqnozu
`vintage_overrides.csv` faylında müqayisə üçün saxlanılır.

Xarici interfeys:
    build_cache(verbose=True, check_md5=True)
    actuals(series_code) -> pd.Series            # indeks = il, HƏMİŞƏ ≤ 2025
    actuals_source(series_code) -> pd.Series     # hər il üçün provenans
    ministry_baseline(series_code) -> pd.Series  # 2026–2030, yalnız müqayisə
    catalog() -> pd.DataFrame                    # series_code, name_az, unit, sheet, row, fr, ...
    series_dictionary() -> pd.DataFrame          # §3 müqaviləsinin sütunları
    vereq_actuals(sheet, row_label, ...) -> pd.Series
    vereq_baseline(sheet, row_label, ...) -> pd.Series
    products(sheet="2.4.1.4.") -> pd.DataFrame   # FR4, kompozit açar
    ref_report() / vintage_notes() / data_gaps() -> pd.DataFrame
    wb_registry() / wb_annual(code) / monthly(code)

Rəsmi açıq mənbələr (§14) və gömrük bülletenləri (§15) — bu buraxılış:
    public_panel() / public_provenance() / public_series(var)
    ssc_production_account() / ssc_branch_wage_ownership() / ssc_branch_employment_ownership()
    vereq_product_concordance() / fx_cross_check() / remittances_annual()
    customs_hs27() / customs_hs27_provenance() / customs_hs27_series(var) / customs_cross_check()
    verify_public_panel(...) / verify_customs_hs27(...)   # mənbədən təkrar oxu + tutuşdurma
Bu sıralar mövcud kodları ƏVƏZ ETMİR: hər biri AYRI kodla kataloqa əlavə olunub, ona görə
onların bağlanması heç bir hazır rəqəmi dəyişmir.

BVF etalon yolları (§16) — R23:
    reference_forecasts() / reference_series(var, block) / imf_reference(var)
    reference_provenance() / reference_match_report() / verify_reference_forecasts()
Bu sıralar YALNIZ MÜQAYİSƏ üçündür: `actuals()`/`ministry_baseline()` kimi ayrı obyektdə
saxlanılır, heç bir tənliyin nümunəsinə və heç bir fərziyyə açarına daxil olmur.
"""
import os
import re
import sys
import json
import math
import shutil
import hashlib
import unicodedata

import pandas as pd

try:
    import openpyxl
except ImportError:      # openpyxl yalnız keş qurularkən tələb olunur
    openpyxl = None

# Modul həm `import data_layer` (src qovluğu yolda), həm də `from src import
# data_layer` (paket kimi) şəklində idxal oluna bilir; hər iki halda config
# vahid mənbə olaraq qalır.
_SRC_DIR = os.path.dirname(os.path.abspath(__file__))
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

from config import (      # noqa: E402
    LAST_ACTUAL, FIRST_FORECAST, LAST_FORECAST,
    XLSX_MAIN, XLSX_MAIN_MD5, XLSX_VEREQ, VEREQ_SHEETS, MAIN_SHEETS,
    P_MACRO, P_MODEL, P_EXTERNAL, P_MONTHLY, P_OFFICIAL_2025,
    C_WB_REGISTRY, C_WB_ANNUAL, C_WB_MONTHLY, C_VEREQ_LONG, C_VEREQ_PRODUCTS,
    C_VEREQ_REF, C_VEREQ_META, C_SPLICE,
    C_ACTUALS, C_BASELINE, C_CATALOG, C_VINTAGE, C_DICT, C_GAPS, O_SERIES_DICT,
    P_MINISTRY_CATALOG, P_MOE_SPEC_PANEL, P_MOE_SPEC_PROV,
    P_PUBLIC_PANEL, P_PUBLIC_PROV, P_CUSTOMS_PANEL, P_CUSTOMS_PROV,
    P_REFERENCE_FORECASTS,
)

# ==========================================================================
# 1. Köməkçi funksiyalar
# ==========================================================================

MONTHS = {"yanvar": 1, "fevral": 2, "mart": 3, "aprel": 4, "may": 5, "iyun": 6,
          "iyul": 7, "avqust": 8, "sentyabr": 9, "oktyabr": 10, "noyabr": 11,
          "dekabr": 12}

ROMAN = {"i": 1, "ii": 2, "iii": 3, "iv": 4}

ABBR = {"Real sektor": "real", "Sosial sektor": "social", "Monetar sektoru": "money",
        "Neft-Qaz sektoru": "oilgas", "Ticarət": "trade", "Fiskal sektor": "fiscal",
        "Emal Sənayesi": "manuf", "Elektrik enerjisi": "elec",
        "Tədiyyə Balansı": "bop", "Mədənçıxarma": "mine", "Su təchizatı": "wsup",
        "DVX üzrə göstəricilər": "dvx", "DİP 2016-2026": "dip"}

NA_TOKENS = ("-", "", "—", "…", "...", "n/a", "N/A", "x", "X", "..")
REF_TOKEN = "#REF!"


def _norm(s) -> str:
    """Kiçik hərfə salır və İ hərfindəki birləşən nöqtəni atır (ə/ş/ç toxunulmur)."""
    return unicodedata.normalize("NFC", str(s)).lower().replace("̇", "")


def _squash(s) -> str:
    """Etiket müqayisəsi üçün: kənar boşluqlar atılır, daxili boşluqlar birləşdirilir."""
    return " ".join(str(s).replace("\xa0", " ").split())


def _key(s) -> str:
    """Etiketin müqayisə açarı — həm boşluq, həm registr normallaşdırılır."""
    return _norm(_squash(s))


def _slug(s) -> str:
    """Ölçü vahidini/mətni identifikator hissəsinə çevirir."""
    t = _norm(_squash(s))
    t = (t.replace("ə", "e").replace("ı", "i").replace("ö", "o").replace("ü", "u")
           .replace("ğ", "g").replace("ş", "s").replace("ç", "c"))
    t = re.sub(r"[^a-z0-9]+", "", t)
    return t or "na"


def num(v):
    """Etibarlı ədədə çevirmə: vergüllü onluq, '-' və boş xana NaN olur."""
    if v is None:
        return math.nan
    if isinstance(v, bool):
        return math.nan
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("\xa0", "")
    if s in NA_TOKENS or s == REF_TOKEN:
        return math.nan
    s = s.replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return math.nan


def _md5(path, chunk=1 << 20):
    h = hashlib.md5()
    with open(path, "rb") as fh:
        for blk in iter(lambda: fh.read(chunk), b""):
            h.update(blk)
    return h.hexdigest()


# ==========================================================================
# 2. Əsas iş kitabının ETL-i (oxlon_realtime/src/etl.py-dən portlanmışdır)
# ==========================================================================

def parse_header(v):
    """Zaman sütununun başlığı -> (il, ay|None, tip) və ya None.

    tip: 'annual' (çılpaq il), 'monthly' (ilin əvvəlindən artan aylıq),
    'quarterly' (rüblük kumulyativ). Nazirliyin iş kitabında üç format var:
      · "2024"                       -> illik
      · "2025 yanvar- sentyabr"      -> YTD, SON ay götürülür
      · "Mart 2021 (I rüb)", "İyun 2021 (6 ayı)" -> rüblük kumulyativ
    """
    s = str(v)
    ym = re.search(r"(19|20)\d{2}", s)
    if not ym:
        return None
    year = int(ym.group(0))
    low = _norm(s)

    # (a) RÜBLÜK budaq — Tədiyyə Balansı vərəqi bu formatı işlədir
    mrub = re.search(r"\(\s*(iv|iii|ii|i)\s*rüb", low)
    if mrub:
        return (year, 3 * ROMAN[mrub.group(1)], "quarterly")
    mays = re.search(r"\(\s*(\d{1,2})\s*ay", low)      # "(6 ayı)", "(9 ayı)"
    if mays:
        m = int(mays.group(1))
        if 1 <= m <= 12:
            return (year, m, "quarterly" if m % 3 == 0 else "monthly")

    # (b) Ay adları. Ticarət vərəqindəki "2021 yanvar- sentyabr" kimi YTD mətnində
    #     interval BAŞLANĞICI deyil, SONU götürülməlidir — əks halda kumulyativ
    #     sütun yanvara bağlanar və de-kumulyasiya tamamilə sürüşərdi.
    hits = sorted((low.rfind(st), mi) for st, mi in MONTHS.items() if st in low)
    if hits:
        first_m = hits[0][1]
        last_m = hits[-1][1]
        if len(hits) > 1 and last_m < first_m:
            raise ValueError(f"YTD interval tərsinədir: {s!r}")
        return (year, last_m, "monthly")

    return (year, None, "annual")


def _iter_rows_capped(ws, hard_cap=1500, empty_run=30, max_col=140):
    """Vərəqi oxuyur, lakin formatlaşma artefaktına görə şişmiş max_row-u kəsir.

    Tədiyyə Balansı vərəqi max_row=1048575 elan edir, real sətir sayı ~90-dır;
    ona görə ardıcıl boş sətirlərin sayı `empty_run`-a çatanda oxu dayandırılır.
    """
    out, empt = [], 0
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=hard_cap,
                                       max_col=max_col, values_only=True), 1):
        if r is None or all(c is None for c in r):
            empt += 1
            if empt >= empty_run and i > 40:
                break
            out.append((i, r if r is not None else ()))
            continue
        empt = 0
        out.append((i, r))
    return out


_ASOF_DATE = re.compile(r"(\d{1,2})[.\-/](\d{1,2})[.\-/]((?:19|20)\d{2})")

# İkinci pillə başlığındakı tanınan keyfiyyətlər (qualifier). Açar = etiketdəki axtarış
# mətni (kiçik hərflə), dəyər = seriya kodunun sonuna əlavə olunan qısa şəkilçi.
_SUBHDR_QUALIFIERS = (("nəzərdə tutul", "plan"), ("faktiki", "fakt"), ("proqnoz", "proqnoz"))


def _subheader_qualifier(text) -> str:
    """İkinci pillə başlığının mətnindən qısa şəkilçi (seriya kodu üçün)."""
    t = _norm(_squash(text))
    for needle, tag in _SUBHDR_QUALIFIERS:
        if needle in t:
            return tag
    return _slug(text)[:12]


def _asof_period(text, year):
    """"Faktiki xərc (01.04.2026)" kimi VƏZİYYƏT TARİXİNDƏN kumulyativ dövr çıxarır.

    Konvensiya: tarix ayın ilk yarısına düşürsə (gün ≤ 15) məlumat ƏVVƏLKİ ayın sonuna
    qədərdir — yəni 01.04.2026 = yanvar–mart kumulyativ (ay = 3); əks halda həmin ayın
    özü götürülür. Tarix tapılmazsa None qaytarılır və sütun illik sayılır.
    """
    m = _ASOF_DATE.search(str(text))
    if not m:
        return None
    day, mon, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if not (1 <= mon <= 12):
        return None
    mm = mon - 1 if day <= 15 else mon
    if mm <= 0:
        return (yr - 1, 12, "monthly")
    return (yr, mm, "monthly")


def load_sheet(ws, sheet):
    """Bir vərəqi oxuyur -> (zaman sütunları, sütun keyfiyyətləri, sətir yazıları).

    Başlıq sətri etiketlə tapılır ("Göstərici..." xanası), ölçü vahidi sütunu
    "Ölçü..." xanası ilə; heç bir sütun mövqeyi sabit qəbul edilmir.

    **İki pilləli başlıq.** Bəzi vərəqlərdə (`DİP 2016-2026`) başlıqdan sonrakı sətir
    ikinci pillədir: illər yuxarı sətirdə, həmin ilin İKİ fərqli oxunuşu isə aşağı
    sətirdə göstərilir — 2026 üçün "Nəzərdə tutulmuş vəsait" (nəzərdə tutulan) və
    "Faktiki xərc (01.04.2026)" (ilin əvvəlindən kumulyativ faktiki icra). Tək pilləli
    məntiq bu vərəqdə ikinci sütunu SƏSSİZ atır və nəzərdə tutulan məbləği adi illik
    müşahidə kimi oxuyardı. Ona görə funksiya ikinci pilləni tanıyır, başlığı boş olan
    sütuna ili SOLDAN miras verir və hər belə sütuna qısa keyfiyyət şəkilçisi qaytarır
    (`plan`, `fakt`); şəkilçili sütunlar ETL-də AYRICA seriya kodu alır, beləliklə
    nəzərdə tutulan rəqəm heç vaxt faktiki müşahidə ilə eyni sırada qarışmır.
    """
    rows = _iter_rows_capped(ws)
    h = None
    for k, (i, r) in enumerate(rows[:10]):
        if r and any(c is not None and str(c).strip().startswith("Göstərici") for c in r[:6]):
            h = k
            break
    if h is None:
        return None
    header = rows[h][1]
    name_col = next((j for j, c in enumerate(header)
                     if c is not None and str(c).strip().startswith("Göstərici")), 0)
    unit_col = next((j for j, c in enumerate(header)
                     if c is not None and str(c).strip().startswith("Ölçü")), name_col + 1)
    msource = next((j for j, c in enumerate(header)
                    if c is not None and str(c).strip() == "Mənbə"), None)
    hi = msource if msource is not None else len(header)

    # --- ikinci pillə başlığı: yalnız (a) ad və ölçü xanaları BOŞ, (b) zaman sütunları
    #     aralığında ən azı bir mətn xanası var, (c) həmin aralıqdakı bütün dolu xanalar
    #     MƏTNdir (ədəd deyil) olduqda qəbul edilir — əks halda sətir adi məlumat sətridir.
    sub = {}
    if h + 1 < len(rows):
        cand = rows[h + 1][1] or ()
        c_name = cand[name_col] if len(cand) > name_col else None
        c_unit = cand[unit_col] if len(cand) > unit_col else None
        span = [(j, cand[j]) for j in range(unit_col + 1, min(hi, len(cand)))
                if cand[j] is not None and str(cand[j]).strip() != ""]
        if (c_name is None or str(c_name).strip() == "") and \
           (c_unit is None or str(c_unit).strip() == "") and span and \
           all(isinstance(v, str) for _, v in span):
            sub = {j: _squash(v) for j, v in span}

    tcols, tqual = {}, {}
    last_year = None
    for j in range(unit_col + 1, hi):
        c = header[j] if j < len(header) else None
        pr = parse_header(c) if c is not None else None
        if pr:
            last_year = pr[0]
        s = sub.get(j)
        if pr and not s:
            tcols[j] = pr
            tqual[j] = ""
            continue
        if pr and s:                       # il yuxarı sətirdə, keyfiyyəti aşağı sətirdə
            asof = _asof_period(s, pr[0])
            tcols[j] = asof if asof else pr
            tqual[j] = _subheader_qualifier(s)
            continue
        if s and last_year is not None:    # başlığı boş sütun — il SOLDAN miras alınır
            asof = _asof_period(s, last_year)
            tcols[j] = asof if asof else (last_year, None, "annual")
            tqual[j] = _subheader_qualifier(s)

    fcols = [msource + 1, msource + 2, msource + 3] if msource is not None else []
    out = []
    data_start = h + 2 if sub else h + 1
    for i, r in rows[data_start:]:
        if not r:
            continue
        nm = r[name_col] if len(r) > name_col else None
        name = (str(nm).strip() if nm is not None
                else (str(r[0]).strip() if r and r[0] is not None else ""))
        if not name:
            continue
        unit = (str(r[unit_col]).strip() if len(r) > unit_col and r[unit_col] is not None else "")
        raw = {j: r[j] for j in tcols if j < len(r) and r[j] is not None
               and str(r[j]).strip() != ""}
        numeric = {j: num(r[j]) for j in raw}
        numeric = {j: x for j, x in numeric.items() if not math.isnan(x)}
        if not unit and not numeric:
            continue                      # bölmə başlığıdır, məlumat sətri deyil
        src = (str(r[msource]).strip() if msource is not None and len(r) > msource
               and r[msource] is not None else "")
        freq, freq_note = "", ""
        for k, lab in zip(fcols, ["monthly", "quarterly", "annual"]):
            if k < len(r) and r[k] is not None and str(r[k]).strip() not in ("-", ""):
                freq, freq_note = lab, str(r[k]).strip()
                break
        out.append(dict(row=i, name=name, unit=unit, source=src, freq=freq,
                        freq_note=freq_note, vals=numeric))
    return tcols, tqual, out


def _is_monotone(ym, tol=0.02):
    """İl daxilində aylıq kumulyativ həqiqətənmi artır? (illər üzrə çoxluq səsi)"""
    by_year = {}
    for (y, m), x in ym.items():
        if m is None:
            continue
        by_year.setdefault(y, []).append((m, x))
    good = bad = 0
    for y, seq in by_year.items():
        seq.sort()
        if len(seq) < 4:
            continue
        diffs = [seq[i + 1][1] - seq[i][1] for i in range(len(seq) - 1)]
        base = max((abs(v) for _, v in seq if v), default=1.0)
        good, bad = (good + 1, bad) if all(d >= -tol * base for d in diffs) else (good, bad + 1)
    return good > bad and good > 0


def value_kind(unit, ym, freq_note=""):
    """nisbət (%/indeks) / axın (YTD kumulyativ) / qalıq (səviyyə) təsnifatı.

    Sıra yalnız o halda AXIN sayılır ki, (a) hər yanvar sıfırdan başlasın və
    (b) il daxilində həqiqətən yığılsın. Vərəqin öz metaməlumatı "kumulyativ"
    yazırsa, bu, təsnifat üçün həlledici sayılır (Tədiyyə Balansı rüblük blokları).
    """
    u = _norm(unit)
    if "%" in u or "indeks" in u:
        return "rate"
    annual = {y: x for (y, m), x in ym.items() if m is None}
    jan = {y: x for (y, m), x in ym.items() if m == 1}
    ratios = [abs(jan[y] / annual[y - 1]) for y in jan
              if annual.get(y - 1) and not math.isnan(annual[y - 1]) and annual[y - 1] != 0]
    if ratios:
        # Yanvar dəyəri əvvəlki ilin cəminə nisbətdə kiçikdirsə, sıra hər il
        # sıfırdan başlayır; bu, vərəqin metaməlumatından daha etibarlı əlamətdir.
        ratios.sort()
        med = ratios[len(ratios) // 2]
        if med < 0.35:
            return "flow" if _is_monotone(ym) else "stock"
        return "stock"                       # il sərhədini keçir -> səviyyədir
    # Yanvar müşahidəsi yoxdur (məsələn Tədiyyə Balansının rüblük blokları):
    # yalnız bu halda vərəqin "kumulyativ" qeydi həlledici sayılır.
    if "kumulyativ" in _norm(freq_note):
        return "flow"
    return "flow" if _is_monotone(ym) else "stock"


def decumulate(kind, ym_vals):
    """{(il, ay)} kumulyativ -> {(il, ay)} dövr axını. Ay 0 == illik uçot yazısı."""
    out = {}
    if kind != "flow":
        for (y, m), x in ym_vals.items():
            out[(y, m if m else 0)] = x
        return out
    by_year = {}
    for (y, m), x in ym_vals.items():
        by_year.setdefault(y, {})[m] = x
    for y, d in by_year.items():
        months = sorted([m for m in d if m is not None])
        prev, last_m = 0.0, 0
        for m in months:
            out[(y, m)] = d[m] - prev
            prev, last_m = d[m], m
        if None in d:
            if last_m >= 1:
                out[(y, 12)] = d[None] - d[last_m]   # ilin qalıq hissəsi
            else:
                out[(y, 0)] = d[None]                # yalnız illik sıra
    return out


def _build_workbook_cache(verbose=True):
    """113 MB-lıq iş kitabını bir dəfə oxuyur -> registr + illik + aylıq keş."""
    if openpyxl is None:
        raise ImportError("openpyxl tələb olunur: pip install openpyxl")
    wb = openpyxl.load_workbook(XLSX_MAIN, read_only=True, data_only=True)
    nm = {n.strip(): n for n in wb.sheetnames}
    reg_rows, ann_rows, mon_rows = [], [], []
    for sheet in MAIN_SHEETS:
        if sheet not in nm:
            if verbose:
                print(f"  [ötürülür] vərəq yoxdur: {sheet}")
            continue
        res = load_sheet(wb[nm[sheet]], sheet)
        if res is None:
            if verbose:
                print(f"  [ötürülür] başlıq tapılmadı: {sheet}")
            continue
        tcols, tqual, recs = res
        n_used = 0
        for rec in recs:
            # İki pilləli başlıqlı vərəqlərdə bir sətir bir neçə OXUNUŞ daşıya bilər
            # (məsələn 2026 üçün "nəzərdə tutulmuş" və "faktiki icra"); hər keyfiyyət
            # AYRICA seriya kodu alır ki, plan rəqəmi faktiki müşahidə ilə qarışmasın.
            groups = {}
            for j, x in rec["vals"].items():
                y, m, _ptype = tcols[j]
                groups.setdefault(tqual.get(j, ""), {})[(y, m)] = x
            for qual, ym in groups.items():
                if not ym:
                    continue
                sid = f"{ABBR[sheet]}_r{rec['row']:03d}" + (f"_{qual}" if qual else "")
                name_az = _squash(rec["name"]) + (f" — {qual}" if qual else "")
                kind = value_kind(rec["unit"], ym, rec["freq_note"])
                n_used += 1
                for (y, m), x in ym.items():
                    if m is None:                     # yalnız çılpaq il = illik dəyər
                        ann_rows.append(dict(series_code=sid, year=y, value=x))
                dec = decumulate(kind, ym)
                has_month = {y for (y, m) in dec if m not in (0,)}
                for (y, m), x in dec.items():
                    if m == 0:
                        # yalnız illik müşahidə: səviyyə sırasında bu, dekabr
                        # səviyyəsidir və o ildə aylıq sıra varsa, dekabr kimi yazılır
                        if kind == "stock" and y in has_month:
                            mon_rows.append(dict(series_code=sid, year=y, month=12, value=x))
                        continue
                    mon_rows.append(dict(series_code=sid, year=y, month=m, value=x))
                yrs = sorted({y for (y, _) in ym})
                reg_rows.append(dict(series_code=sid, sheet=sheet, row=rec["row"],
                                     name_az=name_az, unit=_squash(rec["unit"]),
                                     source=rec["source"], freq=rec["freq"], kind=kind,
                                     n_obs=len(ym), year_min=yrs[0], year_max=yrs[-1]))
        if verbose:
            print(f"  {sheet:<20} -> {n_used} sıra")
    wb.close()
    reg = pd.DataFrame(reg_rows)
    assert reg["series_code"].is_unique, "iş kitabı registrində təkrar identifikator"
    ann = pd.DataFrame(ann_rows).sort_values(["series_code", "year"])
    mon = pd.DataFrame(mon_rows).sort_values(["series_code", "year", "month"])
    reg.to_csv(C_WB_REGISTRY, index=False)
    ann.to_csv(C_WB_ANNUAL, index=False)
    mon.to_csv(C_WB_MONTHLY, index=False)
    return reg, ann, mon


# ==========================================================================
# 3. Statutar şablonun (8 vərəq) parseri
# ==========================================================================

_NUMSEC = re.compile(r"^\d+(\.\d+)*\.?$")

# "real artım tempi" tipli atribut sətirləri öz-özlüyündə bölmə başlığı deyil
_ATTR_LABELS = {_key(x) for x in (
    "real artım tempi", "deflyator", "artım tempi", "əlavə dəyərin xüsusi çəkisi",
    "payı  (ÜDM-də)", "payı (ÜDM-də)", "Cəmi investisiyalarda xüsusi çəkisi",
    "ÜDM-də payı", "əlavə dəyər", "ümumi buraxılış", "aralıq istehlak",
)}


def _scan_vereq_sheet(ws, sheet):
    """Bir statutar vərəqi oxuyur.

    Qaytarır: (il xəritəsi, bölünmə xəritəsi, sətir yazıları, #REF! qeydləri).
    Hesabat/Proqnoz bölgüsü başlıq sətrindəki mətnlə müəyyən edilir və il
    sətri ilə tutuşdurulur; uyğunsuzluq halında istisna atılır.
    """
    rows = _iter_rows_capped(ws, hard_cap=400, empty_run=40, max_col=30)
    col_year, hes_col, proq_col = {}, None, None

    for i, r in rows:
        if not r:
            continue
        # başlıq sətri: birinci xana "№"
        if r[0] is not None and _squash(r[0]) == "№":
            for j, c in enumerate(r):
                if c is None:
                    continue
                t = _key(c)
                if t == "hesabat":
                    hes_col = j if hes_col is None else min(hes_col, j)
                elif t == "proqnoz":
                    proq_col = j if proq_col is None else min(proq_col, j)
            continue
        # il sətri: ən azı 5 dörd rəqəmli il dəyəri
        yrs = {j: int(c) for j, c in enumerate(r)
               if isinstance(c, (int, float)) and not isinstance(c, bool)
               and float(c).is_integer() and 1990 <= int(c) <= 2050}
        if len(yrs) >= 5:
            for j, y in yrs.items():
                if j in col_year and col_year[j] != y:
                    raise ValueError(f"{sheet}: {j} sütununda il ziddiyyəti "
                                     f"({col_year[j]} vs {y})")
                col_year[j] = y

    if not col_year:
        raise ValueError(f"{sheet}: il sətri tapılmadı")
    if hes_col is None or proq_col is None:
        raise ValueError(f"{sheet}: Hesabat/Proqnoz başlığı tapılmadı")

    # Hesabat/Proqnoz bölgüsü şablonun ÖZ başlıq mətni ilə müəyyən edilir və
    # ardıcıllığı MƏCBURİ yoxlanılır (D1). Şablonun sərhədi burada fərz edilmir:
    # 8 vərəq daha köhnə buraxılışdır və onun Hesabat bloku 2024-cü ildə bitir,
    # yəni 2025 sütunu Nazirliyin proqnozudur, faktiki nəticə deyil.
    col_split = {}
    for j, y in col_year.items():
        col_split[j] = "proqnoz" if j >= proq_col else "hesabat"
    hes_years = sorted(y for j, y in col_year.items() if col_split[j] == "hesabat")
    proq_years = sorted(y for j, y in col_year.items() if col_split[j] == "proqnoz")
    if not hes_years or not proq_years:
        raise ValueError(f"{sheet}: Hesabat və ya Proqnoz bloku boşdur")
    if max(hes_years) >= min(proq_years):
        raise ValueError(f"{sheet}: Hesabat ({max(hes_years)}) və Proqnoz "
                         f"({min(proq_years)}) blokları kəsişir")
    if max(hes_years) > LAST_ACTUAL:
        raise ValueError(f"{sheet}: Hesabat bloku {max(hes_years)} ilinə qədər uzanır, "
                         f"faktiki məlumatın son ili isə {LAST_ACTUAL}-dir")
    recs, refs = _scan_rows(rows, col_year, col_split, sheet)
    return dict(col_year=col_year, col_split=col_split, hesabat_last=max(hes_years),
                proqnoz_first=min(proq_years), recs=recs, refs=refs)


def _scan_rows(rows, col_year, col_split, sheet):
    """Vərəqin məlumat sətirlərini bölmə/blok iyerarxiyası ilə oxuyur."""
    recs, refs = [], []
    section_no, section_lab, block_lab = "", "", ""
    for i, r in rows:
        if not r:
            continue
        c0 = _squash(r[0]) if r[0] is not None else ""
        c1 = _squash(r[1]) if len(r) > 1 and r[1] is not None else ""
        c2 = _squash(r[2]) if len(r) > 2 and r[2] is not None else ""
        if c0 == "№":
            section_no, section_lab = "", ""
            continue
        if not c1 and not c2:
            if c0 and not _NUMSEC.match(c0):          # blok başlığı
                block_lab = c0
                section_no, section_lab = "", ""
            continue
        has_num = any(isinstance(c, (int, float)) and not isinstance(c, bool)
                      for j, c in enumerate(r) if j in col_year)
        # bölmə başlığı: ölçü vahidi yoxdur və ədədi dəyər yoxdur
        if not c2 and not has_num:
            if _NUMSEC.match(c0):
                section_no, section_lab = c0, c1
            elif c1 and _key(c1) not in _ATTR_LABELS:
                section_no, section_lab = "", c1
            continue
        if not c2:
            continue                                   # ölçü vahidsiz məlumat sətri sayılmır
        if _NUMSEC.match(c0):
            section_no, section_lab = c0, c1
        if c1 == REF_TOKEN:
            # Sətrin ADI da #REF! qaytarır: göstərici adsız qalır, ona görə
            # susdurulmur, ayrıca qeyd olunur (kompozit açar sətirlə işlədiyi
            # üçün sıra yenə də unikal qalır).
            refs.append(dict(sheet=sheet, row=i, name_az=c1, unit=c2,
                             column=1, year=-1, split="ad"))
        vals, has_ref = {}, []
        for j, y in col_year.items():
            c = r[j] if j < len(r) else None
            if c is None:
                continue
            if isinstance(c, str) and c.strip() == REF_TOKEN:
                has_ref.append(y)
                refs.append(dict(sheet=sheet, row=i, name_az=c1, unit=c2,
                                 column=j, year=y, split=col_split[j]))
                continue
            x = num(c)
            if not math.isnan(x):
                vals[y] = (x, col_split[j])
        recs.append(dict(sheet=sheet, row=i, no=c0, name_az=c1, unit=c2,
                         section_no=section_no, section=section_lab, block=block_lab,
                         vals=vals, ref_years=has_ref))
    return recs, refs


def _build_vereq_cache(verbose=True):
    """8 vərəq şablonunu oxuyur -> uzun format, məhsul kataloqu, #REF! hesabatı."""
    if openpyxl is None:
        raise ImportError("openpyxl tələb olunur: pip install openpyxl")
    wb = openpyxl.load_workbook(XLSX_VEREQ, read_only=True, data_only=True)
    long_rows, prod_rows, ref_rows, meta_rows = [], [], [], []
    for sheet, slug in VEREQ_SHEETS.items():
        if sheet not in wb.sheetnames:
            raise KeyError(f"statutar vərəq yoxdur: {sheet}")
        sc = _scan_vereq_sheet(wb[sheet], sheet)
        col_year, recs, refs = sc["col_year"], sc["recs"], sc["refs"]
        meta_rows.append(dict(sheet=sheet, hesabat_last=sc["hesabat_last"],
                              proqnoz_first=sc["proqnoz_first"],
                              year_min=min(col_year.values()),
                              year_max=max(col_year.values()),
                              n_rows=len(recs), n_ref=len(refs)))
        ref_rows.extend(refs)
        for rec in recs:
            # Kompozit identifikator (vərəq, sətir, ölçü vahidi) — FR4 kataloqunda
            # 10 ad təkrarlanır, ona görə ad heç vaxt açar kimi işlədilmir.
            sid = f"{slug}_r{rec['row']:03d}_{_slug(rec['unit'])}"
            prod_rows.append(dict(series_id=sid, sheet=sheet, row=rec["row"],
                                  no=rec["no"], name_az=rec["name_az"], unit=rec["unit"],
                                  section_no=rec["section_no"], section=rec["section"],
                                  block=rec["block"],
                                  n_hesabat=sum(1 for v in rec["vals"].values() if v[1] == "hesabat"),
                                  n_proqnoz=sum(1 for v in rec["vals"].values() if v[1] == "proqnoz"),
                                  n_ref=len(rec["ref_years"])))
            for y, (x, sp) in rec["vals"].items():
                long_rows.append(dict(series_id=sid, sheet=sheet, row=rec["row"],
                                      name_az=rec["name_az"], unit=rec["unit"],
                                      year=y, value=x, split=sp))
        if verbose:
            print(f"  {sheet:<16} -> {len(recs)} sətir, {len(refs)} #REF! xanası, "
                  f"Hesabat ≤{sc['hesabat_last']}, Proqnoz ≥{sc['proqnoz_first']}")
    wb.close()
    lg = pd.DataFrame(long_rows)
    pr = pd.DataFrame(prod_rows)
    rf = pd.DataFrame(ref_rows)
    mt = pd.DataFrame(meta_rows)
    assert pr["series_id"].is_unique, "statutar şablonda kompozit identifikator təkrarlanır"
    # Şablonun buraxılış sərhədi bütün vərəqlərdə eyni olmalıdır
    assert mt["hesabat_last"].nunique() == 1, \
        f"vərəqlərdə fərqli Hesabat sərhədi: {sorted(mt['hesabat_last'].unique())}"
    lg.to_csv(C_VEREQ_LONG, index=False)
    pr.to_csv(C_VEREQ_PRODUCTS, index=False)
    rf.to_csv(C_VEREQ_REF, index=False)
    mt.to_csv(C_VEREQ_META, index=False)
    return lg, pr, rf, mt


# ==========================================================================
# 4. Seriya kataloqu (FR-lərin tələb etdiyi illik sıralar)
# ==========================================================================

def V(sheet, label, row=None, sec_no=None, sec_prefix=None, block_prefix=None, unit=None):
    """Statutar şablon mənbəyi. Sətir nömrəsi verilsə belə etiket yoxlanılır."""
    return ("vereq", dict(sheet=sheet, label=label, row=row, sec_no=sec_no,
                          sec_prefix=sec_prefix, block_prefix=block_prefix, unit=unit))


def W(sheet, row, label, tr=None, qual=None):
    """Əsas iş kitabı mənbəyi: sətir nömrəsi + MƏCBURİ etiket yoxlaması.

    `qual` iki pilləli başlıqlı vərəqlərdə (məsələn `DİP 2016-2026`) sətrin hansı
    OXUNUŞUNUN götürüldüyünü göstərir: `"plan"` — nəzərdə tutulmuş vəsait,
    `"fakt"` — ilin əvvəlindən faktiki icra. Boş buraxıldıqda əsas (illik) oxunuş alınır.

    `tr` çevrilməsi mənbənin öz konvensiyasını statutar şablonun konvensiyasına
    gətirir:
      · 'idx100' — iş kitabı real artımı "əvvəlki il = 100" indeksi kimi verir,
        şablon isə faiz dəyişməsi kimi; ona görə 100 çıxılır (0,1 f.b. yuvarlaqlaşma
        iş kitabının öz dəqiqliyidir);
      · 'neg'    — tədiyə balansında idxal və ödənişlər debet (mənfi) yazılır,
        şablonda isə müsbət kəmiyyət kimi göstərilir.
    """
    assert tr in (None, "idx100", "neg"), f"naməlum çevrilmə: {tr}"
    return ("wb", dict(sheet=sheet, row=row, label=label, tr=tr, qual=qual))


def C(path, column):
    """CSV panel mənbəyi."""
    return ("csv", dict(path=path, column=column))


def PB(var, label, source_key, panel="public"):
    """Dondurulmuş açıq-mənbə paneli (§14) və ya gömrük paneli (§15) mənbəyi.

    Sətir nömrəsi ilə deyil, sıra adı (`var`) ilə oxunur, LAKİN oxunuş anında provenans
    cədvəlindəki mənbə faylı və ETİKET yoxlanılır: paneli quran kod dəyişsə və sıra başqa
    sətirdən gəlsə, kataloq həllində dərhal istisna atılır. Beləliklə çılpaq mövqe indeksi
    ilə oxu burada da yoxdur — panel §13-ün eyni konvensiyası ilə işləyir.
    """
    assert panel in ("public", "customs"), f"naməlum panel: {panel}"
    return ("panel", dict(var=var, label=label, source_key=source_key, panel=panel))


def A(sheet, row, label, min_months=6):
    """Aylıq səviyyələrin illik ORTALAMASI (etiket yoxlaması ilə).

    Orta illik məzənnə iş kitabında ayrıca sıra kimi verilmir; dövrün sonuna
    aylıq müşahidələrin ortalaması standart yaxınlaşdırmadır və provenansı
    `wb-avg:` prefiksi ilə açıq göstərilir.
    """
    return ("wbavg", dict(sheet=sheet, row=row, label=label, min_months=min_months))


V1, V2, V3, V4, V57, V1314 = ("2.4.1.1.", "2.4.1.2.", "2.4.1.3.", "2.4.1.4.",
                              "2.4.1.5.-7.", "2.4.1.13.-14.")

# --- FR2/FR3 sahə blokları: 2.4.1.2. vərəqinin 11 baş bölməsi ---------------
# Hər sahə üçün iki mənbə saxlanılır: əsas iş kitabının Real sektor vərəqi
# (2000/2005–2025, vahid buraxılış) və statutar şablon (2013–2024 Hesabat +
# 2026–2030 Proqnoz). Səviyyə sıraları üçün ƏSAS mənbə iş kitabıdır, çünki
# statutar şablonun Hesabat bloku 2024-cü ildə bitir; şablon isə Nazirliyin
# baza ssenarisini (ministry_baseline) təmin edir.
# (bölmə nömrəsi, kod kökü, AZ ad, EN ad, bölmə prefiksi,
#  Real sektor: əlavə dəyər sətri + etiketi, ümumi buraxılış sətri + etiketi)
_SECTORS = [
    ("1.", "agri", "Kənd, meşə və balıqçılıq təsərrüfatı", "Agriculture, forestry and fishing",
     "Kənd, meşə", 27, "Kənd, meşə və balıqçılıq təsərrüfatları", 63, "Kənd, meşə və balıqçılıq təsərrüfatları"),
    ("2.", "mining", "Mədənçıxarma sənayesi", "Mining and quarrying",
     "Mədənçıxarma sənayesi", 19, "Mədənşıxarma sənayesi üzrə", 55, "Mədənçıxarma sənayesi üzrə"),
    ("3.", "manuf", "Emal sənayesi", "Manufacturing",
     "Emal sənayesi", 21, "Emal sənayesi üzrə", 57, "Emal sənayesi üzrə"),
    ("4.", "power", "Elektrik enerjisi, qaz və buxar istehsalı və təchizatı",
     "Electricity, gas and steam supply", "Elektrik enerjisi",
     23, "Elektrik enerjisi, qaz və buxar istehsalı, bölüşdürülməsi və təchizatı üzrə",
     59, "Elektrik enerjisi, qaz və buxar istehsalı, bölüşdürülməsi və təchizatı üzrə"),
    ("5.", "water", "Su təchizatı, tullantıların təmizlənməsi və emalı",
     "Water supply and waste management", "Su təchizatı",
     25, "Su təchizatı; tullantıların təmizlənməsi və emalı üzrə",
     61, "Su təchizatı; tullantıların təmizlənməsi və emalı üzrə"),
    ("6.", "constr", "Tikinti", "Construction", "Tikinti", 29, "Tikinti", 65, "Tikinti"),
    ("7.", "trade", "Ticarət, nəqliyyat vasitələrinin təmiri", "Trade and repair of vehicles",
     "Ticarət", 33, "Ticarət; nəqliyyat vasitələrinin təmiri", 69, "Ticarət; nəqliyyat vasitələrinin təmiri"),
    ("8.", "tourism", "Turistlərin yerləşdirilməsi və ictimai iaşə", "Accommodation and food service",
     "Turistlərin yerləşdirilməsi", 35, "Turistlərin yerləşdirilməsi və ictimai iaşə",
     71, "Turistlərin yerləşdirilməsi və ictimai iaşə"),
    ("9.", "transport", "Nəqliyyat və anbar təsərrüfatı", "Transport and storage",
     "Nəqliyyat və anbar", 37, "Nəqliyyat və anbar təsərrüfatı", 73, "Nəqliyyat və anbar təsərrüfatı"),
    ("10.", "ict", "İnformasiya və rabitə", "Information and communication",
     "İnformasiya və rabitə", 39, "İnformasiya və rabitə", 75, "İnformasiya və rabitə"),
    ("11.", "social", "Sosial və digər sahələrdə yaradılmış əlavə dəyər", "Social and other services",
     "Sosial və digər", 41, "Sosial və digər xidmətlər", 77, "Sosial və digər xidmətlər"),
]


def _price_basis(unit, explicit=None):
    """Qiymət bazasının (D5) həlli: açıq verilibsə o, əks halda ölçü vahidindən susmaya görə.

    Qayda `outputs.price_basis_for_unit`-dədir — müqavilə sütununun sahibi həmin moduldur,
    burada yalnız çağırılır (tək mənbə prinsipi).
    """
    if explicit:
        return explicit
    try:
        from outputs import price_basis_for_unit   # type: ignore
    except ImportError:                            # paket kimi idxal edildikdə
        from src.outputs import price_basis_for_unit   # type: ignore
    return price_basis_for_unit(unit)


def _sector_entries():
    """11 sahə üzrə ümumi buraxılış / aralıq istehlak / əlavə dəyər / artım tempi."""
    out = []
    RS = "Real sektor"
    for (sec_no, root, name_az, name_en, prefix,
         va_row, va_lab, go_row, go_lab) in _SECTORS:
        out.append(dict(code=f"go_{root}", unit="mln AZN",
                        name_az=f"{name_az} — ümumi buraxılış",
                        name_en=f"{name_en} — gross output",
                        src=[W(RS, go_row, go_lab),
                             V(V2, "ümumi buraxılış", sec_no=sec_no, sec_prefix=prefix)]))
        out.append(dict(code=f"ic_{root}", unit="mln AZN",
                        name_az=f"{name_az} — aralıq istehlak",
                        name_en=f"{name_en} — intermediate consumption",
                        src=[V(V2, "aralıq istehlak", sec_no=sec_no, sec_prefix=prefix)]))
        out.append(dict(code=f"va_{root}", unit="mln AZN",
                        name_az=f"{name_az} — əlavə dəyər",
                        name_en=f"{name_en} — value added",
                        src=[W(RS, va_row, va_lab),
                             V(V2, "əlavə dəyər", sec_no=sec_no, sec_prefix=prefix)]))
        out.append(dict(code=f"g_{root}", unit="%",
                        name_az=f"{name_az} — real artım tempi",
                        name_en=f"{name_en} — real growth rate",
                        src=[W(RS, va_row + 1, "real artım tempi", tr="idx100"),
                             V(V2, "real artım tempi", sec_no=sec_no, sec_prefix=prefix)]))
    for e in out:
        e["fr"], e["ref"] = "FR2/FR3", V2
    return out


def _catalog_spec():
    """Kataloqun tam təsviri: hər sıra üçün mənbə zənciri və 2025 üstələməsi.

    `src` siyahısındakı BİRİNCİ mənbə əsasdır; sonrakılar yalnız birincidə
    olmayan (adətən daha erkən) illəri doldurur. `official` açarı verilibsə,
    2025-ci il dəyəri DSK-nın rəsmi nəticəsi ilə əvəzlənir və köhnə buraxılış
    `vintage_overrides.csv` faylında qeyd olunur (8 vərəq şablonu 2025 üçün
    daha köhnə buraxılış saxlayır).
    """
    E = []

    def add(code, name_az, name_en, unit, fr, src, ref, official=None, note="", pbasis=None):
        # `pbasis` (D5) — qiymət bazası; verilmədikdə ölçü vahidindən çıxarılır. Kataloqdakı
        # bütün səviyyə sıraları FAKTİKİ müşahidələrdir, yəni CARİ qiymətlərdədir; sabit qiymətli
        # (proqnoz) sıralar dəftərlərdə ayrıca kodla nəşr olunur (məsələn `inv_total_real2025`).
        E.append(dict(code=code, name_az=name_az, name_en=name_en, unit=unit,
                      fr=fr, src=src, ref=ref, official=official, note=note,
                      price_basis=_price_basis(unit, pbasis)))

    # ---------------- FR1: ÜDM nüvəsi ------------------------------------
    add("gdp_nom", "ÜDM bazar qiymətləri ilə", "GDP at market prices", "mln AZN", "FR1",
        [V(V1, "ÜDM bazar qiymətləri ilə", row=4), C(P_MACRO, "gdp_nom")], V1,
        official="gdp_nom")
    add("gdp_realg", "ÜDM real artım tempi", "GDP real growth rate", "%", "FR1",
        [V(V1, "real artım tempi", row=5), C(P_MACRO, "gdp_realg")], V1,
        official="gdp_realg")
    add("gdp_defl", "ÜDM deflyatoru", "GDP deflator", "%", "FR1",
        [V(V1, "deflyator", row=6), C(P_MACRO, "gdp_defl")], V1)
    add("oil_nom", "Neft-qaz sektoru ÜDM-i", "Oil and gas sector GDP", "mln AZN", "FR1",
        [V(V1, "Neft-qaz sektoru, ÜDM, bazar qiymətləri ilə", row=7), C(P_MACRO, "oil_nom")], V1,
        official="oil_nom")
    add("oil_realg", "Neft-qaz sektorunun real artım tempi", "Oil and gas real growth", "%", "FR1",
        [V(V1, "real artım tempi", row=8), C(P_MACRO, "oil_realg")], V1, official="oil_realg")
    add("oil_defl", "Neft-qaz sektorunun deflyatoru", "Oil and gas deflator", "%", "FR1",
        [V(V1, "deflyator", row=9)], V1)
    add("oil_share", "Neft-qaz sektorunun ÜDM-də payı", "Oil and gas share in GDP", "%", "FR1",
        [V(V1, "payı  (ÜDM-də)", row=10), C(P_MACRO, "oil_share")], V1)
    add("nonoil_nom", "Qeyri neft-qaz sektoru ÜDM-i", "Non-oil-gas GDP", "mln AZN", "FR1",
        [V(V1, "Qeyri neft-qaz sektoru, ÜDM, bazar qiymətləri ilə", row=11),
         C(P_MACRO, "nonoil_nom")], V1, official="nonoil_nom")
    add("nonoil_realg", "Qeyri neft-qaz sektorunun real artım tempi", "Non-oil-gas real growth",
        "%", "FR1", [V(V1, "real artım tempi", row=12), C(P_MACRO, "nonoil_realg")], V1,
        official="nonoil_realg")
    add("nonoil_defl", "Qeyri neft-qaz sektorunun deflyatoru", "Non-oil-gas deflator", "%", "FR1",
        [V(V1, "deflyator", row=13), C(P_MACRO, "nonoil_defl")], V1)
    add("nonoil_share", "Qeyri neft-qaz sektorunun ÜDM-də payı", "Non-oil-gas share in GDP",
        "%", "FR1", [V(V1, "payı  (ÜDM-də)", row=14), C(P_MACRO, "nonoil_share")], V1)
    # Mülkiyyət bölgüsü (dövlət/qeyri-dövlət) — statutar şablonda yoxdur,
    # əsas iş kitabının Real sektor vərəqindən götürülür; dövlət payı FR1-də
    # 100 − özəl pay eyniliyi ilə hesablanır.
    add("private_share_gdp", "Özəl sektorun ÜDM-də payı", "Private sector share in GDP",
        "%", "FR1", [W("Real sektor", 10, "Özəl sektorun ÜDM-də payı")], "Real sektor")
    add("private_share_nonoil", "Özəl sektorun qeyri neft-qaz ÜDM-də payı",
        "Private sector share in non-oil GDP", "%", "FR1",
        [W("Real sektor", 11, "Özəl sektorun qeyri-neft-qaz ÜDM-də payı")], "Real sektor")
    add("pop_total", "Əhalinin sayı (dövrün sonuna)", "Population, end of period",
        "min nəfər", "FR1",
        [V(V1, "İşğal olunmuş torpaqlar nəzərə alınmaqla (dövrün sonuna)", row=17)], V1)
    add("pop_avg", "Əhalinin orta illik sayı", "Population, annual average", "min nəfər", "FR1",
        [V(V1, "İşğal olunmuş torpaqlar nəzərə alınmadan (orta illik)", row=18),
         C(P_MACRO, "pop_avg")], V1)
    add("gdp_pc_azn", "Adambaşına düşən ÜDM (manatla)", "GDP per capita (AZN)", "AZN", "FR1",
        [V(V1, "Manatla", row=20), C(P_MACRO, "gdp_pc_azn")], V1, official="gdp_pc_azn")
    add("gdp_pc_realg", "Adambaşına düşən ÜDM-in real artım tempi", "GDP per capita real growth",
        "%", "FR1", [V(V1, "real artım tempi", row=22)], V1)

    # ---------------- FR2/FR3: sahə strukturu ----------------------------
    for e in _sector_entries():
        E.append(dict(code=e["code"], name_az=e["name_az"], name_en=e["name_en"],
                      unit=e["unit"], fr=e["fr"], src=e["src"], ref=e["ref"],
                      official=None, note="", price_basis=_price_basis(e["unit"])))
    add("net_taxes", "Məhsula və idxala xalis vergilər", "Net taxes on products and imports",
        "mln AZN", "FR3",
        [W("Real sektor", 43, "Məhsula və idxala xalis vergilər"),
         V(V3, "əlavə dəyər", sec_no="5.", sec_prefix="Məhsula və idxala")], V3)
    add("net_taxes_g", "Məhsula və idxala xalis vergilərin real artım tempi",
        "Net taxes real growth", "%", "FR3",
        [W("Real sektor", 44, "real artım tempi", tr="idx100"),
         V(V3, "real artım tempi", sec_no="5.", sec_prefix="Məhsula və idxala")], V3)
    add("va_industry", "Sənaye — əlavə dəyər", "Industry — value added", "mln AZN", "FR3",
        [W("Real sektor", 15, "Sənaye"),
         V(V3, "əlavə dəyər", sec_no="2.", sec_prefix="Sənaye")], V3)
    add("va_nonoil_industry", "Qeyri neft-qaz sənayesi — əlavə dəyər",
        "Non-oil industry — value added", "mln AZN", "FR3",
        [W("Real sektor", 17, "Qeyri neft-qaz sənayesi"),
         V(V3, "əlavə dəyər", sec_no="2.2.", sec_prefix="Qeyri neft-qaz sənayesi")], V3)
    add("va_services", "Xidmətlərin istehsalı — əlavə dəyər", "Services — value added",
        "mln AZN", "FR3",
        [W("Real sektor", 31, "Xidmətlərin istehsalı"),
         V(V3, "əlavə dəyər", sec_no="4.", sec_prefix="Xidmətlərin istehsalı")], V3)
    add("go_total", "Cəmi məhsul buraxılışı", "Total gross output", "mln AZN", "FR2",
        [W("Real sektor", 46, "Cəmi buraxılış")], "Real sektor")

    # ---------------- FR5: neftin dünya qiyməti --------------------------
    add("brent_usd", "Neftin dünya bazarında qiyməti (Brent)", "World oil price (Brent)",
        "USD/barel", "FR5",
        [W("Neft-Qaz sektoru", 11, '"Brent" markalı neftin qiyməti, 1 barel (eia.gov)'),
         V(V57, "Neftin qiyməti", row=4, block_prefix="Neftin dünya"),
         C(P_MACRO, "oil_price_usd")], V57)
    add("oil_export_price", "Neftin ixrac qiyməti", "Crude oil export price", "USD/barel", "FR5",
        [W("Neft-Qaz sektoru", 10, "Neftin ixrac qiyməti, 1 barel, DGK*")], "Neft-Qaz sektoru")

    # ---------------- FR6: investisiyalar --------------------------------
    # Səviyyə sıralarında əsas mənbə iş kitabıdır (1995/2000–2025); statutar
    # şablon eyni anlayışı 2013–2024 üçün verir və şablon-nümunə ssenarisini
    # (2026–2030) təmin edir.
    add("inv_total", "Əsas kapitala cəmi investisiyalar", "Total fixed capital investment",
        "mln AZN", "FR6",
        [W("Real sektor", 80, "Əsas kapitala cəmi investisiyalar"),
         V(V57, "Əsas kapitala cəmi investisiyalar", row=10)], V57, official="invest_total")
    add("inv_total_g", "Əsas kapitala investisiyaların real artım tempi",
        "Fixed capital investment real growth", "%", "FR6",
        [W("Real sektor", 81, "real artım tempi", tr="idx100"), V(V57, "real artım tempi", row=11)], V57)
    add("inv_domestic", "Daxili investisiya", "Domestic investment", "mln AZN", "FR6",
        [W("Real sektor", 86, "Daxili investisiyalar"),
         V(V57, "Daxili investisiya", row=13)], V57)
    add("inv_domestic_g", "Daxili investisiyanın real artım tempi",
        "Domestic investment real growth", "%", "FR6",
        [W("Real sektor", 87, "real artım tempi", tr="idx100"), V(V57, "real artım tempi", row=14)], V57)
    add("inv_foreign", "Xarici investisiya", "Foreign investment", "mln AZN", "FR6",
        [W("Real sektor", 92, "Xarici investisiyalar"),
         V(V57, "Xarici investisiya", row=16)], V57)
    add("inv_foreign_g", "Xarici investisiyanın real artım tempi",
        "Foreign investment real growth", "%", "FR6",
        [W("Real sektor", 93, "real artım tempi", tr="idx100"), V(V57, "real artım tempi", row=17)], V57)
    add("inv_oil", "Neft-qaz sektoruna investisiyalar", "Investment in oil and gas sector",
        "mln AZN", "FR6",
        [W("Real sektor", 82, "Neft-qaz sektoruna investisiyalar"),
         V(V57, "Neft və qaz sektoru", row=19)], V57)
    add("inv_nonoil", "Qeyri neft-qaz sektoruna investisiyalar", "Investment in non-oil sectors",
        "mln AZN", "FR6",
        [W("Real sektor", 84, "Qeyri neft-qaz sektoruna investisiyalar"),
         V(V57, "Digər sahələr", row=20)], V57)
    add("inv_state", "Dövlət investisiyaları", "State investment", "mln AZN", "FR6",
        [W("Real sektor", 98, "Dövlət investisiyaları"), V(V57, "Dövlət", row=22)], V57)
    add("inv_state_g", "Dövlət investisiyalarının real artım tempi",
        "State investment real growth", "%", "FR6",
        [W("Real sektor", 99, "real artım tempi", tr="idx100")], "Real sektor")
    add("inv_nonstate", "Qeyri-dövlət investisiyaları", "Non-state investment", "mln AZN", "FR6",
        [W("Real sektor", 100, "Qeyri-dövlət investisiyaları"),
         V(V57, "Qeyri-dövlət", row=24)], V57)
    add("inv_nonstate_g", "Qeyri-dövlət investisiyalarının real artım tempi",
        "Non-state investment real growth", "%", "FR6",
        [W("Real sektor", 101, "real artım tempi", tr="idx100")], "Real sektor")

    # ---------------- FR8: əhalinin gəlirləri, xərcləri, məşğulluq -------
    add("income_nom", "Əhalinin pul gəlirləri", "Money income of population", "mln AZN", "FR8",
        [W("Sosial sektor", 35, "Əhalinin nominal gəlirlərin ümumi dəyəri"),
         V(V57, "Gəlirlər", row=30)], V57, official="population_income")
    add("income_realg", "Əhalinin gəlirlərinin real artım tempi", "Real growth of income",
        "%", "FR8", [V(V57, "real artım tempi", row=31)], V57)
    add("disposable_income", "Sərəncamında qalan gəlirlər", "Disposable income", "mln AZN", "FR8",
        [W("Sosial sektor", 37, "Sərəncamında qalan gəlirlər")], "Sosial sektor")
    add("expend_nom", "Əhalinin xərcləri", "Expenditure of population", "mln AZN", "FR8",
        [W("Sosial sektor", 46, "Əhalinin nominal xərclərinin ümumi dəyəri"),
         V(V57, "Xərclər", row=32)], V57)
    add("expend_realg", "Əhalinin xərclərinin real artım tempi", "Real growth of expenditure",
        "%", "FR8", [V(V57, "real artım tempi", row=33)], V57)
    add("final_consumption", "Son istehlak xərcləri", "Final consumption expenditure",
        "mln AZN", "FR8", [W("Sosial sektor", 47, "Son istehlak xərcləri")], "Sosial sektor",
        note="Sosial sektor vərəqi, sətir 47, 2000–2025 — ƏHALİNİN hesabı üzrə son istehlak "
             "xərcləridir. Milli hesablar mənasında DÖVLƏTİN son istehlak xərcləri bu sırada "
             "DEYİL; DATA-GAP bu buraxılışda BAĞLANDI — dövlət ayağı `gov_cons_total` kodu ilə "
             "DSK-nın 27-ci cədvəlindən (1993–2025) məlumat qatındadır və FR01-in tələb "
             "tənliyində ayrıca sürücü kimi sınaqdan keçirilir.")
    add("wage_avg", "Orta aylıq əmək haqqı", "Average monthly wage", "AZN", "FR8",
        [W("Sosial sektor", 52, "Muzdla çalışan işçilərin orta aylıq nominal əmək haqqı"),
         V(V57, "Orta aylıq əmək haqqı", row=34)], V57)
    add("wage_avg_realg", "Orta aylıq əmək haqqının real artım tempi", "Real wage growth",
        "%", "FR8", [V(V57, "real artım tempi", row=35)], V57)
    add("wage_oil", "Neft-qaz sektorunda orta aylıq əmək haqqı", "Average wage, oil sector",
        "AZN", "FR8",
        [W("Sosial sektor", 53, "Neft-qaz sektorunda ƏH"), V(V57, "Neft", row=37)], V57)
    add("wage_nonoil", "Qeyri neft-qaz sektorunda orta aylıq əmək haqqı",
        "Average wage, non-oil sector", "AZN", "FR8",
        [W("Sosial sektor", 54, "Qeyri neft-qaz sektorunda ƏH"),
         V(V57, "Qeyri-neft", row=39)], V57)
    add("wage_state", "Dövlət sektorunda orta aylıq əmək haqqı", "Average wage, state sector",
        "AZN", "FR8", [W("Sosial sektor", 55, "Dövlət sektorunda ƏH")], "Sosial sektor")
    add("wage_private", "Özəl sektorda orta aylıq əmək haqqı", "Average wage, private sector",
        "AZN", "FR8", [W("Sosial sektor", 56, "Özəl sektorda ƏH")], "Sosial sektor")
    add("paid_services", "Əhaliyə göstərilən ödənişli xidmətlər", "Paid services to population",
        "mln AZN", "FR8",
        [W("Sosial sektor", 27, "Əhaliyə göstərilən ödənişli xidmətlərin dəyəri"),
         V(V57, "Əhaliyə göstərilən ödənişli xidmətlər", row=41)], V57)
    add("paid_services_g", "Ödənişli xidmətlərin real artım tempi", "Paid services real growth",
        "%", "FR8",
        [W("Sosial sektor", 28, "real artım tempi", tr="idx100"), V(V57, "real artım tempi", row=42)], V57)
    add("employment", "Ümumi məşğulluq", "Total employment", "min nəfər", "FR8",
        [W("Sosial sektor", 59, "Məşğul əhalinin sayı"),
         V(V57, "Ümumi məşğulluq", row=43)], V57)
    add("employment_g", "Məşğulluğun artım tempi", "Employment growth", "%", "FR8",
        [V(V57, "Artım tempi", row=44)], V57)
    add("labour_force", "İqtisadi fəal əhalinin sayı", "Labour force", "min nəfər", "FR8",
        [W("Sosial sektor", 58, "İqtisadi fəal əhalinin sayı")], "Sosial sektor")
    add("unemployment_rate", "İşsizlik səviyyəsi", "Unemployment rate", "%", "FR8",
        [W("Sosial sektor", 65, "İşsizlik səviyyəsi")], "Sosial sektor")

    # ---------------- FR8: DVX əmək haqqı fondu bloku (v3-P0-13) ---------
    # `DVX üzrə göstəricilər` vərəqinin 104–110-cu sətirləri əmək haqqı fondunu, muzdla
    # işləyənlərin sayını və orta aylıq əmək haqqını neft-qaz / qeyri-neft və dövlət /
    # qeyri-dövlət kəsimləri ilə verir. ƏHATƏ: yalnız 2021–2025 (beş illik müşahidə) —
    # bu, əmək haqqı fondu × məşğulluq zəncirinin LÖVBƏRİ və çarpaz yoxlaması üçün
    # kifayətdir, qiymətləndirilmiş elastiklik üçün deyil; sıralar bu qeydlə nəşr olunur.
    DVX = "DVX üzrə göstəricilər"
    _DVX_NOTE = ("DVX vərəqi, 2021–2025 (5 illik müşahidə) — lövbər və çarpaz yoxlama üçündür; "
                 "bu qısa nümunə üzərində elastiklik qiymətləndirilmir.")
    add("wage_fund_total", "Cəmi əmək haqqı fondu", "Total wage fund", "mln AZN", "FR8",
        [W(DVX, 104, "Cəmi əmək haqqı fondu")], DVX, note=_DVX_NOTE)
    add("wage_fund_nonoil", "Qeyri neft-qaz sektorunda cəmi əmək haqqı fondu",
        "Wage fund, non-oil-gas sector", "mln AZN", "FR8",
        [W(DVX, 105, "Qeyri-neft qaz sektorunda cəmi əmək haqqı fondu")], DVX, note=_DVX_NOTE)
    add("wage_fund_nonoil_nonstate", "Qeyri neft-qaz qeyri-dövlət bölməsində əmək haqqı fondu",
        "Wage fund, non-oil non-state", "mln AZN", "FR8",
        [W(DVX, 106, "Qeyri-neft qeyri-dövlət bölməsində əmək haqqı fondu")], DVX, note=_DVX_NOTE)
    add("employees_dvx", "Muzdla işləyənlərin sayı (orta aylıq, DVX)",
        "Employees, monthly average (tax service)", "min nəfər", "FR8",
        [W(DVX, 107, "Muzdla işləyənlərin sayı (orta aylıq)")], DVX, note=_DVX_NOTE)
    add("wage_avg_dvx", "Orta aylıq əmək haqqı (DVX)", "Average monthly wage (tax service)",
        "AZN", "FR8", [W(DVX, 108, "Orta aylıq əmək haqqı")], DVX,
        note=_DVX_NOTE + " Sosial sektor vərəqinin `wage_avg` sırası ilə eyni anlayış deyil "
                         "(DVX əmək müqavilələri bazasıdır), ona görə ayrıca kodla saxlanılır.")
    add("employees_nonoil_dvx", "Qeyri neft-qaz sektorunda muzdla işləyənlərin sayı (DVX)",
        "Employees, non-oil-gas (tax service)", "min nəfər", "FR8",
        [W(DVX, 109, "Qeyri-neft qaz sektorunda muzdla işləyənlərin sayı (orta aylıq)")], DVX,
        note=_DVX_NOTE)
    add("wage_avg_nonoil_dvx", "Qeyri neft-qaz sektorunda orta aylıq əmək haqqı (DVX)",
        "Average monthly wage, non-oil-gas (tax service)", "AZN", "FR8",
        [W(DVX, 110, "Qeyri-neft qaz sektorunda orta aylıq əmək haqqı")], DVX, note=_DVX_NOTE)

    # ---------------- FR2/FR3/FR8: sahə vərəqlərinin əmək bloku (v3-P0-13)
    # Dörd sənaye vərəqi (Mədənçıxarma, Emal Sənayesi, Su təchizatı, Elektrik enerjisi)
    # öz sahəsi üzrə orta aylıq əmək haqqını və işçi sayını 2016-cı ildən verir; bunlar
    # sahə əmək məhsuldarlığının və əmək haqqı kanalının yeganə uzun sıralarıdır.
    _BRANCH = [("mining", "Mədənçıxarma", "Mədənçıxarma sənayesi", "Mining and quarrying", 6, 7),
               ("manuf", "Emal Sənayesi", "Emal sənayesi", "Manufacturing", 7, 8),
               ("water", "Su təchizatı", "Su təchizatı", "Water supply", 7, 8),
               ("power", "Elektrik enerjisi", "Elektrik enerjisi, qaz və buxar istehsalı",
                "Electricity, gas and steam", 7, 8)]
    for root, sh, nm_az, nm_en, w_row, e_row in _BRANCH:
        add(f"wage_{root}", f"{nm_az} — orta aylıq əmək haqqı",
            f"{nm_en} — average monthly wage", "AZN", "FR8",
            [W(sh, w_row, "muzdlu işçilərin orta aylıq əmək haqları")], sh,
            note=f"{sh} vərəqi, sətir {w_row}; 2016–2025.")
        add(f"emp_{root}", f"{nm_az} — muzdlu işçilərin orta siyahı sayı",
            f"{nm_en} — average number of employees", "nəfər", "FR8",
            [W(sh, e_row, "muzdlu işçilərin orta siyahı sayı")], sh,
            note=f"{sh} vərəqi, sətir {e_row}; 2016–2025.")

    # Neft-qaz hasilatı üzrə buraxılış və əlavə dəyər (Mədənçıxarma vərəqi) — neft-qaz
    # blokunun əlavə dəyər/buraxılış nisbətinin (ρ) yeganə müşahidə olunan mənbəyi.
    add("go_oilgas_extraction", "Xam neft və təbii qaz hasilatı üzrə buraxılış",
        "Crude oil and natural gas extraction — gross output", "mln AZN", "FR3",
        [W("Mədənçıxarma", 8, "Xam neft və təbii qaz hasilatı üzrə buraxılış")], "Mədənçıxarma",
        note="Mədənçıxarma vərəqi, sətir 8; 2016–2025, LAKİN 2018 xanası mətn kimi yazılıb "
             "('31.417,2') və ədədə çevrilmir — həmin il boşdur (DATA-GAP, düzəliş sorğuya daxildir).")
    add("va_oilgas_extraction", "Xam neft və təbii qaz hasilatı üzrə əlavə dəyər",
        "Crude oil and natural gas extraction — value added", "mln AZN", "FR3",
        [W("Mədənçıxarma", 9, "Xam neft və təbii qaz hasilatı üzrə əlavə dəyər")], "Mədənçıxarma",
        note="Mədənçıxarma vərəqi, sətir 9; 2016–2025.")
    add("va_mining_branch", "Mədənçıxarma sənayesi üzrə əlavə dəyər (sahə vərəqi)",
        "Mining and quarrying — value added (branch sheet)", "mln AZN", "FR3",
        [W("Mədənçıxarma", 4, "Mədənçıxarma sənayesi üzrə əlavə dəyər")], "Mədənçıxarma",
        note="Mədənçıxarma vərəqi, sətir 4; 2016–2025. Real sektor vərəqindən gələn `va_mining` "
             "ilə eyni anlayışdır, lakin ayrı buraxılışdır — çarpaz yoxlama üçün saxlanılır, "
             "nəşr olunan sahə səviyyələri `va_mining` sırasındandır.")

    # ---------------- FR6: dövlət əsaslı vəsait qoyuluşu (DİP, v3-P0-13) -
    # `DİP 2016-2026` vərəqi iki pilləli başlıqlıdır: 2026 sütunu "Nəzərdə tutulmuş vəsait"
    # (plan) və "Faktiki xərc (01.04.2026)" oxunuşlarına bölünür. Faktiki sıra 2016–2025-i
    # əhatə edir; 2026 planı `dip_r004_plan` kodu ilə iş kitabı registrindədir və proqnoz ili
    # olduğu üçün faktiki obyektə (D1) düşmür.
    DIP = "DİP 2016-2026"
    add("state_capex_dip", "Dövlət əsaslı vəsait qoyuluşunun cəmi",
        "State capital investment, total", "mln AZN", "FR6",
        [W(DIP, 4, "Dövlət əsaslı vəsait qoyuluşunun cəmi")], DIP,
        note="DİP vərəqi, sətir 4; faktiki 2016–2025. 2026 üçün nəzərdə tutulmuş vəsait "
             "(2 700,0 mln AZN) `dip_r004_plan`, 01.04.2026 tarixinə faktiki icra "
             "`dip_r004_fakt` kodları ilə registrdədir. Fiskal blokun `fiscal_r019` sırası ilə "
             "üst-üstə düşən 8 ildən 6-sında fərq 0,1 %-dən azdır (2025: 2 305,1 hər iki "
             "mənbədə), LAKİN 2021-ci ildə −46,4 % (2 510,6 vs 4 683,9) və 2024-cü ildə "
             "−13,7 % (2 741,6 vs 3 176,4) fərq var. Ona görə `fiscal_r019` sırasındakı "
             "2022–2023 boşluğu bu sıradan AVTOMATİK doldurulmur — iki sıranın əhatə fərqi "
             "aydınlaşdırılmalıdır (məlumat sorğusuna daxildir).")
    add("state_capex_infra_dip", "Dövlət əsaslı vəsait qoyuluşu — infrastruktur layihələri",
        "State capital investment — infrastructure projects", "mln AZN", "FR6",
        [W(DIP, 5, "o cümlədən: İnfrastruktur layihələri")], DIP,
        note="DİP vərəqi, sətir 5; faktiki 2016–2025.")
    add("state_capex_external_dip",
        "Xarici kreditor qurumlarla birgə maliyyələşmə üzrə kreditlərin məbləği",
        "External co-financing credits", "mln AZN", "FR6",
        [W(DIP, 9, "Xarici kreditor qurumlarla birgə maliyyələşmə əsasında həyata keçirilən "
                   "layihələr üzrə kreditlərin məbləği")], DIP,
        note="DİP vərəqi, sətir 9; faktiki 2020–2025, 2026 planı `dip_r009_plan` kodundadır.")

    # ---------------- FR9: inflyasiya ------------------------------------
    add("cpi_infl", "İnflyasiya (İQİ, orta illik)", "CPI inflation, annual average", "%", "FR9",
        [C(P_MACRO, "cpi_infl"),
         V(V57, "İnflyasiya (İQİ, orta illik)", row=54, block_prefix="Orta illik inflyasiya")],
        V57, official="cpi_infl")
    add("cpi_index", "İstehlak qiymətlərinin bazis indeksi", "CPI, base index", "indeks", "FR9",
        [C(P_MODEL, "cpi_index")], P_MODEL)

    # ---------------- FR10: valyuta məzənnəsi ----------------------------
    # DATA-GAP: statutar şablonun məzənnə sətirləri tam #REF! qaytarır, iş
    # kitabında yalnız dövrün sonuna AZN/USD var, AZN/EUR ümumiyyətlə yoxdur.
    add("fx_usd_azn", "ABŞ dollarının orta illik məzənnəsi", "AZN/USD, annual average",
        "AZN/USD", "FR10",
        [A("Monetar sektoru", 114, "Valyuta məzənnəsi (dövrün sonuna)")], "Monetar sektoru",
        note="DATA-GAP: şablonun məzənnə sətri tam #REF!; orta illik məzənnə iş "
             "kitabının aylıq (dövrün sonuna) müşahidələrinin ortalamasıdır")
    add("fx_usd_azn_eop", "ABŞ dollarının məzənnəsi (dövrün sonuna)", "AZN/USD, end of period",
        "manat/US$", "FR10", [W("Monetar sektoru", 114, "Valyuta məzənnəsi (dövrün sonuna)")],
        "Monetar sektoru")

    # ---------------- FR7/FR11: xarici ticarət ---------------------------
    # Tədiyyə Balansı vərəqi statutar şablonla eyni anlayışı (BPM6, FOB) verir,
    # lakin 2010–2025 dövrünü əhatə edir; ona görə əsas mənbə odur.
    TB = "Tədiyyə Balansı"
    add("exp_goods", "Malların ixracı (FOB)", "Exports of goods (FOB)", "mln USD", "FR7/FR11",
        [W(TB, 6, "Mallar üzrə ixrac (FOB)"), V(V1314, "Malların ixracı", row=4)], V1314)
    add("exp_goods_oil", "Malların ixracı — neft-qaz sektoru", "Exports of goods, oil and gas",
        "mln USD", "FR7/FR11",
        [W(TB, 7, "Neft-qaz sektoru üzrə ixrac"), V(V1314, "Neft və qaz", row=5)], V1314)
    add("exp_goods_nonoil", "Malların ixracı — qeyri neft-qaz sektoru",
        "Exports of goods, non-oil", "mln USD", "FR7/FR11",
        [W(TB, 8, "Qeyri-neft-qaz sektoru üzrə ixrac"), V(V1314, "Digər sektorlar", row=6)], V1314)
    add("imp_goods", "Malların idxalı (FOB)", "Imports of goods (FOB)", "mln USD", "FR7/FR11",
        [W(TB, 9, "Mallar üzrə idxal (FOB)", tr="neg"), V(V1314, "Malların idxalı", row=7)], V1314)
    add("imp_goods_oil", "Malların idxalı — neft-qaz sektoru", "Imports of goods, oil and gas",
        "mln USD", "FR7/FR11",
        [W(TB, 10, "Neft-qaz sektoru üzrə idxal", tr="neg"), V(V1314, "Neft və qaz", row=8)], V1314,
        note="tədiyə balansı oxunuşu: neft-qaz sektorunun APARDIĞI idxal (BPM6, Tədiyyə Balansı "
             "sətir 10). Neft və qaz MƏHSULLARININ idxalı ayrıca sıradır — `imp_oilgas_products` "
             "(Ticarət, sətir 13). İki anlayış eyni deyil və bir-birini əvəz etmir.")
    add("imp_goods_nonoil", "Malların idxalı — qeyri neft-qaz sektoru",
        "Imports of goods, non-oil", "mln USD", "FR7/FR11",
        [W(TB, 11, "Qeyri-neft-qaz sektoru üzrə idxal", tr="neg"), V(V1314, "Digər sektorlar", row=9)], V1314)
    add("exp_services", "Xidmətlərin ixracı", "Exports of services", "mln USD", "FR7/FR11",
        [W(TB, 14, "Xidmətlər üzrə ixrac (daxiloma)"),
         V(V1314, "Xidmətlərin ixracı", row=10)], V1314)
    add("imp_services", "Xidmətlərin idxalı", "Imports of services", "mln USD", "FR7/FR11",
        [W(TB, 34, "Xidmətlər üzrə idxal (ödənişlər)", tr="neg"),
         V(V1314, "Xidmətlərin idxalı", row=11)], V1314)
    # Neft-qaz MƏHSULLARININ idxalı (DGK dəyəri) — `imp_goods_oil` ilə eyni anlayış DEYİL:
    # bu sıra idxal olunan neft/qaz məhsullarının dəyəridir, o biri isə neft-qaz sektorunun
    # apardığı idxaldır (BPM6). Hər ikisi ayrıca nəşr olunur, seçim Nazirliyə buraxılır.
    add("imp_oilgas_products", "Neft/qaz məhsullarının idxalı (DGK dəyəri)",
        "Imports of oil and gas products (customs value)", "mln USD", "FR7/FR11",
        [W("Ticarət", 13, "Neft/qaz idxalı")], "Ticarət",
        note="Ticarət vərəqi, sətir 13; 2017–2025. Tədiyə balansındakı qarşılığı — neft-qaz "
             "sektorunun apardığı idxal — `imp_goods_oil` (Tədiyyə Balansı, sətir 10, 2010–2025).")

    # ---------------- FR12: tədiyə balansı -------------------------------
    add("current_account", "Cari əməliyyatlar hesabı", "Current account balance",
        "mln USD", "FR12",
        [W(TB, 3, "Cari əməliyyatlar hesabı"),
         V(V1314, "CARİ ƏMƏLİYYATLAR HESABI", row=19)], V1314)
    add("trade_balance", "Xarici ticarət balansı", "Trade balance", "mln USD", "FR12",
        [W(TB, 5, "Xarici ticarət balansı"),
         V(V1314, "XARİCİ TİCARƏT BALANSI", row=20)], V1314)
    add("services_balance", "Xidmətlər balansı", "Services balance", "mln USD", "FR12",
        [W(TB, 13, "Xidmətlər balansı, xalis"),
         V(V1314, "XİDMƏTLƏR BALANSI", row=23)], V1314)
    add("primary_income", "İlkin gəlirlər, xalis", "Primary income, net", "mln USD", "FR12",
        [W(TB, 56, "İlkin gəlirlər, xalis"), V(V1314, "İLKİN GƏLİRLƏR", row=26)], V1314)
    add("secondary_income", "Təkrar gəlirlər, xalis", "Secondary income, net", "mln USD", "FR12",
        [W(TB, 78, "Təkrar gəlirlər, xalis"), V(V1314, "TƏKRAR GƏLİRLƏR", row=27)], V1314)
    add("nonoil_current_account", "Qeyri neft-qaz cari əməliyyatlar hesabı",
        "Non-oil current account", "mln USD", "FR12",
        [W(TB, 4, "Qeyri-neft-qaz cari əməliyyatlar hesabı")], TB)

    # ---------------- FR7/FR11: XİDMƏTLƏRİN BÖLGÜSÜ (xidmət bölgüsü) ----------------
    # Tədiyyə Balansı vərəqinin 15–54-cü sətirləri xidmət ixracını və idxalını həm
    # neft-qaz / qeyri-neft-qaz kəsiyində, həm də adlandırılmış xidmət növləri üzrə verir
    # (2010–2025). Debet sətirləri (idxal/ödəniş) vərəqdə MƏNFİ yazılır — `tr="neg"` onları
    # şablonun konvensiyasına (müsbət kəmiyyət) gətirir. Cəm eynilikləri FR07-də HƏR il
    # üçün yoxlanılır: r14 = r15 + r17, r34 = r35 + r38.
    _SERV = [
        # (kod, sətir, etiket, AZ ad, EN ad)
        ("exp_serv_oil", 15, "Neft-qaz sektoru üzrə xidmət ixracı",
         "Xidmət ixracı — neft-qaz sektoru", "Services exports — oil and gas sector"),
        ("exp_serv_nonoil", 17, "Qeyri-neft-qaz sektoru üzrə xidmət ixracı",
         "Xidmət ixracı — qeyri neft-qaz sektoru", "Services exports — non-oil sector"),
        ("exp_serv_transport", 18, "Nəqliyyat xidmətləri",
         "Xidmət ixracı — nəqliyyat", "Services exports — transport"),
        ("exp_serv_travel", 22, "Turizm xidmətləri",
         "Xidmət ixracı — turizm", "Services exports — travel"),
        ("exp_serv_telecom", 28, "Telekomunikasiya (rabitə) xidmətləri",
         "Xidmət ixracı — telekomunikasiya", "Services exports — telecommunications"),
        ("exp_serv_business", 32, "Digər işgüzar xidmətlər",
         "Xidmət ixracı — digər işgüzar xidmətlər", "Services exports — other business services"),
    ]
    _SERV_M = [
        ("imp_serv_oil", 35, "Neft-qaz sektoru üzrə xidmət idxalı",
         "Xidmət idxalı — neft-qaz sektoru", "Services imports — oil and gas sector"),
        ("imp_serv_nonoil", 38, "Qeyri-neft-qaz sektoru üzrə xidmət idxalı",
         "Xidmət idxalı — qeyri neft-qaz sektoru", "Services imports — non-oil sector"),
        ("imp_serv_transport", 39, "Nəqliyyat xidmətləri",
         "Xidmət idxalı — nəqliyyat", "Services imports — transport"),
        ("imp_serv_travel", 43, "Turizm xidmətləri",
         "Xidmət idxalı — turizm", "Services imports — travel"),
        ("imp_serv_telecom", 49, "Telekomunikasiya (rabitə) xidmətləri",
         "Xidmət idxalı — telekomunikasiya", "Services imports — telecommunications"),
        ("imp_serv_business", 53, "Digər işgüzar xidmətlər",
         "Xidmət idxalı — digər işgüzar xidmətlər", "Services imports — other business services"),
    ]
    for code, row, label, nm_az, nm_en in _SERV:
        add(code, nm_az, nm_en, "mln USD", "FR7/FR11", [W(TB, row, label)], TB,
            note=f"Tədiyyə Balansı, sətir {row} (daxilolma, 2010–2025)")
    for code, row, label, nm_az, nm_en in _SERV_M:
        add(code, nm_az, nm_en, "mln USD", "FR7/FR11", [W(TB, row, label, tr="neg")], TB,
            note=f"Tədiyyə Balansı, sətir {row} (ödəniş; vərəqdə debet=mənfi, burada müsbət)")

    # ---------------- FR11/FR12: BRÜT GƏLİR AXINLARI (brüt gəlir axınları) ----------------
    # İlkin gəlirlər vərəqdə xalis (r56) DEYİL, həm də BRÜT daxilolma (r59) və ödəniş (r68)
    # kimi, hər ikisi neft-qaz / qeyri-neft-qaz və alət (birbaşa / portfel / digər) kəsiyində
    # verilir. Alət sətirlərinin etiketləri TƏKRARLANIR ("birbaşa investisiyalardan gəlirlər"
    # dörd dəfə) — etiket yoxlaması (sətir, etiket) cütü üzərində aparıldığı üçün bu problem
    # yaratmır, lakin kodların adı sətri birmənalı göstərir.
    _INC = [
        ("pi_credit", 59, "İlkin gəlirlər üzrə daxilolmalar", 0,
         "İlkin gəlirlər — daxilolmalar (brüt)", "Primary income — credits (gross)"),
        ("pi_credit_oil", 60, "İlkin gəlirlərdə neft-qaz sektoru üzrə daxilolmalar", 0,
         "İlkin gəlir daxilolmaları — neft-qaz", "Primary income credits — oil and gas"),
        ("pi_credit_oil_di", 61, "birbaşa investisiyalardan gəlirlər", 0,
         "İlkin gəlir daxilolmaları — neft-qaz, birbaşa investisiya",
         "Primary income credits — oil and gas, direct investment"),
        ("pi_credit_oil_pf", 62, "portfel investisiyalardan gəlirlər", 0,
         "İlkin gəlir daxilolmaları — neft-qaz, portfel",
         "Primary income credits — oil and gas, portfolio"),
        ("pi_credit_oil_oth", 63, "digər gəlirlər", 0,
         "İlkin gəlir daxilolmaları — neft-qaz, digər",
         "Primary income credits — oil and gas, other"),
        ("pi_credit_nonoil", 64, "İlkin gəlirlərdə qeyri-neft-qaz sektoru üzrə daxilolmalar", 0,
         "İlkin gəlir daxilolmaları — qeyri neft-qaz", "Primary income credits — non-oil"),
        ("pi_credit_nonoil_di", 65, "birbaşa investisiyalardan gəlirlər", 0,
         "İlkin gəlir daxilolmaları — qeyri neft-qaz, birbaşa investisiya",
         "Primary income credits — non-oil, direct investment"),
        ("pi_credit_nonoil_pf", 66, "portfel investisiyalardan gəlirlər", 0,
         "İlkin gəlir daxilolmaları — qeyri neft-qaz, portfel",
         "Primary income credits — non-oil, portfolio"),
        ("pi_credit_nonoil_oth", 67, "digər gəlirlər", 0,
         "İlkin gəlir daxilolmaları — qeyri neft-qaz, digər",
         "Primary income credits — non-oil, other"),
        ("pi_debit", 68, "İlkin gəlirlər üzrə ödənişlər", 1,
         "İlkin gəlirlər — ödənişlər (brüt)", "Primary income — debits (gross)"),
        ("pi_debit_oil", 69, "İlkin gəlirlərdə neft-qaz sektoru üzrə ödənişlər", 1,
         "İlkin gəlir ödənişləri — neft-qaz", "Primary income debits — oil and gas"),
        ("pi_debit_oil_di", 70, "birbaşa investisiyalardan gəlirlər", 1,
         "İlkin gəlir ödənişləri — neft-qaz, birbaşa investisiya gəlirləri",
         "Primary income debits — oil and gas, direct investment income"),
        ("pi_debit_oil_pf", 71, "portfel investisiyalardan gəlirlər", 1,
         "İlkin gəlir ödənişləri — neft-qaz, portfel",
         "Primary income debits — oil and gas, portfolio"),
        ("pi_debit_oil_oth", 72, "digər gəlirlər", 1,
         "İlkin gəlir ödənişləri — neft-qaz, digər",
         "Primary income debits — oil and gas, other"),
        ("pi_debit_nonoil", 73, "İlkin gəlirlərdə qeyri-neft-qaz sektoru üzrə ödənişlər", 1,
         "İlkin gəlir ödənişləri — qeyri neft-qaz", "Primary income debits — non-oil"),
        ("pi_debit_nonoil_di", 74, "birbaşa investisiyalardan gəlirlər", 1,
         "İlkin gəlir ödənişləri — qeyri neft-qaz, birbaşa investisiya",
         "Primary income debits — non-oil, direct investment"),
        ("pi_debit_nonoil_pf", 75, "portfel investisiyalardan gəlirlər", 1,
         "İlkin gəlir ödənişləri — qeyri neft-qaz, portfel",
         "Primary income debits — non-oil, portfolio"),
        ("pi_debit_nonoil_oth", 76, "digər gəlirlər", 1,
         "İlkin gəlir ödənişləri — qeyri neft-qaz, digər",
         "Primary income debits — non-oil, other"),
        ("si_credit", 79, "Təkrar gəlirlər üzrə daxilolmalar", 0,
         "Təkrar gəlirlər — daxilolmalar (brüt)", "Secondary income — credits (gross)"),
        ("si_debit", 80, "Təkrar gəlirlər üzrə ödənişlər", 1,
         "Təkrar gəlirlər — ödənişlər (brüt)", "Secondary income — debits (gross)"),
        ("primary_income_oil", 57, "Neft-qaz", 0,
         "İlkin gəlirlər, xalis — neft-qaz", "Primary income, net — oil and gas"),
        ("primary_income_nonoil", 58, "Qeyri-neft-qaz", 0,
         "İlkin gəlirlər, xalis — qeyri neft-qaz", "Primary income, net — non-oil"),
    ]
    for code, row, label, neg, nm_az, nm_en in _INC:
        add(code, nm_az, nm_en, "mln USD", "FR11",
            [W(TB, row, label, tr=("neg" if neg else None))], TB,
            note=(f"Tədiyyə Balansı, sətir {row}"
                  + ("; vərəqdə debet=mənfi, burada müsbət kəmiyyət" if neg else "")))

    # ---------------- FR12: KAPİTAL VƏ MALİYYƏ HESABI ---------------------
    # Vərəqin 81–109-cu sətirləri. İŞARƏ KONVENSİYASI vərəqin özününküdür və DƏYİŞDİRİLMİR:
    # maliyyə hesabının sətirləri XALİS DAXİLOLMA kimi yazılır (öhdəliklərin artımı müsbət,
    # aktivlərin artımı mənfi). Bu, `bop_r081 = bop_r082 + bop_r083` və
    # `bop_r083 = r084 + r094 + r097 + r100 + r093` eyniliklərinin qapanmasını təmin edir
    # (FR12 hər il üçün yoxlayır). Ehtiyat aktivləri və balanslaşdırıcı maddə bu vərəqdə
    # YOXDUR — onlar FR12-də BPM6 bağlanışından törədilir (bax `moe_bop_closure()`).
    _FIN = [
        ("capital_finance_account", 81, "Kapital və Maliyyənin hərəkəti hesabı",
         "Kapital və maliyyənin hərəkəti hesabı", "Capital and financial account"),
        ("capital_account", 82, "Kapital hesabı", "Kapital hesabı", "Capital account"),
        ("financial_account", 83, "Maliyyə hesabı", "Maliyyə hesabı", "Financial account"),
        ("fdi_net", 84, "Birbaşa investisiyalar, xalis",
         "Birbaşa investisiyalar, xalis", "Direct investment, net"),
        ("fdi_assets", 85, "Aktivlər (Xarici iqtisadiyyata)",
         "Birbaşa investisiyalar — aktivlər (xarici iqtisadiyyata)",
         "Direct investment — assets (abroad)"),
        ("fdi_assets_oil", 86, "Neft-qaz",
         "Birbaşa investisiya aktivləri — neft-qaz", "Direct investment assets — oil and gas"),
        ("fdi_assets_nonoil", 87, "Qeyri-neft-qaz",
         "Birbaşa investisiya aktivləri — qeyri neft-qaz", "Direct investment assets — non-oil"),
        ("fdi_liabilities", 88, "Öhdəliklər",
         "Birbaşa investisiyalar — öhdəliklər", "Direct investment — liabilities"),
        ("fdi_inward", 89, "Azərbaycan iqtisadiyyatına BXİ",
         "Azərbaycan iqtisadiyyatına birbaşa xarici investisiyalar", "Inward FDI"),
        ("fdi_inward_oil", 90, "Neft-qaz",
         "Azərbaycan iqtisadiyyatına BXİ — neft-qaz", "Inward FDI — oil and gas"),
        ("fdi_inward_nonoil", 91, "Qeyri-neft-qaz",
         "Azərbaycan iqtisadiyyatına BXİ — qeyri neft-qaz", "Inward FDI — non-oil"),
        ("fdi_repatriation", 92, "İnvestisiyaların repatriasiyası",
         "İnvestisiyaların repatriasiyası (maliyyə hesabı)",
         "Repatriation of investment (financial account)"),
        ("oil_bonus", 93, "Neft bonusu", "Neft bonusu", "Oil bonus"),
        ("portfolio_net", 94, "Portfel investisiyaları, xalis",
         "Portfel investisiyaları, xalis", "Portfolio investment, net"),
        ("portfolio_assets", 95, "Aktivlər",
         "Portfel investisiyaları — aktivlər", "Portfolio investment — assets"),
        ("portfolio_liabilities", 96, "Öhdəliklər",
         "Portfel investisiyaları — öhdəliklər", "Portfolio investment — liabilities"),
        ("fin_derivatives_net", 97, "Törəmə maliyyə alətləri",
         "Törəmə maliyyə alətləri, xalis", "Financial derivatives, net"),
        ("other_inv_net", 100, "Digər investisiyalar, xalis",
         "Digər investisiyalar, xalis", "Other investment, net"),
        ("other_inv_assets", 101, "Aktivlər (daxilolma)",
         "Digər investisiyalar — aktivlər", "Other investment — assets"),
        ("other_inv_liabilities", 105, "Öhdəliklər (ödəniş)",
         "Digər investisiyalar — öhdəliklər", "Other investment — liabilities"),
    ]
    for code, row, label, nm_az, nm_en in _FIN:
        add(code, nm_az, nm_en, "mln USD", "FR12", [W(TB, row, label)], TB,
            note=(f"Tədiyyə Balansı, sətir {row}; işarə vərəqin öz konvensiyasındadır "
                  "(maliyyə hesabı = xalis daxilolma)"))

    # ---------------- Xarici blok (ekzogen girişlər) ---------------------
    add("partner_gdp_realg", "Tərəfdaş ölkələrin ÜDM artımı", "Partner country GDP growth",
        "%", "FR7", [C(P_EXTERNAL, "partner_gdp_realg")], P_EXTERNAL,
        note="ekzogen fərziyyə — IMF WEO")
    add("import_price_infl", "İdxal qiymətlərinin artımı", "Import price inflation", "%", "FR9",
        [C(P_EXTERNAL, "import_price_infl")], P_EXTERNAL, note="ekzogen fərziyyə")

    # ================= Rəsmi açıq mənbələr (§14) və gömrük bülletenləri (§15) =========
    # `DATA_GAP_NEGATIVE_PROOF.md` sənədinin xana səviyyəsində yoxlanmış yerləri. Bu sıralar
    # mövcud kodları ƏVƏZ ETMİR — yanına AYRI kodla əlavə olunur ki, heç bir hazır rəqəm
    # dəyişməsin; müqayisə və gələcək istifadə üçün açıqdır.
    RA = "AMB 2.16 (rəsmi orta məzənnə)"
    add("fx_usd_azn_avg", "ABŞ dollarının rəsmi orta illik məzənnəsi (AMB 2.16)",
        "US dollar official average annual exchange rate (CBAR 2.16)", "AZN/USD", "FR10",
        [PB("fx_usd_azn_avg", "US dollar", "cbar216")], RA,
        note="AMB cədvəl 2.16, ABŞ dolları sütunu, 1995–2025; DSK 010en-in implisit "
             "məzənnəsi ilə müstəqil tutuşdurulur (`fx_cross_check`). FR10-un BAZA sırası — "
             "spesifikasiyanın tələb etdiyi orta illik anlayış. Mövcud `fx_usd_azn` "
             "(aylıq dövrün-sonuna ortalaması, 2021–2025) SAXLANILIR və dəyişdirilmir")
    add("fx_eur_azn_avg", "Avronun rəsmi orta illik məzənnəsi (AMB 2.16)",
        "Euro official average annual exchange rate (CBAR 2.16)", "AZN/EUR", "FR10",
        [PB("fx_eur_azn", "EURO", "cbar216")], RA,
        note="AMB cədvəl 2.16, AVRO sütunu, 1999–2025; DSK 010en ilə 2025-ci il üçün "
             "dəqiq üst-üstə düşür (1,9210). FR10 bu sıranı `fx_eur_azn` kodu ilə nəşr "
             "edir — çarpaz hesablama və xarici ECB istinadı LƏĞV EDİLİB")

    SNA27 = "DSK 27 (027en)"
    add("gov_cons_indiv", "Dövlət idarələrinin fərdi xidmət xərcləri (P.3)",
        "Government expenditure on individual services (P.3)", "mln AZN", "FR1",
        [PB("gov_cons_indiv",
            "expenditures of government  institutions providing individual services",
            "ssc027")], SNA27,
        note="DSK 27, sətir 10 (P.3), 1993–2025; 2025 ilkin məlumatdır (`2025*`)")
    add("gov_cons_collect", "Dövlət idarələrinin kollektiv xidmət xərcləri (P.4)",
        "Government actual final consumption, collective services (P.4)", "mln AZN", "FR1",
        [PB("gov_cons_collect",
            "Actual final consumption expenditures of government institutions providing "
            "collective services", "ssc027")], SNA27,
        note="DSK 27, sətir 12 (P.4), 1993–2025")
    add("gov_cons_total", "Dövlət idarələrinin son istehlak xərcləri (cəmi)",
        "Government final consumption expenditure (total)", "mln AZN", "FR1",
        [PB("gov_cons_total",
            "expenditures of government institutions providing individual services + "
            "Actual final consumption expenditures of government institutions providing "
            "collective services", "ssc027")], SNA27,
        note="P.3 (fərdi) + P.4 (kollektiv), 1993–2025 — xərc tərəfinin müstəqil oxunuşu üçün milli hesablar "
             "əsaslı dövlət son istehlakı; interim büdcə törəməsini əvəz edən rəsmi sıra. "
             "FR01-in tələb tənliyində AYRICA sürücü kimi sınaqdan keçirilir (bu buraxılış)")
    add("hh_cons_actual", "Ev təsərrüfatlarının faktiki son istehlak xərcləri (P.4)",
        "Actual final consumption expenditure of households (P.4)", "mln AZN", "FR1",
        [PB("hh_cons_actual", "Actual final consumption expenditures of households",
            "ssc027")], SNA27,
        note="DSK 27, sətir 7 (P.4), 1993–2025 — dövlət sırası ilə eyni cədvəldən; iş "
             "kitabının `final_consumption` sırasından (ƏHALİNİN hesabı üzrə) FƏRQLİ "
             "anlayışdır, ona görə ayrıca kodla nəşr olunur")

    # --- ÜDM-in TAM İSTİFADƏ (xərc) hesabı — DSK 27-nin qalan sətirləri ------------------
    # Bu on bir sıra bu buraxılışda ƏLAVƏ OLUNDU. Onlarsız paketdə yalnız istehsal hesabı və
    # davranış tənlikləri ilə qurulmuş «tələb nüvəsi» var idi, resurslar–istifadə (tələb-təklif)
    # balansı isə qurula bilmirdi: xüsusilə ehtiyatların dəyişməsi (P.52) DATA-GAP elan
    # edilmişdi. Sıra mənbədə 1993-cü ildən var — boşluq OXUNUŞDA idi, məlumatda deyil.
    # FR01b bu on bir sıranın üzərində resurslar–istifadə balansını qurur (bax `notebooks/FR01b`).
    add("cons_final_total", "Faktiki son istehlak xərcləri, cəmi (P.4)",
        "Actual final consumption expenditure, total (P.4)", "mln AZN", "FR1B",
        [PB("cons_final_total", "Actual final consumption expenditures", "ssc027")], SNA27,
        note="DSK 27, sətir 5 (P.4), 1993–2025 — istifadə hesabının birinci baş sətri; "
             "ev təsərrüfatlarının (P.4) və dövlətin kollektiv (P.4) faktiki son istehlakının cəmi")
    add("hh_cons_final", "Ev təsərrüfatlarının son istehlak xərcləri (P.3)",
        "Household final consumption expenditure (P.3)", "mln AZN", "FR1B",
        [PB("hh_cons_final", "final consumption expenditure of households", "ssc027")], SNA27,
        note="DSK 27, sətir 9 (P.3), 1993–2025 — ev təsərrüfatlarının ÖZ xərci; faktiki son "
             "istehlakdan (`hh_cons_actual`, P.4) fərqi dövlətin və QHT-lərin fərdi "
             "xidmətlərinin transferidir")
    add("npish_cons_final", "QHT-lərin son istehlak xərcləri (P.3)",
        "NPISH final consumption expenditure (P.3)", "mln AZN", "FR1B",
        [PB("npish_cons_final",
            "expenditures of non - profit institutions providing services for households",
            "ssc027")], SNA27,
        note="DSK 27, sətir 11 (P.3), 1993–2025 — ev təsərrüfatlarına xidmət göstərən "
             "qeyri-kommersiya təşkilatları")
    add("gcf_total", "Ümumi yığım (P.5)", "Gross capital formation (P.5)", "mln AZN", "FR1B",
        [PB("gcf_total", "Gross saving", "ssc027")], SNA27,
        note="DSK 27, sətir 13 (P.5), 1993–2025 = əsas kapitalın ümumi yığımı (P.51) + "
             "ehtiyatların dəyişməsi (P.52). Mənbənin İngilis etiketi «Gross saving»dir, "
             "lakin SNA kodu və cədvəlin öz cəmi bunun YIĞIM olduğunu göstərir")
    add("gfcf_total", "Əsas kapitalın ümumi yığımı (P.51)",
        "Gross fixed capital formation (P.51)", "mln AZN", "FR1B",
        [PB("gfcf_total", "gross fixed capital formation", "ssc027")], SNA27,
        note="DSK 27, sətir 14 (P.51), 1993–2025 — MİLLİ HESABLAR anlayışı; iş kitabının "
             "«əsas kapitala yönəldilmiş investisiyalar» sırası (FR06, `inv_total_nom`) ayrı "
             "statistik müşahidədir və eyni kəmiyyət DEYİL (2025-də nisbət 0,991)")
    add("inventories_chg", "Ehtiyatların dəyişməsi (P.52)",
        "Changes in inventories (P.52)", "mln AZN", "FR1B",
        [PB("inventories_chg", "changes in inventories (+, -)", "ssc027")], SNA27,
        note="DSK 27, sətir 15 (P.52), 1993–2025 — istifadə hesabının BUFER maddəsi və "
             "qısa dövrdə mal bazarının tarazlaşma kanalı. Əvvəlki buraxılışlarda bu sıra "
             "«heç bir mənbədə yoxdur» kimi qeyd olunurdu; qeyd SƏHV İDİ və ləğv edilib")
    add("net_exports_gs", "Mal və xidmətlərin xalis ixracı (P.6n)",
        "Net exports of goods and services (P.6n)", "mln AZN", "FR1B",
        [PB("net_exports_gs", "Net exports", "ssc027")], SNA27,
        note="DSK 27, sətir 16 (P.6n), 1993–2025 = ixrac (P.6) − idxal (P.7), manatla")
    add("exports_gs", "Mal və xidmətlərin ixracı (P.6)",
        "Exports of goods and services (P.6)", "mln AZN", "FR1B",
        [PB("exports_gs", "Exports", "ssc027")], SNA27,
        note="DSK 27, sətir 17 (P.6), 1993–2025 — manatla; FR07-nin dollarla ixracı ilə "
             "orta illik məzənnə (FR10) vasitəsilə tutuşdurulur")
    add("imports_gs", "Mal və xidmətlərin idxalı (P.7)",
        "Imports of goods and services (P.7)", "mln AZN", "FR1B",
        [PB("imports_gs", "Imports (-)", "ssc027")], SNA27,
        note="DSK 27, sətir 18 (P.7), 1993–2025 — cədvəldə MÜSBƏT yazılır, eynilikdə çıxılır")
    add("stat_discrepancy", "Statistik fərq (istifadə hesabı)",
        "Statistical discrepancy (use account)", "mln AZN", "FR1B",
        [PB("stat_discrepancy", "Statistical discrepancy", "ssc027")], SNA27,
        note="DSK 27, sətir 19, 1993–2025 — DSK-nın öz balanslaşdırma qalığı; 2013–2025 "
             "aralığında yalnız üç ildə (2014, 2018, 2019) sıfırdan fərqlidir")
    add("gdp_use", "ÜDM — istifadə (xərc) hesabı üzrə", "GDP by expenditure (use account)",
        "mln AZN", "FR1B", [PB("gdp_use", "GDP", "ssc027")], SNA27,
        note="DSK 27, sətir 20 (B.1*g), 1993–2025 — ÜDM-in İKİNCİ, müstəqil ölçüsü. "
             "İstehsal hesabının ÜDM-i ilə buraxılış fərqi ola bilər (2024: 187,2 mln AZN); "
             "fərq FR01b-də nəşr olunur, gizlədilmir")

    LAB = "DSK 4.5–4.8 (004_5-8en)"
    for suf, own_az, own_en in (("", "cəmi", "total"), ("_state", "dövlət", "state"),
                                ("_nonstate", "qeyri-dövlət", "non-state")):
        add(f"wage_ssc_total{suf}",
            f"İqtisadiyyat üzrə orta aylıq nominal əmək haqqı — {own_az} (DSK)",
            f"Average monthly nominal wage, economy — {own_en} (SSC)", "AZN", "FR8",
            [PB(f"wage_ssc_total{suf}", "On economy, total", "ssc0458")], LAB,
            note="DSK 4.5–4.8 `Dynamics`, 2005–2024; 19 fəaliyyət növü üzrə tam kəsişmə "
                 "`ssc_branch_wage_ownership()` ilə əlçatandır (ANNEX2 sətir 17(iii) "
                 "bağlanır)")

    ODZ = "opendata.az (investisiyalar, maliyyə mənbələri)"
    for code, az, en, note in (
        ("inv_own_funds", "Müəssisə və təşkilatların öz vəsaitləri üzrə investisiyalar",
         "Investment financed from enterprises' own funds",
         "opendata.az, 1999–2024; iş kitabının `real_r114` sırası (yalnız 2023–2025) ilə "
         "üst-üstə düşən illərdə DƏQİQ eynidir — SOE/öz vəsait tarixçəsini 1999-a qədər açır"),
        ("inv_budget_funds", "Büdcə vəsaitləri üzrə investisiyalar",
         "Investment financed from budget funds", "opendata.az, 1999–2024"),
        ("inv_state_budget_funds", "Dövlət büdcəsi vəsaitləri üzrə investisiyalar",
         "Investment financed from the state budget", "opendata.az, 1995–2024"),
    ):
        prov_label = {
            "inv_own_funds": "Müəssisə və təşkilatların öz vəsaitləri hesabına yönəldilmiş "
                             "investisiyalar (min manat)",
            "inv_budget_funds": "Büdcə vəsaitləri hesabına yönəldilmiş investisiyalar "
                                "(min manat)",
            "inv_state_budget_funds": "Dövlət büdcəsi vəsaitləri hesabına yönəldilmiş "
                                      "investisiyalar (min manat)",
        }[code]
        add(code, az, en, "mln AZN", "FR6", [PB(code, prov_label, "odazinv")], ODZ, note=note)

    SNA13 = "DSK 13 (013en)"
    _S13 = {r: [lt for lt, _l, rt in SSC013_SECTIONS if rt == r]
            for _lt, _lb, r in SSC013_SECTIONS}
    for root, letters in _S13.items():
        label = "; ".join(lb for lt, lb, rt in SSC013_SECTIONS if rt == root)
        name = dict((r, n) for _s, r, n, _e, _p, _vr, _vl, _gr, _gl in _SECTORS)[root]
        add(f"ic_ssc_{root}", f"{name} — aralıq istehlak (DSK 13)",
            f"{root} — intermediate consumption (SSC 13)", "mln AZN", "FR2/FR3",
            [PB(f"ic_ssc_{root}", label, "ssc013")], SNA13,
            note=f"DSK 13, NACE {'+'.join(letters)}, sütun P.2, 2005–2025 — statutar "
                 f"şablonun `ic_{root}` sırası 2024-də bitir, bu sıra 2025-i də əhatə edir "
                 "və `go − va` törəməsini əvəz edir")
    for m, mname_az, mname_en, mcode in (("go", "ümumi buraxılış", "gross output", "P.1"),
                                         ("ic", "aralıq istehlak",
                                          "intermediate consumption", "P.2"),
                                         ("va", "əlavə dəyər", "value added", "B.1g")):
        add(f"{m}_ssc_total", f"İqtisadiyyat üzrə {mname_az} (DSK 13)",
            f"Economy-wide {mname_en} (SSC 13)", "mln AZN", "FR2/FR3",
            [PB(f"{m}_ssc_total", "Total", "ssc013")], SNA13,
            note=f"DSK 13 `Total` sətri, sütun {mcode}, 2005–2025 (FISIM daxil)")

    BOPINC = "MOE BOP `Income account` / AMB NSDP"
    add("remit_credit", "Fiziki şəxslərin pul baratları — daxilolmalar",
        "Personal remittances — credit", "mln USD", "FR11",
        [PB("remit_credit", "Fiziki şəxslərin pul baratları", "moebop")], BOPINC,
        note="MOE BOP `Income account` sətir 74, YALNIZ faktiki sütunlar 1996–2024; "
             "AMB NSDP sırası (`remit_credit_cbar`) 2013–2024 üzrə eyni dəyərləri verir")
    add("remit_debit", "Fərdlərin pul köçürmələri — ödənişlər",
        "Personal remittances — debit", "mln USD", "FR11",
        [PB("remit_debit", "Fərdlərin pul köçürmələri", "moebop")], BOPINC,
        note="MOE BOP `Income account` sətir 83, YALNIZ faktiki sütunlar 1996–2024")
    add("remit_credit_cbar", "Pul köçürmələri — daxilolmalar (AMB NSDP)",
        "Personal remittances — credit (CBAR NSDP)", "mln USD", "FR11",
        [PB("remit_credit_cbar", "- Receipts", "nsdpbop")], BOPINC,
        note="AMB NSDP `BOP Analytical` sətir 48, rüblükdən illiyə, 2013–2025 — "
             "2025 ili üçün YEGANƏ mənbə (Nazirliyin sütunu «gözlənilən»dir)")
    add("remit_debit_cbar", "Pul köçürmələri — ödənişlər (AMB NSDP)",
        "Personal remittances — debit (CBAR NSDP)", "mln USD", "FR11",
        [PB("remit_debit_cbar", "- Receipts", "nsdpbop")], BOPINC,
        note="AMB NSDP `BOP Analytical`: ödəniş = daxilolma − xalis (sətir 48 − sətir 46); "
             "mənbənin `- Payments` sətri 2024-Q4-də eyniliyi pozur, bax provenans qeydi")

    DGK = "DGK gömrük statistikası bülleteni, Cədvəl 7"
    for (code, hs, flow, meas, _su, _f, unit, az, en) in CUSTOMS_SPEC:
        add(code, az, en, unit, "FR11",
            [PB(code, CUSTOMS_HS_LABELS[hs], "dgk_bulleten", panel="customs")], DGK,
            note=f"HS {hs}, {'ixrac' if flow == 'exp' else 'idxal'} "
                 f"{'miqdarı' if meas == 'qty' else 'statistik dəyəri'}, 2016–2025 — "
                 "gömrük bülletenlərindən (10 RAR arxivi) çıxarılıb"
                 + (f". {CUSTOMS_EXPORT_BREAK_NOTE}" if flow == "exp" else ""))
    for code, qcode, vcode, _mult, unit, az, en in CUSTOMS_UNIT_VALUES:
        hs = dict((c, h) for c, h, *_ in ((s[0], s[1]) + tuple(s[2:])
                                          for s in CUSTOMS_SPEC))[qcode]
        add(code, az, en, unit, "FR11",
            [PB(code, CUSTOMS_HS_LABELS[hs], "dgk_bulleten", panel="customs")], DGK,
            note=f"vahid qiymət = {vcode} ÷ {qcode}; HS {hs}, 2016–2025 — sübut sənədinin "
                 "(xi) bəndinin tələb etdiyi gömrük vahid qiymətləri")

    PRD = "DSK 18 / 1.46 / 1.47 (natural göstəricilər)"
    for (vrow, vlabel, vunit, code, key, _loc, label, sunit, factor,
         status) in VEREQ_DEAD_PRODUCTS:
        add(code, f"{vlabel} — natural istehsal (DSK)",
            f"{label} — output in kind (SSC)", (vunit if factor else sunit), "FR4",
            [PB(code, label, key)], PRD,
            note=f"statutar şablon 2.4.1.4. sətir {vrow} («{vlabel}») 2018–2030 üçün "
                 f"tam `#REF!`-dir; DSK qarşılığı: {status}")

    return E


# ==========================================================================
# 5. Mənbələrin həlli (etiket yoxlaması ilə)
# ==========================================================================

class _Resolver:
    """Kataloq mənbələrini keşdən oxuyur və hər müraciəti etiketlə yoxlayır."""

    def __init__(self, wb_reg, vereq_long, vereq_prod):
        self.wb_reg = wb_reg.set_index("series_code")
        # (vərəq, sətir, keyfiyyət) -> (seriya kodu, ETİKET). İki pilləli başlıqlı vərəqlərdə
        # bir sətir bir neçə kod verir (`..._plan`, `..._fakt`); açara keyfiyyət daxil edilir
        # ki, ƏSAS (keyfiyyətsiz) oxunuş variant tərəfindən üstələnməsin.
        self.wb_by_sheet_row = {}
        for code, r in wb_reg.set_index("series_code").iterrows():
            parts = str(code).split("_")
            qual = "_".join(parts[2:]) if len(parts) > 2 else ""
            label = str(r.name_az)
            if qual and label.endswith(f" — {qual}"):
                label = label[: -len(f" — {qual}")]
            self.wb_by_sheet_row[(r.sheet, int(r.row), qual)] = (code, label)
        self.vereq_long = vereq_long
        self.vereq_prod = vereq_prod
        self._wb_annual = None
        self._wb_monthly = None
        self._csv_cache = {}
        self._public_prov = None
        self._customs_prov = None

    # -- əsas iş kitabı ---------------------------------------------------
    @property
    def wb_annual(self):
        if self._wb_annual is None:
            self._wb_annual = pd.read_csv(C_WB_ANNUAL)
        return self._wb_annual

    def _from_wb(self, spec):
        sheet, row, label = spec["sheet"], int(spec["row"]), spec["label"]
        qual = spec.get("qual") or ""
        hit = self.wb_by_sheet_row.get((sheet, row, qual))
        if hit is None:
            raise KeyError(f"iş kitabı: '{sheet}' vərəqində {row} sətri"
                           + (f" ({qual} oxunuşu)" if qual else "") + " registrdə yoxdur")
        code, found = hit
        if _key(found) != _key(label):
            raise AssertionError(
                f"ETİKET UYĞUNSUZLUĞU — {sheet}!{row}: gözlənilən {label!r}, tapılan {found!r}")
        d = self.wb_annual
        s = d.loc[d["series_code"] == code].set_index("year")["value"].astype(float)
        tr = spec.get("tr")
        if tr == "idx100":
            s = s - 100.0
        elif tr == "neg":
            s = -s
        return s.sort_index(), f"wb:{sheet}!r{row}", int(row)

    # -- statutar şablon --------------------------------------------------
    def _find_vereq_row(self, spec):
        p = self.vereq_prod
        cand = p[p["sheet"] == spec["sheet"]]
        if spec.get("row") is not None:
            cand = cand[cand["row"] == int(spec["row"])]
            if len(cand) != 1:
                raise KeyError(f"{spec['sheet']}!{spec['row']}: məlumat sətri tapılmadı")
            got = cand.iloc[0]
            if _key(got["name_az"]) != _key(spec["label"]):
                raise AssertionError(
                    f"ETİKET UYĞUNSUZLUĞU — {spec['sheet']}!{spec['row']}: "
                    f"gözlənilən {spec['label']!r}, tapılan {got['name_az']!r}")
        else:
            cand = cand[cand["name_az"].map(_key) == _key(spec["label"])]
            if spec.get("sec_no"):
                cand = cand[cand["section_no"] == spec["sec_no"]]
            if spec.get("block_prefix"):
                cand = cand[cand["block"].fillna("").map(_key).str.startswith(_key(spec["block_prefix"]))]
            if spec.get("unit"):
                cand = cand[cand["unit"].map(_key) == _key(spec["unit"])]
            if len(cand) != 1:
                raise KeyError(f"{spec['sheet']}: {spec['label']!r} etiketi "
                               f"{len(cand)} sətrə uyğun gəldi — dəqiqləşdirin "
                               f"(sətirlər: {list(cand['row'])})")
            got = cand.iloc[0]
        if spec.get("sec_prefix"):
            sec = _key(got["section"] or "")
            if not sec.startswith(_key(spec["sec_prefix"])):
                raise AssertionError(
                    f"BÖLMƏ UYĞUNSUZLUĞU — {spec['sheet']}!{got['row']}: "
                    f"gözlənilən bölmə {spec['sec_prefix']!r}, tapılan {got['section']!r}")
        return got

    def _from_vereq(self, spec, split):
        got = self._find_vereq_row(spec)
        d = self.vereq_long
        s = d[(d["series_id"] == got["series_id"]) & (d["split"] == split)]
        s = s.set_index("year")["value"].astype(float).sort_index()
        return s, f"vereq:{spec['sheet']}!r{int(got['row'])}", int(got["row"])

    # -- CSV panelləri ----------------------------------------------------
    def _from_csv(self, spec):
        path, col = spec["path"], spec["column"]
        if path not in self._csv_cache:
            self._csv_cache[path] = pd.read_csv(path)
        d = self._csv_cache[path]
        if col not in d.columns:
            raise KeyError(f"{os.path.basename(path)}: '{col}' sütunu yoxdur")
        s = d.set_index("year")[col].astype(float).dropna().sort_index()
        return s, f"csv:{os.path.basename(path)}:{col}", None

    @property
    def wb_monthly_all(self):
        if self._wb_monthly is None:
            self._wb_monthly = pd.read_csv(C_WB_MONTHLY)
        return self._wb_monthly

    def _from_wb_avg(self, spec):
        sheet, row, label = spec["sheet"], int(spec["row"]), spec["label"]
        hit = self.wb_by_sheet_row.get((sheet, row, spec.get("qual") or ""))
        if hit is None:
            raise KeyError(f"iş kitabı: '{sheet}' vərəqində {row} sətri registrdə yoxdur")
        code, found = hit
        if _key(found) != _key(label):
            raise AssertionError(
                f"ETİKET UYĞUNSUZLUĞU — {sheet}!{row}: gözlənilən {label!r}, tapılan {found!r}")
        d = self.wb_monthly_all
        d = d[d["series_code"] == code]
        g = d.groupby("year")["value"].agg(["count", "mean"])
        g = g[g["count"] >= spec.get("min_months", 6)]
        s = g["mean"].astype(float).sort_index()
        return s, f"wb-avg:{sheet}!r{row}", int(row)

    # -- dondurulmuş panellər (§14 açıq mənbələr, §15 gömrük) --------------
    @property
    def public_prov(self):
        if self._public_prov is None:
            self._public_prov = public_provenance().set_index("var")
        return self._public_prov

    @property
    def customs_prov(self):
        if self._customs_prov is None:
            self._customs_prov = customs_hs27_provenance().set_index("var")
        return self._customs_prov

    def _from_panel(self, spec):
        var, which = spec["var"], spec["panel"]
        prov = self.public_prov if which == "public" else self.customs_prov
        if var not in prov.index:
            raise KeyError(f"{which} paneli: `{var}` sırası provenans cədvəlində yoxdur")
        pr = prov.loc[var]
        if _key(pr["row_label"]) != _key(spec["label"]):
            raise AssertionError(
                f"ETİKET UYĞUNSUZLUĞU — {which}:{var}: gözlənilən {spec['label']!r}, "
                f"tapılan {str(pr['row_label'])!r}")
        if str(pr["source_key"]) != str(spec["source_key"]):
            raise AssertionError(
                f"MƏNBƏ UYĞUNSUZLUĞU — {which}:{var}: gözlənilən {spec['source_key']!r}, "
                f"tapılan {str(pr['source_key'])!r}")
        p = public_panel() if which == "public" else customs_hs27()
        d = p[p["var"] == var]
        s = pd.Series(d["value"].values, index=d["year"].astype(int).values,
                      dtype=float).sort_index()
        tag = (f"{which}:{pr['source_key']}!{pr['table']}!{pr['locator']}"
               if which == "public" else f"customs:{pr['hs_code']}!{pr['flow']}")
        return s, tag, None

    def series(self, src, split="hesabat"):
        kind, spec = src
        if kind == "wb":
            return self._from_wb(spec)
        if kind == "wbavg":
            return self._from_wb_avg(spec)
        if kind == "vereq":
            return self._from_vereq(spec, split)
        if kind == "csv":
            return self._from_csv(spec)
        if kind == "panel":
            return self._from_panel(spec)
        raise ValueError(f"naməlum mənbə tipi: {kind}")


def _official_2025():
    d = pd.read_csv(P_OFFICIAL_2025)
    return dict(zip(d["series"], d["value"].astype(float)))


def _write_starter_dict(dct, path=None):
    """`outputs/series_dictionary.csv`-i yalnız güvənli halda başlanğıc cədvəllə yazır.

    Təhvil faylı notebook-ların əlavə etdiyi sətirlərlə tamamlanır (§3 müqaviləsi:
    hər nəşr olunan sıra üçün bir sətir). Keşin yenidən qurulması bu tamamlanmış
    faylı kataloq-yalnız başlanğıc cədvələ qədər qısaltmamalıdır, ona görə yazı
    yalnız iki halda aparılır: fayl ümumiyyətlə yoxdursa, yaxud mövcud fayl
    başlanğıc cədvəldən az sətir saxlayırsa (yarımçıq qalmış nüsxə).
    Böyük mövcud fayl heç vaxt üstündən yazılmır.
    """
    path = path or O_SERIES_DICT
    if os.path.exists(path):
        try:
            have = len(pd.read_csv(path))
        except Exception:
            have = -1          # oxunmayan/zədələnmiş nüsxə → yenidən yazılır
        if have >= len(dct):
            return False
    dct.to_csv(path, index=False)
    return True


def _build_series_cache(verbose=True):
    """Kataloqu həll edir -> actuals_long / baseline_long / catalog / lüğət."""
    wb_reg = pd.read_csv(C_WB_REGISTRY)
    vereq_long = pd.read_csv(C_VEREQ_LONG)
    vereq_prod = pd.read_csv(C_VEREQ_PRODUCTS)
    rs = _Resolver(wb_reg, vereq_long, vereq_prod)
    official = _official_2025()

    spec = _catalog_spec()
    codes = [e["code"] for e in spec]
    assert len(codes) == len(set(codes)), "kataloqda təkrar seriya kodu var"

    act_rows, base_rows, cat_rows, vint_rows, splice_rows = [], [], [], [], []
    for e in spec:
        code = e["code"]
        # --- faktiki məlumat: birinci mənbə əsasdır, qalanları boşluq doldurur
        vals, prov, row_no, sheet_ref = {}, {}, None, e["ref"]
        for k, src in enumerate(e["src"]):
            try:
                s, tag, rw = rs.series(src, split="hesabat")
            except (KeyError, AssertionError) as ex:
                if k == 0:
                    raise
                if verbose:
                    print(f"    [qeyd] {code}: köməkçi mənbə oxunmadı — {ex}")
                continue
            if k == 0:
                row_no, first_tag = rw, tag
                if src[0] in ("wb", "wbavg", "vereq"):
                    src_sheet = src[1].get("sheet")
                elif src[0] == "panel":
                    # Dondurulmuş panel: "vərəq" sütununa mənbə faylının açarı yazılır
                    # (§14/§15 provenans cədvəlində tam yol və cədvəl adı saxlanılır).
                    src_sheet = src[1]["source_key"]
                else:
                    src_sheet = os.path.basename(src[1]["path"])
            is_rate = e["unit"] in ("%", "indeks")
            ov_n, ov_max, ov_year = 0, 0.0, None
            for y, x in s.items():
                y = int(y)
                if y > LAST_ACTUAL:
                    continue            # D1: faktiki obyektə proqnoz ili düşə bilməz
                if pd.isna(x):
                    continue
                if y not in vals:
                    vals[y] = float(x)
                    prov[y] = tag
                else:
                    # Birləşdirmə yoxlaması: mənbələr üst-üstə düşən illərdə
                    # nə qədər fərqlənir? Səviyyə sıçrayışı gizlədilmir.
                    # Faiz sıraları üçün mütləq (f.b.), səviyyələr üçün nisbi fərq.
                    ov_n += 1
                    dif = (abs(float(x) - vals[y]) if is_rate
                           else abs(float(x) - vals[y]) / max(abs(vals[y]), 1e-9))
                    if dif > ov_max:
                        ov_max, ov_year = dif, y
            if k > 0 and ov_n:
                splice_rows.append(dict(
                    series_code=code, unit=e["unit"],
                    primary=first_tag, secondary=tag, n_overlap=ov_n,
                    metric="f.b." if is_rate else "nisbi",
                    max_diff=round(ov_max, 6), max_diff_year=ov_year))
        # --- 2025 buraxılışının rəsmi göstərici ilə üstələnməsi (vahid 2025 bazası)
        if e["official"] and e["official"] in official:
            new = official[e["official"]]
            old = vals.get(LAST_ACTUAL)
            if old is None or abs(old - new) > 1e-9:
                vint_rows.append(dict(series_code=code, year=LAST_ACTUAL,
                                      value_used=new, value_prior=old,
                                      prior_source=prov.get(LAST_ACTUAL, ""),
                                      note="DSK-nın rəsmi 2025 nəticəsi ilə əvəzləndi"))
            vals[LAST_ACTUAL] = new
            prov[LAST_ACTUAL] = "official_2025_actuals.csv"
        for y in sorted(vals):
            act_rows.append(dict(series_code=code, year=y, value=vals[y], source=prov[y]))

        # --- Nazirliyin Proqnoz sütunları (yalnız müqayisə üçün)
        for src in e["src"]:
            if src[0] != "vereq":
                continue
            try:
                s, tag, _ = rs.series(src, split="proqnoz")
            except (KeyError, AssertionError):
                continue
            for y, x in s.items():
                y = int(y)
                if pd.isna(x):
                    continue
                if y == LAST_ACTUAL:
                    # Şablonun 2025 sütunu Nazirliyin proqnozudur, faktiki nəticə
                    # deyil: faktiki dəyərlə yanaşı qeyd olunur, lakin nə faktiki
                    # obyektə, nə də baza ssenarisinə daxil edilmir.
                    got = vals.get(LAST_ACTUAL)
                    vint_rows.append(dict(
                        series_code=code, year=LAST_ACTUAL,
                        value_used=got, value_prior=float(x),
                        prior_source="8 vərəq Proqnoz 2025",
                        note="şablonun 2025 sütunu proqnozdur; faktiki nəticə ayrıca mənbədəndir"))
                    continue
                if y <= LAST_ACTUAL:
                    continue
                # Şablonun öz Proqnoz sütunu — mənbə etiketi `template_sample`
                # ("şablon-nümunə ssenarisi (75 saylı qərar formatı üzrə)"). Bu, Nazirliyin
                # RƏSMİ proqnozu deyil; rəsmi rəqəmlər üçün `ministry_official` slotu boş saxlanılır.
                base_rows.append(dict(series_code=code, year=y, value=float(x),
                                      source="template_sample"))
            break

        yrs = sorted(vals)
        cat_rows.append(dict(series_code=code, name_az=e["name_az"], name_en=e["name_en"],
                             unit=e["unit"], fr=e["fr"],
                             sheet=src_sheet, row=row_no,
                             statutory_sheet_ref=_statutory_ref(sheet_ref),
                             price_basis=e["price_basis"],
                             year_min=yrs[0] if yrs else None,
                             year_max=yrs[-1] if yrs else None,
                             n_obs=len(yrs), note=e["note"]))

    act = pd.DataFrame(act_rows).drop_duplicates(["series_code", "year"])
    base = pd.DataFrame(base_rows).drop_duplicates(["series_code", "year"])
    cat = pd.DataFrame(cat_rows)
    vint = pd.DataFrame(vint_rows)
    spl = pd.DataFrame(splice_rows)
    spl.to_csv(C_SPLICE, index=False)

    # D1-in konstruksiya ilə təminatı
    assert act["year"].max() <= LAST_ACTUAL, "faktiki məlumat proqnoz üfüqünə sızıb"
    if len(base):
        assert base["year"].min() >= FIRST_FORECAST, "baza ssenarisinə faktiki il düşüb"

    act.sort_values(["series_code", "year"]).to_csv(C_ACTUALS, index=False)
    base.sort_values(["series_code", "year"]).to_csv(C_BASELINE, index=False)
    cat.to_csv(C_CATALOG, index=False)
    vint.to_csv(C_VINTAGE, index=False)
    # §3 müqaviləsi: series_dictionary.csv — sütunlar dəqiq həmin ardıcıllıqla
    # (`price_basis` v3-də əlavə olunub — D5: hər nəşr olunan sıra qiymət bazasını elan edir)
    dct = cat[["series_code", "name_az", "name_en", "unit", "fr", "statutory_sheet_ref",
               "price_basis"]]
    dct.to_csv(C_DICT, index=False)
    # Başlanğıc nüsxə (`data/cache/…_starter.csv`) hər dəfə şərtsiz yenilənir.
    # `outputs/series_dictionary.csv` isə YALNIZ mövcud olmadıqda və ya başlanğıc
    # nüsxədən az sətir saxladıqda yazılır: notebook-lar (FR01–FR12) bu fayla öz
    # sıralarını `outputs.write_series_dictionary` vasitəsilə əlavə edir və keşin
    # yenidən qurulması həmin sətirləri silməməlidir — əks halda tam lüğət
    # tək başına çağırılan `build_cache()` ilə başlanğıc cədvələ qədər qısalır.
    _write_starter_dict(dct)
    data_gaps().to_csv(C_GAPS, index=False)
    return act, base, cat, vint


_DATA_GAPS = [
    # BAĞLANDI (bu buraxılış) — sıra AMB-nin 2.16 cədvəlində mövcuddur; qeyd tarixçə üçün
    # saxlanılır, lakin artıq boşluq deyil.
    ("fx_eur_azn", "FR10",
     "BAĞLANDI. Əvvəllər manatın avroya nisbətdə məzənnəsinin nə iş kitabında, nə statutar "
     "şablonda olmadığı qeyd edilirdi (şablonun 31-ci sətri bütün illər üzrə #REF! "
     "qaytarır) — bu, YANLIŞ idi: sıra AMB-nin 2.16 cədvəlində 1999–2025 üçün rəsmi orta "
     "illik məzənnə kimi dərc olunur.",
     "TAM BAĞLANDI. `fx_eur_azn` sırası AMB cədvəl 2.16 (AVRO sütunu) əsasında məlumat "
     "qatına daxil edilib (1999–2025) və DSK 010en-in implisit məzənnəsi ilə müstəqil "
     "tutuşdurulub (2025: hər iki mənbədə 1,9210). FR10 ARTIQ çarpaz məzənnə hesablamır və "
     "xarici ECB istinadından istifadə etmir: faktiki sıra üzərində peg arifmetikası ilə "
     "təsadüfi gəzişmə yarışdırılır, yelpik zolağının sigması isə həmin sıranın peg "
     "dövrünün öz volatilliyindən qiymətləndirilir."),
    ("fx_usd_azn", "FR10",
     "BAĞLANDI. Əvvəllər orta illik AZN/USD məzənnəsinin yalnız 2021–2025 üçün hesablana "
     "bildiyi qeyd edilirdi (şablonun 30-cu sətri tam #REF!, iş kitabında yalnız dövrün "
     "sonuna sıra) — bu, YANLIŞ idi: AMB-nin 2.16 cədvəli rəsmi orta illik məzənnəni "
     "1995-ci ilə qədər verir. model_series_annual.csv panelindəki `ner_usd` sırası isə "
     "əsl məzənnə deyil (miqyaslanmış indeksdir, 2024 üçün 1,657 verir) və istifadə edilmir.",
     "`fx_usd_azn_avg` sırası AMB cədvəl 2.16 əsasında məlumat qatına daxil edilib "
     "(1995–2025) və FR10-un qiymətləndirmə, sınaq və nəşr bazasıdır (UÇOT DÜZƏLİŞİ: "
     "əvvəlki buraxılış dövrün-sonuna sırasını işlədirdi; iki anlayış 2015-ci ildə "
     "0,5333 AZN fərqlənir). Mövcud `fx_usd_azn` (aylıq dövrün-sonuna müşahidələrin "
     "ortalaması, 2021–2025) və `fx_usd_azn_eop` DƏYİŞDİRİLMƏDƏN saxlanılır və müqayisə "
     "üçün nəşr olunur."),
    ("gdp_pc_usd", "FR1",
     "Adambaşına ÜDM-in dollar ifadəsi (şablon, sətir 21) 2013–2024 üçün #REF!-dir.",
     "Manatla göstərici (`gdp_pc_azn`) və məzənnə əsasında FR1-də hesablanır."),
    ("ic_*", "FR2",
     "BAĞLANDI. Aralıq istehlakın yalnız statutar şablonda (2013–2024) verildiyi və 2025-ci "
     "ilin `go − va` eyniliyi ilə doldurulduğu qeyd edilirdi — bu, YANLIŞ idi: DSK-nın 13-cü "
     "cədvəli («Production and generation of income account») aralıq istehlakı (P.2) NACE "
     "bölmələri üzrə 2005–2025 dövrü üçün dərc edir; həmçinin 015_1en (sənaye, bölmə "
     "səviyyəsi) və 019en (buraxılışda pay) mövcuddur.",
     "NƏŞR OLUNAN `ic_<sahə>` sırası artıq DSK-nın ÖZ P.2 sütunudur (2005–2025); 2000–2004 "
     "illəri `go − va` törəməsi ilə splice edilir və splice ayağı FR02-də sətir-sətir "
     "göstərilir. DSK oxunuşu ilə törəmə 231 sahə-ilin 222-sində yuvarlaqlaşdırma "
     "səviyyəsində eynidir; fərqli 9 sahə-il vintaj fərqidir və gizlədilmir. Xam DSK "
     "sıraları `ic_ssc_*` (11 model sahəsi) və `ic_ssc_total` kodları ilə paralel saxlanılır. "
     "`go_*` və `va_*` DƏYİŞDİRİLMİR (iş kitabının öz buraxılışı) — yeni aralıq istehlak "
     "modelə yalnız τ = 1 − IC/GO vasitəsilə daxil olur, ÜDM səviyyəsinə toxunmadan."),
    ("oil_defl, gdp_pc_realg, income_realg, expend_realg, wage_avg_realg, employment_g", "FR1/FR8",
     "Bu artım və deflyator sıraları yalnız statutar şablonda hazır verilir "
     "(2013–2024); iş kitabında müvafiq sətir yoxdur.",
     "2025-ci il üçün səviyyə sıralarından və İQİ-dən FR1/FR8-də hesablanır."),
    ("pop_total", "FR1",
     "Dövrün sonuna əhali sayı iş kitabında yoxdur; şablonda 2013–2024 verilir.",
     "Orta illik sıra (`pop_avg`, 1991–2025) istifadə olunur, dövrün sonuna sıra "
     "yalnız 2024-ə qədər göstərilir."),
    ("2.4.1.4. məhsul kataloqu", "FR4",
     "Şablonun məhsul vərəqində 272 xana #REF! qaytarır; bir sətrin (179) ADI da "
     "#REF!-dir; 10 ad 23 sətirdə təkrarlanır.",
     "Sıralar (vərəq, sətir, ölçü vahidi) kompozit açarı ilə identifikasiya olunur; "
     "#REF! xanaları `vereq_ref_cells.csv` hesabatında sətir-sətir göstərilir və "
     "heç bir yerdə sıfır kimi oxunmur. Sətir 179 (adı özü #REF!) tamamilə istisna "
     "edilir."),
    ("2.4.1.4. sətir 48–58 — 11 qida məhsulu", "FR4",
     "QİSMƏN BAĞLANDI. 48–58-ci sətirlər yalnız 5 real il (2013–2017) daşıyırdı, "
     "qalan 13 xana #REF! idi; `n_eff<8` qaydası hər on bir sətri SABİT (2017 "
     "səviyyəsində düz xətt) yoluna yönləndirirdi.",
     "10 sətir DİRÇƏLDİLDİ: DSK-nın 18-ci cədvəli (`018en`, 1995–2024) və ət/süd üçün "
     "kənd təsərrüfatı cədvəlləri 1.46/1.47 (1985–2024) əsasında, şablonun 2017-ci il "
     "lövbərinə yenidən bazalaşdırma ilə (`vereq_product_concordance()` hər sətrin "
     "mənbəyini, etiketini, vahidini və əmsalını qaytarır); n_eff 5 → 24–35, ayrıca "
     "sınaqda 10-dan 10-u təsadüfi gəzişməni üstələyir. Sətir 50 (meyvə-tərəvəz "
     "şirələri) DİRÇƏLDİLMİR: şablon min dkl, DSK min ton ölçür və DSK sətri şirələri "
     "konservlərlə birlikdə sayır (ortaq beş ildə nisbət 2,41–24,57) — sətir SABİT "
     "qalır; tələb olunan konkordans deyil, ÖLÇÜ VAHİDİ uzlaşdırılmasıdır."),
    ("cust_imp_qty_* / cust_imp_uv_*", "FR7/FR12",
     "BAĞLANDI. Neft-qaz idxalının NATURAL həcmi və orta idxal qiymətinin heç bir "
     "mənbədə olmadığı, yalnız dəyər sıralarının (`bop_r010`, `trade_r013`) mövcud "
     "olduğu qeyd edilirdi — bu, YANLIŞ idi: DGK-nın «Gömrük statistikası bülleteni» "
     "arxivləri (`Bulleten_2016.rar` … `Bulleten_2025.rar`) Cədvəl 7-də HS mövqeyi "
     "üzrə həm miqdar, həm statistik dəyər verir.",
     "Arxivlər `bsdtar` (libarchive) ilə açılıb (heç bir proqram quraşdırılmadan) və "
     "HS 2709 / 2710 / 2711 mövqelərinin idxal və ixrac sətirləri `data/"
     "customs_hs27_annual.csv` panelinə çıxarılıb: 18 sıra × 2016–2025 (miqdar, "
     "statistik dəyər və törəmə vahid qiymət, hər iki axın üzrə). Provenans hər sıra "
     "üçün il bazasını (illik buraxılış / dörd rübün cəmi) göstərir. ÖLÇÜLMÜŞ ƏHATƏ "
     "QIRILMASI gizlədilmir: DGK-nın 2016-cı il bülletenində İXRAC sətirləri iş "
     "kitabından aşağıdır (xam neft 6 575,4 vs 10 692,8 mln USD), 2017-dən tam üst-üstə "
     "düşür; İDXAL tərəfində qırılma yoxdur. QALAN BOŞLUQ: bülletenlər 2016-dan "
     "başlayır (2013–2015 əhatə olunmur) və HS-4 səviyyəsindədir — 2710 daxilində "
     "avtobenzin/dizel ayrıca verilmir."),
    ("sahə əlavə dəyəri — buraxılış fərqi", "FR2/FR3",
     "İş kitabı (cari buraxılış) ilə statutar şablon (köhnə buraxılış) sahə üzrə "
     "real artım tempində 2024-cü il üçün 2,2 f.b.-a, 2015-ci il üçün 4,9 f.b.-a "
     "qədər fərqlənir.",
     "Faktiki sıralarda vahid buraxılış saxlanılır: əsas mənbə iş kitabıdır; fərqlər "
     "`splice_check.csv` faylında ölçülüb qeyd olunur."),
    ("ÜDM komponentlərinin cəmi", "FR1",
     "Şablonun 2016-cı il sütununda neft-qaz və qeyri neft-qaz ÜDM-in cəmi ümumi "
     "ÜDM-dən 32,7 mln AZN (5,4×10⁻⁴ nisbi) fərqlənir.",
     "Fərq gizlədilmir; eyniliyin bərpası mütənasib uzlaşdırma ilə FR1/validate.py "
     "mərhələsində aparılır."),
]


def data_gaps() -> pd.DataFrame:
    """Bilinən məlumat boşluqları və onların rəsmiləşdirilmiş həlli.

    Hesabatın "məhdudiyyətlər" bölməsi və FR notebook-larındakı DATA-GAP
    qeydləri bu cədvəldən gəlir.
    """
    return pd.DataFrame(_DATA_GAPS, columns=["series_code", "fr", "problem_az", "treatment_az"])


def _statutory_ref(ref):
    """Statutar istinad mətni: 8 vərəq nömrəsi və ya iş kitabı vərəqinin adı."""
    r = str(ref)
    if r in VEREQ_SHEETS:
        return f"8 vərəq, {r}"
    if r.endswith(".csv"):
        return os.path.basename(r)
    return r


# ==========================================================================
# 6. Keşin qurulması
# ==========================================================================

def build_cache(verbose=True, check_md5=True):
    """Bir dəfəlik: bütün mənbələri oxuyur və `data/cache/*.csv` yazır.

    Notebook-lar bundan sonra 113 MB-lıq faylı heç vaxt açmır.
    """
    if check_md5:
        got = _md5(XLSX_MAIN)
        if verbose:
            print(f"MD5 {os.path.basename(XLSX_MAIN)}: {got}")
        if got != XLSX_MAIN_MD5:
            raise ValueError(f"iş kitabının MD5-i uyğun gəlmir: {got} != {XLSX_MAIN_MD5}")

    if verbose:
        print("\n[1/3] Əsas iş kitabı oxunur (113 MB) ...")
    reg, ann, mon = _build_workbook_cache(verbose=verbose)

    if verbose:
        print("\n[2/3] Statutar şablon (8 vərəq) oxunur ...")
    lg, pr, rf, mt = _build_vereq_cache(verbose=verbose)

    if verbose:
        print("\n[3/3] Seriya kataloqu həll edilir ...")
    act, base, cat, vint = _build_series_cache(verbose=verbose)

    if verbose:
        _print_summary(reg, ann, mon, lg, pr, rf, act, base, cat, vint)
    return dict(registry=reg, wb_annual=ann, wb_monthly=mon, vereq_long=lg,
                vereq_products=pr, vereq_ref=rf, vereq_meta=mt, actuals=act,
                baseline=base, catalog=cat, vintage=vint)


def _print_summary(reg, ann, mon, lg, pr, rf, act, base, cat, vint):
    print("\n" + "=" * 72)
    print("KEŞ HAZIRDIR — əhatə xülasəsi")
    print("=" * 72)
    print(f"iş kitabı registri : {len(reg):5d} sıra, {reg['sheet'].nunique()} vərəq")
    print(f"illik uzun format  : {len(ann):5d} sətir, illər "
          f"{int(ann['year'].min())}–{int(ann['year'].max())}")
    print(f"aylıq uzun format  : {len(mon):5d} sətir")
    print(f"statutar şablon    : {len(pr):5d} sətir, {len(lg)} dəyər, "
          f"{len(rf)} #REF! xanası")
    print(f"kataloq            : {len(cat):5d} sıra")
    print(f"faktiki (≤{LAST_ACTUAL})   : {len(act):5d} sətir, "
          f"{act['series_code'].nunique()} sıra")
    print(f"Nazirlik bazası    : {len(base):5d} sətir, "
          f"{base['series_code'].nunique()} sıra ({FIRST_FORECAST}–{LAST_FORECAST})")
    print(f"buraxılış qeydləri : {len(vint):5d} sıra 2025 üçün rəsmi göstərici ilə əvəzləndi")


# ==========================================================================
# 7. İllik API
# ==========================================================================

_CACHE = {}


def _load(path, what):
    if what not in _CACHE:
        if not os.path.exists(path):
            raise FileNotFoundError(
                f"keş faylı yoxdur: {path}\nƏvvəlcə data_layer.build_cache() işlədin.")
        _CACHE[what] = pd.read_csv(path)
    return _CACHE[what]


def clear_cache():
    """Yaddaşdakı keşi boşaldır (yenidən qurulduqdan sonra)."""
    _CACHE.clear()


def catalog() -> pd.DataFrame:
    """Seriya lüğəti: series_code, name_az, name_en, unit, fr, sheet, row, ...

    `outputs/series_dictionary.csv` faylı üçün başlanğıc məlumat bu cədvəldir.
    """
    return _load(C_CATALOG, "catalog").copy()


def series_dictionary() -> pd.DataFrame:
    """§3 müqaviləsinə uyğun lüğət sütunları (`price_basis` daxil olmaqla — D5)."""
    c = catalog()
    return c[["series_code", "name_az", "name_en", "unit", "fr", "statutory_sheet_ref",
              "price_basis"]].copy()


def actuals(series_code: str) -> pd.Series:
    """Faktiki illik sıra. İndeks = il, HƏMİŞƏ ≤ LAST_ACTUAL (2025).

    Proqnoz üfüqünün dəyəri bu funksiyadan heç bir halda qayıtmır — bu, D1
    qaydasının konstruksiya ilə təminatıdır.
    """
    d = _load(C_ACTUALS, "actuals")
    s = d[d["series_code"] == series_code]
    if s.empty:
        raise KeyError(f"kataloqda '{series_code}' sırası yoxdur "
                       f"(mövcud kodlar üçün catalog() çağırın)")
    out = s.set_index("year")["value"].astype(float).sort_index()
    out = out[out.index <= LAST_ACTUAL]
    assert out.index.max() <= LAST_ACTUAL, f"{series_code}: nümunə {LAST_ACTUAL}-i keçir"
    out.name = series_code
    return out


def actuals_source(series_code: str) -> pd.Series:
    """Hər il üçün faktiki dəyərin mənbəyi (provenans)."""
    d = _load(C_ACTUALS, "actuals")
    s = d[d["series_code"] == series_code]
    return s.set_index("year")["source"].sort_index()


def ministry_baseline(series_code: str) -> pd.Series:
    """8 vərəq şablonunun 2026–2030 Proqnoz sütunları — YALNIZ müqayisə üçün.

    Bu dəyərlər heç bir reqressiyaya və ya fərziyyə faylına daxil olmur. Nəşrdə mənbə
    etiketi `template_sample`-dir: şablon-nümunə ssenarisi (75 saylı qərar formatı üzrə),
    Nazirliyin RƏSMİ proqnozu deyil (rəsmi rəqəmlər üçün `ministry_official` slotu boşdur).
    """
    d = _load(C_BASELINE, "baseline")
    s = d[d["series_code"] == series_code]
    if s.empty:
        return pd.Series(dtype=float, name=series_code)
    out = s.set_index("year")["value"].astype(float).sort_index()
    out = out[(out.index >= FIRST_FORECAST) & (out.index <= LAST_FORECAST)]
    out.name = series_code
    return out


def vintage_notes() -> pd.DataFrame:
    """2025-ci il üçün rəsmi göstərici ilə əvəzlənmiş şablon dəyərləri."""
    return _load(C_VINTAGE, "vintage").copy()


def ref_report() -> pd.DataFrame:
    """Statutar şablondakı #REF! xanalarının açıq hesabatı (susdurulmur)."""
    return _load(C_VEREQ_REF, "vereq_ref").copy()


def products(sheet: str = "2.4.1.4.") -> pd.DataFrame:
    """FR4 məhsul kataloqu: kompozit (vərəq, sətir, ölçü vahidi) identifikatoru.

    Ad açar kimi işlədilmir: şablonda 10 ad təkrarlanır.
    """
    p = _load(C_VEREQ_PRODUCTS, "vereq_products")
    out = p[p["sheet"] == sheet].copy()
    assert out["series_id"].is_unique, "kompozit identifikator təkrarlanır"
    return out


def _vereq_lookup(sheet, row_label, row=None, sec_no=None, sec_prefix=None,
                  block_prefix=None, unit=None):
    p = _load(C_VEREQ_PRODUCTS, "vereq_products")
    lg = _load(C_VEREQ_LONG, "vereq_long")
    rs = _Resolver(_load(C_WB_REGISTRY, "wb_registry"), lg, p)
    spec = dict(sheet=sheet, label=row_label, row=row, sec_no=sec_no,
                sec_prefix=sec_prefix, block_prefix=block_prefix, unit=unit)
    return rs._find_vereq_row(spec), lg


def vereq_actuals(sheet: str, row_label: str, row=None, sec_no=None, sec_prefix=None,
                  block_prefix=None, unit=None) -> pd.Series:
    """8 vərəq şablonunun Hesabat (≤2025) hissəsi, etiket yoxlaması ilə.

    Sətir nömrəsi verilə bilər, lakin həmin sətrin etiketi `row_label` ilə
    tam üst-üstə düşməlidir; əks halda AssertionError atılır.
    """
    got, lg = _vereq_lookup(sheet, row_label, row, sec_no, sec_prefix, block_prefix, unit)
    s = lg[(lg["series_id"] == got["series_id"]) & (lg["split"] == "hesabat")]
    out = s.set_index("year")["value"].astype(float).sort_index()
    out = out[out.index <= LAST_ACTUAL]
    out.name = got["series_id"]
    return out


def vereq_baseline(sheet: str, row_label: str, row=None, sec_no=None, sec_prefix=None,
                   block_prefix=None, unit=None) -> pd.Series:
    """8 vərəq şablonunun Proqnoz (2026–2030) hissəsi — yalnız müqayisə üçün."""
    got, lg = _vereq_lookup(sheet, row_label, row, sec_no, sec_prefix, block_prefix, unit)
    s = lg[(lg["series_id"] == got["series_id"]) & (lg["split"] == "proqnoz")]
    out = s.set_index("year")["value"].astype(float).sort_index()
    out = out[(out.index >= FIRST_FORECAST) & (out.index <= LAST_FORECAST)]
    out.name = got["series_id"]
    return out


def wb_annual(series_code: str) -> pd.Series:
    """Əsas iş kitabının xam illik sırası (registr kodu ilə)."""
    d = _load(C_WB_ANNUAL, "wb_annual")
    s = d[d["series_code"] == series_code]
    if s.empty:
        raise KeyError(f"iş kitabı registrində '{series_code}' yoxdur")
    return s.set_index("year")["value"].astype(float).sort_index()


def monthly_panel() -> pd.DataFrame:
    """Aylıq panel (İQİ, pul aqreqatları, faiz dərəcələri, brent) — CSV mənbəsi."""
    d = pd.read_csv(P_MONTHLY)
    return d


def external_block() -> pd.DataFrame:
    """Xarici blok (IMF/ECB göstəriciləri, NEER/REER) — ekzogen girişlər."""
    return pd.read_csv(P_EXTERNAL)


def wb_registry() -> pd.DataFrame:
    """Əsas iş kitabından çıxarılmış bütün sıraların registri."""
    return _load(C_WB_REGISTRY, "wb_registry").copy()


def monthly(series_code: str) -> pd.DataFrame:
    """De-kumulyasiya edilmiş aylıq/rüblük axın sırası."""
    d = _load(C_WB_MONTHLY, "wb_monthly")
    s = d[d["series_code"] == series_code].copy()
    if s.empty:
        raise KeyError(f"iş kitabı registrində '{series_code}' yoxdur")
    return s.sort_values(["year", "month"]).reset_index(drop=True)


# ==========================================================================
# 8. Neft-qaz bloku — həcm, qiymət və çəki konstruksiyaları
# ==========================================================================
# Blok üç dəftər tərəfindən işlədilir (FR05b — qurulma; FR01 — neft-qaz deflyatoru;
# FR02 — mədənçıxarma deflyatoru), ona görə konstruksiya BURADA, bir yerdə saxlanılır:
# eyni indeks üç dəftərdə üç dəfə qurulsaydı, onların bir-birindən sürüşməsi mümkün olardı.
# Hər sıra registrdən sətir nömrəsi VƏ etiket yoxlaması ilə götürülür (`wb_labelled`).

BBL_PER_TON = 7.4        # iş kitabının öz konvensiyası: ixrac dəyəri ÷ (həcm × 7,4) = USD/barel

# Daxili bazarın qiymətləri heç bir statistik nəşrdə verilmir. Aşağıdakı iki sabit Nazirliyin
# öz nümunə iş kitabından (MOE OIL.xlsx, `O&G` vərəqi, sətir 24 "Domestic (estimated)" AZN/t və
# sətir 56 "Domestic" AZN/1000 m³) götürülüb; həmin fayl YALNIZ oxunur və onun 2024-dən sonrakı
# sütunları nümunə-model proyeksiyasıdır, ona görə burada faktiki dövrün (≤2024) sabit dəyəri
# işlədilir. Qiymətlər TƏSDİQLƏNMƏMİŞ (UNVERIFIED) sayılır və tənzimlənən satış tarifi sorğusu
# məlumat tələbi siyahısındadır; blokun mərkəzi konstruksiyaları (çəkili qiymət indeksi və
# əmtəəlik həcm indeksi) bu sabitlərdən ASILI DEYİL — onlar yalnız buraxılışın səviyyəsində və
# daxili/ixrac bölgüsünün dəyər ifadəsində iştirak edir.
OILGAS_P_DOM_OIL_AZN_T = 53.6        # MOE OIL 'O&G' r24, 1995–2024 boyu sabit
OILGAS_P_DOM_GAS_AZN_KM3 = 44.87     # MOE OIL 'O&G' r56, 2022–2024 dəyəri

# Müqavilə səviyyəsində qaz ixrac qiymətləri — eyni nümunə iş kitabının `O&G` vərəqindən
# (sətir 54 "SD" = Şahdəniz, sətir 55 "SOCAR", hər ikisi USD/min m³). Bu sıralar modelin GİRİŞİ
# DEYİL: onlar çəkili qiymət indeksi arqumentinin DƏSTƏKLƏYİCİ SÜBUTUDUR — qazın ixrac qiyməti vahid bir qiymət deyil,
# iki fərqli müqavilə rejimidir, ona görə onun yolunu neftin qiymətindən çıxarmaq mümkün deyil.
# Fayl NÜMUNƏ-MODEL faylıdır və yalnız oxunur; 2024-dən sonrakı sütunları proyeksiya olduğu üçün
# burada YALNIZ faktiki dövr (≤2024) saxlanılır.
OILGAS_MOE_GAS_CONTRACT_PRICES = {      # il -> (Şahdəniz [r54], SOCAR [r55]), USD/min m³
    2018: (161.00, 163.90), 2019: (189.28, 165.84), 2020: (211.63, 155.09),
    2021: (291.60, 150.94), 2022: (786.60, 190.21), 2023: (411.69, 199.72),
    2024: (326.69, 181.65),
}
OILGAS_MOE_SOURCE_NOTE = (
    "MOE OIL.xlsx, `O&G` vərəqi, sətir 54 (SD — Şahdəniz) və 55 (SOCAR), USD/min m³; "
    "Nazirliyin nümunə-model faylı — YALNIZ OXUNUR, yalnız faktiki sütunlar (≤2024)")


# Kənd təsərrüfatı məhsullarının istehsalçı qiymətləri — Nazirliyin nümunə-model faylından
# (MOE AGRI.xlsx, `add4agr_` vərəqi). Fayl YALNIZ OXUNUR; onun 2025–2030 sütunları vərəqin öz
# annotasiya sətrində "proqnoz" kimi işarələnib, ona görə burada YALNIZ faktiki dövr (≤2024)
# saxlanılır. Vərəq aqreqat indeks vermir — məhsul üzrə qiymət (sətir 82–101, AZN/ton) və
# məhsul üzrə buraxılış dəyəri (sətir 105–124, mln AZN) verir; aşağıdakı sıra həmin iki blokdan
# GECİKMİŞ DƏYƏR ÇƏKİLƏRİ ilə qurulmuş Laspeyres qiymət indeksinin illik artımıdır:
#     g(t) = Σ_i V_i(t−1)·[P_i(t)/P_i(t−1)] / Σ_i V_i(t−1) − 1
# QIRILMA: 2010-cu ildə vərəqin qiymət blokunda metodoloji qırılma var (tərəvəz 223→802,
# bostan 148→601, üzüm 320→941 AZN/ton — 2010-a qədərki dəyərlər tam ədəd, sonrakılar onluq
# hissəlidir). Ona görə indeks YALNIZ 2011-dən başlayır. Heyvandarlıq qiymətləri vərəqdə
# 2018-dən mövcud olduğuna görə çəkilər bitkiçilik məhsullarındandır (11 məhsul).
MOE_AGRI_PPI_SOURCE_NOTE = (
    "MOE AGRI.xlsx, `add4agr_` vərəqi: qiymətlər sətir 82–92 (AZN/ton), buraxılış dəyərləri "
    "sətir 105–115 (mln AZN); Nazirliyin nümunə-model faylı — YALNIZ OXUNUR, yalnız faktiki "
    "sütunlar (≤2024); 2010-cu ilin metodoloji qırılmasına görə indeks 2011-dən verilir")
MOE_AGRI_PPI_ROWS = [                    # (qiymət sətri, dəyər sətri, gözlənilən etiket)
    (82, 105, "Dənli və dənli paxlalılar"), (83, 106, "Tərəvəz"), (84, 107, "Bostan"),
    (85, 108, "Kartof"), (86, 109, "Pambıq"), (87, 110, "Tütün"), (88, 111, "Şəkər çuğunduru"),
    (89, 112, "Dən üçün günəbaxan"), (90, 113, "Meyvə və giləmeyvə"), (91, 114, "Üzüm"),
    (92, 115, "Çay yarpağı"),
]
MOE_AGRI_PPI_G = {       # il -> illik artım, %
    2011: 5.6274, 2012: -5.1679, 2013: 0.8352, 2014: 0.9291, 2015: 0.3757, 2016: -8.7520,
    2017: 8.5331, 2018: 0.4138, 2019: 3.7395, 2020: 6.0583, 2021: 6.1419, 2022: 22.5822,
    2023: 7.6870, 2024: 2.6370,
}
MOE_AGRI_XLSX = os.environ.get("MOE_AGRI_XLSX", "")  # inkişaf mühitindən kənar YOXDUR (qəsdən) —
# `verify_moe_agri_ppi()` yalnız könüllü, əl ilə çağırılan çarpaz yoxlamadır; boş yol `os.path.exists`
# ilə False verir və funksiya `None` qaytarır (aşağıda), ona görə bu, xəta DEYİL. Nazirliyin nümunə-model
# faylının öz maşınındakı yolu təhvil paketinə YAZILMIR (şəxsi yol sızması olmasın deyə); lokal təkrarlama
# üçün `MOE_AGRI_XLSX` mühit dəyişənini "…/Modellər/Makroekonometrik model/MOE AGRI.xlsx" olaraq təyin edin.


def moe_agri_ppi_growth() -> pd.Series:
    """Kənd təsərrüfatı istehsalçı qiymətlərinin dəyər-çəkili indeksinin artımı, % (2011–2024).

    `eq_dfpi_az` (Nazirliyin öz tənliyi) tənliyinin qida-qiymət termi üçün qurulub. Sıra Nazirliyin nümunə-model
    faylından TÖRƏDİLİB və burada sabit kimi saxlanılır (neft-qaz blokunun müqavilə qiymətləri
    ilə eyni konvensiya): təhvil paketi paketdən kənar fayla ASILI OLMAMALIDIR. Provenans
    `MOE_AGRI_PPI_SOURCE_NOTE`-dadır; `verify_moe_agri_ppi()` fayl əlçatan olduqda sıranı
    mənbədən yenidən qurub tutuşdurur.
    """
    return pd.Series(MOE_AGRI_PPI_G, name="fpi_az_moe").sort_index()


def verify_moe_agri_ppi(path: str = MOE_AGRI_XLSX, tol: float = 1e-3):
    """Sabit sıranı MƏNBƏDƏN yenidən qurur (etiket yoxlaması ilə) və fərqi qaytarır.

    Fayl yoxdursa `None` qaytarır — təhvil mühitində fayl olmaya bilər və bu, xəta deyil.
    Fayl varsa hər sətir üçün gözlənilən etiket yoxlanılır; uyğunsuzluqda `AssertionError`.
    """
    if not os.path.exists(path):
        return None
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["add4agr_"]
    rows = list(ws.iter_rows(min_row=1, max_row=130, max_col=45, values_only=True))
    hdr = rows[2]
    ycol = {hdr[i]: i for i in range(5, 42) if isinstance(hdr[i], int)}
    P, V = {}, {}
    for pr, vr, label in MOE_AGRI_PPI_ROWS:
        for rr in (pr, vr):
            got = str(rows[rr - 1][1] or "").strip()
            if _key(got) != _key(label):
                raise AssertionError(
                    f"ETİKET UYĞUNSUZLUĞU — MOE AGRI add4agr_!{rr}: gözlənilən {label!r}, tapılan {got!r}")
        P[label] = pd.Series({y: rows[pr - 1][c] for y, c in ycol.items()
                              if isinstance(rows[pr - 1][c], (int, float))})
        V[label] = pd.Series({y: rows[vr - 1][c] for y, c in ycol.items()
                              if isinstance(rows[vr - 1][c], (int, float))})
    P = pd.DataFrame(P).sort_index(); V = pd.DataFrame(V).sort_index()
    P = P[P.index <= 2024]; V = V[V.index <= 2024]
    out = {}
    for t in P.index:
        if t < 2011 or (t - 1) not in P.index:
            continue
        num = den = 0.0
        for c in P.columns:
            p1, p0, v0 = P[c].get(t), P[c].get(t - 1), V[c].get(t - 1)
            if pd.isna(p1) or pd.isna(p0) or p0 == 0 or pd.isna(v0) or v0 <= 0:
                continue
            num += v0 * (p1 / p0); den += v0
        if den > 0:
            out[int(t)] = (num / den - 1.0) * 100.0
    got = pd.Series(out).sort_index()
    want = moe_agri_ppi_growth()
    common = sorted(set(got.index) & set(want.index))
    diff = float((got.loc[common] - want.loc[common]).abs().max()) if common else float("nan")
    if common and diff > tol:
        raise AssertionError(f"MOE AGRI qiymət indeksi mənbədən fərqlənir: maks. fərq {diff:.6f}")
    return diff


def moe_gas_contract_prices() -> pd.DataFrame:
    """Müqavilə səviyyəsində qaz ixrac qiymətləri (Şahdəniz / SOCAR), USD/min m³, 2018–2024.

    **Modelə daxil edilmir.** Sıralar Nazirliyin nümunə-model faylından yalnız oxunur və çəkili qiymət indeksinin
    arqumentini sənədləşdirmək üçün çap olunur: iki müqavilə rejimi arasındakı fərq 2022-ci ildə
    dörd dəfədən çox olub, yəni "qazın ixrac qiyməti" tək bir davranış tənliyi ilə təsvir oluna
    bilməyən, rejimlərə bölünmüş kəmiyyətdir. Provenans `OILGAS_MOE_SOURCE_NOTE`-dadır.
    """
    d = OILGAS_MOE_GAS_CONTRACT_PRICES
    return pd.DataFrame({"sd_usd_km3": {y: v[0] for y, v in d.items()},
                         "socar_usd_km3": {y: v[1] for y, v in d.items()}}).sort_index()


# registr kodu -> (vərəq, sətir, gözlənilən etiket)
OILGAS_SOURCES = {
    "oilgas_r003": ("Neft-Qaz sektoru", 3, "Neft hasilatı"),
    "oilgas_r004": ("Neft-Qaz sektoru", 4, "Əmtəəlik neft hasilatı"),
    "oilgas_r005": ("Neft-Qaz sektoru", 5, "Neft ixracı"),
    "oilgas_r006": ("Neft-Qaz sektoru", 6, "Neft ixracı"),
    "oilgas_r010": ("Neft-Qaz sektoru", 10, "Neftin ixrac qiyməti, 1 barel, DGK*"),
    "oilgas_r012": ("Neft-Qaz sektoru", 12, "Qaz hasilatı"),
    "oilgas_r013": ("Neft-Qaz sektoru", 13, "Əmtəəlik qaz hasilatı"),
    "oilgas_r014": ("Neft-Qaz sektoru", 14, "Qaz ixracı (DGK)"),
    "oilgas_r015": ("Neft-Qaz sektoru", 15, "Qaz ixracı (DGK)"),
    "oilgas_r018": ("Neft-Qaz sektoru", 18, "Qazın ixrac qiyməti, min kub metr (DGK)"),
    "mine_r003": ("Mədənçıxarma", 3, "Mədənçıxarma sənayesi üzrə buraxılış"),
    "mine_r004": ("Mədənçıxarma", 4, "Mədənçıxarma sənayesi üzrə əlavə dəyər"),
    "mine_r008": ("Mədənçıxarma", 8, "Xam neft və təbii qaz hasilatı üzrə buraxılış"),
    "mine_r009": ("Mədənçıxarma", 9, "Xam neft və təbii qaz hasilatı üzrə əlavə dəyər"),
}


# ==========================================================================
# 8b. Sahə sürücüləri və sahə qiymət indeksləri (FR02 — sürücü və qiymət sırası xəritəsi)
# ==========================================================================
# Hər sahə üçün BİR adlandırılmış sürücü (adlandırılmış sürücü qaydası) və hər deflyator üçün BİR sahə-spesifik
# qiymət sırası (sahə qiymət sıraları qaydası). Sıralar əsas iş kitabının registrindən sətir nömrəsi VƏ etiket
# yoxlaması ilə götürülür — `wb_labelled` eyni mexanizmi neft-qaz bloku üçün artıq tətbiq edir.
# DVX vərəqinin 5 müşahidəli dövriyyə sətirləri (r53-r66) BURADA YALNIZ lövbər/çarpaz yoxlama
# kimi qeyd olunur: onların üzərində elastiklik qiymətləndirilmir (D5).
SECTOR_DRIVER_SOURCES = {
    # --- sahə sürücüləri (həcm / kredit / xərc) ---
    "money_r044": ("Monetar sektoru", 44, "Kənd təsərrüfatı, meşə təsərrüfatı və balıqçılıq sektoru"),
    "money_r046": ("Monetar sektoru", 46, "Sənaye və istehsal sektoru"),
    "money_r043": ("Monetar sektoru", 43,
                   "Mədənçıxarma və elektrik enerjisi, qaz, buxar və su təsərrüfatı sektoru"),
    "money_r047": ("Monetar sektoru", 47, "Nəqliyyat və rabitə sektoru"),
    "money_r042": ("Monetar sektoru", 42, "Ticarət və xidmət sektoru"),
    "money_r045": ("Monetar sektoru", 45, "İnşaat və tikinti sektoru"),
    "real_r109": ("Real sektor", 109,
                  "Tikinti-quraşdırma işləri üzrə əsas kapitala yönəldilən vəsaitin həcmi"),
    "real_r110": ("Real sektor", 110, "real artım tempi"),
    "social_r015": ("Sosial sektor", 15, "Pərakəndə ticarət dövriyyəsi"),
    "social_r016": ("Sosial sektor", 16, "real artım tempi"),
    "social_r005": ("Sosial sektor", 5, "İctimai iaşə dövriyyəsinin həcmi"),
    "social_r006": ("Sosial sektor", 6, "real artım tempi"),
    "social_r027": ("Sosial sektor", 27, "Əhaliyə göstərilən ödənişli xidmətlərin dəyəri"),
    "social_r028": ("Sosial sektor", 28, "real artım tempi"),
    "social_r055": ("Sosial sektor", 55, "Dövlət sektorunda ƏH"),
    "social_r071": ("Sosial sektor", 71, "Orta aylıq pensiya"),
    "elec_r004": ("Elektrik enerjisi", 4,
                  "Elektrik enerjisi, qaz və buxar istehsalı, bölüşdürülməsi və təchizatı üzrə buraxılış"),
    "wsup_r004": ("Su təchizatı", 4,
                  "Su təchizatı, tullantıların təmizlənməsi və emalı üzrə məhsul buraxılışı"),
    "real_r060": ("Real sektor", 60, "real artım tempi"),
    "real_r062": ("Real sektor", 62, "real artım tempi"),
    "fiscal_r017": ("Fiskal sektor", 17, "Cari xərclər"),
    "fiscal_r019": ("Fiskal sektor", 19, "Dövlət büdcəsinin dövlət əsaslı vəsait qoyuluşu"),
    "dip_r005": ("DİP 2016-2026", 5, "o cümlədən: İnfrastruktur layihələri"),
    "bop_r018": ("Tədiyyə Balansı", 18, "Nəqliyyat xidmətləri"),
    "bop_r022": ("Tədiyyə Balansı", 22, "Turizm xidmətləri"),
    "bop_r028": ("Tədiyyə Balansı", 28, "Telekomunikasiya (rabitə) xidmətləri"),
    # --- sahə buraxılışının nominal səviyyəsi + real artım tempi (implisit sahə deflyatoru) ---
    # EYNİ vərəqin qonşu sətirləri işlədilir: nominal səviyyə bir vərəqdən, real artım tempi
    # başqasından götürülsəydi, iki mənbənin səviyyə tərifi üst-üstə düşməyə bilərdi və implisit
    # deflyator həmin tərif fərqini qiymət hərəkəti kimi göstərərdi.
    "real_r053": ("Real sektor", 53, "Qeyri neft-qaz sənayesi"),
    "real_r054": ("Real sektor", 54, "real artım tempi"),
    "real_r057": ("Real sektor", 57, "Emal sənayesi üzrə"),
    "real_r058": ("Real sektor", 58, "real artım tempi"),
    "real_r059": ("Real sektor", 59,
                  "Elektrik enerjisi, qaz və buxar istehsalı, bölüşdürülməsi və təchizatı üzrə"),
    "real_r061": ("Real sektor", 61, "Su təchizatı; tullantıların təmizlənməsi və emalı üzrə"),
    "real_r063": ("Real sektor", 63, "Kənd, meşə və balıqçılıq təsərrüfatları"),
    "real_r064": ("Real sektor", 64, "real artım tempi"),
    # --- sahə qiymət indeksləri (deflyator sürücüləri) ---
    "real_r188": ("Real sektor", 188, "Sənaye məhsullarının orta illik istehsalçı qiymət indeksi"),
    "real_r189": ("Real sektor", 189,
                  "Qeyri neft-qaz sənayesində orta illik istehsalçı qiymət indeksləri"),
    "real_r190": ("Real sektor", 190, "Mədənçıxarma sənayesi üzrə"),
    "real_r191": ("Real sektor", 191, "Emal sənayesi üzrə"),
    "real_r192": ("Real sektor", 192,
                  "Elektrik enerjisi, qaz və buxar istehsalı, bölüşdürülməsi və təchizatı üzrə"),
    "real_r193": ("Real sektor", 193, "Su təchizatı; tullantıların təmizlənməsi və emalı üzrə"),
    "real_r194": ("Real sektor", 194,
                  "Kənd təsərrüfatı məhsullarının orta illik istehsalçı qiymət indeksi"),
    "real_r195": ("Real sektor", 195, "Bitkiçilik məhsulları"),
    "real_r196": ("Real sektor", 196, "heyvandarlıq məhsullarə"),
    "real_r197": ("Real sektor", 197,
                  "Nəqliyyat və anbar təsərrüfatı xidmətlərinin orta illik qiymət indeksi"),
    "real_r199": ("Real sektor", 199, "Rabitə xidmətlərinin orta illik qiymət indeksi"),
    # --- DVX dövriyyəsi: YALNIZ lövbər (n=5, D5 — tənlik sürücüsü kimi işlədilmir) ---
    "dvx_r057": ("DVX üzrə göstəricilər", 57, "qeyri neft-qaz sənayesi üzrə dövriyyə"),
    "dvx_r058": ("DVX üzrə göstəricilər", 58, "Kənd təsərrüfatı üzrə dövriyyə"),
    "dvx_r059": ("DVX üzrə göstəricilər", 59, "Tikinti sektrou üzrə dövriyyə"),
    "dvx_r060": ("DVX üzrə göstəricilər", 60, "Ticarət sektrou üzrə dövriyyə"),
    "dvx_r063": ("DVX üzrə göstəricilər", 63, "Turizm və ictimai iaşə sektrou üzrə dövriyyə"),
    "dvx_r064": ("DVX üzrə göstəricilər", 64, "Nəqliyyat və anbar təsərrüfatı sektrou üzrə dövriyyə"),
    "dvx_r065": ("DVX üzrə göstəricilər", 65, "İnformasiya və rabitə sektrou üzrə dövriyyə"),
    "dvx_r066": ("DVX üzrə göstəricilər", 66, "Sosial və digər xidmətlər sektrou üzrə dövriyyə"),
    # --- A4: dolayı vergilərin neft-qaz payı (φ-nin zamana görə dəyişməsi, kalibrləmə) ---
    "dvx_r067": ("DVX üzrə göstəricilər", 67, "Əlavə dəyər vergisi-CƏMİ"),
    "dvx_r068": ("DVX üzrə göstəricilər", 68, "Aksiz vergisi"),
    "dvx_r074": ("DVX üzrə göstəricilər", 74, "Mədən vergisi"),
    "dvx_r075": ("DVX üzrə göstəricilər", 75, "Yol vergisi"),
    "dvx_r079": ("DVX üzrə göstəricilər", 79, "Əlavə dəyər vergisi (qeyri-neft-qaz sektoru üzrə)"),
    "dvx_r080": ("DVX üzrə göstəricilər", 80, "Aksiz vergisi"),
    "dvx_r086": ("DVX üzrə göstəricilər", 86, "Mədən vergisi"),
    "dvx_r087": ("DVX üzrə göstəricilər", 87, "Yol vergisi"),
}

# --- Tədiyə balansının bağlanışı üçün monetar sektorun ehtiyat sıraları (FR11/FR12) -------
# `money_r003` STOK göstəricisidir (dövrün sonuna, mlrd USD) və üç komponentin cəmidir:
# ARDNF (r4), MB (r5), MN (r6). FR11 ilkin gəlir daxilolmalarının sürücüsü kimi məhz bu stoku
# işlədir (investisiya gəlirinin bazası əvvəlki ilin ehtiyat həcmidir); FR12 isə onun illik
# dəyişməsini tədiyə balansından törədilən ehtiyat AXINI ilə tutuşdurur. İki kəmiyyət EYNİ
# DEYİL — stok dəyişməsi qiymət/məzənnə yenidənqiymətləndirmələrini də daşıyır — və bu fərq
# FR12-də açıq nəşr olunur.
BOP_SOURCES = {
    "money_r003": ("Monetar sektoru", 3, "Strateji valyuta ehtiyatları"),
    "money_r004": ("Monetar sektoru", 4, "ARDNF ehtiyatları"),
    "money_r005": ("Monetar sektoru", 5, "MB ehtiyatları"),
}


def labelled_sources() -> dict:
    """`wb_labelled`-in tanıdığı bütün sıralar: neft-qaz bloku + sahə sürücüləri + fiskal blok.

    Birləşdirmə HƏR ÇAĞIRIŞDA aparılır (modul səviyyəsində bir dəfə deyil): əks halda üç
    lüğətdən birinə sonradan edilən dəyişiklik — məsələn testin yanlış etiketi qəsdən
    yerləşdirməsi — yoxlamadan yayınardı. `FISCAL_SOURCES` aşağıda, 9-cu bölmədə təyin olunur;
    modul tam yükləndikdən sonra çağırıldığı üçün irəli istinad problemi yaratmır.
    """
    return {**OILGAS_SOURCES, **SECTOR_DRIVER_SOURCES, **FISCAL_SOURCES, **BOP_SOURCES}


def wb_labelled(series_code: str) -> pd.Series:
    """`wb_annual` + MƏCBURİ etiket yoxlaması (registrin `name_az` sütunu üzrə).

    Sətir nömrəsinə görə oxumaq özlüyündə kifayət deyil: iş kitabının yeni buraxılışında
    sətirlər sürüşə bilər. Ona görə hər müraciətdə vərəq, sətir və etiket üçlüyü
    `labelled_sources()`-dakı (neft-qaz bloku + sahə sürücüləri) gözlənilən dəyərlərlə
    tutuşdurulur.
    """
    _src = labelled_sources()
    if series_code not in _src:
        raise KeyError(f"'{series_code}' etiket-yoxlamalı mənbə siyahısında yoxdur")
    sheet, row, label = _src[series_code]
    reg = _load(C_WB_REGISTRY, "wb_registry")
    r = reg[reg["series_code"] == series_code]
    if r.empty:
        raise KeyError(f"iş kitabı registrində '{series_code}' yoxdur")
    got_sheet, got_row, got_label = str(r.iloc[0]["sheet"]), int(r.iloc[0]["row"]), str(r.iloc[0]["name_az"])
    if got_sheet != sheet or got_row != row or _key(got_label) != _key(label):
        raise AssertionError(
            f"ETİKET UYĞUNSUZLUĞU — {series_code}: gözlənilən {sheet}!{row} {label!r}, "
            f"tapılan {got_sheet}!{got_row} {got_label!r}")
    return wb_annual(series_code)


def sector_driver(series_code: str) -> pd.Series:
    """Sahə sürücüsünün / sahə qiymət indeksinin FAKTİKİ hissəsi — etiket yoxlaması ilə, ≤2025.

    `wb_labelled` xam registr sırasını qaytarır; illik yığım natamam illəri artıq atır, lakin
    burada kəsim AÇIQ tətbiq olunur: sürücü sırası heç bir halda proqnoz üfüqünə uzana bilməz
    (BUILD_CONTEXT §D6 birinci bənd — heç bir qiymətləndirmə nümunəsi 2026–2030-a toxunmur).
    """
    s = wb_labelled(series_code).dropna()
    return s[s.index <= LAST_ACTUAL]


def weighted_price_growth(p_oil, p_gas, w_oil) -> pd.Series:
    """Dəyər payı ilə çəkilənmiş, zəncirlənmiş neft+qaz qiymət indeksinin artımı, %.

        d(t) = w(t-1)·[P^neft(t)/P^neft(t-1)] + (1−w(t-1))·[P^qaz(t)/P^qaz(t-1)] − 1

    `w` — ixrac DƏYƏRİNDƏ neftin payı; gecikmiş çəki (Laspeyres konvensiyası) işlədilir ki,
    cari ilin qiymət hərəkəti öz çəkisini özü təyin etməsin. Qaz qiyməti verilməyən illərdə
    (2006-dan əvvəl qaz ixracı praktik olaraq sıfırdır) çəki 1-ə bərabər tutulur və indeks
    neft qiymətinin indeksinə çevrilir.
    """
    p_oil = pd.Series(p_oil).astype(float).sort_index()
    p_gas = pd.Series(p_gas).astype(float).sort_index()
    w_oil = pd.Series(w_oil).astype(float).sort_index()
    out = {}
    for t in p_oil.index:
        if (t - 1) not in p_oil.index:
            continue
        po1, po0 = p_oil.get(t), p_oil.get(t - 1)
        if pd.isna(po1) or pd.isna(po0) or po0 == 0:
            continue
        wo = w_oil.get(t - 1)
        if pd.isna(wo):
            continue
        ro = po1 / po0
        pg1, pg0 = p_gas.get(t), p_gas.get(t - 1)
        rg = (pg1 / pg0) if (pd.notna(pg1) and pd.notna(pg0) and pg0 != 0) else ro
        out[int(t)] = (wo * ro + (1.0 - wo) * rg - 1.0) * 100.0
    return pd.Series(out, name="oilgas_pw_g").sort_index()


def laspeyres_volume_growth(q_oil, q_gas, p_oil_lag, p_gas_lag) -> pd.Series:
    """Neft və qazın ƏVVƏLKİ İLİN qiymətləri ilə çəkilənmiş həcm indeksinin artımı, %.

        g(t) = [Q^n(t)·P^n(t-1) + Q^q(t)·P^q(t-1)] / [Q^n(t-1)·P^n(t-1) + Q^q(t-1)·P^q(t-1)] − 1

    Bu, iki məhsulun natural həcmlərini vahid ölçüyə gətirməyin standart üsuludur (Laspeyres
    həcm indeksi): mln ton və mlrd m³ birbaşa toplana bilməz, dəyər çəkisi onları toplanan
    edir. Qiymətlər manatla, eyni ölçü vahidi başına verilməlidir.
    """
    q_oil = pd.Series(q_oil).astype(float).sort_index()
    q_gas = pd.Series(q_gas).astype(float).sort_index()
    p_oil_lag = pd.Series(p_oil_lag).astype(float).sort_index()
    p_gas_lag = pd.Series(p_gas_lag).astype(float).sort_index()
    out = {}
    for t in q_oil.index:
        if (t - 1) not in q_oil.index:
            continue
        vals = [q_oil.get(t), q_oil.get(t - 1), q_gas.get(t), q_gas.get(t - 1),
                p_oil_lag.get(t - 1), p_gas_lag.get(t - 1)]
        if any(pd.isna(v) for v in vals):
            continue
        qo1, qo0, qg1, qg0, po0, pg0 = (float(v) for v in vals)
        den = qo0 * po0 + qg0 * pg0
        if den <= 0:
            continue
        out[int(t)] = ((qo1 * po0 + qg1 * pg0) / den - 1.0) * 100.0
    return pd.Series(out, name="oilgas_volg").sort_index()


def oilgas_block() -> pd.DataFrame:
    """Neft-qaz blokunun tarixi cədvəli (neft-qaz bloku: həcm × qiymət) — həcm, qiymət, dəyər, çəki və ρ.

    Sütunlar (indeks = il, hamısı ≤ LAST_ACTUAL):
      `q_oil_prod/mkt/exp/dom` mln ton · `q_gas_prod/mkt/exp/dom` mlrd m³ ·
      `v_oil_exp/v_gas_exp` mln USD · `p_oil_exp` USD/barel · `p_gas_exp` USD/min m³ ·
      `p_brent` USD/barel · `fx_avg` AZN/USD (orta illik) · `w_oil_exp` neftin ixrac
      dəyərindəki payı · `p_oil_azn_t`, `p_gas_azn_km3` manatla ixrac qiymətləri ·
      `go_oilgas_obs`, `va_oilgas_obs` müşahidə olunan buraxılış və əlavə dəyər (mln AZN) ·
      `rho` = əlavə dəyər / buraxılış.

    Daxili həcm EYNİLİKlə hesablanır: daxili = əmtəəlik − ixrac. **DATA-GAP**: neftin ixrac
    həcmi sırası 2016-cı ilə qədər əmtəəlik hasilatla uzlaşmır (2008-ci ildə ixrac hasilatdan
    böyük görünür), ona görə neft üzrə daxili həcm yalnız 2016-dan etibarən etibarlıdır; qaz
    üzrə eynilik 2006-dan qapanır. Bu, sıraların özündə olan uyğunsuzluqdur, gizlədilmir.
    """
    cut = lambda s: s[s.index <= LAST_ACTUAL]
    q_oil_prod = cut(wb_labelled("oilgas_r003"))
    q_oil_mkt = cut(wb_labelled("oilgas_r004"))
    q_oil_exp = cut(wb_labelled("oilgas_r005"))
    v_oil_exp = cut(wb_labelled("oilgas_r006"))
    p_oil_exp = cut(wb_labelled("oilgas_r010"))
    q_gas_prod = cut(wb_labelled("oilgas_r012"))
    q_gas_mkt = cut(wb_labelled("oilgas_r013"))
    q_gas_exp = cut(wb_labelled("oilgas_r014"))
    v_gas_exp = cut(wb_labelled("oilgas_r015"))
    p_gas_exp = cut(wb_labelled("oilgas_r018"))
    go_obs = cut(wb_labelled("mine_r008"))
    va_obs = cut(wb_labelled("mine_r009"))
    go_mine = cut(wb_labelled("mine_r003"))
    va_mine = cut(wb_labelled("mine_r004"))

    brent = actuals("brent_usd")
    eop = actuals("fx_usd_azn_eop")
    fx_avg = ((eop + eop.shift(1)) / 2.0)          # orta illik məzənnəyə iki nöqtəli yaxınlaşma

    df = pd.DataFrame({
        "q_oil_prod": q_oil_prod, "q_oil_mkt": q_oil_mkt, "q_oil_exp": q_oil_exp,
        "q_gas_prod": q_gas_prod, "q_gas_mkt": q_gas_mkt, "q_gas_exp": q_gas_exp,
        "v_oil_exp": v_oil_exp, "v_gas_exp": v_gas_exp,
        "p_oil_exp": p_oil_exp, "p_gas_exp": p_gas_exp,
        "p_brent": brent, "fx_avg": fx_avg,
        "go_oilgas_obs": go_obs, "va_oilgas_obs": va_obs,
        "go_mine_obs": go_mine, "va_mine_obs": va_mine,
    })
    df = df[df.index <= LAST_ACTUAL].sort_index()
    df["q_oil_dom"] = df["q_oil_mkt"] - df["q_oil_exp"]
    df["q_gas_dom"] = df["q_gas_mkt"] - df["q_gas_exp"]
    # ixrac dəyərində neftin payı; qaz ixracı olmayan illərdə pay 1-dir
    tot = df["v_oil_exp"] + df["v_gas_exp"].fillna(0.0)
    df["w_oil_exp"] = (df["v_oil_exp"] / tot).where(tot > 0)
    df.loc[df["v_gas_exp"].isna() & df["v_oil_exp"].notna(), "w_oil_exp"] = 1.0
    df["p_oil_azn_t"] = df["p_oil_exp"] * BBL_PER_TON * df["fx_avg"]      # AZN / ton
    df["p_gas_azn_km3"] = df["p_gas_exp"] * df["fx_avg"]                  # AZN / min m³
    df["rho"] = df["va_oilgas_obs"] / df["go_oilgas_obs"]
    return df


def oilgas_price_growth(oil_leg: str = "brent") -> pd.Series:
    """Çəkili neft+qaz ixrac qiyməti indeksinin artımı, % (tarixi hissə).

    `oil_leg="brent"` — neft ayağı dünya qiyməti (Brent) ilə ölçülür; bu, köhnə
    spesifikasiya ilə BİR dəyişən fərqi olan variantdır (yeganə fərq qaz ayağının əlavə
    olunmasıdır) və ona görə d22 süni dəyişəninin sınağında məhz bu variant işlədilir.
    `oil_leg="realised"` — Azərbaycan neftinin faktiki ixrac qiyməti.
    """
    b = oilgas_block()
    p_oil = b["p_brent"] if oil_leg == "brent" else b["p_oil_exp"]
    if oil_leg not in ("brent", "realised"):
        raise ValueError("oil_leg yalnız 'brent' və ya 'realised' ola bilər")
    return weighted_price_growth(p_oil, b["p_gas_exp"], b["w_oil_exp"])


def oilgas_volume_growth() -> pd.Series:
    """Əmtəəlik neft və qaz hasilatının dəyər çəkili həcm artımı, % (tarixi hissə)."""
    b = oilgas_block()
    return laspeyres_volume_growth(b["q_oil_mkt"], b["q_gas_mkt"],
                                   b["p_oil_azn_t"], b["p_gas_azn_km3"])


# ==========================================================================
# 9. Fiskal blok — büdcə sıraları, DİP və qeyri-neft baza kəsiri
# ==========================================================================
# İlkin icmalın fiskal blok tapıntısı: paketin əvvəlki buraxılışında NƏ bir fiskal sıra, NƏ də bir fiskal
# tənlik var idi. Məlumat isə tarixi hissə üçün tamdır. Aşağıdakı lüğət həmin sıraları
# `wb_labelled`-in etiket yoxlamasına daxil edir — sətir nömrəsi iş kitabının yeni
# buraxılışında sürüşərsə blok səssizcə başqa sətri oxumur, DAYANIR.
FISCAL_SOURCES = {
    "fiscal_r003": ("Fiskal sektor", 3, "Gəlirlər"),
    "fiscal_r005": ("Fiskal sektor", 5, "Neft-qaz gəlirləri"),
    "fiscal_r007": ("Fiskal sektor", 7, "ARDNF transfert"),
    "fiscal_r008": ("Fiskal sektor", 8, "Qeyri-neft-qaz gəlirləri"),
    "fiscal_r016": ("Fiskal sektor", 16, "Dövlət büdcəsinin xərcləri"),
    "fiscal_r018": ("Fiskal sektor", 18, "Əsaslı xərclər"),
    "fiscal_r020": ("Fiskal sektor", 20, "Borca xidmət xərcləri"),
    # funksional təsnifat — sosial sahə tənliyinin fiskal kanalının yenidən sınanması üçün
    # (sosial sətri). Sıralar 2019-dan başlayır: n = 6-7, D5 çərçivəsində qiymətləndirilir.
    "fiscal_r036": ("Fiskal sektor", 36, "400 Təhsil"),
    "fiscal_r045": ("Fiskal sektor", 45, "500 Səhiyyə"),
    "fiscal_r051": ("Fiskal sektor", 51, "600 Sosial müdafiə və sosial təminat"),
    "fiscal_r021": ("Fiskal sektor", 21, "Büdcə kəsiri/profisiti"),
    "dip_r004": ("DİP 2016-2026", 4, "Dövlət əsaslı vəsait qoyuluşunun cəmi"),
    "dip_r004_plan": ("DİP 2016-2026", 4, "Dövlət əsaslı vəsait qoyuluşunun cəmi — plan"),
    "dip_r009": ("DİP 2016-2026", 9,
                 "Xarici kreditor qurumlarla birgə maliyyələşmə əsasında həyata keçirilən "
                 "layihələr üzrə kreditlərin məbləği"),
    "real_r005": ("Real sektor", 5, "Qeyri neft-qaz ÜDM, bazar qiymətləri ilə"),
    "real_r080": ("Real sektor", 80, "Əsas kapitala cəmi investisiyalar"),
    "real_r081": ("Real sektor", 81, "real artım tempi"),
    "real_r082": ("Real sektor", 82, "Neft-qaz sektoruna investisiyalar"),
    "real_r084": ("Real sektor", 84, "Qeyri neft-qaz sektoruna investisiyalar"),
    "real_r098": ("Real sektor", 98, "Dövlət investisiyaları"),
    "real_r100": ("Real sektor", 100, "Qeyri-dövlət investisiyaları"),
    "real_r102": ("Real sektor", 102, "Qeyri-dövlət qeyri neft-qaz investisiyaları"),
    "real_r114": ("Real sektor", 114,
                  "Müəssisə və təşkilatların öz vəsaitləri üzrə əsas kapitala yönəldilmiş "
                  "investisiyalar"),
    "real_r116": ("Real sektor", 116,
                  "Büdcə vəsaitləri hesabına əsas kapitala yönəldilmiş investisiyalar"),
}


def fiscal_annual() -> pd.DataFrame:
    """Dövlət büdcəsinin illik sıraları (mln AZN, ≤ LAST_ACTUAL) — hamısı etiket yoxlaması ilə.

    Sütunlar: `revenue_total`, `revenue_oil`, `revenue_nonoil`, `sofaz_transfer`,
    `exp_total`, `exp_current`, `exp_capital`, `state_capex_fiscal`, `debt_service`,
    `balance`. `state_capex_fiscal` XAM sıradır (2022 və 2023 boşdur) — DİP ilə
    birləşdirilmiş variant üçün `state_capex()`.
    """
    cut = lambda s: s[s.index <= LAST_ACTUAL]
    cols = {
        "revenue_total": "fiscal_r003", "revenue_oil": "fiscal_r005",
        "revenue_nonoil": "fiscal_r008", "sofaz_transfer": "fiscal_r007",
        "exp_total": "fiscal_r016", "exp_current": "fiscal_r017",
        "exp_capital": "fiscal_r018", "state_capex_fiscal": "fiscal_r019",
        "debt_service": "fiscal_r020", "balance": "fiscal_r021",
    }
    return pd.DataFrame({k: cut(wb_labelled(v)) for k, v in cols.items()}).sort_index()


# `fiscal_r019` (Dövlət büdcəsinin dövlət əsaslı vəsait qoyuluşu) illik sütunu 2022 və 2023
# üçün BOŞDUR. `DİP 2016-2026` vərəqinin 4-cü sətri həmin iki ili doldurur. Birləşdirmə
# YALNIZ çatışmayan illərə tətbiq olunur: hər iki mənbənin dəyəri olduğu illərdə fiskal
# vərəqin öz rəqəmi saxlanılır, çünki nəşr olunan büdcə icrası odur.
STATE_CAPEX_SPLICE_YEARS = (2022, 2023)
STATE_CAPEX_SOURCE_NOTE = (
    "fiscal_r019 (Fiskal sektor, 4 vərəq sətir 19); 2022 və 2023 illəri `DİP 2016-2026` "
    "vərəqinin 4-cü sətrindən doldurulub. 2025-ci ildə iki mənbə üst-üstə düşür "
    "(2 305,1 mln AZN), 2021 və 2024-cü illərdə isə fərqlənir — fərq gizlədilmir, "
    "`state_capex_gap()` ilə çap olunur.")


def state_capex() -> pd.Series:
    """Dövlət büdcəsinin əsaslı vəsait qoyuluşu (DƏVQ), mln AZN — DİP ilə birləşdirilmiş.

    Dövlət investisiyası tənliyinin birinci termi. Birləşdirmə qaydası `STATE_CAPEX_SOURCE_NOTE`-dadır.
    """
    base = wb_labelled("fiscal_r019")
    base = base[base.index <= LAST_ACTUAL].copy()
    dip = wb_labelled("dip_r004")
    dip = dip[dip.index <= LAST_ACTUAL]
    for y in STATE_CAPEX_SPLICE_YEARS:
        if y in dip.index and (y not in base.index or pd.isna(base.get(y))):
            base.loc[y] = float(dip.loc[y])
    return base.dropna().sort_index()


def state_capex_gap() -> pd.DataFrame:
    """İki mənbənin (Fiskal sektor r19 və DİP r4) üst-üstə düşən illərdəki fərqi, mln AZN və %.

    Fərq DÜZƏLDİLMİR — nəşr olunur: iki rəsmi mənbə eyni anlayış üçün fərqli rəqəm verirsə,
    bu, məlumat sorğusuna gedən bir sualdır, modelin gizlədəcəyi bir şey deyil.
    """
    a = wb_labelled("fiscal_r019"); a = a[a.index <= LAST_ACTUAL]
    b = wb_labelled("dip_r004"); b = b[b.index <= LAST_ACTUAL]
    df = pd.DataFrame({"fiscal_r019": a, "dip_r004": b}).dropna()
    df["fərq"] = df["dip_r004"] - df["fiscal_r019"]
    df["fərq_%"] = (df["dip_r004"] / df["fiscal_r019"] - 1.0) * 100.0
    return df


NONOIL_BASE_LABEL_AZ = (
    "dövlət büdcəsi üzrə qeyri-neft balansı / qeyri-neft ÜDM (icmal büdcə perimetri və "
    "rəsmi 'baza' düzəlişi təqdim olunduqdan sonra yenilənəcək)")


def nonoil_base_balance() -> pd.DataFrame:
    """Qeyri-neft baza kəsiri göstəricisinin DÖVLƏT BÜDCƏSİ analoqu (fiskal kanal 3).

    ```
    balans = fiscal_r008 (qeyri-neft-qaz gəlirləri) − fiscal_r016 (dövlət büdcəsinin xərcləri)
    nisbət = balans / real_r005 (qeyri neft-qaz ÜDM, bazar qiymətləri)
    ```

    **İCMAL BÜDCƏ PERİMETRİ DEYİL.** Əsas iş kitabında icmal büdcə (DSMF, İSF, İTSF, Naxçıvan
    MR) üzrə bir sıra da yoxdur və rəsmi "baza" düzəliş qaydası bizdə deyil (S2 §4c/§4d).
    Sıra `NONOIL_BASE_LABEL_AZ` etiketi ilə nəşr olunur; icmal rəqəmi UYDURULMUR.
    """
    rev = wb_labelled("fiscal_r008"); rev = rev[rev.index <= LAST_ACTUAL]
    exp = wb_labelled("fiscal_r016"); exp = exp[exp.index <= LAST_ACTUAL]
    gdp = wb_labelled("real_r005"); gdp = gdp[gdp.index <= LAST_ACTUAL]
    df = pd.DataFrame({"revenue_nonoil": rev, "expenditure": exp, "gdp_nonoil": gdp}).dropna()
    df["balance"] = df["revenue_nonoil"] - df["expenditure"]
    df["ratio_nonoil_gdp"] = df["balance"] / df["gdp_nonoil"] * 100.0
    return df


# --- MOE FISCAL.xlsx `İcmal büdcə` — YALNIZ DƏSTƏKLƏYİCİ SÜBUT ---------------------------
# Nazirliyin NÜMUNƏ-MODEL faylı. Sıralar burada sabit kimi saxlanılır (neft-qaz blokunun
# müqavilə qiymətləri ilə eyni konvensiya: təhvil paketi paketdən kənar fayla asılı olmamalıdır)
# və YALNIZ faktiki dövr (≤2024) götürülür — vərəqin annotasiya sətri 2025-i "gözlənilən",
# 2026-2030-u "proqnoz" kimi işarələyir. Bu sıralar MODELƏ DAXİL EDİLMİR və `forecast_long.csv`-ə
# yazılmır: onlar yuxarıdakı dövlət büdcəsi analoqunun ÇARPAZ YOXLAMASIDIR (R17).
MOE_ICMAL_SOURCE_NOTE = (
    "MOE FISCAL.xlsx, `İcmal büdcə` vərəqi: sətir 4 (icmal gəlirlər), 12 (icmal xərclər), "
    "21 (kəsir/profisit), 38 (qeyri-neft büdcə gəlirləri), 39 (qeyri-neft icmal baza kəsiri), "
    "41 (baza kəsirinin qeyri-neft ÜDM-ə nisbəti); Nazirliyin nümunə-model faylı — YALNIZ "
    "OXUNUR, yalnız faktiki sütunlar (≤2024), modelə daxil edilmir (R17 / S2 §4c)")
MOE_ICMAL_ROWS = [(4, "I. İCMAL BÜDCƏNİN GƏLİRLƏRİ"), (12, "II. İCMAL BÜDCƏNİN XƏRCLƏRİ"),
                  (21, "III. İCMAL BÜDCƏNİN -KƏSİRİ/PROFİSİTİ"), (38, "Qeyri-neft büdcə gəlirləri"),
                  (39, "Qeyri-neft icmal büdcə baza kəsiri"),
                  (40, "Qeyri-neft icmal büdcənin baza kəsirinin ÜDM-ə nisbəti, %"),
                  (41, "Qeyri-neft icmal büdcə baza kəsirinin qeyri-neft ÜDM-ə nisbəti, %")]
# MƏNBƏ QÜSURU (gizlədilmir): 41-ci sətrin məxrəci 2015-ci ilə qədər sınıqdır — həmin illərdə
# sətir −25-dən −2186-ya qədər dəyərlər verir, yəni nisbət deyil. Ona görə qeyri-neft ÜDM-ə
# nisbətin çarpaz yoxlama pəncərəsi 2015–2024-dür; 40-cı sətir (ÜDM-ə nisbət) bütün dövr üçün
# düzgündür və `moe_icmal_budget()` hər ikisini qaytarır.
MOE_ICMAL_RATIO_VALID_FROM = 2015
MOE_ICMAL_BUDGET = {   # il -> (gəlirlər, xərclər, balans, qeyri-neft gəlirlər,
                       #        qeyri-neft baza kəsiri, kəsir/ÜDM, kəsir/qeyri-neft ÜDM)
    2001: (1125.597400, 1077.515710, 48.081690, None, None, None, None),
    2002: (1710.965080, 1650.527739, 60.437341, None, None, None, None),
    2003: (2086.440000, 2095.537600, -9.097600, 1871.194920, -224.342680, -0.031392, -25.493486),
    2004: (2338.557480, 2246.612430, 91.945050, 1560.577480, -686.034950, -0.080424, -70.003566),
    2005: (3099.470000, 3001.309140, 98.160860, 2232.612520, -768.696620, -0.061385, -71.175613),
    2006: (4908.425600, 5004.574066, -96.148466, 2575.682745, -2428.891321, -0.129567, -205.838248),
    2007: (8050.072613, 7380.413233, 669.659380, 4151.652745, -3228.760488, -0.113847, -283.224604),
    2008: (19552.787089, 12459.899292, 7092.887797, 14244.587003, 1784.687712, 0.044465, 116.426680),
    2009: (14353.424820, 12157.876094, 2195.548726, 801.794820, -11356.081274, -0.318978, -2185.908516),
    2010: (19264.485079, 13485.900667, 5778.584412, 9055.462479, -4430.438188, -0.104332, -555.582772),
    2011: (20248.821182, 17518.964433, 2729.856749, 4568.590782, -12950.373651, -0.248654, -1404.533536),
    2012: (22009.294168, 20027.811313, 1981.482855, 6543.683168, -13484.128145, -0.246314, -1416.035745),
    2013: (22850.829910, 21931.788059, 919.041851, 6276.657310, -15655.130749, -0.269072, -1538.677950),
    2014: (23002.557424, 21384.053715, 1618.503709, 6594.957424, -14789.096291, -0.250603, -2146.974580),
    2015: (18415.966481, 21014.676792, -2598.710311, 3597.654281, -17417.022511, -0.320284, -0.459305),
    2016: (20697.209500, 21429.566281, -732.356781, 11366.026700, -10063.539581, -0.166545, -0.251740),
    2017: (24076.000000, 25183.000000, -1107.000000, 12881.855905, -12301.144095, -0.174887, -0.277604),
    2018: (30924.627409, 26457.034300, 4467.593109, 10869.091709, -15587.942591, -0.194625, -0.333770),
    2019: (33966.674500, 26506.080000, 7460.594500, 12695.761700, -13810.318300, -0.168632, -0.272661),
    2020: (24465.197400, 29160.100000, -4694.902600, 13302.320000, -15268.680000, -0.210376, -0.298616),
    2021: (33940.081300, 29990.477900, 3949.603400, 15768.109000, -13596.368900, -0.145879, -0.236737),
    2022: (43097.699867, 35073.672600, 8024.027267, 18824.521753, -15616.450847, -0.116564, -0.223846),
    2023: (50048.954495, 40333.328200, 9715.626295, 21968.272114, -17496.438502, -0.142099, -0.221501),
    2024: (47965.445061, 42840.400921, 5125.044140, 25145.958513, -17248.873988, -0.136531, -0.201241),
}
MOE_FISCAL_XLSX = os.environ.get("MOE_FISCAL_XLSX", "")   # bax: MOE_AGRI_XLSX qeydi


def moe_icmal_budget() -> pd.DataFrame:
    """İcmal büdcənin faktiki tarixçəsi (2001–2024, mln AZN) — DƏSTƏKLƏYİCİ SÜBUT.

    `nonoil_base_balance()`-in dövlət büdcəsi analoqunun yanında çap olunur ki, iki perimetr
    arasındakı fərq görünsün. Provenans `MOE_ICMAL_SOURCE_NOTE`-dadır.
    """
    cols = ["revenue_total", "expenditure_total", "balance", "revenue_nonoil",
            "nonoil_base_balance", "ratio_gdp", "ratio_nonoil_gdp"]
    df = pd.DataFrame.from_dict(MOE_ICMAL_BUDGET, orient="index", columns=cols).sort_index()
    df["ratio_gdp"] = df["ratio_gdp"] * 100.0                   # faylda pay, burada faiz
    df["ratio_nonoil_gdp"] = df["ratio_nonoil_gdp"] * 100.0
    df.loc[df.index < MOE_ICMAL_RATIO_VALID_FROM, "ratio_nonoil_gdp"] = float("nan")
    return df


def verify_moe_icmal_budget(path: str = MOE_FISCAL_XLSX, tol: float = 1e-3):
    """Sabit cədvəli MƏNBƏDƏN yenidən oxuyur (etiket yoxlaması ilə). Fayl yoxdursa None."""
    if not path or not os.path.exists(path):
        return None
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["İcmal büdcə"]
    rows = list(ws.iter_rows(min_row=1, max_row=91, max_col=47, values_only=True))
    hdr = rows[2]
    ycol = {hdr[i]: i for i in range(47) if isinstance(hdr[i], int) and 1990 < hdr[i] < 2040}
    got = {}
    for r, label in MOE_ICMAL_ROWS:
        found = str(rows[r - 1][1] or "").strip()
        if _key(found)[:40] != _key(label)[:40]:
            raise AssertionError(
                f"ETİKET UYĞUNSUZLUĞU — MOE FISCAL İcmal büdcə!{r}: gözlənilən {label!r}, "
                f"tapılan {found!r}")
        got[r] = {y: rows[r - 1][c] for y, c in ycol.items()
                  if isinstance(rows[r - 1][c], (int, float)) and y <= 2024}
    wb.close()
    diff = 0.0
    for y, vals in MOE_ICMAL_BUDGET.items():
        for k, (r, _) in enumerate(MOE_ICMAL_ROWS):
            if y in got[r]:
                diff = max(diff, abs(float(got[r][y]) - float(vals[k])))
    if diff > tol:
        raise AssertionError(f"MOE İcmal büdcə mənbədən fərqlənir: maks. fərq {diff:.6f}")
    return diff


# --- MOE SNA.xlsx `Investisiya` — investisiya deflyatorları --------------------
# `DEFID` (Nazirliyin öz `eq_defid` tənliyinin asılı dəyişəni) MOE SNA `Investisiya` vərəqinin
# 45-ci sətridir: "Daxili investisiyaların deflyatoru", ingilis etiketi "deflator, domestic
# investment, defid", indeks 2015 = 1, 1998–2024. 49-cu sətir xarici investisiyanın deflyatorudur.
# Hər iki sıra YALNIZ oxunur və yalnız faktiki dövr (≤2024) saxlanılır.
MOE_INV_DEFL_SOURCE_NOTE = (
    "MOE SNA.xlsx, `Investisiya` vərəqi, sətir 45 (DEFID — daxili investisiyanın deflyatoru, "
    "indeks 2015=1) və 49 (xarici investisiyanın deflyatoru); Nazirliyin nümunə-model faylı — "
    "YALNIZ OXUNUR, yalnız faktiki sütunlar (≤2024)")
MOE_INV_DEFL_ROWS = [(45, "Daxili investisiyaların deflyatoru"),
                     (49, "Xarici investisiyaların deflyatoru")]
MOE_INV_DEFL = {          # il -> (defid, defif), indeks 2015 = 1
    1998: (0.9261234, 0.9119409),
    1999: (0.9261234, 0.9119409),
    2000: (0.9257940, 0.9116206),
    2001: (0.9257955, 0.9118791),
    2002: (0.9261127, 0.9245871),
    2003: (0.9263393, 0.9243700),
    2004: (0.9263281, 0.9240799),
    2005: (0.9262393, 0.9244340),
    2006: (0.9264986, 0.9248929),
    2007: (0.9263002, 0.9253317),
    2008: (0.9354756, 0.9350318),
    2009: (0.8907620, 0.8898471),
    2010: (0.9422769, 0.9412516),
    2011: (0.9563503, 0.9557995),
    2012: (0.9711356, 0.9708181),
    2013: (0.9772860, 0.9773552),
    2014: (0.9811505, 0.9816477),
    2015: (1.0000000, 1.0000000),
    2016: (1.0228173, 1.5546931),
    2017: (1.0955166, 1.6758383),
    2018: (1.1329513, 1.7338345),
    2019: (1.1548598, 1.7661250),
    2020: (1.1573505, 1.7697038),
    2021: (1.1831368, 1.8093757),
    2022: (1.2172989, 1.8609178),
    2023: (1.2598342, 1.9251652),
    2024: (1.2761704, 1.9503714),
}
MOE_SNA_XLSX = os.environ.get("MOE_SNA_XLSX", "")          # bax: MOE_AGRI_XLSX qeydi
MOE_EQ_DEFID_TEXT = (
    'DLOG(DEFID) = -0.0662721611849*@BEFORE("2003") + -0.00255232925979*@DURING("2003 2009") '
    '+ 0.0344790916665*@AFTER("2010") + -0.374221077246*DLOG(DEFID(-1)) '
    '+ -0.101836275949*DLOG(NEER(-1))   [MOE SNA.xlsx, `eq` vərəqi, sətir 32, `eq_defid`]')


def moe_investment_deflators() -> pd.DataFrame:
    """Nazirliyin investisiya deflyatorları (1998–2024, indeks 2015 = 1).

    Sütunlar `defid` (daxili) və `defif` (xarici). İnvestisiyanın ayrıca qiymət indeksinin hədəf sırasıdır (onların
    `eq_defid` tənliyi bizim məlumat qatında yenidən qiymətləndirilir).
    """
    return pd.DataFrame.from_dict(MOE_INV_DEFL, orient="index",
                                 columns=["defid", "defif"]).sort_index()


def verify_moe_investment_deflators(path: str = MOE_SNA_XLSX, tol: float = 1e-5):
    """Sabit deflyator cədvəlini MƏNBƏDƏN yenidən oxuyur (etiket yoxlaması ilə)."""
    if not path or not os.path.exists(path):
        return None
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Investisiya"]
    rows = list(ws.iter_rows(min_row=1, max_row=110, max_col=50, values_only=True))
    hdr = rows[2]
    ycol = {hdr[i]: i for i in range(50) if isinstance(hdr[i], int) and 1990 < hdr[i] < 2040}
    got = {}
    for r, label in MOE_INV_DEFL_ROWS:
        found = str(rows[r - 1][1] or "").strip()
        if _key(found) != _key(label):
            raise AssertionError(
                f"ETİKET UYĞUNSUZLUĞU — MOE SNA Investisiya!{r}: gözlənilən {label!r}, "
                f"tapılan {found!r}")
        got[r] = {y: rows[r - 1][c] for y, c in ycol.items()
                  if isinstance(rows[r - 1][c], (int, float)) and y <= 2024}
    wb.close()
    diff = 0.0
    for y, (a, b) in MOE_INV_DEFL.items():
        if y in got[45]:
            diff = max(diff, abs(float(got[45][y]) - a))
        if y in got[49]:
            diff = max(diff, abs(float(got[49][y]) - b))
    if diff > tol:
        raise AssertionError(f"MOE investisiya deflyatorları mənbədən fərqlənir: {diff:.8f}")
    return diff


def _implied_deflator_growth(nominal_code: str, real_index_code: str) -> pd.Series:
    """İki qonşu iş kitabı sətrindən (nominal səviyyə + "əvvəlki il = 100" real indeks)
    implisit deflyatorun illik artımı, %."""
    lvl = wb_labelled(nominal_code); lvl = lvl[lvl.index <= LAST_ACTUAL]
    idx = wb_labelled(real_index_code); idx = idx[idx.index <= LAST_ACTUAL]
    gn = (lvl / lvl.shift(1) - 1.0) * 100.0
    gr = idx - 100.0
    return (((1 + gn / 100.0) / (1 + gr.reindex(gn.index) / 100.0) - 1.0) * 100.0).dropna()


# Əsas kapitala CƏMİ investisiyanın implisit deflyatoru 2006-cı ilə qədər etibarsızdır:
# `inv_total` (Real sektor r80) 2000–2005 illərində komponent sıralarından fərqli buraxılışa
# aiddir (FR06, Bölmə 2-də sənədləşdirilib) və implisit deflyator həmin buraxılış fərqini
# qiymət hərəkəti kimi göstərir (2006: −42 %). Ona görə sıra 2008-dən verilir — 2008 həm də
# MOE DEFID sırasının hərəkətə başladığı ildir.
INV_DEFLATOR_FIRST_YEAR = 2008


def investment_deflator_growth(first_year: int = INV_DEFLATOR_FIRST_YEAR) -> pd.Series:
    """Əsas kapitala cəmi investisiyanın implisit deflyatorunun illik artımı, % (bizim qat).

    `real_r080` (nominal səviyyə) və `real_r081` (əvvəlki il = 100 real indeks) qonşu
    sətirlərindən qurulur. 2016-cı il dəyəri (26,24 %) MOE SNA `Investisiya` r7-nin öz
    "deflyator" sətri ilə (26,23964) beş onluq rəqəmə qədər üst-üstə düşür — bu, investisiya deflyatorunun
    "ayrıca investisiya deflyatoru yoxdur" boşluğunu bizim öz qatımızda bağlayan çarpaz
    yoxlamadır.
    """
    s = _implied_deflator_growth("real_r080", "real_r081")
    return s[s.index >= int(first_year)]


def construction_works_deflator_growth() -> pd.Series:
    """Tikinti-quraşdırma işlərinin implisit deflyatorunun artımı, % (`real_r109` / `real_r110`).

    Cəmi investisiyanın deflyatorunun əsas komponentidir: işlərin cəmi investisiyadakı payı
    2008–2025-ci illərdə 0,60–0,76 aralığındadır (`construction_works_share()`).
    """
    return _implied_deflator_growth("real_r109", "real_r110")


def construction_works_share() -> pd.Series:
    """Tikinti-quraşdırma işlərinin cəmi investisiyadakı payı (0–1) — çəkili düsturun `w_c` çəkisi."""
    w = wb_labelled("real_r109") / wb_labelled("real_r080")
    return w[w.index <= LAST_ACTUAL].dropna()


def soe_investment_annual() -> pd.Series:
    """Müəssisə və təşkilatların ÖZ VƏSAİTLƏRİ ilə əsas kapitala investisiya, mln AZN.

    dövlət investisiyası tənliyinin ikinci termi üçün nəzərdə tutulan sıra (`real_r114`). **Etiket dəqiqliyi:** sıra
    yalnız dövlət müəssisələrini deyil, BÜTÜN müəssisələri əhatə edir — maliyyələşmə mənbəyi
    üzrə bölgüdür, mülkiyyət üzrə deyil. İllik sütunu yalnız 2023-cü ildən doludur (2021–2022
    aylıq sətirlər natamam olduğu üçün illik yazı yoxdur), ona görə n = 3: D5-ə görə bu sıranın
    üzərində elastiklik QİYMƏTLƏNDİRİLMİR, o yalnız lövbər/çarpaz yoxlamadır.
    """
    s = wb_labelled("real_r114")
    return s[s.index <= LAST_ACTUAL].dropna()


def own_funds_investment_annual(tol: float = 1e-6) -> pd.Series:
    """Öz vəsaitləri ilə əsas kapitala investisiya — TAM tarixçə, mln AZN (1999–2025).

    bu buraxılış. `soe_investment_annual()` (iş kitabının `real_r114` sətri) yalnız 2023–2025 illik
    dəyərlərini daşıyır və dövlət investisiyası tənliyinin ikinci termi məhz buna görə qiymətləndirilə bilmirdi.
    `DATA_GAP_NEGATIVE_PROOF.md` (x) bəndi eyni anlayışın tam tarixçəsini opendata.az-ın
    *«Maliyyə mənbələri üzrə əsas kapitala yönəldilmiş investisiyalar»* resursunda tapdı
    (`inv_own_funds`, 1999–2024).

    **Birləşdirmə qaydası və onun sübutu:** iki sıra üst-üstə düşən illərdə (2023, 2024)
    **dəqiq** eynidir (fərq ≤ `tol`), ona görə miqyaslama və ya yenidən bazalaşdırma TƏLƏB
    OLUNMUR — sıra sadəcə uzadılır: 1999–2024 açıq mənbədən, 2025 isə iş kitabından
    (açıq mənbənin son ili 2024-dür). Uyğunluq hər çağırışda YOXLANILIR; pozularsa icra dayanır.
    """
    hist = public_series("inv_own_funds")
    wb = soe_investment_annual()
    common = sorted(set(hist.index) & set(wb.index))
    if common:
        diff = float((hist.loc[common] - wb.loc[common]).abs().max())
        if diff > tol:
            raise AssertionError(
                "öz vəsait sıraları üst-üstə düşən illərdə fərqlənir: "
                f"maks. |fərq| = {diff:.6f} mln AZN ({common[0]}–{common[-1]})")
    tail = wb[wb.index > (int(hist.index.max()) if len(hist) else -1)]
    out = pd.concat([hist, tail]).sort_index()
    return out[out.index <= LAST_ACTUAL].rename("inv_own_funds")



# =========================================================================================
# 11. MOE SOCIAL.xlsx + MOE FISCAL.xlsx `Pension Fund` — FƏALİYYƏT NÖVLƏRİ ÜZRƏ ƏMƏK BLOKU
#     (əmək haqqı → gəlir kanalı; FR08)
# =========================================================================================
# Nazirliyin NÜMUNƏ-MODEL faylları. Konvensiya yuxarıdakı MOE bloklarının eynisidir: sıralar
# burada SABİT kimi saxlanılır (təhvil paketi paketdən kənar fayla asılı olmamalıdır), provenans
# `*_SOURCE_NOTE`-dadır, `verify_*` funksiyaları fayl əlçatan olduqda etiket yoxlaması ilə
# mənbədən yenidən qurub tutuşdurur.
#
# GUARDRAİL: hər iki faylın illik başlıq sətrinin üstündəki annotasiya sətri
# **2025-i "gözlənilən", 2026–2030-u "proqnoz"** kimi işarələyir, ona görə BURADA YALNIZ ≤2024
# faktiki sütunlar saxlanılır. `model` və `add factor` sətirləri (MOE SOCIAL `7_10` r30/r32/r33)
# nümunə modelin öz proqnozu və mülahizə düzəlişidir — MƏLUMAT DEYİL və oxunmur.
#
# NƏ ÜÇÜN LAZIMDIR: paketin öz iş kitabı fəaliyyət növləri üzrə əmək haqqı və məşğulluğu yalnız
# dörd sənaye vərəqi üçün (2016-dan) və DVX kəsikləri üçün (2021-dən) daşıyır. MOE SOCIAL 8 və 9
# vərəqləri həmin bölgünü **19 fəaliyyət növü üzrə 1999–2024** verir; `add8` isə həmin bölgü üzrə
# əmək haqqı fondunu (2000–2024). Fond sətirləri MUZDLU işçilərin fonduna aiddir (ümumi
# məşğulluğa deyil) — bu, blokun `moe_branch_employees()` törəməsinin bütün mənasıdır.

MOE_SOCIAL_SOURCE_NOTE = (
    "MOE SOCIAL.xlsx — `8` vərəqi sətir 7 (iqtisadiyyat üzrə orta aylıq əmək haqqı), 9–45 (19 "
    "fəaliyyət növü, tək sətirlər), 48/50 (dövlət/qeyri-dövlət), 58/60 (neft/qeyri-neft), 62 "
    "(minimum əmək haqqı); `9` vərəqi eyni sətir sxemi ilə muzdlu/ümumi məşğulluq; `add8` vərəqi "
    "sətir 29 (iqtisadiyyat üzrə əmək haqqı fondu, AZN mln) və 30–48 (19 fəaliyyət növü üzrə "
    "fond). Nazirliyin nümunə-model faylı — YALNIZ OXUNUR, yalnız faktiki sütunlar (≤2024; "
    "annotasiya sətri 2025-i 'gözlənilən', 2026–2030-u 'proqnoz' kimi işarələyir)")
MOE_SSPF_SOURCE_NOTE = (
    "MOE FISCAL.xlsx, `Pension Fund` vərəqi (Dövlət Sosial Müdafiə Fondunun büdcəsi, mln AZN): "
    "sətir 4 (gəlirlər), 10 (dövlət büdcəsindən transfertlər), 12 (xərclər), 13 (əhaliyə "
    "ödənişlər), 14 (əmək pensiyaları). Nazirliyin nümunə-model faylı — YALNIZ OXUNUR, yalnız "
    "faktiki sütunlar (≤2024)")

# fəaliyyət növü kodu -> (AZ etiket, `8`/`9` vərəqlərinin sətri, `add8` fond sətri)
MOE_SOCIAL_BRANCHES = [
    ("agr", "Kənd, meşə və balıqçılıq təsərrüfatı", 9, 30),
    ("mine", "Mədənçıxarma sənayesi", 11, 31),
    ("manu", "Emal sənayesi", 13, 32),
    ("elec", "Elektrik enerjisi, qaz və buxar istehsalı, bölüşdürülməsi və təchizatı", 15, 33),
    ("wate", "Su təchizatı; tullantıların təmizlənməsi və emalı", 17, 34),
    ("const", "Tikinti", 19, 35),
    ("trade", "Ticarət; nəqliyyat vasitələrinin təmiri", 21, 36),
    ("accom", "Turistlərin yerləşdirilməsi və ictimai iaşə", 23, 37),
    ("trans", "Nəqliyyat və anbar təsərrüfatı", 25, 38),
    ("inform", "İnformasiya və rabitə", 27, 39),
    ("finance", "Maliyyə və sığorta fəaliyyəti", 29, 40),
    ("prof", "Peşə, elmi və texniki fəaliyyət", 31, 41),
    ("admn", "İnzibati və yardımçı xidmətlərin göstərilməsi", 33, 42),
    ("rest", "Daşınmaz əmlakla əlaqədar əməliyyatlar", 35, 43),
    ("stmanag", "Dövlət idarəetmə və müdafiə; sosial təminat", 37, 44),
    ("educ", "Təhsil", 39, 45),
    ("health", "Əhaliyə səhiyyə və sosial xidmətlərin göstərilməsi", 41, 46),
    ("art", "İstirahət əyləncə və incəsənət sahəsində  fəaliyyət", 43, 47),
    ("other", "Digər sahələrdə xidmətlərin göstərilməsi", 45, 48),
]
# aqreqat sətirlər: kod -> (vərəq, sətir, gözlənilən etiket)
MOE_SOCIAL_AGG_ROWS = {
    "w_avg": ("8", 7, "Orta aylıq əmək haqqı"),
    "w_min": ("8", 62, "Minimum əmək haqqı"),
    "w_state": ("8", 48, "Dövlət"),
    "w_private": ("8", 50, "Qeyri-dövlət"),
    "w_oil": ("8", 58, "Neft"),
    "w_nonoil": ("8", 60, "Qeyri-neft"),
    "fund_total": ("add8", 29, "İqtisadiyyatda orta aylıq əmək haqqı"),
    "emp_total": ("9", 7, "Ümumi məşğulluq"),
    "emp_state": ("9", 48, "Dövlət"),
    "emp_private": ("9", 50, "Qeyri-dövlət"),
    "emp_oil": ("9", 58, "Neft (muzdlu)"),
    "emp_nonoil": ("9", 60, "Qeyri-neft"),
}
MOE_SSPF_ROWS = {
    "revenue": (4, "I. Revenues"),
    "transfer_from_budget": (10, "3. Transfers from  state budget"),
    "expenditure": (12, "II. Expenditure"),
    "payments_to_population": (13, "1. Payment to population"),
    "labour_pensions": (14, "   Payments to labor pensions"),
}
MOE_SOCIAL_XLSX = os.environ.get("MOE_SOCIAL_XLSX", "")   # bax: MOE_AGRI_XLSX qeydi

MOE_LAST_ACTUAL = 2024          # MOE iş kitablarının son FAKTİKİ ili

MOE_SOCIAL_W = {
    "agr": (1999, (12.9, 13.6, 15.1, 17.9, 23.0, 29.8, 41.6, 52.5, 86.7, 114.5, 134.3, 160.3, 196.4, 201.1, 
        217.9, 241.3, 245.8, 253.8, 261.5, 281.1, 371.4, 433.5, 456.9, 514.8, 559.0, 589.6)),
    "mine": (1999, (126.0, 162.9, 217.2, 258.1, 389.0, 451.3, 507.3, 636.8, 851.2, 1008.2, 992.8, 1004.7, 
        1180.406756, 1402.0, 1516.3, 1753.8, 2171.1, 2807.2, 3071.9, 2964.2, 3055.6, 3278.6, 3088.3, 3240.3, 
        3347.5, 3564.8)),
    "manu": (1999, (48.8, 56.9, 58.0, 69.9, 89.1, 98.3, 115.9, 141.0, 190.4, 251.7, 267.5, 320.5, 354.473348, 
        398.8, 439.3, 495.4, 527.9, 542.9, 554.0, 554.1, 632.3, 644.4, 685.2, 765.6, 844.3, 905.5)),
    "elec": (1999, (53.4, 64.3, 69.9, 75.8, 90.5, 110.4, 134.6, 164.1, 226.6, 314.9, 322.7, 349.4, 413.370692, 
        443.7, 467.2, 489.3, 513.2, 504.7, 547.8, 583.3, 638.9, 698.3, 770.0, 897.2, 959.7, 1097.2)),
    "wate": (1999, (19.7, 23.7, 26.9, 30.5, 41.3, 51.7, 64.4, 80.8, 135.6, 182.6, 189.9, 197.7, 231.854868, 
        274.8, 324.9, 331.8, 333.3, 321.4, 324.6, 304.6, 468.9, 558.3, 549.6, 647.3, 710.2, 745.3)),
    "const": (1999, (70.0, 83.8, 87.2, 105.8, 154.4, 218.2, 233.3, 293.2, 372.3, 398.4, 440.6, 505.8, 
        519.435954, 587.5, 625.5, 626.9, 677.7, 812.9, 783.3, 698.7, 737.3, 790.6, 887.9, 1005.5, 1055.0, 1086.5)),
    "trade": (1999, (23.6, 24.0, 30.0, 92.7, 96.9, 117.8, 120.0, 129.1, 173.5, 211.5, 215.2, 282.8, 335.163011, 
        343.7, 363.8, 374.0, 378.1, 382.3, 385.6, 390.3, 480.3, 521.4, 536.8, 573.3, 631.2, 691.3)),
    "accom": (1999, (68.7, 62.0, 72.9, 96.8, 101.7, 112.6, 157.9, 164.6, 212.3, 265.4, 297.9, 333.7, 385.0, 
        404.6, 444.6, 463.9, 464.6, 476.1, 511.2, 541.7, 568.3, 537.7, 584.9, 664.0, 720.5, 749.2)),
    "trans": (1999, (48.0, 51.4, 58.4, 66.1, 78.1, 98.8, 124.6, 157.2, 226.6, 300.0, 349.2, 395.1, 446.849569, 
        511.5, 536.3, 530.3, 575.8, 649.3, 733.9, 806.9, 855.8, 886.0, 953.0, 1124.3, 1278.0, 1395.1)),
    "inform": (1999, (74.0, 86.6, 96.2, 106.9, 171.1, 188.0, 204.8, 223.0, 364.2, 465.6, 496.7, 531.3, 
        576.620311, 621.9, 675.4, 735.0, 747.2, 782.0, 870.6, 881.3, 1005.3, 1082.6, 1155.4, 1342.4, 1512.8, 
        1670.8)),
    "finance": (1999, (118.2, 126.5, 157.4, 209.6, 196.0, 217.3, 311.7, 535.5, 707.8, 785.4, 812.8, 990.2, 
        1004.497406, 1055.5, 1126.4, 1198.8, 1210.9, 1229.2, 1387.7, 1459.3, 1607.3, 1726.4, 1937.6, 2066.8, 
        2325.5, 2682.4)),
    "prof": (1999, (99.5, 115.1, 137.0, 162.8, 202.5, 256.3, 324.4, 410.5, 556.2, 596.1, 585.3, 592.2, 600.0, 
        620.7, 667.1, 669.7, 752.4, 886.5, 1032.7, 1054.7, 1180.3, 1188.4, 1208.9, 1407.6, 1573.5, 1710.7)),
    "admn": (1999, (28.3, 32.8, 47.6, 68.9, 373.9, 410.8, 448.5, 489.8, 534.2, 477.9, 456.9, 526.7, 535.105442, 
        563.2, 583.9, 566.6, 542.6, 543.6, 547.8, 548.3, 386.1, 426.2, 431.6, 486.8, 563.9, 620.0)),
    "rest": (1999, (30.6, 32.2, 33.2, 34.2, 34.7, 40.8, 56.3, 77.7, 104.9, 147.2, 175.8, 168.1, 228.288597, 
        255.6, 293.9, 308.2, 308.2, 337.0, 354.6, 422.5, 558.2, 610.5, 660.0, 789.0, 910.1, 1073.3)),
    "stmanag": (1999, (36.3, 38.9, 42.2, 51.0, 63.0, 86.8, 133.9, 157.8, 209.1, 287.0, 350.1, 376.5, 402.7, 
        452.6, 455.0, 479.6, 494.7, 510.6, 534.5, 596.0, 799.9, 957.2, 985.3, 1213.3, 1356.4, 1417.6)),
    "educ": (1999, (28.3, 31.2, 31.5, 33.8, 42.3, 51.3, 66.0, 79.6, 145.4, 214.4, 260.0, 271.8, 283.410277, 
        287.3, 293.6, 298.0, 301.1, 307.0, 322.2, 365.3, 438.9, 515.7, 526.1, 630.2, 713.3, 754.5)),
    "health": (1999, (14.2, 14.7, 14.8, 18.0, 21.9, 29.6, 45.2, 69.6, 94.2, 130.1, 154.1, 155.2, 163.957087, 
        175.1, 181.6, 197.8, 204.2, 214.8, 222.5, 239.4, 352.4, 524.7, 568.2, 695.0, 804.1, 848.5)),
    "art": (1999, (17.2, 18.4, 20.0, 21.8, 33.8, 41.8, 51.4, 63.2, 103.8, 143.5, 204.2, 208.4, 210.976398, 
        211.3, 220.6, 250.1, 252.7, 258.9, 268.4, 296.0, 423.8, 529.7, 535.7, 622.2, 726.1, 747.0)),
    "other": (1999, (41.2, 46.6, 50.7, 55.2, 47.3, 59.9, 75.8, 96.0, 155.3, 227.8, 250.5, 280.3, 331.833671, 
        367.6, 377.6, 400.2, 441.9, 473.4, 662.6, 631.8, 631.8, 678.8, 758.5, 931.1, 1104.7, 1308.4)),
}

MOE_SOCIAL_L = {
    "agr": (2001, (1521.7, 1530.4, 1546.1, 1551.6, 1573.6, 1583.2, 1597.6, 1611.3, 1628.6, 1655.0, 1657.4, 
        1673.8, 1677.4, 1691.7, 1698.4, 1729.6, 1752.9, 1769.3, 1720.4, 1696.5, 1732.9, 1753.1, 1777.4, 1782.3)),
    "mine": (1999, (40.5, 41.2, 38.8, 39.3, 41.9, 42.1, 42.3, 44.1, 44.3, 44.2, 42.7, 41.5, 41.2, 41.8, 42.3, 
        41.5, 39.1, 38.1, 37.9, 40.3, 40.3, 40.1, 39.4, 39.3, 39.7, 39.7)),
    "manu": (1999, (184.5, 187.3, 188.0, 191.1, 192.3, 198.5, 198.4, 203.4, 206.7, 211.4, 214.2, 208.9, 210.3, 
        215.6, 224.1, 227.1, 229.8, 242.2, 249.1, 254.8, 259.5, 255.3, 266.0, 276.3, 285.0, 293.8)),
    "elec": (1999, (20.3, 25.2, 25.9, 26.7, 22.7, 25.2, 27.9, 30.9, 34.6, 34.9, 30.0, 30.6, 30.8, 31.2, 32.3, 
        29.8, 27.1, 27.3, 27.5, 27.5, 30.5, 28.3, 27.2, 27.9, 29.1, 29.7)),
    "wate": (1999, (19.4, 24.0, 24.7, 24.9, 22.8, 23.5, 23.9, 24.3, 25.4, 27.5, 25.3, 25.2, 24.6, 24.7, 25.7, 
        25.8, 25.4, 30.6, 30.8, 41.1, 30.2, 29.9, 32.7, 33.6, 34.6, 42.2)),
    "const": (1999, (158.1, 189.0, 191.0, 195.4, 202.8, 208.4, 211.9, 217.0, 222.5, 226.0, 229.0, 287.5, 308.9, 
        321.8, 325.5, 334.1, 336.4, 343.8, 347.9, 354.5, 363.6, 358.0, 373.9, 379.7, 388.6, 403.8)),
    "trade": (1999, (606.9, 612.2, 617.2, 617.8, 629.1, 630.9, 634.8, 649.2, 656.5, 670.0, 678.9, 626.7, 635.4, 
        646.8, 664.0, 681.9, 693.7, 699.6, 705.9, 706.1, 699.4, 682.3, 695.2, 702.2, 707.2, 718.4)),
    "accom": (1999, (10.0, 20.2, 20.7, 21.2, 21.9, 24.0, 25.1, 25.3, 25.7, 26.0, 26.3, 46.9, 48.1, 48.9, 49.2, 
        55.7, 61.5, 68.4, 73.5, 78.2, 77.5, 65.5, 76.9, 87.9, 95.1, 103.6)),
    "trans": (1999, (146.7, 165.7, 167.9, 169.6, 170.8, 172.6, 174.6, 176.7, 178.9, 181.2, 183.6, 179.1, 181.8, 
        182.7, 183.8, 185.1, 197.1, 198.4, 201.0, 203.2, 198.6, 195.6, 199.7, 199.8, 202.7, 207.6)),
    "inform": (1999, (25.3, 28.5, 29.1, 29.5, 30.3, 31.1, 32.3, 32.7, 33.2, 33.5, 34.0, 55.8, 58.0, 58.7, 58.1, 
        59.2, 60.3, 61.2, 61.7, 62.8, 60.3, 58.6, 59.6, 60.2, 59.5, 60.2)),
    "finance": (1999, (15.5, 16.0, 15.9, 16.6, 16.5, 17.0, 18.1, 20.9, 21.2, 21.5, 21.8, 24.4, 26.3, 26.9, 
        30.6, 32.8, 33.0, 27.1, 26.9, 27.4, 28.0, 28.9, 31.9, 35.1, 37.2, 39.6)),
    "prof": (1999, (25.6, 40.4, 40.9, 41.5, 41.9, 42.9, 43.4, 43.9, 44.4, 45.0, 45.6, 45.6, 46.7, 54.6, 56.3, 
        58.5, 59.6, 68.4, 73.5, 74.3, 62.4, 61.1, 61.6, 62.9, 66.7, 70.7)),
    "admn": (1999, (16.8, 26.7, 28.1, 28.6, 35.4, 33.8, 38.7, 38.0, 34.9, 33.8, 38.7, 46.5, 47.4, 49.2, 52.4, 
        53.7, 55.2, 57.1, 58.0, 59.5, 70.4, 90.3, 94.9, 97.1, 96.8, 94.7)),
    "rest": (1999, (100.8, 78.3, 78.5, 81.9, 82.6, 83.8, 82.1, 83.0, 84.1, 85.2, 86.3, 69.6, 71.2, 74.8, 79.4, 
        85.6, 89.7, 88.0, 88.8, 90.8, 87.5, 84.4, 85.2, 85.8, 84.9, 86.0)),
    "stmanag": (1999, (265.8, 244.7, 246.9, 247.8, 248.9, 253.7, 256.6, 259.6, 262.9, 266.3, 269.8, 279.1, 
        281.0, 281.7, 282.3, 285.2, 287.3, 285.4, 284.2, 281.5, 242.6, 238.1, 236.7, 234.9, 230.5, 228.4)),
    "educ": (1999, (319.1, 330.9, 336.3, 345.5, 343.6, 348.8, 345.1, 342.8, 348.9, 355.2, 361.0, 349.8, 349.9, 
        349.0, 366.2, 367.3, 373.5, 374.8, 377.8, 380.2, 370.6, 369.9, 373.8, 375.5, 376.1, 371.8)),
    "health": (1999, (171.8, 179.4, 182.1, 183.6, 184.7, 186.4, 188.8, 189.1, 192.6, 192.4, 201.9, 170.3, 
        165.2, 165.4, 171.8, 176.5, 180.8, 185.6, 189.0, 190.0, 188.2, 187.1, 186.3, 188.7, 189.7, 190.3)),
    "art": (1999, (30.8, 48.8, 49.3, 50.3, 49.7, 51.4, 52.4, 53.2, 53.2, 54.3, 56.9, 59.6, 60.3, 61.1, 61.8, 
        67.8, 69.6, 77.4, 80.6, 81.5, 66.6, 63.1, 64.3, 65.2, 65.4, 66.3)),
    "other": (1999, (55.3, 87.6, 88.4, 89.4, 88.6, 91.2, 92.3, 93.5, 94.6, 95.8, 97.1, 127.0, 130.7, 136.6, 
        138.0, 143.6, 154.1, 156.9, 155.1, 156.3, 189.0, 188.2, 192.9, 195.9, 197.1, 200.7)),
}

MOE_SOCIAL_FUND = {
    "agr": (2000, (14.688, 11.99544, 9.3438, 12.1716, 13.01664, 22.81344, 29.673, 47.65032, 61.9674, 70.9104, 
        79.25232, 96.6288, 102.31968, 114.26676, 131.17068, 136.56648, 143.75232, 155.0172, 173.04516, 
        241.55856, 318.3624, 337.1922, 389.1888, 395.1012, 334.65696)),
    "mine": (2000, (77.41008, 90.18144, 109.02144, 182.9856, 214.99932, 245.33028, 313.3056, 417.76896, 
        485.14584, 422.9328, 418.35708, 484.438933, 627.5352, 678.69588, 747.1188, 891.01944, 1148.70624, 
        1242.27636, 1212.95064, 1279.68528, 1341.60312, 1197.02508, 1310.37732, 1337.661, 1411.6608)),
    "manu": (2000, (97.77696, 76.3512, 89.16444, 103.28472, 113.59548, 145.3386, 186.966, 251.55648, 318.6522, 
        305.271, 326.1408, 396.017625, 467.07456, 538.75752, 611.71992, 598.00512, 656.04036, 698.7048, 
        725.42772, 924.92844, 982.83888, 1045.88928, 1191.57984, 1310.01588, 1466.91)),
    "elec": (2000, (19.44432, 21.72492, 24.28632, 24.6522, 33.38496, 45.06408, 60.84828, 94.08432, 131.88012, 
        116.172, 128.29968, 150.797628, 163.99152, 181.08672, 174.97368, 166.89264, 165.33972, 180.774, 
        192.489, 233.07072, 237.14268, 251.328, 300.38256, 335.115724, 391.04208)),
    "wate": (2000, (6.8256, 7.97316, 9.1134, 11.29968, 14.5794, 18.46992, 23.56128, 41.33088, 60.258, 57.65364, 
        52.66728, 62.87904, 75.51504, 100.19916, 102.72528, 101.58984, 118.01808, 119.97216, 150.22872, 
        169.92936, 200.31804, 215.66304, 260.99136, 294.943219, 376.52556)),
    "const": (2000, (72.10152, 57.13344, 71.98632, 102.45984, 155.00928, 172.73532, 233.9736, 317.64636, 
        388.20096, 366.93168, 483.14016, 573.457294, 786.075, 816.6528, 813.96696, 775.01772, 875.00556, 
        955.93932, 890.42328, 1058.17296, 1271.2848, 1353.1596, 1503.4236, 1433.112, 1529.3574)),
    "trade": (2000, (35.3664, 66.6, 212.69088, 274.53708, 348.59376, 360.432, 400.31328, 556.7268, 676.377, 
        692.0832, 973.28448, 1087.134744, 1143.28368, 1225.42392, 1274.592, 1292.19456, 1338.66168, 1349.75424, 
        1345.12992, 1697.3802, 1956.50136, 2045.208, 2227.61448, 2452.59072, 2735.05932)),
    "accom": (2000, (6.1752, 4.28652, 6.27264, 7.9326, 10.8096, 25.39032, 24.69, 37.95924, 51.59376, 63.27396, 
        67.27392, 82.236, 92.2488, 104.56992, 120.79956, 121.53936, 130.83228, 139.86432, 152.10936, 186.85704, 
        163.24572, 209.16024, 280.4736, 381.2886, 559.20288)),
    "trans": (2000, (48.1104, 55.78368, 61.55232, 66.5412, 86.78592, 111.09336, 142.23456, 203.12424, 285.84, 
        356.60304, 307.70388, 351.223761, 433.3428, 490.39272, 453.08832, 505.09176, 547.74948, 622.64076, 
        697.1616, 776.38176, 811.2216, 845.1204, 991.6326, 1164.0024, 1267.30884)),
    "inform": (2000, (19.12128, 21.3564, 23.60352, 55.4364, 64.9728, 67.33824, 69.8436, 110.57112, 135.21024, 
        143.0496, 154.92708, 170.91026, 194.77908, 211.53528, 231.966, 234.02304, 234.6, 260.13528, 281.31096, 
        341.39988, 387.13776, 420.10344, 513.87072, 600.88416, 687.70128)),
    "finance": (2000, (20.493, 11.3328, 16.09728, 20.2272, 25.29372, 44.13672, 89.964, 148.638, 185.66856, 
        202.87488, 249.5304, 263.981918, 312.8502, 383.87712, 444.51504, 437.37708, 368.76, 396.32712, 
        443.04348, 540.0528, 598.71552, 741.71328, 870.53616, 1038.1032, 1274.67648)),
    "prof": (2000, (38.25924, 48.498, 61.14768, 85.05, 113.18208, 146.36928, 189.1584, 275.65272, 307.5876, 
        285.15816, 308.41776, 282.96, 377.63388, 437.08392, 458.0748, 514.6416, 588.2814, 644.4048, 682.17996, 
        800.2434, 832.83072, 837.04236, 1030.3632, 1219.7772, 1395.9312)),
    "admn": (2000, (3.5424, 5.712, 9.01212, 25.12608, 35.49312, 49.5144, 68.76792, 112.82304, 137.6352, 
        115.1388, 132.09636, 151.541861, 177.07008, 170.96592, 167.94024, 160.82664, 161.12304, 180.11664, 
        188.17656, 318.30084, 389.71728, 446.44704, 509.38752, 592.77168, 584.784)),
    "rest": (2000, (1.66152, 1.75296, 1.88784, 1.95708, 2.54592, 3.64824, 4.38228, 7.17516, 10.42176, 14.13432, 
        17.54964, 28.490417, 35.57952, 44.43768, 47.70936, 55.84584, 76.0272, 80.8488, 95.823, 114.54264, 
        131.868, 131.472, 163.7964, 198.76584, 252.44016)),
    "stmanag": (2000, (20.95932, 25.52256, 31.824, 36.8928, 51.66336, 77.44776, 98.27784, 125.71092, 176.3328, 
        218.88252, 236.7432, 279.79596, 519.22272, 543.27, 578.97312, 601.95096, 659.89944, 696.5604, 778.1376, 
        1087.54404, 1286.4768, 1330.155, 1633.58712, 1813.23552, 1917.16224)),
    "educ": (2000, (119.02176, 120.9978, 132.14448, 163.29492, 203.148, 258.7464, 315.12048, 591.31272, 
        885.55776, 1082.328, 1126.23048, 1147.471529, 1158.3936, 1191.54624, 1205.8272, 1222.70688, 1241.508, 
        1288.67112, 1465.43748, 1771.22484, 2060.11836, 2106.71484, 2532.64776, 2852.05872, 2985.1038)),
    "health": (2000, (21.85596, 22.14672, 27.4968, 32.87628, 46.24704, 69.86112, 110.2464, 150.00408, 
        211.38648, 248.53248, 256.08, 265.80723, 273.99648, 285.91104, 311.89104, 327.61848, 343.33632, 
        350.838, 380.07144, 571.31088, 882.75528, 940.9392, 1183.446, 1414.57272, 1537.482)),
    "art": (2000, (10.35552, 11.304, 12.47832, 18.69816, 24.12696, 30.77832, 39.28512, 64.02384, 91.266, 
        134.52696, 134.29296, 138.231736, 140.47224, 150.62568, 173.16924, 171.93708, 174.60216, 174.88944, 
        194.2944, 280.21656, 346.4238, 348.41928, 403.93224, 467.02752, 483.1596)),
    "other": (2000, (4.7532, 5.53644, 6.35904, 5.50572, 7.33176, 9.18696, 11.52, 19.75416, 36.35688, 41.7834, 
        47.42676, 62.517464, 64.84464, 70.68672, 79.71984, 89.61732, 95.43744, 134.37528, 124.33824, 147.8412, 
        153.95184, 172.0278, 202.23492, 235.96392, 299.88528)),
}

MOE_SOCIAL_AGG = {
    "w_avg": (1995, (12.502459, 17.886846, 28.348959, 33.708008, 36.9, 44.3, 52.0, 63.1, 77.4, 99.4, 123.6, 
        149.0, 215.8, 274.4, 298.0, 331.5, 364.2, 398.4, 425.1, 444.5, 466.9, 499.8, 528.5, 544.6, 635.1, 
        707.7, 732.1, 840.0, 933.9, 1009.2)),
    "w_min": (1995, (1.075, 1.1, 1.1, 1.1, 1.1, 1.1, 5.5, 5.5, 6.666667, 16.0, 26.25, 30.0, 49.166667, 65.0, 
        75.0, 78.333333, 85.708333, 93.5, 97.333333, 105.0, 105.0, 105.0, 116.0, 130.0, 195.0, 250.0, 250.0, 
        300.0, 345.0, 345.0)),
    "w_state": (2005, (89.1, 111.4, 171.9, 235.8, 263.4, 276.4, 303.3, 332.9, 346.9, 355.7, 360.4, 375.2, 
        395.6, 435.2, 531.3, 631.4, 654.2, 795.7, 904.0, 975.2)),
    "w_private": (2005, (191.9, 221.0, 296.5, 341.3, 357.2, 423.9, 466.7, 500.3, 542.1, 571.1, 617.4, 679.0, 
        711.1, 690.7, 768.2, 801.6, 823.2, 889.0, 966.0, 1044.1)),
    "w_oil": (2005, (482.2, 624.4, 840.9, 1007.1, 991.6, 986.7, 1175.4, 1430.4, 1554.7, 1758.5, 2197.2, 2802.4, 
        3082.6, 3007.6, 3115.8, 3348.6, 3166.6, 3367.9, 3506.3, 3729.4)),
    "w_nonoil": (2005, (113.7, 136.0, 190.2, 246.0, 277.5, 312.1, 341.5, 370.1, 395.5, 411.2, 425.0, 444.8, 
        468.8, 488.4, 579.2, 651.1, 683.6, 790.4, 885.0, 959.6)),
    "fund_total": (2000, (637.92168, 666.18948, 915.48264, 1230.92916, 1564.77912, 1903.69476, 2412.13164, 
        3573.51336, 4637.33856, 4938.24084, 5499.41424, 6076.522199, 7146.22872, 7739.985, 8129.94108, 
        8404.46184, 9067.68072, 9672.11004, 10171.77852, 12540.6414, 14352.51396, 14974.78008, 17499.4662, 
        19536.991223, 21490.04988)),
    "emp_total": (1996, (3686.7, 3694.1, 3701.5, 3782.8, 3855.5, 3891.4, 3931.1, 3972.6, 4016.9, 4062.3, 
        4110.8, 4162.2, 4215.5, 4271.7, 4329.1, 4375.2, 4445.3, 4521.2, 4602.9, 4671.6, 4759.9, 4822.1, 4879.3, 
        4785.6, 4721.2, 4831.1, 4901.1, 4963.3, 5029.8)),
    "emp_state": (1998, (1710.2, 1342.0, 1278.2, 1240.0, 1192.0, 1180.0, 1209.3, 1229.8, 1271.9, 1234.6, 
        1244.4, 1149.7, 1142.7, 1143.2, 1157.7, 1169.4, 1178.2, 1176.1, 1171.4, 1158.4, 1154.9, 1156.8, 1123.2, 
        1115.3, 1075.7, 1064.4, 1053.6)),
    "emp_private": (1998, (1991.3, 2440.8, 2577.3, 2651.4, 2739.1, 2792.6, 2807.6, 2832.5, 2838.9, 2927.6, 
        2971.1, 3122.0, 3186.4, 3232.0, 3287.6, 3351.8, 3424.7, 3495.5, 3588.5, 3663.7, 3724.4, 3628.8, 3598.0, 
        3715.8, 3825.4, 3898.9, 3976.2)),
    "emp_oil": (2005, (44.3, 43.9, 42.7, 42.5, 38.8, 37.8, 36.6, 37.0, 37.2, 36.0, 34.4, 34.4, 33.9, 34.2, 
        34.7, 33.8, 31.6, 31.7, 31.4, 31.8)),
    "emp_nonoil": (2000, (3704.5, 3715.0, 3726.5, 3747.0, 3809.1, 4018.0, 4066.9, 4119.5, 4173.0, 4232.9, 
        4291.3, 4338.6, 4408.3, 4484.0, 4566.9, 4637.2, 4725.5, 4788.2, 4845.1, 4750.9, 4687.4, 4799.5, 4869.4, 
        4931.9, 4998.0)),
}

MOE_SSPF = {
    "revenue": (1995, (107.8, 158.6, 175.4, 195.6, 231.4, 268.0, 269.9, 303.0, 361.6, 394.1, 495.7, 562.0, 
        1078.91945, 1403.618562, 1615.848813, 1755.334107, 2160.275, 2475.7, 2676.9336, 2894.3, 2912.4, 
        3175.74599, 3398.8925, 3747.85, 4116.7, 4748.3, 5178.0, 5729.7, 6675.2, 7020.4)),
    "transfer_from_budget": (1995, (34.8, 68.0, 55.8, 68.0, 94.0, 104.8, 103.0, 109.0, 153.6, 137.0, 183.0, 
        167.0, 278.74, 435.0, 560.0, 646.7048, 886.0, 1044.3, 1077.0336, 1142.0, 1100.0, 1246.0, 1303.0, 
        1337.3, 1182.7, 1177.9, 1307.4, 1131.967, 1467.77, 1232.0)),
    "expenditure": (1995, (107.8, 158.6, 175.4, 195.6, 231.4, 268.0, 269.9, 310.0, 361.6, 394.0, 496.8, 557.9, 
        1004.324178, 1392.670316, 1699.524424, 1768.892852, 2125.7, 2458.9, 2588.1848, 2868.768, 2934.2, 
        3194.8, 3464.01339, 3752.186, 3948.83, 4711.1, 4843.0, 5252.1, 6080.8, 6888.8)),
    "payments_to_population": (2000, (266.9, 261.9, 293.8, 349.7, 380.6, 467.8, 529.4, 954.166382, 1339.398576, 
        1616.942847, 1704.383448, 2066.7, 2403.4, 2530.5, 2809.1, 2872.0, 3194.8, 3384.9608, 3673.53, 3843.14, 
        4606.08, 4737.98, 5145.999497, 5975.78, 6650.52)),
    "labour_pensions": (2005, (450.012, 501.298, 926.422641, 1302.270305, 1568.710242, 1654.881731, 2010.8, 
        2344.8, 2468.2, 2738.5, 2797.3, 3042.84226, 3288.7034, 3593.03, 3747.48, 4505.13, 4613.23, 4951.0, 
        5780.780503, 6468.94)),
}



def _moe_table(const: dict, name: str) -> pd.DataFrame:
    """`{ad: (ilk_il, (dəyərlər...))}` sabitini illik DataFrame-ə çevirir (≤2024)."""
    cols = {}
    for k, (y0, vals) in const.items():
        cols[k] = pd.Series({int(y0) + i: (float(v) if v is not None else float("nan"))
                             for i, v in enumerate(vals)})
    df = pd.DataFrame(cols).sort_index()
    if not df.empty and int(df.index.max()) > MOE_LAST_ACTUAL:
        raise AssertionError(f"{name}: sabit cədvəl {MOE_LAST_ACTUAL}-dən sonrakı ili daşıyır")
    return df


def moe_branch_wages() -> pd.DataFrame:
    """19 fəaliyyət növü üzrə muzdlu işçilərin orta aylıq əmək haqqı, AZN/ay, 1999–2024."""
    return _moe_table(MOE_SOCIAL_W, "MOE_SOCIAL_W")


def moe_branch_employment() -> pd.DataFrame:
    """19 fəaliyyət növü üzrə məşğulluq (ümumi, muzdlu olmayanlar daxil), min nəfər, 1999–2024.

    **Etiket dəqiqliyi:** `9` vərəqinin sətirləri ÜMUMİ məşğulluqdur — kənd təsərrüfatında bu,
    əsasən öz hesabına çalışanlardır (2024: 1 782,3 min nəfər), ona görə bu sıra ilə əmək haqqı
    sırasının hasili əmək haqqı fondu DEYİL. Fond üçün `moe_branch_wage_bill()`, muzdlu işçi sayı
    üçün `moe_branch_employees()` istifadə olunur.
    """
    return _moe_table(MOE_SOCIAL_L, "MOE_SOCIAL_L")


def moe_branch_wage_bill() -> pd.DataFrame:
    """19 fəaliyyət növü üzrə əmək haqqı fondu (muzdlu işçilər), mln AZN, 2000–2024."""
    return _moe_table(MOE_SOCIAL_FUND, "MOE_SOCIAL_FUND")


def moe_social_aggregates() -> pd.DataFrame:
    """MOE SOCIAL-ın aqreqat sətirləri (orta/minimum əmək haqqı, sektor kəsikləri, fond, məşğulluq)."""
    return _moe_table(MOE_SOCIAL_AGG, "MOE_SOCIAL_AGG")


def moe_branch_employees() -> pd.DataFrame:
    """19 fəaliyyət növü üzrə MUZDLU işçilərin sayı, min nəfər, 2000–2024 — TÖRƏMƏ sıra.

        E_i(t) = fond_i(t) × 1000 / (12 × W_i(t))

    Fond (mln AZN/il) və əmək haqqı (AZN/ay) eyni vərəq ailəsindən gəldiyi üçün bölmə tərifə görə
    həmin fondun arxasındakı işçi sayını verir. Cəm 2000-ci ildə 1 217,8, 2024-də 1 778,5 min
    nəfərdir; DVX-nin öz muzdlu işçi sırası (`employees_dvx`, 2021–2025) eyni illərdə 1 807,6 və
    2 073,8 min nəfər göstərir — iki inzibati perimetr (DVX əmək müqavilələrini sayır) arasındakı
    fərq 5–14 %-dir və bu, çarpaz yoxlama kimi FR08-də çap olunur.
    """
    W = moe_branch_wages()
    F = moe_branch_wage_bill()
    cols = [c for c, *_ in [(b[0],) for b in MOE_SOCIAL_BRANCHES]]
    E = (F[cols] * 1000.0) / (12.0 * W[cols])
    return E.dropna(how="all")


def moe_payroll_fund() -> pd.Series:
    """İqtisadiyyat üzrə əmək haqqı fondu (19 sahənin cəmi), mln AZN, 2000–2024.

    `moe_social_aggregates()["fund_total"]` ilə eyniliyi `verify_moe_social()` yoxlayır.
    """
    return moe_branch_wage_bill().sum(axis=1, min_count=1).rename("moe_payroll_fund")


def moe_employees_total() -> pd.Series:
    """Muzdlu işçilərin ümumi sayı (19 sahənin cəmi), min nəfər, 2000–2024."""
    return moe_branch_employees().sum(axis=1, min_count=1).rename("moe_employees_total")


def moe_sspf() -> pd.DataFrame:
    """Dövlət Sosial Müdafiə Fondunun büdcəsi, mln AZN, 1995/2000/2005–2024.

    Əmək haqqı → gəlir kanalının TRANSFERT ayağının sürücüsüdür: ev təsərrüfatlarının aldığı cari və əsaslı transfertlərin
    (`social_r044`) ən böyük tərkib hissəsi DSMF-nin əhaliyə ödənişləridir.
    """
    return _moe_table(MOE_SSPF, "MOE_SSPF")


def moe_wage_splice_check() -> pd.DataFrame:
    """VİNTAJ TUTUŞDURMASI: MOE SOCIAL-ın aqreqat sıraları vs paketin öz iş kitabı və 8 vərəq.

    Blok yalnız SAHƏ bölgüsü üçün istifadə olunur; hər aqreqat sıra paketin öz qatından götürülür.
    Bu funksiya iki qatın üst-üstə düşdüyü illərdə fərqi ölçür və birləşdirmə (splice) qaydasını
    ədədlə əsaslandırır: fərq sıfıra bərabər olduğu üçün sahə bölgüsü aqreqatla uzlaşır və heç bir
    yenidən miqyaslama tətbiq edilmir.

    Sütunlar: `moe`, `paket` (iş kitabı), `fərq_%`; sətir indeksi il × sıra.
    """
    agg = moe_social_aggregates()
    pairs = [("w_avg", "social_r052", "orta aylıq əmək haqqı"),
             ("w_state", "social_r055", "dövlət sektorunda əmək haqqı"),
             ("w_oil", "social_r053", "neft-qaz sektorunda əmək haqqı"),
             ("w_nonoil", "social_r054", "qeyri-neft sektorunda əmək haqqı"),
             ("emp_total", "social_r059", "məşğul əhalinin sayı")]
    rows = []
    for key, code, label in pairs:
        a = agg[key].dropna()
        b = wb_annual(code)
        yrs = sorted(set(a.index) & set(b.index))
        for y in yrs:
            if b.loc[y] == 0 or pd.isna(b.loc[y]):
                continue
            rows.append({"sıra": label, "kod": code, "il": int(y), "moe": float(a.loc[y]),
                         "paket": float(b.loc[y]),
                         "fərq_%": float((a.loc[y] / b.loc[y] - 1.0) * 100.0)})
    return pd.DataFrame(rows)


def _moe_read_rows(path, sheet, rows_wanted, hdr_row=3, maxrow=120, label_col=1):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    rr = list(ws.iter_rows(min_row=1, max_row=maxrow, max_col=50, values_only=True))
    wb.close()
    hdr = rr[hdr_row - 1]
    ycol = {hdr[i]: i for i in range(50) if isinstance(hdr[i], int) and 1990 < hdr[i] < 2040}
    out = {}
    for r, label in rows_wanted:
        got = str(rr[r - 1][label_col] or "").strip()
        if _key(got) != _key(label):
            raise AssertionError(
                f"ETİKET UYĞUNSUZLUĞU — {os.path.basename(path)} {sheet}!{r}: "
                f"gözlənilən {label!r}, tapılan {got!r}")
        out[r] = {int(y): float(rr[r - 1][c]) for y, c in ycol.items()
                  if isinstance(rr[r - 1][c], (int, float)) and y <= MOE_LAST_ACTUAL}
    return out


def verify_moe_social(path: str = MOE_SOCIAL_XLSX, tol: float = 1e-6):
    """Sabit sahə cədvəllərini MƏNBƏDƏN yenidən oxuyur (19 × 3 etiket yoxlaması ilə).

    Fayl yoxdursa `None` qaytarır — təhvil mühitində fayl olmaya bilər və bu, xəta deyil.
    """
    if not path or not os.path.exists(path):
        return None
    want8 = [(r8, lab) for _, lab, r8, _ in MOE_SOCIAL_BRANCHES]
    want9 = [(r8, lab) for _, lab, r8, _ in MOE_SOCIAL_BRANCHES]
    wanta = [(rf, lab) for _, lab, _, rf in MOE_SOCIAL_BRANCHES]
    for key, (sh, r, lab) in MOE_SOCIAL_AGG_ROWS.items():
        (want8 if sh == "8" else want9 if sh == "9" else wanta).append((r, lab))
    got8 = _moe_read_rows(path, "8", want8)
    got9 = _moe_read_rows(path, "9", want9)
    gota = _moe_read_rows(path, "add8", wanta)
    diff = 0.0
    for code, lab, r_lab, r_fund in MOE_SOCIAL_BRANCHES:
        for src, const, r in ((got8, MOE_SOCIAL_W, r_lab), (got9, MOE_SOCIAL_L, r_lab),
                              (gota, MOE_SOCIAL_FUND, r_fund)):
            y0, vals = const[code]
            for i, v in enumerate(vals):
                if v is None:
                    continue
                y = int(y0) + i
                if y in src[r]:
                    diff = max(diff, abs(src[r][y] - float(v)))
    for key, (sh, r, lab) in MOE_SOCIAL_AGG_ROWS.items():
        src = {"8": got8, "9": got9, "add8": gota}[sh]
        y0, vals = MOE_SOCIAL_AGG[key]
        for i, v in enumerate(vals):
            if v is None:
                continue
            y = int(y0) + i
            if y in src[r]:
                diff = max(diff, abs(src[r][y] - float(v)))
    if diff > tol:
        raise AssertionError(f"MOE SOCIAL sahə cədvəlləri mənbədən fərqlənir: maks. fərq {diff:.8f}")
    return diff


def verify_moe_sspf(path: str = MOE_FISCAL_XLSX, tol: float = 1e-6):
    """DSMF cədvəlini MƏNBƏDƏN yenidən oxuyur (etiket yoxlaması ilə). Fayl yoxdursa None."""
    if not path or not os.path.exists(path):
        return None
    got = _moe_read_rows(path, "Pension Fund",
                         [(r, lab) for r, lab in MOE_SSPF_ROWS.values()], maxrow=113)
    diff = 0.0
    for key, (r, _lab) in MOE_SSPF_ROWS.items():
        y0, vals = MOE_SSPF[key]
        for i, v in enumerate(vals):
            if v is None:
                continue
            y = int(y0) + i
            if y in got[r]:
                diff = max(diff, abs(got[r][y] - float(v)))
    if diff > tol:
        raise AssertionError(f"MOE DSMF cədvəli mənbədən fərqlənir: maks. fərq {diff:.8f}")
    return diff


def household_income_decomposition() -> pd.DataFrame:
    """Ev təsərrüfatlarının gəlir hesabının üç ayağı — PAKETİN ÖZ iş kitabından (əmək haqqı → gəlir kanalı).

        gəlir (`social_r035`) = əmək ödənişləri (`r041`) + transfertlər (`r044`) + digər (`r042` + `r043`)

    Bölgü DSK-nın öz hesabıdır və iş kitabında EYNİLİK kimi qapanır (maks. sapma < 1e-10 mln AZN,
    2000–2025) — `FR08` hər il üçün bunu ayrıca yoxlayır. Sütunlar: `wage_bill`, `transfers`,
    `other`, `income`, `primary` (r040).
    """
    r035, r040 = wb_annual("social_r035"), wb_annual("social_r040")
    r041, r042 = wb_annual("social_r041"), wb_annual("social_r042")
    r043, r044 = wb_annual("social_r043"), wb_annual("social_r044")
    df = pd.DataFrame({"wage_bill": r041, "transfers": r044, "other": r042 + r043,
                       "primary": r040, "income": r035}).dropna(how="all")
    df = df[df.index <= LAST_ACTUAL]
    return df.dropna(subset=["income", "wage_bill", "transfers", "other"])

def strategic_reserves_usd() -> pd.Series:
    """Strateji valyuta ehtiyatları, dövrün sonuna, **mln USD** (`money_r003`, ≤2025).

    İş kitabı mlrd dollarla verir; burada mln dollara çevrilir ki, tədiyə balansının bütün
    sıraları ilə eyni vahiddə olsun. Sıra STOK-dur — axın deyil.
    """
    s = wb_labelled("money_r003") * 1000.0
    s = s[s.index <= LAST_ACTUAL]
    s.name = "strategic_reserves_usd"
    return s


# ==========================================================================
# 12. MOE BOP.xlsx `BOP` vərəqi — tədiyə balansının BAĞLANIŞ sətirləri (FR12)
# ==========================================================================
# Paketin əsas iş kitabının `Tədiyyə Balansı` vərəqi 109-cu sətirdə bitir: onda EHTİYAT
# AKTİVLƏRİ və BALANSLAŞDIRICI MADDƏ (səhv və buraxılışlar) sətirləri YOXDUR. Nazirliyin öz
# nümunə-model iş kitabı (`MOE BOP.xlsx`, `BOP` vərəqi) isə tam BPM6 hesabatını saxlayır:
#   r4  A. CARİ ƏMƏLİYYATLAR HESABI        r32 B. KAPİTAL VƏ MALİYYƏNİN HƏRƏKƏTİ HESABI
#   r57 C. EHTİYAT AKTİVLƏRİ               r61 Ç. BALANSLAŞDIRICI MADDƏLƏR
#   r62 ÜMUMİ BALANS  (= r4 + r32 + r57 + r61)
# İşarə konvensiyası: r57 MƏNFİ olduqda ehtiyatlar ARTIR (aktivlərin alınması debetdir).
#
# GUARDRAIL: bu fayldan YALNIZ ≤2024 sütunları götürülür — 2025 "gözlənilən",
# 2026–2030 "proqnoz" kimi işarələnib və nümunə modelin öz proyeksiyasıdır, məlumat deyil.
# Fayl YALNIZ OXUNUR; dəyərlər aşağıda sabit kimi saxlanılır və `verify_moe_bop_closure()`
# fayl əlçatan olduqda etiketlərlə birlikdə yenidən yoxlayır.
#
# İKİ BURAXILIŞ FƏRQİ, gizlədilmir: (i) MOE-nin öz ÜMUMİ BALANS sətri 2010, 2011 və 2012-ci
# illərdə sıfır deyil (+31,8 / +566,5 / +157,8 mln USD) — yəni onların bu üç ili öz-özünə
# bağlanmır; (ii) 2024-cü il üçün MOE-nin maliyyə hesabı bizim iş kitabımızın sətrindən
# 1 627,7 mln USD fərqlənir (buraxılış fərqi). Ona görə FR12 ehtiyat aktivlərini MOE-nin
# rəqəmindən KÖÇÜRMÜR, öz məlumat qatının eyniliyindən TÖRƏDİR və MOE sətrini yanaşı,
# müstəqil çarpaz yoxlama kimi nəşr edir.
MOE_BOP_SOURCE_NOTE = (
    "MOE BOP.xlsx, `BOP` vərəqi: sətir 57 (C. EHTİYAT AKTİVLƏRİ / III. CHANGE IN RESERVES), "
    "58 (3.1. Mərkəzi Bank), 59 (3.2. Neft Fondu), 61 (Ç. BALANSLAŞDIRICI MADDƏLƏR / "
    "IV. ERRORS AND OMISSIONS). YALNIZ OXUNUR; yalnız ≤2024 faktiki sütunları götürülür "
    "(2025 «gözlənilən», 2026–2030 «proqnoz»). İşarə: mənfi = ehtiyatların artması.")
MOE_BOP_ROWS = [                      # (sətir, gözlənilən EN etiketi, gözlənilən AZ etiketi)
    (57, "III. CHANGE IN RESERVES", "C. EHTİYAT AKTİVLƏRİ"),
    (58, "3.1. Central Bank", "3.1. Mərkəzi Bank"),
    (59, "3.2. State Oil Fund", "3.2. Neft Fondu"),
    (61, "IV. ERRORS AND OMISSIONS", "Ç. BALANSLAŞDIRICI MADDƏLƏR"),
]
MOE_BOP_LAST_ACTUAL = 2024            # guardrail: 2025+ sütunları nümunə modelin proqnozudur
MOE_BOP_CLOSURE = {   # il -> (ehtiyat aktivləri r57, MB r58, ARDNF r59, balanslaşdırıcı r61)
    2010: (-10461.170, -1357.480, -8740.100, -988.729),
    2011: (-12356.220, -4166.160, -8385.900, -770.052),
    2012: (-4955.610, -1092.790, -3878.800, -1879.191),
    2013: (-3079.411, -2223.510, -1673.500, -2736.667),
    2014: (-4194.231, -209.560, -3489.200, -2810.254),
    2015: (11329.031, 8430.670, 2437.770, -2037.400),
    2016: (539.221, 992.398, -279.880, 3625.857),
    2017: (-1971.382, -1294.830, -554.578, 387.790),
    2018: (-3508.120, -318.762, -3331.693, 646.397),
    2019: (-5142.924, -642.473, -3598.243, 661.809),
    2020: (1987.666, -75.659, 1816.463, 648.042),
    2021: (-3081.380, -799.298, -2338.110, -153.579),
    2022: (-9375.155, -2151.337, -7755.889, -1617.763),
    2023: (-5356.361, -2378.237, -2734.627, 1320.607),
    2024: (2065.900, 741.018, -302.860, -106.686),
}
MOE_BOP_XLSX = os.environ.get("MOE_BOP_XLSX", "")     # bax: MOE_AGRI_XLSX qeydi


def moe_bop_closure() -> pd.DataFrame:
    """Nazirliyin öz tədiyə balansı iş kitabının bağlanış sətirləri, 2010–2024 (YALNIZ OXUNUR).

    Sütunlar: `reserve_assets` (r57), `reserve_cba` (r58), `reserve_sofaz` (r59),
    `errors_omissions` (r61). Provenans `MOE_BOP_SOURCE_NOTE`-dadır.
    """
    df = pd.DataFrame.from_dict(
        MOE_BOP_CLOSURE, orient="index",
        columns=["reserve_assets", "reserve_cba", "reserve_sofaz", "errors_omissions"]).sort_index()
    df.index.name = "year"
    assert int(df.index.max()) <= MOE_BOP_LAST_ACTUAL, "MOE BOP: 2024-dən sonrakı sütun götürülüb"
    return df


def verify_moe_bop_closure(path: str = MOE_BOP_XLSX, tol: float = 1e-3):
    """Könüllü çarpaz yoxlama: `MOE_BOP_CLOSURE` sabitlərini faylla tutuşdurur (etiketlə).

    Fayl yolu verilmədikdə (təhvil mühiti) `None` qaytarır — sabitlər öz-özlüyündə mənbədir.
    """
    if not path or not os.path.exists(path):
        return None
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["BOP"]
    rows = {}
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=109, max_col=54, values_only=True), 1):
        rows[i] = row
    years = {}
    for j, v in enumerate(rows[3]):
        try:
            years[j] = int(str(v).strip())
        except Exception:
            pass
    for r, en, az in MOE_BOP_ROWS:
        got_en, got_az = str(rows[r][1] or "").strip(), str(rows[r][2] or "").strip()
        if _key(got_en) != _key(en) or _key(got_az) != _key(az):
            raise AssertionError(
                f"ETİKET UYĞUNSUZLUĞU — MOE BOP `BOP`!{r}: gözlənilən {en!r}/{az!r}, "
                f"tapılan {got_en!r}/{got_az!r}")
    diff = 0.0
    for j, y in years.items():
        if y not in MOE_BOP_CLOSURE:
            continue
        for k, (r, _, _) in enumerate(MOE_BOP_ROWS):
            v = rows[r][j]
            if isinstance(v, (int, float)):
                diff = max(diff, abs(float(v) - MOE_BOP_CLOSURE[y][k]))
    if diff > tol:
        raise AssertionError(f"MOE BOP bağlanış sətirləri mənbədən fərqlənir: maks. fərq {diff:.6f}")
    return diff


# =========================================================================================
# 13. NAZİRLİK SPESİFİKASİYASI ÜÇÜN BİRLƏŞDİRİLMİŞ MOE PANELİ (FR13)
#     Nazirliyin 92 sətirlik tənlik kataloqunun tələb etdiyi sıralar, ONUN ÖZ iş
#     kitablarından, YALNIZ OXUNAN rejimdə.
# =========================================================================================
# Konvensiya §10–§12 bloklarının eynisidir, bir fərqlə: burada sıraların sayı (≈70 sıra ×
# 30 il) Python sabitləri kimi saxlanmaq üçün çox böyükdür, ona görə dəyərlər paketin öz
# giriş qovluğunda — `data/moe_spec_panel.csv` (dəyərlər) və `data/moe_spec_provenance.csv`
# (provenans) — donmuş şəkildə saxlanılır. Paket beləliklə yenə də ÖZÜNÜ-TAM qalır (paketdən
# kənar fayla asılı deyil); `verify_moe_spec_panel()` isə iş kitabları əlçatan olduqda hər
# sətri ETİKET YOXLAMASI ilə mənbədən yenidən oxuyub tutuşdurur.
#
# GUARDRAİL — üç qayda, kodda məcburi:
#   (1) yalnız ≤2024 faktiki sütunlar oxunur (annotasiya sətri 2025-i "gözlənilən",
#       2026–2030-u "proqnoz" kimi işarələyir);
#   (2) `model` və `add factor` sətirləri nümunə modelin öz proqnozu və mülahizə
#       düzəlişidir — MƏLUMAT DEYİL; spesifikasiyaya belə sətir daxil edilə bilməz
#       (`_moe_spec_assert_not_model` bunu yoxlayır);
#   (3) `PBCRHCF` 2020-dən 2030-a qədər −0.0047 səviyyəsində DONDURULUB — həmin illər
#       tarix deyil, ona görə bu sıra yalnız ≤2019 oxunur (`MOE_SPEC_YEAR_CAP`).

MOE_SPEC_SOURCE_NOTE = (
    "Nazirliyin makroekonometrik nümunə-model iş kitabları (MOE SOCIAL / MOE SNA / MOE INF / "
    "MOE BOP / INDUSTRY / MOE T&C / MOE FISCAL). YALNIZ OXUNUR; yalnız faktiki sütunlar "
    "(≤2024); `model` və `add factor` sətirləri oxunmur; hər sətir üçün gözlənilən etiket "
    "`MOE_SPEC_ROWS`-da saxlanılır və `verify_moe_spec_panel()` tərəfindən yoxlanılır")

# Yollar `config.py`-dədir (vahid mənbə): `P_MINISTRY_CATALOG` — Nazirliyin 92 sətirlik
# tənlik kataloqu; `P_MOE_SPEC_PANEL` / `P_MOE_SPEC_PROV` — donmuş panel və provenans.

# İş kitablarının olduğu qovluq — YALNIZ yoxlama/yenidən qurma üçün (bax MOE_AGRI_XLSX qeydi).
MOE_MODEL_DIR = os.environ.get("MOE_MODEL_DIR", "")

MOE_SPEC_WORKBOOKS = {
    "SOCIAL": "MOE SOCIAL.xlsx", "SNA": "MOE SNA.xlsx", "INF": "MOE INF.xlsx",
    "BOP": "MOE BOP.xlsx", "IND": "INDUSTRY.xlsx", "TC": "MOE T&C.xlsx",
    "FISCAL": "MOE FISCAL.xlsx",
}

# Sıra üzrə xüsusi il sərhədi (guardrail 3).
MOE_SPEC_YEAR_CAP = {"PBCRHC": 2019}

_MOE_SOCIAL_BRANCH_VARS = [
    ("AGR", "Kənd, meşə və balıqçılıq təsərrüfatı", 9, 8),
    ("MINE", "Mədənçıxarma sənayesi", 11, 9),
    ("MANU", "Emal sənayesi", 13, 10),
    ("ELEC", "Elektrik enerjisi, qaz və buxar istehsalı, bölüşdürülməsi və təchizatı", 15, 11),
    ("WATE", "Su təchizatı; tullantıların təmizlənməsi və emalı", 17, 12),
    ("CONST", "Tikinti", 19, 13),
    ("TRADE", "Ticarət; nəqliyyat vasitələrinin təmiri", 21, 14),
    ("ACCOM", "Turistlərin yerləşdirilməsi və ictimai iaşə", 23, 15),
    ("TRANS", "Nəqliyyat və anbar təsərrüfatı", 25, 16),
    ("INFORM", "İnformasiya və rabitə", 27, 17),
    ("FINANCE", "Maliyyə və sığorta fəaliyyəti", 29, 18),
    ("PROF", "Peşə, elmi və texniki fəaliyyət", 31, 19),
    ("ADMN", "İnzibati və yardımçı xidmətlərin göstərilməsi", 33, 20),
    ("REST", "Daşınmaz əmlakla əlaqədar əməliyyatlar", 35, 21),
    ("STMANAG", "Dövlət idarəetmə və müdafiə; sosial təminat", 37, 22),
    ("EDUC", "Təhsil", 39, 23),
    ("HEALTH", "Əhaliyə səhiyyə və sosial xidmətlərin göstərilməsi", 41, 24),
    ("ART", "İstirahət əyləncə və incəsənət sahəsində  fəaliyyət", 43, 25),
    ("OTHER", "Digər sahələrdə xidmətlərin göstərilməsi", 45, 26),
]


def _moe_spec_branch_rows():
    """19 fəaliyyət növü üzrə əmək haqqı (`8`), məşğulluq (`9`) və məhsuldarlıq (`add8`) sətirləri."""
    out = []
    for code, label, r89, rlp in _MOE_SOCIAL_BRANCH_VARS:
        out.append((f"W_{code}", "SOCIAL", "8", r89, label, "AZN/ay",
                    "fəaliyyət növü üzrə orta aylıq əmək haqqı"))
        out.append((f"L_{code}", "SOCIAL", "9", r89, label, "min nəfər",
                    "fəaliyyət növü üzrə məşğulluq"))
        out.append((f"LP_{code}", "SOCIAL", "add8", rlp, label, "%, illik",
                    "fəaliyyət növü üzrə əmək məhsuldarlığının real artım tempi"))
    return out


# (dəyişən, iş kitabı, vərəq, sətir, gözlənilən etiket, ölçü vahidi, AZ qeyd)
MOE_SPEC_ROWS = _moe_spec_branch_rows() + [
    # --- MOE SOCIAL ---------------------------------------------------------------------
    ("W", "SOCIAL", "8", 7, "Orta aylıq əmək haqqı", "AZN/ay", "iqtisadiyyat üzrə orta aylıq əmək haqqı"),
    ("MW", "SOCIAL", "8", 62, "Minimum əmək haqqı", "AZN/ay", "minimum əmək haqqı"),
    ("NONOILWAGE", "SOCIAL", "8", 60, "Qeyri-neft", "AZN/ay", "qeyri neft-qaz sektorunda orta aylıq əmək haqqı"),
    ("L", "SOCIAL", "9", 7, "Ümumi məşğulluq", "min nəfər", "ümumi məşğulluq"),
    ("POP", "SOCIAL", "add9", 5, "İşğal olunmuş torpaqlar nəzərə alınmaqla (dövrün sonuna)", "min nəfər", "əhalinin sayı"),
    ("HI", "SOCIAL", "7_10", 6, "Gəlirlər", "mln AZN", "əhalinin gəlirləri, nominal"),
    ("SERVICES_PAID", "SOCIAL", "7_10", 11, "Əhaliyə göstərilən pullu xidmətlər", "mln AZN",
     "əhaliyə göstərilən ödənişli xidmətlər, nominal"),
    ("IEA", "SOCIAL", "7_10", 31, "sahibkarlıq fəaliyyətindən gəlirlər", "mln AZN",
     "sahibkarlıq fəaliyyətindən gəlirlər, nominal"),
    ("POPINCOME", "SOCIAL", "7_10", 50, "Gəlirlər", "mln AZN, sabit 2015",
     "əhalinin gəlirləri, 2015-ci ilin qiymətləri ilə"),
    ("RIEA", "SOCIAL", "7_10", 51, "sahibkarlıq fəaliyyətindən gəlirlər", "mln AZN, sabit 2015",
     "sahibkarlıq fəaliyyətindən gəlirlər, 2015-ci ilin qiymətləri ilə"),
    # --- MOE SNA: `ÜDM xərc metodu` ------------------------------------------------------
    ("RGDPO", "SNA", "ÜDM xərc metodu", 13, "Neft sektoru, ÜDM, əsas qiymətlərlə", "mln AZN, sabit 2015",
     "neft-qaz sektorunun ÜDM-i, sabit qiymətlərlə — `RGDPO` üçün NAMİZƏD sıra (tərif təsdiqi gözlənilir)"),
    ("FINALC_NOM", "SNA", "ÜDM xərc metodu", 20, "Son istehlak xərcləri", "mln AZN",
     "son istehlak xərcləri, nominal (`REALFINALC` törəməsi üçün)"),
    ("HC", "SNA", "ÜDM xərc metodu", 21, "Ev təsərrüfatlarının son istehlak xərcləri", "mln AZN",
     "ev təsərrüfatlarının son istehlak xərcləri, nominal"),
    ("GC", "SNA", "ÜDM xərc metodu", 22, "Dövlət idarələrinin istehlakı", "mln AZN",
     "dövlət idarələrinin istehlakı"),
    ("GFCF", "SNA", "ÜDM xərc metodu", 25, "Əsas fondların ümumi yığımı", "mln AZN",
     "əsas fondların ümumi yığımı, nominal"),
    ("CI", "SNA", "ÜDM xərc metodu", 26, "Maddi dövriyyə vəsaitlərinin dəyişməsi", "mln AZN",
     "maddi dövriyyə vəsaitlərinin dəyişməsi"),
    ("PBCRHC", "SNA", "ÜDM xərc metodu", 45, "PBCRHCF", "nisbət",
     "kataloqun `PBCRHC` adı iş kitabında `PBCRHCF` kimidir; 2020–2030 −0.0047-də DONDURULUB, "
     "ona görə yalnız ≤2019 oxunur"),
    ("RHC", "SNA", "ÜDM xərc metodu", 49, "Ev təsərrüfatlarında", "mln AZN, sabit 2015",
     "ev təsərrüfatlarının son istehlakı, 2015-ci ilin qiymətləri ilə"),
    # --- MOE SNA: `ÜDM sahələr üzrə` -----------------------------------------------------
    ("NONOILIND", "SNA", "ÜDM sahələr üzrə", 20, "əlavə dəyər", "mln AZN", "qeyri-neft sənayesi, əlavə dəyər"),
    ("VA_MANU", "SNA", "ÜDM sahələr üzrə", 63, "əlavə dəyər", "mln AZN", "emal sənayesi, əlavə dəyər"),
    ("AGRI", "SNA", "ÜDM sahələr üzrə", 262, "əlavə dəyər", "mln AZN", "kənd təsərrüfatı, əlavə dəyər"),
    ("VA_CONST", "SNA", "ÜDM sahələr üzrə", 271, "əlavə dəyər", "mln AZN", "tikinti, əlavə dəyər"),
    ("TRADE", "SNA", "ÜDM sahələr üzrə", 289, "əlavə dəyər", "mln AZN", "ticarət, əlavə dəyər"),
    ("VA_INFORM", "SNA", "ÜDM sahələr üzrə", 320, "əlavə dəyər", "mln AZN", "rabitə, əlavə dəyər"),
    ("VA_DIGER", "SNA", "ÜDM sahələr üzrə", 327, "əlavə dəyər", "mln AZN",
     "«digər sahələr» əlavə dəyəri — kataloqun `VA_DIGER` adı üçün SƏNƏDLƏŞDİRİLMİŞ əvəzedici"),
    ("ADDVALUE", "SNA", "ÜDM sahələr üzrə", 345, "əlavə dəyər", "mln AZN", "iqtisadiyyat üzrə cəmi əlavə dəyər"),
    ("NONOILGDP", "SNA", "ÜDM sahələr üzrə", 363, "Qeyri-neft ÜDM-i bazar qiymətləri ilə", "mln AZN",
     "qeyri neft-qaz ÜDM-i, bazar qiymətləri ilə"),
    ("RVA_MANU", "SNA", "ÜDM sahələr üzrə", 374, "Emal sənayesi", "mln AZN, sabit 2015",
     "emal sənayesi, əlavə dəyər, sabit qiymətlərlə"),
    ("RVA_CONST", "SNA", "ÜDM sahələr üzrə", 377, "Tikinti (inv daxili və xarici)", "mln AZN, sabit 2015",
     "tikinti, əlavə dəyər, sabit qiymətlərlə"),
    ("RVA_TRADE", "SNA", "ÜDM sahələr üzrə", 379, "Ticarət; nəqliyyat vasitələrinin təmiri",
     "mln AZN, sabit 2015", "ticarət, əlavə dəyər, sabit qiymətlərlə"),
    ("RVA_ACCOM", "SNA", "ÜDM sahələr üzrə", 380, "Turistlərin yerləşdirilməsi və ictimai iaşə",
     "mln AZN, sabit 2015", "turizm və ictimai iaşə, əlavə dəyər, sabit qiymətlərlə"),
    ("RVA_TRANS", "SNA", "ÜDM sahələr üzrə", 381, "Nəqliyyat və anbar təsərrüfatı", "mln AZN, sabit 2015",
     "nəqliyyat, əlavə dəyər, sabit qiymətlərlə"),
    ("RVA_INFORM", "SNA", "ÜDM sahələr üzrə", 382, "İnformasiya və rabitə", "mln AZN, sabit 2015",
     "rabitə, əlavə dəyər, sabit qiymətlərlə"),
    ("FPI_AZ", "SNA", "ÜDM sahələr üzrə", 394, "Kənd, meşə və balıqçılıq təsərrüfatı", "indeks, 2015=1",
     "kənd təsərrüfatı deflyatoru — kataloqun `FPI_AZ` adı üçün SƏNƏDLƏŞDİRİLMİŞ əvəzedici"),
    ("DEF_CONSTR", "SNA", "ÜDM sahələr üzrə", 395, "Tikinti", "indeks, 2015=1", "tikinti deflyatoru"),
    ("DEF_ACCOM", "SNA", "ÜDM sahələr üzrə", 397, "Turistlərin yerləşdirilməsi və ictimai iaşə",
     "indeks, 2015=1", "turizm və ictimai iaşə deflyatoru"),
    ("DEF_TRANSP", "SNA", "ÜDM sahələr üzrə", 398, "Nəqliyyat və anbar təsərrüfatı", "indeks, 2015=1",
     "nəqliyyat deflyatoru"),
    ("DEF_NTP", "SNA", "ÜDM sahələr üzrə", 400, "Məhsula və idxala xalis vergilər", "indeks, 2015=1",
     "məhsula və idxala xalis vergilərin deflyatoru"),
    # --- MOE SNA: `sosial və digər xidmətlər` -------------------------------------------
    ("VA_FINANCE", "SNA", "sosial və digər xidmətlər", 33, "Maliyyə və sığorta fəaliyyəti", "mln AZN",
     "maliyyə və sığorta, əlavə dəyər, nominal (`DEF_FINANCE` törəməsi üçün)"),
    ("VA_ADMN", "SNA", "sosial və digər xidmətlər", 35, "İnzibati və yardımçı xidmətlərin göstərilməsi",
     "mln AZN", "inzibati xidmətlər, əlavə dəyər (2009-dan — QISA NÜMUNƏ)"),
    ("VA_REST", "SNA", "sosial və digər xidmətlər", 36, "Daşınmaz əmlakla əlaqədar əməliyyatlar", "mln AZN",
     "daşınmaz əmlak, əlavə dəyər, nominal (`DEF_REST` törəməsi üçün)"),
    ("RVA_MINE", "SNA", "sosial və digər xidmətlər", 56, "Mədənçıxarma sənayesi", "mln AZN, sabit 2015",
     "mədənçıxarma, əlavə dəyər, sabit qiymətlərlə"),
    ("RVA_FINANCE", "SNA", "sosial və digər xidmətlər", 57, "Maliyyə və sığorta fəaliyyəti",
     "mln AZN, sabit 2015", "maliyyə və sığorta, əlavə dəyər, sabit qiymətlərlə"),
    ("RVA_REST", "SNA", "sosial və digər xidmətlər", 58, "Daşınmaz əmlakla əlaqədar əməliyyatlar",
     "mln AZN, sabit 2015", "daşınmaz əmlak, əlavə dəyər, sabit qiymətlərlə"),
    # --- MOE SNA: `Investisiya` ----------------------------------------------------------
    ("I", "SNA", "Investisiya", 5, "Əsas kapitala cəmi investisiyalar", "mln AZN", "cəmi investisiya, nominal"),
    ("ID", "SNA", "Investisiya", 8, "Daxili investisiya", "mln AZN", "daxili investisiya, nominal"),
    ("IF", "SNA", "Investisiya", 10, "Xarici investisiya", "mln AZN", "xarici investisiya, nominal"),
    ("IO", "SNA", "Investisiya", 12, "Neft və qaz sektoru", "mln AZN", "neft-qaz sektoruna investisiya, nominal"),
    ("INO", "SNA", "Investisiya", 14, "Digər sahələr", "mln AZN", "digər sahələrə investisiya, nominal"),
    ("IOD", "SNA", "Investisiya", 22, "Neft və qaz sektoru", "mln AZN", "daxili neft-qaz investisiyası"),
    ("IOF", "SNA", "Investisiya", 34, "Neft və qaz sektoru", "mln AZN", "xarici neft-qaz investisiyası"),
    ("RID", "SNA", "Investisiya", 44, "Daxili investisiyalar 2015-ci ilin qiymətləri ilə",
     "mln AZN, sabit 2015", "daxili investisiya, sabit qiymətlərlə"),
    ("DEFID", "SNA", "Investisiya", 45, "Daxili investisiyaların deflyatoru", "indeks, 2015=1",
     "daxili investisiyaların deflyatoru"),
    ("RIF", "SNA", "Investisiya", 48, "Xarici investisiyalar 2015-ci ilin qiymətləri ilə",
     "mln AZN, sabit 2015", "xarici investisiya, sabit qiymətlərlə"),
    ("RIDNS", "SNA", "Investisiya", 100, "Daxili qeyri-dövlət investisiyaları 2015-ci ilin qiymətləri ilə",
     "mln AZN, sabit 2015", "daxili qeyri-dövlət investisiyası, sabit qiymətlərlə"),
    # --- MOE SNA: `IN` -------------------------------------------------------------------
    ("RXGNO", "SNA", "IN", 35, "Digər sektorlar", "mln USD", "qeyri neft-qaz malların ixracı"),
    ("MGO", "SNA", "IN", 37, "Neft və qaz", "mln USD", "neft-qaz malların idxalı"),
    ("MGNO", "SNA", "IN", 38, "Digər sektorlar", "mln USD", "qeyri neft-qaz malların idxalı"),
    ("FOREIGNTURIST", "SNA", "IN", 26, "Xaricdən gələn turistlərin sayı DTA məlumatı", "min nəfər",
     "xaricdən gələn turistlərin sayı (2006-dan — QISA NÜMUNƏ)"),
    # --- MOE INF -------------------------------------------------------------------------
    ("OPWTI", "INF", "IN", 4, "Neftin qiyməti", "USD/barel", "neftin qiyməti"),
    ("OPIWTI", "INF", "IN", 5, "Crude Oil (petroleum), West Texas Intermediate 40 API, Midland Texas",
     "USD/barel", "WTI neft qiyməti"),
    ("INFLATION", "INF", "11_12", 6, "İnflyasiya", "%", "orta illik inflyasiya"),
    ("ER", "INF", "11_12", 9, "ABŞ dolları", "AZN/USD", "orta illik məzənnə"),
    ("LCPI_HP", "INF", "11_12", 20, "lcpi_hp", "loq", "İQİ loqarifminin HP trendi"),
    ("CPI", "INF", "11_12", 24, "İQİ, indeks, 2015=1", "indeks, 2015=1", "İstehlak Qiymətləri İndeksi"),
    ("NER", "INF", "11_12", 28, "Nominal valyuta məzənnəsi(ABŞ dolları)", "indeks", "nominal məzənnə (model)"),
    ("FPI_WEO", "INF", "11_12", 29, "Commodity Food Price Index, WEO", "indeks, 2015=1",
     "dünya ərzaq qiymətləri indeksi"),
    ("NIRR", "INF", "Mon", 84, "Refinancing rate (nirr)", "%", "uçot dərəcəsi"),
    ("NIRD", "INF", "Mon", 85, "Interest rate on deposits (nird)", "%", "depozit faizi"),
    ("NIRDFC", "INF", "Mon", 86, "Interest rate on foreign currency deposits (nirdfc)", "%", "xarici valyutada depozit faizi"),
    ("RIRD", "INF", "Mon", 89, "rird", "%", "real depozit faizi"),
    ("RIRL", "INF", "Mon", 90, "rirl (as in eviews)", "%", "real kredit faizi"),
    ("MB", "INF", "Mon", 94, "Pul bazası", "mln AZN", "pul bazası"),
    ("M1", "INF", "Mon", 103, "M1", "mln AZN", "M1 pul aqreqatı"),
    ("M3", "INF", "Mon", 107, "M3", "mln AZN", "M3 pul aqreqatı"),
    ("M2_M1", "INF", "Mon", 108, "M2-M1", "mln AZN", "M2 − M1"),
    ("M3_M2USD", "INF", "Mon", 112, "M3-M2, USD", "mln USD", "M3 − M2, USD"),
    ("CPIMTP", "INF", "Mon", 115, "cpimtp", "indeks, 2015=1", "ticarət tərəfdaşlarının İQİ-si"),
    ("NEER", "INF", "Mon", 116, "neer", "indeks, 2015=1", "nominal effektiv məzənnə"),
    ("REER", "INF", "Mon", 117, "reer", "indeks, 2015=1", "real effektiv məzənnə"),
    ("RGDPMTP", "INF", "reerEA", 287, "rgdpmtp", "indeks", "tərəfdaş ölkələrin real ÜDM indeksi"),
    ("NERMTP", "INF", "reerEA", 311, "nermtp", "indeks", "tərəfdaş ölkələrin nominal məzənnə indeksi"),
    # --- MOE BOP -------------------------------------------------------------------------
    ("XS_TP", "BOP", "Trade of Service", 15, "1. Nəqliyyat xidmətləri", "mln USD",
     "nəqliyyat xidmətlərinin ixracı"),
    ("XS_TRAVB", "BOP", "Trade of Service", 20, "İşgüzar səfərlər", "mln USD",
     "işgüzar səfərlər — xidmət ixracı"),
    ("XS_IFORM", "BOP", "Trade of Service", 22, "3. Rabitə xidmətləri", "mln USD",
     "rabitə xidmətlərinin ixracı"),
    ("MS_TP", "BOP", "Trade of Service", 38, "1. Nəqliyyat xidmətləri", "mln USD",
     "nəqliyyat xidmətlərinin idxalı"),
    ("MS_TRAVB", "BOP", "Trade of Service", 45, "İşgüzar səfərlər", "mln USD",
     "işgüzar səfərlər — xidmət idxalı"),
    ("MS_TRAVP", "BOP", "Trade of Service", 48, "Şəxsi səfərlər", "mln USD",
     "şəxsi səfərlər — xidmət idxalı"),
    # --- INDUSTRY ------------------------------------------------------------------------
    ("PI_IND3_1", "IND", "3.1", 4, "Qida məhsullarının istehsalı", "min AZN",
     "qida məhsulları istehsalı, müqayisəli qiymətlərlə (2009-dan — QISA NÜMUNƏ)"),
    ("PI_IND3_2", "IND", "3.2. VAR", 4, "İçki istehsalı", "min AZN",
     "içki istehsalı, müqayisəli qiymətlərlə (2009-dan — QISA NÜMUNƏ)"),
    # --- MOE T&C -------------------------------------------------------------------------
    ("TRANS_TURN_G", "TC", "NvR", 18, "Avtomobil", "mln t-km", "avtomobil yük dövriyyəsi"),
    ("TRANS_P", "TC", "NvR", 37, "Nəqliyyat sektorunda sərnişinin daşınması", "min sərn.",
     "nəqliyyat sektorunda sərnişin daşınması"),
    # --- MOE FISCAL ----------------------------------------------------------------------
    ("BE_NONEC", "FISCAL", "Dövlət büdcəsi", 68,
     "Əsaslı xərclər çıxmaqla dövlət büdcəsinin xərcləri", "mln AZN",
     "əsaslı xərclər çıxmaqla büdcə xərcləri (QISA NÜMUNƏ)"),
]

# Kataloqun bəzi adları iş kitablarında AYRI sıra kimi mövcud deyil; onlar ya eyni sıranın
# ikinci adıdır (ALIAS), ya da bir sətirlik AÇIQ törəmədir (DERIVED). Hər ikisi bayraqlanır.
MOE_SPEC_ALIAS = {
    "WAGE": ("W", "kataloq eyni sıranı iki adla çağırır"),
    "MINWAGE": ("MW", "kataloq eyni sıranı iki adla çağırır"),
    "CPICMTP": ("CPIMTP", "kataloqda yalnız bir sətirdə (17) rast gəlinir; iş kitablarında "
                          "`cpicmtp` adlı sıra YOXDUR — `cpimtp` ilə eyniləşdirilir, BAYRAQLANIR"),
    "TIKINTI": ("RVA_CONST", "kataloqun 51-ci sətri tikintinin real əlavə dəyərini hədəfləyir "
                             "(48/32-ci sətirlərlə eyni forma) — `RVA_CONST` ilə eyniləşdirilir, BAYRAQLANIR"),
}
MOE_SPEC_DERIVED = {
    "RW": ("W / CPI", "real əmək haqqı — iş kitabında ayrıca sıra yoxdur"),
    "DEF_FINANCE": ("VA_FINANCE / RVA_FINANCE",
                    "maliyyə sektorunun deflyatoru — nominal ÷ sabit (2015 = 1 yoxlanılır)"),
    "DEF_REST": ("VA_REST / RVA_REST",
                 "daşınmaz əmlak deflyatoru — nominal ÷ sabit (2015 = 1 yoxlanılır)"),
    "RIO": ("IO / DEFID", "neft-qaz investisiyası, sabit qiymətlərlə — daxili investisiya deflyatoru ilə"),
    "RINO": ("INO / DEFID", "qeyri-neft investisiyası, sabit qiymətlərlə — daxili investisiya deflyatoru ilə"),
    "REALFINALC": ("FINALC_NOM / (HC / RHC)",
                   "son istehlakın real həcmi — ev təsərrüfatlarının deflyatoru ilə"),
}


def _moe_spec_yearcols(rows):
    """İl başlığı sətrini AVTOMATİK aşkarlayır (vərəqdən vərəqə 1, 2, 3 və 4-cü sətir olur)."""
    best, brow = {}, None
    for i, row in enumerate(rows[:8], start=1):
        yrs = {j + 1: int(v) for j, v in enumerate(row)
               if isinstance(v, (int, float)) and not isinstance(v, bool)
               and 1990 <= v <= 2035 and float(v) == int(v)}
        if len(yrs) > len(best):
            best, brow = yrs, i
    return best, brow


def _moe_spec_assert_not_model(cells, ref):
    """Guardrail 2: `model` / `add factor` sətri MƏLUMAT kimi oxuna bilməz."""
    for c in cells[:5]:
        t = _key(c or "")
        if t in ("model", "addfactor", "add factor".replace(" ", "")):
            raise AssertionError(f"NÜMUNƏ-MODEL SƏTRİ — {ref}: `model`/`add factor` sətri məlumat deyil")


def extract_moe_spec_panel(dirpath: str = ""):
    """İş kitablarından paneli YENİDƏN qurur (etiket yoxlaması ilə). (panel, provenans) qaytarır.

    Qovluq verilməyibsə və ya mövcud deyilsə `(None, None)` qaytarır — təhvil mühitində
    Nazirliyin nümunə-model faylları olmaya bilər və bu, xəta deyil.
    """
    dirpath = dirpath or MOE_MODEL_DIR
    if not dirpath or not os.path.isdir(dirpath):
        return None, None
    import openpyxl as _ox
    cache = {}
    recs, prov = [], []
    for var, wbkey, sheet, row, label, unit, note in MOE_SPEC_ROWS:
        path = os.path.join(dirpath, MOE_SPEC_WORKBOOKS[wbkey])
        key = (wbkey, sheet)
        if key not in cache:
            wb = _ox.load_workbook(path, read_only=True, data_only=True)
            ws = wb[sheet]
            cache[key] = list(ws.iter_rows(min_row=1, max_row=max(420, row + 5),
                                           max_col=60, values_only=True))
            wb.close()
        rows = cache[key]
        ycol, hdr_row = _moe_spec_yearcols(rows)
        cells = rows[row - 1]
        _moe_spec_assert_not_model(cells, f"{MOE_SPEC_WORKBOOKS[wbkey]} {sheet}!{row}")
        labels = [str(c).strip() for c in cells[:5] if isinstance(c, str)]
        if not any(_key(g) == _key(label) for g in labels):
            raise AssertionError(
                f"ETİKET UYĞUNSUZLUĞU — {MOE_SPEC_WORKBOOKS[wbkey]} {sheet}!{row}: "
                f"gözlənilən {label!r}, tapılan {labels!r}")
        cap = MOE_SPEC_YEAR_CAP.get(var, MOE_LAST_ACTUAL)
        vals = {}
        for c, y in sorted(ycol.items(), key=lambda kv: kv[1]):
            if y > cap or c - 1 >= len(cells):
                continue
            v = cells[c - 1]
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                vals[int(y)] = float(v)
        for y in sorted(vals):
            recs.append({"var": var, "year": y, "value": vals[y]})
        prov.append({"var": var, "workbook": MOE_SPEC_WORKBOOKS[wbkey], "sheet": sheet,
                     "row": row, "row_label": label, "header_row": hdr_row, "unit": unit,
                     "year_first": (min(vals) if vals else None),
                     "year_last": (max(vals) if vals else None), "n": len(vals),
                     "note_az": note})
    return (pd.DataFrame(recs, columns=["var", "year", "value"]),
            pd.DataFrame(prov, columns=["var", "workbook", "sheet", "row", "row_label",
                                        "header_row", "unit", "year_first", "year_last",
                                        "n", "note_az"]))


def rebuild_moe_spec_panel(dirpath: str = "", verbose: bool = True):
    """Paneli iş kitablarından yenidən qurub `data/` altına yazır (builder-in çağırdığı funksiya)."""
    panel, prov = extract_moe_spec_panel(dirpath)
    if panel is None:
        raise FileNotFoundError(
            "MOE nümunə-model qovluğu tapılmadı; `MOE_MODEL_DIR` mühit dəyişənini təyin edin")
    panel.to_csv(P_MOE_SPEC_PANEL, index=False)
    prov.to_csv(P_MOE_SPEC_PROV, index=False)
    if verbose:
        print(f"moe_spec_panel: {len(panel)} sətir, {panel['var'].nunique()} sıra → {P_MOE_SPEC_PANEL}")
    return panel, prov


def moe_spec_panel() -> pd.DataFrame:
    """Nazirlik spesifikasiyasının tələb etdiyi sıralar (donmuş panel, ≤2024). Sütunlar: var, year, value."""
    if not os.path.exists(P_MOE_SPEC_PANEL):
        raise FileNotFoundError(
            f"{P_MOE_SPEC_PANEL} yoxdur — donmuş panel faylı əvvəlcədən qurulmalıdır")
    return pd.read_csv(P_MOE_SPEC_PANEL, float_precision="round_trip")


def moe_spec_provenance() -> pd.DataFrame:
    """Panelin sətir-sətir provenansı (iş kitabı, vərəq, sətir, etiket, dövr, qeyd)."""
    if not os.path.exists(P_MOE_SPEC_PROV):
        raise FileNotFoundError(f"{P_MOE_SPEC_PROV} yoxdur")
    return pd.read_csv(P_MOE_SPEC_PROV)


def moe_spec_series(var: str) -> pd.Series:
    """Bir sıranı il indeksi ilə qaytarır (yoxdursa boş sıra)."""
    p = moe_spec_panel()
    s = p[p["var"] == var]
    return pd.Series(s["value"].values, index=s["year"].astype(int).values,
                     name=var, dtype=float).sort_index()


def verify_moe_spec_panel(dirpath: str = "", tol: float = 1e-9):
    """Donmuş paneli MƏNBƏDƏN yenidən oxuyur (etiket yoxlaması ilə) və maksimal fərqi qaytarır.

    Fayllar yoxdursa `None` qaytarır — bu, xəta deyil.
    """
    got, _ = extract_moe_spec_panel(dirpath)
    if got is None:
        return None
    want = moe_spec_panel()
    m = want.merge(got, on=["var", "year"], how="outer", suffixes=("_disk", "_src"))
    miss = m[m["value_disk"].isna() | m["value_src"].isna()]
    if len(miss):
        raise AssertionError(
            f"MOE spesifikasiya paneli mənbədən fərqlənir: {len(miss)} sətir yalnız bir tərəfdə "
            f"(ilk: {miss.iloc[0]['var']} {int(miss.iloc[0]['year'])})")
    diff = float((m["value_disk"] - m["value_src"]).abs().max())
    if diff > tol:
        raise AssertionError(f"MOE spesifikasiya paneli mənbədən fərqlənir: maks. fərq {diff:.3e}")
    return diff


# =========================================================================================
# 14. RƏSMİ AÇIQ MƏNBƏLƏRİN PANELİ (DSK / AMB / opendata.az / MOE BOP) — bu buraxılış
#     `DATA_GAP_NEGATIVE_PROOF.md` sübutunun xanaları ilə yoxlanmış yerləri məlumat
#     qatına bağlayır. Konvensiya §13-ün eynisidir: spesifikasiya BURADADIR, dəyərlər
#     `data/public_sources_panel.csv` faylında DONDURULUB, provenans isə
#     `data/public_sources_provenance.csv` faylında sətir-sətir saxlanılır.
# =========================================================================================
# Niyə dondurulmuş panel: mənbə faylları (`data_collection_20260603/raw/…` və Nazirliyin
# nümunə-model iş kitabları) təhvil paketinin XARİCİNDƏDİR. §10–§13-də olduğu kimi paket
# özünü-tam qalmalıdır, ona görə dəyərlər paketin öz `data/` qovluğunda saxlanılır;
# `verify_public_panel()` mənbələr əlçatan olduqda hər sıranı ETİKET YOXLAMASI ilə yenidən
# oxuyub tutuşdurur.
#
# GUARDRAİL-lər, kodda məcburi:
#   (1) Nazirliyin iş kitablarından YALNIZ faktiki sütunlar (≤2024) oxunur — `Income account`
#       vərəqinin annotasiya sətri 2025-i "gözlənilən", 2026–2030-u "proqnoz" kimi işarələyir;
#   (2) DSK/AMB nəşrlərinin ulduzlu (`2025*`) sütunu ilkin məlumatdır, faktiki sayılır və
#       provenansda AÇIQ göstərilir;
#   (3) hər sətir üçün gözlənilən etiket spesifikasiyada saxlanılır və oxunuş anında
#       yoxlanılır; uyğunsuzluqda icra dayanır (çılpaq mövqe indeksi ilə oxu yoxdur).

PUBLIC_RETRIEVED = "2026-06-03"      # `data_collection_20260603` anlıq nüsxəsinin tarixi

PUBLIC_SOURCE_NOTE = (
    "Rəsmi açıq mənbələr: DSK (Dövlət Statistika Komitəsi) nəşr cədvəlləri, AMB (Mərkəzi "
    "Bank) monetar göstəriciləri və NSDP buraxılışı, opendata.az açıq məlumat resursu, "
    f"Nazirliyin MOE BOP iş kitabı. Anlıq nüsxə {PUBLIC_RETRIEVED}. Hər sıra üçün gözlənilən "
    "etiket `PUBLIC_SPEC`-də saxlanılır və `verify_public_panel()` tərəfindən yoxlanılır")

# Mənbə faylları — YALNIZ yoxlama/yenidən qurma üçün (bax MOE_AGRI_XLSX qeydi: şəxsi yol
# təhvil paketinə yazılmır, mühit dəyişəni ilə verilir).
PUBLIC_DATA_DIR = os.environ.get("PUBLIC_DATA_DIR", "")

# açar -> (kök qovluğa nisbi yol, mənbənin insan oxunaqlı adı)
PUBLIC_FILES = {
    "cbar216": ("cbar/monetary_indicators/xlsx/5d454b461eaeb9cd59e4efeca.xlsx",
                "AMB, cədvəl 2.16 «Manatın rəsmi orta məzənnəsi»"),
    "ssc010": ("ssc/system_nat_accounts/010en.xls",
               "DSK, cədvəl 10 «Gross domestic product — manats, dollars, in euro»"),
    "ssc013": ("ssc/system_nat_accounts/013en.xls",
               "DSK, cədvəl 13 «Production and generation of income account»"),
    "ssc027": ("ssc/system_nat_accounts/027en.xls",
               "DSK, cədvəl 27 «Gross domestic product use, at current prices»"),
    "ssc018": ("ssc/industry/018en.xls",
               "DSK, cədvəl 18 «Manufacture of the most important types of industrial "
               "products in kind»"),
    "ssc146": ("ssc/agriculture/1.46en.xls",
               "DSK, cədvəl 1.46 «Meat production by type, in slaughtered weight»"),
    "ssc147": ("ssc/agriculture/1.47en.xls",
               "DSK, cədvəl 1.47 «Milk production by type, by farm categories»"),
    "ssc0458": ("ssc/labour/004_5-8en.xls",
                "DSK, cədvəl 4.5–4.8 (`Dynamics`) «Average monthly nominal wages and "
                "salaries of employees by economic activities and property forms»"),
    "ssc0212": ("ssc/labour/002_12-13en.xls",
                "DSK, cədvəl 2.12 (`Dynamics_2.12`) «Number of employees by property "
                "forms and economic activities»"),
    "odazinv": ("local_existing_snapshots/opendata_az/resources/"
                "investments-in-fixed-capital-by-financial-sources__"
                "Maliyyə_mənbələri_üzrə_əsas_kapitala_yönəldilmiş_investisiyalar.csv",
                "opendata.az «Maliyyə mənbələri üzrə əsas kapitala yönəldilmiş "
                "investisiyalar»"),
    "nsdpbop": ("local_existing_snapshots/nsdp_latest/nsdp_balance_of_payments_cbar.xlsx",
                "AMB NSDP buraxılışı, `BOP Analytical` vərəqi"),
    "moebop": ("MOE BOP.xlsx", "Nazirliyin MOE BOP.xlsx nümunə-model iş kitabı, "
                               "`Income account` vərəqi"),
}

# `moebop` açarı Nazirliyin iş kitabları qovluğundan (`MOE_MODEL_DIR`), qalanları isə açıq
# məlumat kökündən (`PUBLIC_DATA_DIR`) oxunur.
PUBLIC_MOE_KEYS = {"moebop"}

# DSK 13-ün NACE bölmələri: hərf -> (gözlənilən etiket, model sahəsinin kökü və ya None)
SSC013_SECTIONS = [
    ("A", "Agriculture, forestry and fishing", "agri"),
    ("B", "Mining", "mining"),
    ("C", "Manufacturing", "manuf"),
    ("D", "Electricity, gas and steam production, distribution and supply", "power"),
    ("E", "Water supply, waste treatment and disposal", "water"),
    ("F", "Construction", "constr"),
    ("G", "Trade: repair of transport means", "trade"),
    ("H", "Transportation and storage", "transport"),
    ("I", "Accommodation and food service activities", "tourism"),
    ("J", "Information and communication", "ict"),
    ("K", "Financial and insurance activities", "social"),
    ("L", "Real estate activities", "social"),
    ("M", "Professional, scientific and technical activities", "social"),
    ("N", "Administrative and support service activities", "social"),
    ("O", "Public administration and defence; social security", "social"),
    ("P", "Education", "social"),
    ("Q", "Human health and social work activities", "social"),
    ("R", "Arts, entertainment and recreation", "social"),
    ("S", "Other service activities", "social"),
]
SSC013_MEASURES = [("go", "Output", "P.1"), ("ic", "Intermediate consumption", "P.2"),
                   ("va", "Value added", "B.1g")]

# DSK 4.5–4.8 `Dynamics` sətirləri: (kod kökü, gözlənilən etiket)
SSC_ACTIVITY_ROWS = [
    ("total", "On economy, total"),
    ("agri", "Agriculture, forestry and fishing"),
    ("mining", "Mining"),
    ("manuf", "Manufacturing"),
    ("power", "Electricity, gas and steam production, distribution and supply"),
    ("water", "Water supply; waste treatment and disposal"),
    ("constr", "Construction"),
    ("trade", "Trade; repair of transport means"),
    ("transport", "Transportation and storage"),
    ("tourism", "Accommodation and food service activities"),
    ("ict", "Information and communication"),
    ("finance", "Financial and insurance activities"),
    ("realest", "Real estate activities"),
    ("prof", "Professional, scientific and technical activities"),
    ("admin", "Administrative and support service activities"),
    ("pubadm", "Public administration and defence; social security"),
    ("educ", "Education"),
    ("health", "Human health and social work activities"),
    ("art", "Art, entertainment and recreation"),
    ("other", "Other service activities"),
]
SSC_OWNERSHIP = [("", "  Total"), ("_state", "state"), ("_nonstate", "non-state")]

# Statutar şablonun 2.4.1.4. vərəqindəki 11 ÖLÜ (`#REF!`) məhsul sətrinin DSK qarşılığı
# (sübut sənədinin (vii) bəndi). Sahə: (şablon sətri, şablon adı, şablon vahidi, kod,
#  mənbə açarı, yerləşmə, gözlənilən etiket, mənbə vahidi, çevirmə əmsalı, status).
# `factor` — mənbə vahidindən şablon vahidinə çevirmə; `None` isə vahidlər uyğun gəlmir
# (sıra öz vahidində nəşr olunur, çevrilmir) və status bunu AÇIQ göstərir.
VEREQ_DEAD_PRODUCTS = [
    (48, "Mal əti", "min ton", "prod_ssc_beef", "ssc146", ("col", 3), "beef",
     "min ton", 1.0, "uyğun — kənd təsərrüfatı mənbəyi (sənaye cədvəlində yoxdur)"),
    (49, "Quş əti", "min ton", "prod_ssc_poultry", "ssc146", ("col", 6), "poultry meat",
     "min ton", 1.0, "uyğun — kənd təsərrüfatı mənbəyi (sənaye cədvəlində yoxdur)"),
    (50, "Meyvə və tərəvəz şirələri", "min dkl", "prod_ssc_juice", "ssc018", ("row", 27),
     "Fruit and vegetables tinned, juices, thsd. tons", "min ton", None,
     "ÖLÇÜ VAHİDİ UYĞUN DEYİL — şablon min dkl, DSK min ton; həm də DSK sətri "
     "konservləri şirələrlə birlikdə verir"),
    (51, "Emal edilmiş duru süd", "min ton", "prod_ssc_milk", "ssc147", ("col", 2),
     "Milk-total", "min ton", 1.0,
     "ANLAYIŞ FƏRQİ — DSK sırası xam süd istehsalıdır, şablon isə emal edilmiş duru süddür"),
    (52, "Qənd-rafinad və şəkər tozu", "min ton", "prod_ssc_sugar", "ssc018", ("row", 39),
     "Sugar, thsd. tons", "min ton", 1.0, "uyğun"),
    (53, "Kərə yağı", "min ton", "prod_ssc_butter", "ssc018", ("row", 24),
     "Butter, thsd. tons", "min ton", 1.0, "uyğun"),
    (54, "Pendir və kəsmik", "min ton", "prod_ssc_cheese", "ssc018", ("row", 23),
     "Cheese and curd, ton", "ton", 0.001, "uyğun — vahid ton → min ton çevrilir"),
    (55, "Bitki yağları", "min ton", "prod_ssc_vegoil", "ssc018", ("row", 31),
     "Vegetable oils, thsd. tons", "min ton", 1.0, "uyğun"),
    (56, "Un", "min ton", "prod_ssc_flour", "ssc018", ("row", 40),
     "Flour, thsd. tons", "min ton", 1.0, "uyğun"),
    (57, "Çörək və çörək-bulka məmulatları", "min ton", "prod_ssc_bread", "ssc018",
     ("row", 41), "Bread, thsd. tons", "min ton", 1.0, "uyğun"),
    (58, "Xörək duzu", "min ton", "prod_ssc_salt", "ssc018", ("row", 44),
     "Table salt, ton", "ton", 0.001, "uyğun — vahid ton → min ton çevrilir"),
]


def public_source_path(key, dirpath="", moe_dir=""):
    """Mənbə faylının tam yolu; kök qovluq verilməyibsə boş sətir qaytarır."""
    rel, _ = PUBLIC_FILES[key]
    root = (moe_dir or MOE_MODEL_DIR) if key in PUBLIC_MOE_KEYS else (dirpath or PUBLIC_DATA_DIR)
    return os.path.join(root, rel) if root else ""


def _xls_sheet(path, sheet=None):
    import xlrd
    bk = xlrd.open_workbook(path)
    return bk.sheet_by_name(sheet) if sheet else bk.sheet_by_index(0)


def _cell(sh, r, c):
    return sh.cell_value(r, c) if (r < sh.nrows and c < sh.ncols) else ""


def _assert_label(got, want, ref):
    if _key(got) != _key(want):
        raise AssertionError(
            f"ETİKET UYĞUNSUZLUĞU — {ref}: gözlənilən {want!r}, tapılan {str(got)!r}")


def _year_of(v):
    """`2025*`, `  2010 `, 2013.0 → 2025 / 2010 / 2013; il deyilsə None."""
    s = _squash(v).replace("*", "").strip()
    try:
        y = int(float(s))
    except (TypeError, ValueError):
        return None
    return y if 1985 <= y <= 2035 else None


def _prov(var, key, table, locator, label, unit, vals, note):
    ys = sorted(vals)
    return dict(var=var, source_key=key, source_file=PUBLIC_FILES[key][0],
                source_name=PUBLIC_FILES[key][1], table=table, locator=locator,
                row_label=label, unit=unit,
                year_first=(ys[0] if ys else None), year_last=(ys[-1] if ys else None),
                n=len(vals), retrieved=PUBLIC_RETRIEVED, note_az=note)


# --- mənbə-mənbə oxucular ----------------------------------------------------------------

def _read_cbar216(path):
    """AMB 2.16: illik ORTA rəsmi məzənnə — ABŞ dolları (B) və AVRO (C)."""
    import openpyxl as _ox
    wb = _ox.load_workbook(path, read_only=True, data_only=True)
    ws = wb["2.16"]
    rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=4, values_only=True))
    _assert_label(rows[5][0], "Tarix", "AMB 2.16!A6")
    _assert_label(rows[6][1], "US dollar", "AMB 2.16!B7")
    _assert_label(rows[6][2], "EURO", "AMB 2.16!C7")
    usd, eur = {}, {}
    for r in rows[8:]:
        y = _year_of(r[0])
        if y is None:
            continue                          # aylıq sətirlər ('01'…'12') atlanır
        for col, tgt in ((1, usd), (2, eur)):
            v = r[col]
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                tgt[y] = float(v)
    wb.close()
    out = {}
    out["fx_usd_azn_avg"] = (usd, _prov(
        "fx_usd_azn_avg", "cbar216", "2.16", "sütun B", "US dollar", "AZN/USD", usd,
        "manatın ABŞ dollarına nisbətdə RƏSMİ ORTA illik məzənnəsi (dövrün sonuna deyil)"))
    out["fx_eur_azn"] = (eur, _prov(
        "fx_eur_azn", "cbar216", "2.16", "sütun C", "EURO", "AZN/EUR", eur,
        "manatın avroya nisbətdə rəsmi orta illik məzənnəsi"))
    return out


def _read_ssc010(path):
    """DSK 10: ÜDM-in manat/dollar/avro ifadəsindən doğan implisit orta məzənnələr."""
    sh = _xls_sheet(path)
    _assert_label(_cell(sh, 3, 8), "1$= manats", "DSK 010en!I4")
    _assert_label(_cell(sh, 3, 9), "1euro= manats", "DSK 010en!J4")
    usd, eur = {}, {}
    for r in range(5, sh.nrows):
        y = _year_of(_cell(sh, r, 1))
        if y is None:
            continue
        for col, tgt in ((8, usd), (9, eur)):
            v = num(_cell(sh, r, col))
            if not (isinstance(v, float) and math.isnan(v)):
                tgt[y] = float(v)
    return {
        "fx_usd_azn_ssc": (usd, _prov(
            "fx_usd_azn_ssc", "ssc010", "10", "sütun `1$= manats`", "1$= manats",
            "AZN/USD", usd, "AMB 2.16 üçün MÜSTƏQİL çarpaz yoxlama sırası")),
        "fx_eur_azn_ssc": (eur, _prov(
            "fx_eur_azn_ssc", "ssc010", "10", "sütun `1euro= manats`", "1euro= manats",
            "AZN/EUR", eur, "AMB 2.16 üçün MÜSTƏQİL çarpaz yoxlama sırası")),
    }


def _read_ssc013(path):
    """DSK 13: NACE bölmələri üzrə buraxılış (P.1) / aralıq istehlak (P.2) / əlavə dəyər (B.1g)."""
    sh = _xls_sheet(path, "013")
    blocks = []
    for r in range(sh.nrows):
        if _squash(_cell(sh, r, 3)) == "Output":
            y = _year_of(_cell(sh, r - 1, 3))
            if y is not None:
                blocks.append((r, y))
    if len(blocks) < 20:
        raise AssertionError(f"DSK 013en: {len(blocks)} il bloku tapıldı, ≥20 gözlənilirdi")
    acc = {}          # var -> {year: value}
    for hdr, year in blocks:
        _assert_label(_cell(sh, hdr, 4), "Intermediate consumption", f"DSK 013en!E{hdr + 1}")
        _assert_label(_cell(sh, hdr, 5), "Value added", f"DSK 013en!F{hdr + 1}")
        _assert_label(_cell(sh, hdr + 1, 3), "P.1", f"DSK 013en!D{hdr + 2}")
        found = {}
        for r in range(hdr + 2, min(hdr + 32, sh.nrows)):
            letter = _squash(_cell(sh, r, 1))
            label = _squash(_cell(sh, r, 2))
            if _key(label) == _key("Total"):
                found["TOTAL"] = r
                break
            if letter and letter in dict((a, b) for a, b, _ in SSC013_SECTIONS):
                found[letter] = r
        for letter, label, _root in SSC013_SECTIONS:
            r = found.get(letter)
            if r is None:
                raise AssertionError(f"DSK 013en {year}: `{letter}` bölməsi tapılmadı")
            _assert_label(_cell(sh, r, 2), label, f"DSK 013en!C{r + 1} ({year})")
            for m, _mlabel, _code in SSC013_MEASURES:
                col = {"go": 3, "ic": 4, "va": 5}[m]
                v = num(_cell(sh, r, col))
                if not (isinstance(v, float) and math.isnan(v)):
                    acc.setdefault(f"{m}_ssc_{letter.lower()}", {})[year] = float(v)
        rt = found.get("TOTAL")
        if rt is None:
            raise AssertionError(f"DSK 013en {year}: `Total` sətri tapılmadı")
        for m in ("go", "ic", "va"):
            col = {"go": 3, "ic": 4, "va": 5}[m]
            v = num(_cell(sh, rt, col))
            if not (isinstance(v, float) and math.isnan(v)):
                acc.setdefault(f"{m}_ssc_total", {})[year] = float(v)
    out = {}
    for letter, label, _root in SSC013_SECTIONS:
        for m, mlabel, mcode in SSC013_MEASURES:
            var = f"{m}_ssc_{letter.lower()}"
            out[var] = (acc[var], _prov(
                var, "ssc013", "13", f"NACE bölmə {letter}, sütun {mcode}", label,
                "mln AZN", acc[var], f"{mlabel} ({mcode}) — DSK-nın 21 illik blokundan"))
    for m, mlabel, mcode in SSC013_MEASURES:
        var = f"{m}_ssc_total"
        out[var] = (acc[var], _prov(var, "ssc013", "13", f"`Total` sətri, sütun {mcode}",
                                    "Total", "mln AZN", acc[var],
                                    f"{mlabel} ({mcode}) — bütün NACE bölmələri, FISIM daxil"))
    # Model sahələrinə (11 sahə) aqreqasiya: 1:1 uyğunlaşma, `social` isə K–S cəmi.
    for m, mlabel, mcode in SSC013_MEASURES:
        agg, parts = {}, {}
        for letter, _label, root in SSC013_SECTIONS:
            if root is None:
                continue
            parts.setdefault(root, []).append(letter)
            for y, v in acc[f"{m}_ssc_{letter.lower()}"].items():
                agg.setdefault(root, {}).setdefault(y, 0.0)
                agg[root][y] += v
        for root, vals in agg.items():
            var = f"{m}_ssc_{root}"
            letters = "+".join(parts[root])
            out[var] = (vals, _prov(
                var, "ssc013", "13", f"NACE {letters}, sütun {mcode}",
                "; ".join(lb for lt, lb, rt in SSC013_SECTIONS if rt == root),
                "mln AZN", vals,
                f"{mlabel} ({mcode}); model sahəsi `{root}` = NACE {letters} cəmi"))
    return out


def _read_ssc027(path):
    """DSK 27 «Gross domestic product use, at current prices» — ÜDM-in TAM İSTİFADƏ hesabı.

    Cədvəl ÜDM-i XƏRC (istifadə) metodu ilə verir və onun BÜTÜN sətirlərini daşıyır:
    faktiki son istehlak (P.4) ev təsərrüfatı / dövlət / QHT bölgüsü ilə, ümumi yığım
    (P.5 = P.51 əsas kapital + P.52 ehtiyatların dəyişməsi), xalis ixrac (P.6n = P.6 − P.7)
    və statistik fərq. Cəmi eyniliklə ÜDM-i (B.1*g) verir:

        B.1*g = P.4 + P.5 + (P.6 − P.7) + statistik fərq

    Əvvəlki buraxılışda bu cədvəldən YALNIZ üç sətir (dövlətin fərdi/kollektiv xərcləri və
    ev təsərrüfatlarının faktiki son istehlakı) oxunurdu; qalan sətirlər — xüsusilə
    **ehtiyatların dəyişməsi (P.52)** — paketdə DATA-GAP elan edilmişdi və məhz bu boşluq
    tələb-təklif (resurslar–istifadə) balansının qurulmasına mane olurdu. Sətirlər mənbədə
    1993-cü ildən mövcuddur; boşluq oxunuşda idi, məlumatda deyil. Bu buraxılışda bütün
    on dörd sətir oxunur və FR01b resurslar–istifadə balansını onların üzərində qurur.

    Mövcud üç kod (`gov_cons_indiv`, `gov_cons_collect`, `hh_cons_actual`) və onların
    törəməsi (`gov_cons_total`) DƏYİŞMİR — yeni sətirlər yanlarına AYRI kodlarla əlavə olunur.
    """
    sh = _xls_sheet(path, "027")
    ycol = {}
    for c in range(3, sh.ncols):
        y = _year_of(_cell(sh, 3, c))
        if y is not None:
            ycol[c] = y
    if len(ycol) < 30:
        raise AssertionError(f"DSK 027en: {len(ycol)} il sütunu tapıldı, ≥30 gözlənilirdi")
    # (sətir indeksi 0-dan, gözlənilən İngilis etiketi, kod, SNA əməliyyat kodu, AZ qeyd).
    # Sətir sırası cədvəldəki KİMİDİR; `None` əməliyyat kodu — mənbədə B sütunu boşdur
    # (statistik fərq sətri), ona görə orada kod yoxlaması aparılmır.
    ROWS = [
        (4, "Actual final consumption expenditures", "cons_final_total", "P.4",
         "faktiki son istehlak xərcləri, CƏMİ (ev təsərrüfatları + dövlət); istifadə "
         "hesabının birinci baş sətri"),
        (6, "Actual final consumption expenditures of households", "hh_cons_actual", "P.4",
         "ev təsərrüfatlarının faktiki son istehlak xərcləri — dövlət sırası ilə eyni cədvəldən"),
        (8, "final consumption expenditure of households", "hh_cons_final", "P.3",
         "ev təsərrüfatlarının ÖZ son istehlak xərcləri (P.3) — faktiki son istehlakdan "
         "(P.4) FƏRQLİ anlayış: P.4 dövlətin və QHT-lərin fərdi xidmətlərini də daşıyır"),
        (9, "expenditures of government  institutions providing individual services",
         "gov_cons_indiv", "P.3", "fərdi xidmətlər göstərən dövlət idarələrinin xərcləri"),
        (10, "expenditures of non - profit institutions providing services for households",
         "npish_cons_final", "P.3",
         "ev təsərrüfatlarına xidmət göstərən qeyri-kommersiya təşkilatlarının (QHT) "
         "son istehlak xərcləri"),
        (11, "Actual final consumption expenditures of government institutions providing "
             "collective services", "gov_cons_collect", "P.4",
         "kollektiv xidmətlər göstərən dövlət idarələrinin faktiki son istehlak xərcləri"),
        (12, "Gross saving", "gcf_total", "P.5",
         "ÜMUMİ YIĞIM = əsas kapitalın ümumi yığımı (P.51) + ehtiyatların dəyişməsi (P.52). "
         "DİQQƏT: mənbənin İngilis etiketi «Gross saving»dir, lakin SNA kodu (P.5) və "
         "cədvəlin öz cəmi bunun YIĞIM (capital formation) olduğunu göstərir — əmanət deyil"),
        (13, "gross fixed capital formation", "gfcf_total", "P.51",
         "əsas kapitalın ümumi yığımı — milli hesablar anlayışı; iş kitabının «əsas kapitala "
         "yönəldilmiş investisiyalar» sırasından (FR06) FƏRQLİ statistik anlayışdır"),
        (14, "changes in inventories (+, -)", "inventories_chg", "P.52",
         "ehtiyatların dəyişməsi — istifadə hesabının BUFER maddəsi; paketdə əvvəllər "
         "DATA-GAP elan edilmişdi, mənbədə 1993-cü ildən mövcuddur"),
        (15, "Net exports", "net_exports_gs", "P.6 n",
         "mal və xidmətlərin xalis ixracı = ixrac (P.6) − idxal (P.7)"),
        (16, "Exports", "exports_gs", "P.6",
         "mal və xidmətlərin ixracı, milli hesablar qiymətləndirməsi (manatla)"),
        (17, "Imports (-)", "imports_gs", "P.7",
         "mal və xidmətlərin idxalı, milli hesablar qiymətləndirməsi (manatla); cədvəldə "
         "MÜSBƏT yazılır, eynilikdə çıxılır"),
        (18, "Statistical discrepancy", "stat_discrepancy", None,
         "istehsal və istifadə hesablarının DSK-nın öz balanslaşdırmasından qalan fərqi; "
         "faktiki illərin əksəriyyətində sıfırdır"),
        (19, "GDP", "gdp_use", "B.1*g",
         "ÜDM — istifadə (xərc) hesabının cəmi; istehsal hesabının ÜDM-i ilə eyni "
         "kəmiyyətin ikinci ölçüsüdür"),
    ]
    out, got = {}, {}
    for r, label, var, code, note in ROWS:
        _assert_label(_cell(sh, r, 2), label, f"DSK 027en!C{r + 1}")
        if code is not None:
            _assert_label(_cell(sh, r, 1), code, f"DSK 027en!B{r + 1}")
        vals = {}
        for c, y in ycol.items():
            v = num(_cell(sh, r, c))
            if not (isinstance(v, float) and math.isnan(v)):
                vals[y] = float(v)
        got[var] = vals
        loc = f"sətir {r + 1}" + (f", {code}" if code is not None else "")
        out[var] = (vals, _prov(var, "ssc027", "27", loc, label, "mln AZN", vals, note))
    tot = {y: got["gov_cons_indiv"][y] + got["gov_cons_collect"][y]
           for y in sorted(set(got["gov_cons_indiv"]) & set(got["gov_cons_collect"]))}
    out["gov_cons_total"] = (tot, _prov(
        "gov_cons_total", "ssc027", "27", "sətir 10 + sətir 12 (P.3 + P.4)",
        "expenditures of government institutions providing individual services + "
        "Actual final consumption expenditures of government institutions providing "
        "collective services", "mln AZN", tot,
        "dövlət idarələrinin ÜMUMİ son istehlak xərcləri = fərdi (P.3) + kollektiv (P.4)"))

    # MƏNBƏNİN ÖZ EYNİLİYİ — oxunuşun yoxlanışı. Cədvəl balanslıdırsa hər il üçün
    #   B.1*g = P.4 + P.5 + (P.6 − P.7) + statistik fərq
    # bağlanmalıdır. Bağlanmırsa sətir xəritəsi sürüşüb və oxunuş DAYANDIRILIR.
    ys = sorted(set(got["gdp_use"]) & set(got["cons_final_total"]) & set(got["gcf_total"])
                & set(got["exports_gs"]) & set(got["imports_gs"]))
    worst, worst_y = 0.0, None
    for y in ys:
        lhs = got["gdp_use"][y]
        rhs = (got["cons_final_total"][y] + got["gcf_total"][y]
               + got["exports_gs"][y] - got["imports_gs"][y]
               + got["stat_discrepancy"].get(y, 0.0))
        d = abs(lhs - rhs)
        if d > worst:
            worst, worst_y = d, y
    if worst > 0.15:      # mənbənin öz yuvarlaqlaşması 0,1 mln AZN səviyyəsindədir
        raise AssertionError(
            f"DSK 027en: istifadə hesabının eyniliyi bağlanmır — {worst_y}-ci ildə "
            f"{worst:.3f} mln AZN fərq (sətir xəritəsi sürüşmüş ola bilər)")
    return out


def _read_ssc018(path):
    """DSK 18: ən mühüm sənaye məhsullarının NATURAL ifadədə istehsalı (1995–2024)."""
    sh = _xls_sheet(path, "18")
    ycol = {}
    for c in range(2, sh.ncols):
        y = _year_of(_cell(sh, 3, c))
        if y is not None:
            ycol[c] = y
    if len(ycol) < 25:
        raise AssertionError(f"DSK 018en: {len(ycol)} il sütunu tapıldı, ≥25 gözlənilirdi")
    return ycol, sh


def _read_products(dirpath, moe_dir):
    """11 ölü şablon sətrinin DSK qarşılığı (`VEREQ_DEAD_PRODUCTS` konkordansı)."""
    ycol018, sh018 = _read_ssc018(public_source_path("ssc018", dirpath, moe_dir))
    cache = {}
    out = {}
    for (vrow, vlabel, vunit, code, key, loc, label, sunit, factor, status) in VEREQ_DEAD_PRODUCTS:
        if key == "ssc018":
            sh, ycol = sh018, ycol018
            r = loc[1]
            _assert_label(_cell(sh, r, 1), label, f"DSK 018en!B{r + 1}")
            vals = {}
            for c, y in ycol.items():
                v = num(_cell(sh, r, c))
                if not (isinstance(v, float) and math.isnan(v)):
                    vals[y] = float(v) * (factor if factor else 1.0)
            locator = f"sətir {r + 1}"
        else:
            if key not in cache:
                cache[key] = _xls_sheet(public_source_path(key, dirpath, moe_dir))
            sh = cache[key]
            c = loc[1]
            hdr = _squash(_cell(sh, 4, c)) or _squash(_cell(sh, 3, c))
            _assert_label(hdr, label, f"{PUBLIC_FILES[key][0]}!col{c + 1}")
            _assert_label(_cell(sh, 5, 1), "All categories of farms",
                          f"{PUBLIC_FILES[key][0]}!B6")
            vals = {}
            for r in range(6, sh.nrows):
                y = _year_of(_cell(sh, r, 1))
                if y is None:
                    continue
                v = num(_cell(sh, r, c))
                if not (isinstance(v, float) and math.isnan(v)):
                    vals[y] = float(v) * (factor if factor else 1.0)
            locator = f"sütun {c + 1} («bütün təsərrüfat kateqoriyaları» bloku)"
        out[code] = (vals, _prov(
            code, key, PUBLIC_FILES[key][1].split("cədvəl ")[-1].split(" ")[0], locator,
            label, (vunit if factor else sunit), vals,
            f"statutar şablon 2.4.1.4. sətir {vrow} («{vlabel}», {vunit}) — konkordans: {status}"))
    return out


def _read_ssc_crosstab(path, sheet, hdr_row, first_row, unit, prefix, table, note):
    """DSK-nın «fəaliyyət növü × mülkiyyət forması» dinamika vərəqləri (eyni forma)."""
    sh = _xls_sheet(path, sheet)
    ycol = {}
    for c in range(3, sh.ncols):
        y = _year_of(_cell(sh, hdr_row, c))
        if y is not None:
            ycol[y] = c                       # ilin BİRİNCİ (Total) sütunu
    if len(ycol) < 15:
        raise AssertionError(f"{sheet}: {len(ycol)} il tapıldı, ≥15 gözlənilirdi")
    _assert_label(_cell(sh, hdr_row + 1, 3), "Total", f"{sheet}!D{hdr_row + 2}")
    _assert_label(_cell(sh, hdr_row + 2, 4), "state", f"{sheet}!E{hdr_row + 3}")
    _assert_label(_cell(sh, hdr_row + 2, 5), "non-state", f"{sheet}!F{hdr_row + 3}")
    out = {}
    for i, (root, label) in enumerate(SSC_ACTIVITY_ROWS):
        r = first_row + i
        _assert_label(_cell(sh, r, 2), label, f"{sheet}!C{r + 1}")
        for suffix, own in SSC_OWNERSHIP:
            off = {"": 0, "_state": 1, "_nonstate": 2}[suffix]
            vals = {}
            for y, c0 in ycol.items():
                v = num(_cell(sh, r, c0 + off))
                if not (isinstance(v, float) and math.isnan(v)):
                    vals[y] = float(v)
            var = f"{prefix}_{root}{suffix}"
            out[var] = (vals, _prov(
                var, "ssc0458" if prefix == "wage_ssc" else "ssc0212", table,
                f"sətir {r + 1}, {own.strip()} sütunu", label, unit, vals, note))
    return out


def _read_odaz_investments(path):
    """opendata.az: maliyyə mənbələri üzrə əsas kapitala investisiyalar (min manat)."""
    d = pd.read_csv(path)
    COLS = [
        ("inv_own_funds",
         "Müəssisə və təşkilatların öz vəsaitləri hesabına yönəldilmiş investisiyalar (min manat)",
         "müəssisə və təşkilatların öz vəsaitləri — iş kitabının `real_r114` sırasının "
         "TAM tarixçəsi (2023/2024 dəyərləri ilə dəqiq üst-üstə düşür)"),
        ("inv_budget_funds",
         "Büdcə vəsaitləri hesabına yönəldilmiş investisiyalar (min manat)",
         "büdcə vəsaitləri hesabına investisiyalar"),
        ("inv_state_budget_funds",
         "Dövlət büdcəsi vəsaitləri hesabına yönəldilmiş investisiyalar (min manat)",
         "dövlət büdcəsi vəsaitləri hesabına investisiyalar"),
    ]
    out = {}
    for var, col, note in COLS:
        if col not in d.columns:
            raise AssertionError(
                f"ETİKET UYĞUNSUZLUĞU — opendata.az investisiya faylı: {col!r} sütunu yoxdur; "
                f"tapılanlar: {list(d.columns)}")
        vals = {}
        for _, row in d.iterrows():
            y = _year_of(row["Year"])
            v = num(row[col])
            # 1995–1998 üçün fayl SIFIR saxlayır — bu, müşahidə deyil, boşluqdur:
            # həmin illər üçün DSK ayrılıqda göstərici dərc etməyib.
            if y is None or (isinstance(v, float) and math.isnan(v)) or v == 0.0:
                continue
            vals[y] = float(v) / 1000.0        # min manat → mln AZN
        out[var] = (vals, _prov(var, "odazinv", "investments-by-financial-sources",
                                f"sütun «{col}»", col, "mln AZN", vals,
                                note + "; mənbə min manatdadır, mln AZN-ə çevrilib; "
                                       "faylın 1995–1998 sıfırları boşluq kimi oxunur"))
    return out


def _read_nsdp_remittances(path):
    """AMB NSDP `BOP Analytical`: fiziki şəxslərin pul köçürmələri, rüblük → illik."""
    import openpyxl as _ox
    wb = _ox.load_workbook(path, read_only=True, data_only=True)
    ws = wb["BOP Analytical"]
    rows = list(ws.iter_rows(min_row=1, max_row=60, max_col=ws.max_column, values_only=True))
    _assert_label(rows[5][0], "UNIT_MULT", "NSDP BOP Analytical!A6")
    _assert_label(rows[5][2], "Scale = Million", "NSDP BOP Analytical!C6")
    _assert_label(rows[6][2], "Frequency = Quarterly", "NSDP BOP Analytical!C7")
    per = {}
    for j, v in enumerate(rows[9]):
        m = re.fullmatch(r"(\d{4})-Q([1-4])", _squash(v))
        if m:
            per[j] = int(m.group(1))
    for r, code, label in ((46, "AZE_BISRI_BP6_USD", "Remittances of individuals"),
                           (48, "AZE_BISRIR_BP6_USD", "- Receipts"),
                           (49, "AZE_BISRIP_BP6_USD", "- Payments")):
        _assert_label(rows[r - 1][0], code, f"NSDP BOP Analytical!A{r}")
        _assert_label(rows[r - 1][1], label, f"NSDP BOP Analytical!B{r}")
    net, rec, pay = rows[45], rows[47], rows[48]

    def _q(cells, j):
        v = cells[j] if j < len(cells) else None
        return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None

    # Ödənişlər EYNİLİKDƏN çıxarılır: ödəniş = daxilolma − xalis. Mənbənin öz `- Payments`
    # sətri bir rübdə (2024-Q4) daxili eyniliyi POZUR — dəyər 935,059 kimi verilib, halbuki
    # xalis və daxilolma sətirləri −118,459 tələb edir. Fərq gizlədilmir: pozulan rüblər
    # aşağıda sayılır və provenans qeydinə yazılır.
    credit, debit, nq, broken = {}, {}, {}, []
    for j, y in per.items():
        a, b, c = _q(rec, j), _q(net, j), _q(pay, j)
        if a is None or b is None:
            continue
        d = a - b                              # ödəniş (müsbət kəmiyyət)
        if c is not None and abs((-c) - d) > 1e-6:
            broken.append(_squash(rows[9][j]))
        credit[y] = credit.get(y, 0.0) + a
        debit[y] = debit.get(y, 0.0) + d
        nq[y] = nq.get(y, 0) + 1
    credit = {y: v for y, v in credit.items() if nq[y] == 4}   # yalnız TAM illər
    debit = {y: v for y, v in debit.items() if nq[y] == 4}
    base = "; rüblük sıradan yalnız TAM illər (4 rüb) yığılır"
    fix = (f"; mənbənin `- Payments` sətri {len(broken)} rübdə ({', '.join(broken)}) "
           "xalis/daxilolma eyniliyini pozur, ona görə ödəniş EYNİLİKDƏN "
           "(daxilolma − xalis) hesablanır" if broken else "")
    wb.close()
    return {
        "remit_credit_cbar": (credit, _prov(
            "remit_credit_cbar", "nsdpbop", "BOP Analytical", "sətir 48", "- Receipts",
            "mln USD", credit,
            "fiziki şəxslərin pul köçürmələri — DAXİLOLMALAR (kredit)" + base)),
        "remit_debit_cbar": (debit, _prov(
            "remit_debit_cbar", "nsdpbop", "BOP Analytical", "sətir 48 − sətir 46",
            "- Receipts", "mln USD", debit,
            "fiziki şəxslərin pul köçürmələri — ÖDƏNİŞLƏR (debet), müsbət kəmiyyət kimi"
            + base + fix)),
    }


def _read_moe_bop_remittances(path):
    """MOE BOP `Income account`: pul baratları — YALNIZ faktiki sütunlar (≤2024)."""
    import openpyxl as _ox
    wb = _ox.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Income account"]
    rows = list(ws.iter_rows(min_row=1, max_row=90, max_col=45, values_only=True))
    _assert_label(rows[1][4], "Göstəricilər", "MOE BOP Income account!E2")
    ycol = {j: int(v) for j, v in enumerate(rows[2])
            if isinstance(v, (int, float)) and not isinstance(v, bool)
            and 1990 <= v <= 2035 and float(v) == int(v)}
    ROWS = [(74, "Fiziki şəxslərin pul baratları", "remit_credit",
             "təkrar gəlirlərin KREDİT budağı (daxilolmalar) — fiziki şəxslərin pul baratları"),
            (83, "Fərdlərin pul köçürmələri", "remit_debit",
             "təkrar gəlirlərin DEBET budağı (ödənişlər) — fərdlərin pul köçürmələri")]
    out = {}
    for r, label, var, note in ROWS:
        cells = rows[r - 1]
        _moe_spec_assert_not_model(cells[:6], f"MOE BOP Income account!{r}")
        _assert_label(cells[4], label, f"MOE BOP Income account!E{r}")
        vals = {}
        for j, y in ycol.items():
            if y > MOE_LAST_ACTUAL:            # guardrail 1: 2025 «gözlənilən», 2026+ «proqnoz»
                continue
            v = cells[j] if j < len(cells) else None
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                vals[y] = float(v)
        out[var] = (vals, _prov(var, "moebop", "Income account", f"sətir {r}", label,
                                "mln USD", vals,
                                note + f"; YALNIZ faktiki sütunlar (≤{MOE_LAST_ACTUAL})"))
    wb.close()
    return out


def extract_public_panel(dirpath: str = "", moe_dir: str = ""):
    """Bütün açıq mənbələri oxuyub (panel, provenans) qaytarır (etiket yoxlaması ilə).

    Kök qovluqlar verilməyibsə və ya mövcud deyilsə `(None, None)` qaytarır — təhvil
    mühitində mənbə faylları olmaya bilər və bu, xəta deyil.
    """
    root = dirpath or PUBLIC_DATA_DIR
    mroot = moe_dir or MOE_MODEL_DIR
    if not root or not os.path.isdir(root) or not mroot or not os.path.isdir(mroot):
        return None, None
    P = lambda k: public_source_path(k, root, mroot)      # noqa: E731
    got = {}
    got.update(_read_cbar216(P("cbar216")))
    got.update(_read_ssc010(P("ssc010")))
    got.update(_read_ssc013(P("ssc013")))
    got.update(_read_ssc027(P("ssc027")))
    got.update(_read_products(root, mroot))
    got.update(_read_ssc_crosstab(
        P("ssc0458"), "Dynamics", 5, 8, "AZN", "wage_ssc", "4.5–4.8 `Dynamics`",
        "fəaliyyət növü × mülkiyyət forması üzrə orta aylıq nominal əmək haqqı"))
    got.update(_read_ssc_crosstab(
        P("ssc0212"), "Dynamics_2.12", 7, 10, "min nəfər", "emp_ssc", "2.12 `Dynamics_2.12`",
        "fəaliyyət növü × mülkiyyət forması üzrə muzdlu işçilərin sayı"))
    got.update(_read_odaz_investments(P("odazinv")))
    got.update(_read_nsdp_remittances(P("nsdpbop")))
    got.update(_read_moe_bop_remittances(P("moebop")))

    recs, prov = [], []
    for var in sorted(got):
        vals, pr = got[var]
        for y in sorted(vals):
            recs.append({"var": var, "year": int(y), "value": float(vals[y])})
        prov.append(pr)
    panel = pd.DataFrame(recs, columns=["var", "year", "value"])
    prov = pd.DataFrame(prov, columns=["var", "source_key", "source_file", "source_name",
                                       "table", "locator", "row_label", "unit",
                                       "year_first", "year_last", "n", "retrieved", "note_az"])
    return panel, prov


def rebuild_public_panel(dirpath: str = "", moe_dir: str = "", verbose: bool = True):
    """Paneli mənbələrdən yenidən qurub `data/` altına yazır (builder-in çağırdığı funksiya)."""
    panel, prov = extract_public_panel(dirpath, moe_dir)
    if panel is None:
        raise FileNotFoundError(
            "açıq mənbə kökü tapılmadı; `PUBLIC_DATA_DIR` və `MOE_MODEL_DIR` təyin edin")
    panel.to_csv(P_PUBLIC_PANEL, index=False)
    prov.to_csv(P_PUBLIC_PROV, index=False)
    if verbose:
        print(f"public_sources_panel: {len(panel)} sətir, {panel['var'].nunique()} sıra "
              f"→ {P_PUBLIC_PANEL}")
    return panel, prov


def public_panel() -> pd.DataFrame:
    """Rəsmi açıq mənbələrin dondurulmuş paneli. Sütunlar: var, year, value."""
    if not os.path.exists(P_PUBLIC_PANEL):
        raise FileNotFoundError(
            f"{P_PUBLIC_PANEL} yoxdur — açıq mənbə builder skripti ilə qurulmalıdır")
    return pd.read_csv(P_PUBLIC_PANEL, float_precision="round_trip")


def public_provenance() -> pd.DataFrame:
    """Panelin sətir-sətir provenansı (fayl, cədvəl, sətir/sütun, etiket, dövr, tarix)."""
    if not os.path.exists(P_PUBLIC_PROV):
        raise FileNotFoundError(f"{P_PUBLIC_PROV} yoxdur")
    return pd.read_csv(P_PUBLIC_PROV)


def public_series(var: str) -> pd.Series:
    """Bir açıq-mənbə sırasını il indeksi ilə qaytarır."""
    p = public_panel()
    s = p[p["var"] == var]
    return pd.Series(s["value"].values, index=s["year"].astype(int).values,
                     name=var, dtype=float).sort_index()


def verify_public_panel(dirpath: str = "", moe_dir: str = "", tol: float = 1e-9):
    """Dondurulmuş paneli MƏNBƏLƏRDƏN yenidən oxuyur və maksimal fərqi qaytarır.

    Fayllar yoxdursa `None` qaytarır — bu, xəta deyil.
    """
    got, _ = extract_public_panel(dirpath, moe_dir)
    if got is None:
        return None
    want = public_panel()
    m = want.merge(got, on=["var", "year"], how="outer", suffixes=("_disk", "_src"))
    miss = m[m["value_disk"].isna() | m["value_src"].isna()]
    if len(miss):
        raise AssertionError(
            f"açıq mənbə paneli fərqlənir: {len(miss)} sətir yalnız bir tərəfdə "
            f"(ilk: {miss.iloc[0]['var']} {int(miss.iloc[0]['year'])})")
    diff = float((m["value_disk"] - m["value_src"]).abs().max())
    if diff > tol:
        raise AssertionError(f"açıq mənbə paneli fərqlənir: maks. fərq {diff:.3e}")
    return diff


# --- rahat interfeyslər (dəftərlər üçün) --------------------------------------------------

def _panel_wide(prefix, roots, suffixes):
    p = public_panel()
    cols = {}
    for root in roots:
        for suf in suffixes:
            var = f"{prefix}_{root}{suf}"
            s = p[p["var"] == var]
            if len(s):
                cols[var] = pd.Series(s["value"].values,
                                      index=s["year"].astype(int).values, dtype=float)
    return pd.DataFrame(cols).sort_index()


def ssc_production_account() -> pd.DataFrame:
    """DSK 13 — NACE bölmələri üzrə buraxılış / aralıq istehlak / əlavə dəyər (2005–2025)."""
    letters = [lt.lower() for lt, _l, _r in SSC013_SECTIONS] + ["total"]
    return _panel_wide("", [""], [f"{m}_ssc_{x}" for m in ("go", "ic", "va") for x in letters])


def ssc_branch_wage_ownership() -> pd.DataFrame:
    """DSK 4.5–4.8 — 19 fəaliyyət növü × dövlət/qeyri-dövlət orta aylıq əmək haqqı (2005–2024)."""
    return _panel_wide("wage_ssc", [r for r, _ in SSC_ACTIVITY_ROWS],
                       [s for s, _ in SSC_OWNERSHIP])


def ssc_branch_employment_ownership() -> pd.DataFrame:
    """DSK 2.12 — 19 fəaliyyət növü × dövlət/qeyri-dövlət muzdlu işçi sayı (2005–2024)."""
    return _panel_wide("emp_ssc", [r for r, _ in SSC_ACTIVITY_ROWS],
                       [s for s, _ in SSC_OWNERSHIP])


def vereq_product_concordance() -> pd.DataFrame:
    """Statutar şablonun 11 ölü (`#REF!`) məhsul sətri ilə DSK sıralarının konkordansı.

    Sübut sənədinin (vii) bəndinin tələb etdiyi cədvəl: hər şablon sətrinin adı, vahidi,
    tapılan DSK mənbəyi, onun öz etiketi/vahidi, çevirmə əmsalı, əhatə və uyğunluq statusu.
    """
    prov = public_provenance().set_index("var")
    rows = []
    for (vrow, vlabel, vunit, code, key, _loc, label, sunit, factor, status) in VEREQ_DEAD_PRODUCTS:
        pr = prov.loc[code] if code in prov.index else None
        rows.append(dict(
            vereq_sheet="2.4.1.4.", vereq_row=vrow, vereq_name_az=vlabel, vereq_unit=vunit,
            series_code=code, source_file=PUBLIC_FILES[key][0],
            source_name=PUBLIC_FILES[key][1], source_label=label, source_unit=sunit,
            factor=factor, unit_out=(vunit if factor else sunit),
            year_first=(None if pr is None else pr["year_first"]),
            year_last=(None if pr is None else pr["year_last"]),
            n=(0 if pr is None else int(pr["n"])), status_az=status))
    return pd.DataFrame(rows)


def fx_cross_check(tol: float = 5e-4) -> pd.DataFrame:
    """AMB 2.16 ilə DSK 10-un orta illik məzənnələrinin müstəqil tutuşdurulması.

    DSK sırası dörd onluq işarə ilə dərc olunduğu üçün müqayisə həmin dəqiqliyə
    yuvarlaqlaşdırılır; sübut sənədinin (xv) bəndi 2025 üçün 1,9210 dəqiq uyğunluğunu
    qeyd edir və bu funksiya onu ölçüb qaytarır.
    """
    rows = []
    for a, b, name in (("fx_usd_azn_avg", "fx_usd_azn_ssc", "AZN/USD"),
                       ("fx_eur_azn", "fx_eur_azn_ssc", "AZN/EUR")):
        sa, sb = public_series(a), public_series(b)
        common = sorted(set(sa.index) & set(sb.index))
        d = (sa.loc[common].round(4) - sb.loc[common]).abs()
        rows.append(dict(
            pair=name, cbar=a, ssc=b, n_overlap=len(common),
            n_within_tol=int((d <= tol).sum()),
            max_abs_diff=float(d.max()) if len(d) else float("nan"),
            max_diff_year=(int(d.idxmax()) if len(d) else None),
            diff_2025=(float(d.loc[LAST_ACTUAL]) if LAST_ACTUAL in d.index else None),
            diff_2015=(float(d.loc[2015]) if 2015 in d.index else None)))
    return pd.DataFrame(rows)


def remittances_annual() -> pd.DataFrame:
    """Pul baratları: MOE BOP (1996–2024) və AMB NSDP (2013–2025) yanaşı, mln USD."""
    return _panel_wide("remit", ["credit", "debit"], ["", "_cbar"])


# =========================================================================================
# 15. GÖMRÜK BÜLLETENLƏRİ — HS-27 (xam neft, neft məhsulları, qaz) İDXAL/İXRAC — bu buraxılış
#     `DATA_GAP_NEGATIVE_PROOF.md` (xi) bəndinin AÇILMAMIŞ aparıcı ucu: DGK-nın
#     `Bulleten_2016.rar` … `Bulleten_2025.rar` illik bülletenləri.
# =========================================================================================
# Arxivlər RAR5 formatındadır; `bsdtar` (libarchive, macOS-da hazır) onları AÇIR — heç bir
# proqram quraşdırılmır. `Cədvəl 7` («mallar və ölkələr üzrə ixrac və idxal») HS-4 rəqəmli
# mövqe üzrə həm MİQDAR, həm STATİSTİK DƏYƏR verir; sübut sənədinin (xi) bəndinin
# "yalnız dəyər var, miqdar yoxdur" hökmü məhz bununla bağlanır.
#
# Ölçü vahidləri mənbədə: miqdar kq / 1000 l / m³; dəyər min ABŞ dolları.
# Panelə çevrilmiş halda: miqdar min ton (kq ÷ 1e6) və ya mln m³ (m³ ÷ 1e6);
# dəyər mln USD (min USD ÷ 1e3); vahid qiymət USD/ton, USD/min m³.
#
# İl bazası: mövcud olduqda İLLİK («illik») bülleten; olmadıqda dörd rübün CƏMİ. Metod
# 2024-cü ildə YOXLANILIR (hər iki buraxılış mövcuddur) və nəticə provenansda göstərilir.

CUSTOMS_RETRIEVED = "2026-06-03"
CUSTOMS_RAR_DIR = os.environ.get(
    "CUSTOMS_RAR_DIR", "")      # `…/data_collection_20260603/raw/customs/statistics_bulletin/rar`

CUSTOMS_SOURCE_NOTE = (
    "Azərbaycan Respublikası Dövlət Gömrük Komitəsi (DGK), «Gömrük statistikası bülleteni», "
    "Cədvəl 7 «Azərbaycan Respublikasının mallar və ölkələr üzrə ixracı və idxalı». "
    "Arxivlər: Bulleten_2016.rar … Bulleten_2025.rar (10 fayl), `bsdtar` ilə açılıb. "
    f"Anlıq nüsxə {CUSTOMS_RETRIEVED}")

# ÖLÇÜLMÜŞ ƏHATƏ QIRILMASI — gizlədilmir, provenansda və `customs_cross_check()`-də
# göstərilir. DGK-nın 2016-cı il bülleteni ümumi İXRACI 9 340,2 mln USD kimi verir, halbuki
# iş kitabının (DSK/AMB buraxılışı) həmin il üçün göstəricisi 13 457,6 mln USD-dir; fərqin
# demək olar hamısı xam neft sətrindədir (6 575,4 vs 10 692,8 mln USD). 2017-ci ildən
# bülletenin xam neft sətri iş kitabı ilə TAM üst-üstə düşür (nisbi fərq < 1e-5), İDXAL
# tərəfi isə 2016-da da uyğundur (8 542,5 vs 8 489,1 mln USD — buraxılış düzəlişi həddində).
# Ona görə İXRAC sıraları üçün müqayisə bazası 2017-dən başlayır; 2016 dəyəri panelə daxildir,
# lakin əhatə qırılması kimi işarələnir.
CUSTOMS_EXPORT_BREAK_YEAR = 2016
CUSTOMS_EXPORT_BREAK_NOTE = (
    "ƏHATƏ QIRILMASI: DGK-nın 2016-cı il bülletenində ixrac sətirləri DSK/AMB buraxılışından "
    "əhəmiyyətli dərəcədə aşağıdır (xam neft: 6 575,4 vs 10 692,8 mln USD); 2017-dən "
    "sıralar tam üst-üstə düşür. İdxal tərəfində belə qırılma YOXDUR")

CUSTOMS_HS_LABELS = {
    "2709": "Xam neft və bitumlu minerallardan alınan xam neft məhsulları",
    "2710": "Bitumlu minerallardan alınmış neft və neft məhsulları, xam məhsullar istisna "
            "olmaqla",
    "2711": "Neft qazları və digər qaz halında karbohidrogenlər",
}

# (kod, HS, axın, ölçü, mənbə vahidi, çevirici, panel vahidi, AZ ad, EN ad)
CUSTOMS_SPEC = [
    ("cust_exp_qty_crude", "2709", "exp", "qty", "kq", 1e-6, "min ton",
     "Xam neft ixracı — miqdar (DGK)", "Crude oil exports — quantity (customs)"),
    ("cust_exp_val_crude", "2709", "exp", "val", "min USD", 1e-3, "mln USD",
     "Xam neft ixracı — statistik dəyər (DGK)", "Crude oil exports — value (customs)"),
    ("cust_imp_qty_crude", "2709", "imp", "qty", "kq", 1e-6, "min ton",
     "Xam neft idxalı — miqdar (DGK)", "Crude oil imports — quantity (customs)"),
    ("cust_imp_val_crude", "2709", "imp", "val", "min USD", 1e-3, "mln USD",
     "Xam neft idxalı — statistik dəyər (DGK)", "Crude oil imports — value (customs)"),
    ("cust_exp_qty_petprod", "2710", "exp", "qty", "kq", 1e-6, "min ton",
     "Neft məhsullarının ixracı — miqdar (DGK, kq bölməsi)",
     "Petroleum products exports — quantity (customs, kg leg)"),
    ("cust_exp_val_petprod", "2710", "exp", "val", "min USD", 1e-3, "mln USD",
     "Neft məhsullarının ixracı — dəyər (DGK, kq bölməsi)",
     "Petroleum products exports — value (customs, kg leg)"),
    ("cust_imp_qty_petprod", "2710", "imp", "qty", "kq", 1e-6, "min ton",
     "Neft məhsullarının idxalı — miqdar (DGK, kq bölməsi)",
     "Petroleum products imports — quantity (customs, kg leg)"),
    ("cust_imp_val_petprod", "2710", "imp", "val", "min USD", 1e-3, "mln USD",
     "Neft məhsullarının idxalı — dəyər (DGK, kq bölməsi)",
     "Petroleum products imports — value (customs, kg leg)"),
    ("cust_exp_qty_gas", "2711", "exp", "qty", "m3", 1e-6, "mln m3",
     "Qaz ixracı — miqdar (DGK)", "Gas exports — quantity (customs)"),
    ("cust_exp_val_gas", "2711", "exp", "val", "min USD", 1e-3, "mln USD",
     "Qaz ixracı — statistik dəyər (DGK)", "Gas exports — value (customs)"),
    ("cust_imp_qty_gas", "2711", "imp", "qty", "m3", 1e-6, "mln m3",
     "Qaz idxalı — miqdar (DGK)", "Gas imports — quantity (customs)"),
    ("cust_imp_val_gas", "2711", "imp", "val", "min USD", 1e-3, "mln USD",
     "Qaz idxalı — statistik dəyər (DGK)", "Gas imports — value (customs)"),
]

# Vahid qiymətlər (törəmə): (kod, miqdar kodu, dəyər kodu, əmsal, vahid, AZ ad, EN ad)
CUSTOMS_UNIT_VALUES = [
    ("cust_exp_uv_crude", "cust_exp_qty_crude", "cust_exp_val_crude", 1e3, "USD/ton",
     "Xam neft ixracının vahid qiyməti (DGK)", "Crude oil export unit value (customs)"),
    ("cust_imp_uv_crude", "cust_imp_qty_crude", "cust_imp_val_crude", 1e3, "USD/ton",
     "Xam neft idxalının vahid qiyməti (DGK)", "Crude oil import unit value (customs)"),
    ("cust_exp_uv_petprod", "cust_exp_qty_petprod", "cust_exp_val_petprod", 1e3, "USD/ton",
     "Neft məhsulları ixracının vahid qiyməti (DGK)",
     "Petroleum products export unit value (customs)"),
    ("cust_imp_uv_petprod", "cust_imp_qty_petprod", "cust_imp_val_petprod", 1e3, "USD/ton",
     "Neft məhsulları idxalının vahid qiyməti (DGK)",
     "Petroleum products import unit value (customs)"),
    ("cust_exp_uv_gas", "cust_exp_qty_gas", "cust_exp_val_gas", 1e3, "USD/min m3",
     "Qaz ixracının vahid qiyməti (DGK)", "Gas export unit value (customs)"),
    ("cust_imp_uv_gas", "cust_imp_qty_gas", "cust_imp_val_gas", 1e3, "USD/min m3",
     "Qaz idxalının vahid qiyməti (DGK)", "Gas import unit value (customs)"),
]

_CUST_NOISE = re.compile(
    r"Dövlət Gömrük Komitəsi|cədvəlin ardı|XİF|^\s*MN\s*$|^\s*üzrə\s*$|^\s*kodu\s*$|"
    r"Mal mövqeyinin adı|Ölçü vahidi|Miqdar|rüb|ci il|cı il|cü il|cu il|^\s*$", re.M)
_CUST_UNIT = re.compile(r"^(kq|1000 l|1000 kvt|m3|ton|əd)$")
_CUST_ROMAN = {"I": 1, "II": 2, "III": 3, "IV": 4}


def _cust_fields(line):
    return [f.strip() for f in re.split(r"\s{2,}", line.replace("\xa0", " ").strip()) if f.strip()]


def _cust_num(tok):
    t = tok.replace(" ", "").replace(" ", "")
    if t in ("-", "—", "–"):
        return ("nil", None)
    return ("num", float(t)) if re.fullmatch(r"\d+(?:\.\d+)?", t) else None


def customs_period(text):
    """Bülletenin dövrü: (il, rüb) — illik buraxılışda rüb `None`."""
    m = re.search(r"(\d{4})-(?:cı|ci|cu|cü)\s+ilin\s+([IVX]+)\s*r[üu]b", text)
    if m:
        return int(m.group(1)), _CUST_ROMAN.get(m.group(2))
    m = re.search(r"(\d{4})-(?:cı|ci|cu|cü)\s+il\b", text)
    return (int(m.group(1)), None) if m else (None, None)


def parse_customs_hs(text, hs):
    """Bir HS mövqeyinin `Cəmi` bloku: {ölçü vahidi -> [ixrac miqdar, ixrac dəyər,
    idxal miqdar, idxal dəyər]}; vahidsiz ümumi sətir `__total__` açarı ilə."""
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if not re.match(r"^\s{0,4}%s\s{2,}\S" % hs, ln):
            continue
        # Etiket yoxlaması: HS mövqeyinin adı gözlənilən mətnlə başlamalıdır.
        head = _squash(re.sub(r"^\s*%s\s+" % hs, "", ln))
        want = CUSTOMS_HS_LABELS[hs]
        if not _key(head).startswith(_key(want)[:40]):
            raise AssertionError(
                f"ETİKET UYĞUNSUZLUĞU — DGK Cədvəl 7, HS {hs}: gözlənilən {want!r}, "
                f"tapılan {head[:80]!r}")
        out, started, j = {}, False, i + 1
        while j < len(lines) and j < i + 40:
            s = lines[j]
            j += 1
            if re.search(r"o c[üu]ml[əe]d[əe]n", s):
                break
            if re.match(r"^\s{0,4}\d{4}\s{2,}\S", s):
                break
            f = _cust_fields(s)
            if not f:
                continue
            if _CUST_NOISE.search(s) and not started:
                continue
            if f[0].startswith("Cəmi"):
                started = True
                rest = f[1:]
                unit = rest[0] if rest and _CUST_UNIT.match(rest[0]) else None
                toks = [_cust_num(t) for t in (rest[1:] if unit else rest)]
                if any(t is None for t in toks):
                    continue
                out[unit or "__total__"] = [(v if k == "num" else None) for k, v in toks]
                continue
            if not started:
                continue
            if _CUST_UNIT.match(f[0]):
                toks = [_cust_num(t) for t in f[1:]]
                if any(t is None for t in toks):
                    continue
                out[f[0]] = [(v if k == "num" else None) for k, v in toks]
                continue
            if _CUST_NOISE.search(s):
                continue
            break
        return out
    return {}


def _customs_extract_archives(rar_dir, workdir, verbose=True):
    """RAR arxivlərini `bsdtar` ilə (iç-içə arxivlər daxil) `workdir`-ə açır.

    `bsdtar` macOS və əksər Linux paylanmalarında hazır gəlir və RAR5-i dəstəkləyir; heç
    bir proqram QURAŞDIRILMIR. Açılan fayllar müvəqqəti qovluqdadır — paketə yazılmır.
    """
    import glob
    import subprocess
    os.makedirs(workdir, exist_ok=True)
    tar = shutil.which("bsdtar") or "/usr/bin/bsdtar"
    if not os.path.exists(tar):
        raise FileNotFoundError("bsdtar tapılmadı — RAR arxivləri açıla bilmir")
    archives = sorted(glob.glob(os.path.join(rar_dir, "Bulleten_*.rar")))
    if not archives:
        raise FileNotFoundError(f"{rar_dir}: `Bulleten_*.rar` arxivi tapılmadı")
    for a in archives:
        subprocess.run([tar, "-xf", a], cwd=workdir, check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(4):                      # iç-içə arxivlər (2016–2023 buraxılışları)
        nested = [p for p in glob.glob(os.path.join(workdir, "**", "*.rar"), recursive=True)]
        if not nested:
            break
        for p in nested:
            d = os.path.join(os.path.dirname(p), "_x_" + os.path.basename(p)[:-4])
            os.makedirs(d, exist_ok=True)
            subprocess.run([tar, "-xf", p], cwd=d, check=False,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            os.remove(p)
    pdfs = glob.glob(os.path.join(workdir, "**", "*.pdf"), recursive=True)
    if verbose:
        print(f"gömrük arxivləri açıldı: {len(archives)} arxiv → {len(pdfs)} PDF")
    return sorted(pdfs)


def extract_customs_hs27(rar_dir: str = "", workdir: str = "", verbose: bool = True):
    """DGK bülletenlərindən HS-27 illik panelini qurur. (panel, provenans) qaytarır.

    Arxiv qovluğu verilməyibsə və ya mövcud deyilsə `(None, None)` qaytarır — bu, xəta deyil.
    """
    import subprocess
    import tempfile
    rar_dir = rar_dir or CUSTOMS_RAR_DIR
    if not rar_dir or not os.path.isdir(rar_dir):
        return None, None
    pdftotext = shutil.which("pdftotext")
    if not pdftotext:
        raise FileNotFoundError("pdftotext tapılmadı — bülletenlərin mətni oxuna bilmir")
    tmp = workdir or tempfile.mkdtemp(prefix="dgk_bulleten_")
    try:
        pdfs = _customs_extract_archives(rar_dir, tmp, verbose=verbose)
        seen, per = {}, {}       # (il, rüb) -> {hs: bloklar}; təkrar nüsxələr atılır
        files = {}
        for p in pdfs:
            if "Cədvəl 7" not in os.path.basename(p):
                continue
            h = _md5(p)
            if h in seen:
                continue
            seen[h] = p
            txt = subprocess.run([pdftotext, "-layout", p, "-"],
                                 capture_output=True, text=True).stdout
            y, q = customs_period(txt)
            if y is None:
                continue
            blocks = {hs: parse_customs_hs(txt, hs) for hs in CUSTOMS_HS_LABELS}
            if not any(blocks.values()):
                continue
            per[(y, q)] = blocks
            files[(y, q)] = os.path.relpath(p, tmp)
    finally:
        if not workdir:
            shutil.rmtree(tmp, ignore_errors=True)

    def leg(blocks, hs, unit):
        d = blocks.get(hs) or {}
        return d.get(unit)

    def pick(blocks, code, hs, flow, meas):
        unit = {"2709": "kq", "2710": "kq", "2711": "m3"}[hs]
        row = leg(blocks, hs, unit)
        if row is None or len(row) != 4:
            return None
        idx = {("exp", "qty"): 0, ("exp", "val"): 1,
               ("imp", "qty"): 2, ("imp", "val"): 3}[(flow, meas)]
        return row[idx]

    years = sorted({y for y, _ in per})
    recs, basis = [], {}
    for y in years:
        ann = per.get((y, None))
        qs = [per.get((y, k)) for k in (1, 2, 3, 4)]
        use_annual = ann is not None
        basis[y] = "illik buraxılış" if use_annual else "dörd rübün cəmi"
        for (code, hs, flow, meas, _su, factor, _pu, _az, _en) in CUSTOMS_SPEC:
            if use_annual:
                v = pick(ann, code, hs, flow, meas)
            else:
                if any(b is None for b in qs):
                    continue
                parts = [pick(b, code, hs, flow, meas) for b in qs]
                v = None if any(p is None for p in parts) else sum(parts)
            if v is None:
                continue
            recs.append({"var": code, "year": y, "value": float(v) * factor})
    panel = pd.DataFrame(recs, columns=["var", "year", "value"])

    # Vahid qiymətlər — miqdar və dəyər eyni ildə mövcud olduqda törədilir.
    wide = panel.pivot(index="year", columns="var", values="value") if len(panel) else pd.DataFrame()
    uv = []
    for code, qcode, vcode, mult, _u, _az, _en in CUSTOMS_UNIT_VALUES:
        if qcode not in wide.columns or vcode not in wide.columns:
            continue
        q, v = wide[qcode], wide[vcode]
        r = (v / q.where(q > 0)) * mult
        for y, x in r.dropna().items():
            uv.append({"var": code, "year": int(y), "value": float(x)})
    panel = pd.concat([panel, pd.DataFrame(uv, columns=["var", "year", "value"])],
                      ignore_index=True) if uv else panel
    panel = panel.sort_values(["var", "year"]).reset_index(drop=True)

    prov = []
    src_files = "; ".join(f"{y}: {files.get((y, None)) or 'I–IV rüb'}" for y in years)
    for (code, hs, flow, meas, su, factor, pu, az, en) in CUSTOMS_SPEC:
        s = panel[panel["var"] == code]
        ys = sorted(s["year"].astype(int))
        prov.append(dict(
            var=code, hs_code=hs, hs_label=CUSTOMS_HS_LABELS[hs], flow=flow, measure=meas,
            source_unit=su, unit=pu, name_az=az, name_en=en,
            table="Cədvəl 7", row_label=CUSTOMS_HS_LABELS[hs],
            source_key="dgk_bulleten", source_file="Bulleten_2016.rar … Bulleten_2025.rar",
            year_first=(ys[0] if ys else None), year_last=(ys[-1] if ys else None),
            n=len(ys), retrieved=CUSTOMS_RETRIEVED,
            basis_az="; ".join(f"{y}: {basis[y]}" for y in years),
            note_az=(f"DGK Cədvəl 7, HS {hs} `Cəmi` sətri, `{su}` bölməsi; "
                     f"mənbə vahidi {su} → {pu} (əmsal {factor:g})"
                     + (f". {CUSTOMS_EXPORT_BREAK_NOTE}" if flow == "exp" else ""))))
    for code, qcode, vcode, mult, u, az, en in CUSTOMS_UNIT_VALUES:
        s = panel[panel["var"] == code]
        ys = sorted(s["year"].astype(int))
        prov.append(dict(
            var=code, hs_code=[r["hs_code"] for r in prov if r["var"] == qcode][0],
            hs_label="", flow=("exp" if "_exp_" in code else "imp"), measure="uv",
            source_unit="", unit=u, name_az=az, name_en=en, table="Cədvəl 7",
            row_label=CUSTOMS_HS_LABELS[[r["hs_code"] for r in prov if r["var"] == qcode][0]],
            source_key="dgk_bulleten", source_file="Bulleten_2016.rar … Bulleten_2025.rar",
            year_first=(ys[0] if ys else None), year_last=(ys[-1] if ys else None),
            n=len(ys), retrieved=CUSTOMS_RETRIEVED,
            basis_az="törəmə", note_az=f"vahid qiymət = {vcode} ÷ {qcode} × {mult:g}"))
    prov = pd.DataFrame(prov)
    prov.attrs["source_files"] = src_files
    return panel, prov


def rebuild_customs_hs27(rar_dir: str = "", verbose: bool = True):
    """HS-27 panelini arxivlərdən yenidən qurub `data/` altına yazır."""
    panel, prov = extract_customs_hs27(rar_dir, verbose=verbose)
    if panel is None:
        raise FileNotFoundError(
            "gömrük arxivi qovluğu tapılmadı; `CUSTOMS_RAR_DIR` mühit dəyişənini təyin edin")
    panel.to_csv(P_CUSTOMS_PANEL, index=False)
    prov.to_csv(P_CUSTOMS_PROV, index=False)
    if verbose:
        print(f"customs_hs27_annual: {len(panel)} sətir, {panel['var'].nunique()} sıra "
              f"→ {P_CUSTOMS_PANEL}")
    return panel, prov


def verify_customs_hs27(rar_dir: str = "", tol: float = 1e-9):
    """Dondurulmuş gömrük panelini ARXİVLƏRDƏN yenidən qurub tutuşdurur (yoxsa `None`)."""
    got, _ = extract_customs_hs27(rar_dir, verbose=False)
    if got is None:
        return None
    want = customs_hs27()
    m = want.merge(got, on=["var", "year"], how="outer", suffixes=("_disk", "_src"))
    miss = m[m["value_disk"].isna() | m["value_src"].isna()]
    if len(miss):
        raise AssertionError(
            f"gömrük paneli fərqlənir: {len(miss)} sətir yalnız bir tərəfdə "
            f"(ilk: {miss.iloc[0]['var']} {int(miss.iloc[0]['year'])})")
    diff = float((m["value_disk"] - m["value_src"]).abs().max())
    if diff > tol:
        raise AssertionError(f"gömrük paneli fərqlənir: maks. fərq {diff:.3e}")
    return diff


def customs_hs27() -> pd.DataFrame:
    """DGK bülletenlərindən çıxarılmış HS-27 illik paneli. Sütunlar: var, year, value."""
    if not os.path.exists(P_CUSTOMS_PANEL):
        raise FileNotFoundError(
            f"{P_CUSTOMS_PANEL} yoxdur — gömrük builder skripti ilə qurulmalıdır")
    return pd.read_csv(P_CUSTOMS_PANEL, float_precision="round_trip")


def customs_hs27_provenance() -> pd.DataFrame:
    """HS-27 panelinin provenansı (HS mövqeyi, axın, ölçü, il bazası, mənbə buraxılışları)."""
    if not os.path.exists(P_CUSTOMS_PROV):
        raise FileNotFoundError(f"{P_CUSTOMS_PROV} yoxdur")
    return pd.read_csv(P_CUSTOMS_PROV)


def customs_hs27_series(var: str) -> pd.Series:
    p = customs_hs27()
    s = p[p["var"] == var]
    return pd.Series(s["value"].values, index=s["year"].astype(int).values,
                     name=var, dtype=float).sort_index()


def customs_cross_check() -> pd.DataFrame:
    """DGK panelinin əsas iş kitabının DGK-mənbəli sıraları ilə tutuşdurulması.

    İş kitabının `Neft-Qaz sektoru` vərəqindəki neft/qaz ixracı sıraları DGK mənbəlidir,
    ona görə müstəqil deyil, LAKİN oxunuşun düzgünlüyünü sübut edən ən sərt yoxlamadır:
    fərq sıfıra yaxın olmalıdır.
    """
    pairs = [("cust_exp_qty_crude", "oilgas_r005", 1e-3, "mln ton"),      # min ton → mln ton
             ("cust_exp_val_crude", "oilgas_r006", 1.0, "mln USD"),
             ("cust_exp_qty_gas", "oilgas_r014", 1e-3, "mlrd m3"),        # mln m³ → mlrd m³
             ("cust_exp_val_gas", "oilgas_r015", 1.0, "mln USD")]
    rows = []
    for var, wb_code, factor, unit in pairs:
        a = customs_hs27_series(var) * factor
        b = wb_annual(wb_code)
        common = sorted(set(a.index) & set(b.index))
        if not common:
            continue
        rel = ((a.loc[common] - b.loc[common]).abs()
               / b.loc[common].abs().clip(lower=1e-9))
        post = rel[rel.index > CUSTOMS_EXPORT_BREAK_YEAR]
        rows.append(dict(customs_var=var, wb_series=wb_code, unit=unit, n_overlap=len(common),
                         max_rel_diff=float(rel.max()), max_diff_year=int(rel.idxmax()),
                         max_rel_diff_post_break=(float(post.max()) if len(post) else None),
                         post_break_year=(int(post.idxmax()) if len(post) else None)))
    return pd.DataFrame(rows)


# =========================================================================================
# 16. BVF (Beynəlxalq Valyuta Fondu, beynəlxalq mətnlərdə IMF) ETALON YOLLARI —
#     Maddə IV (26/112) və WEO — v3 R23
#     Müqavilə üzrə FR8-in 7-ci alt-tapşırığı «beynəlxalq təşkilatların proqnozları ilə
#     müqayisə» tələb edir. Bu bölmə həmin müqayisənin YEGANƏ giriş nöqtəsidir.
# =========================================================================================
# QAYDA (D1-in davamı). BVF yolları modelin heç bir yerinə girmir: nə bir tənliyin nümunəsinə,
# nə bir fərziyyə açarına, nə də bir kalibrləmə hədəfinə. `actuals()` və `ministry_baseline()`
# kimi AYRI obyektdə saxlanılır və yalnız `source=imf_reference` etiketi ilə müqayisə sətri
# şəklində nəşr olunur. Ona görə bu bölmənin bütün funksiyaları YALNIZ OXUYUR.
#
# BURAXILIŞ (vintaj) QEYDİ. Rəqəmlər BVF-nin Azərbaycan üzrə Maddə IV məsləhətləşməsinin
# ölkə hesabatından (Country Report 26/112, Cədvəl 1) və eyni buraxılışın WEO yolundan
# götürülüb. Bu, MÜƏYYƏN bir buraxılışdır: BVF hər yeni raundda rəqəmlərini yeniləyir, ona görə
# müqayisə oxunarkən buraxılış tarixi mütləq yanında göstərilməlidir (aşağıdakı sabit).
#
# ƏHATƏ. `data/reference_forecasts.json` DONDURULMUŞ ÇIXARIŞDIR: yalnız BVF-nin iki bloku
# saxlanılır (faylın `_provenance` bölməsi mənbəni və buraxılışı adı ilə göstərir). Başqa
# modellərin (məsələn Nazirliyin CAEM iş kitabının) və köhnə mühərrikin arxivləşdirilmiş yolları
# fayla QƏSDƏN salınmayıb — onlar bu paketin rəqəmləri ilə yan-yana durarsa İKİNCİ bir "kanonik"
# dəst təəssüratı yaradardı. Aşağıdakı ağ siyahı həmin qərarı KODDA da bağlayır: siyahıda olmayan
# blok, faylda görünsə belə, oxunmur.
REFERENCE_VINTAGE_NOTE = "BVF Maddə IV 26/112, aprel 2026 vintajı"
REFERENCE_SOURCE_NOTE = (
    "Beynəlxalq Valyuta Fondu (BVF), Azərbaycan üzrə Maddə IV məsləhətləşməsi — ölkə hesabatı "
    "26/112, Cədvəl 1; eyni buraxılışın «World Economic Outlook» (WEO) yolu ilə birlikdə. "
    f"Buraxılış: {REFERENCE_VINTAGE_NOTE}")

# Ağ siyahı: blok açarı -> (qısa açar, AZ etiket). Siyahıda olmayan blok OXUNMUR.
REFERENCE_BLOCKS = {
    "IMF_ArticleIV_26_112": ("artiv", "BVF Maddə IV (26/112)"),
    "IMF_WEO": ("weo", "BVF WEO"),
}
REFERENCE_PRIMARY_BLOCK = "IMF_ArticleIV_26_112"   # eyni sıra hər iki blokda varsa bu üstündür

# Bu paketin sıra kodları ilə ÜST-ÜSTƏ DÜŞƏN BVF göstəriciləri. Yalnız bunlar nəşr olunur:
# adı oxşayan, tərifi fərqli göstəricilər (məsələn WEO-nun `fiscbal_pgdp`, `ca_pgdp`,
# `grossdebt_pgdp` sıraları) bu paketdə eyni tərifli qarşılığa malik olmadığı üçün
# BURAXILIR — uydurma uyğunluq yaradılmır (bax `reference_match_report`).
REFERENCE_SERIES = {
    "gdp_realg": ("ÜDM real artım tempi", "%"),
    "nonoil_realg": ("Qeyri neft-qaz sektorunun real artım tempi", "%"),
    "cpi_infl": ("İnflyasiya (İQİ, orta illik)", "%"),
}
# Fayldakı, lakin bu paketdə eyni tərifli qarşılığı OLMAYAN BVF sıraları — susmaq əvəzinə
# adları ilə birlikdə saxlanılır ki, `reference_match_report()` onları açıq sadalaya bilsin.
REFERENCE_UNMATCHED_NOTE = {
    "fiscbal_pgdp": "BVF-nin ümumi hökumət balansı (ÜDM-ə nisbətdə); bu paketin fiskal bloku "
                    "(FR13) dövlət büdcəsi əhatəsindədir, ümumi hökumət əhatəsində deyil",
    "ca_pgdp": "BVF-nin cari hesab nisbəti; bu paketin `ca_gdp_ratio` sırası öz ÜDM və "
               "məzənnə bazası ilə hesablanır — eyni tərif deyil, avtomatik tutuşdurulmur",
    "grossdebt_pgdp": "BVF-nin ümumi dövlət borcu (ÜDM-ə nisbətdə); bu paketdə borc sırası "
                      "modelləşdirilmir",
}

_REF_CACHE = {}


def reference_forecasts() -> dict:
    """`data/reference_forecasts.json` faylının ağ siyahı ilə süzülmüş oxunuşu.

    Qaytarılan quruluş: `{blok_açarı: {sıra_kodu: {il: dəyər}}}` — YALNIZ `REFERENCE_BLOCKS`
    siyahısındakı bloklar. Fayl dəyişdirilmir; modul onu heç vaxt yazmır.
    """
    if "data" in _REF_CACHE:
        return _REF_CACHE["data"]
    if not os.path.exists(P_REFERENCE_FORECASTS):
        raise FileNotFoundError(
            f"{P_REFERENCE_FORECASTS} yoxdur — BVF etalon yolları oxuna bilmir")
    with open(P_REFERENCE_FORECASTS, encoding="utf-8") as fh:
        raw = json.load(fh)
    out = {}
    for block in REFERENCE_BLOCKS:                 # sıra ağ siyahı ilə müəyyən olunur (sabit)
        if block not in raw:
            raise KeyError(f"reference_forecasts.json: gözlənilən blok yoxdur — {block}")
        body = raw[block]
        out[block] = {var: {int(y): float(v) for y, v in sorted(vals.items())}
                      for var, vals in body.items() if isinstance(vals, dict)}
    _REF_CACHE["data"] = out
    return out


def reference_series(var: str, block: str = REFERENCE_PRIMARY_BLOCK) -> pd.Series:
    """Bir BVF sırası — indeks = il, dəyər = faizlə temp. Blok tapılmazsa boş sıra qaytarılır."""
    if block not in REFERENCE_BLOCKS:
        raise KeyError(f"blok ağ siyahıda deyil: {block}; icazə verilənlər: "
                       f"{sorted(REFERENCE_BLOCKS)}")
    body = reference_forecasts()[block].get(var, {})
    return pd.Series(body, dtype=float, name=f"{block}:{var}").sort_index()


def imf_reference(var: str, years=None) -> pd.Series:
    """Nəşr olunan BVF müqayisə yolu: əvvəlcə Maddə IV (26/112), həmin sıra orada yoxdursa WEO.

    `years` verildikdə nəticə həmin illərə məhdudlaşdırılır (proqnoz üfüqü üçün adi istifadə:
    `imf_reference("gdp_realg", FY)`). Sıra heç bir blokda yoxdursa BOŞ sıra qaytarılır —
    çağıran tərəf bunu «BVF-də uyğun sıra yoxdur» kimi oxumalı və UYDURMAMALIDIR.
    """
    s = reference_series(var, REFERENCE_PRIMARY_BLOCK)
    if s.empty:
        s = reference_series(var, "IMF_WEO")
    if years is not None:
        yrs = [int(y) for y in years]
        s = s.reindex([y for y in yrs if y in s.index])
    return s.rename(var)


def reference_provenance() -> pd.DataFrame:
    """Sətir-sətir provenans: hansı sıra hansı blokdan, hansı illər, hansı buraxılışdan.

    Digər panellərin (`public_provenance`, `customs_hs27_provenance`) provenans cədvəlləri ilə
    eyni formadadır ki, hesabatda və dashboard-da eyni şəkildə oxuna bilsin.
    """
    rows = []
    data = reference_forecasts()
    for block, (short, label) in REFERENCE_BLOCKS.items():
        for var, vals in sorted(data[block].items()):
            yrs = sorted(vals)
            rows.append(dict(
                var=var, block=block, block_short=short, label_az=label,
                published=var in REFERENCE_SERIES,
                year_min=int(yrs[0]) if yrs else None,
                year_max=int(yrs[-1]) if yrs else None,
                n_years=len(yrs), unit="%",
                vintage=REFERENCE_VINTAGE_NOTE, source_note=REFERENCE_SOURCE_NOTE,
                note=("bu paketin eyni adlı sırası ilə tutuşdurulur"
                      if var in REFERENCE_SERIES
                      else REFERENCE_UNMATCHED_NOTE.get(
                          var, "bu paketdə eyni tərifli qarşılığı yoxdur — nəşr olunmur")),
            ))
    return pd.DataFrame(rows).sort_values(["block", "var"], kind="stable").reset_index(drop=True)


def reference_match_report() -> pd.DataFrame:
    """«Hansı BVF göstəricisinin bu paketdə qarşılığı var, hansının yoxdur» cədvəli.

    Uyğunluq gəlmədikdə səbəb AÇIQ yazılır (susmaq yoxdur, uydurma sıra yoxdur) — FR8-in
    7-ci alt-tapşırığının «uyğun sıra olmadıqda bunu bir cümlə ilə bildir» şərti buradan
    qidalanır.
    """
    data = reference_forecasts()
    all_vars = sorted({v for block in data.values() for v in block})
    rows = []
    for var in all_vars:
        blocks = [short for b, (short, _) in REFERENCE_BLOCKS.items() if var in data[b]]
        matched = var in REFERENCE_SERIES
        rows.append(dict(
            imf_var=var, in_blocks=", ".join(blocks),
            package_series_code=(var if matched else None),
            name_az=(REFERENCE_SERIES[var][0] if matched else None),
            matched=matched,
            reason=("" if matched else REFERENCE_UNMATCHED_NOTE.get(
                var, "bu paketdə eyni tərifli sıra modelləşdirilmir")),
        ))
    return pd.DataFrame(rows)


def reference_rows(years, fr_of: dict) -> pd.DataFrame:
    """`forecast_long.csv` müqavilə sütunları ilə hazır BVF müqayisə sətirləri.

    `fr_of`: sıra kodu -> həmin sıranın SAHİBİ olan FR (məsələn `{"gdp_realg": "FR1"}`). Sətirlər
    sahibi olan FR-in partiyası ilə birlikdə yazılmalıdır, çünki `outputs.write_forecast_long`
    əhatəni `fr` sütunu üzrə yerinə-yazır — ayrıca çağırış həmin FR-in qalan sətirlərini silərdi.
    Yelpik zolağı sütunları BOŞ qalır: BVF nəşrində qeyri-müəyyənlik zolağı verilmir və
    uydurulmur.
    """
    yrs = [int(y) for y in years]
    rows = []
    for code, fr in sorted(fr_of.items()):
        if code not in REFERENCE_SERIES:
            raise KeyError(f"`{code}` BVF uyğunluq siyahısında deyil ({sorted(REFERENCE_SERIES)})")
        name_az, unit = REFERENCE_SERIES[code]
        s = imf_reference(code, yrs)
        for y, v in s.items():
            rows.append({"fr": fr, "series_code": code, "series_name_az": name_az, "unit": unit,
                         "year": int(y), "kind": "forecast", "source": "imf_reference",
                         "value": round(float(v), 6),
                         "lo80": None, "hi80": None, "lo50": None, "hi50": None})
    return pd.DataFrame(rows).astype({"lo80": "float64", "hi80": "float64",
                                      "lo50": "float64", "hi50": "float64"})


def verify_reference_forecasts(tol: float = 1e-9) -> pd.DataFrame:
    """Mənbə faylının təkrar oxunuşu və iki blokun üst-üstə düşən sıralarının tutuşdurulması.

    Maddə IV və WEO eyni buraxılışdan gəldiyi üçün ortaq sıralar (ÜDM real artımı, İQİ)
    eyni olmalıdır; fərq varsa bu, buraxılış qarışığının əlamətidir və AÇIQ göstərilir.
    """
    data = reference_forecasts()
    a, w = data[REFERENCE_PRIMARY_BLOCK], data["IMF_WEO"]
    rows = []
    for var in sorted(set(a) & set(w)):
        yrs = sorted(set(a[var]) & set(w[var]))
        diffs = [abs(a[var][y] - w[var][y]) for y in yrs]
        rows.append(dict(var=var, n_years=len(yrs),
                         max_abs_diff=(max(diffs) if diffs else float("nan")),
                         agree=bool(diffs and max(diffs) <= tol)))
    return pd.DataFrame(rows)


if __name__ == "__main__":
    build_cache()
