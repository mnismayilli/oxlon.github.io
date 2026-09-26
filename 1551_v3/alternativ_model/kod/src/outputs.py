"""outputs.py — Nazirlik dashboard-u üçün maşın-oxunan çıxış yazıcıları (§3 müqavilə sütunları).

Hər yazıcı: (1) sütunları §3 müqaviləsinə uyğun DƏQİQ yoxlayır, (2) təbii açar (natural key) üzrə
dublikatları təmizləyir, (3) CSV-yə İDEMPOTENT şəkildə əlavə edir — eyni FR təkrar işlədildikdə
özünün əvvəlki sətirlərini əvəz edir, heç vaxt dublikat yaratmır. "Əhatə sütunları" (scope_cols) həmin
partiyanın hansı FR/fərziyyə açarına aid olduğunu göstərir: yeni partiyada rast gəlinən əhatə
dəyərlərinə uyğun bütün köhnə sətirlər silinir, sonra yeni sətirlər əlavə olunur — beləliklə bir FR-in
modeli dəyişəndə (məsələn bir tənlik artıq saxlanılmırsa) köhnə iz qalmır.
"""
from __future__ import annotations

import os
import re
import warnings
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd

# ---------------------------------------------------------------------------
# Yol həlli: config.py mövcud olarsa oradan İSTİFADƏ OLUNUR (vahid mənbə); əks halda
# bu fayl öz yerindən (src/..) törəyərək DATA/OUT yollarını müstəqil müəyyən edir. Bu, config.py
# hələ tikilməzdən əvvəl də bu modulun sərbəst test edilə bilməsini təmin edir.
# ---------------------------------------------------------------------------
try:
    from . import config as _cfg  # type: ignore
    DATA_DIR: Path = Path(_cfg.DATA)
    OUT_DIR: Path = Path(_cfg.OUT)
except Exception:
    _BASE_DIR = Path(__file__).resolve().parent.parent
    DATA_DIR = _BASE_DIR / "data"
    OUT_DIR = _BASE_DIR / "outputs"

DATA_DIR.mkdir(parents=True, exist_ok=True)
OUT_DIR.mkdir(parents=True, exist_ok=True)

ASSUMPTIONS_PATH = DATA_DIR / "assumptions.csv"
ASSUMPTION_REGISTRY_PATH = DATA_DIR / "assumption_registry.csv"
EQUATIONS_CATALOG_PATH = OUT_DIR / "equations_catalog.csv"
FORECAST_LONG_PATH = OUT_DIR / "forecast_long.csv"
VALIDATION_BACKTEST_PATH = OUT_DIR / "validation_backtest.csv"
SERIES_DICTIONARY_PATH = OUT_DIR / "series_dictionary.csv"


# ---------------------------------------------------------------------------
# §3 müqaviləsi — sütun adları DƏQİQ olaraq BUILD_CONTEXT.md-dən köçürülüb, sıra saxlanılır.
# ---------------------------------------------------------------------------
EQUATIONS_CATALOG_COLUMNS: List[str] = [
    "workbook", "sheet", "eq_name", "description", "lhs", "rhs",
    "series_code", "fr", "sample_start", "sample_end", "adj_r2", "se_regression",
]
FORECAST_LONG_COLUMNS: List[str] = [
    "fr", "series_code", "series_name_az", "unit", "year", "kind", "source",
    "value", "lo80", "hi80", "lo50", "hi50",
]
ASSUMPTIONS_COLUMNS: List[str] = ["assumption_key", "year", "value", "unit", "note"]
# Əvəzləmə (override) reyestri — fərziyyə dəstinin İDARƏETMƏ qatı: hər açarın hansı
# rejimdə əvəz oluna biləcəyi və dəyərin rəsmi sahibi. Bax məzmun paketi, sənəd 01.
ASSUMPTION_REGISTRY_COLUMNS: List[str] = ["assumption_key", "override_class", "owner_org", "basis"]
OVERRIDE_CLASSES: Dict[str, str] = {
    "A": "rəsmi/ekzogen giriş — əvəzləmə nəzərdə tutulub, baza yolunu dəyişir",
    "B": "modul nəticəsi — əvəzləmə yalnız adlandırılmış ssenari kimi",
    "C": "törəmə, eynilik, zolaq və ya audit möhürü — əvəzləmə qapalıdır",
}
VALIDATION_BACKTEST_COLUMNS: List[str] = [
    "series_code", "model", "vintage_year", "horizon_h", "actual", "forecast",
    "error", "abs_error", "rmse_h", "rw_rmse_h", "coverage80",
]
SERIES_DICTIONARY_COLUMNS: List[str] = [
    "series_code", "name_az", "name_en", "unit", "fr", "statutory_sheet_ref", "price_basis",
]

# ---------------------------------------------------------------------------
# Mənbə (source) lüğəti — `forecast_long.csv`-in `source` sütununun İCAZƏLİ dəyərləri.
# Hər dəyər üçün dashboard-da göstərilən AZ etiketi burada, bir yerdə saxlanılır ki,
# ekranlarda və hesabatda eyni ad işlədilsin.
#   · `ours`              — bu paketin modeli;
#   · `template_sample`   — Nazirliyin 8 vərəq şablonundakı nümunə (şablon) rəqəmləri,
#                           YALNIZ müqayisə üçün; modelin heç bir tənliyi onları görmür;
#   · `ministry_official` — Nazirliyin RƏSMİ proqnozu üçün ayrılmış BOŞ slot: təhvil verilən
#                           paketdə bu mənbə ilə bir sətir də yoxdur, rəsmi rəqəmlər daxil
#                           olduqda dolduralacaq;
#   · `ministry_spec`     — «Nazirlik spesifikasiyası» SSENARİSİ (FR13): Nazirliyin öz 92
#                           sətirlik tənlik kataloqunun ÖZ əmsalları ilə, bizim məlumat qatı və
#                           bizim ekzogen yollarımız üzərində həll edilmiş yolu. Bu, ayrıca
#                           ssenaridir — baza proqnoz DEYİL və `ours` ilə eyni ekranda müqayisə
#                           sütunu kimi göstərilir;
#   · `imf_reference`     — BVF-nin (Beynəlxalq Valyuta Fondu) NƏŞR OLUNMUŞ etalon yolu:
#                           Maddə IV məsləhətləşməsinin ölkə hesabatı 26/112 (Cədvəl 1) və eyni
#                           buraxılışın WEO yolu. Yalnız hər iki tərəfdə EYNİ tərifli sıralar
#                           üçün (ÜDM real artımı, qeyri-neft real artımı, İQİ inflyasiyası);
#                           buraxılış qeydi `data_layer.REFERENCE_VINTAGE_NOTE`-dadır. Yelpik
#                           zolağı YOXDUR — BVF nəşrində verilmir və uydurulmur.
#                           bu buraxılışda yenidən adlandırıldı: əvvəlki BOŞ `imf_weo` / `imf_artiv` slotları bir
#                           populyasiya olunmuş `imf_reference` dəyəri ilə əvəzlənib (eyni
#                           qayda ilə ki, P0-da `ministry_decree75` → `template_sample`).
# ---------------------------------------------------------------------------
SOURCE_LABELS_AZ: Dict[str, str] = {
    "ours": "bu paketin modeli",
    "template_sample": "şablon-nümunə ssenarisi (75 saylı qərar formatı üzrə)",
    "ministry_official": "Nazirliyin rəsmi proqnozu",
    "ministry_spec": "Nazirlik spesifikasiyası ssenarisi (Nazirliyin 92 tənlikli kataloqu)",
    "imf_reference": "BVF etalon proqnozu (Maddə IV 26/112 · WEO)",
}
FORECAST_SOURCES: List[str] = list(SOURCE_LABELS_AZ)
# Təhvil anında BOŞ olması gözlənilən slotlar (məlumat gəldikdə doldurulur).
EMPTY_SOURCE_SLOTS: List[str] = ["ministry_official"]

# ---------------------------------------------------------------------------
# Qiymət bazası (D5) — hər nəşr olunan sıra öz qiymət bazasını ELAN edir:
#   · `cari`       — cari (nominal) qiymətlər;
#   · `sabit-2025` — 2025-ci ilin sabit qiymətləri (real səviyyə);
#   · `indeks`     — faiz/indeks göstəricisi (artım tempi, deflyator, pay);
#   · `natural`    — natural göstərici (miqdar, sayı) — qiymət bazası anlayışı tətbiq olunmur.
# Sütun `series_dictionary.csv`-dədir; verilmədikdə ölçü vahidindən çıxarılır (aşağıdakı qayda),
# lakin qiyməti sabit ilin qiymətləri ilə verilən sıralar bunu AÇIQ göstərməlidir.
# ---------------------------------------------------------------------------
PRICE_BASIS_VALUES: List[str] = ["cari", "sabit-2025", "indeks", "natural"]
_MONETARY_UNIT_TOKENS = ("azn", "manat", "usd", "dollar", "eur", "avro")


def price_basis_for_unit(unit: object) -> str:
    """Ölçü vahidindən qiymət bazasının SUSMAYA görə dəyəri (D5).

    Faiz və indekslər `indeks`; pul ifadəsindəki sıralar `cari`; qalan natural göstəricilər
    (min nəfər, mln ton, mlrd m³ və s.) `natural` sayılır. Sabit qiymətli sıralar bu qaydadan
    çıxa bilmədiyi üçün onları yazan dəftər `price_basis` sütununu AÇIQ verməlidir.
    """
    u = str(unit or "").strip().lower()
    if not u:
        return "natural"
    if "%" in u or "indeks" in u or "faiz" in u:
        return "indeks"
    if any(tok in u for tok in _MONETARY_UNIT_TOKENS):
        return "cari"
    return "natural"

# Hər cədvəl üçün: təbii açar (KEY) — sətri unikal identifikasiya edir; əhatə (SCOPE) — hansı
# FR/fərziyyə partiyasının köhnə sətirlərinin silinəcəyini müəyyən edir (SCOPE həmişə KEY-in alt
# çoxluğudur ya da ona bərabərdir).
_KEY_COLS = {
    "equations_catalog": ["fr", "eq_name"],
    "forecast_long": ["fr", "series_code", "year", "kind", "source"],
    "assumptions": ["assumption_key", "year"],
    "validation_backtest": ["series_code", "model", "vintage_year", "horizon_h"],
    "series_dictionary": ["series_code"],
}
_SCOPE_COLS = {
    "equations_catalog": ["fr"],
    "forecast_long": ["fr"],
    "assumptions": ["assumption_key"],
    "validation_backtest": ["series_code", "model"],
    "series_dictionary": ["series_code"],
}
_NULLABLE = {
    "equations_catalog": {"adj_r2", "se_regression"},          # kalibrlənmiş tənliklərdə ola bilər
    "forecast_long": {"lo80", "hi80", "lo50", "hi50"},          # yelpik zolağı hər sətirdə tələb olunmur
    "assumptions": {"note"},
    "validation_backtest": {"coverage80"},                       # yalnız bant mövcud olduqda hesablanır
    "series_dictionary": {"statutory_sheet_ref"},
}


class OutputsContractError(ValueError):
    """§3 dashboard müqaviləsi pozulduqda atılır (sütun uyğunsuzluğu, boş açar, vs.)."""


class AssumptionRegistryError(OutputsContractError):
    """`data/assumption_registry.csv` fərziyyə dəsti ilə uyğunsuz olduqda atılır: təsnif
    olunmamış açar, reyestrdə qalmış yetim açar, naməlum sinif və ya A sinfində boş sahib
    qurum. Bu yoxlama sənəd 01-in (əvəzləmə müqaviləsi) köhnəlməsinin qarşısını alır —
    icmalın B21 bəndi məhz həmin köhnəlmə üzündən yazılmışdı."""


class BacktestOverwriteError(OutputsContractError):
    """Bir sınaq partiyası BAŞQA dəftərin eyni `(series_code, model)` sətirlərini əvəz etməyə
    çalışdıqda atılır (bax `_guard_backtest_overwrite`)."""


def _validate_columns(df: pd.DataFrame, contract: List[str], key_cols: List[str], name: str) -> pd.DataFrame:
    got, want = set(df.columns), set(contract)
    missing, extra = want - got, got - want
    if missing or extra:
        raise OutputsContractError(
            f"{name}: sütunlar §3 müqaviləsinə uyğun deyil — çatışmayan={sorted(missing)}, "
            f"artıq={sorted(extra)}"
        )
    nullable = _NULLABLE.get(name, set())
    required_non_null = [c for c in contract if c not in nullable]
    bad = df[required_non_null].isna().any()
    bad_cols = bad[bad].index.tolist()
    if bad_cols:
        raise OutputsContractError(f"{name}: bu sütunlarda boş (NaN) dəyərə icazə verilmir: {bad_cols}")
    missing_key = [c for c in key_cols if df[c].isna().any()]
    if missing_key:
        raise OutputsContractError(f"{name}: təbii açar sütununda boş dəyər ola bilməz: {missing_key}")
    return df[contract].copy()  # sıralamanı §3-ə uyğunlaşdır


def _upsert_csv(path: Path, new_rows: pd.DataFrame, key_cols: List[str], scope_cols: List[str]) -> pd.DataFrame:
    """Əhatə-əsaslı yerinə-yazma: `scope_cols` üzrə yeni partiyada rast gəlinən dəyərlərə uyğun
    bütün köhnə sətirlər silinir, sonra yeni sətirlər əlavə olunur, açar üzrə son dublikat-təmizləmə
    aparılır (təhlükəsizlik toru) və nəticə açar sırası ilə diskə yazılır.

    `float_precision="round_trip"`: pandas-ın SUSMAYA görə sürətli float oxuyucusu bəzi dəyərləri
    son bitdə (1 ULP) dəyişdirir; hər dəftər faylı bütövlükdə oxuyub yenidən yazdığı üçün bu, təkrar
    icralarda rəqəmlərin səssizcə sürüşməsinə (byte-idempotentliyin pozulmasına) səbəb olurdu.
    Dəqiq (round-trip) oxuma bunu bağlayır: təkrar icra eyni faylı bayt-bayt verir."""
    if path.exists() and os.path.getsize(path) > 0:
        existing = pd.read_csv(path, float_precision="round_trip")
    else:
        existing = pd.DataFrame(columns=new_rows.columns)

    if not existing.empty:
        scope_keys_new = set(map(tuple, new_rows[scope_cols].itertuples(index=False, name=None)))
        existing_scope = list(map(tuple, existing[scope_cols].itertuples(index=False, name=None)))
        keep_mask = [t not in scope_keys_new for t in existing_scope]
        existing = existing.loc[keep_mask]

    # `existing` boşdursa birbaşa `new_rows`-u istifadə et — boş DataFrame ilə concat pandas-da
    # tip xəbərdarlığı yaradır və lazımsızdır. Qalan hallarda da (məsələn `lo80`/`hi80` bütün
    # sətirlərdə NA olan bir sıra) pandas eyni FutureWarning-i verə bilər; xəbərdarlıq mətni öz
    # daxilində çağıran faylın MÜTLƏQ yolunu daşıyır (bu, təhvil sızma yoxlamasını pozur —
    # DELIVERY_SPEC §4.2), ona görə burada NÖQTƏLİ şəkildə (yalnız bu concat üçün) susdurulur.
    if existing.empty:
        combined = new_rows.copy()
    else:
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", message=".*empty or all-NA entries.*",
                                    category=FutureWarning)
            combined = pd.concat([existing, new_rows], ignore_index=True)
    # Təhlükəsizlik toru: eyni açar üzrə iki sətir qalarsa, sonuncunu (yeni məlumatı) saxla.
    combined = combined.drop_duplicates(subset=key_cols, keep="last")
    combined = combined.sort_values(by=key_cols, kind="stable").reset_index(drop=True)
    combined.to_csv(path, index=False)
    return combined


def write_equations_catalog(df: pd.DataFrame) -> pd.DataFrame:
    """Bir qiymətləndirilmiş tənlik = bir sətir. `fr` əhatəsi üzrə yerinə-yazılır."""
    clean = _validate_columns(df, EQUATIONS_CATALOG_COLUMNS, _KEY_COLS["equations_catalog"], "equations_catalog")
    return _upsert_csv(EQUATIONS_CATALOG_PATH, clean, _KEY_COLS["equations_catalog"], _SCOPE_COLS["equations_catalog"])


def write_forecast_long(df: pd.DataFrame) -> pd.DataFrame:
    """Uzun formatda proqnoz/aktual sıra. `fr` əhatəsi üzrə yerinə-yazılır.

    `source` sütunu yalnız `SOURCE_LABELS_AZ` lüğətindəki dəyərləri qəbul edir — beləliklə
    dashboard-un mənbə filtri heç vaxt tanınmayan etiketlə qarşılaşmır.
    """
    clean = _validate_columns(df, FORECAST_LONG_COLUMNS, _KEY_COLS["forecast_long"], "forecast_long")
    bad_src = sorted(set(clean["source"].astype(str)) - set(SOURCE_LABELS_AZ))
    if bad_src:
        raise OutputsContractError(
            f"forecast_long: tanınmayan `source` dəyəri: {bad_src}; "
            f"icazə verilənlər: {FORECAST_SOURCES}")
    return _upsert_csv(FORECAST_LONG_PATH, clean, _KEY_COLS["forecast_long"], _SCOPE_COLS["forecast_long"])


def write_assumptions(df: pd.DataFrame) -> pd.DataFrame:
    """VAHİD fərziyyə dəsti (D1). `assumption_key` əhatəsi üzrə yerinə-yazılır — məsələn FR5 öz
    `brent_usd` yolunu təkrar hesabladıqda bu açarın BÜTÜN illəri əvəzlənir, digər açarlara toxunulmur."""
    clean = _validate_columns(df, ASSUMPTIONS_COLUMNS, _KEY_COLS["assumptions"], "assumptions")
    return _upsert_csv(ASSUMPTIONS_PATH, clean, _KEY_COLS["assumptions"], _SCOPE_COLS["assumptions"])


def _guard_backtest_overwrite(new_rows: pd.DataFrame, tol: float = 1e-9) -> None:
    """Sınaq sətirlərinin SƏSSİZ itməsinin qarşısını alır.

    NİYƏ: `validation_backtest.csv` `(series_code, model)` əhatəsi üzrə yerinə-yazılır — yəni eyni
    sıra kodu və eyni model adı ilə yazan İKİNCİ dəftər birincinin bütün sətirlərini heç bir xəbərdarlıq
    olmadan silir. İki dəftər eyni göstərici üçün fərqli pəncərədə (fərqli nümunə, fərqli etalon xətası)
    sınaq apardıqda bu, nəticələrin birinin itməsi deməkdir və Nazirliyin dashboard-unda yalnız sonuncu
    yazanın rəqəmi qalır. Ona görə yeni partiya diskdəki eyni `(series_code, model)` cütünü FƏRQLİ
    etalon xətası (`rw_rmse_h`) ilə əvəz etməyə çalışırsa, bu, iki müxtəlif sınaq deməkdir və yazma
    DAYANDIRILIR: sıra kodu ad fəzasında ayrılmalıdır (məsələn `inv_total_g` → `inv_total_g_fr1_core`).
    Eyni sınağın təkrar icrası isə eyni `rw_rmse_h` verdiyi üçün maneəsiz keçir (idempotentlik pozulmur).
    """
    if not VALIDATION_BACKTEST_PATH.exists() or os.path.getsize(VALIDATION_BACKTEST_PATH) == 0:
        return
    existing = pd.read_csv(VALIDATION_BACKTEST_PATH, float_precision="round_trip")
    if existing.empty:
        return

    def _first_rw(df: pd.DataFrame) -> Dict[tuple, float]:
        out: Dict[tuple, float] = {}
        for k, g in df.groupby(["series_code", "model"], dropna=False):
            vals = pd.to_numeric(g["rw_rmse_h"], errors="coerce").dropna()
            if len(vals):
                out[tuple(k)] = float(vals.iloc[0])
        return out

    old_rw, new_rw = _first_rw(existing), _first_rw(new_rows)
    clashes = []
    for key, val in new_rw.items():
        if key in old_rw:
            prev = old_rw[key]
            if abs(prev - val) > tol * max(1.0, abs(prev)):
                clashes.append((key, prev, val))
    if clashes:
        detail = "; ".join(
            f"{sc}/{md}: diskdə rw_rmse_h={prev:.6g}, yeni partiyada {val:.6g}"
            for (sc, md), prev, val in clashes[:6])
        raise BacktestOverwriteError(
            "validation_backtest: yeni partiya BAŞQA bir sınağın sətirlərini əvəz edərdi — "
            f"{detail}. Eyni `(series_code, model)` cütü iki fərqli sınaq pəncərəsində istifadə "
            "olunur; sıra kodunu ad fəzasında ayırın (məsələn `<kod>_fr1_core`), sonra yenidən yazın."
        )


def write_backtest(df: pd.DataFrame) -> pd.DataFrame:
    """Geriyə doğru sınaq nəticələri. `(series_code, model)` əhatəsi üzrə yerinə-yazılır."""
    clean = _validate_columns(df, VALIDATION_BACKTEST_COLUMNS, _KEY_COLS["validation_backtest"], "validation_backtest")
    _guard_backtest_overwrite(clean)   # başqa dəftərin sətirlərinin səssiz silinməsinə qarşı qoruma
    return _upsert_csv(VALIDATION_BACKTEST_PATH, clean, _KEY_COLS["validation_backtest"], _SCOPE_COLS["validation_backtest"])


def write_series_dictionary(df: pd.DataFrame) -> pd.DataFrame:
    """Sıra lüğəti — bir `series_code` = bir sətir, `series_code` həm açar həm əhatədir.

    `price_basis` (D5) hər sətirdə olmalıdır. Sütun verilməyibsə ölçü vahidindən susmaya görə
    doldurulur (`price_basis_for_unit`); sabit qiymətli sıralar isə onu AÇIQ verməlidir, çünki
    vahid ("mln AZN") cari və sabit qiymətləri bir-birindən ayırmır — v2-də məhz bu fərq
    görünmürdü.
    """
    out = df.copy()
    if "price_basis" not in out.columns:
        out["price_basis"] = out["unit"].map(price_basis_for_unit)
    else:
        miss = out["price_basis"].isna() | (out["price_basis"].astype(str).str.strip() == "")
        if miss.any():
            out.loc[miss, "price_basis"] = out.loc[miss, "unit"].map(price_basis_for_unit)
    bad_pb = sorted(set(out["price_basis"].astype(str)) - set(PRICE_BASIS_VALUES))
    if bad_pb:
        raise OutputsContractError(
            f"series_dictionary: tanınmayan `price_basis` dəyəri: {bad_pb}; "
            f"icazə verilənlər: {PRICE_BASIS_VALUES}")
    clean = _validate_columns(out, SERIES_DICTIONARY_COLUMNS, _KEY_COLS["series_dictionary"], "series_dictionary")
    return _upsert_csv(SERIES_DICTIONARY_PATH, clean, _KEY_COLS["series_dictionary"], _SCOPE_COLS["series_dictionary"])


def read_assumptions() -> Dict[str, pd.Series]:
    """`data/assumptions.csv`-i oxuyur və `assumption_key -> pd.Series(index=year, value=dəyər)`
    lüğəti qaytarır (hər açar üçün bir "yol" / path). Fayl yoxdursa boş lüğət qaytarılır — FR5/FR9/FR10
    hələ işləməmiş ilk mərhələdə DAG-ın uğursuz olmaması üçün."""
    if not ASSUMPTIONS_PATH.exists() or os.path.getsize(ASSUMPTIONS_PATH) == 0:
        return {}
    raw = pd.read_csv(ASSUMPTIONS_PATH, float_precision="round_trip")
    out: Dict[str, pd.Series] = {}
    for key, g in raw.groupby("assumption_key"):
        out[key] = g.sort_values("year").set_index("year")["value"]
    return out


# ---------------------------------------------------------------------------
# Əvəzləmə reyestri — hər fərziyyə açarı təsnif olunmalıdır (B21)
# ---------------------------------------------------------------------------
def read_assumption_registry() -> pd.DataFrame:
    """Əvəzləmə reyestrini oxuyur. Fayl yoxdursa BOŞ cədvəl deyil, XƏTA qaytarılır:
    reyestrsiz fərziyyə dəstinin idarəetmə qatı yoxdur."""
    if not ASSUMPTION_REGISTRY_PATH.exists() or os.path.getsize(ASSUMPTION_REGISTRY_PATH) == 0:
        raise AssumptionRegistryError(
            f"əvəzləmə reyestri tapılmadı: {ASSUMPTION_REGISTRY_PATH}")
    reg = pd.read_csv(ASSUMPTION_REGISTRY_PATH)
    return _validate_columns(reg, ASSUMPTION_REGISTRY_COLUMNS, ["assumption_key"],
                             "assumption_registry")


def validate_assumption_registry() -> pd.DataFrame:
    """Reyestr ilə `assumptions.csv` arasında BİRƏBİR uyğunluğu yoxlayır və reyestri qaytarır.

    Dörd şərt: (1) hər fərziyyə açarı təsnif olunub; (2) reyestrdə artıq (yetim) açar yoxdur;
    (3) sinif yalnız A/B/C-dir; (4) A sinfindəki hər açarın sahib qurumu göstərilib.
    Yeni açar əlavə edən istənilən mərhələ reyestri də yeniləməyə məcburdur — sənəd 01-in
    açar siyahısı bu fayldan qurulur."""
    reg = read_assumption_registry()
    keys_reg = set(reg["assumption_key"].astype(str))
    keys_asm = set(read_assumptions().keys())

    missing = sorted(keys_asm - keys_reg)
    orphan = sorted(keys_reg - keys_asm)
    bad_class = sorted(set(reg.loc[~reg["override_class"].isin(OVERRIDE_CLASSES),
                                   "assumption_key"].astype(str)))
    a_rows = reg[reg["override_class"] == "A"]
    no_owner = sorted(a_rows.loc[a_rows["owner_org"].astype(str).str.strip().isin(["", "nan"]),
                                 "assumption_key"].astype(str))

    problems = []
    if missing:
        problems.append(f"reyestrdə təsnif olunmamış açar(lar): {missing}")
    if orphan:
        problems.append(f"reyestrdə qalmış yetim açar(lar): {orphan}")
    if bad_class:
        problems.append(f"naməlum sinif (yalnız A/B/C olmalıdır): {bad_class}")
    if no_owner:
        problems.append(f"A sinfində sahib qurum boşdur: {no_owner}")
    if problems:
        raise AssumptionRegistryError(
            "əvəzləmə reyestri fərziyyə dəsti ilə uyğun deyil — " + "; ".join(problems))
    return reg


# ---------------------------------------------------------------------------
# D2 — dinamiklik göstəricisinin ƏSASLANDIRMASININ nəşri
# ---------------------------------------------------------------------------
_DYN_NOTE_RE = re.compile(r"\s*DİNAMİKLİK\s+[0-9.,]+\s*:.*$", re.S)


def _dyn_note_of(r: object) -> str:
    """Nəşr olunacaq qeyd mətni: mümkün olduqda `DynamicsResult.note_text` (sinif izahı +
    ölçülmüş diaqnostika), əks halda yalnız `justification` (geriyə uyğunluq)."""
    txt = getattr(r, "note_text", None)
    return str(txt if txt else getattr(r, "justification", "") or "")


def publish_dynamics_notes(results: Iterable[object], verbose: bool = True) -> pd.DataFrame:
    """`ek.sigma_ratio` nəticələrinin YAZILI əsaslandırmasını seriya lüğətinin qeyd sütununa yazır.

    D2-nin qaydası: qəbul ARALIĞI nəşr olunmur, lakin göstəricinin dəyəri və onun səbəbi
    nəşr olunur. Əvvəllər əsaslandırma yalnız dəftərin içində qalırdı (`assert_dynamics_band`
    onu ekranda göstərirdi) — bu funksiya onu çıxış müqaviləsinə çıxarır ki, dashboard-u oxuyan
    rəqəmin yanında izahı görsün.

    İDEMPOTENT: mövcud qeydin sonundakı əvvəlki "DİNAMİKLİK …" hissəsi silinib yenisi ilə
    əvəzlənir, ona görə dəftərin təkrar icrası qeydi uzatmır. Lüğətdə hələ olmayan sıra üçün
    sətir YARADILMIR (o sıranı öz FR-i yazır) — belə sıralar `atlandı` kimi çap olunur.
    """
    res = [r for r in results if str(_dyn_note_of(r)).strip()]
    if not res:
        return pd.DataFrame(columns=SERIES_DICTIONARY_COLUMNS)
    cur = pd.read_csv(SERIES_DICTIONARY_PATH) if SERIES_DICTIONARY_PATH.exists() else pd.DataFrame()
    rows, skipped = [], []
    for r in res:
        base = cur[cur["series_code"] == r.series_code] if len(cur) else cur
        if base is None or len(base) == 0:
            skipped.append(r.series_code)
            continue
        row = base.iloc[0].to_dict()
        ref = _DYN_NOTE_RE.sub("", str(row.get("statutory_sheet_ref") or "")).strip()
        row["statutory_sheet_ref"] = (ref + f" DİNAMİKLİK {r.ratio:.3f}: {_dyn_note_of(r)}").strip()
        rows.append({k: row.get(k) for k in SERIES_DICTIONARY_COLUMNS})
    if verbose:
        msg = f"D2: {len(rows)} sıra üçün dinamiklik əsaslandırması seriya lüğətinə yazıldı"
        n_rev = sum(1 for r in res if getattr(r, "under_review", False))
        if n_rev:
            msg += f"; bunlardan {n_rev} sıra YENİDƏN BAXILIR sinfindədir"
        if skipped:
            msg += f"; lüğətdə olmayan {len(skipped)} sıra atlandı ({', '.join(skipped[:6])})"
        print(msg)
    return write_series_dictionary(pd.DataFrame(rows)) if rows else pd.DataFrame()
