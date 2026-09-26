# -*- coding: utf-8 -*-
"""Mərkəzi konfiqurasiya — MİİS §15.5.1 makroiqtisadi proqnoz paketi.

Bütün yollar, üfüq sərhədləri və baza sabitləri BURADA bir dəfə təyin olunur.
Heç bir notebook və ya modul öz nüsxəsini saxlamamalıdır: nümunə sərhədi,
proqnoz üfüqü və 2025-ci il bazası yalnız bu fayldan oxunur.
"""
import os

# --------------------------------------------------------------------------
# 1. Zaman sərhədləri
# --------------------------------------------------------------------------
LAST_ACTUAL = 2025                 # faktiki (hesabat) məlumatın son ili
FIRST_FORECAST = 2026              # proqnoz üfüqünün ilk ili
LAST_FORECAST = 2030               # proqnoz üfüqünün son ili
FY = range(FIRST_FORECAST, LAST_FORECAST + 1)   # 2026–2030

# Qiymətləndirmə nümunəsi heç bir halda LAST_ACTUAL-dan kənara çıxmır;
# Nazirliyin Proqnoz sütunları yalnız müqayisə (overlay) məqsədi ilə oxunur.

# --------------------------------------------------------------------------
# 2. Baza göstəriciləri
# --------------------------------------------------------------------------
SSC_GDP_2025 = 129094.0            # DSK-nın rəsmi 2025 ÜDM göstəricisi, mln AZN
VEREQ_GDP_2025_VINTAGE = 128134.2131721304   # 8 vərəq şablonunun köhnə buraxılışı (yalnız qeyd üçün)
USD_AZN_PEG = 1.70                 # manatın ABŞ dollarına faktiki bağlılığı

# --------------------------------------------------------------------------
# 3. Yollar
# --------------------------------------------------------------------------
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, "src")
DATA = os.path.join(BASE, "data")
CACHE = os.path.join(DATA, "cache")
OUT = os.path.join(BASE, "outputs")
FIG = os.path.join(OUT, "figures")
NOTEBOOKS = os.path.join(BASE, "notebooks")
REPORT = os.path.join(BASE, "report")

for _d in (DATA, CACHE, OUT, FIG):
    os.makedirs(_d, exist_ok=True)

# --------------------------------------------------------------------------
# 4. Giriş faylları
# --------------------------------------------------------------------------
XLSX_MAIN = os.path.join(DATA, "Statistik data dinamika 05.06.2026 +.xlsx")
XLSX_MAIN_MD5 = "028fcc0f36736482dcaceacfbb9dc827"   # mənbənin bütövlüyü yoxlanılır
XLSX_VEREQ = os.path.join(DATA, "8_vereq_original.xlsx")

P_MACRO = os.path.join(DATA, "macro_annual.csv")            # illik makro panel (DSK nəşrləri)
P_MODEL = os.path.join(DATA, "model_series_annual.csv")     # model üçün hazırlanmış illik sıralar
P_EXTERNAL = os.path.join(DATA, "external_block_annual.csv")  # xarici blok (IMF/ECB mənbələri)
P_MONTHLY = os.path.join(DATA, "monthly_panel.csv")         # aylıq panel (İQİ, brent, faiz)
P_OFFICIAL_2025 = os.path.join(DATA, "official_2025_actuals.csv")  # DSK-nın rəsmi 2025 nəticələri

ASSUMPTIONS = os.path.join(DATA, "assumptions.csv")         # YEGANƏ fərziyyə dəsti (D1)
ASSUMPTION_REGISTRY = os.path.join(DATA, "assumption_registry.csv")  # əvəzləmə reyestri (B21)

# Nazirlik spesifikasiyası ssenarisi (FR13). Kataloq Nazirliyin öz təhvil paketindən
# DƏYİŞDİRİLMƏDƏN köçürülüb; panel və provenans onun nümunə-model iş kitablarından yalnız
# faktiki (≤2024) sütunlar üzrə çıxarılıb (bax `data_layer` §13).
P_MINISTRY_CATALOG = os.path.join(DATA, "ministry_equations_catalog.csv")
P_MOE_SPEC_PANEL = os.path.join(DATA, "moe_spec_panel.csv")
P_MOE_SPEC_PROV = os.path.join(DATA, "moe_spec_provenance.csv")

# Rəsmi açıq mənbələrin (DSK / AMB / opendata.az / Nazirliyin BOP iş kitabı) dondurulmuş
# paneli — bax `data_layer` §14. Panel daxili builder skripti ilə qurulur;
# paket bundan sonra mənbə fayllarına ASILI DEYİL (özünü-tam qalır).
P_PUBLIC_PANEL = os.path.join(DATA, "public_sources_panel.csv")
P_PUBLIC_PROV = os.path.join(DATA, "public_sources_provenance.csv")

# DGK-nın (Dövlət Gömrük Komitəsi) gömrük statistika bülletenlərindən çıxarılmış
# HS-27 (neft, neft məhsulları, qaz) illik miqdar/dəyər paneli — bax `data_layer` §15.
P_CUSTOMS_PANEL = os.path.join(DATA, "customs_hs27_annual.csv")
P_CUSTOMS_PROV = os.path.join(DATA, "customs_hs27_provenance.csv")

# BVF-nin (Beynəlxalq Valyuta Fondu) nəşr olunmuş etalon yolları — Maddə IV məsləhətləşməsinin
# ölkə hesabatı 26/112 və WEO buraxılışı. Fayl YALNIZ OXUNUR: heç bir tənliyə, heç bir
# fərziyyə açarına daxil olmur, yalnız müqayisə sətirləri kimi nəşr edilir (bax `data_layer` §16).
P_REFERENCE_FORECASTS = os.path.join(DATA, "reference_forecasts.json")

# --------------------------------------------------------------------------
# 5. Nazirliyin statutar şablonu (8 vərəq) — vərəq adları
# --------------------------------------------------------------------------
VEREQ_SHEETS = {
    "2.4.1.1.": "v1",        # ÜDM (real artım, deflyator, neft/qeyri-neft)
    "2.4.1.2.": "v2",        # ÜDM istehsal strukturu (buraxılış, aralıq istehlak, əlavə dəyər)
    "2.4.1.3.": "v3",        # ÜDM sahə strukturu
    "2.4.1.4.": "v4",        # əsas məhsul növləri (natural göstəricilər)
    "2.4.1.5.-7.": "v57",    # neftin qiyməti, investisiyalar, əhalinin gəlirləri, inflyasiya
    "2.4.1.13.-14.": "v1314",  # xarici ticarət və tədiyə balansı
}

# Əsas iş kitabından götürülən vərəqlər (aylıq/rüblük daşıyan bloklar).
# Son dörd vərəq v3-ün P0 mərhələsində əlavə edilib:
#   · `Mədənçıxarma`  — neft-qaz hasilatı üzrə buraxılış/əlavə dəyər, sahə əmək haqqı və işçi sayı;
#   · `Su təchizatı`  — sahə üzrə ÜDM, investisiya, əmək haqqı, işçi sayı;
#   · `DVX üzrə göstəricilər` — vergi daxilolmaları, sahə dövriyyəsi, əmək haqqı fondu;
#   · `DİP 2016-2026` — dövlət əsaslı vəsait qoyuluşu (iki pilləli başlıq: 2026 üçün
#     "Nəzərdə tutulmuş vəsait" və "Faktiki xərc (01.04.2026)" sütunları).
MAIN_SHEETS = ["Real sektor", "Sosial sektor", "Monetar sektoru", "Neft-Qaz sektoru",
               "Ticarət", "Fiskal sektor", "Emal Sənayesi", "Elektrik enerjisi",
               "Tədiyyə Balansı", "Mədənçıxarma", "Su təchizatı",
               "DVX üzrə göstəricilər", "DİP 2016-2026"]

# --------------------------------------------------------------------------
# 6. Keş faylları
# --------------------------------------------------------------------------
C_WB_REGISTRY = os.path.join(CACHE, "wb_registry.csv")
C_WB_ANNUAL = os.path.join(CACHE, "wb_annual.csv")
C_WB_MONTHLY = os.path.join(CACHE, "wb_monthly.csv")
C_VEREQ_LONG = os.path.join(CACHE, "vereq_long.csv")
C_VEREQ_PRODUCTS = os.path.join(CACHE, "vereq_products.csv")
C_VEREQ_REF = os.path.join(CACHE, "vereq_ref_cells.csv")
C_VEREQ_META = os.path.join(CACHE, "vereq_meta.csv")
C_SPLICE = os.path.join(CACHE, "splice_check.csv")
C_ACTUALS = os.path.join(CACHE, "actuals_long.csv")
C_BASELINE = os.path.join(CACHE, "baseline_long.csv")
C_CATALOG = os.path.join(CACHE, "catalog.csv")
C_VINTAGE = os.path.join(CACHE, "vintage_overrides.csv")
C_DICT = os.path.join(CACHE, "series_dictionary_starter.csv")
C_GAPS = os.path.join(CACHE, "data_gaps.csv")

# §3 müqaviləsi üzrə çıxış faylı (məlumat qatı başlanğıc versiyanı yazır)
O_SERIES_DICT = os.path.join(OUT, "series_dictionary.csv")

# --------------------------------------------------------------------------
# 7. Ədədi dəqiqlik
# --------------------------------------------------------------------------
TOL_IDENTITY = 1e-6        # balans eyniliklərinin qapanma həddi
TOL_SOLVER = 1e-8          # sabit nöqtə həllinin yığılma həddi
MAX_ITER_SOLVER = 200
