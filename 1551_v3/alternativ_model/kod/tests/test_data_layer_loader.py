"""tests/test_data_layer_loader.py — iş kitabı oxuyucusunun (data_layer.load_sheet) reqressiya testləri.

Diqqət mərkəzində İKİ PİLLƏLİ BAŞLIQ: `DİP 2016-2026` vərəqində illər başlıq sətrindədir, 2026-cı il
isə aşağıdakı ikinci pillə sətrində iki oxunuşa bölünür — "Nəzərdə tutulmuş vəsait" (plan) və
"Faktiki xərc (01.04.2026)" (ilin əvvəlindən kumulyativ icra). Tək pilləli məntiq ikinci sütunu
səssiz atır və planı adi illik müşahidə kimi oxuyur; bu testlər həmin geri-düşməni bağlayır.
Həmçinin: vərəq adlarındakı sonda boşluq (`Sosial sektor `) və ikinci pillənin YANLIŞ tanınmaması
(adi məlumat sətri) yoxlanılır.

İşə salınma:
    python3 -m unittest tests.test_data_layer_loader -v
"""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl  # noqa: E402

from src import config as C  # noqa: E402
from src import data_layer as dl  # noqa: E402


def _sheet(rows, title="Test"):
    """Yaddaşda bir vərəq qurur (sətirlər siyahı-siyahı verilir)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    for r in rows:
        ws.append(list(r))
    return wb, ws


# İki pilləli başlıq — `DİP 2016-2026` vərəqinin quruluşunun kiçildilmiş nüsxəsi.
TWO_TIER = [
    ["Dövlət büdcəsində dövlət əsaslı vəsait qoyuluşu", None, None, None, None, None, None],
    ["Göstəricilər", "Ölçü vahidi", 2024, 2025, 2026, None, "Mənbə"],
    [None, None, None, None, "Nəzərdə tutulmuş vəsait", "Faktiki xərc\n(01.04.2026)", None],
    ["Dövlət əsaslı vəsait qoyuluşunun cəmi", "mln.manat", 2741.6, 2305.1, 2700.0, 365.4, "MN"],
    ["Sosial yönümlü layihələr", "mln.manat", 251.1, 331.8, 115.9, 9.6, "MN"],
]

# Tək pilləli başlıq — başlıqdan sonrakı sətir ADİ məlumat sətridir.
ONE_TIER = [
    ["Göstəricilər", "Ölçü vahidi", 2024, 2025, "Mənbə"],
    ["Sahə üzrə ÜDM", "milyon manat", 121.7, 134.2, "DSK"],
    ["muzdlu işçilərin orta siyahı sayı", "nəfər", 30560, 30758, "DSK"],
]


class TwoTierHeaderTests(unittest.TestCase):
    def test_year_row_plus_subheader_row(self) -> None:
        wb, ws = _sheet(TWO_TIER)
        tcols, tqual, recs = dl.load_sheet(ws, "DİP 2016-2026")

        # zaman sütunları: 2024/2025 adi illik; 2026 iki oxunuşa bölünür
        self.assertEqual(tcols[2], (2024, None, "annual"))
        self.assertEqual(tcols[3], (2025, None, "annual"))
        self.assertEqual(tcols[4], (2026, None, "annual"))       # nəzərdə tutulmuş vəsait
        self.assertEqual(tqual[4], "plan")
        # başlığı BOŞ olan sütun ili SOLDAN miras alır və vəziyyət tarixindən dövr çıxarılır
        self.assertEqual(tcols[5], (2026, 3, "monthly"))          # 01.04.2026 -> yanvar-mart
        self.assertEqual(tqual[5], "fakt")
        self.assertEqual(tqual[2], "")

        # ikinci pillə sətri MƏLUMAT sətri kimi oxunmur
        self.assertEqual([r["row"] for r in recs], [4, 5])
        self.assertEqual(recs[0]["name"], "Dövlət əsaslı vəsait qoyuluşunun cəmi")
        self.assertAlmostEqual(recs[0]["vals"][4], 2700.0)
        self.assertAlmostEqual(recs[0]["vals"][5], 365.4)

    def test_asof_date_convention(self) -> None:
        # ayın ilk yarısındakı tarix ƏVVƏLKİ ayın sonuna qədər kumulyativ deməkdir
        self.assertEqual(dl._asof_period("Faktiki xərc (01.04.2026)", 2026), (2026, 3, "monthly"))
        self.assertEqual(dl._asof_period("Faktiki xərc (20.04.2026)", 2026), (2026, 4, "monthly"))
        self.assertEqual(dl._asof_period("Faktiki xərc (01.01.2026)", 2026), (2025, 12, "monthly"))
        self.assertIsNone(dl._asof_period("Nəzərdə tutulmuş vəsait", 2026))

    def test_qualifier_tags(self) -> None:
        self.assertEqual(dl._subheader_qualifier("Nəzərdə tutulmuş vəsait"), "plan")
        self.assertEqual(dl._subheader_qualifier("Faktiki xərc\n(01.04.2026)"), "fakt")

    def test_single_tier_sheet_is_untouched(self) -> None:
        """Adi vərəqlərdə ikinci pillə TANINMAMALIDIR — əks halda ilk məlumat sətri itərdi."""
        wb, ws = _sheet(ONE_TIER)
        tcols, tqual, recs = dl.load_sheet(ws, "Su təchizatı")
        self.assertEqual(set(tqual.values()), {""})
        self.assertEqual([r["name"] for r in recs],
                         ["Sahə üzrə ÜDM", "muzdlu işçilərin orta siyahı sayı"])
        self.assertEqual(tcols[2], (2024, None, "annual"))

    def test_numeric_row_after_header_is_not_a_subheader(self) -> None:
        """Başlıqdan sonrakı sətirdə ƏDƏD varsa, o, ikinci pillə deyil (mətn tələb olunur)."""
        rows = [["Göstəricilər", "Ölçü vahidi", 2024, 2025, "Mənbə"],
                [None, None, 1.0, 2.0, None],
                ["Sahə üzrə ÜDM", "milyon manat", 121.7, 134.2, "DSK"]]
        wb, ws = _sheet(rows)
        tcols, tqual, recs = dl.load_sheet(ws, "Test")
        self.assertEqual(set(tqual.values()), {""})
        self.assertEqual([r["row"] for r in recs], [3])

    def test_trailing_space_sheet_name_resolves(self) -> None:
        """İş kitabında `Sosial sektor ` və `Elektrik enerjisi ` adları sonda boşluqla yazılıb;
        ETL vərəqi kəsilmiş adla tapmalıdır (kataloq və ABBR kəsilmiş adı işlədir)."""
        wb, ws = _sheet(ONE_TIER, title="Elektrik enerjisi ")
        nm = {n.strip(): n for n in wb.sheetnames}
        self.assertIn("Elektrik enerjisi", nm)
        tcols, tqual, recs = dl.load_sheet(wb[nm["Elektrik enerjisi"]], "Elektrik enerjisi")
        self.assertEqual(len(recs), 2)


class WorkbookRegistryTests(unittest.TestCase):
    """Keşdən (rebuild olunmuş registrdən) yoxlamalar — keş yoxdursa test ötürülür."""

    @classmethod
    def setUpClass(cls) -> None:
        if not os.path.exists(C.C_WB_REGISTRY) or not os.path.exists(C.C_WB_ANNUAL):
            raise unittest.SkipTest("keş qurulmayıb: data_layer.build_cache() işlədin")
        import pandas as pd
        cls.reg = pd.read_csv(C.C_WB_REGISTRY)
        cls.ann = pd.read_csv(C.C_WB_ANNUAL)

    def test_new_sheets_are_wired(self) -> None:
        sheets = set(self.reg["sheet"].unique())
        for s in ("Mədənçıxarma", "Su təchizatı", "DVX üzrə göstəricilər", "DİP 2016-2026"):
            self.assertIn(s, sheets, f"{s} vərəqi registrdə yoxdur")

    def test_dip_plan_and_actual_are_separate_series(self) -> None:
        codes = set(self.reg["series_code"])
        for c in ("dip_r004", "dip_r004_plan", "dip_r004_fakt"):
            self.assertIn(c, codes)
        base = self.ann[self.ann.series_code == "dip_r004"]
        plan = self.ann[self.ann.series_code == "dip_r004_plan"]
        # faktiki sıra 2025-də bitir (2026 planı ora düşmür), plan ayrıca koddadır
        self.assertEqual(int(base.year.max()), 2025)
        self.assertEqual(list(plan.year), [2026])
        self.assertAlmostEqual(float(plan.value.iloc[0]), 2700.0, places=6)

    def test_dip_cross_checks_fiscal_capex_2025(self) -> None:
        """DİP 2025 cəmi Fiskal sektorun dövlət əsaslı vəsait qoyuluşu sətri ilə üst-üstə düşür
        (iş kitabının öz yuvarlaqlaşdırması həddində: 2 305,12 vs 2 305,1).

        Qeyd: bu, YALNIZ üst-üstə düşən illərin bir hissəsi üçün doğrudur — 2021 və 2024-cü
        illərdə iki sıra əhəmiyyətli dərəcədə fərqlənir (kataloq qeydinə bax), ona görə burada
        yalnız 2025 çarpaz yoxlaması sabitlənir."""
        dip = self.ann[(self.ann.series_code == "dip_r004") & (self.ann.year == 2025)]
        fis = self.ann[(self.ann.series_code == "fiscal_r019") & (self.ann.year == 2025)]
        self.assertEqual(len(dip), 1)
        self.assertEqual(len(fis), 1)
        self.assertLess(abs(float(dip.value.iloc[0]) - float(fis.value.iloc[0])), 0.05)

    def test_no_actual_series_reaches_the_forecast_horizon(self) -> None:
        """D1: kataloqun faktiki obyekti heç bir halda 2025-dən kənara çıxmır."""
        import pandas as pd
        act = pd.read_csv(C.C_ACTUALS)
        self.assertLessEqual(int(act.year.max()), C.LAST_ACTUAL)


class SeriesDictionaryGuardTests(unittest.TestCase):
    """`outputs/series_dictionary.csv` keşin yenidən qurulması ilə QISALDILMAMALIDIR.

    Geri-düşmə: `build_cache()` təhvil faylının üstünə kataloq-yalnız başlanğıc cədvəli
    (~142 sətir) yazırdı; notebook-ların əlavə etdiyi sətirlər (tam fayl 402 sıra) tək
    başına çağırılan `build_cache()`-dən sonra itirdi və yalnız tam `run_all.py` bərpa
    edirdi. Qoruyucu: mövcud fayl başlanğıc cədvəldən böyükdürsə, ona toxunulmur.
    """

    def _starter(self, n):
        import pandas as pd
        return pd.DataFrame({
            "series_code": [f"kod_{i:03d}" for i in range(n)],
            "name_az": ["ad"] * n, "name_en": ["name"] * n, "unit": ["mln AZN"] * n,
            "fr": ["FR01"] * n, "statutory_sheet_ref": [""] * n,
            "price_basis": ["cari"] * n})

    def test_fat_file_survives_a_starter_write(self) -> None:
        import pandas as pd
        import tempfile
        with tempfile.TemporaryDirectory() as td:
            p = os.path.join(td, "series_dictionary.csv")
            fat = self._starter(402)
            fat.to_csv(p, index=False)
            with open(p, "rb") as fh:
                before = fh.read()
            wrote = dl._write_starter_dict(self._starter(142), path=p)
            self.assertFalse(wrote, "başlanğıc cədvəl böyük faylın üstünə yazıldı")
            with open(p, "rb") as fh:
                self.assertEqual(fh.read(), before)
            self.assertEqual(len(pd.read_csv(p)), 402)

    def test_missing_file_is_created(self) -> None:
        import pandas as pd
        import tempfile
        with tempfile.TemporaryDirectory() as td:
            p = os.path.join(td, "series_dictionary.csv")
            self.assertTrue(dl._write_starter_dict(self._starter(142), path=p))
            self.assertEqual(len(pd.read_csv(p)), 142)

    def test_short_or_corrupt_file_is_refreshed(self) -> None:
        import pandas as pd
        import tempfile
        with tempfile.TemporaryDirectory() as td:
            p = os.path.join(td, "series_dictionary.csv")
            self._starter(10).to_csv(p, index=False)          # yarımçıq nüsxə
            self.assertTrue(dl._write_starter_dict(self._starter(142), path=p))
            self.assertEqual(len(pd.read_csv(p)), 142)
            with open(p, "w"):                                # zədələnmiş (boş) nüsxə
                pass
            self.assertTrue(dl._write_starter_dict(self._starter(142), path=p))
            self.assertEqual(len(pd.read_csv(p)), 142)

    def test_build_cache_does_not_truncate_the_delivered_dictionary(self) -> None:
        """Tam `build_cache()` axını: ağır mərhələlər əvəz olunur, kataloq həlli əsldir."""
        import pandas as pd
        if not os.path.exists(C.C_WB_REGISTRY) or not os.path.exists(C.O_SERIES_DICT):
            raise unittest.SkipTest("keş və ya təhvil lüğəti yoxdur")
        with open(C.O_SERIES_DICT, "rb") as fh:
            before = fh.read()
        n_before = len(pd.read_csv(C.O_SERIES_DICT))
        if n_before <= len(pd.read_csv(C.C_DICT)):
            raise unittest.SkipTest("təhvil lüğəti başlanğıc cədvəldən böyük deyil")

        real_wb, real_vq = dl._build_workbook_cache, dl._build_vereq_cache
        dl._build_workbook_cache = lambda verbose=True: (None, None, None)
        dl._build_vereq_cache = lambda verbose=True: (None, None, None, None)
        try:
            dl.build_cache(verbose=False, check_md5=False)
        finally:
            dl._build_workbook_cache, dl._build_vereq_cache = real_wb, real_vq
            dl.clear_cache()

        with open(C.O_SERIES_DICT, "rb") as fh:
            self.assertEqual(fh.read(), before, "build_cache() təhvil lüğətini dəyişdi")
        self.assertEqual(len(pd.read_csv(C.O_SERIES_DICT)), n_before)


class OilGasBlockTests(unittest.TestCase):
    """Neft-qaz blokunun konstruksiyaları: etiket yoxlaması, eynilik, indekslər."""

    @classmethod
    def setUpClass(cls) -> None:
        if not os.path.exists(C.C_WB_REGISTRY) or not os.path.exists(C.C_ACTUALS):
            raise unittest.SkipTest("keş qurulmayıb: data_layer.build_cache() işlədin")
        cls.b = dl.oilgas_block()

    def test_every_source_row_keeps_its_label(self) -> None:
        """Sətir nömrəsi sürüşərsə və ya etiket dəyişərsə blok SƏSSİZ işləməməlidir."""
        for code in dl.OILGAS_SOURCES:
            s = dl.wb_labelled(code)
            self.assertGreater(len(s), 0, f"{code}: boş sıra")
        real = dict(dl.OILGAS_SOURCES)
        try:
            dl.OILGAS_SOURCES["oilgas_r003"] = ("Neft-Qaz sektoru", 3, "YANLIŞ ETİKET")
            with self.assertRaises(AssertionError):
                dl.wb_labelled("oilgas_r003")
        finally:
            dl.OILGAS_SOURCES.clear()
            dl.OILGAS_SOURCES.update(real)

    def test_domestic_identity_holds_every_year(self) -> None:
        """daxili = əmtəəlik − ixrac, hər il üçün, tolerans 1e-9."""
        d = self.b.dropna(subset=["q_oil_mkt", "q_oil_exp"])
        self.assertTrue(((d["q_oil_mkt"] - d["q_oil_exp"] - d["q_oil_dom"]).abs() < 1e-9).all())
        g = self.b.dropna(subset=["q_gas_mkt", "q_gas_exp"])
        self.assertTrue(((g["q_gas_mkt"] - g["q_gas_exp"] - g["q_gas_dom"]).abs() < 1e-9).all())

    def test_export_value_weight_is_a_share(self) -> None:
        w = self.b["w_oil_exp"].dropna()
        self.assertTrue(((w >= 0) & (w <= 1)).all())
        # qazın ixrac dəyərindəki payı 2013-dən 2025-ə kəskin qalxıb
        self.assertLess(1 - float(w.loc[2013]), 0.05)
        self.assertGreater(1 - float(w.loc[C.LAST_ACTUAL]), 0.35)

    def test_price_index_reduces_to_the_oil_index_when_gas_weight_is_one(self) -> None:
        import numpy as np
        import pandas as pd
        p_oil = pd.Series({2020: 100.0, 2021: 110.0, 2022: 121.0})
        p_gas = pd.Series({2020: 100.0, 2021: 200.0, 2022: 400.0})
        w1 = pd.Series({2020: 1.0, 2021: 1.0, 2022: 1.0})
        g = dl.weighted_price_growth(p_oil, p_gas, w1)
        self.assertTrue(np.allclose(g.values, [10.0, 10.0]))
        w0 = pd.Series({2020: 0.0, 2021: 0.0, 2022: 0.0})
        g0 = dl.weighted_price_growth(p_oil, p_gas, w0)
        self.assertTrue(np.allclose(g0.values, [100.0, 100.0]))

    def test_volume_index_is_a_value_weighted_average(self) -> None:
        import pandas as pd
        q_oil = pd.Series({2024: 100.0, 2025: 110.0})     # +10 %
        q_gas = pd.Series({2024: 100.0, 2025: 90.0})      # −10 %
        p_oil = pd.Series({2024: 3.0, 2025: 3.0})
        p_gas = pd.Series({2024: 1.0, 2025: 1.0})
        g = dl.laspeyres_volume_growth(q_oil, q_gas, p_oil, p_gas)
        self.assertAlmostEqual(float(g.loc[2025]), 0.75 * 10.0 + 0.25 * (-10.0), places=9)

    def test_constructions_never_reach_the_forecast_horizon(self) -> None:
        self.assertLessEqual(int(self.b.index.max()), C.LAST_ACTUAL)
        self.assertLessEqual(int(dl.oilgas_price_growth().index.max()), C.LAST_ACTUAL)
        self.assertLessEqual(int(dl.oilgas_volume_growth().index.max()), C.LAST_ACTUAL)


class FiscalBlockTests(unittest.TestCase):
    """Fiskal blok və investisiya deflyatoru: etiket yoxlaması, eynilik, provenans."""

    @classmethod
    def setUpClass(cls) -> None:
        if not os.path.exists(C.C_WB_REGISTRY) or not os.path.exists(C.C_ACTUALS):
            raise unittest.SkipTest("keş qurulmayıb: data_layer.build_cache() işlədin")

    def test_every_fiscal_source_row_keeps_its_label(self) -> None:
        """Fiskal sıralar da neft-qaz bloku ilə eyni etiket yoxlamasından keçir."""
        for code in dl.FISCAL_SOURCES:
            s = dl.wb_labelled(code)
            self.assertGreater(len(s), 0, f"{code}: boş sıra")
        real = dict(dl.FISCAL_SOURCES)
        try:
            dl.FISCAL_SOURCES["fiscal_r016"] = ("Fiskal sektor", 16, "YANLIŞ ETİKET")
            with self.assertRaises(AssertionError):
                dl.wb_labelled("fiscal_r016")
        finally:
            dl.FISCAL_SOURCES.clear()
            dl.FISCAL_SOURCES.update(real)

    def test_state_capex_splice_fills_only_the_missing_years(self) -> None:
        """DİP birləşməsi YALNIZ 2022-2023-ü doldurur; digər illərdə fiskal vərəq saxlanılır."""
        raw = dl.wb_labelled("fiscal_r019")
        raw = raw[raw.index <= C.LAST_ACTUAL]
        cap = dl.state_capex()
        self.assertTrue({2022, 2023} <= set(cap.index))
        common = [y for y in raw.dropna().index if y in cap.index]
        for y in common:
            self.assertAlmostEqual(float(cap.loc[y]), float(raw.loc[y]), places=9,
                                   msg=f"{y}: birləşmə mövcud fiskal dəyəri əvəz etdi")

    def test_state_capex_gap_is_published_not_corrected(self) -> None:
        """İki mənbənin fərqi düzəldilmir — 2021 və 2024 fərqləri görünməlidir."""
        gap = dl.state_capex_gap()
        self.assertIn(2021, gap.index)
        self.assertGreater(abs(float(gap.loc[2021, "fərq_%"])), 10.0)

    def test_nonoil_base_balance_identity(self) -> None:
        """balans = qeyri-neft gəlirlər − xərclər; nisbət = balans / qeyri-neft ÜDM, HƏR il."""
        nob = dl.nonoil_base_balance()
        self.assertTrue(((nob["balance"] - (nob["revenue_nonoil"] - nob["expenditure"])).abs()
                         < 1e-9).all())
        self.assertTrue(((nob["ratio_nonoil_gdp"] - nob["balance"] / nob["gdp_nonoil"] * 100.0).abs()
                         < 1e-9).all())
        self.assertLessEqual(int(nob.index.max()), C.LAST_ACTUAL)

    def test_moe_sample_tables_stop_at_2024(self) -> None:
        """Nazirliyin nümunə-model faylları YALNIZ faktiki sütunlara qədər saxlanılır (sətir-sətir təsnifat §4d)."""
        self.assertLessEqual(int(max(dl.MOE_ICMAL_BUDGET)), 2024)
        self.assertLessEqual(int(max(dl.MOE_INV_DEFL)), 2024)
        self.assertLessEqual(int(dl.moe_icmal_budget().index.max()), 2024)
        self.assertLessEqual(int(dl.moe_investment_deflators().index.max()), 2024)

    def test_moe_icmal_ratio_masked_before_the_broken_years(self) -> None:
        """41-ci sətrin məxrəci 2015-ə qədər sınıqdır — həmin illər NaN olmalıdır."""
        ic = dl.moe_icmal_budget()
        pre = ic.loc[ic.index < dl.MOE_ICMAL_RATIO_VALID_FROM, "ratio_nonoil_gdp"]
        self.assertTrue(pre.isna().all())
        self.assertFalse(ic.loc[2020:2024, "ratio_nonoil_gdp"].isna().any())

    def test_moe_constants_match_source_when_available(self) -> None:
        """Fayl əlçatandırsa, dondurulmuş cədvəllər mənbədən yenidən qurulub tutuşdurulur."""
        for fn, path in ((dl.verify_moe_icmal_budget, dl.MOE_FISCAL_XLSX),
                         (dl.verify_moe_investment_deflators, dl.MOE_SNA_XLSX)):
            diff = fn(path)
            if diff is None:
                continue                                  # fayl yoxdur — xəta deyil
            self.assertLess(float(diff), 1e-3)

    def test_investment_deflator_matches_the_ministry_2016_value(self) -> None:
        """Çarpaz yoxlama: implisit deflyatorun 2016 dəyəri Nazirliyin öz cəmi
        deflyator sətri (26,23964 %) ilə üst-üstə düşür."""
        g = dl.investment_deflator_growth()
        self.assertAlmostEqual(float(g.loc[2016]), 26.23964, places=3)
        self.assertLessEqual(int(g.index.max()), C.LAST_ACTUAL)
        self.assertGreaterEqual(int(g.index.min()), dl.INV_DEFLATOR_FIRST_YEAR)

    def test_construction_share_is_a_share(self) -> None:
        w = dl.construction_works_share()
        self.assertTrue(((w > 0) & (w < 1)).all())

    def test_soe_series_is_too_short_to_estimate(self) -> None:
        """D5 intizamı: `real_r114` illik sütunu üç müşahidədir — sənədləşdirilmiş fakt."""
        s = dl.soe_investment_annual()
        self.assertLessEqual(len(s), 5)
        self.assertLessEqual(int(s.index.max()), C.LAST_ACTUAL)


class ReferenceForecastTests(unittest.TestCase):
    """§16 — BVF etalon yolları. Bu sıralar YALNIZ müqayisə üçündür; testlər onların
    (i) ağ siyahı ilə məhdudlaşdığını, (ii) buraxılış qeydi daşıdığını, (iii) uyğun sıra
    olmadıqda susmadığını və (iv) heç bir fərziyyə açarına sızmadığını qoruyur."""

    def test_only_the_two_imf_blocks_are_read(self) -> None:
        """Yalnız BVF-nin iki bloku oxunur. Fayl özü dondurulmuş çıxarışdır, lakin ağ siyahı
        qərarı KODDA da bağlayır: başqa bir modelin bloku fayla düşsə belə, oxunmur — paketdə
        ikinci bir «kanonik» rəqəm dəsti yaranmır."""
        data = dl.reference_forecasts()
        self.assertEqual(sorted(data), sorted(dl.REFERENCE_BLOCKS))
        self.assertNotIn("CAEM_xlsb", data)
        self.assertNotIn("AZ_baseline", data)
        with self.assertRaises(KeyError):
            dl.reference_series("gdp_realg", block="CAEM_xlsb")

    def test_vintage_note_is_attached_to_every_provenance_row(self) -> None:
        prov = dl.reference_provenance()
        self.assertTrue(len(prov) > 0)
        self.assertTrue((prov["vintage"] == dl.REFERENCE_VINTAGE_NOTE).all())
        self.assertEqual(dl.REFERENCE_VINTAGE_NOTE, "BVF Maddə IV 26/112, aprel 2026 vintajı")

    def test_unmatched_series_are_named_with_a_reason(self) -> None:
        """Uyğun sıra olmayan göstəricilər susmaqla ötürülmür: hər biri səbəbi ilə sadalanır."""
        rep = dl.reference_match_report()
        matched = set(rep.loc[rep["matched"], "imf_var"])
        self.assertEqual(matched, {"gdp_realg", "nonoil_realg", "cpi_infl"})
        unmatched = rep[~rep["matched"]]
        self.assertTrue(len(unmatched) > 0)
        self.assertTrue((unmatched["reason"].astype(str).str.len() > 20).all())
        self.assertTrue(unmatched["package_series_code"].isna().all())

    def test_published_rows_carry_no_bands_and_stay_in_horizon(self) -> None:
        rows = dl.reference_rows(C.FY, {"gdp_realg": "FR1", "cpi_infl": "FR9"})
        self.assertEqual(len(rows), 2 * len(list(C.FY)))
        self.assertEqual(set(rows["source"]), {"imf_reference"})
        self.assertEqual(set(rows["kind"]), {"forecast"})
        self.assertGreaterEqual(int(rows["year"].min()), C.FIRST_FORECAST)
        self.assertLessEqual(int(rows["year"].max()), C.LAST_FORECAST)
        for col in ("lo80", "hi80", "lo50", "hi50"):
            self.assertTrue(rows[col].isna().all(), col)
        with self.assertRaises(KeyError):
            dl.reference_rows(C.FY, {"ca_gdp_ratio": "FR7"})

    def test_article_iv_and_weo_agree_where_they_overlap(self) -> None:
        chk = dl.verify_reference_forecasts()
        self.assertTrue(len(chk) > 0)
        self.assertTrue(bool(chk["agree"].all()))

    def test_reference_paths_never_enter_the_assumption_set(self) -> None:
        """D1 sərhədi: BVF yolları ekzogen giriş DEYİL — fərziyyə faylında izi olmamalıdır."""
        import csv
        with open(C.ASSUMPTIONS, encoding="utf-8") as fh:
            keys = {row["assumption_key"] for row in csv.DictReader(fh)}
        self.assertFalse({k for k in keys if "imf" in k.lower() or "bvf" in k.lower()})


if __name__ == "__main__":
    unittest.main()
