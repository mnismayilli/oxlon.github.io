"""tests/test_outputs_contract.py — §3 dashboard müqaviləsinin qorunması üçün yoxlamalar.

Real `outputs/` və `data/` qovluqlarına TOXUNMUR: hər test modulun yol sabitlərini müvəqqəti
qovluğa yönləndirir (monkeypatch), sonra əsl paket faylları toxunulmamış qalır. İşə salınma:
    python3 -m unittest tests.test_outputs_contract -v
(stdlib `unittest` istifadə olunur — əlavə asılılıq (məsələn pytest) tələb edilmir, §7-yə uyğun.)
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import pandas as pd  # noqa: E402

from src import outputs as O  # noqa: E402


class OutputsContractTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        tmp = Path(self._tmp.name)
        # Modulun qlobal yol sabitlərini müvəqqəti qovluğa yönləndiririk ki, əsl paket
        # məlumatlarına (outputs/, data/assumptions.csv) toxunulmasın.
        self._orig = {
            "DATA_DIR": O.DATA_DIR, "OUT_DIR": O.OUT_DIR,
            "ASSUMPTIONS_PATH": O.ASSUMPTIONS_PATH,
            "EQUATIONS_CATALOG_PATH": O.EQUATIONS_CATALOG_PATH,
            "FORECAST_LONG_PATH": O.FORECAST_LONG_PATH,
            "VALIDATION_BACKTEST_PATH": O.VALIDATION_BACKTEST_PATH,
            "SERIES_DICTIONARY_PATH": O.SERIES_DICTIONARY_PATH,
        }
        O.DATA_DIR = tmp / "data"
        O.OUT_DIR = tmp / "outputs"
        O.DATA_DIR.mkdir(parents=True, exist_ok=True)
        O.OUT_DIR.mkdir(parents=True, exist_ok=True)
        O.ASSUMPTIONS_PATH = O.DATA_DIR / "assumptions.csv"
        O.EQUATIONS_CATALOG_PATH = O.OUT_DIR / "equations_catalog.csv"
        O.FORECAST_LONG_PATH = O.OUT_DIR / "forecast_long.csv"
        O.VALIDATION_BACKTEST_PATH = O.OUT_DIR / "validation_backtest.csv"
        O.SERIES_DICTIONARY_PATH = O.OUT_DIR / "series_dictionary.csv"

    def tearDown(self) -> None:
        for k, v in self._orig.items():
            setattr(O, k, v)
        self._tmp.cleanup()

    def test_series_dictionary_headers_and_idempotency(self) -> None:
        df1 = pd.DataFrame({
            "series_code": ["gdp_real_growth", "cpi_yoy"],
            "name_az": ["ÜDM real artım tempi", "İQİ illik artımı"],
            "name_en": ["Real GDP growth", "CPI inflation"],
            "unit": ["%", "%"],
            "fr": [1, 9],
            "statutory_sheet_ref": ["2.4.1.1", "2.4.1.9"],
        })
        out = O.write_series_dictionary(df1)
        self.assertEqual(list(out.columns), O.SERIES_DICTIONARY_COLUMNS)
        with open(O.SERIES_DICTIONARY_PATH) as fh:
            header = fh.readline().strip().split(",")
        self.assertEqual(header, O.SERIES_DICTIONARY_COLUMNS)
        self.assertEqual(len(out), 2)

        # Eyni sətirlərlə TƏKRAR yazma — dublikat yaranmamalıdır.
        out2 = O.write_series_dictionary(df1)
        self.assertEqual(len(out2), 2)

        # Bir FR-in sətirini DƏYİŞDİRİB yenidən yazmaq — köhnə dəyər əvəzlənməli, sətir sayı sabit qalmalı.
        df1_mod = df1.copy()
        df1_mod.loc[0, "name_az"] = "ÜDM real artım tempi (yenilənmiş)"
        out3 = O.write_series_dictionary(df1_mod)
        self.assertEqual(len(out3), 2)
        self.assertEqual(
            out3.loc[out3.series_code == "gdp_real_growth", "name_az"].iloc[0],
            "ÜDM real artım tempi (yenilənmiş)",
        )

    def test_assumptions_scope_replace(self) -> None:
        df_brent_v1 = pd.DataFrame({
            "assumption_key": ["brent_usd"] * 3,
            "year": [2026, 2027, 2028],
            "value": [70.0, 71.0, 72.0],
            "unit": ["USD/barel"] * 3,
            "note": ["FR5 modeli"] * 3,
        })
        O.write_assumptions(df_brent_v1)
        # FR5 təkrar hesablanır: fərqli illər/dəyərlər (2029 əlavə olunur, 2026 dəyəri dəyişir).
        df_brent_v2 = pd.DataFrame({
            "assumption_key": ["brent_usd"] * 4,
            "year": [2026, 2027, 2028, 2029],
            "value": [69.5, 71.0, 72.5, 73.0],
            "unit": ["USD/barel"] * 4,
            "note": ["FR5 modeli (yenidən hesablanmış)"] * 4,
        })
        result = O.write_assumptions(df_brent_v2)
        brent_rows = result[result.assumption_key == "brent_usd"]
        self.assertEqual(len(brent_rows), 4)  # köhnə 3 sətir tam əvəzləndi, 4 yeni sətir qaldı
        self.assertAlmostEqual(float(brent_rows[brent_rows.year == 2026].value.iloc[0]), 69.5)

        paths = O.read_assumptions()
        self.assertIn("brent_usd", paths)
        self.assertEqual(list(paths["brent_usd"].index), [2026, 2027, 2028, 2029])

    def test_missing_column_raises(self) -> None:
        bad = pd.DataFrame({"assumption_key": ["cbar_target"], "year": [2026], "value": [4.0]})
        with self.assertRaises(O.OutputsContractError):
            O.write_assumptions(bad)  # `unit`, `note` çatışmır

    def test_null_in_required_column_raises(self) -> None:
        bad = pd.DataFrame({
            "assumption_key": ["cbar_target"], "year": [2026], "value": [None],
            "unit": ["%"], "note": ["ekzogen fərziyyə"],
        })
        with self.assertRaises(O.OutputsContractError):
            O.write_assumptions(bad)  # `value` boş ola bilməz

    def test_forecast_long_nullable_bands(self) -> None:
        df = pd.DataFrame({
            "fr": [5, 5],
            "series_code": ["brent_usd", "brent_usd"],
            "series_name_az": ["Brent neft qiyməti", "Brent neft qiyməti"],
            "unit": ["USD/barel", "USD/barel"],
            "year": [2025, 2026],
            "kind": ["actual", "forecast"],
            "source": ["ours", "ours"],
            "value": [68.0, 70.0],
            "lo80": [None, 65.0], "hi80": [None, 75.0],
            "lo50": [None, 67.0], "hi50": [None, 73.0],
        })
        out = O.write_forecast_long(df)
        self.assertEqual(len(out), 2)
        self.assertEqual(list(out.columns), O.FORECAST_LONG_COLUMNS)

    def test_forecast_long_source_vocabulary(self) -> None:
        """`source` yalnız müqavilədəki dəyərləri qəbul edir: `template_sample` (şablon-nümunə
        ssenarisi) və boş `ministry_official` slotu keçir, köhnə `ministry_decree75` etiketi isə
        artıq QƏBUL EDİLMİR — relabel geriyə sürüşə bilməz."""
        def row(source: str) -> pd.DataFrame:
            return pd.DataFrame([{
                "fr": "FR6", "series_code": "inv_total_nom", "series_name_az": "Cəmi investisiya",
                "unit": "mln AZN", "year": 2026, "kind": "forecast", "source": source,
                "value": 24437.0, "lo80": None, "hi80": None, "lo50": None, "hi50": None,
            }])

        self.assertIn("template_sample", O.SOURCE_LABELS_AZ)
        self.assertEqual(O.SOURCE_LABELS_AZ["template_sample"],
                         "şablon-nümunə ssenarisi (75 saylı qərar formatı üzrə)")
        self.assertIn("ministry_official", O.EMPTY_SOURCE_SLOTS)
        self.assertEqual(len(O.write_forecast_long(row("template_sample"))), 1)
        # «Nazirlik spesifikasiyası» ssenarisinin mənbəsi (FR13) də lüğətdədir və qəbul olunur
        self.assertIn("ministry_spec", O.SOURCE_LABELS_AZ)
        self.assertEqual(list(O.write_forecast_long(row("ministry_spec"))["source"]),
                         ["ministry_spec"])
        # BVF etalonu — beşinci, POPULYASİYA OLUNMUŞ mənbə
        self.assertIn("imf_reference", O.SOURCE_LABELS_AZ)
        self.assertEqual(O.SOURCE_LABELS_AZ["imf_reference"],
                         "BVF etalon proqnozu (Maddə IV 26/112 · WEO)")
        self.assertEqual(list(O.write_forecast_long(row("imf_reference"))["source"]),
                         ["imf_reference"])
        # `fr` əhatəsi üzrə yerinə-yazma: eyni FR-in yeni partiyası köhnəni əvəz edir
        out = O.write_forecast_long(row("ministry_official"))
        self.assertEqual(list(out["source"]), ["ministry_official"])
        with self.assertRaises(O.OutputsContractError):
            O.write_forecast_long(row("ministry_decree75"))
        # BOŞ `imf_weo` / `imf_artiv` slotları `imf_reference` ilə əvəzlənib —
        # köhnə etiketlər geri sürüşə bilməz
        for retired in ("imf_weo", "imf_artiv"):
            self.assertNotIn(retired, O.SOURCE_LABELS_AZ)
            with self.assertRaises(O.OutputsContractError):
                O.write_forecast_long(row(retired))

    def test_series_dictionary_price_basis(self) -> None:
        """D5: hər sətir qiymət bazasını daşıyır — açıq verilmədikdə ölçü vahidindən çıxarılır,
        sabit qiymətli sıra isə onu AÇIQ elan edir."""
        df = pd.DataFrame({
            "series_code": ["inv_total_nom", "inv_total_real2025", "inv_total_g", "emp_mining"],
            "name_az": ["Cəmi investisiya (cari)", "Cəmi investisiya (2025 sabit)",
                        "Real artım tempi", "İşçi sayı"],
            "name_en": ["Investment (current)", "Investment (constant 2025)",
                        "Real growth", "Employees"],
            "unit": ["mln AZN", "mln AZN", "%", "nəfər"],
            "fr": ["FR6", "FR6", "FR6", "FR8"],
            "statutory_sheet_ref": ["8 vərəq", "8 vərəq", "8 vərəq", "Mədənçıxarma"],
            "price_basis": [None, "sabit-2025", None, None],
        })
        out = O.write_series_dictionary(df)
        got = out.set_index("series_code")["price_basis"].to_dict()
        self.assertEqual(got["inv_total_nom"], "cari")        # vahiddən: pul ifadəsi
        self.assertEqual(got["inv_total_real2025"], "sabit-2025")   # açıq elan
        self.assertEqual(got["inv_total_g"], "indeks")        # faiz sırası
        self.assertEqual(got["emp_mining"], "natural")        # natural göstərici
        self.assertIn("price_basis", O.SERIES_DICTIONARY_COLUMNS)

        bad = df.copy()
        bad.loc[1, "price_basis"] = "constant-2025"
        with self.assertRaises(O.OutputsContractError):
            O.write_series_dictionary(bad)

    def test_backtest_overwrite_guard(self) -> None:
        """İki müxtəlif sınaq eyni `(series_code, model)` cütünə yazdıqda yazma DAYANDIRILMALIDIR;
        eyni sınağın təkrar icrası isə maneəsiz keçməlidir (idempotentlik)."""
        def batch(code: str, rw: float, rmse: float) -> pd.DataFrame:
            return pd.DataFrame([{
                "series_code": code, "model": "STRUKTUR", "vintage_year": 2020, "horizon_h": 1,
                "actual": 1.0, "forecast": 1.1, "error": 0.1, "abs_error": 0.1,
                "rmse_h": rmse, "rw_rmse_h": rw, "coverage80": 0.8,
            }])

        O.write_backtest(batch("inv_total_g", 12.24, 8.15))
        # eyni sınaq, eyni rəqəmlər → dublikat yaranmır, istisna qalxmır
        self.assertEqual(len(O.write_backtest(batch("inv_total_g", 12.24, 8.15))), 1)
        # BAŞQA pəncərədə aparılmış sınaq (fərqli etalon xətası) → səssiz əvəzlənmə əvəzinə istisna
        with self.assertRaises(O.BacktestOverwriteError):
            O.write_backtest(batch("inv_total_g", 9.67, 7.08))
        # ad fəzasında ayrıldıqdan sonra hər iki sınaq yan-yana yaşayır
        out = O.write_backtest(batch("inv_total_g_fr1_core", 9.67, 7.08))
        self.assertEqual(sorted(out.series_code.unique()), ["inv_total_g", "inv_total_g_fr1_core"])


if __name__ == "__main__":
    unittest.main()
