"""tests/test_assumption_registry.py — fərziyyə açarlarının əvəzləmə reyestri (B21).

Reyestr `assumptions.csv`-in İDARƏETMƏ qatıdır: hər açarın hansı rejimdə rəsmi göstərici ilə
əvəz oluna biləcəyini və dəyərin sahibi olan qurumu təsbit edir. Məzmun paketinin 01 saylı
sənədi (əvəzləmə müqaviləsi) məhz bu fayldan qurulur, ona görə fayl fərziyyə dəsti ilə
BİRƏBİR uyğun olmalıdır — əks halda sənəd səssizcə köhnəlir.

Testlərin bir hissəsi müvəqqəti qovluqda süni fayllarla işləyir (əsl paket fayllarına
toxunulmur), sonuncu test isə paketin ÖZ reyestrini yoxlayır.
    python3 -m unittest tests.test_assumption_registry -v
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import pandas as pd  # noqa: E402

from src import outputs as O  # noqa: E402

_ASM = pd.DataFrame({
    "assumption_key": ["brent_usd", "brent_usd", "oil_realg", "oil_realg"],
    "year": [2026, 2027, 2026, 2027],
    "value": [68.98, 68.97, -2.41, -1.26],
    "unit": ["USD/barel", "USD/barel", "%", "%"],
    "note": ["FR05", "FR05", "hasilat planı", "hasilat planı"],
})
_REG = pd.DataFrame({
    "assumption_key": ["brent_usd", "oil_realg"],
    "override_class": ["B", "A"],
    "owner_org": ["modul — FR05", "İqtisadiyyat Nazirliyi (hasilat planı)"],
    "basis": ["struktur tənliyi", "rəsmi hasilat planı"],
})


class AssumptionRegistryTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        tmp = Path(self._tmp.name)
        self._orig = {"DATA_DIR": O.DATA_DIR,
                      "ASSUMPTIONS_PATH": O.ASSUMPTIONS_PATH,
                      "ASSUMPTION_REGISTRY_PATH": O.ASSUMPTION_REGISTRY_PATH}
        O.DATA_DIR = tmp / "data"
        O.DATA_DIR.mkdir(parents=True, exist_ok=True)
        O.ASSUMPTIONS_PATH = O.DATA_DIR / "assumptions.csv"
        O.ASSUMPTION_REGISTRY_PATH = O.DATA_DIR / "assumption_registry.csv"
        _ASM.to_csv(O.ASSUMPTIONS_PATH, index=False)

    def tearDown(self) -> None:
        for k, v in self._orig.items():
            setattr(O, k, v)
        self._tmp.cleanup()

    def _write(self, reg: pd.DataFrame) -> None:
        reg.to_csv(O.ASSUMPTION_REGISTRY_PATH, index=False)

    def test_matching_registry_passes(self) -> None:
        self._write(_REG)
        out = O.validate_assumption_registry()
        self.assertEqual(list(out.columns), O.ASSUMPTION_REGISTRY_COLUMNS)
        self.assertEqual(len(out), 2)

    def test_missing_key_is_rejected(self) -> None:
        """Yeni fərziyyə açarı reyestrə salınmasa, boru xətti dayanmalıdır."""
        self._write(_REG[_REG.assumption_key != "oil_realg"])
        with self.assertRaises(O.AssumptionRegistryError) as ctx:
            O.validate_assumption_registry()
        self.assertIn("oil_realg", str(ctx.exception))

    def test_orphan_key_is_rejected(self) -> None:
        extra = pd.concat([_REG, pd.DataFrame([{
            "assumption_key": "kohne_acar", "override_class": "B",
            "owner_org": "modul — FR05", "basis": "artıq mövcud deyil"}])])
        self._write(extra)
        with self.assertRaises(O.AssumptionRegistryError) as ctx:
            O.validate_assumption_registry()
        self.assertIn("kohne_acar", str(ctx.exception))

    def test_unknown_class_is_rejected(self) -> None:
        bad = _REG.copy()
        bad.loc[bad.assumption_key == "brent_usd", "override_class"] = "D"
        self._write(bad)
        with self.assertRaises(O.AssumptionRegistryError):
            O.validate_assumption_registry()

    def test_official_key_needs_owner(self) -> None:
        """A sinfi rəsmi girişdir: sahib qurum olmadan əvəzləmə qeydi mənbəsiz qalar.
        (Tam boş xana artıq §3 sütun yoxlamasında dayanır — burada yalnız boşluqdan ibarət
        dəyər sınanır, çünki o, NaN yoxlamasından keçir.)"""
        bad = _REG.copy()
        bad.loc[bad.assumption_key == "oil_realg", "owner_org"] = "   "
        self._write(bad)
        with self.assertRaises(O.AssumptionRegistryError) as ctx:
            O.validate_assumption_registry()
        self.assertIn("sahib qurum", str(ctx.exception))

    def test_missing_registry_file_is_rejected(self) -> None:
        with self.assertRaises(O.AssumptionRegistryError):
            O.validate_assumption_registry()


class PackageRegistryTests(unittest.TestCase):
    """Paketin ÖZ faylları — müvəqqəti qovluq yoxdur, real reyestr yoxlanılır."""

    def test_package_registry_matches_assumption_set(self) -> None:
        reg = O.validate_assumption_registry()
        self.assertEqual(len(reg), len(set(O.read_assumptions().keys())))
        self.assertTrue((reg["override_class"] == "A").any(),
                        "ən azı bir rəsmi giriş açarı olmalıdır")


if __name__ == "__main__":
    unittest.main()
