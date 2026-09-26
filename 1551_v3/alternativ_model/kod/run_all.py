#!/usr/bin/env python3
"""run_all.py — MİİS §15.5.1 (FR1–FR12) borusunun tək-əmrli işə salınması.

DAG sırası (D2): FR05 (neft) → FR10 (məzənnə) → FR05b (neft-qaz bloku) →
FR09 (İQİ) → FR01 (1-ci keçid: tələb nüvəsi və deflyatorlar) → FR02 (aşağıdan-yuxarı istehsal bloku)
→ FR01 (2-ci keçid: baş göstərici FR02-nin aqreqatından) → {FR03, FR04, FR06, FR08} (paralel qrup —
sıra əhəmiyyətsizdir, hamısı yuxarı axının nəticəsindən asılıdır) → FR07 → FR01b (resurslar–istifadə
balansı) → FR11 → FR12 → FR13_nazirlik_spes («Nazirlik spesifikasiyası» ssenarisi, SONUNCU mərhələ —
yuxarı axının bütün çıxışlarını oxuyur).
Beləliklə 15 dəftər üzrə 16 mərhələ icrası aparılır: FR01 layihə üzrə İKİ dəfə işləyir və bunun
səbəbi aşağıda, DAG tərifinin yanında izah olunub.

16 mərhələ bitdikdən sonra İKİ yekun addım işləyir: `neticeler_15_5_1.xlsx` (Excel təqdimatı) və
`neticeler_15_5_1.json` (JSON təqdimatı) hazır çıxış fayllarından yenidən yığılır. MİİS §15.5.1-in
FR1–FR12 üzrə alt-tapşırıq 5-i hər iki formatı tələb etdiyi üçün bu addımlar müqavilə çıxışıdır və
uğursuzluq halında boru xətti UĞURSUZ sayılır. Onlar DAG-ın mərhələsi DEYİL (dəftər icra etmirlər)
və mərhələ sayına daxil edilmirlər.

Hər dəftər `jupyter nbconvert --to notebook --execute --inplace` ilə işə salınır (nəticə eyni
faylın üzərinə yazılır — icra izləri saxlanılır). Dəftərlər hələ yaradılmayıbsa, boru XƏTA ilə
YIXILMIR — çatışmayan dəftərlərin siyahısını aydın cədvəl şəklində göstərib səliqəli çıxır.
Bir mərhələ icra zamanı uğursuz olarsa (nbconvert sıfırdan fərqli çıxış kodu qaytarırsa), boru
DAYANIR — aşağı axın mərhələləri yuxarı axının çıxışından asılı olduğu üçün davam etmək yanlış
nəticəyə aparardı.

İstifadə:
    python3 run_all.py            # bütün DAG-ı ardıcıllıqla işə salır
    python3 run_all.py --stage FR09   # yalnız bir FR-i işə salır (məsələn təkrar-icra üçün)
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import List, Optional

BASE_DIR = Path(__file__).resolve().parent
try:
    from src import config as _cfg  # type: ignore
    NOTEBOOKS_DIR = Path(_cfg.BASE) / "notebooks"
except Exception:
    _cfg = None
    NOTEBOOKS_DIR = BASE_DIR / "notebooks"

# DAG (D2) — hər daxili siyahı bir "mərhələ"dir; mərhələ daxilindəki elementlər bir-birindən asılı
# deyil (yalnız yuxarı mərhələlərdən asılıdır), ona görə ardıcıl icra edilsə də sıra önəmsizdir.
DAG: List[List[str]] = [
    ["FR05"],
    ["FR10"],
    ["FR05b"],      # neft-qaz bloku: həcm × qiymət → oil_realg və çəkili qiymət indeksi
    ["FR09"],
    # BLOK SIRASI: baş göstərici FR02-nin aşağıdan-yuxarı istehsal aqreqatıdır,
    # halbuki FR02 özü FR01-in yazdığı açarlardan (oil_defl, g_inv_core, partner_gdp_realg,
    # nonoil_realg_demand, nonoil_defl_demand) asılıdır. Blok səviyyəsində asılılıq DÖVRƏVİ DEYİL —
    # FR01-in nüvəsi və deflyator tənlikləri FR02-dən asılı olmadığı üçün ardıcıllıq belədir:
    #   FR01-nüvə → FR02-istehsal bloku → FR01-baş göstərici.
    # Dəftər bölünmədiyi üçün FR01 DAG-da iki dəfə görünür; ikinci keçid birincinin yazdığı bütün
    # açarları eyni dəyərlərlə təkrar yazır (nüvə FR02-dən asılı deyil), ona görə boru xətti iki
    # keçiddə DƏQİQ bağlanır və təkrar icra eyni nəticəni verir (idempotent).
    ["FR01"],       # 1-ci keçid: nüvə, deflyatorlar, FR02-nin tələb etdiyi fərziyyə açarları
    ["FR02"],       # aşağıdan-yuxarı istehsal bloku (11 sahə + xalis vergilərin A4 bölgüsü)
    ["FR01"],       # 2-ci keçid: baş göstərici FR02-nin aqreqatından yığılır
    ["FR03", "FR04", "FR06", "FR08"],
    ["FR07"],
    # Resurslar–istifadə (tələb-təklif) balansı. FR07-dən SONRA gəlir, çünki ixracın və idxalın
    # manatla yolu FR07-nin dollarla axınlarından və FR10-un orta illik məzənnəsindən qurulur;
    # FR11/FR12-dən ƏVVƏL gəlir, çünki FR12-nin paket səviyyəli eynilik zənciri onun
    # `use_*` sətirlərini oxuyur. Öz sətirlərindən başqa heç nəyə yazmır (`fr=FR1B`).
    ["FR01b"],
    ["FR11"],
    ["FR12"],
    # «Nazirlik spesifikasiyası» ssenarisi. Boru xəttinin SONUNCU mərhələsidir, çünki
    # yuxarı axının bütün çıxışlarını (fərziyyələr, proqnoz yolları, sınaq cədvəli) OXUYUR və
    # yalnız öz sətirlərini (`fr=FR13`, `source=ministry_spec`, `model=NAZIRLIK_SPES`) yazır.
    ["FR13_nazirlik_spes"],
]
ALL_FRS = sorted({fr for stage in DAG for fr in stage})

NBCONVERT_TIMEOUT_SEC = 1800  # bir dəftərin maksimum icra vaxtı (30 dəq) — asılı qalmanın qarşısı


def _normalise_stage(raw: str) -> str:
    """"FR9", "fr09", "9" → "FR09"; "fr5b", "05B" → "FR05b" formatına salır.

    Hərfi şəkilçi (məsələn neft-qaz blokunun `FR05b` mərhələsi) SAXLANILIR: şəkilçi
    atılsaydı `--stage FR05b` səssizcə FR05-i işə salardı.
    """
    m = re.match(r"^\s*(?:fr)?\s*(\d+)\s*([a-z]?)\s*(_[A-Za-z0-9_]+)?\s*$",
                 str(raw).strip(), flags=re.IGNORECASE)
    if not m:
        raise ValueError(f"tanınmayan mərhələ adı: {raw!r}")
    base = f"FR{int(m.group(1)):02d}{(m.group(2) or '').lower()}"
    if m.group(3):                       # məsələn `FR13_nazirlik_spes`
        return base + m.group(3)
    # Şəkilçili dəftər adları (`FR13_nazirlik_spes`) qısa formada da yazıla bilsin deyə
    # DAG-dakı adlar arasında prefiks üzrə axtarılır.
    for fr in ALL_FRS:
        if fr == base or fr.startswith(base + "_"):
            return fr
    return base


def _notebook_path(fr: str) -> Path:
    return NOTEBOOKS_DIR / f"{fr}.ipynb"


def _print_table(rows: List[List[str]], headers: List[str]) -> None:
    widths = [max(len(str(h)), *(len(str(r[i])) for r in rows)) if rows else len(h) for i, h in enumerate(headers)]
    def fmt_row(cells: List[str]) -> str:
        return "  ".join(str(c).ljust(w) for c, w in zip(cells, widths))
    print(fmt_row(headers))
    print("  ".join("-" * w for w in widths))
    for r in rows:
        print(fmt_row(r))


def _preflight(frs: List[str]) -> bool:
    """Bütün tələb olunan dəftərlərin mövcudluğunu YOXLAYIR. Çatışan varsa cədvəl çap edir və
    False qaytarır (icra başlamamış dayandırılmalıdır)."""
    rows = []
    all_ok = True
    seen: set = set()
    frs = [fr for fr in frs if not (fr in seen or seen.add(fr))]   # FR01 DAG-da iki dəfədir
    for fr in frs:
        p = _notebook_path(fr)
        ok = p.exists()
        all_ok &= ok
        rows.append([fr, str(p.relative_to(BASE_DIR)), "VAR" if ok else "YOXDUR (mövcud deyil)"])
    print("\n=== Ön-yoxlama: dəftərlərin mövcudluğu ===")
    _print_table(rows, ["FR", "Dəftər yolu", "Status"])
    if not all_ok:
        missing = [fr for fr in frs if not _notebook_path(fr).exists()]
        print(
            f"\nDAYANDIRILDI: {len(missing)} dəftər hələ yaradılmayıb ({', '.join(missing)}). "
            "Boru YALNIZ mövcud dəftərləri icra edə bilər; əvvəlcə notebooks/ altında müvafiq "
            "FRxx.ipynb fayllarını yaradın, sonra yenidən işə salın."
        )
    return all_ok


def _ensure_cache() -> bool:
    """Mərhələ 0: `data/cache/` hazır olduğunu YOXLAYIR, olmadıqda və ya iş kitabından köhnə
    qaldıqda `data_layer.build_cache()` çağırır. Heç bir dəftər keşsiz işləyə bilmədiyi üçün bu
    addım DAG-dan əvvəl, preflight-dan da əvvəl icra olunur (metodoloji xartiya, D6)."""
    try:
        from src import data_layer as _dl  # type: ignore
    except Exception as e:
        print(f"XƏTA: src.data_layer idxal edilə bilmədi ({e}); keş yoxlanıla bilmir.")
        return False

    cache_csv = Path(_cfg.C_ACTUALS) if _cfg is not None else BASE_DIR / "data" / "cache" / "actuals_long.csv"
    xlsx_main = Path(_cfg.XLSX_MAIN) if _cfg is not None else BASE_DIR / "data" / "Statistik data dinamika 05.06.2026 +.xlsx"

    fresh = (
        cache_csv.exists()
        and cache_csv.stat().st_size > 0
        and (not xlsx_main.exists() or cache_csv.stat().st_mtime >= xlsx_main.stat().st_mtime)
    )
    print("\n=== Mərhələ 0: keş hazırlığı ===")
    if fresh:
        print(f"keş yenidir (mövcud): {cache_csv.relative_to(BASE_DIR)}")
        return True

    print("keş yoxdur və ya iş kitabından köhnədir — data_layer.build_cache() çağırılır ...")
    try:
        _dl.build_cache(verbose=True, check_md5=True)
    except Exception as e:
        print(f"XƏTA: keş qurulması uğursuz oldu: {e}")
        return False
    return True


# Kernel adı ARTIQ sabit deyil. Əvvəllər burada `python3` bərkidilmişdi; bu, dəftərin öz
# kernelspec metadatasını ƏVƏZ EDİR və host sistemində `python3` kerneli başqa (paketin
# asılılıqları quraşdırılmamış) interpretatora baxdıqda bütün boru xətti idxal mərhələsində
# dağılır — üstəlik uğursuz icra dəftərin çıxışlarını da silir. İndi ad dəftərin özündən
# oxunur; yoxdursa layihənin `.venv` kerneli işlədilir (MIIS_KERNEL ilə əvəzlənə bilər).
DEFAULT_KERNEL = os.environ.get("MIIS_KERNEL", "miis-model")


def _kernel_for(path: Path) -> str:
    try:
        meta = json.loads(path.read_text(encoding="utf-8")).get("metadata", {})
        return (meta.get("kernelspec", {}) or {}).get("name") or DEFAULT_KERNEL
    except Exception:
        return DEFAULT_KERNEL


def _run_one(fr: str) -> tuple[bool, str]:
    """Bir dəftəri nbconvert ilə yerində icra edir. (uğur?, qısa mesaj) qaytarır."""
    path = _notebook_path(fr)
    # `jupyter` ADI ilə deyil, boru xəttini işə salan interpretatorun ÖZ modulu ilə çağırılır:
    # PATH-da başqa (asılılıqsız) jupyter ola bilər və `jupyter-nbconvert not found` verə bilər.
    cmd = [
        sys.executable, "-m", "nbconvert", "--to", "notebook", "--execute", "--inplace",
        # Kernel AÇIQ verilir ki, nbconvert naməlum defolt kernelə sürüşməsin (əvvəlki davranışın
        # məqsədi bu idi) — lakin ad indi dəftərin öz metadatasından gəlir, bərkidilmir.
        f"--ExecutePreprocessor.kernel_name={_kernel_for(path)}",
        str(path),
    ]
    try:
        proc = subprocess.run(
            cmd, cwd=str(BASE_DIR), capture_output=True, text=True, timeout=NBCONVERT_TIMEOUT_SEC
        )
    except FileNotFoundError:
        return False, ("nbconvert bu interpretatorda yoxdur "
                       f"({sys.executable}) — `pip install -r requirements.txt` icra edin")
    except subprocess.TimeoutExpired:
        return False, f"vaxt aşımı (> {NBCONVERT_TIMEOUT_SEC}s)"
    if proc.returncode == 0:
        return True, "UĞURLU"
    tail = (proc.stderr or proc.stdout or "").strip().splitlines()
    last_line = tail[-1] if tail else "naməlum xəta"
    return False, f"XƏTA (kod={proc.returncode}): {last_line}"


# YEKUN ADDIMLAR — nəticələrin Excel (xlsx) və json formatında yadda saxlanılması.
# MİİS §15.5.1-in FR1–FR12 üzrə alt-tapşırıq 5-i hər iki formatı TƏLƏB EDİR, ona görə bu iki
# addım paketin müqavilə çıxışlarının bir hissəsidir və hər tam icradan sonra işləyir.
# Onlar DAG-ın 15 mərhələsinə DAXİL DEYİL (dəftər icra etmirlər — hazır çıxışları yenidən
# yığırlar), ona görə mərhələ cədvəlində və «BÜTÜN … MƏRHƏLƏ» sətrində sayılmırlar.
POST_BUILDERS: List[tuple[str, Path]] = [
    ("nəticələrin Excel iş kitabı (xlsx)", BASE_DIR / "build_results_xlsx.py"),
    ("nəticələrin JSON faylı", BASE_DIR / "build_results_json.py"),
]


def _check_assumption_registry() -> bool:
    """Fərziyyə dəstinin İDARƏETMƏ qatını yoxlayır: hər açar əvəzləmə reyestrində təsnif
    olunmalıdır (məzmun paketi, sənəd 01). Yeni açar əlavə edən mərhələ reyestri yeniləməsə,
    boru xətti burada dayanır — sənəd 01-in açar siyahısının səssizcə köhnəlməsi belə
    bağlanır (icmalın B21 bəndi)."""
    sys.path.insert(0, str(BASE_DIR))
    try:
        from src import outputs as O
        reg = O.validate_assumption_registry()
    except Exception as exc:                                  # noqa: BLE001 — mesaj istifadəçiyə çıxır
        print(f"XƏTA: əvəzləmə reyestri yoxlanışı uğursuz oldu:\n  {exc}")
        return False
    n = reg.groupby("override_class").size().to_dict()
    print(f"əvəzləmə reyestri: {len(reg)} açar təsnif olunub — "
          f"A (rəsmi giriş) {n.get('A', 0)}, B (modul nəticəsi) {n.get('B', 0)}, "
          f"C (törəmə/qapalı) {n.get('C', 0)}")
    return True


def _post_stage_outputs() -> bool:
    """Hər iki yazıcını ardıcıl işə salır. Biri uğursuz olarsa boru xətti UĞURSUZ sayılır —
    hər ikisi müqavilə çıxışıdır, ona görə səssiz atlama yolverilməzdir."""
    print("\n=== Yekun addım: fərziyyə açarlarının əvəzləmə reyestri ===")
    ok_all = _check_assumption_registry()
    for label, script in POST_BUILDERS:
        print(f"\n=== Yekun addım: {label} ===")
        if not script.exists():
            print(f"XƏTA: {script.name} tapılmadı — bu fayl paketin tərkib hissəsidir.")
            ok_all = False
            continue
        proc = subprocess.run([sys.executable, "-B", str(script)],
                              cwd=str(BASE_DIR), capture_output=True, text=True)
        print((proc.stdout or "").rstrip())
        if proc.returncode != 0:
            print(f"XƏTA: {label} yığıla bilmədi:\n" + (proc.stderr or "")[-1500:])
            ok_all = False
    return ok_all


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--stage", type=str, default=None, help="Yalnız bir FR-i işə sal, məsələn --stage FR09")
    args = ap.parse_args()

    if args.stage:
        try:
            target = _normalise_stage(args.stage)
        except ValueError as e:
            print(f"XƏTA: {e}")
            return 1
        if target not in ALL_FRS:
            print(f"XƏTA: {target} DAG-da tanınmır (mümkün: {', '.join(ALL_FRS)})")
            return 1
        run_stages = [[target]]
    else:
        run_stages = DAG

    if not _ensure_cache():
        print("\nDAYANDIRILDI: keş hazır deyil, heç bir dəftər etibarlı işləyə bilməz.")
        return 1

    scoped_frs = [fr for stage in run_stages for fr in stage]
    if not _preflight(scoped_frs):
        return 1

    print("\n=== İcra (DAG sırası ilə) ===")
    summary_rows: List[List[str]] = []
    aborted = False
    for stage in run_stages:
        for fr in stage:
            print(f"-- {fr} icra olunur ...")
            ok, msg = _run_one(fr)
            summary_rows.append([fr, "UĞURLU" if ok else "UĞURSUZ", msg])
            if not ok:
                aborted = True
                break
        if aborted:
            break

    print("\n=== Nəticə cədvəli (mərhələ üzrə uğur/uğursuzluq) ===")
    _print_table(summary_rows, ["FR", "Nəticə", "Qeyd"])

    if aborted:
        print(
            "\nBORU DAYANDIRILDI: yuxarıda göstərilən mərhələ uğursuz oldu. Aşağı axın mərhələləri "
            "bu nəticədən asılı olduğu üçün (D2 DAG) davam edilmədi."
        )
        return 2

    print(f"\nBÜTÜN {len(scoped_frs)} MƏRHƏLƏ UĞURLA TAMAMLANDI.")

    # Yalnız TAM icradan sonra: tək mərhələ (`--stage`) işlədildikdə iş kitabı və JSON
    # qismən yenilənmiş çıxışlardan yığılardı, ona görə o halda addımlar keçilir.
    if not args.stage and not _post_stage_outputs():
        return 3
    return 0


if __name__ == "__main__":
    sys.exit(main())
