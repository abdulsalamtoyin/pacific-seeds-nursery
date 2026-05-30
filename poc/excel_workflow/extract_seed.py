"""Extract a vbaProject.bin from a known-good .xlsm and stash it as a seed
so future build_workbooks.py runs can bake the same VBA into fresh files.

Usage:
    python extract_seed.py output/Nursery_Template.xlsm
    python extract_seed.py output/Nursery_Hub.xlsm

The seed is saved into excel_workflow/seeds/<stem>.vbaProject.bin.
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SEEDS = HERE / "seeds"


def extract(xlsm_path: Path) -> Path:
    if not xlsm_path.exists():
        raise FileNotFoundError(xlsm_path)
    if xlsm_path.suffix != ".xlsm":
        raise ValueError(f"Expected .xlsm, got {xlsm_path.suffix}")

    SEEDS.mkdir(exist_ok=True)
    seed_path = SEEDS / f"{xlsm_path.stem}.vbaProject.bin"

    with zipfile.ZipFile(xlsm_path, "r") as zf:
        names = zf.namelist()
        if "xl/vbaProject.bin" not in names:
            raise RuntimeError(f"No vbaProject.bin found in {xlsm_path}")
        seed_path.write_bytes(zf.read("xl/vbaProject.bin"))

    return seed_path


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    for arg in sys.argv[1:]:
        try:
            out = extract(Path(arg))
            print(f"  ✓ {arg} → {out} ({out.stat().st_size:,} bytes)")
        except Exception as e:
            print(f"  ✗ {arg}: {e}")
            return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
