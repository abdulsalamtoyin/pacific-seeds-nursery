"""Bundle the Windows app and the Excel workbook into one zip for users.

The workbook is only included if its baked VBA is current. The .xlsx half is
regenerated from Python on every build, but the macros only reach the file
through the manual bootstrap step in Excel — so a workbook can look freshly
built while carrying months-old macros that reference tabs which no longer
exist. Shipping that is worse than shipping nothing, because every button
fails on click and the user has no way to tell why.

Usage:
    python scripts/make_release.py
    python scripts/make_release.py --installer ~/Downloads/PacificSeedsSetup-1.0.0.exe
    python scripts/make_release.py --allow-stale-vba    # not for users
"""
from __future__ import annotations

import argparse
import hashlib
import sys
import zipfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = ROOT / "excel_workflow" / "output" / "Nursery_Template.xlsm"
HUB = ROOT / "excel_workflow" / "output" / "Nursery_Hub.xlsm"
SEED = ROOT / "excel_workflow" / "seeds" / "Nursery_Template.vbaProject.bin"
QUICKREF = ROOT / "excel_workflow" / "dist" / "QuickReference.pdf"
OUT_DIR = ROOT / "dist" / "release"

DEFAULT_INSTALLER_DIRS = [
    Path.home() / "Downloads" / "pacific-seeds-build",
    Path.home() / "Downloads",
    ROOT / "dist" / "installer",
]

# Procedures added in the v2 rebuild. If the baked seed has none of these,
# the workbook's macros predate the twelve-tab restructure.
V2_MARKERS = [
    "btnFieldMapWizard",
    "btnBuildFieldbook",
    "btnResolveSplitFromQR",
    "SpecConstants",
]


def seed_has_v2_macros() -> tuple[bool, list[str]]:
    """Look for the v2 procedure names inside the compiled VBA blob.

    Crude — the blob is compressed, so absence is strong evidence while a
    single hit is enough to call the seed current.
    """
    if not SEED.exists():
        return False, list(V2_MARKERS)
    blob = SEED.read_bytes()
    missing = [m for m in V2_MARKERS
               if m.encode("utf-16-le") not in blob and m.encode() not in blob]
    return len(missing) < len(V2_MARKERS), missing


def find_installer(explicit: str | None) -> Path | None:
    if explicit:
        p = Path(explicit).expanduser()
        return p if p.exists() else None
    for folder in DEFAULT_INSTALLER_DIRS:
        if not folder.is_dir():
            continue
        hits = sorted(folder.glob("PacificSeedsSetup*.exe"))
        if hits:
            return hits[-1]
    return None


README = """Pacific Seeds — Sorghum Nursery Workflow
========================================
Release: {today}

WHAT'S IN HERE
--------------
{contents}

INSTALLING THE APP (Windows)
----------------------------
1. Double-click PacificSeedsSetup-*.exe
2. Windows may warn "Windows protected your PC" because the installer is
   not code-signed. Click "More info" then "Run anyway".
3. It installs for the current user only — no admin rights needed.
4. Launch it from the Start Menu. A tray icon appears and your browser
   opens on the nursery workbook.

To remove it: Settings > Apps > Pacific Seeds Nursery Fieldbook.

USING THE EXCEL WORKBOOK
------------------------
1. Open Nursery_Template.xlsm.
2. Click "Enable Content" when Excel asks — the workflow is macro-driven
   and nothing works without it.
3. The Home tab lists every step in order. Start at step 1.

The app and the workbook do the same job. Use whichever suits you; they
follow the same steps, in the same order, with the same rules.

SUPPORT
-------
If the app will not start, it writes a log to:
    %LOCALAPPDATA%\\PacificSeeds
Send that file on and it will say what went wrong.
"""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--installer", help="path to PacificSeedsSetup-*.exe")
    ap.add_argument("--allow-stale-vba", action="store_true",
                    help="include the workbook even if its macros are old "
                         "(for internal testing, never for users)")
    args = ap.parse_args()

    fresh, missing = seed_has_v2_macros()
    installer = find_installer(args.installer)

    items: list[tuple[Path, str]] = []
    notes: list[str] = []

    if installer:
        items.append((installer, installer.name))
        notes.append(f"{installer.name}   the Windows app installer")
    else:
        print("!  No PacificSeedsSetup-*.exe found. Pass --installer PATH.",
              file=sys.stderr)

    if fresh or args.allow_stale_vba:
        for wb in (WORKBOOK, HUB):
            if wb.exists():
                items.append((wb, wb.name))
                notes.append(f"{wb.name}   the Excel workbook")
        if not fresh:
            print("!  Including a workbook with STALE macros "
                  "(--allow-stale-vba).", file=sys.stderr)
    else:
        print("!  EXCLUDING the Excel workbook — its baked VBA is out of date.",
              file=sys.stderr)
        print(f"   Missing from the seed: {', '.join(missing)}", file=sys.stderr)
        print("   The .xlsx structure is current but the macros are not, so the",
              file=sys.stderr)
        print("   buttons would reference tabs that no longer exist.",
              file=sys.stderr)
        print("   Fix: run the bootstrap step in Excel (excel_workflow/SETUP.md,",
              file=sys.stderr)
        print("   'Changing VBA — the full loop'), then re-run this script.",
              file=sys.stderr)

    if QUICKREF.exists():
        items.append((QUICKREF, QUICKREF.name))
        notes.append(f"{QUICKREF.name}   printable one-page reference")

    if not items:
        print("Nothing to package.", file=sys.stderr)
        return 1

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    today = date.today().isoformat()
    zip_path = OUT_DIR / f"PacificSeeds-Nursery-{today}.zip"
    if zip_path.exists():
        zip_path.unlink()

    readme = README.format(today=today, contents="\n".join(notes))

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for src, arcname in items:
            zf.write(src, arcname)
        zf.writestr("README.txt", readme)

    digest = hashlib.sha256(zip_path.read_bytes()).hexdigest()
    size_mb = zip_path.stat().st_size / 1048576

    print(f"\nWrote {zip_path}")
    print(f"  {size_mb:.1f} MB")
    print(f"  sha256 {digest}")
    print("\nContents:")
    for _, arcname in items:
        print(f"  - {arcname}")
    print("  - README.txt")
    if not fresh and not args.allow_stale_vba:
        print("\nNOTE: the Excel workbook is NOT in this zip. See above.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
