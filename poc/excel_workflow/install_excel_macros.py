"""Auto-install the VBA modules into Nursery_Template.xlsx and Nursery_Hub.xlsx,
then save each as .xlsm. Requires macOS Excel + xlwings.

If Excel's 'Trust access to the VBA project object model' is NOT enabled, this
script will print clear remediation steps and exit cleanly.
"""
from __future__ import annotations

import re
import sys
import site
import time
from pathlib import Path

# appscript (xlwings dep on macOS) ships in an "aeosa" subdir activated by a .pth
# file. Anaconda-based interpreters sometimes skip .pth processing inside venvs.
# Add the aeosa path manually so `import appscript` works regardless.
for _sp in site.getsitepackages() + [p for p in sys.path if "site-packages" in p]:
    _aeosa = Path(_sp) / "aeosa"
    if _aeosa.exists() and str(_aeosa) not in sys.path:
        sys.path.insert(0, str(_aeosa))

HERE = Path(__file__).resolve().parent
VBA_DIR = HERE / "vba"
OUT_DIR = HERE / "output"

XLSX_TEMPLATE = OUT_DIR / "Nursery_Template.xlsx"
XLSM_TEMPLATE = OUT_DIR / "Nursery_Template.xlsm"
XLSX_HUB      = OUT_DIR / "Nursery_Hub.xlsx"
XLSM_HUB      = OUT_DIR / "Nursery_Hub.xlsm"

# VBE constants
vbext_ct_StdModule = 1
xlOpenXMLWorkbookMacroEnabled = 52


def _green(s): return f"\033[1;32m{s}\033[0m"
def _red(s):   return f"\033[1;31m{s}\033[0m"
def _cyan(s):  return f"\033[1;36m{s}\033[0m"
def _dim(s):   return f"\033[2m{s}\033[0m"


def strip_cls_header(text: str) -> str:
    """Strip the VERSION/Attribute boilerplate from a .cls file so the body
    can be inserted into an existing class module via AddFromString."""
    lines = text.splitlines()
    out = []
    skipping = True
    for ln in lines:
        if skipping:
            if ln.startswith("VERSION ") or ln == "BEGIN" or ln == "END" \
               or ln.startswith("  MultiUse") or ln.startswith("Attribute "):
                continue
            # Allow empty lines and comments while we're still in the header.
            if ln.strip() == "":
                continue
            skipping = False
        out.append(ln)
    return "\n".join(out).strip() + "\n"


def _offer_open_excel_prefs() -> None:
    """Best-effort: open Excel's Preferences window so the user can find the
    setting in two clicks. Falls back to printing instructions."""
    import subprocess
    print(_cyan("Opening Excel Preferences for you…"))
    apple = '''tell application "Microsoft Excel"
    activate
    try
        tell application "System Events" to keystroke "," using {command down}
    end try
end tell'''
    try:
        subprocess.run(["osascript", "-e", apple], timeout=5, check=False)
    except Exception:
        pass


def trust_check_message() -> str:
    return _red("\n⚠  Could not write VBA into the workbook.\n") + (
        "\nOn Mac, TWO separate permissions are needed:\n\n"
        "  " + _cyan("1) Excel's VBA-trust setting") + "\n"
        "     Excel → Preferences → Security & Privacy → Macro Security\n"
        "     → tick \"Trust access to the VBA project object model\"\n\n"
        "  " + _cyan("2) macOS Automation permission") + "\n"
        "     System Settings → Privacy & Security → Automation\n"
        "     → expand \"Terminal\" (or whichever app you ran this from)\n"
        "     → tick \"Microsoft Excel\"\n"
        "     (If Microsoft Excel doesn't appear yet, run the installer once\n"
        "     and accept the popup — macOS will then list it.)\n\n"
        "After enabling both, fully quit Excel (⌘Q) and re-run install.command.\n"
    )


def open_excel():
    """Open xlwings and return the running Excel instance, launching Excel if needed."""
    import xlwings as xw
    apps = xw.apps
    if len(apps) == 0:
        print(_dim("  Launching Excel…"))
        app = xw.App(visible=True, add_book=False)
        time.sleep(1.5)
    else:
        app = apps.active or apps[0]
        try:
            app.visible = True
        except Exception:
            pass
    return app


def install_workbook(app, xlsx_path: Path, xlsm_path: Path,
                     std_module_basfile: Path,
                     cls_files: list[tuple[str, Path]],
                     label: str) -> bool:
    """Inject VBA into one workbook and save as .xlsm.

    cls_files: list of (component_name, cls_path) to overwrite. The
    component_name is what VBE shows (e.g. "ThisWorkbook" or the sheet's
    codename like "Sheet1").
    """
    if not xlsx_path.exists():
        print(_red(f"  ✗ Source not found: {xlsx_path}"))
        return False

    print(_cyan(f"\n→ {label}"))
    # Close the file if it's already open in Excel from a prior run.
    for b in list(app.books):
        try:
            if Path(b.fullname).resolve() == xlsx_path.resolve() or \
               Path(b.fullname).name in (xlsx_path.name, xlsm_path.name):
                b.close()
        except Exception:
            pass
    print(_dim(f"  Opening {xlsx_path.name}…"))
    book = app.books.open(str(xlsx_path))

    # Probe the VBProject; if access is blocked we close cleanly and bail.
    blocked = False
    underlying_err = None
    vb_project = None
    try:
        vb_project = book.api.VBProject
        # Accessing .VBComponents triggers the security check on Mac.
        _ = vb_project.VBComponents.Count
    except Exception as e:
        blocked = True
        underlying_err = e

    if blocked:
        try: book.close(save_changes=False)
        except TypeError:
            try: book.close()
            except Exception: pass
        except Exception: pass
        # Show the actual underlying error so we can diagnose which permission
        # layer (Mac Automation vs Excel VBA-trust) is the real problem.
        print(_red("  ✗ Could not reach VBProject."))
        if underlying_err:
            print(_dim("  Underlying error:"))
            print(_dim(f"    {type(underlying_err).__name__}: {underlying_err}"))
        print(trust_check_message())
        return False

    # 1. Add the main .bas as a Standard Module
    print(_dim(f"  Adding module {std_module_basfile.name}…"))
    try:
        new_mod = vb_project.VBComponents.Add(vbext_ct_StdModule)
        bas_text = std_module_basfile.read_text(encoding="utf-8")
        # AddFromString handles the Attribute VB_Name header
        new_mod.CodeModule.AddFromString(bas_text)
    except Exception as e:
        print(_red(f"  ✗ Failed to add std module: {e}"))
        return False

    # 2. Replace each class module's code
    for comp_name, cls_path in cls_files:
        print(_dim(f"  Updating {comp_name} from {cls_path.name}…"))
        try:
            comp = _find_component(vb_project, comp_name)
            if comp is None:
                print(_red(f"  ✗ Component '{comp_name}' not found in {xlsx_path.name}"))
                continue
            code = strip_cls_header(cls_path.read_text(encoding="utf-8"))
            cm = comp.CodeModule
            line_count = cm.CountOfLines
            if line_count > 0:
                cm.DeleteLines(1, line_count)
            cm.AddFromString(code)
        except Exception as e:
            print(_red(f"  ✗ Failed to update {comp_name}: {e}"))

    # 3. Save as .xlsm
    print(_dim(f"  Saving as {xlsm_path.name}…"))
    try:
        if xlsm_path.exists():
            xlsm_path.unlink()
        book.api.SaveAs(str(xlsm_path), xlOpenXMLWorkbookMacroEnabled)
    except Exception as e:
        print(_red(f"  ✗ SaveAs failed: {e}"))
        return False

    try: book.close()
    except Exception: pass

    print(_green(f"  ✓ Wrote {xlsm_path}"))
    return True


def _find_component(vb_project, name: str):
    """Find a VBComponent by its Name (case-insensitive)."""
    target = name.lower()
    for c in vb_project.VBComponents:
        try:
            if c.Name.lower() == target:
                return c
        except Exception:
            continue
    # Some xlwings/appscript paths require indexed access:
    try:
        n = vb_project.VBComponents.Count
        for i in range(1, n + 1):
            c = vb_project.VBComponents.Item(i)
            if c.Name.lower() == target:
                return c
    except Exception:
        pass
    return None


def find_home_sheet_codename(app, xlsx_path: Path) -> str | None:
    """Try to find the VBE codename of the 'Home' sheet.

    Mac Excel's AppleScript bridge doesn't reliably expose `CodeName`, so we
    fall back to assuming the Home sheet is the first sheet (which our build
    script guarantees). In a freshly-generated workbook, codename = 'Sheet1'.
    """
    # Best effort via appscript:
    try:
        book = app.books.open(str(xlsx_path))
        try:
            for ws in book.sheets:
                if ws.name == "Home":
                    try:
                        return ws.api.code_name
                    except Exception:
                        pass
        finally:
            try: book.close()
            except Exception: pass
    except Exception:
        pass
    return "Sheet1"


def main() -> int:
    print(_cyan("Pacific Seeds — Nursery VBA installer"))
    print(_dim("Source: ") + str(VBA_DIR))
    print(_dim("Output: ") + str(OUT_DIR))

    if not VBA_DIR.exists():
        print(_red(f"VBA dir missing: {VBA_DIR}"))
        return 2

    try:
        app = open_excel()
    except Exception as e:
        print(_red(f"Could not start Excel via xlwings: {e}"))
        return 3

    # We need to know each workbook's "Home" sheet codename (Sheet1 etc.).
    template_home_code = find_home_sheet_codename(app, XLSX_TEMPLATE) or "Sheet1"
    hub_home_code      = find_home_sheet_codename(app, XLSX_HUB)      or "Sheet1"
    print(_dim(f"  Template Home codename = {template_home_code}"))
    print(_dim(f"  Hub      Home codename = {hub_home_code}"))

    ok_t = install_workbook(
        app,
        XLSX_TEMPLATE, XLSM_TEMPLATE,
        std_module_basfile=VBA_DIR / "NurseryTemplate.bas",
        cls_files=[
            ("ThisWorkbook", VBA_DIR / "ThisWorkbook_Template.cls"),
            (template_home_code, VBA_DIR / "Sheet_Home_Template.cls"),
        ],
        label="Nursery_Template",
    )

    # If the first one was blocked by VBA trust, no point trying the second.
    if not ok_t:
        print()
        _offer_open_excel_prefs()
        return 1

    ok_h = install_workbook(
        app,
        XLSX_HUB, XLSM_HUB,
        std_module_basfile=VBA_DIR / "NurseryHub.bas",
        cls_files=[
            (hub_home_code, VBA_DIR / "Sheet_Home_Hub.cls"),
        ],
        label="Nursery_Hub",
    )

    print()
    if ok_t and ok_h:
        print(_green("✅ All done."))
        print(_dim("Files ready at:"))
        print(f"   {XLSM_TEMPLATE}")
        print(f"   {XLSM_HUB}")
        return 0
    print(_red("⚠ Finished with errors above."))
    return 1


if __name__ == "__main__":
    sys.exit(main())
