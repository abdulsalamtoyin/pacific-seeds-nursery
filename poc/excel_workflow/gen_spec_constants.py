"""Emit vba/SpecConstants.bas from nursery_spec.json.

This module is the only writer of SpecConstants.bas. Hand-editing that file
puts the VBA out of step with the Python builder, which is exactly what the
declarative spec exists to prevent.

Run: python -m excel_workflow.gen_spec_constants
"""
from __future__ import annotations

import re
from pathlib import Path

from excel_workflow.spec.loader import load_spec

OUT_PATH = Path(__file__).resolve().parent / "vba" / "SpecConstants.bas"

HEADER = '''Attribute VB_Name = "SpecConstants"
'==============================================================================
'  SpecConstants - GENERATED FILE, DO NOT EDIT
'
'  Produced by excel_workflow/gen_spec_constants.py from
'  excel_workflow/spec/nursery_spec.json. Edit the JSON and re-run the
'  generator; any manual change here is overwritten.
'==============================================================================
Option Explicit

'--- Sheet names --------------------------------------------------------------
'''


def const_name(tab: str) -> str:
    """'Replacements and Errors' -> 'SHEET_REPLACEMENTS_AND_ERRORS'."""
    slug = re.sub(r"[^A-Za-z0-9]+", "_", tab).strip("_").upper()
    return f"SHEET_{slug}"


def render_constants() -> str:
    spec = load_spec()
    lines = [HEADER]

    tabs = (list(spec["always_first"]) + list(spec["hidden"])
            + list(spec["default_tabs"]) + list(spec["conditional"])
            + list(spec["extras"]))
    width = max(len(const_name(t)) for t in tabs)
    for tab in tabs:
        lines.append(
            f'Public Const {const_name(tab):<{width}} As String = "{tab}"')

    lines.append("")
    lines.append(
        "'--- Fieldbook column positions -----------------------------------------------")
    columns = spec["fieldbook_columns"]
    col_names = [const_name(c).replace("SHEET_", "FB_COL_") for c in columns]
    col_width = max(len(n) for n in col_names)
    for index, name in enumerate(col_names, start=1):
        lines.append(f"Public Const {name:<{col_width}} As Long = {index}")

    lines.append("")
    return "\n".join(lines)


def main() -> None:
    OUT_PATH.write_text(render_constants(), encoding="utf-8")
    print(f"wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
