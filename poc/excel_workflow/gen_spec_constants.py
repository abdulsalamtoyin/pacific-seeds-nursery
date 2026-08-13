"""Emit vba/SpecConstants.bas from nursery_spec.json.

This module is the only writer of SpecConstants.bas. Hand-editing that file
puts the VBA out of step with the Python builder, which is exactly what the
declarative spec exists to prevent.

Run: python -m excel_workflow.gen_spec_constants
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from excel_workflow.spec.loader import load_spec

OUT_PATH = Path(__file__).resolve().parent / "vba" / "SpecConstants.bas"
JS_OUT_PATH = (Path(__file__).resolve().parent.parent
               / "pwa" / "nursery-spec.js")

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


def render_js() -> str:
    """The same spec as an ES module, for the desktop app.

    The app has to open from the filesystem as well as over HTTP, so the spec
    ships as a module rather than something fetched at runtime.
    """
    spec = load_spec()
    body = json.dumps(spec, indent=2, ensure_ascii=False)
    return (
        "// GENERATED FILE, DO NOT EDIT\n"
        "//\n"
        "// Produced by excel_workflow/gen_spec_constants.py from\n"
        "// excel_workflow/spec/nursery_spec.json. Edit the JSON and re-run\n"
        "// the generator; any manual change here is overwritten.\n"
        "//\n"
        "// This is the same spec the Excel builder and SpecConstants.bas read,\n"
        "// so the desktop app's tabs cannot drift from the workbook's.\n"
        "\n"
        f"export const SPEC = {body};\n"
        "\n"
        "export function defaultTabs() {\n"
        "  return [...SPEC.default_tabs];\n"
        "}\n"
        "\n"
        "// Mirrors tabs_for() in excel_workflow/spec/loader.py.\n"
        "export function tabsFor(nurseryTypes, plantingDates) {\n"
        "  if (plantingDates < 1) {\n"
        "    throw new RangeError(\n"
        "      `plantingDates must be >= 1, got ${plantingDates}`);\n"
        "  }\n"
        "  const selected = new Set(nurseryTypes);\n"
        "  const hits = (types) => types.some((t) => selected.has(t));\n"
        "\n"
        "  const tabs = [...SPEC.always_first, ...SPEC.hidden];\n"
        "  for (const name of SPEC.default_tabs) {\n"
        "    for (const [condName, rule] of Object.entries(SPEC.conditional)) {\n"
        "      if (rule.before === name && hits(rule.types)) tabs.push(condName);\n"
        "    }\n"
        "    if (name in SPEC.fan_out) {\n"
        "      for (let i = 1; i <= plantingDates; i++) tabs.push(`${name} ${i}`);\n"
        "    } else {\n"
        "      tabs.push(name);\n"
        "    }\n"
        "  }\n"
        "  for (const [tab, types] of Object.entries(SPEC.extras)) {\n"
        "    if (hits(types)) tabs.push(tab);\n"
        "  }\n"
        "  return tabs;\n"
        "}\n"
    )


def main() -> None:
    OUT_PATH.write_text(render_constants(), encoding="utf-8")
    print(f"wrote {OUT_PATH}")
    JS_OUT_PATH.write_text(render_js(), encoding="utf-8")
    print(f"wrote {JS_OUT_PATH}")


if __name__ == "__main__":
    main()
