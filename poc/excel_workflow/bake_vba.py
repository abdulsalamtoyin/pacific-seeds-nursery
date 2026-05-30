"""Bake a pre-built vbaProject.bin into a freshly-generated .xlsx and rewrite it
as a fully-working .xlsm. Bypasses the bootstrap dance entirely.

The .xlsx → .xlsm transformation needs three steps:
  1. Inject xl/vbaProject.bin into the package.
  2. Change the workbook content type in [Content_Types].xml from .sheet to
     .sheet.macroEnabled.
  3. Add a relationship from the workbook to vbaProject.bin.

Companion: extract_seed.py pulls a vbaProject.bin out of a known-good .xlsm
so it can be locked in as a seed for future builds.
"""
from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

SHEET_TYPE_PLAIN = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
)
SHEET_TYPE_MACRO = (
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
)
VBA_REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"


def bake_vba(xlsx_src: Path, vba_bin: Path, xlsm_dst: Path) -> None:
    """Combine an .xlsx with a vbaProject.bin into a .xlsm at xlsm_dst."""
    if not xlsx_src.exists():
        raise FileNotFoundError(xlsx_src)
    if not vba_bin.exists():
        raise FileNotFoundError(vba_bin)

    if xlsm_dst.exists():
        xlsm_dst.unlink()

    # zipfile can't update entries in place; read everything, mutate, rewrite.
    with zipfile.ZipFile(xlsx_src, "r") as zf:
        entries: dict[str, bytes] = {n: zf.read(n) for n in zf.namelist()}

    # 1. Add the vbaProject.bin
    entries["xl/vbaProject.bin"] = vba_bin.read_bytes()

    # 2. Update Content Types — switch workbook to macroEnabled + register .bin
    ct = entries["[Content_Types].xml"].decode("utf-8")
    if SHEET_TYPE_PLAIN in ct:
        ct = ct.replace(SHEET_TYPE_PLAIN, SHEET_TYPE_MACRO)
    if "vnd.ms-office.vbaProject" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Default Extension="bin" '
            'ContentType="application/vnd.ms-office.vbaProject"/></Types>',
        )
    entries["[Content_Types].xml"] = ct.encode("utf-8")

    # 3. Add relationship from workbook to vbaProject.bin
    rels_path = "xl/_rels/workbook.xml.rels"
    if rels_path in entries:
        rels = entries[rels_path].decode("utf-8")
        if "vbaProject" not in rels:
            # Pick a relationship id that doesn't collide with existing ones.
            new_rel = (
                f'<Relationship Id="rIdVBA" Type="{VBA_REL_TYPE}" '
                'Target="vbaProject.bin"/>'
            )
            rels = rels.replace("</Relationships>", new_rel + "</Relationships>")
            entries[rels_path] = rels.encode("utf-8")

    # 4a. Inject codeName="ThisWorkbook" on the workbook's <workbookPr> so the
    #     ThisWorkbook.cls class module from the seed links to the actual
    #     workbook. Without this, Excel creates a duplicate ThisWorkbook1
    #     object and the `ThisWorkbook` keyword inside VBA refers to the
    #     wrong one — breaking every sheet lookup.
    wb_xml = entries["xl/workbook.xml"].decode("utf-8")
    if 'codeName="ThisWorkbook"' not in wb_xml:
        if "<workbookPr/>" in wb_xml:
            wb_xml = wb_xml.replace(
                "<workbookPr/>", '<workbookPr codeName="ThisWorkbook"/>')
        elif "<workbookPr>" in wb_xml:
            wb_xml = wb_xml.replace(
                "<workbookPr>", '<workbookPr codeName="ThisWorkbook">')
        elif "<workbookPr " in wb_xml:
            # workbookPr already has other attributes — insert codeName in the
            # first attribute slot
            wb_xml = wb_xml.replace(
                "<workbookPr ", '<workbookPr codeName="ThisWorkbook" ', 1)
        else:
            # No workbookPr element at all — insert one after <workbook ...>
            i = wb_xml.find("<workbook ")
            if i >= 0:
                close = wb_xml.find(">", i)
                wb_xml = wb_xml[:close + 1] + \
                    '<workbookPr codeName="ThisWorkbook"/>' + wb_xml[close + 1:]
        entries["xl/workbook.xml"] = wb_xml.encode("utf-8")

    # 4b. Inject codeName="SheetN" on each worksheet so the VBA class modules
    #    (Sheet1.cls, Sheet2.cls, …) get linked to the right sheet objects.
    #    Without this, Worksheet_BeforeDoubleClick etc. never fire — the
    #    handlers exist in the project but Excel can't find their host sheet.
    sheet_files = sorted(
        (n for n in entries if n.startswith("xl/worksheets/sheet")
         and n.endswith(".xml")),
        key=lambda n: int("".join(c for c in n.rsplit("/", 1)[-1] if c.isdigit())),
    )
    for idx, sheet_file in enumerate(sheet_files, start=1):
        xml = entries[sheet_file].decode("utf-8")
        code_name = f"Sheet{idx}"
        if 'codeName="' in xml:
            continue  # already set
        if "<sheetPr/>" in xml:
            xml = xml.replace(
                "<sheetPr/>", f'<sheetPr codeName="{code_name}"/>')
        elif "<sheetPr>" in xml:
            xml = xml.replace(
                "<sheetPr>", f'<sheetPr codeName="{code_name}">')
        else:
            # Insert a fresh sheetPr right after the opening <worksheet ...>
            i = xml.find("<dimension")
            if i < 0:
                i = xml.find("<sheetViews")
            if i > 0:
                xml = xml[:i] + f'<sheetPr codeName="{code_name}"/>' + xml[i:]
        entries[sheet_file] = xml.encode("utf-8")

    # Rewrite the zip as .xlsm
    with zipfile.ZipFile(xlsm_dst, "w", zipfile.ZIP_DEFLATED) as zf_out:
        for name, content in entries.items():
            zf_out.writestr(name, content)
