"""Generate Nursery_Template.xlsx and Nursery_Hub.xlsx for the proposed workflow.

These workbooks are .xlsx (no embedded VBA) and ship with styled "button" cells
on the Home tab. After import (see SETUP.md), the included VBA wires up a
double-click router so each button cell runs its corresponding macro.

Pacific Seeds brand colours throughout (white + blue, navy headers).
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Protection, Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# Run either as `python excel_workflow/build_workbooks.py` (script mode, which
# puts excel_workflow/ on sys.path rather than poc/) or as
# `python -m excel_workflow.build_workbooks`. The guard makes the package
# import below resolve in both cases.
if __package__ in (None, ""):
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from excel_workflow.spec.loader import load_spec, tabs_for

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"
SEEDS = ROOT / "seeds"

# ---------- Pacific Seeds palette ----------
PS_NAVY     = "092A40"
PS_BLUE     = "0678CD"
PS_DEEP     = "0075BD"
PS_SKY      = "3AB3E5"
PS_LIGHT    = "8FCEEE"
PS_PALE     = "E9F3FC"
PS_WHITE    = "FFFFFF"
PS_GREY     = "F3F7FB"
PS_BORDER   = "D8E3ED"
PS_ORANGE   = "E45138"
PS_WHEAT    = "DDB318"
PS_GREEN    = "28A745"
PS_INK      = "092A40"
PS_INK_SOFT = "34526E"
PS_MUTED    = "5A7896"

THIN  = Side(style="thin",  color=PS_BORDER)
MED   = Side(style="medium", color=PS_BLUE)
HAIR  = Side(style="hair",  color=PS_BORDER)


# ---------- Style helpers ----------
def banner(ws, title: str, subtitle: str = "", cols: int = 12) -> int:
    """Pacific Seeds branded banner; returns the next free row."""
    end = get_column_letter(cols)
    ws.merge_cells(f"A1:{end}1")
    c = ws["A1"]
    c.value = "PACIFIC SEEDS"
    c.font = Font(name="Calibri", size=10, bold=True, color=PS_SKY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.fill = PatternFill("solid", fgColor=PS_NAVY)
    ws.row_dimensions[1].height = 18

    ws.merge_cells(f"A2:{end}2")
    c = ws["A2"]
    c.value = title
    c.font = Font(name="Calibri", size=18, bold=True, color=PS_WHITE)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.fill = PatternFill("solid", fgColor=PS_NAVY)
    ws.row_dimensions[2].height = 30

    if subtitle:
        ws.merge_cells(f"A3:{end}3")
        c = ws["A3"]
        c.value = subtitle
        c.font = Font(size=11, italic=True, color=PS_INK_SOFT)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c.fill = PatternFill("solid", fgColor=PS_PALE)
        ws.row_dimensions[3].height = 22
        next_row = 5
    else:
        # Thin sky underline
        ws.merge_cells(f"A3:{end}3")
        c = ws["A3"]
        c.fill = PatternFill("solid", fgColor=PS_SKY)
        ws.row_dimensions[3].height = 4
        next_row = 5
    return next_row


def header_row(ws, row: int, headers: list[str], start_col: int = 1) -> None:
    for i, h in enumerate(headers, start=start_col):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(bold=True, color=PS_NAVY, size=11)
        c.fill = PatternFill("solid", fgColor=PS_PALE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(top=THIN, bottom=MED, left=THIN, right=THIN)
    ws.row_dimensions[row].height = 28


def button_cell(ws, cell_range: str, label: str, macro: str,
                color: str = PS_BLUE, text_color: str = PS_WHITE,
                description: str = "") -> None:
    """Style a merged cell to look like a button. Macro name stored in a comment
    AND in a parallel hidden cell for the VBA router."""
    ws.merge_cells(cell_range)
    top_left = cell_range.split(":")[0]
    c = ws[top_left]
    c.value = label
    c.font = Font(bold=True, size=13, color=text_color)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.fill = PatternFill("solid", fgColor=color)
    thick = Side(style="medium", color=color)
    c.border = Border(top=thick, bottom=thick, left=thick, right=thick)
    if description:
        # Tooltip via comment
        from openpyxl.comments import Comment
        c.comment = Comment(f"{description}\n\n(Double-click to run: {macro})", "PS Workflow")


# ---------- Shared tab layouts ----------
REPLACEMENT_HEADERS = [
    "Stage", "Timestamp", "Plot", "Material ID", "Source ID",
    "QR", "Reason", "Technician", "Split no.", "Status",
]

STATUS_DEFAULT = "Waiting decision"
REPLACEMENT_LAST_ROW = 500

GROWTH_STAGES = [
    ("Seedling",      "FEC000"),
    ("Vegetative",    "A9D08E"),
    ("Heading",       "FF6B6B"),
    ("Flowering",     "ADD8E6"),
    ("Grain filling", "FFFF99"),
    ("Harvest",       "CCFFCC"),
    ("Post Harvest",  "CCCCFF"),
]


def decorate_replacements_sheet(ws, header_row_no: int) -> None:
    """Dropdowns, QR width and default status for the merged log.

    Stage is what distinguishes a packeting replacement from a planting error,
    so one row format now covers what used to be two tabs.
    """
    first = header_row_no + 1
    qr_col = get_column_letter(REPLACEMENT_HEADERS.index("QR") + 1)
    status_col = get_column_letter(REPLACEMENT_HEADERS.index("Status") + 1)
    stage_col = get_column_letter(REPLACEMENT_HEADERS.index("Stage") + 1)

    # The QR payload is a full comma-separated record, not a short code.
    ws.column_dimensions[qr_col].width = 150

    stage_dv = DataValidation(
        type="list", formula1='"Packeting,Planting"', allow_blank=True)
    status_dv = DataValidation(
        type="list", formula1=f'"{STATUS_DEFAULT},Changes made in PRISM"',
        allow_blank=True)
    ws.add_data_validation(stage_dv)
    ws.add_data_validation(status_dv)

    stage_dv.add(f"{stage_col}{first}:{stage_col}{REPLACEMENT_LAST_ROW}")
    status_dv.add(f"{status_col}{first}:{status_col}{REPLACEMENT_LAST_ROW}")

    for row in range(first, REPLACEMENT_LAST_ROW + 1):
        ws[f"{status_col}{row}"] = STATUS_DEFAULT


def build_nursery_data_sheet(ws) -> None:
    """The printable front page. Labels only — values are filled in by hand.

    Layout mirrors the 'Nursery data' tab of the client's File for Toyin.xlsx,
    including the deliberately blank rows 3-4 and right-hand column.
    """
    spec = load_spec()

    ws.merge_cells("A1:D1")
    ws["A1"] = "R&D Fieldbook - Grain Sorghum"
    ws["A1"].font = Font(bold=True, size=20)
    ws["A1"].alignment = Alignment(horizontal="center")

    # Rows 3 and 4 are merged but deliberately empty — the nursery name and
    # block order are written in by the breeder at print time.
    for row in (3, 4):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
    ws.cell(row=3, column=2).font = Font(bold=True, size=16)
    ws.cell(row=4, column=2).font = Font(size=12)
    ws.cell(row=4, column=2).border = Border(bottom=THIN)

    for offset, label in enumerate(spec["nursery_data_labels"]):
        row = 5 + offset
        cell = ws.cell(row=row, column=2, value=label)
        cell.font = Font(size=12)
        cell.border = Border(bottom=THIN)
        ws.cell(row=row, column=3).border = Border(bottom=THIN)

    ws.merge_cells("A24:D24")

    for col, width in spec["nursery_data_widths"].items():
        ws.column_dimensions[col].width = width


def build_grouped_sheet(ws, title: str) -> None:
    """Operations and Comments share one bordered Stage x Group table.

    The client asked to keep the existing structure and only rename the three
    free-text columns, so the growth-stage rows stay.
    """
    ws.merge_cells("A1:D1")
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=16, color=PS_NAVY)
    ws["A1"].alignment = Alignment(horizontal="center")

    headers = ["Stage", "Group 1", "Group 2", "Group 3"]
    for col, name in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=name)
        cell.font = Font(bold=True, color=PS_NAVY)
        cell.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 18
    for letter in ("B", "C", "D"):
        ws.column_dimensions[letter].width = 34

    last_row = 2 + len(GROWTH_STAGES)
    for i, (stage, rgb) in enumerate(GROWTH_STAGES, start=3):
        cell = ws.cell(row=i, column=1, value=stage)
        cell.font = Font(bold=True, color=PS_NAVY)
        cell.fill = PatternFill("solid", fgColor=rgb)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for col in (2, 3, 4):
            ws.cell(row=i, column=col).alignment = Alignment(
                wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 40

    for row in range(2, last_row + 1):
        for col in range(1, 5):
            ws.cell(row=row, column=col).border = Border(
                top=THIN, bottom=THIN, left=THIN, right=THIN)


# ---------- Build: Nursery Template ----------
def build_nursery_template(nursery_types: list[str] | None = None,
                           planting_dates: int = 1) -> Path:
    """Build the template workbook for one nursery.

    The sheet set and its order come from nursery_spec.json, so the tabs here
    and the SHEET_* constants the VBA uses cannot drift apart.
    """
    nursery_types = nursery_types or ["Selection"]
    sheet_names = tabs_for(nursery_types, planting_dates)

    wb = Workbook()
    wb.remove(wb.active)

    # ---- Home ----
    home = wb.create_sheet("Home")
    home.sheet_view.showGridLines = False

    next_row = banner(home, "Nursery Workflow", "Click ▶ buttons to run each step in order", cols=14)

    # Quick status box
    home.merge_cells(f"A{next_row}:E{next_row+3}")
    c = home[f"A{next_row}"]
    c.value = ("📋  This workbook follows the proposed Sorghum nursery workflow.\n"
               "    Run each step in order. Stats below auto-refresh after each step.")
    c.font = Font(size=11, color=PS_INK_SOFT)
    c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
    c.fill = PatternFill("solid", fgColor=PS_GREY)
    home.row_dimensions[next_row].height = 50
    next_row += 5

    # ---- Workflow steps (3 phases, color-coded) ----
    home.cell(row=next_row, column=1, value="PHASE 1 — PRE-FIELD PREP").font = Font(
        bold=True, color=PS_NAVY, size=12)
    home.cell(row=next_row, column=1).alignment = Alignment(indent=1)
    next_row += 1

    prep_steps = [
        ("▶ 1. Initialise from PRISM export",
         "btnInitNursery",
         "Asks for the nursery code, validates that PRISM data is pasted into the "
         "Nursery site tab, stamps Nursery data."),
        ("▶ 2. Generate ALL workbook tabs",
         "btnGenerateAllTabs",
         "Builds every output tab from the pasted PRISM data: Map, Material Map, "
         "Packet Prep (with QR text + colored digits), Nursery list, Fieldbook, "
         "BC0 labels, Date recording, Pulling bags, TFMSA Spray plots, Hy Heights."),
        ("▶ 3. Sort packets for racking (LSD Radix)",
         "btnSortForRacking",
         "Re-sorts Packet Prep by Rack Order ↑ then Spike ↑ — the physical "
         "pick-up order for racking."),
    ]

    field_steps = [
        ("▶ 4. Record replacement (Packeting/Planting)",
         "btnAddReplacement",
         "Opens a dialog to capture a replacement event with stage dropdown."),
        ("▶ 5. Record planting error",
         "btnAddPlantingError",
         "Logs a planting error in the Replacements done tab."),
        ("▶ 6. Spray track + date recording",
         "btnRecordSpray",
         "Adds a TFMSA / IMI / HPPD spray application with date."),
        ("▶ 7. AB bag pulling",
         "btnRecordABPull",
         "Records pulled bags with date and count."),
    ]

    post_steps = [
        ("▶ 8. Pull updated Nursery site from PRISM",
         "btnImportUpdated",
         "After Breeder updates PRISM with replacements + errors, paste the "
         "refreshed Nursery site export and re-run Step 2 to regenerate every tab."),
        ("▶ 9. Refresh dashboard",
         "btnRefreshDashboard",
         "Recalculates the live stats panel below."),
        ("▶ 10. Push to Hub",
         "btnPushToHub",
         "Writes this nursery's summary to the shared registry.csv so the Nursery "
         "Hub workbook sees it on its dashboard."),
    ]

    def render_step_block(start_row: int, steps: list, color: str) -> int:
        r = start_row
        for label, macro, desc in steps:
            cell_range = f"A{r}:E{r+1}"
            button_cell(home, cell_range, label, macro, color=color, description=desc)
            # Macro name stored in hidden col G for the VBA router
            home.cell(row=r, column=7, value=macro).font = Font(color="BBBBBB", size=9)
            home.cell(row=r, column=7).alignment = Alignment(indent=1)
            # Description in F
            home.cell(row=r, column=6, value=desc).font = Font(
                size=10, color=PS_INK_SOFT, italic=True)
            home.cell(row=r, column=6).alignment = Alignment(
                wrap_text=True, vertical="center", indent=1)
            home.merge_cells(start_row=r, end_row=r+1, start_column=6, end_column=6)
            home.row_dimensions[r].height = 22
            home.row_dimensions[r+1].height = 22
            r += 3
        return r

    next_row = render_step_block(next_row, prep_steps, PS_BLUE)
    next_row += 1

    home.cell(row=next_row, column=1, value="PHASE 2 — FIELD OPERATIONS").font = Font(
        bold=True, color=PS_NAVY, size=12)
    home.cell(row=next_row, column=1).alignment = Alignment(indent=1)
    next_row += 1
    next_row = render_step_block(next_row, field_steps, PS_DEEP)
    next_row += 1

    home.cell(row=next_row, column=1, value="PHASE 3 — POST-FIELD & SYNC").font = Font(
        bold=True, color=PS_NAVY, size=12)
    home.cell(row=next_row, column=1).alignment = Alignment(indent=1)
    next_row += 1
    next_row = render_step_block(next_row, post_steps, PS_NAVY)
    next_row += 2

    # ---- Local dashboard ----
    home.cell(row=next_row, column=1, value="LIVE DASHBOARD").font = Font(
        bold=True, color=PS_NAVY, size=14)
    home.cell(row=next_row, column=1).alignment = Alignment(indent=1)
    next_row += 1
    home.cell(row=next_row, column=1, value="Refreshed by step 12. Pulled from the working sheets.").font = Font(
        size=10, color=PS_MUTED, italic=True)
    home.cell(row=next_row, column=1).alignment = Alignment(indent=1)
    next_row += 2

    metrics = [
        ("Total packets", "DASH_TotalPackets", "=IFERROR(COUNTA('Nursery site'!A:A)-1,0)"),
        ("Unique source IDs", "DASH_UniqueSources", "=IFERROR(COUNTA('Nursery list'!A:A)-1,0)"),
        # Replacements and planting errors now share one tab, split by Stage.
        ("Packeting replacements", "DASH_Replacements",
         "=IFERROR(COUNTIF('Replacements and Errors'!A:A,\"Packeting\"),0)"),
        ("Planting errors logged", "DASH_PlantingErrors",
         "=IFERROR(COUNTIF('Replacements and Errors'!A:A,\"Planting\"),0)"),
        ("Awaiting decision", "DASH_AwaitingDecision",
         "=IFERROR(COUNTIF('Replacements and Errors'!J:J,\"Waiting decision\"),0)"),
        ("Last synced to Hub", "DASH_LastSync", "Never"),
    ]
    for i, (label, name, formula) in enumerate(metrics):
        col = (i % 3) * 4 + 1   # 1, 5, 9
        row_block = next_row + (i // 3) * 5
        # Label
        home.cell(row=row_block, column=col, value=label).font = Font(
            bold=True, color=PS_MUTED, size=10)
        home.cell(row=row_block, column=col).alignment = Alignment(indent=1)
        # Big value
        target = home.cell(row=row_block+1, column=col, value=formula)
        target.font = Font(bold=True, size=24, color=PS_BLUE)
        target.alignment = Alignment(horizontal="left", indent=1)
        home.merge_cells(start_row=row_block+1, end_row=row_block+2,
                         start_column=col, end_column=col+2)
        # Named range for VBA to reference
        try:
            dn = DefinedName(name=name,
                             attr_text=f"Home!${get_column_letter(col)}${row_block+1}")
            wb.defined_names[name] = dn
        except Exception:
            pass

    # Column widths
    for col_letter, width in [("A", 12), ("B", 24), ("C", 24), ("D", 14),
                              ("E", 18), ("F", 60), ("G", 28)]:
        home.column_dimensions[col_letter].width = width

    # ---- Settings tab ----
    settings = wb.create_sheet("Settings")
    banner(settings, "Settings", "Workbook & Hub configuration", cols=4)
    settings.cell(row=5, column=1, value="Setting").font = Font(bold=True, color=PS_NAVY)
    settings.cell(row=5, column=2, value="Value").font = Font(bold=True, color=PS_NAVY)
    settings.cell(row=5, column=3, value="Notes").font = Font(bold=True, color=PS_NAVY)
    for r, (k, v, note) in enumerate([
        ("Nursery code", "(set after step 1)", "Short code, e.g. AUGT1-26S-IMI"),
        ("Nursery name", "", "Full nursery name"),
        ("Breeder", "", "Lead breeder for this nursery"),
        ("Season", "", "e.g. 2026S"),
        ("Hub registry folder", "~/Documents/PacificSeeds/Nurseries/",
         "Shared folder where Hub looks for registry.csv. Same on every machine."),
        ("Qty per packet", 1.4, "Multiplier for required qty per Source ID"),
        ("Filter bulk treatment above", 10, "Source IDs with > X reps go to bulk list"),
    ], start=6):
        settings.cell(row=r, column=1, value=k).font = Font(color=PS_INK_SOFT)
        sv = settings.cell(row=r, column=2, value=v)
        sv.fill = PatternFill("solid", fgColor=PS_PALE)
        sv.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
        settings.cell(row=r, column=3, value=note).font = Font(italic=True, color=PS_MUTED, size=10)
    settings.column_dimensions["A"].width = 26
    settings.column_dimensions["B"].width = 42
    settings.column_dimensions["C"].width = 55

    # ---- Tab definitions, keyed by base name ----
    # The workbook's sheet list comes from nursery_spec.json; this maps each
    # name to its subtitle and headers. 'Packet Prep N' resolves to the
    # 'Packet Prep' entry, so every split shares one definition.
    PRISM_HEADERS = [
        "Range", "Row", "Material ID", "Inbred Code", "Source ID", "CMS reaction",
        "Generation", "Comments", "Pedigree", "Hybrid Code", "Trait Name",
        "Plant #", "Loc Seq#", "SubSeq Flag", "Entry Book Project",
        "Entry Book Name", "Entry #",
    ]

    tab_defs = {
        "Nursery site": (
            "Paste the PRISM export here (headers row 5, data row 6+). Step 2 reads this.",
            PRISM_HEADERS),
        "Material Map": (
            "Auto-built by Step 2. Range x Row grid carrying Material ID, "
            "Inbred code and Hybrid code.", []),
        "Field Map": (
            "Same grid as Material Map, without material information. Use the "
            "wizard button to set planting dates, spikes and run direction.", []),
        "Nursery data": (
            "Printable front page. Fill the right-hand column in by hand.", []),
        "Packet Prep": (
            "Auto-built by Step 2. QR CODE text + colored digit columns. "
            "Feed this tab to your barcode printer machine.",
            ["QR CODE", "Range", "Row", "Plot", "SPIKE#", "RACK ORDER",
             "Split no.",
             "Thousands/Black", "Hundreds/Red", "Tens/Green", "Ones/Blue",
             "Material ID", "Inbred Code", "Source ID", "CMS reaction",
             "Generation", "Comments", "Pedigree", "Hybrid Code",
             "Trait Name", "Plant #", "Loc Seq#", "SubSeq Flag",
             "Entry Book Project", "Entry Book Name", "Entry #"]),
        "Nursery list": (
            "Auto-built by Step 2 — unique Source IDs ascending, with repeats and qty.",
            ["", "", "Source ID (Hybrid Code)", "Repeats", "Qty Required",
             "Inbred Code", "Hybrid Code", "Notes"]),
        "Replacements and Errors": (
            "Log replacements and planting errors here. Stage marks which.",
            REPLACEMENT_HEADERS),
        "Updated nursery site": (
            "Paste the re-downloaded PRISM export here, then run the Fieldbook "
            "button on this tab.",
            PRISM_HEADERS),
        "Fieldbook": (
            "Auto-built from Updated nursery site — print-ready field reference.",
            []),
        "Date recording": (
            "AB nurseries only. S 1 and S 2 hold day-of-month numbers.", []),
        "Operations": (
            "Growth-stage tracker — fill in each group per stage.", []),
        "Comments": (
            "Free-form team notes, same layout as Operations.", []),
        # --- nursery-type extras -------------------------------------------
        "BC0 labels": (
            "Auto-built by Step 2 — BC* generations only (TFMSA / Pollen tracking).",
            ["Range", "Row", "Crossed bags", "TFMSA", "Pollen", "TFMSA/Pollen",
             "Nursery Name", "Bagging Info", "Material ID", "Source ID", "Gen",
             "CMS", "Comments"]),
        "Pulling bags": (
            "Auto-built by Step 2 — same population as Date recording, bag-pulling tracker.",
            ["Range", "Row", "Plot", "1", "2",
             "Material ID", "Source ID", "Gen", "CMS", "Comments"]),
        "BC0 TFMSA record": (
            "Day 7 / Day 10 / Day 13 TFMSA spray date observations.", []),
        "TFMSA Spray plots": (
            "Auto-built by Step 2 — BC* plots that get TFMSA.",
            ["Range", "Row", "Source ID", "CMS", "Gen"]),
        "Hy Heights": (
            "Auto-built by Step 2 — F1 hybrids height tracker.",
            ["Range", "Row", "Height in CM", "Material ID", "Source ID", "Gen", "CMS"]),
    }

    for name in sheet_names:
        if name in ("Home", "Settings"):
            continue                       # already created above
        base = "Packet Prep" if name.startswith("Packet Prep ") else name
        subtitle, headers = tab_defs[base]

        ws = wb.create_sheet(name)
        ws.sheet_view.showGridLines = False

        # Tabs whose layout is dictated by the client's sample bypass the
        # standard banner+header treatment entirely.
        if base == "Nursery data":
            build_nursery_data_sheet(ws)
            continue
        if base in ("Operations", "Comments"):
            build_grouped_sheet(
                ws,
                "Operations Overview" if base == "Operations" else "Field Comments")
            continue

        n = banner(ws, name, subtitle, cols=max(8, len(headers) or 8))
        if headers:
            header_row(ws, n, headers)
            for i, _ in enumerate(headers, start=1):
                ws.column_dimensions[get_column_letter(i)].width = 16
            ws.freeze_panes = ws.cell(row=n + 1, column=1)
        else:
            ws.cell(row=n, column=1,
                    value="(auto-built — do not edit by hand)").font = Font(
                italic=True, color=PS_MUTED)

        if base == "Replacements and Errors":
            decorate_replacements_sheet(ws, n)

    # Settings is machinery, not a workflow tab.
    for hidden_name in load_spec()["hidden"]:
        wb[hidden_name].sheet_state = "hidden"

    OUT.mkdir(exist_ok=True)
    path = OUT / "Nursery_Template.xlsx"
    wb.save(path)
    return path


# ---------- Build: Nursery Hub ----------
def build_nursery_hub() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    # Home
    home = wb.create_sheet("Home")
    home.sheet_view.showGridLines = False
    next_row = banner(home, "Pacific Seeds · Nursery Hub",
                      "Central view of every nursery across the program", cols=8)

    button_cell(home, f"A{next_row}:C{next_row+1}", "▶ Refresh Dashboard",
                "hubRefreshDashboard", color=PS_BLUE,
                description="Re-reads registry.csv from the shared folder and rebuilds the Dashboard tab.")
    home.row_dimensions[next_row].height = 22
    home.row_dimensions[next_row+1].height = 22

    button_cell(home, f"E{next_row}:G{next_row+1}", "▶ Register This Hub Folder",
                "hubRegisterFolder", color=PS_DEEP,
                description="Sets the shared folder where each Nursery workbook will write its summary.")

    next_row += 3
    button_cell(home, f"A{next_row}:C{next_row+1}", "▶ Open Nursery Folder",
                "hubOpenFolder", color=PS_NAVY,
                description="Opens the registry folder in Finder / File Explorer so you can pick a nursery to edit.")
    home.row_dimensions[next_row].height = 22
    home.row_dimensions[next_row+1].height = 22

    button_cell(home, f"E{next_row}:G{next_row+1}", "▶ New Nursery from Template",
                "hubCreateFromTemplate", color=PS_ORANGE,
                description="Clones Nursery_Template.xlsm into the registry folder with a new nursery code.")

    next_row += 4
    home.cell(row=next_row, column=1, value="QUICK STATS").font = Font(
        bold=True, color=PS_NAVY, size=14)
    home.cell(row=next_row, column=1).alignment = Alignment(indent=1)
    next_row += 1
    home.cell(row=next_row, column=1, value="Aggregated across every nursery registered in this Hub.").font = Font(
        size=10, color=PS_MUTED, italic=True)
    home.cell(row=next_row, column=1).alignment = Alignment(indent=1)
    next_row += 2

    hub_metrics = [
        ("Nurseries registered", "HUB_NurseryCount"),
        ("Total packets across all nurseries", "HUB_TotalPackets"),
        ("Replacements logged", "HUB_TotalReplacements"),
        ("Planting errors logged", "HUB_TotalErrors"),
    ]
    for i, (label, name) in enumerate(hub_metrics):
        col = (i % 2) * 4 + 1
        row_block = next_row + (i // 2) * 5
        home.cell(row=row_block, column=col, value=label).font = Font(
            bold=True, color=PS_MUTED, size=10)
        home.cell(row=row_block, column=col).alignment = Alignment(indent=1)
        target = home.cell(row=row_block+1, column=col, value=0)
        target.font = Font(bold=True, size=26, color=PS_BLUE)
        target.alignment = Alignment(horizontal="left", indent=1)
        home.merge_cells(start_row=row_block+1, end_row=row_block+2,
                         start_column=col, end_column=col+2)
        try:
            dn = DefinedName(name=name,
                             attr_text=f"Home!${get_column_letter(col)}${row_block+1}")
            wb.defined_names[name] = dn
        except Exception:
            pass

    for col_letter, width in [("A", 24), ("B", 24), ("C", 24), ("D", 6),
                              ("E", 24), ("F", 24), ("G", 24), ("H", 8)]:
        home.column_dimensions[col_letter].width = width

    # Dashboard
    dash = wb.create_sheet("Dashboard")
    dash.sheet_view.showGridLines = False
    n = banner(dash, "Nurseries — Live Dashboard",
               "Auto-populated from registry.csv when you press Refresh.", cols=10)
    dash_headers = ["Nursery code", "Season", "Breeder", "Packets",
                    "Replacements", "Errors", "Sprays", "Last update",
                    "File path", "Status"]
    header_row(dash, n, dash_headers)
    dash.freeze_panes = dash.cell(row=n + 1, column=1)
    for i, w in enumerate([22, 10, 16, 10, 14, 10, 10, 22, 50, 14], start=1):
        dash.column_dimensions[get_column_letter(i)].width = w

    # Settings
    settings = wb.create_sheet("Settings")
    settings.sheet_view.showGridLines = False
    banner(settings, "Hub Settings", "Set once, then Refresh.", cols=4)
    settings_rows = [
        ("Registry folder", "~/Documents/PacificSeeds/Nurseries/",
         "All nurseries write registry rows into this folder. Same on every workstation."),
        ("Registry file", "registry.csv",
         "Shared CSV inside the folder above. Don't rename — Hub looks for this name."),
        ("Template path", "~/Documents/PacificSeeds/Templates/Nursery_Template.xlsm",
         "Used by 'New Nursery from Template' button."),
    ]
    settings.cell(row=5, column=1, value="Setting").font = Font(bold=True, color=PS_NAVY)
    settings.cell(row=5, column=2, value="Value").font = Font(bold=True, color=PS_NAVY)
    settings.cell(row=5, column=3, value="Notes").font = Font(bold=True, color=PS_NAVY)
    for r, (k, v, note) in enumerate(settings_rows, start=6):
        settings.cell(row=r, column=1, value=k).font = Font(color=PS_INK_SOFT)
        sv = settings.cell(row=r, column=2, value=v)
        sv.fill = PatternFill("solid", fgColor=PS_PALE)
        sv.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
        settings.cell(row=r, column=3, value=note).font = Font(italic=True, color=PS_MUTED, size=10)
        # Named ranges for VBA
        nm = "HUB_" + k.replace(" ", "_")
        try:
            wb.defined_names[nm] = DefinedName(name=nm,
                                               attr_text=f"Settings!$B${r}")
        except Exception:
            pass
    settings.column_dimensions["A"].width = 22
    settings.column_dimensions["B"].width = 55
    settings.column_dimensions["C"].width = 60

    wb.move_sheet("Home", offset=-wb.sheetnames.index("Home"))

    OUT.mkdir(exist_ok=True)
    path = OUT / "Nursery_Hub.xlsx"
    wb.save(path)
    return path


def _maybe_bake(xlsx_path: Path, label: str) -> Path:
    """If a seed vbaProject.bin exists for this workbook, bake it in and
    return the .xlsm path. Otherwise just return the .xlsx path unchanged.
    Seeds live in excel_workflow/seeds/<stem>.vbaProject.bin and are
    created by running extract_seed.py against a known-good .xlsm."""
    seed = SEEDS / f"{xlsx_path.stem}.vbaProject.bin"
    if not seed.exists():
        print(f"     ℹ {label}: no seed at {seed.relative_to(ROOT)} — "
              f"keeping .xlsx only. Run extract_seed.py once to lock in VBA.")
        return xlsx_path

    try:
        from excel_workflow.bake_vba import bake_vba
    except ImportError:            # script mode puts excel_workflow/ on sys.path
        from bake_vba import bake_vba
    xlsm = xlsx_path.with_suffix(".xlsm")
    bake_vba(xlsx_path, seed, xlsm)
    print(f"     ✓ Baked VBA into {xlsm.name} (seed: {seed.name})")
    return xlsm


def main() -> None:
    print("→ Building workbooks…")
    t = build_nursery_template()
    print(f"  · {t.name}")
    out_t = _maybe_bake(t, "Template")

    h = build_nursery_hub()
    print(f"  · {h.name}")
    out_h = _maybe_bake(h, "Hub")

    print()
    print("Outputs:")
    print(f"  {out_t}")
    print(f"  {out_h}")


if __name__ == "__main__":
    main()
