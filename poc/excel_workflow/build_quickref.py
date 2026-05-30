"""Generate a one-page A4 quick-reference PDF for the technician to print
and pin next to their monitor. Pacific Seeds branding throughout.
"""
from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table,
    TableStyle,
)

HERE = Path(__file__).resolve().parent
OUT = HERE / "dist" / "QuickReference.pdf"

PS_NAVY   = colors.HexColor("#092A40")
PS_BLUE   = colors.HexColor("#0678CD")
PS_SKY    = colors.HexColor("#3AB3E5")
PS_LIGHT  = colors.HexColor("#8FCEEE")
PS_PALE   = colors.HexColor("#E9F3FC")
PS_ORANGE = colors.HexColor("#E45138")
PS_WHEAT  = colors.HexColor("#DDB318")
PS_GREEN  = colors.HexColor("#28A745")
PS_GREY   = colors.HexColor("#F3F7FB")
PS_INK    = colors.HexColor("#092A40")
PS_SOFT   = colors.HexColor("#34526E")
PS_MUTED  = colors.HexColor("#5A7896")
PS_BORDER = colors.HexColor("#D8E3ED")


def banner(canvas, doc):
    page_w, page_h = A4
    canvas.saveState()
    # Navy header bar
    canvas.setFillColor(PS_NAVY)
    canvas.rect(0, page_h - 22 * mm, page_w, 22 * mm, fill=1, stroke=0)
    # Sky underline
    canvas.setFillColor(PS_SKY)
    canvas.rect(0, page_h - 23 * mm, page_w, 1.2 * mm, fill=1, stroke=0)
    # Wordmark
    canvas.setFillColor(PS_SKY)
    canvas.setFont("Helvetica-Bold", 9)
    canvas.drawString(15 * mm, page_h - 9 * mm, "PACIFIC SEEDS")
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 16)
    canvas.drawString(15 * mm, page_h - 17 * mm, "Sorghum Nursery Workflow — Quick Reference")
    # Footer
    canvas.setFillColor(PS_MUTED)
    canvas.setFont("Helvetica-Oblique", 7)
    canvas.drawRightString(page_w - 15 * mm, 10 * mm, "Print and pin · v1.0 · See README.md for full instructions")
    canvas.restoreState()


def build():
    doc = BaseDocTemplate(
        str(OUT), pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=28 * mm, bottomMargin=15 * mm,
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin,
                  doc.width, doc.height, id="main")
    doc.addPageTemplates(PageTemplate(id="ps", frames=[frame], onPage=banner))

    styles = getSampleStyleSheet()
    H = ParagraphStyle("H", parent=styles["Heading2"],
                       textColor=PS_NAVY, fontName="Helvetica-Bold",
                       fontSize=11, spaceAfter=3, spaceBefore=6, leading=13)
    body = ParagraphStyle("body", parent=styles["BodyText"],
                          textColor=PS_INK, fontName="Helvetica",
                          fontSize=8.5, leading=11, spaceAfter=2)
    small = ParagraphStyle("small", parent=body, fontSize=7.5,
                           textColor=PS_SOFT, leading=10)

    story = []

    # ── First-time setup ──
    story.append(Paragraph("FIRST-TIME SETUP (≈ 2 minutes)", H))
    setup_data = [
        ["1.",
         "Unzip the pack into  <font name='Courier'>Documents\\PacificSeeds\\</font>",
         "Also create the sub-folder <font name='Courier'>Nurseries\\</font>"],
        ["2.",
         "Open each <font name='Courier'>.xlsm</font> in Excel",
         "Click <b>Enable Content</b> on the yellow security banner"],
        ["3.",
         "Open <font name='Courier'>Nursery_Hub.xlsm</font>",
         "Confirm the <b>Registry folder</b> on the Settings tab"],
    ]
    setup_tbl = Table(
        [[s, Paragraph(t, body), Paragraph(d, small)] for s, t, d in setup_data],
        colWidths=[8 * mm, 80 * mm, 92 * mm],
    )
    setup_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 11),
        ("TEXTCOLOR", (0, 0), (0, -1), PS_BLUE),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(setup_tbl)

    # ── Per-nursery workflow ──
    story.append(Paragraph("EVERY NEW NURSERY — 7 STEPS", H))
    story.append(Paragraph(
        "<b>1.</b> Duplicate <font name='Courier'>Nursery_Template.xlsm</font> in your <font name='Courier'>Nurseries\\</font> folder. "
        "<b>2.</b> Rename to the nursery code (e.g. <font name='Courier'>AUGT1-26S-IMI.xlsm</font>). "
        "<b>3.</b> Open it, Enable Macros. "
        "<b>4.</b> Click the <b>Nursery site</b> tab, paste the PRISM export starting at <b>cell A5</b>. "
        "<b>5.</b> Click the <b>Home</b> tab. "
        "<b>6.</b> <b>Double-click each blue button</b> on the Home tab, top-to-bottom. "
        "<b>7.</b> <b>Ctrl + S</b> to save (auto-publishes to Hub).",
        body,
    ))

    # ── Button reference ──
    story.append(Paragraph("BUTTONS ON THE HOME TAB", H))
    rows = [
        ["#", "Button", "What it does"],
        ["1",  "Initialise from PRISM",          "Asks for nursery code; stamps Nursery data"],
        ["2",  "Build Nursery list",             "Auto: unique Source IDs, Repeats, Qty Required (1.4 × reps)"],
        ["3",  "Design Field Map",               "Opens Field Map tab for manual layout / spike zones"],
        ["4",  "Generate Packet Prep & QRs",     "Creates Plot, Spike, Rack Order, QR payload per packet"],
        ["5",  "Sort for racking (LSD Radix)",   "Re-orders packets in physical pick-up order"],
        ["6",  "Record replacement",             "Dialog: stage, original source, replacement"],
        ["7",  "Record planting error",          "Dialog: severity, note"],
        ["8",  "Spray track + date",             "TFMSA / IMI / HPPD application + date"],
        ["9",  "AB bag pulling",                 "Bag count + date pulled"],
        ["10", "Pull updated PRISM",             "Opens Updated nursery site tab for fresh paste"],
        ["11", "Generate Fieldbook",             "Column reorder + serpentine sort + landscape print"],
        ["12", "Refresh dashboard",              "Recalculates the Home stats panel"],
        ["13", "Push to Hub",                    "Writes summary row to shared registry.csv"],
    ]
    tbl = Table(rows, colWidths=[8 * mm, 50 * mm, 122 * mm])
    tbl.setStyle(TableStyle([
        # Header row
        ("BACKGROUND", (0, 0), (-1, 0), PS_NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        # Body
        ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
        ("FONT", (1, 1), (1, -1), "Helvetica-Bold", 8),
        ("TEXTCOLOR", (1, 1), (1, -1), PS_NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, PS_GREY]),
        ("GRID", (0, 0), (-1, -1), 0.4, PS_BORDER),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]))
    # Colour the # column by phase
    phase_colors = {
        range(1, 6): PS_BLUE,            # Pre-field
        range(6, 10): colors.HexColor("#0075BD"),  # Field ops
        range(10, 14): PS_NAVY,          # Post-field
    }
    extra = []
    for i in range(1, 14):
        col = next(c for r, c in phase_colors.items() if i in r)
        extra.append(("BACKGROUND", (0, i), (0, i), col))
        extra.append(("TEXTCOLOR", (0, i), (0, i), colors.white))
        extra.append(("FONT", (0, i), (0, i), "Helvetica-Bold", 8.5))
    tbl.setStyle(TableStyle(extra))
    story.append(tbl)

    # ── Hub ──
    story.append(Paragraph("THE HUB — VIEW ALL NURSERIES", H))
    story.append(Paragraph(
        "Open <font name='Courier'>Nursery_Hub.xlsm</font>, double-click <b>▶ Refresh Dashboard</b>. "
        "The Dashboard tab lists every nursery that's been saved, with packet counts, "
        "replacements, errors and a hyperlink to open each workbook. "
        "Auto-syncs every time a nursery file is saved.",
        body,
    ))

    # ── Troubleshooting ──
    story.append(Paragraph("IF SOMETHING GOES WRONG", H))
    trouble = [
        ["Buttons don't react",
         "Macros aren't enabled — click Enable Content on the banner, or Trust Center → Macro Settings"],
        ["\"Path not found\" on save",
         "Create the folder Documents\\PacificSeeds\\Nurseries\\ manually"],
        ["\"Nursery site is empty\"",
         "Paste the PRISM export into the Nursery site tab at cell A5, then re-run Step 1"],
        ["Hub shows no nurseries",
         "Save the nursery workbook at least once; confirm both Hub and Nursery use the same Registry folder"],
        ["Compile error on open",
         "Close, redownload the .xlsm files (don't edit in SharePoint Web Excel — download to desktop)"],
    ]
    trouble_tbl = Table(
        [[Paragraph(f"<b>{a}</b>", small), Paragraph(b, small)] for a, b in trouble],
        colWidths=[55 * mm, 125 * mm],
    )
    trouble_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BACKGROUND", (0, 0), (0, -1), PS_PALE),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("BOX", (0, 0), (-1, -1), 0.4, PS_BORDER),
        ("LINEBEFORE", (1, 0), (1, -1), 0.4, PS_BORDER),
    ]))
    story.append(trouble_tbl)

    doc.build(story)
    print(f"  → {OUT} ({OUT.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    build()
