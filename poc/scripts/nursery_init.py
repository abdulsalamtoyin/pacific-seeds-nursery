"""Initialise a nursery from a PRISM export.

Replaces three VBA modules:
  - Recurrent file template.bas  → creates workbook tabs
  - Packetprinting.bas           → sorts, inserts Plot/SPIKE#/RACK ORDER columns
  - (new) packet QR PDF generation

Outputs:
  - data/nursery.sqlite (packets table seeded for this nursery)
  - output/<nursery>_workbook.xlsx (the tab-structured workbook)
  - output/<nursery>_packets.pdf (printable QR-coded packet labels)
"""
from __future__ import annotations

import argparse
import io
import sqlite3
import uuid
from pathlib import Path

import qrcode
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

from map_parser import MapLayout, parse_map

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
OUT_DIR = ROOT / "output"
DB_PATH = DATA_DIR / "nursery.sqlite"

PRISM_COLS = [
    "Range", "Row", "Material ID", "Inbred Code", "Source ID", "CMS reaction",
    "Generation", "Comments", "Pedigree", "Hybrid Code", "Trait Name",
    "Plant #", "Loc Seq#", "SubSeq Flag", "Entry Book Project",
    "Entry Book Name", "Entry #",
]

WORKBOOK_TABS = [
    "Nursery site", "Field Map", "Material Map", "Nursery data", "Nursery list",
    "Packet prep", "Replacements and errors", "Updated nursery site",
    "Fieldbook", "Operations", "Comments", "Additionals",
]

GROWTH_STAGES = [
    ("Seedling",      (254, 192, 0)),
    ("Vegetative",    (169, 208, 142)),
    ("Heading",       (255, 0, 0)),
    ("Flowering",     (173, 216, 230)),
    ("Grain filling", (255, 255, 153)),
    ("Harvest",       (204, 255, 204)),
    ("Post Harvest",  (204, 204, 255)),
]


def init_db() -> sqlite3.Connection:
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS nurseries (
      code TEXT PRIMARY KEY,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS packets (
      uuid TEXT PRIMARY KEY,
      nursery_code TEXT NOT NULL REFERENCES nurseries(code),
      range_n INTEGER, row_n INTEGER, plot TEXT,
      spike INTEGER, rack_order INTEGER,
      material_id TEXT, source_id TEXT, inbred_code TEXT,
      generation TEXT, cms TEXT, comments TEXT,
      pedigree TEXT, hybrid_code TEXT,
      entry_no INTEGER
    );
    CREATE INDEX IF NOT EXISTS idx_packets_nursery ON packets(nursery_code);
    CREATE TABLE IF NOT EXISTS events (
      uuid TEXT PRIMARY KEY,
      packet_uuid TEXT REFERENCES packets(uuid),
      nursery_code TEXT NOT NULL,
      type TEXT NOT NULL,           -- replacement | planting_error | spray | ab_pull | note
      payload TEXT,                 -- JSON
      tech_id TEXT,
      captured_at TEXT NOT NULL,    -- ISO from device
      gps_lat REAL, gps_lon REAL,
      synced_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE INDEX IF NOT EXISTS idx_events_nursery ON events(nursery_code);
    """)
    return conn


def read_prism_export(path: Path, sheet: str) -> list[dict]:
    wb = load_workbook(path, read_only=True, data_only=True)
    if sheet not in wb.sheetnames:
        raise SystemExit(f"Sheet '{sheet}' not found. Available: {wb.sheetnames}")
    ws = wb[sheet]
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(rows)]
    missing = [c for c in PRISM_COLS if c not in header]
    if missing:
        print(f"WARNING: PRISM export is missing columns: {missing}")
    idx = {c: header.index(c) for c in PRISM_COLS if c in header}

    def cell(r, c):
        return r[idx[c]] if c in idx and idx[c] < len(r) else None

    out = []
    for r in rows:
        if r is None or all(v is None for v in r):
            continue
        if cell(r, "Range") is None or cell(r, "Row") is None:
            continue
        out.append({
            "range_n":     int(cell(r, "Range")),
            "row_n":       int(cell(r, "Row")),
            "material_id": cell(r, "Material ID"),
            "inbred_code": cell(r, "Inbred Code"),
            "source_id":   cell(r, "Source ID"),
            "cms":         cell(r, "CMS reaction"),
            "generation":  cell(r, "Generation"),
            "comments":    cell(r, "Comments"),
            "pedigree":    cell(r, "Pedigree"),
            "hybrid_code": cell(r, "Hybrid Code"),
            "entry_no":    cell(r, "Entry #"),
        })
    return out


def assign_spike_rack(packets: list[dict], layout: MapLayout | None) -> None:
    """Set spike + rack_order on each packet.

    If a parsed MapLayout is available, spike numbers come from the Map (1, 2, …)
    and rack order is a serpentine walk down the ranges within that spike.
    Otherwise we fall back to "spike = range_n" with per-range serpentine.
    """
    for p in packets:
        p["plot"] = f"{p['range_n']}_{p['row_n']}"

    if layout is None:
        by_range: dict[int, list[dict]] = {}
        for p in packets:
            by_range.setdefault(p["range_n"], []).append(p)
        for rng, items in by_range.items():
            reverse = (rng % 2 == 0)
            items.sort(key=lambda p: p["row_n"], reverse=reverse)
            for i, p in enumerate(items, start=1):
                p["spike"] = rng
                p["rack_order"] = i
        return

    # Group by spike using field-row → spike mapping from the Map.
    # Rack-order = serpentine walk: for each range (ascending), pick that
    # range's packets within the spike's field-rows in alternating direction.
    by_spike: dict[int, list[dict]] = {}
    for p in packets:
        spike = layout.row_to_spike.get(p["row_n"], 0)  # 0 = off-map ("Other")
        p["spike"] = spike
        by_spike.setdefault(spike, []).append(p)

    for spike, items in by_spike.items():
        field_row_order = layout.row_order_in_spike.get(spike, [])
        row_pos = {fr: i for i, fr in enumerate(field_row_order)}

        def sort_key(pp):
            rng = pp["range_n"]
            i = row_pos.get(pp["row_n"], 9999)
            # Even-numbered ranges traverse the spike in reverse (serpentine).
            return (rng, -i if rng % 2 == 0 else i)

        items.sort(key=sort_key)
        for i, p in enumerate(items, start=1):
            p["rack_order"] = i


def insert_packets(conn: sqlite3.Connection, nursery_code: str, packets: list[dict]) -> None:
    conn.execute("INSERT OR REPLACE INTO nurseries(code) VALUES (?)", (nursery_code,))
    conn.execute("DELETE FROM packets WHERE nursery_code = ?", (nursery_code,))
    for p in packets:
        p["uuid"] = uuid.uuid4().hex[:8]
        conn.execute("""
            INSERT INTO packets(uuid, nursery_code, range_n, row_n, plot, spike, rack_order,
                material_id, source_id, inbred_code, generation, cms, comments,
                pedigree, hybrid_code, entry_no)
            VALUES (:uuid, :nc, :range_n, :row_n, :plot, :spike, :rack_order,
                :material_id, :source_id, :inbred_code, :generation, :cms, :comments,
                :pedigree, :hybrid_code, :entry_no)
        """, {**p, "nc": nursery_code})
    conn.commit()


def qr_payload(nursery_code: str, packet_uuid: str) -> str:
    return f"SNUR:{nursery_code}:{packet_uuid}"


def write_workbook(out_path: Path, nursery_code: str, packets: list[dict]) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    # All standard tabs from the proposed workflow.
    for name in WORKBOOK_TABS:
        wb.create_sheet(name)

    # Fieldbook banner.
    fb = wb["Fieldbook"]
    fb["A1"] = ('Download Nursery file from PRISM once the "Replacements and errors" '
                'tab status is UPDATED')
    fb["A1"].fill = PatternFill("solid", fgColor="FF0000")
    fb["A1"].font = Font(bold=True)

    # Replacements and errors header.
    re_ws = wb["Replacements and errors"]
    re_ws.append(["QR payload", "Plot", "Stage", "Original Source ID",
                  "Replaced with", "Tech", "Captured at", "Status"])
    for c in re_ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
    re_ws.column_dimensions["A"].width = 28
    re_ws.column_dimensions["E"].width = 30

    # Operations tab styled with growth stages.
    ops = wb["Operations"]
    ops["A1"] = "Operations Overview"
    ops["A1"].font = Font(bold=True)
    ops["A1"].fill = PatternFill("solid", fgColor="A9D08E")
    for i, (stage, rgb) in enumerate(GROWTH_STAGES, start=4):
        c = ops.cell(row=i, column=1, value=stage)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="{:02X}{:02X}{:02X}".format(*rgb))
        c.alignment = Alignment(horizontal="center")
        ops.cell(row=i, column=2, value="*")
        ops.cell(row=i, column=3, value="*")

    # Packet prep tab — what used to come from Step1/Step2 packet printing VBA.
    pp = wb["Packet prep"]
    pp.append(["QR payload", "Plot", "Range", "Row", "Spike", "Rack order",
               "Material ID", "Source ID", "Generation", "CMS"])
    for c in pp[1]:
        c.font = Font(bold=True)
    for p in sorted(packets, key=lambda p: (p["spike"], p["rack_order"])):
        pp.append([
            qr_payload(nursery_code, p["uuid"]),
            p["plot"], p["range_n"], p["row_n"], p["spike"], p["rack_order"],
            p["material_id"], p["source_id"], p["generation"], p["cms"],
        ])

    # Nursery list — unique source IDs with repeat count and qty required.
    nl = wb["Nursery list"]
    nl.append(["Source ID", "Repeats", "Qty required (1.4 × repeats)"])
    for c in nl[1]:
        c.font = Font(bold=True)
    counts: dict[str, int] = {}
    for p in packets:
        sid = p["source_id"] or "(unknown)"
        counts[sid] = counts.get(sid, 0) + 1
    for sid in sorted(counts):
        nl.append([sid, counts[sid], round(counts[sid] * 1.4, 1)])

    out_path.parent.mkdir(exist_ok=True)
    wb.save(out_path)


def write_packet_pdf(out_path: Path, nursery_code: str, packets: list[dict]) -> None:
    """One QR-coded label per packet, 4 columns × 8 rows per A4 page.

    Header band on each page carries the Pacific Seeds brand colour and
    nursery code so a misplaced sheet of labels is still identifiable.
    """
    from reportlab.lib.colors import HexColor
    from reportlab.lib.utils import ImageReader

    PS_BLUE = HexColor("#0678CD")
    PS_NAVY = HexColor("#092A40")
    PS_ORANGE = HexColor("#e45138")

    c = canvas.Canvas(str(out_path), pagesize=A4)
    page_w, page_h = A4
    band_h = 12 * mm
    cols, rows = 4, 8
    margin_x, margin_y = 8 * mm, 8 * mm
    cell_w = (page_w - 2 * margin_x) / cols
    cell_h = (page_h - band_h - 2 * margin_y) / rows

    sorted_packets = sorted(packets, key=lambda p: (p["spike"], p["rack_order"]))
    total_pages = (len(sorted_packets) + cols * rows - 1) // (cols * rows)

    def draw_band(page_no: int):
        c.setFillColor(PS_NAVY)
        c.rect(0, page_h - band_h, page_w, band_h, fill=1, stroke=0)
        c.setFillColor(HexColor("#3AB3E5"))
        c.rect(0, page_h - band_h - 2, page_w, 2, fill=1, stroke=0)
        c.setFillColor(HexColor("#FFFFFF"))
        c.setFont("Helvetica-Bold", 11)
        c.drawString(margin_x, page_h - 8 * mm, "Pacific Seeds")
        c.setFont("Helvetica", 9)
        c.setFillColor(HexColor("#8FCEEE"))
        c.drawString(margin_x + 28 * mm, page_h - 8 * mm,
                     f"Nursery {nursery_code}  ·  Page {page_no}/{total_pages}")

    for i, p in enumerate(sorted_packets):
        page_no = i // (cols * rows) + 1
        if i % (cols * rows) == 0:
            if i > 0:
                c.showPage()
            draw_band(page_no)
        idx_on_page = i % (cols * rows)
        col = idx_on_page % cols
        row = idx_on_page // cols
        x = margin_x + col * cell_w
        y = page_h - band_h - margin_y - (row + 1) * cell_h

        payload = qr_payload(nursery_code, p["uuid"])
        qr_img = qrcode.make(payload, box_size=4, border=1)
        buf = io.BytesIO(); qr_img.save(buf, format="PNG"); buf.seek(0)
        c.drawImage(ImageReader(buf), x + 2 * mm, y + 8 * mm,
                    width=cell_h - 12 * mm, height=cell_h - 12 * mm,
                    preserveAspectRatio=True)

        text_x = x + cell_h - 8 * mm
        text_y = y + cell_h - 6 * mm
        c.setFillColor(PS_NAVY)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(text_x, text_y, f"Plot {p['plot']}")
        c.setFillColor(PS_BLUE)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(text_x, text_y - 11, f"Spike {p['spike']} · Rack {p['rack_order']}")
        c.setFillColor(HexColor("#34526e"))
        c.setFont("Helvetica", 7)
        c.drawString(text_x, text_y - 20, f"Mat: {str(p['material_id'] or '')[:14]}")
        c.drawString(text_x, text_y - 29, f"Src: {str(p['source_id'] or '')[:14]}")
        c.drawString(text_x, text_y - 38, f"{p['generation'] or ''} · CMS {p['cms'] or ''}")
        c.setFillColor(PS_ORANGE)
        c.setFont("Helvetica-Oblique", 5)
        c.drawString(x + 2 * mm, y + 2 * mm, payload)

    c.save()


def init_nursery(input_path: Path, nursery_code: str,
                 sheet: str = "Sheet1") -> dict:
    """Run the full nursery initialisation pipeline. Callable from API code."""
    OUT_DIR.mkdir(exist_ok=True)
    packets = read_prism_export(input_path, sheet)
    layout = parse_map(input_path)
    assign_spike_rack(packets, layout)
    conn = init_db()
    insert_packets(conn, nursery_code, packets)
    wb_path = OUT_DIR / f"{nursery_code}_workbook.xlsx"
    write_workbook(wb_path, nursery_code, packets)
    pdf_path = OUT_DIR / f"{nursery_code}_packets.pdf"
    write_packet_pdf(pdf_path, nursery_code, packets)
    return {
        "nursery_code": nursery_code,
        "packet_count": len(packets),
        "spike_count": (layout.spike_count if layout else 0),
        "map_source": (layout.sheet_name if layout else "(no Map found — serpentine fallback)"),
        "workbook_path": str(wb_path),
        "pdf_path": str(pdf_path),
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, type=Path,
                    help="PRISM export .xlsx")
    ap.add_argument("--nursery-code", required=True,
                    help="Short nursery code, e.g. AUGT1-26S-IMI")
    ap.add_argument("--sheet", default="Sheet1", help="Sheet name in the export")
    args = ap.parse_args()

    print(f"Reading PRISM export {args.input} (sheet={args.sheet})…")
    result = init_nursery(args.input, args.nursery_code, args.sheet)
    print(f"  → {result['packet_count']} packets")
    print(f"  → Map: {result['map_source']} ({result['spike_count']} spike(s))")
    print(f"  → wrote {result['workbook_path']}")
    print(f"  → wrote {result['pdf_path']}")
    print("Done.")


if __name__ == "__main__":
    main()
