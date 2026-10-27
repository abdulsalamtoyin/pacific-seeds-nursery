"""The browser's PRISM reader must see a file exactly as the server did.

POST /parse/prism was the only way into the app until it had to run as
static files on SharePoint. pwa/prism-reader.js replaces it there, so every
file below goes through both and the results have to match field for field —
a disagreement would mean the same export imports differently depending on
where the app is hosted.
"""
from __future__ import annotations

import asyncio
import os
import tempfile
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pytest
from fastapi import HTTPException, UploadFile
from openpyxl import Workbook

import browser_xlsx

# Importing the app creates its data folders; keep them out of the project.
_scratch = tempfile.mkdtemp(prefix="prism-reader-")
os.environ.setdefault("PS_DATA_DIR", _scratch)
os.environ.setdefault("PS_OUTPUT_DIR", _scratch)
from backend.app import parse_prism  # noqa: E402

HEADERS = ["Range", "Row", "Source ID", "Material ID", "Pedigree",
           "Generation", "CMS reaction", "Planted"]


def server_read(path: Path) -> dict:
    upload = UploadFile(file=BytesIO(path.read_bytes()), filename=path.name)
    return asyncio.run(parse_prism(file=upload, sheet=None))


def both(path: Path) -> tuple[dict, dict]:
    return server_read(path), browser_xlsx.read_prism(path)


def xlsx(tmp_path: Path, rows: list[list], name="prism.xlsx",
         title="Sheet 1", merge: str | None = None) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = title
    for row in rows:
        ws.append(row)
    if merge:
        ws.merge_cells(merge)
    path = tmp_path / name
    wb.save(path)
    return path


def text_file(tmp_path: Path, name: str, body: str,
              encoding="utf-8") -> Path:
    path = tmp_path / name
    path.write_bytes(body.encode(encoding))
    return path


PACKETS = [
    [1, 1, "SRC-0001", "MAT-01", "A/B", "F3", "A", date(2026, 8, 20)],
    [1, 2, "SRC-0002", "MAT-02", "C/D", "BC1", "B", datetime(2026, 9, 1, 0, 0)],
    [2, 1, "SRC-0003", "MAT-03", "  padded  ", "F4", "R", None],
]


# ------------------------------------------------------------------ xlsx


def test_header_on_the_first_row(tmp_path):
    server, browser = both(xlsx(tmp_path, [HEADERS, *PACKETS]))
    assert browser == server
    assert browser["count"] == 3


def test_header_found_below_prism_title_rows(tmp_path):
    """PRISM sometimes saves four lines of title above the header."""
    rows = [["Nursery export"], ["Generated 2026-08-20"], [], [],
            HEADERS, *PACKETS]
    server, browser = both(xlsx(tmp_path, rows))
    assert browser == server
    assert browser["rows"][0]["Source ID"] == "SRC-0001"


def test_dates_become_iso_text(tmp_path):
    _, browser = both(xlsx(tmp_path, [HEADERS, *PACKETS]))
    assert browser["rows"][0]["Planted"] == "2026-08-20"
    assert browser["rows"][1]["Planted"] == "2026-09-01"


def test_numbers_and_whitespace_match(tmp_path):
    server, browser = both(xlsx(tmp_path, [HEADERS, *PACKETS]))
    assert browser["rows"][0]["Range"] == server["rows"][0]["Range"] == "1"
    assert browser["rows"][2]["Pedigree"] == "padded"


def test_blank_rows_and_footers_are_dropped(tmp_path):
    rows = [HEADERS, PACKETS[0], [], ["", "", "", "", "Total: 1"]]
    server, browser = both(xlsx(tmp_path, rows))
    assert browser == server
    assert browser["count"] == 1


def test_short_rows_are_padded_with_blanks(tmp_path):
    rows = [HEADERS, [3, 4, "SRC-0009"]]
    server, browser = both(xlsx(tmp_path, rows))
    assert browser == server
    assert browser["rows"][0]["Pedigree"] == ""


def test_blank_header_cells_are_skipped(tmp_path):
    rows = [["Range", "Row", None, "Source ID"], [1, 1, "stray", "SRC-1"]]
    server, browser = both(xlsx(tmp_path, rows))
    assert browser == server
    assert browser["headers"] == ["Range", "Row", "Source ID"]


def test_merged_cells_hold_their_value_once(tmp_path):
    rows = [HEADERS, [1, 1, "SRC-0001", "MAT-01", "wide note", "", "", ""]]
    server, browser = both(xlsx(tmp_path, rows, merge="E2:G2"))
    assert browser == server


def test_sheet_name_is_reported(tmp_path):
    server, browser = both(xlsx(tmp_path, [HEADERS, *PACKETS]))
    assert browser["sheet"] == server["sheet"] == "Sheet 1"


def test_xlsm_is_read_like_xlsx(tmp_path):
    path = xlsx(tmp_path, [HEADERS, *PACKETS], name="prism.xlsm")
    server, browser = both(path)
    assert browser == server


# ------------------------------------------------------------- delimited


def test_csv_with_quotes_bom_and_crlf(tmp_path):
    body = ("﻿Range,Row,Source ID,Pedigree\r\n"
            '1,1,SRC-0001,"A/B, selected"\r\n'
            '1,2,SRC-0002,"says ""hi"""\r\n')
    server, browser = both(text_file(tmp_path, "prism.csv", body))
    assert browser == server
    assert browser["rows"][0]["Pedigree"] == "A/B, selected"
    assert browser["rows"][1]["Pedigree"] == 'says "hi"'


def test_tsv_by_extension(tmp_path):
    body = "Range\tRow\tSource ID\n1\t1\tSRC-0001\n"
    server, browser = both(text_file(tmp_path, "prism.tsv", body))
    assert browser == server


def test_tab_delimited_txt_is_sniffed(tmp_path):
    body = "Range\tRow\tSource ID\n1\t1\tSRC-0001\n2\t3\tSRC-0002"
    server, browser = both(text_file(tmp_path, "prism.txt", body))
    assert browser == server
    assert browser["count"] == 2


# ------------------------------------------------------------- refusals


def test_no_header_row_is_refused_with_the_same_message(tmp_path):
    path = xlsx(tmp_path, [["Plot", "Entry"], [1, 2]])
    with pytest.raises(HTTPException) as server_err:
        server_read(path)
    with pytest.raises(browser_xlsx.BrowserError) as browser_err:
        browser_xlsx.read_prism(path)
    assert str(browser_err.value) == server_err.value.detail


def test_a_file_that_is_not_a_spreadsheet_is_refused(tmp_path):
    path = text_file(tmp_path, "prism.xlsx", "not a zip at all")
    with pytest.raises(browser_xlsx.BrowserError, match="Could not read file"):
        browser_xlsx.read_prism(path)
