"""Grid rules that decide what a cell shows and which rows survive a filter.

The dangerous one is the hand-edit overlay. Computed tabs (Packet Prep,
Fieldbook, Material Map) rebuild their rows from PRISM on every render, so a
typed correction has nowhere durable to live and is kept as an overlay keyed
by plot. If that keying is wrong, corrections silently attach to the wrong
row after a re-import — or vanish. These run the real module under node.
"""
from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

import pytest

CORE_JS = Path(__file__).resolve().parent.parent / "pwa" / "grid-core.js"

PREAMBLE = """
import * as core from {module!r};

const valueOf = (rec, key) =>
  core.effectiveValue(rec, key, {{ owned: false, edits: EDITS }});
"""


def run_js(tmp_path: Path, body: str, edits: dict | None = None) -> dict:
    if shutil.which("node") is None:
        pytest.skip("node not installed — cannot exercise grid-core.js")

    driver = tmp_path / "grid_case.mjs"
    driver.write_text(
        f"const EDITS = {json.dumps(edits or {})};\n"
        + PREAMBLE.format(module=str(CORE_JS))
        + body,
        encoding="utf-8",
    )
    proc = subprocess.run(["node", str(driver)], capture_output=True,
                          text=True, timeout=30)
    if proc.returncode != 0:
        pytest.fail(f"node driver failed:\n{proc.stderr}")
    return json.loads(proc.stdout)


# Three plots as the grid wraps them: a stable key, the computed source row.
RECORDS_JS = """
const recs = [
  { key: "1_2", src: { Range: 1, Row: 2, CMS: "A", Reason: "" } },
  { key: "1_10", src: { Range: 1, Row: 10, CMS: "B", Reason: "" } },
  { key: "2_3", src: { Range: 2, Row: 3, CMS: "A", Reason: "" } },
];
"""


# ------------------------------------------------------------------ compare


def test_numbers_sort_as_numbers_not_text(tmp_path):
    out = run_js(tmp_path, """
      const rows = ["10", "2", "1", "20"];
      console.log(JSON.stringify({ sorted: rows.sort(core.compare) }));
    """)
    assert out == {"sorted": ["1", "2", "10", "20"]}


def test_blanks_sort_last(tmp_path):
    out = run_js(tmp_path, """
      const rows = ["b", "", "a"];
      console.log(JSON.stringify({ sorted: rows.sort(core.compare) }));
    """)
    assert out == {"sorted": ["a", "b", ""]}


# ------------------------------------------------------------------ overlay


def test_computed_cell_shows_its_source_value_when_unedited(tmp_path):
    out = run_js(tmp_path, RECORDS_JS + """
      console.log(JSON.stringify({ v: valueOf(recs[0], "CMS") }));
    """)
    assert out == {"v": "A"}


def test_hand_edit_wins_over_the_computed_value(tmp_path):
    out = run_js(tmp_path, RECORDS_JS + """
      console.log(JSON.stringify({
        v: valueOf(recs[0], "Reason"),
        edited: core.isEditedCell(recs[0], "Reason",
          { owned: false, edits: EDITS }),
        untouched: core.isEditedCell(recs[1], "Reason",
          { owned: false, edits: EDITS }),
      }));
    """, edits={"1_2": {"Reason": "reseeded"}})
    assert out == {"v": "reseeded", "edited": True, "untouched": False}


def test_overlay_follows_the_plot_not_the_row_position(tmp_path):
    """A re-import that reorders rows must not move the correction."""
    out = run_js(tmp_path, RECORDS_JS + """
      // PRISM comes back with the same plots in a different order.
      const reimported = [recs[2], recs[0], recs[1]];
      console.log(JSON.stringify({
        values: reimported.map((r) => valueOf(r, "Reason")),
        keys: reimported.map((r) => r.key),
      }));
    """, edits={"1_2": {"Reason": "reseeded"}})
    assert out["keys"] == ["2_3", "1_2", "1_10"]
    # The edit stayed on plot 1_2 rather than on "the first row".
    assert out["values"] == ["", "reseeded", ""]


def test_owned_rows_ignore_the_overlay(tmp_path):
    """Replacements own their data — the overlay must not shadow it."""
    out = run_js(tmp_path, RECORDS_JS + """
      console.log(JSON.stringify({
        v: core.effectiveValue(recs[0], "Reason", { owned: true, edits: EDITS }),
        edited: core.isEditedCell(recs[0], "Reason",
          { owned: true, edits: EDITS }),
      }));
    """, edits={"1_2": {"Reason": "reseeded"}})
    assert out == {"v": "", "edited": False}


def test_user_added_rows_hold_their_own_values(tmp_path):
    out = run_js(tmp_path, """
      const rec = { key: "extra:a1", src: { Reason: "typed" }, userAdded: true };
      console.log(JSON.stringify({
        v: core.effectiveValue(rec, "Reason", { owned: false, edits: EDITS }),
        overlaid: core.isOverlaid(rec, false),
      }));
    """, edits={"extra:a1": {"Reason": "overlay should not win"}})
    assert out == {"v": "typed", "overlaid": False}


def test_orphaned_edits_are_reported_not_dropped(tmp_path):
    out = run_js(tmp_path, RECORDS_JS + """
      console.log(JSON.stringify({
        orphans: core.orphanKeys(EDITS, recs.map((r) => r.key)),
      }));
    """, edits={"1_2": {"Reason": "kept"}, "9_9": {"Reason": "plot gone"}})
    assert out == {"orphans": ["9_9"]}


# ------------------------------------------------------------------ filter


def test_a_column_with_no_filter_entry_is_unfiltered(tmp_path):
    out = run_js(tmp_path, RECORDS_JS + """
      const kept = core.filterRecords(recs, {}, valueOf);
      console.log(JSON.stringify({ keys: kept.map((r) => r.key) }));
    """)
    assert out == {"keys": ["1_2", "1_10", "2_3"]}


def test_filter_keeps_only_allowed_values(tmp_path):
    out = run_js(tmp_path, RECORDS_JS + """
      const kept = core.filterRecords(recs, { CMS: ["A"] }, valueOf);
      console.log(JSON.stringify({ keys: kept.map((r) => r.key) }));
    """)
    assert out == {"keys": ["1_2", "2_3"]}


def test_filters_on_two_columns_are_combined(tmp_path):
    out = run_js(tmp_path, RECORDS_JS + """
      const kept = core.filterRecords(
        recs, { CMS: ["A"], Range: ["2"] }, valueOf);
      console.log(JSON.stringify({ keys: kept.map((r) => r.key) }));
    """)
    assert out == {"keys": ["2_3"]}


def test_filter_sees_the_hand_edited_value(tmp_path):
    """Filtering must match what the user can see, not the stale source."""
    out = run_js(tmp_path, RECORDS_JS + """
      const kept = core.filterRecords(recs, { CMS: ["B"] }, valueOf);
      console.log(JSON.stringify({ keys: kept.map((r) => r.key) }));
    """, edits={"1_2": {"CMS": "B"}})
    assert out == {"keys": ["1_2", "1_10"]}


# ------------------------------------------------------------------ sort


def test_sort_ascending_then_descending(tmp_path):
    out = run_js(tmp_path, RECORDS_JS + """
      const asc = core.sortRecords(recs, { key: "Row", dir: "asc" }, valueOf);
      const desc = core.sortRecords(recs, { key: "Row", dir: "desc" }, valueOf);
      console.log(JSON.stringify({
        asc: asc.map((r) => r.key), desc: desc.map((r) => r.key),
      }));
    """)
    # Row 2 before row 10 — the point of numeric-aware comparison.
    assert out == {"asc": ["1_2", "2_3", "1_10"],
                   "desc": ["1_10", "2_3", "1_2"]}


def test_sort_null_leaves_the_natural_order(tmp_path):
    out = run_js(tmp_path, RECORDS_JS + """
      const ordered = core.sortRecords(recs, null, valueOf);
      console.log(JSON.stringify({ keys: ordered.map((r) => r.key) }));
    """)
    assert out == {"keys": ["1_2", "1_10", "2_3"]}


def test_header_clicks_cycle_asc_desc_off(tmp_path):
    out = run_js(tmp_path, """
      const a = core.nextSort(null, "Row");
      const b = core.nextSort(a, "Row");
      const c = core.nextSort(b, "Row");
      const other = core.nextSort(b, "Range");
      console.log(JSON.stringify({ a, b, c, other }));
    """)
    assert out["a"] == {"key": "Row", "dir": "asc"}
    assert out["b"] == {"key": "Row", "dir": "desc"}
    assert out["c"] is None
    # Clicking a different header starts that column fresh.
    assert out["other"] == {"key": "Range", "dir": "asc"}


# ------------------------------------------------------------------ rack order


@pytest.mark.parametrize("count,width", [
    (1, 3), (999, 3), (1000, 4), (1001, 4), (9999, 4),
])
def test_rack_order_widens_past_999(tmp_path, count, width):
    out = run_js(tmp_path, f"""
      console.log(JSON.stringify({{ w: core.radixWidth({count}) }}));
    """)
    assert out == {"w": width}


def test_rack_order_pads_to_a_single_width(tmp_path):
    out = run_js(tmp_path, """
      console.log(JSON.stringify({
        three: ["7", "42", "356"].map((v) => core.padRack(v, 3)),
        four: ["7", "1000"].map((v) => core.padRack(v, 4)),
        nonNumeric: core.padRack("A1", 3),
      }));
    """)
    assert out == {"three": ["007", "042", "356"],
                   "four": ["0007", "1000"],
                   "nonNumeric": "A1"}
