"""The nursery store must not lose a nursery.

Everything the user types lives in browser storage, so the two dangerous
moments are the upgrade from the old single-nursery payload and switching
between nurseries. Both are checked here by running the real module under
node against an in-memory storage backend.
"""
from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

import pytest

STORE_JS = Path(__file__).resolve().parent.parent / "pwa" / "store.js"

PREAMBLE = """
import * as store from {module!r};

// Stand-in for localStorage, seeded with whatever the case needs.
const cell = new Map(Object.entries({seed}));
store.useBackend({{
  getItem: (k) => (cell.has(k) ? cell.get(k) : null),
  setItem: (k, v) => cell.set(k, String(v)),
  removeItem: (k) => cell.delete(k),
}});
const raw = () => JSON.parse(cell.get("ps-nursery-workbook-v3") ?? "null");
"""


def run_js(tmp_path: Path, body: str, seed: dict | None = None) -> dict:
    """Execute a snippet against store.js and return its JSON output."""
    if shutil.which("node") is None:
        pytest.skip("node not installed — cannot exercise store.js")

    driver = tmp_path / "store_case.mjs"
    driver.write_text(
        PREAMBLE.format(module=str(STORE_JS), seed=json.dumps(seed or {}))
        + body,
        encoding="utf-8",
    )
    proc = subprocess.run(["node", str(driver)], capture_output=True,
                          text=True, timeout=30)
    if proc.returncode != 0:
        pytest.fail(f"node driver failed:\n{proc.stderr}")
    return json.loads(proc.stdout)


LEGACY = {
    "code": "OLD-1",
    "fileName": "OLD-1",
    "types": ["AB"],
    "plantingDates": 2,
    "seedQty": "40",
    "prism": [{"Range": 1, "Row": 2, "Material ID": "M-1"}],
    "replacements": [{"Reason": "spilled"}],
    "operations": {"Seedling|1": "watered"},
}


def test_starts_empty_with_one_untitled_nursery(tmp_path):
    out = run_js(tmp_path, """
      const n = store.activeNursery();
      console.log(JSON.stringify({
        count: store.listNurseries().length,
        code: n.code,
        prism: n.prism.length,
      }));
    """)
    assert out == {"count": 1, "code": "", "prism": 0}


def test_migrates_a_legacy_single_nursery_payload(tmp_path):
    out = run_js(tmp_path, """
      const n = store.activeNursery();
      console.log(JSON.stringify({
        count: store.listNurseries().length,
        code: n.code,
        types: n.types,
        plantingDates: n.plantingDates,
        prism: n.prism.length,
        replacements: n.replacements.length,
        operations: n.operations,
        plantingErrors: n.plantingErrors,
        schema: raw().schema,
      }));
    """, seed={"ps-nursery-workbook": json.dumps(LEGACY)})

    assert out["count"] == 1
    assert out["code"] == "OLD-1"
    assert out["types"] == ["AB"]
    assert out["plantingDates"] == 2
    assert out["prism"] == 1
    assert out["replacements"] == 1
    assert out["operations"] == {"Seedling|1": "watered"}
    # Fields the old shape never had are filled in, not left undefined.
    assert out["plantingErrors"] == []
    assert out["schema"] == 3


def test_migration_is_written_back_so_it_runs_once(tmp_path):
    """After migrating, the v3 key holds the nursery — no second migration."""
    out = run_js(tmp_path, """
      store.activeNursery();
      console.log(JSON.stringify({
        persisted: Object.keys(raw().nurseries).length,
        code: Object.values(raw().nurseries)[0].code,
        activePointsAtIt: raw().activeId in raw().nurseries,
      }));
    """, seed={"ps-nursery-workbook": json.dumps(LEGACY)})
    assert out == {"persisted": 1, "code": "OLD-1", "activePointsAtIt": True}


def test_a_new_nursery_does_not_overwrite_the_old_one(tmp_path):
    """The bug this store exists to fix."""
    out = run_js(tmp_path, """
      store.activeNursery();
      const fresh = store.createNursery({ code: "NEW-2", types: ["Selection"] });
      const list = store.listNurseries();
      console.log(JSON.stringify({
        codes: list.map((n) => n.code),
        active: store.activeNursery().code,
        bothKept: list.length === 2,
        freshIsActive: store.activeId() === fresh,
      }));
    """, seed={"ps-nursery-workbook": json.dumps(LEGACY)})

    assert out["codes"] == ["OLD-1", "NEW-2"]
    assert out["active"] == "NEW-2"
    assert out["bothKept"] is True
    assert out["freshIsActive"] is True


def test_switching_back_restores_the_earlier_nursery(tmp_path):
    out = run_js(tmp_path, """
      const a = store.createNursery({ code: "A" });
      store.activeNursery().seedQty = "11";
      store.save();
      store.createNursery({ code: "B" });
      store.activeNursery().seedQty = "22";
      store.save();
      store.switchNursery(a);
      console.log(JSON.stringify({
        code: store.activeNursery().code,
        qty: store.activeNursery().seedQty,
      }));
    """)
    assert out == {"code": "A", "qty": "11"}


def test_nurseries_do_not_share_arrays(tmp_path):
    """Two nurseries built from the same blank must not alias each other."""
    out = run_js(tmp_path, """
      store.createNursery({ code: "A" });
      store.activeNursery().prism.push({ Range: 1 });
      store.save();
      store.createNursery({ code: "B" });
      console.log(JSON.stringify({ b: store.activeNursery().prism.length }));
    """)
    assert out == {"b": 0}


def test_duplicate_is_a_deep_copy(tmp_path):
    out = run_js(tmp_path, """
      const a = store.createNursery({ code: "A" });
      store.activeNursery().prism.push({ Range: 1 });
      store.save();
      store.duplicateNursery(a);
      store.activeNursery().prism.push({ Range: 2 });
      store.save();
      store.switchNursery(a);
      console.log(JSON.stringify({
        original: store.activeNursery().prism.length,
        codes: store.listNurseries().map((n) => n.code),
      }));
    """)
    assert out["original"] == 1
    assert out["codes"] == ["A", "A (copy)"]


def test_deleting_the_active_nursery_repoints_it(tmp_path):
    out = run_js(tmp_path, """
      const a = store.createNursery({ code: "A" });
      const b = store.createNursery({ code: "B" });
      store.deleteNursery(b);
      console.log(JSON.stringify({ active: store.activeNursery().code }));
    """)
    assert out == {"active": "A"}


def test_deleting_an_unknown_nursery_is_refused(tmp_path):
    out = run_js(tmp_path, """
      store.createNursery({ code: "A" });
      let message = "NO ERROR";
      try { store.deleteNursery("n-nope"); } catch (e) { message = e.message; }
      console.log(JSON.stringify({ message }));
    """)
    assert out == {"message": "no such nursery: n-nope"}


def test_grid_state_is_per_tab_and_per_nursery(tmp_path):
    out = run_js(tmp_path, """
      const a = store.createNursery({ code: "A" });
      store.setGridState("Fieldbook", { sort: { key: "Row", dir: "asc" } });
      store.setGridState("Nursery list",
        { sort: { key: "Source ID", dir: "desc" } });
      store.createNursery({ code: "B" });
      const bFieldbook = store.gridState("Fieldbook").sort;
      store.switchNursery(a);
      console.log(JSON.stringify({
        bFieldbook,
        aFieldbook: store.gridState("Fieldbook").sort,
        aList: store.gridState("Nursery list").sort,
      }));
    """)
    assert out["bFieldbook"] is None
    assert out["aFieldbook"] == {"key": "Row", "dir": "asc"}
    assert out["aList"] == {"key": "Source ID", "dir": "desc"}


def test_grid_state_backfills_keys_added_later(tmp_path):
    """A grid saved before a key existed must not crash the render."""
    out = run_js(tmp_path, """
      store.createNursery({ code: "A" });
      // Simulate an older payload that predates extraColumns/edits/hidden.
      store.activeNursery().grids["Fieldbook"] = { sort: null };
      store.save();
      const g = store.gridState("Fieldbook");
      console.log(JSON.stringify({
        extraColumns: g.extraColumns, edits: g.edits, hidden: g.hidden,
      }));
    """)
    assert out == {"extraColumns": [], "edits": {}, "hidden": []}


def test_corrupt_storage_falls_back_to_empty(tmp_path):
    out = run_js(tmp_path, """
      console.log(JSON.stringify({ count: store.listNurseries().length }));
    """, seed={"ps-nursery-workbook-v3": "{not json"})
    assert out == {"count": 0}
