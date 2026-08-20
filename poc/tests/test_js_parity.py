"""The JS and Python ordering rules must agree, value for value.

The Excel workbook and the desktop app are meant to behave identically. The
rules are written three times — Python, JavaScript, VBA — so something has to
hold them together. This runs the real JS under node and compares its output
against the Python original over the same inputs.

VBA cannot be executed here; it is checked by hand against the same cases.
"""
from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

import pytest

from excel_workflow import nursery_algos as py

PWA_DIR = Path(__file__).resolve().parent.parent / "pwa"
ALGOS_JS = PWA_DIR / "nursery-algos.js"

ROWS = list(range(1, 41))
PLOTS = [(r, w) for r in range(1, 9) for w in range(1, 7)]
BAND_PLOTS = [(r, w) for r in range(2, 8) for w in (2, 3, 26, 27)]
DOP_ROWS = [[1, 2, 5, 6, 9, 10], [3, 4, 7, 8, 11, 12]]

DRIVER = """
import {{
  spikeForRow, runDirection, bandForRow, parityForRow,
  serpentineByRange, serpentineTwoRowBands, assignSplits,
}} from {module!r};

const rows = {rows};
const plots = {plots};
const bandPlots = {band_plots};
const dopRows = {dop_rows};

const splits = assignSplits(dopRows);

console.log(JSON.stringify({{
  spikes: rows.map(spikeForRow),
  runs: rows.map(runDirection),
  bands: rows.map(bandForRow),
  parities: rows.map(parityForRow),
  serpRange: serpentineByRange(plots),
  serpBands: serpentineTwoRowBands(bandPlots),
  splits: [...splits.entries()].sort((a, b) => a[0] - b[0]),
}}));
"""


@pytest.fixture(scope="module")
def js_output(tmp_path_factory) -> dict:
    if shutil.which("node") is None:
        pytest.skip("node not installed — cannot verify JS/Python parity")

    driver = tmp_path_factory.mktemp("parity") / "driver.mjs"
    driver.write_text(
        DRIVER.format(
            module=str(ALGOS_JS),
            rows=json.dumps(ROWS),
            plots=json.dumps([list(p) for p in PLOTS]),
            band_plots=json.dumps([list(p) for p in BAND_PLOTS]),
            dop_rows=json.dumps(DOP_ROWS),
        ),
        encoding="utf-8",
    )
    proc = subprocess.run(
        ["node", str(driver)], capture_output=True, text=True, timeout=30)
    if proc.returncode != 0:
        pytest.fail(f"node driver failed:\n{proc.stderr}")
    return json.loads(proc.stdout)


def test_js_module_exists():
    assert ALGOS_JS.exists(), f"missing {ALGOS_JS}"


def test_spike_matches_python(js_output):
    assert js_output["spikes"] == [py.spike_for_row(r) for r in ROWS]


def test_run_direction_matches_python(js_output):
    assert js_output["runs"] == [py.run_direction(r) for r in ROWS]


def test_band_matches_python(js_output):
    assert js_output["bands"] == [py.band_for_row(r) for r in ROWS]


def test_parity_matches_python(js_output):
    assert js_output["parities"] == [py.parity_for_row(r) for r in ROWS]


def test_serpentine_by_range_matches_python(js_output):
    expected = [list(p) for p in py.serpentine_by_range(PLOTS)]
    assert js_output["serpRange"] == expected


def test_serpentine_two_row_bands_matches_python(js_output):
    expected = [list(p) for p in py.serpentine_two_row_bands(BAND_PLOTS)]
    assert js_output["serpBands"] == expected


def test_splits_match_python(js_output):
    expected = sorted(py.assign_splits(DOP_ROWS).items())
    assert [tuple(pair) for pair in js_output["splits"]] == expected


@pytest.mark.parametrize("types,dates", [
    (["Selection"], 1),
    (["Selection"], 3),
    (["AB"], 2),
    (["Hybrid"], 1),
    (["AB", "Hybrid"], 2),
    (["Selection", "AB", "Hybrid", "Other"], 4),
])
def test_tabs_for_matches_python(tmp_path, types, dates):
    """The generated JS spec must produce the same sheet list as the loader.

    The app's spec is the base plus app_overrides, so the comparison is against
    tabs_for(..., spec=app_spec()) — the same thing render_js() emits.
    """
    if shutil.which("node") is None:
        pytest.skip("node not installed")

    from excel_workflow.spec.loader import app_spec, tabs_for

    spec_js = PWA_DIR / "nursery-spec.js"
    assert spec_js.exists(), "run: python -m excel_workflow.gen_spec_constants"

    driver = tmp_path / "tabs.mjs"
    driver.write_text(
        f"import {{ tabsFor }} from {str(spec_js)!r};\n"
        f"console.log(JSON.stringify(tabsFor({json.dumps(types)}, {dates})));\n",
        encoding="utf-8",
    )
    proc = subprocess.run(["node", str(driver)], capture_output=True,
                          text=True, timeout=30)
    if proc.returncode != 0:
        pytest.fail(f"node driver failed:\n{proc.stderr}")
    assert json.loads(proc.stdout) == tabs_for(types, dates, spec=app_spec())


def test_js_rejects_a_conditional_with_two_anchors(tmp_path):
    """anchorOf() must refuse the same rule the Python _anchor() refuses."""
    if shutil.which("node") is None:
        pytest.skip("node not installed")

    spec_js = PWA_DIR / "nursery-spec.js"
    driver = tmp_path / "anchors.mjs"
    driver.write_text(
        f"import {{ SPEC, tabsFor }} from {str(spec_js)!r};\n"
        "SPEC.conditional['Date recording'] = "
        "{ types: ['AB'], before: 'Operations', after: 'Fieldbook' };\n"
        "try { tabsFor(['AB'], 1); console.log('NO ERROR'); }\n"
        "catch (e) { console.log(e.message); }\n",
        encoding="utf-8",
    )
    out = subprocess.run(["node", str(driver)], capture_output=True,
                         text=True, timeout=30).stdout.strip()
    # Same wording as the Python ValueError, so either side reads the same.
    assert out == (
        "conditional rule 'Date recording' must set exactly one of "
        "'before' or 'after'")


def test_js_rejects_overlapping_splits(tmp_path):
    if shutil.which("node") is None:
        pytest.skip("node not installed")
    driver = tmp_path / "overlap.mjs"
    driver.write_text(
        f"import {{ assignSplits }} from {str(ALGOS_JS)!r};\n"
        "try { assignSplits([[3, 4], [4, 5]]); console.log('NO ERROR'); }\n"
        "catch (e) { console.log(e.message); }\n",
        encoding="utf-8",
    )
    out = subprocess.run(["node", str(driver)], capture_output=True,
                         text=True, timeout=30).stdout.strip()
    # Same wording as the Python ValueError, so either side reads the same.
    assert out == "row 4 appears in splits 1 and 2"
