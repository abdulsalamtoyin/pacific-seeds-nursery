"""SpecConstants.bas is generated, never hand-edited.

These tests pin the contract the VBA relies on: every tab reachable from the
spec gets a constant, and the renamed tabs leave no stale constant behind.
"""
from excel_workflow.gen_spec_constants import const_name, render_constants
from excel_workflow.spec.loader import load_spec


def test_const_name_slugifies_tab_names():
    assert const_name("Nursery site") == "SHEET_NURSERY_SITE"
    assert const_name("Replacements and Errors") == "SHEET_REPLACEMENTS_AND_ERRORS"
    assert const_name("BC0 labels") == "SHEET_BC0_LABELS"


def test_const_name_collapses_punctuation():
    assert const_name("O/E") == "SHEET_O_E"
    assert const_name("In. Code") == "SHEET_IN_CODE"


def test_generated_bas_declares_the_module_name():
    assert 'Attribute VB_Name = "SpecConstants"' in render_constants()


def test_generated_bas_warns_against_hand_editing():
    assert "DO NOT EDIT" in render_constants()


def test_every_spec_tab_gets_a_constant():
    spec = load_spec()
    out = render_constants()
    tabs = (spec["always_first"] + spec["hidden"] + spec["default_tabs"]
            + list(spec["conditional"]) + list(spec["extras"]))
    for tab in tabs:
        assert const_name(tab) in out, f"missing constant for {tab}"
        assert f'"{tab}"' in out, f"missing literal for {tab}"


def test_renamed_tabs_have_no_stale_constants():
    out = render_constants()
    # 'Map' became 'Field Map'; the two error tabs merged into one.
    assert 'As String = "Map"' not in out
    assert 'As String = "Replacements done"' not in out
    assert 'As String = "Planting error noted"' not in out


def test_fieldbook_column_positions_are_one_based_and_ordered():
    out = render_constants()
    assert "Public Const FB_COL_RANGE" in out
    assert "As Long = 1" in out
    # Plot moved from column 3 to column 7 in the v2 layout.
    plot_line = [l for l in out.splitlines() if "FB_COL_PLOT" in l][0]
    assert plot_line.strip().endswith("= 7")


def test_generated_bas_is_valid_vba_option_explicit():
    assert "Option Explicit" in render_constants()
