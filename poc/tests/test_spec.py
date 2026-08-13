"""The declarative spec is the single source of truth for workbook structure.

Tab names and ordering here come from Workbook changes.docx; if these tests
and the document disagree, the document wins.
"""
import pytest

from excel_workflow.spec.loader import default_tabs, load_spec, tabs_for


def test_default_tabs_are_in_the_specified_order():
    assert default_tabs() == [
        "Nursery site", "Material Map", "Field Map", "Nursery data",
        "Packet Prep", "Nursery list", "Replacements and Errors",
        "Updated nursery site", "Fieldbook", "Operations", "Comments",
    ]


def test_selection_nursery_has_no_date_recording():
    assert "Date recording" not in tabs_for(["Selection"], planting_dates=1)


def test_ab_nursery_inserts_date_recording_before_operations():
    tabs = tabs_for(["AB"], planting_dates=1)
    assert tabs.index("Date recording") == tabs.index("Operations") - 1


def test_packet_prep_fans_out_per_planting_date():
    tabs = tabs_for(["Selection"], planting_dates=3)
    assert [t for t in tabs if t.startswith("Packet Prep")] == [
        "Packet Prep 1", "Packet Prep 2", "Packet Prep 3",
    ]


def test_single_planting_date_still_numbers_the_tab():
    assert "Packet Prep 1" in tabs_for(["Selection"], planting_dates=1)


def test_home_is_first_and_settings_present():
    tabs = tabs_for(["Selection"], planting_dates=1)
    assert tabs[0] == "Home"
    assert "Settings" in tabs


def test_ab_unlocks_its_extras_only():
    tabs = tabs_for(["AB"], planting_dates=1)
    assert "BC0 labels" in tabs
    assert "Pulling bags" in tabs
    assert "Hy Heights" not in tabs


def test_hybrid_unlocks_hy_heights():
    assert "Hy Heights" in tabs_for(["Hybrid"], planting_dates=1)


def test_multiple_types_union_their_extras():
    tabs = tabs_for(["AB", "Hybrid"], planting_dates=1)
    assert "BC0 labels" in tabs and "Hy Heights" in tabs


def test_core_sheet_count_matches_the_documented_formula():
    # Home + Settings + the 10 non-fanned default tabs = 12, plus N Packet
    # Prep tabs, plus Date recording for AB. Packet Prep is replaced by its
    # fan-out copies, not joined by them.
    n = 2
    tabs = tabs_for(["AB"], planting_dates=n)
    core = [t for t in tabs if t not in load_spec()["extras"]]
    assert len(core) == 12 + n + 1


def test_selection_core_sheet_count():
    tabs = tabs_for(["Selection"], planting_dates=1)
    core = [t for t in tabs if t not in load_spec()["extras"]]
    assert len(core) == 12 + 1


def test_tabs_for_rejects_zero_planting_dates():
    with pytest.raises(ValueError, match="planting_dates must be >= 1"):
        tabs_for(["Selection"], planting_dates=0)


def test_no_duplicate_tab_names():
    tabs = tabs_for(["AB", "Hybrid", "Other"], planting_dates=3)
    assert len(tabs) == len(set(tabs))
