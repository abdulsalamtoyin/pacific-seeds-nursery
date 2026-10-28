"""The declarative spec is the single source of truth for workbook structure.

Tab names and ordering here come from Workbook changes.docx; if these tests
and the document disagree, the document wins.
"""
import copy

import pytest

from excel_workflow.spec.loader import (
    app_spec, default_tabs, load_spec, tabs_for,
)


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


# --------------------------------------------------------------- app_overrides
#
# The desktop app's tabs moved ahead of the workbook's. Rather than fork the
# spec, the app's differences live in an app_overrides block that only the
# desktop side merges — so the workbook keeps building exactly as before.


def test_base_spec_still_has_the_combined_replacements_tab():
    """The workbook builder must not see the app's tab split."""
    assert "Replacements and Errors" in default_tabs()
    assert "Planting errors" not in tabs_for(["AB"], planting_dates=1)


def test_app_spec_splits_replacements_from_planting_errors():
    tabs = tabs_for(["Selection"], planting_dates=1, spec=app_spec())
    assert "Replacements and Errors" not in tabs
    assert tabs.index("Planting errors") == tabs.index("Replacements") + 1


def test_app_spec_moves_date_recording_after_fieldbook():
    tabs = tabs_for(["AB"], planting_dates=1, spec=app_spec())
    assert tabs.index("Date recording") == tabs.index("Fieldbook") + 1


def test_app_spec_still_hides_date_recording_from_non_ab():
    tabs = tabs_for(["Selection"], planting_dates=1, spec=app_spec())
    assert "Date recording" not in tabs


def test_app_spec_carries_four_selection_columns_and_their_months():
    names = [c["name"] for c in app_spec()["date_recording_columns"]]
    for expected in ["S 1", "S1 Month", "S 2", "S2 Month",
                     "S 3", "S3 Month", "S 4", "S4 Month"]:
        assert expected in names, f"{expected} missing from {names}"


def test_app_spec_carries_the_nursery_data_defaults():
    defaults = app_spec()["nursery_data_defaults"]
    assert defaults["Planter"] == "Almaco Precision Planter"
    assert defaults["Seeds/side"] == "34"
    assert defaults["Plot length"] == "4.5m"
    assert defaults["Alley way spacing"] == "0.75m"


def test_app_overrides_do_not_leak_into_the_merged_spec():
    assert "app_overrides" not in app_spec()


# ------------------------------------------------------------ anchor placement


def _spec_with_rule(rule: dict) -> dict:
    spec = copy.deepcopy(app_spec())
    spec["conditional"] = {"Date recording": rule}
    return spec


def test_before_anchor_places_the_tab_immediately_before():
    spec = _spec_with_rule({"types": ["AB"], "before": "Operations"})
    tabs = tabs_for(["AB"], planting_dates=1, spec=spec)
    assert tabs.index("Date recording") == tabs.index("Operations") - 1


def test_a_rule_with_both_anchors_is_rejected():
    spec = _spec_with_rule(
        {"types": ["AB"], "before": "Operations", "after": "Fieldbook"})
    with pytest.raises(ValueError, match="exactly one of 'before' or 'after'"):
        tabs_for(["AB"], planting_dates=1, spec=spec)


def test_a_rule_with_no_anchor_is_rejected():
    spec = _spec_with_rule({"types": ["AB"]})
    with pytest.raises(ValueError, match="exactly one of 'before' or 'after'"):
        tabs_for(["AB"], planting_dates=1, spec=spec)


def test_app_spec_has_no_duplicate_tab_names():
    tabs = tabs_for(["AB", "Hybrid", "Other"], planting_dates=3,
                    spec=app_spec())
    assert len(tabs) == len(set(tabs))
