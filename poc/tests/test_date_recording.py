"""Month assignment, checked against the client's own sample data.

Selection columns hold a bare day of the month, so the month has to be
inferred: recording starts on a known day and runs forward, and a day lower
than the start day has already wrapped into the next month. That rule comes
from GenerateS1S2TrendAnalysis in the appendix of the client's change
document; these check our port of it against the 396 rows they sent.

The sample was recorded starting 28 July 2026, which the data itself shows:
days 28-30 carry "Jul" and days 1-27 carry "Aug".
"""
from __future__ import annotations

from pathlib import Path

import pytest

from excel_workflow.nursery_algos import assign_month, trend_counts

SAMPLE = Path(__file__).resolve().parent.parent / "Tab information.xlsx"

# What the sample was recorded from.
START_DAY, START_MONTH, START_YEAR = 28, 7, 2026

MONTH_NAMES = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}


def sample_rows():
    """(S1 day, S1 month, S2 day, S2 month) from the client's sheet."""
    openpyxl = pytest.importorskip("openpyxl")
    if not SAMPLE.exists():
        pytest.skip(f"client sample not present at {SAMPLE}")
    wb = openpyxl.load_workbook(SAMPLE, data_only=True)
    ws = wb["Date recording"]
    return [(r[4], r[5], r[6], r[7])
            for r in ws.iter_rows(min_row=2, values_only=True)]


# ------------------------------------------------------------------ the rule


def test_a_day_on_or_after_the_start_day_keeps_the_start_month():
    assert assign_month(28, 28, 7, 2026) == (2026, 7, 28)
    assert assign_month(30, 28, 7, 2026) == (2026, 7, 30)


def test_a_day_below_the_start_day_has_wrapped_into_the_next_month():
    assert assign_month(1, 28, 7, 2026) == (2026, 8, 1)
    assert assign_month(27, 28, 7, 2026) == (2026, 8, 27)


def test_december_wraps_into_january_of_the_next_year():
    assert assign_month(5, 20, 12, 2026) == (2027, 1, 5)
    assert assign_month(25, 20, 12, 2026) == (2026, 12, 25)


@pytest.mark.parametrize("value", ["", None, "  ", "n/a", "-", 0, "0", -3])
def test_unrecorded_selections_resolve_to_nothing(value):
    """A blank or zero is an unrecorded selection, not a recording on day 0.

    DateSerial(y, m, 0) would silently mean the last day of the previous
    month, so the VBA guards on `dayValue > 0` and so do we.
    """
    assert assign_month(value, 28, 7, 2026) is None


def test_numeric_strings_are_accepted():
    """The grid stores what was typed, which may be text."""
    assert assign_month("28", 28, 7, 2026) == (2026, 7, 28)


# ------------------------------------------------- against the client's sheet


def test_every_recorded_day_in_the_sample_matches_our_month():
    rows = sample_rows()
    checked = 0
    mismatches = []

    for s1, s1_month, s2, s2_month in rows:
        for day, month_name in ((s1, s1_month), (s2, s2_month)):
            stamp = assign_month(day, START_DAY, START_MONTH, START_YEAR)
            if stamp is None:
                continue
            checked += 1
            if MONTH_NAMES[stamp[1]] != month_name:
                mismatches.append((day, month_name, MONTH_NAMES[stamp[1]]))

    assert checked > 300, f"expected the full sample, only checked {checked}"
    assert not mismatches, f"month mismatches: {mismatches[:10]}"


def test_the_sample_only_spans_the_two_expected_months():
    """A sanity check on the fixture itself, so a swapped file is obvious."""
    rows = sample_rows()
    months = {m for _, m, _, _ in rows if m} | {m for _, _, _, m in rows if m}
    assert months == {"Jul", "Aug"}


# ------------------------------------------------------------------ counting


def test_counts_are_per_column_per_date():
    rows = [
        {"S 1": 28, "S 2": 29},
        {"S 1": 28, "S 2": 1},
        {"S 1": 1, "S 2": ""},
    ]
    series = dict(trend_counts(rows, ["S 1", "S 2"], 28, 7, 2026))
    assert series[(2026, 7, 28)] == {"S 1": 2, "S 2": 0}
    assert series[(2026, 7, 29)] == {"S 1": 0, "S 2": 1}
    assert series[(2026, 8, 1)] == {"S 1": 1, "S 2": 1}


def test_days_with_no_recording_are_still_reported():
    """The chart must show the gaps rather than joining across them."""
    rows = [{"S 1": 28}, {"S 1": 30}]
    series = dict(trend_counts(rows, ["S 1"], 28, 7, 2026))
    assert series[(2026, 7, 29)] == {"S 1": 0}
    assert sorted(series) == [(2026, 7, 28), (2026, 7, 29), (2026, 7, 30)]


def test_the_series_starts_at_the_start_date_even_if_nothing_was_recorded():
    series = trend_counts([{"S 1": 30}], ["S 1"], 28, 7, 2026)
    assert series[0][0] == (2026, 7, 28)


def test_zero_entries_are_excluded_from_the_counts():
    """18 rows of the client's sample carry a 0; none is a recording."""
    rows = [{"S 1": 0}, {"S 1": 28}, {"S 1": ""}]
    series = dict(trend_counts(rows, ["S 1"], 28, 7, 2026))
    assert series[(2026, 7, 28)] == {"S 1": 1}


def test_counting_spans_a_year_boundary():
    rows = [{"S 1": 20}, {"S 1": 2}]
    series = trend_counts(rows, ["S 1"], 20, 12, 2026)
    assert series[0][0] == (2026, 12, 20)
    assert series[-1][0] == (2027, 1, 2)


def test_four_selection_columns_are_supported():
    """The appendix handles S1 and S2; the app carries four passes."""
    rows = [{"S 1": 28, "S 2": 28, "S 3": 28, "S 4": 28}]
    columns = ["S 1", "S 2", "S 3", "S 4"]
    series = dict(trend_counts(rows, columns, 28, 7, 2026))
    assert series[(2026, 7, 28)] == {c: 1 for c in columns}
