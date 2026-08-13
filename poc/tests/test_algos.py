"""Ordering and numbering rules shared by the Excel and PWA implementations.

Cases marked "client sample" come from File for Toyin.xlsx and are the
authority — if an implementation disagrees with one, the implementation is
wrong.
"""
import pytest

from excel_workflow.nursery_algos import (
    assign_splits,
    band_for_row,
    parity_for_row,
    run_direction,
    serpentine_by_range,
    serpentine_two_row_bands,
    spike_for_row,
)


# --- Task 1: spike pattern and run direction ---------------------------------

def test_spike_follows_period_four_cycle():
    rows = list(range(1, 13))
    assert [spike_for_row(r) for r in rows] == [1, 2, 2, 1, 1, 2, 2, 1, 1, 2, 2, 1]


def test_run_direction_flips_every_two_rows_starting_forward():
    rows = list(range(1, 9))
    assert [run_direction(r) for r in rows] == [
        "forward", "forward", "reverse", "reverse",
        "forward", "forward", "reverse", "reverse",
    ]


def test_row_one_is_spike_one_forward():
    assert spike_for_row(1) == 1
    assert run_direction(1) == "forward"


def test_spike_rejects_rows_below_one():
    with pytest.raises(ValueError, match="row must be >= 1"):
        spike_for_row(0)


def test_run_direction_rejects_rows_below_one():
    with pytest.raises(ValueError, match="row must be >= 1"):
        run_direction(0)


# --- Task 2: serpentine ordering ---------------------------------------------

def test_serpentine_by_range_reverses_row_order_on_even_ranges():
    plots = [(1, 1), (1, 2), (1, 3), (2, 1), (2, 2), (2, 3)]
    assert serpentine_by_range(plots) == [
        (1, 1), (1, 2), (1, 3),   # odd range -> rows ascending
        (2, 3), (2, 2), (2, 1),   # even range -> rows descending
    ]


def test_band_groups_two_adjacent_rows():
    assert band_for_row(2) == 1
    assert band_for_row(3) == 1
    assert band_for_row(26) == 13
    assert band_for_row(27) == 13


def test_parity_labels_odd_and_even_rows():
    assert parity_for_row(3) == "O"
    assert parity_for_row(2) == "E"


def test_serpentine_two_row_bands_matches_client_sample():
    # client sample: 'Date recording', Group 1 (rows 2 and 3), ranges 2-5.
    plots = [(r, row) for r in range(2, 6) for row in (2, 3)]
    assert serpentine_two_row_bands(plots) == [
        (2, 3), (2, 2),
        (3, 2), (3, 3),
        (4, 3), (4, 2),
        (5, 2), (5, 3),
    ]


def test_serpentine_two_row_bands_orders_bands_before_ranges():
    plots = [(2, 2), (2, 26), (3, 26), (3, 2)]
    assert serpentine_two_row_bands(plots) == [(2, 2), (3, 2), (2, 26), (3, 26)]


# --- Task 3: planting-date splits --------------------------------------------

def test_assign_splits_maps_each_row_to_its_planting_date():
    # client sample: 'Nursery data' 1st/2nd DOP row lists.
    dop_rows = [[1, 2, 5, 6], [3, 4, 7, 8]]
    assert assign_splits(dop_rows) == {
        1: 1, 2: 1, 5: 1, 6: 1,
        3: 2, 4: 2, 7: 2, 8: 2,
    }


def test_assign_splits_supports_three_planting_dates():
    assert assign_splits([[1], [2], [3]]) == {1: 1, 2: 2, 3: 3}


def test_assign_splits_rejects_a_row_in_two_planting_dates():
    with pytest.raises(ValueError, match="row 4 appears in splits 1 and 2"):
        assign_splits([[3, 4], [4, 5]])


def test_assign_splits_rejects_empty_input():
    with pytest.raises(ValueError, match="at least one planting date"):
        assign_splits([])
