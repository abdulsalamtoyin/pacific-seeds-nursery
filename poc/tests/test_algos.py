"""Ordering and numbering rules shared by the Excel and PWA implementations.

Cases marked "client sample" come from File for Toyin.xlsx and are the
authority — if an implementation disagrees with one, the implementation is
wrong.
"""
import pytest

from excel_workflow.nursery_algos import (
    assign_splits,
    band_for_row,
    packet_prep_order,
    parity_for_row,
    rack_digits,
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


# --- Packet printing (advanced): PP2026_Run --------------------------------
#
# From "Packetprinting VBA (advanced one).docx". Packets have to come off the
# rack in the order the planter needs them, so the sheet is ordered before the
# rack numbers are handed out.


def test_spike_matches_the_documents_row_mod_4_rule():
    """Step 7 states the spikes as row lists; they must agree with our cycle."""
    spike1 = [r for r in range(1, 17) if spike_for_row(r) == 1]
    spike2 = [r for r in range(1, 17) if spike_for_row(r) == 2]
    assert spike1 == [1, 4, 5, 8, 9, 12, 13, 16]
    assert spike2 == [2, 3, 6, 7, 10, 11, 14, 15]


def test_forward_rows_are_the_ones_the_document_lists():
    """Step 8 calls 1,2,5,6,9,10... forward. run_direction must agree."""
    forward = [r for r in range(1, 17) if run_direction(r) == "forward"]
    assert forward == [1, 2, 5, 6, 9, 10, 13, 14]


def test_packet_order_groups_spike_one_before_spike_two():
    plots = [(r, w) for w in range(1, 5) for r in range(1, 4)]
    order = packet_prep_order(plots)
    spikes = [spike_for_row(row) for _, row in order]
    assert spikes == sorted(spikes), "spike 1 must come out before spike 2"


def test_rows_descend_within_a_spike():
    plots = [(1, w) for w in (1, 4, 5, 8)]        # all spike 1
    assert [row for _, row in packet_prep_order(plots)] == [8, 5, 4, 1]


def test_range_descends_on_a_forward_row():
    """Row 1 is forward, so its ranges come out high to low."""
    plots = [(1, 1), (2, 1), (3, 1)]
    assert packet_prep_order(plots) == [(3, 1), (2, 1), (1, 1)]


def test_range_ascends_on_a_reverse_row():
    """Row 3 is reverse, so its ranges come out low to high."""
    plots = [(3, 3), (1, 3), (2, 3)]
    assert packet_prep_order(plots) == [(1, 3), (2, 3), (3, 3)]


def test_the_snake_turns_between_rows():
    """Forward row 5 descends, reverse row 4 ascends — the serpentine."""
    plots = [(1, 5), (2, 5), (1, 4), (2, 4)]
    order = packet_prep_order(plots)
    # Both rows are spike 1; row 5 comes first because rows descend.
    assert order == [(2, 5), (1, 5), (1, 4), (2, 4)]


def test_every_plot_survives_the_ordering():
    plots = [(r, w) for w in range(1, 9) for r in range(1, 5)]
    assert sorted(packet_prep_order(plots)) == sorted(plots)


@pytest.mark.parametrize("value,expected", [
    (7, ["0", "0", "0", "7"]),
    (85, ["0", "0", "8", "5"]),
    (326, ["0", "3", "2", "6"]),
    (1458, ["1", "4", "5", "8"]),
])
def test_rack_order_splits_into_the_documents_four_digits(value, expected):
    assert rack_digits(value) == expected


def test_a_rack_order_beyond_four_digits_widens_rather_than_truncating():
    """Clipping a rack number would misfeed the rack with no sign of why."""
    assert rack_digits(12345) == ["1", "2", "3", "4", "5"]


@pytest.mark.parametrize("value", ["", None, "n/a"])
def test_a_missing_rack_order_gives_blank_digits(value):
    assert rack_digits(value) == ["", "", "", ""]
