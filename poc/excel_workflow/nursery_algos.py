"""Pure ordering and numbering rules for the nursery workbook.

These are the routines most at risk of drifting between the Excel (VBA) and
PWA (JS) implementations, so they live here with tests and are ported rather
than reinvented.

See docs/superpowers/specs/2026-08-13-nursery-workbook-v2-design.md
"""
from __future__ import annotations

# A two-cone planter's cones swap sides when it turns around, so spike numbers
# run 1,2 down the forward pass and 2,1 back up the reverse pass.
SPIKE_CYCLE = (1, 2, 2, 1)

FORWARD = "forward"
REVERSE = "reverse"


def spike_for_row(row: int) -> int:
    """Spike (planter cone) number for a field row. Rows are 1-based."""
    if row < 1:
        raise ValueError(f"row must be >= 1, got {row}")
    return SPIKE_CYCLE[(row - 1) % 4]


def run_direction(row: int) -> str:
    """Planting run direction for a field row. First two rows are forward."""
    if row < 1:
        raise ValueError(f"row must be >= 1, got {row}")
    return FORWARD if ((row - 1) // 2) % 2 == 0 else REVERSE


def band_for_row(row: int) -> int:
    """Two-row band index. Rows 2 and 3 are band 1; rows 26 and 27 are band 13."""
    return row // 2


def parity_for_row(row: int) -> str:
    """The 'O/E' helper column: O for odd rows, E for even."""
    return "O" if row % 2 else "E"


def _row_key(range_no: int, row_no: int) -> tuple[int, int]:
    """Serpentine within a range: rows descend on even ranges, ascend on odd.

    Negating the row on even ranges collapses the snake into a single sortable
    key, rather than needing a second pass per range.
    """
    return (range_no, -row_no if range_no % 2 == 0 else row_no)


def serpentine_by_range(plots) -> list[tuple[int, int]]:
    """Walk ranges ascending, snaking through the rows within each range."""
    return sorted(plots, key=lambda p: _row_key(p[0], p[1]))


def serpentine_two_row_bands(plots) -> list[tuple[int, int]]:
    """Walk each two-row band in turn, snaking by range within the band."""
    return sorted(plots, key=lambda p: (band_for_row(p[1]),) + _row_key(p[0], p[1]))


def assign_splits(dop_rows) -> dict[int, int]:
    """Map each field row to its 1-based planting-date split.

    A row belonging to two planting dates is a data-entry error, not something
    to resolve silently — the packets would be printed onto the wrong split.
    """
    dop_rows = list(dop_rows)
    if not dop_rows:
        raise ValueError("need at least one planting date")

    splits: dict[int, int] = {}
    for index, rows in enumerate(dop_rows, start=1):
        for row in rows:
            if row in splits:
                raise ValueError(
                    f"row {row} appears in splits {splits[row]} and {index}")
            splits[row] = index
    return splits


# -------------------------------------------------------------- packet prep
#
# Ported from PP2026_Run in the client's "Packetprinting VBA (advanced one)"
# document, which supersedes the rack-order note in the earlier change list.
#
# Packets come off the rack in the order the planter needs them, so the sheet
# is ordered before the rack numbers are handed out:
#
#   step 8  row descending; within a row the range descends on a forward row
#           and ascends on a reverse row
#   step 9  then grouped by spike, spike 1 before spike 2, keeping that order
#   step 10 rack order counts 1..n within each spike
#
# The spike rule in step 7 (row mod 4 -> 0,1 = spike 1; 2,3 = spike 2) is the
# same assignment SPIKE_CYCLE already makes, so spike_for_row is reused.


def packet_prep_order(plots) -> list[tuple[int, int]]:
    """Order ``(range, row)`` plots the way packets are racked."""
    def serpentine_key(plot):
        range_no, row_no = plot
        # Forward rows are 1,2,5,6,9,10... — the same rows run_direction()
        # calls forward, expressed as the VBA writes it.
        forward = row_no % 4 in (1, 2)
        return (-row_no, -range_no if forward else range_no)

    ordered = sorted(plots, key=serpentine_key)
    # Stable, so the serpentine order survives inside each spike.
    return sorted(ordered, key=lambda plot: spike_for_row(plot[1]))


def rack_digits(value, width: int = 4) -> list[str]:
    """Split a rack order into one column per digit: 7 -> ['0','0','0','7'].

    Four columns by default, as the document shows. A value too big for the
    width widens rather than being truncated — a clipped rack number would
    misfeed the rack silently.
    """
    text = str(value if value not in (None, "") else "").strip()
    if not text.isdigit():
        return [""] * width
    return list(text.zfill(max(width, len(text))))


# --------------------------------------------------------------------- dates
#
# Ported from GenerateS1S2TrendAnalysis in the client's appendix VBA.
#
# Selection columns hold a bare day of the month — the crew writes "28", not a
# full date — so the month has to be inferred. Recording starts on a known day
# and runs forward: a day at or after the start day belongs to the starting
# month, and anything lower has already wrapped into the next one.


def assign_month(day, start_day: int, start_month: int, start_year: int):
    """Resolve a bare day-of-month into ``(year, month, day)``.

    Returns ``None`` for blank, non-numeric or zero entries — an unrecorded
    selection, which must not be counted as a recording on any date.
    """
    try:
        day_no = int(str(day).strip())
    except (TypeError, ValueError):
        return None
    if day_no <= 0:
        return None

    if day_no >= start_day:
        month, year = start_month, start_year
    else:
        month, year = start_month + 1, start_year
        if month > 12:
            month, year = 1, year + 1
    return year, month, day_no


def trend_counts(rows, columns, start_day: int, start_month: int,
                 start_year: int):
    """Recordings per calendar date, for the selection trend chart.

    ``rows`` are dicts of selection column -> day value. Returns
    ``[((year, month, day), {column: count}), ...]`` over every date from the
    start date to the last one recorded, including days where nothing was
    recorded so the chart shows the gaps.
    """
    from datetime import date, timedelta

    counts: dict[tuple[int, int, int], dict[str, int]] = {}
    latest = (start_year, start_month, start_day)

    for row in rows:
        for column in columns:
            stamp = assign_month(row.get(column), start_day, start_month,
                                 start_year)
            if stamp is None:
                continue
            bucket = counts.setdefault(stamp, {})
            bucket[column] = bucket.get(column, 0) + 1
            latest = max(latest, stamp)

    cursor = date(start_year, start_month, start_day)
    last = date(*latest)
    out = []
    while cursor <= last:
        stamp = (cursor.year, cursor.month, cursor.day)
        here = counts.get(stamp, {})
        out.append((stamp, {c: here.get(c, 0) for c in columns}))
        cursor += timedelta(days=1)
    return out
