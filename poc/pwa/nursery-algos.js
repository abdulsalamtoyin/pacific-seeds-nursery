// Ordering and numbering rules for the nursery workbook.
//
// This is a deliberate port of excel_workflow/nursery_algos.py. The desktop
// app and the Excel workbook are meant to behave identically, so these rules
// exist three times — here, in Python, and in VBA — and are pinned by the
// same cases. tests/test_js_parity.py runs this file under node and compares
// it against the Python original; if they diverge, that test fails.
//
// Do not "improve" a rule here without changing the other two.

// A two-cone planter's cones swap sides when it turns around, so spike
// numbers run 1,2 down the forward pass and 2,1 back up the reverse pass.
export const SPIKE_CYCLE = [1, 2, 2, 1];

export const FORWARD = "forward";
export const REVERSE = "reverse";

export function spikeForRow(row) {
  if (row < 1) throw new RangeError(`row must be >= 1, got ${row}`);
  return SPIKE_CYCLE[(row - 1) % 4];
}

export function runDirection(row) {
  if (row < 1) throw new RangeError(`row must be >= 1, got ${row}`);
  return Math.floor((row - 1) / 2) % 2 === 0 ? FORWARD : REVERSE;
}

export function bandForRow(row) {
  // Rows 2 and 3 are band 1; rows 26 and 27 are band 13.
  return Math.floor(row / 2);
}

export function parityForRow(row) {
  return row % 2 ? "O" : "E";
}

// Serpentine within a range: rows descend on even ranges, ascend on odd.
// Negating the row on even ranges collapses the snake into one sortable key.
function rowKey(rangeNo, rowNo) {
  return rangeNo % 2 === 0 ? -rowNo : rowNo;
}

function byKeys(keysOf) {
  return (a, b) => {
    const ka = keysOf(a);
    const kb = keysOf(b);
    for (let i = 0; i < ka.length; i++) {
      if (ka[i] !== kb[i]) return ka[i] - kb[i];
    }
    return 0;
  };
}

export function serpentineByRange(plots) {
  return [...plots].sort(byKeys(([r, w]) => [r, rowKey(r, w)]));
}

export function serpentineTwoRowBands(plots) {
  return [...plots].sort(byKeys(([r, w]) => [bandForRow(w), r, rowKey(r, w)]));
}

export function assignSplits(dopRows) {
  // A row belonging to two planting dates is a data-entry error, not something
  // to resolve silently — the packets would be printed onto the wrong split.
  if (!dopRows || dopRows.length === 0) {
    throw new RangeError("need at least one planting date");
  }
  const splits = new Map();
  dopRows.forEach((rows, i) => {
    const index = i + 1;
    for (const row of rows) {
      if (splits.has(row)) {
        throw new RangeError(
          `row ${row} appears in splits ${splits.get(row)} and ${index}`);
      }
      splits.set(row, index);
    }
  });
  return splits;
}

// ----------------------------------------------------------- packet prep
//
// Ported from PP2026_Run in the client's "Packetprinting VBA (advanced one)"
// document, and mirrored by packet_prep_order() / rack_digits() in
// nursery_algos.py.
//
// Packets come off the rack in the order the planter needs them, so the sheet
// is ordered before the rack numbers are handed out: row descending, range
// descending on a forward row and ascending on a reverse one, then grouped by
// spike with spike 1 first, then numbered 1..n within each spike.

/** Order [range, row] plots the way packets are racked. */
export function packetPrepOrder(plots) {
  const serpentineKey = ([rangeNo, rowNo]) => {
    // Forward rows are 1,2,5,6,9,10... — the same rows runDirection() calls
    // forward, expressed as the VBA writes it.
    const forward = rowNo % 4 === 1 || rowNo % 4 === 2;
    return [-rowNo, forward ? -rangeNo : rangeNo];
  };

  const ordered = [...plots].sort((a, b) => {
    const ka = serpentineKey(a);
    const kb = serpentineKey(b);
    return ka[0] - kb[0] || ka[1] - kb[1];
  });
  // Array.prototype.sort is stable, so the serpentine order survives inside
  // each spike.
  return ordered.sort((a, b) => spikeForRow(a[1]) - spikeForRow(b[1]));
}

/**
 * Split a rack order into one column per digit: 7 -> ["0","0","0","7"].
 *
 * Four columns by default, as the document shows. A value too big for the
 * width widens rather than being truncated — a clipped rack number would
 * misfeed the rack silently.
 */
export function rackDigits(value, width = 4) {
  const text = String(value ?? "").trim();
  if (!/^\d+$/.test(text)) return Array(width).fill("");
  return text.padStart(Math.max(width, text.length), "0").split("");
}

// ----------------------------------------------------------------- dates
//
// Ported from GenerateS1S2TrendAnalysis in the client's appendix VBA, and
// mirrored by assign_month() / trend_counts() in nursery_algos.py.
//
// Selection columns hold a bare day of the month — the crew writes "28", not
// a full date — so the month has to be inferred. Recording starts on a known
// day and runs forward: a day at or after the start day belongs to the
// starting month, and anything lower has already wrapped into the next one.

/**
 * Resolve a bare day-of-month into [year, month, day].
 *
 * Returns null for blank, non-numeric or zero entries — an unrecorded
 * selection, which must not be counted as a recording on any date.
 */
export function assignMonth(day, startDay, startMonth, startYear) {
  const text = String(day ?? "").trim();
  if (!/^-?\d+$/.test(text)) return null;
  const dayNo = Number(text);
  if (dayNo <= 0) return null;

  let month = startMonth;
  let year = startYear;
  if (dayNo < startDay) {
    month = startMonth + 1;
    if (month > 12) {
      month = 1;
      year += 1;
    }
  }
  return [year, month, dayNo];
}

/**
 * Recordings per calendar date, for the selection trend chart.
 *
 * Includes days where nothing was recorded, so the chart shows the gaps
 * rather than joining across them. UTC throughout — these are calendar
 * labels, not instants, and local DST would shift the day boundaries.
 */
export function trendCounts(rows, columns, startDay, startMonth, startYear) {
  const counts = new Map();
  const first = Date.UTC(startYear, startMonth - 1, startDay);
  let latest = first;

  for (const row of rows) {
    for (const column of columns) {
      const stamp = assignMonth(row[column], startDay, startMonth, startYear);
      if (!stamp) continue;
      const key = Date.UTC(stamp[0], stamp[1] - 1, stamp[2]);
      if (!counts.has(key)) counts.set(key, {});
      const bucket = counts.get(key);
      bucket[column] = (bucket[column] ?? 0) + 1;
      if (key > latest) latest = key;
    }
  }

  const DAY = 86400000;
  const out = [];
  for (let t = first; t <= latest; t += DAY) {
    const d = new Date(t);
    const here = counts.get(t) ?? {};
    out.push([
      [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()],
      Object.fromEntries(columns.map((c) => [c, here[c] ?? 0])),
    ]);
  }
  return out;
}
