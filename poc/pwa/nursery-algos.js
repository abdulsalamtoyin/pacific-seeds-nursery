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
