// Grid logic with no DOM in it.
//
// Sorting, filtering and — most importantly — working out what a cell's value
// actually is once hand-edits are layered over computed data. That last part
// is where work gets lost if it is wrong, so it lives here where it can be
// tested under node instead of inside a rendering function.

const NUMERIC = /^-?\d+(\.\d+)?$/;

/**
 * Order two cell values.
 *
 * Numbers sort as numbers, so row 2 comes before row 10 rather than after it.
 * Blanks sort last, because an empty plot at the top of a fieldbook is never
 * what the user was looking for.
 */
export function compare(a, b) {
  const sa = String(a ?? "");
  const sb = String(b ?? "");
  if (!sa && !sb) return 0;
  if (!sa) return 1;
  if (!sb) return -1;
  if (NUMERIC.test(sa) && NUMERIC.test(sb)) return Number(sa) - Number(sb);
  return sa.localeCompare(sb, undefined, { numeric: true, sensitivity: "base" });
}

/**
 * Does this record read its values through the hand-edit overlay?
 *
 * Rows the user typed themselves, and rows on a tab that owns its data, hold
 * their own values. Only computed rows — rebuilt from PRISM on every render —
 * need an overlay, because there is no durable object to write into.
 */
export function isOverlaid(rec, owned) {
  return !owned && !rec.userAdded;
}

/** The value a cell should show: the hand-edit if there is one, else source. */
export function effectiveValue(rec, key, { owned = false, edits = {} } = {}) {
  if (isOverlaid(rec, owned)) {
    const patch = edits[rec.key];
    if (patch && key in patch) return patch[key];
  }
  return rec.src?.[key] ?? "";
}

/** Today as YYYY-MM-DD — what an <input type="date"> expects. */
export function todayISO(now = new Date()) {
  const pad = (n) => String(n).padStart(2, "0");
  // Deliberately local, not toISOString(): a nursery recording an evening date
  // in Australia must not have UTC roll it back to yesterday.
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
}

/**
 * A blank date cell reads as today.
 *
 * Dates are picked from a calendar and default to the current date. Applying
 * that here rather than in the renderer means the value shown, filtered,
 * sorted and exported are all the same value.
 */
export function defaultedValue(value, column, today) {
  if (value !== "" && value !== null && value !== undefined) return value;
  return column?.type === "date" ? today : (value ?? "");
}

/** True when this cell carries a hand-edit over a computed value. */
export function isEditedCell(rec, key, { owned = false, edits = {} } = {}) {
  return isOverlaid(rec, owned) && !!(edits[rec.key] && key in edits[rec.key]);
}

/**
 * Filters are per column: a list of the values allowed through. A column with
 * no entry is unfiltered — that is what lets "everything ticked" mean "off".
 */
export function filterRecords(records, filters, valueOf) {
  const active = Object.entries(filters ?? {}).filter(([, v]) => Array.isArray(v));
  if (!active.length) return [...records];
  return records.filter((rec) => active.every(([key, allowed]) =>
    allowed.includes(String(valueOf(rec, key) ?? ""))));
}

/** Sort a copy; `sort` of null means leave the natural order alone. */
export function sortRecords(records, sort, valueOf) {
  const out = [...records];
  if (!sort) return out;
  const { key, dir } = sort;
  out.sort((a, b) => {
    const r = compare(valueOf(a, key), valueOf(b, key));
    return dir === "desc" ? -r : r;
  });
  return out;
}

/** Header click order: unsorted -> ascending -> descending -> unsorted. */
export function nextSort(current, key) {
  if (!current || current.key !== key) return { key, dir: "asc" };
  if (current.dir === "asc") return { key, dir: "desc" };
  return null;
}

/**
 * Hand-edits whose row is no longer present.
 *
 * Re-importing a corrected PRISM export can drop a plot the user had already
 * annotated. Deleting those edits silently would lose the work, so they are
 * kept and surfaced instead.
 */
export function orphanKeys(edits, liveKeys) {
  const live = new Set(liveKeys);
  return Object.keys(edits ?? {}).filter((k) => !live.has(k));
}

/**
 * Rack order is printed onto packets and read back by a barcode scanner, so
 * every value in a tab must be the same width or the sort breaks. Three digits
 * normally; four once any spike passes 999.
 */
export function radixWidth(largestCount) {
  return String(Math.max(0, largestCount)).length <= 3 ? 3 : 4;
}

/** Zero-pad to `width`, leaving anything non-numeric alone. */
export function padRack(value, width) {
  const s = String(value ?? "");
  if (!NUMERIC.test(s)) return s;
  return s.padStart(width, "0");
}
