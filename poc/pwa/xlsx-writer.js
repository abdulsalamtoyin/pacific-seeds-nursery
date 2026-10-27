// Turn a grid from the app into a real .xlsx file, in the browser.
//
// This is the browser twin of backend/xlsx_export.py, so the app can run as
// plain static files (SharePoint, a static host) with no Python behind it.
// The layout rules are the same: an optional merged band row, bold headers,
// fills and text styles addressed by data position, column widths, Excel's own
// filter dropdowns, and frozen headers. tests/test_export.py runs every case
// against both writers to keep them identical.

// Excel forbids these in a sheet name, caps it at 31 characters, and will not
// accept an empty one. "History" is reserved.
const ILLEGAL = /[[\]:*?/\\]/g;
const MAX_NAME = 31;
const RESERVED = new Set(["history"]);

/** A sheet name Excel will accept, unique against `taken`. */
export function sanitiseSheetName(name, taken = null) {
  let cleaned = String(name ?? "").replace(ILLEGAL, "-").trim()
    .replace(/^'+|'+$/g, "");
  if (!cleaned) cleaned = "Sheet";
  if (RESERVED.has(cleaned.toLowerCase())) cleaned = `${cleaned}_`;
  cleaned = cleaned.slice(0, MAX_NAME);

  if (taken === null) return cleaned;

  // Uniquify by suffixing " (2)", trimming the stem so the result still fits.
  const lowered = new Set([...taken].map((t) => t.toLowerCase()));
  let candidate = cleaned;
  for (let n = 2; lowered.has(candidate.toLowerCase()); n++) {
    const suffix = ` (${n})`;
    candidate = cleaned.slice(0, MAX_NAME - suffix.length) + suffix;
  }
  return candidate;
}

/** '#ffd9d9' or 'ffd9d9' -> 'FFFFD9D9' (Excel wants aRGB). */
export function argb(colour) {
  if (!colour) return null;
  let text = String(colour).replace(/^#+/, "").trim();
  if (text.length === 3) text = [...text].map((c) => c + c).join("");
  if (text.length === 6) text = `FF${text}`;
  if (!/^[0-9A-Fa-f]{8}$/.test(text)) return null;
  return text.toUpperCase();
}

/** '2,3' -> [2, 3]; anything malformed is ignored rather than fatal. */
function cellRef(key) {
  const parts = String(key).split(",");
  if (parts.length !== 2) return null;
  const [row, col] = parts.map((p) => Number(p.trim()));
  if (!Number.isInteger(row) || !Number.isInteger(col)) return null;
  if (row < 0 || col < 0) return null;
  return [row, col];
}

/** 1 -> A, 27 -> AA. */
function columnLetter(n) {
  let s = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  }
  return s;
}

function writeSheet(ws, sheet) {
  const headers = sheet.headers || [];
  const rows = sheet.rows || [];
  const bands = sheet.bands || [];
  const colours = sheet.colours || {};
  const styles = sheet.styles || {};
  const widths = sheet.widths || [];

  let headerRow = 1;
  if (bands.length) {
    // A band spans several columns above the real headers, as in the
    // client's format: Original entry | Replaced entry | Reason.
    let col = 1;
    for (const [label, rawSpan] of bands) {
      const span = Math.max(1, Math.trunc(Number(rawSpan)) || 1);
      const cell = ws.getCell(1, col);
      cell.value = label;
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center" };
      if (span > 1) ws.mergeCells(1, col, 1, col + span - 1);
      col += span;
    }
    headerRow = 2;
  }

  headers.forEach((head, i) => {
    const cell = ws.getCell(headerRow, i + 1);
    cell.value = head;
    cell.font = { bold: true };
  });

  const firstDataRow = headerRow + 1;
  rows.forEach((row, r) => {
    row.forEach((value, c) => {
      ws.getCell(firstDataRow + r, c + 1).value = value ?? null;
    });
  });

  // Fills and text styles are addressed by the position the user saw them in,
  // zero-based over the data rows only.
  for (const [key, colour] of Object.entries(colours)) {
    const ref = cellRef(key);
    const rgb = argb(colour);
    if (!ref || !rgb) continue;
    ws.getCell(firstDataRow + ref[0], ref[1] + 1).fill = {
      type: "pattern", pattern: "solid", fgColor: { argb: rgb },
    };
  }

  for (const [key, style] of Object.entries(styles)) {
    const ref = cellRef(key);
    if (!ref || !style) continue;
    const cell = ws.getCell(firstDataRow + ref[0], ref[1] + 1);
    const font = { bold: Boolean(style.bold), italic: Boolean(style.italic) };
    const size = Number(style.size);
    if (style.size && Number.isFinite(size)) font.size = size;
    const colour = argb(style.colour);
    if (colour) font.color = { argb: colour };
    cell.font = font;
    if (["left", "center", "right"].includes(style.align)) {
      cell.alignment = { horizontal: style.align };
    }
  }

  widths.forEach((width, i) => {
    const w = Number(width);
    if (width !== null && width !== "" && Number.isFinite(w)) {
      ws.getColumn(i + 1).width = w;
    }
  });

  // Excel's own filter dropdowns, so the exported sheet behaves like the tab.
  if (headers.length && rows.length) {
    const last = columnLetter(headers.length);
    ws.autoFilter = `A${headerRow}:${last}${headerRow + rows.length}`;
  }
  ws.views = [{
    state: "frozen", xSplit: 0, ySplit: headerRow,
    topLeftCell: `A${firstDataRow}`,
  }];
}

/** Build an ExcelJS workbook from `{filename, sheets: [...]}`. */
export function buildWorkbook(payload, ExcelJS = globalThis.ExcelJS) {
  const sheets = payload?.sheets || [];
  if (!sheets.length) {
    throw new Error("Nothing to export — no sheets were supplied.");
  }

  // An export that would produce a file of nothing but headers is almost
  // always a mistake, so say so rather than hand over an empty spreadsheet.
  if (sheets.every((s) => !(s.rows || []).length)) {
    const names = sheets.map((s) => String(s.name ?? "?")).join(", ");
    throw new Error(`Nothing to export — ${names} has no rows yet.`);
  }

  const wb = new ExcelJS.Workbook();
  const taken = new Set();
  for (const sheet of sheets) {
    const name = sanitiseSheetName(sheet.name, taken);
    taken.add(name);
    writeSheet(wb.addWorksheet(name), sheet);
  }
  return wb;
}

/** The finished file as bytes, ready for a Blob or a disk write. */
export async function workbookBytes(payload, ExcelJS = globalThis.ExcelJS) {
  return buildWorkbook(payload, ExcelJS).xlsx.writeBuffer();
}
