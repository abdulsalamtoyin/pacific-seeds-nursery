// Read a PRISM export in the browser and hand back its rows.
//
// The browser twin of POST /parse/prism in backend/app.py, so importing works
// with no Python behind the app. Accepts .xlsx/.xlsm or .csv/.tsv. PRISM
// downloads arrive with the data on a sheet called 'Sheet 1', and the header
// row sits at row 1 or row 5 depending on how it was saved, so both are found
// by scanning rather than assumed. tests/test_prism_reader.py feeds the same
// files to both readers and requires identical output.

const DELIMITED = [".csv", ".tsv", ".txt"];
const HEADER_SCAN_ROWS = 15;

const pad = (n) => String(n).padStart(2, "0");

/**
 * One spreadsheet cell as the text the app stores.
 *
 * ExcelJS hands back objects for formulas, rich text and hyperlinks; the
 * Python reader opened the file with data_only, so formulas give their cached
 * result and everything else its visible text.
 */
export function cleanCell(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) {
    // Spreadsheet dates carry no zone; ExcelJS builds them as UTC.
    return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
  }
  if (typeof v === "boolean") return v ? "True" : "False";
  if (typeof v === "object") {
    if ("result" in v) return cleanCell(v.result);
    if (Array.isArray(v.richText)) {
      return v.richText.map((part) => part.text ?? "").join("").trim();
    }
    if ("text" in v) return cleanCell(v.text);
    if ("error" in v) return String(v.error).trim();
    return "";
  }
  return String(v).trim();
}

/** Split CSV/TSV text into rows, honouring quoted fields and "" escapes. */
export function parseDelimited(text, delimiter) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false;
      } else {
        field += ch;
      }
    } else if (ch === '"' && field === "") {
      quoted = true;
    } else if (ch === delimiter) {
      row.push(field);
      field = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += ch;
    }
  }
  if (field !== "" || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/** Text file -> grid of cleaned cells. */
export function gridFromText(text, fileName) {
  const body = text.replace(/^﻿/, "");
  const firstLine = body.split("\n")[0];
  const delimiter = (fileName.toLowerCase().endsWith(".tsv")
    || firstLine.includes("\t")) ? "\t" : ",";
  return parseDelimited(body, delimiter).map((r) => r.map(cleanCell));
}

/** Worksheet -> grid of cleaned cells, every row padded to the sheet width. */
export function gridFromWorksheet(ws) {
  const grid = [];
  const width = ws.columnCount;
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const cells = [];
    for (let c = 1; c <= width; c++) {
      const cell = row.getCell(c);
      // ExcelJS repeats a merged range's value into every cell it covers;
      // Excel itself (and the Python reader) only has it in the first.
      const covered = cell.isMerged && cell.master.address !== cell.address;
      cells.push(covered ? "" : cleanCell(cell.value));
    }
    grid.push(cells);
  }
  return grid;
}

/** Find the header row and turn what follows into one object per packet. */
export function rowsFromGrid(grid, sheetName) {
  const headerIdx = grid.slice(0, HEADER_SCAN_ROWS).findIndex((row) => {
    const lowered = row.map((c) => c.toLowerCase());
    return lowered.includes("range") && lowered.includes("row");
  });
  if (headerIdx < 0) {
    throw new Error(
      "No header row with 'Range' and 'Row' found in the first 15 rows. " +
      "Check the right sheet was picked.");
  }

  const headers = grid[headerIdx];
  const rows = [];
  for (const rawRow of grid.slice(headerIdx + 1)) {
    if (!rawRow.some((c) => c)) continue;
    const rec = {};
    headers.forEach((h, i) => {
      if (h) rec[h] = i < rawRow.length ? rawRow[i] : "";
    });
    // A row with neither Range nor Row is a footer or a stray note.
    if (!(rec.Range || rec.Row)) continue;
    rows.push(rec);
  }

  return {
    sheet: sheetName,
    headers: headers.filter((h) => h),
    rows,
    count: rows.length,
  };
}

/**
 * Read a File (from an <input type=file>) as a PRISM export.
 *
 * @returns {Promise<{sheet, headers, rows, count}>}
 */
export async function readPrism(file, ExcelJS = globalThis.ExcelJS) {
  const name = String(file.name || "");
  const lowered = name.toLowerCase();

  if (DELIMITED.some((ext) => lowered.endsWith(ext))) {
    return rowsFromGrid(gridFromText(await file.text(), name), name || "csv");
  }

  let ws;
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    ws = wb.worksheets[0];
    if (!ws) throw new Error("the workbook has no sheets");
  } catch (err) {
    throw new Error(`Could not read file: ${err.message || err}`);
  }
  return rowsFromGrid(gridFromWorksheet(ws), ws.name);
}
