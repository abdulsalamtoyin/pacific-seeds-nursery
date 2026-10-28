// Downloading tabs as real Excel files.
//
// The workbook is built in the browser (xlsx-writer.js, on the vendored
// ExcelJS), so export works with no server behind the app — which is what
// lets it be hosted as static files on SharePoint. Four shapes are needed:
//
//   one tab      — the Export button on a grid
//   whole book   — every tab as its own sheet, one file
//   each tab     — the same tabs, one file apiece
//   Bartender    — the packet list, Source ID descending
//
// Sheets are described the same way in all four, so the writer only has one
// format to understand.

import { loadExcelJS } from "./excel-lib.js";
import { workbookBytes } from "./xlsx-writer.js";

const XLSX_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/** Strip characters Windows will not accept in a file name. */
function safeName(name) {
  return String(name ?? "").replace(/[\\/:*?"<>|]+/g, "-").trim()
    || "nursery-export";
}

/**
 * Turn a rendered grid into a sheet description.
 *
 * Takes what the user is actually looking at — filtered, sorted, with their
 * own columns and fills — rather than the underlying data, so the spreadsheet
 * matches the screen.
 */
export function sheetFromGrid(name, api, extra = {}) {
  const columns = api.columns();
  const rows = api.rows();
  return {
    name,
    headers: columns.map((c) => c.label ?? c.key),
    rows: rows.map((row) => columns.map((c) => row[c.key] ?? "")),
    colours: api.colours(),
    styles: api.styles(),
    ...extra,
  };
}

/** Build the workbook from these sheets and save it. */
export async function download(filename, sheets) {
  const ExcelJS = await loadExcelJS();
  const bytes = await workbookBytes(
    { filename: safeName(filename), sheets }, ExcelJS);

  const blob = new Blob([bytes], { type: XLSX_TYPE });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${safeName(filename)}.xlsx`;
  document.body.append(link);
  link.click();
  link.remove();
  // Revoking immediately can cancel the download in some browsers.
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

/** One tab, one file. The Export button on every grid. */
export async function exportGrid(nurseryCode, tabName, api, extra = {}) {
  await download(`${nurseryCode || "nursery"} - ${tabName}`,
    [sheetFromGrid(tabName, api, extra)]);
}

/** Every tab as its own sheet in one workbook. */
export async function exportWorkbook(nurseryCode, sheets) {
  await download(nurseryCode || "nursery-export", sheets);
}

/**
 * The same tabs, one file each.
 *
 * Sequential rather than parallel: browsers throttle simultaneous downloads,
 * and a dozen at once tends to leave some silently missing. Empty tabs are
 * skipped and named back to the caller rather than failing the whole run.
 */
export async function exportEachTab(nurseryCode, sheets) {
  const skipped = [];
  for (const sheet of sheets) {
    if (!sheet.rows.length) {
      skipped.push(sheet.name);
      continue;
    }
    await download(`${nurseryCode || "nursery"} - ${sheet.name}`, [sheet]);
  }
  return skipped;
}

/**
 * Packet list for Bartender.
 *
 * Always sorted by Source ID descending on the way out, whatever the tab is
 * showing — the label printer expects that order regardless of how the
 * operator happened to sort the screen.
 */
export async function exportBartender(nurseryCode, tabName, api) {
  const sheet = sheetFromGrid(tabName, api);
  const columns = api.columns();
  const sourceIndex = columns.findIndex(
    (c) => String(c.label ?? c.key).toLowerCase() === "source id");

  if (sourceIndex < 0) {
    throw new Error(
      `${tabName} has no Source ID column, so it cannot be sorted for ` +
      "Bartender. Un-hide the column and try again.");
  }

  sheet.rows = [...sheet.rows].sort((a, b) =>
    String(b[sourceIndex] ?? "").localeCompare(
      String(a[sourceIndex] ?? ""), undefined,
      { numeric: true, sensitivity: "base" }));

  // Cell fills are positional, and the rows just moved — carrying them over
  // would colour the wrong packets, so they are dropped for this export.
  sheet.colours = {};
  sheet.styles = {};

  await download(`${nurseryCode || "nursery"} - ${tabName} - Bartender`, [sheet]);
}
