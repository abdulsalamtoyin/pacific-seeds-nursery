// Loading the spreadsheet library on first use.
//
// ExcelJS is vendored rather than fetched from a CDN so the app works offline
// in the Windows build and on a SharePoint page whose script policy blocks
// outside hosts. It is close to 1 MB, so it is only pulled in the first time
// someone imports or exports — opening a tab never pays for it.
//
// Tests under node set globalThis.ExcelJS before importing, which skips the
// script tag entirely.

let pending = null;

export function loadExcelJS() {
  if (globalThis.ExcelJS) return Promise.resolve(globalThis.ExcelJS);
  if (pending) return pending;

  pending = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = new URL("./vendor/exceljs.min.js", import.meta.url).href;
    script.onload = () => resolve(globalThis.ExcelJS);
    script.onerror = () => {
      pending = null;               // let a later click try again
      reject(new Error(
        "The spreadsheet library (vendor/exceljs.min.js) could not be " +
        "loaded. Check it was copied alongside the app."));
    };
    document.head.append(script);
  });
  return pending;
}
