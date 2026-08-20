// Nursery Workbook — desktop app.
//
// This mirrors the Excel workbook: the same tabs in the same order, driven by
// the same nursery_spec.json, using the same ordering rules. Where the
// workbook has a macro button, this has a button that does the same thing.
//
// Tab list comes from nursery-spec.js and the rules from nursery-algos.js —
// both generated/ported from the Python, and pinned by tests/test_js_parity.py.

import { SPEC, tabsFor } from "./nursery-spec.js";
import {
  spikeForRow, runDirection, bandForRow, parityForRow,
  serpentineByRange, serpentineTwoRowBands, assignSplits,
  assignMonth, trendCounts, packetPrepOrder, rackDigits,
} from "./nursery-algos.js";
import * as store from "./store.js";
import { grid } from "./grid.js";
import { todayISO } from "./grid-core.js";
import {
  download, exportBartender, exportEachTab, exportGrid, exportWorkbook,
} from "./exporter.js";

const $ = (sel) => document.querySelector(sel);

const STATUS_DEFAULT = "Waiting decision";
const STATUS_OPTIONS = [STATUS_DEFAULT, "Changes made in PRISM"];
const STAGE_OPTIONS = ["Packeting", "Planting"];

const QR_ORIGINAL = "QR code (original entry)";
const QR_REPLACED = "QR code (replaced entry)";

// Timestamp, Material ID, Source ID, Technician and Split no. were dropped from
// these tabs on request. The material fields are not lost — the export expands
// each QR back into the full nine-column block the client's format wants, so
// the tab stays short while the spreadsheet stays complete.
const REPLACEMENT_COLUMNS = [
  { key: "Stage", type: "select", options: STAGE_OPTIONS },
  { key: "Plot", type: "text" },
  { key: QR_ORIGINAL, type: "text", wide: true },
  { key: QR_REPLACED, type: "text", wide: true },
  { key: "Reason", type: "text" },
  { key: "Status", type: "select", options: STATUS_OPTIONS },
];

const PLANTING_ERROR_COLUMNS =
  REPLACEMENT_COLUMNS.filter((c) => c.key !== QR_REPLACED);

// Paperbags & labels was added above Seedling and Post Harvest removed, both
// on request. The stage column's own heading is left blank, as asked.
const GROWTH_STAGES = [
  ["Paperbags & labels", "#d9d2e9"],
  ["Seedling", "#fec000"], ["Vegetative", "#a9d08e"], ["Heading", "#ff6b6b"],
  ["Flowering", "#add8e6"], ["Grain filling", "#ffff99"],
  ["Harvest", "#ccffcc"],
];

const STAGE_COLOURS = new Map(GROWTH_STAGES);

// ---------------------------------------------------------------- state
//
// `state` is the *active* nursery, held in a rebindable binding so switching
// nurseries is a reassignment plus a render. Views read `state.x` as before;
// store.js owns the collection, the active pointer and persistence.

let state = store.activeNursery();
let activeTab = "Home";

function save() {
  store.save();
}

/** Point `state` at whichever nursery is active now, then redraw. */
function reloadActive() {
  state = store.activeNursery();
  activeTab = "Home";
  render();
}

function isType(t) {
  return state.types.some((x) => x.toLowerCase() === t.toLowerCase());
}

// ---------------------------------------------------------------- helpers

function el(tag, attrs = {}, ...kids) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") node.className = v;
    else if (k.startsWith("on")) node.addEventListener(k.slice(2), v);
    else if (v !== null && v !== undefined) node.setAttribute(k, v);
  }
  for (const kid of kids.flat()) {
    if (kid === null || kid === undefined || kid === false) continue;
    node.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
  }
  return node;
}

function table(headers, rows, opts = {}) {
  const thead = el("tr", {}, headers.map((h) => el("th", {}, h)));
  const body = rows.map((cells, i) =>
    el("tr", {}, cells.map((c, j) => {
      const cls = opts.cellClass ? opts.cellClass(i, j, c) : "";
      return el("td", { class: cls }, c === null || c === undefined ? "" : c);
    })));
  return el("div", { class: "scroll" },
    el("table", {}, el("thead", {}, thead), el("tbody", {}, body)));
}

function pageHead(title, subtitle) {
  return [
    title ? el("h1", {}, title) : null,
    subtitle ? el("p", { class: "sub" }, subtitle) : null,
  ].filter(Boolean);
}

function emptyState(msg) {
  return el("div", { class: "empty" }, msg);
}

// Tabs register how to build their sheet here, so "export the whole book" can
// produce every tab without first rendering each one. A grid-backed view fills
// this in as it renders; tabs not yet converted are simply absent.
const SHEETS = new Map();

function registerSheet(tabName, build) {
  SHEETS.set(tabName, build);
}

/** A rendered grid as a sheet description, for the whole-book export. */
function sheetOf(name, api, extra = {}) {
  const columns = api.columns();
  return {
    name,
    headers: columns.map((c) => c.label ?? c.key),
    rows: api.rows().map((r) => columns.map((c) => r[c.key] ?? "")),
    colours: api.colours(),
    styles: api.styles(),
    ...extra,
  };
}

/** Build every registered tab's sheet, in workbook tab order. */
function allSheets() {
  return visibleTabs()
    .filter((t) => SHEETS.has(t))
    .map((t) => SHEETS.get(t)());
}

/** Wrap an export so a failure explains itself rather than doing nothing. */
async function runExport(label, fn) {
  try {
    await fn();
  } catch (e) {
    alert(`${label} failed.\n\n${e.message}`);
  }
}

/**
 * Choose rows from the rows that actually exist, rather than typing a list.
 *
 * Rows already claimed by an earlier planting date are shown disabled — a row
 * in two planting dates would print packets onto the wrong split, and catching
 * it here is friendlier than failing after the whole wizard has been filled in.
 *
 * @returns {Promise<number[]|null>} chosen rows, or null if cancelled.
 */
function pickRows(label, allRows, taken) {
  return new Promise((resolve) => {
    const chosen = new Set();
    const close = (value) => { backdrop.remove(); resolve(value); };

    const boxes = allRows.map((row) => {
      const claimed = taken.has(row);
      const box = el("input", { type: "checkbox" });
      box.disabled = claimed;
      box.addEventListener("change", () => {
        if (box.checked) chosen.add(row); else chosen.delete(row);
      });
      return el("label", {
        class: "filter-row",
        title: claimed ? "Already assigned to an earlier planting date" : "",
        style: claimed ? "opacity:.45" : "",
      }, box, el("span", {}, `Row ${row}`));
    });

    const backdrop = el("div", { class: "modal-backdrop" },
      el("div", { class: "modal" },
        el("h3", {}, label),
        el("div", { class: "filter-list" }, boxes),
        el("div", { class: "filter-actions" },
          el("button", {
            class: "action ghost",
            onclick: () => {
              boxes.forEach((l, i) => {
                const box = l.querySelector("input");
                if (box.disabled) return;
                box.checked = true;
                chosen.add(allRows[i]);
              });
            },
          }, "All free rows"),
          el("button", {
            class: "action ghost",
            onclick: () => {
              boxes.forEach((l) => { l.querySelector("input").checked = false; });
              chosen.clear();
            },
          }, "None")),
        el("div", { class: "btnrow", style: "margin:14px 0 0" },
          el("button", {
            class: "action",
            onclick: () => {
              if (!chosen.size) {
                alert("Pick at least one row.");
                return;
              }
              close([...chosen].sort((a, b) => a - b));
            },
          }, "OK"),
          el("button", { class: "action ghost", onclick: () => close(null) },
            "Cancel"))));

    backdrop.addEventListener("mousedown", (e) => {
      if (e.target === backdrop) close(null);
    });
    document.body.append(backdrop);
  });
}

/**
 * Ask for a date with a calendar, not a typed string.
 *
 * A typed date could arrive as 3/4/26, 03-Apr-26 or 2026-04-03, and the tabs
 * that group by planting date would treat those as three different dates. The
 * picker removes the ambiguity, and opens on today.
 *
 * @returns {Promise<string|null>} ISO YYYY-MM-DD, or null if cancelled.
 */
function pickDate(label, initial) {
  return new Promise((resolve) => {
    const input = el("input", {
      class: "f",
      type: "date",
      value: initial || todayISO(),
    });

    const close = (value) => { backdrop.remove(); resolve(value); };

    const backdrop = el("div", { class: "modal-backdrop" },
      el("div", { class: "modal" },
        el("h3", {}, label),
        input,
        el("div", { class: "btnrow", style: "margin:14px 0 0" },
          el("button", {
            class: "action",
            onclick: () => close(input.value || null),
          }, "OK"),
          el("button", {
            class: "action ghost",
            onclick: () => close(null),
          }, "Cancel"))));

    backdrop.addEventListener("mousedown", (e) => {
      if (e.target === backdrop) close(null);
    });
    document.body.append(backdrop);
    input.focus();
  });
}

function num(v) {
  const n = parseInt(String(v ?? "").trim(), 10);
  return Number.isFinite(n) ? n : 0;
}

function plotOf(r) {
  return `${num(r.Range)}_${num(r.Row)}`;
}

// The QR payload is plot,material,inbred,source,cms,gen,comments.
// Format is fixed — printed labels depend on it.
function qrText(r) {
  return [plotOf(r), r["Material ID"], r["Inbred Code"], r["Source ID"],
    r["CMS reaction"], r.Generation, r.Comments]
    .map((v) => v ?? "").join(",");
}

function splitForRow(row) {
  const hit = state.fieldMap.find((e) => e.row === row);
  if (!hit) return 1;      // before the wizard runs, everything is split 1
  const dops = [...new Set(state.fieldMap.map((e) => e.dop))];
  return dops.indexOf(hit.dop) + 1;
}

function duplicateSet(values) {
  const seen = new Map();
  for (const v of values) {
    if (!v) continue;
    seen.set(v, (seen.get(v) || 0) + 1);
  }
  return new Set([...seen.entries()].filter(([, n]) => n > 1).map(([v]) => v));
}

// ---------------------------------------------------------------- chrome

function visibleTabs() {
  return tabsFor(
    state.types.length ? state.types : ["Selection"],
    Math.max(1, state.plantingDates),
  ).filter((t) => !SPEC.hidden.includes(t));
}

function renderNav() {
  const nav = $("#nav");
  nav.replaceChildren();
  nav.append(el("div", { class: "group" }, "Workbook"));
  for (const name of visibleTabs()) {
    nav.append(el("a", {
      class: name === activeTab ? "active" : "",
      onclick: () => { activeTab = name; render(); },
    }, name));
  }
  $("#topCode").textContent = state.code || "—";
  $("#topTypes").textContent = state.types.length ? state.types.join(", ") : "";
}

function render() {
  if (!visibleTabs().includes(activeTab)) activeTab = "Home";
  renderNav();
  const main = $("#main");
  main.replaceChildren();
  main.scrollTop = 0;
  const base = activeTab.startsWith("Packet Prep ") ? "Packet Prep" : activeTab;
  (VIEWS[base] || viewMissing)(main, activeTab);
}

// ---------------------------------------------------------------- views

const VIEWS = {};

// Nurseries no longer overwrite each other, so Home has to show which one is
// open and let the user move between them.
function nurseryPicker() {
  const all = store.listNurseries();
  const current = store.activeId();

  const select = el("select", {
    class: "f",
    style: "min-width:260px",
    onchange: (e) => { store.switchNursery(e.target.value); reloadActive(); },
  }, all.map((n) => el("option",
    n.id === current ? { value: n.id, selected: "selected" } : { value: n.id },
    n.code ? `${n.code}${n.types.length ? ` — ${n.types.join(", ")}` : ""}`
      : "(untitled nursery)")));

  return el("div", { class: "nursery-bar" },
    el("label", {}, "Open nursery"),
    select,
    el("button", { class: "action", onclick: initNursery }, "New nursery"),
    el("button", {
      class: "action ghost",
      onclick: () => {
        store.duplicateNursery(store.activeId());
        reloadActive();
      },
    }, "Duplicate"),
    el("button", {
      class: "action ghost",
      onclick: () => {
        const name = state.code || "this untitled nursery";
        if (!confirm(
          `Delete ${name}? Its PRISM data, field map, replacements and ` +
          "comments go with it. This cannot be undone.")) return;
        store.deleteNursery(store.activeId());
        reloadActive();
      },
    }, "Delete"));
}

// "Export whole book (in individual tabs) and individually" — both shapes.
function exportBar() {
  return el("div", { class: "btnrow" },
    el("button", {
      class: "action",
      onclick: () => runExport("Whole-book export", async () => {
        const sheets = allSheets();
        if (!sheets.length) {
          alert("Nothing to export yet — import the PRISM export first.");
          return;
        }
        await exportWorkbook(state.fileName || state.code, sheets);
      }),
    }, "Export whole book"),
    el("button", {
      class: "action ghost",
      onclick: () => runExport("Per-tab export", async () => {
        const sheets = allSheets();
        if (!sheets.length) {
          alert("Nothing to export yet — import the PRISM export first.");
          return;
        }
        const skipped = await exportEachTab(state.fileName || state.code, sheets);
        if (skipped.length) {
          alert(`Exported ${sheets.length - skipped.length} tab(s).\n\n` +
            `Skipped, having no rows yet: ${skipped.join(", ")}.`);
        }
      }),
    }, "Export each tab separately"));
}

VIEWS.Home = (main) => {
  main.append(...pageHead("Nursery Workflow",
    "Run each step in order. The same steps as the workbook's Home tab."));

  main.append(nurseryPicker());
  main.append(exportBar());

  if (!state.code) {
    main.append(el("div", { class: "note" },
      "No nursery initialised yet — click ‘New nursery’ above to set " +
      "the code and nursery type."));
  }

  const steps = [
    ["1. Initialise from PRISM export",
      "Set the nursery code, type and file name, then import the PRISM " +
      "export file into Nursery site.",
      () => { activeTab = "Nursery site"; render(); }],
    ["2. Generate all workbook tabs",
      "Builds Material Map, Field Map, Packet Prep and Nursery list from the " +
      "imported data.",
      generateAll],
    ["3. Design Field Map",
      "Enter planting dates and rows. Spike numbers and forward/reverse runs " +
      "are assigned automatically.",
      () => { activeTab = "Field Map"; render(); }],
    ["4. Record replacements and planting errors",
      "Replacements carry an original and a replaced entry; planting errors " +
      "carry the original entry only.",
      () => { activeTab = "Replacements"; render(); }],
    ["5. Pull updated Nursery site from PRISM",
      "Import the refreshed export, then build the Fieldbook from it.",
      () => { activeTab = "Updated nursery site"; render(); }],
  ];

  for (const [title, desc, fn] of steps) {
    main.append(el("div", { class: "home-card" },
      el("h3", {}, title),
      el("p", {}, desc),
      el("button", { class: "action", onclick: fn }, "Go")));
  }

  main.append(el("h1", { style: "margin-top:26px;font-size:16px" }, "Status"));
  main.append(el("div", { class: "kv" },
    el("div", {}, "Nursery code"), el("div", {}, state.code || "—"),
    el("div", {}, "Nursery types"), el("div", {}, state.types.join(", ") || "—"),
    el("div", {}, "Planting dates"), el("div", {}, String(state.plantingDates)),
    el("div", {}, "Packets in Nursery site"), el("div", {}, String(state.prism.length)),
    el("div", {}, "Rows assigned in Field Map"), el("div", {}, String(state.fieldMap.length)),
    el("div", {}, "Updated nursery site rows"), el("div", {}, String(state.updatedPrism.length)),
    el("div", {}, "Replacements logged"), el("div", {}, String(state.replacements.length)),
    el("div", {}, "Planting errors logged"), el("div", {}, String(state.plantingErrors.length)),
    el("div", {}, "Nurseries stored"), el("div", {}, String(store.listNurseries().length))));
};

// Import the PRISM export straight from the downloaded file. Parsing happens
// on the backend, which already has openpyxl — the browser would otherwise
// need a spreadsheet library bundled just to read one sheet.
async function importFile(file, key) {
  const body = new FormData();
  body.append("file", file);
  const res = await fetch("/parse/prism", { method: "POST", body });
  if (!res.ok) {
    let detail = `${res.status} ${res.statusText}`;
    try {
      detail = (await res.json()).detail || detail;
    } catch { /* response was not JSON */ }
    throw new Error(detail);
  }
  const data = await res.json();
  if (!data.rows.length) {
    throw new Error("The file parsed, but no data rows were found.");
  }
  state[key] = data.rows;
  save();
  return data;
}

function importView(main, title, subtitle, key) {
  main.append(...pageHead(title, subtitle));
  const rows = state[key];

  const status = el("p", { class: "sub" }, "");
  const picker = el("input", {
    type: "file",
    accept: ".xlsx,.xlsm,.csv,.tsv,.txt",
    style: "display:none",
    onchange: async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      status.textContent = `Reading ${file.name}…`;
      try {
        const data = await importFile(file, key);
        alert(`${data.count} packets read from “${data.sheet}”.`);
        render();
      } catch (err) {
        status.textContent = "";
        alert(`Could not import ${file.name}.\n\n${err.message}`);
      } finally {
        e.target.value = "";      // allow re-picking the same file
      }
    },
  });

  const drop = el("div", {
    class: "empty",
    style: "text-align:center;cursor:pointer",
    onclick: () => picker.click(),
    ondragover: (e) => { e.preventDefault(); drop.classList.add("dragging"); },
    ondragleave: () => drop.classList.remove("dragging"),
    ondrop: async (e) => {
      e.preventDefault();
      drop.classList.remove("dragging");
      const file = e.dataTransfer.files[0];
      if (!file) return;
      status.textContent = `Reading ${file.name}…`;
      try {
        const data = await importFile(file, key);
        alert(`${data.count} packets read from “${data.sheet}”.`);
        render();
      } catch (err) {
        status.textContent = "";
        alert(`Could not import ${file.name}.\n\n${err.message}`);
      }
    },
  },
    el("b", {}, "Choose the PRISM export"),
    el("div", { style: "margin-top:6px" },
      "Click to browse, or drag the file here — .xlsx, .xlsm, .csv or .tsv. " +
      "The header row is found automatically, whether it sits at row 1 or row 5."));

  main.append(picker, drop, status);

  main.append(el("div", { class: "btnrow" },
    el("button", { class: "action", onclick: () => picker.click() },
      "Choose file…"),
    rows.length
      ? el("button", {
        class: "action ghost",
        onclick: () => { state[key] = []; save(); render(); },
      }, "Clear imported data")
      : null));

  if (!rows.length) {
    main.append(emptyState("No data imported yet."));
    return;
  }
  const cols = Object.keys(rows[0]).slice(0, 10);
  main.append(el("p", { class: "sub" },
    `${rows.length} rows imported — showing the first 50.`));
  main.append(table(cols, rows.slice(0, 50).map((r) => cols.map((c) => r[c]))));
}

VIEWS["Nursery site"] = (main) => importView(main, "Nursery site",
  "Import the PRISM export. This is the source for every generated tab.",
  "prism");

VIEWS["Updated nursery site"] = (main) => {
  importView(main, "Updated nursery site",
    "Import the re-downloaded PRISM file after the breeder applies " +
    "replacements and errors.",
    "updatedPrism");
  main.append(el("div", { class: "btnrow" },
    el("button", {
      class: "action",
      onclick: () => {
        if (!state.updatedPrism.length) {
          alert("Import the updated PRISM export first.");
          return;
        }
        activeTab = "Fieldbook";
        render();
      },
    }, "Build Fieldbook from this tab")));
};

// Fields offered in the Material Map cell checkboxes. Anything PRISM supplies
// can be shown; these are the ones worth defaulting to.
const MAP_FIELDS = [
  "Material ID", "Inbred Code", "Hybrid Code", "Source ID",
  "Generation", "CMS reaction", "Pedigree",
];
const MAP_FIELDS_DEFAULT = ["Material ID", "Inbred Code", "Hybrid Code"];

/**
 * The field map as a map, not a table.
 *
 * Range numbers run down both sides and row numbers across top and bottom, as
 * in the workbook. `cellText` decides what goes in each plot and `cellFill`
 * what colour it is, so Material Map and Field Map share the same geometry.
 */
function buildMapGrid({ cellText, cellFill }) {
  const maxRange = Math.max(...state.prism.map((r) => num(r.Range)));
  const maxRow = Math.max(...state.prism.map((r) => num(r.Row)));
  const byPlot = new Map(
    state.prism.map((r) => [`${num(r.Range)}:${num(r.Row)}`, r]));

  const rowNums = Array.from({ length: maxRow }, (_, i) => i + 1);
  const headCells = ["Rng \\ Row", ...rowNums, "Rng"]
    .map((h) => el("th", {}, h));

  const body = [];
  // Ranges descend so the top range prints at the top, as in the workbook.
  for (let rng = maxRange; rng >= 1; rng--) {
    const cells = [el("td", { class: "hdr" }, rng)];
    for (const w of rowNums) {
      const record = byPlot.get(`${rng}:${w}`);
      const fill = record ? cellFill?.(record, rng, w) : null;
      cells.push(el("td", { style: fill ? `background:${fill}` : "" },
        record ? cellText(record, rng, w) : ""));
    }
    cells.push(el("td", { class: "hdr" }, rng));
    body.push(el("tr", {}, cells));
  }
  // Row numbers repeat along the foot, as the client asked.
  body.push(el("tr", {},
    ["Rng \\ Row", ...rowNums, "Rng"].map((h) => el("td", { class: "hdr" }, h))));

  return el("div", { class: "scroll" },
    el("table", { class: "grid" },
      el("thead", {}, el("tr", {}, headCells)), el("tbody", {}, body)));
}

VIEWS["Material Map"] = (main) => {
  main.append(...pageHead("Material Map",
    "Range numbers down both sides, field rows across top and bottom."));
  if (!state.prism.length) {
    main.append(emptyState("Import Nursery site data first."));
    return;
  }

  // Which fields appear in each cell is the user's choice, and it sticks.
  const chosen = new Set(
    state.mapFields?.length ? state.mapFields : MAP_FIELDS_DEFAULT);

  const holder = el("div");
  const draw = () => {
    const keys = MAP_FIELDS.filter((k) => chosen.has(k));
    holder.replaceChildren(buildMapGrid({
      cellText: (r) => keys
        .map((k) => String(r[k] ?? "").slice(0, 16))
        .filter(Boolean).join("\n"),
    }));
  };

  main.append(el("div", { class: "nursery-bar" },
    el("label", {}, "Show in each cell"),
    ...MAP_FIELDS.map((field) => {
      const box = el("input", { type: "checkbox" });
      box.checked = chosen.has(field);
      box.addEventListener("change", () => {
        if (box.checked) chosen.add(field); else chosen.delete(field);
        state.mapFields = [...chosen];
        save();
        draw();
      });
      return el("label", { class: "filter-row" }, box, el("span", {}, field));
    })));

  main.append(holder);
  draw();

  main.append(el("div", { class: "btnrow" },
    el("button", {
      class: "action",
      onclick: () => runExport("Export", () =>
        download(`${state.code || "nursery"} - Material Map`,
          [materialMapSheet(chosen)])),
    }, "Export")));

  registerSheet("Material Map", () => materialMapSheet(chosen));
};

/** Material Map as a grid of plots, laid out the way it is displayed. */
function materialMapSheet(chosen) {
  const keys = MAP_FIELDS.filter((k) => chosen.has(k));
  const maxRange = Math.max(...state.prism.map((r) => num(r.Range)));
  const maxRow = Math.max(...state.prism.map((r) => num(r.Row)));
  const byPlot = new Map(
    state.prism.map((r) => [`${num(r.Range)}:${num(r.Row)}`, r]));
  const rowNums = Array.from({ length: maxRow }, (_, i) => i + 1);

  const rows = [];
  for (let rng = maxRange; rng >= 1; rng--) {
    rows.push([rng, ...rowNums.map((w) => {
      const r = byPlot.get(`${rng}:${w}`);
      return r ? keys.map((k) => r[k] ?? "").filter(Boolean).join(" / ") : "";
    }), rng]);
  }
  return {
    name: "Material Map",
    headers: ["Rng \\ Row", ...rowNums.map(String), "Rng"],
    rows,
  };
}

// One colour per planting date, so the map reads at a glance.
const DOP_COLOURS = [
  "#cfe2f3", "#d9ead3", "#fff2cc", "#f4cccc", "#e6d0f0",
  "#d0e0e3", "#fce5cd", "#ead1dc",
];

VIEWS["Field Map"] = (main) => {
  main.append(...pageHead("Field Map",
    "Plots coloured by planting date, with the spike and run in each cell."));

  main.append(el("div", { class: "btnrow" },
    el("button", { class: "action", onclick: fieldMapWizard },
      "Set planting dates and rows"),
    state.fieldMap.length
      ? el("button", {
        class: "action ghost",
        onclick: () => {
          if (!confirm("Clear the planting dates and row assignments?")) return;
          state.fieldMap = [];
          state.plantingDates = 1;
          save();
          render();
        },
      }, "Clear")
      : null));

  if (!state.fieldMap.length) {
    main.append(emptyState(
      "No planting dates set. Spike numbers follow the 1,2,2,1 cycle and runs " +
      "alternate every two rows, starting forward."));
    return;
  }

  const dops = [...new Set(state.fieldMap.map((e) => e.dop))];
  const byRow = new Map(state.fieldMap.map((e) => [e.row, e]));
  const colourFor = (dop) => DOP_COLOURS[dops.indexOf(dop) % DOP_COLOURS.length];

  main.append(el("p", { class: "sub" },
    `${dops.length} planting date(s), ${state.fieldMap.length} rows assigned. ` +
    `Seed qty/plot: ${state.seedQty || "—"} g`));

  // Legend, so the colours mean something without hovering.
  main.append(el("div", { class: "nursery-bar" },
    el("label", {}, "Planting dates"),
    ...dops.map((dop, i) => el("span", { class: "legend-chip" },
      el("span", {
        class: "legend-swatch",
        style: `background:${colourFor(dop)}`,
      }),
      `${i + 1}. ${dop}`))));

  if (state.prism.length) {
    main.append(buildMapGrid({
      // Each plot shows which spike it belongs to and which way the run goes,
      // which is what the planting crew needs off the map itself.
      cellText: (_r, _rng, row) => {
        const hit = byRow.get(row);
        if (!hit) return "";
        return `S${hit.spike}\n${hit.run === "forward" ? "▶" : "◀"}`;
      },
      cellFill: (_r, _rng, row) => {
        const hit = byRow.get(row);
        return hit ? colourFor(hit.dop) : null;
      },
    }));
  }

  // The row assignments themselves stay available as an editable table.
  main.append(el("h1", { style: "margin-top:24px;font-size:16px" },
    "Row assignments"));

  const rows = state.fieldMap.map((e) => ({
    "Planting date": e.dop,
    Row: e.row,
    Spike: e.spike,
    Run: e.run,
    "Split no.": dops.indexOf(e.dop) + 1,
    "Seed qty/plot (g)": e.qty,
  }));

  const node = grid({
    id: "Field Map",
    columns: [
      { key: "Planting date", type: "date" },
      { key: "Row", type: "number" },
      { key: "Spike", type: "number" },
      { key: "Run", type: "select", options: ["forward", "reverse"] },
      { key: "Split no.", type: "number" },
      { key: "Seed qty/plot (g)", type: "text" },
    ],
    rows,
    rowKey: (r) => `row:${r.Row}`,
    cellColour: (row) => colourFor(row["Planting date"]),
    onExport: (api) => runExport("Export", () =>
      exportGrid(state.code, "Field Map", api)),
  });
  registerSheet("Field Map", () => sheetOf("Field Map", node.gridApi));
  main.append(node);
};

async function fieldMapWizard() {
  if (!state.prism.length) {
    alert("Import Nursery site data first.");
    return;
  }
  const n = num(prompt("How many planting dates?",
    String(state.plantingDates || 2)));
  if (n < 1) return;

  const qty = prompt("Seed quantity per plot (grams):", state.seedQty || "");
  if (qty === null) return;

  const allRows = [...new Set(state.prism.map((r) => num(r.Row)))]
    .filter((r) => r > 0).sort((a, b) => a - b);

  const dopRows = [];
  const dops = [];
  for (let i = 1; i <= n; i++) {
    // Picked from a calendar so every planting date is the same ISO shape —
    // the Packet Prep split and the Nursery data rows both group on it.
    // Picked from a calendar so every planting date is the same ISO shape —
    // the Packet Prep split and the Nursery data rows both group on it.
    const dop = await pickDate(`Date of planting ${i}`, dops[i - 2]);
    if (!dop) return;
    const already = new Set(dopRows.flat());
    const picked = await pickRows(
      `Rows for planting date ${i} (${dop})`, allRows, already);
    if (!picked) return;
    dops.push(dop);
    dopRows.push(picked);
  }

  // A row in two planting dates would print packets onto the wrong split,
  // so refuse rather than silently overwrite.
  try {
    assignSplits(dopRows);
  } catch (e) {
    alert(`${e.message}\n\nFix the row lists and run the wizard again.`);
    return;
  }

  const out = [];
  dopRows.forEach((rows, i) => {
    for (const row of rows) {
      out.push({
        dop: dops[i],
        row,
        spike: spikeForRow(row),
        run: runDirection(row),
        qty,
      });
    }
  });

  state.fieldMap = out;
  state.plantingDates = n;
  state.seedQty = qty;
  save();
  alert(`${n} planting date(s) recorded, ${out.length} rows assigned.`);
  render();
}

VIEWS["Nursery data"] = (main) => {
  main.append(...pageHead("Nursery data",
    "Printable front page. The right-hand column is filled in by hand."));
  main.append(el("div",
    { style: "font-weight:700;font-size:19px;margin:6px 0 14px" },
    "R&D Fieldbook - Grain Sorghum"));
  // "1st DOP and rows" etc. are filled from Field Map. The base spec names
  // three; a nursery with more planting dates gets the extra labels appended
  // rather than quietly losing a date.
  const dops = [...new Set(state.fieldMap.map((e) => e.dop))];
  const ordinal = (n) => {
    const teens = n % 100 >= 11 && n % 100 <= 13;
    const suffix = teens ? "th" : ({ 1: "st", 2: "nd", 3: "rd" }[n % 10] ?? "th");
    return `${n}${suffix}`;
  };
  const labels = [...SPEC.nursery_data_labels];
  for (let i = 4; i <= dops.length; i++) {
    const label = `${ordinal(i)} DOP and rows`;
    if (!labels.includes(label)) {
      labels.splice(labels.indexOf("Tag rows"), 0, label);
    }
  }

  const dopFor = (label) => {
    const m = /^(\d+)(?:st|nd|rd|th) DOP and rows$/.exec(label);
    if (!m) return null;
    const dop = dops[Number(m[1]) - 1];
    if (!dop) return "";
    const rows = state.fieldMap.filter((e) => e.dop === dop)
      .map((e) => e.row).sort((a, b) => a - b);
    return `${dop} — rows ${rows.join(", ")}`;
  };

  const defaults = SPEC.nursery_data_defaults ?? {};
  const isDate = (label) => /date/i.test(label);

  const rows = labels.map((label) => {
    const saved = state.nurseryData[label] ?? {};
    const computed = dopFor(label);
    return {
      Item: label,
      Value: saved.value ?? computed ?? defaults[label] ?? "",
      Comments: saved.comment ?? "",
    };
  });

  const node = grid({
    id: "Nursery data",
    columns: [
      { key: "Item", type: "text", readOnly: true },
      // Only the rows whose label mentions a date get a calendar — "Planter"
      // and "Seeds/side" are on the same column but are not dates.
      { key: "Value", type: "text", typeFor: (row) => (isDate(row.Item) ? "date" : "text") },
      { key: "Comments", type: "multiline" },
    ],
    rows,
    rowKey: (r) => r.Item,
    owned: true,
    onEdit: (row, key, value) => {
      const entry = state.nurseryData[row.Item] ?? {};
      if (key === "Value") entry.value = value;
      if (key === "Comments") entry.comment = value;
      state.nurseryData[row.Item] = entry;
      save();
    },
    onExport: (api) => runExport("Export", () =>
      exportGrid(state.code, "Nursery data", api)),
  });

  main.append(el("div", { class: "btnrow" },
    el("button", { class: "action ghost", onclick: captureGps }, "Capture GPS"),
    el("button", {
      class: "action ghost",
      onclick: () => {
        for (const [label, value] of Object.entries(defaults)) {
          state.nurseryData[label] = {
            ...(state.nurseryData[label] ?? {}), value,
          };
        }
        save();
        render();
      },
    }, "Restore defaults")));

  registerSheet("Nursery data", () => sheetOf("Nursery data", node.gridApi));
  main.append(node);
};

/** Write the device's coordinates into the GPS row; still editable after. */
function captureGps() {
  if (!navigator.geolocation) {
    alert("This device has no location service. Type the coordinates in by hand.");
    return;
  }
  navigator.geolocation.getCurrentPosition(
    (pos) => {
      const { latitude, longitude } = pos.coords;
      const value = `${latitude.toFixed(6)}, ${longitude.toFixed(6)}`;
      state.nurseryData["GPS coordinates"] = {
        ...(state.nurseryData["GPS coordinates"] ?? {}), value,
      };
      save();
      render();
      alert(`GPS captured: ${value}`);
    },
    (err) => alert(
      `Could not read the location (${err.message}).\n\n` +
      "Type the coordinates in by hand instead."),
    { enableHighAccuracy: true, timeout: 15000 });
}

VIEWS["Packet Prep"] = (main, tabName) => {
  const splitNo = num(tabName.split(" ").pop()) || 1;
  main.append(...pageHead(tabName,
    `Packets for planting date ${splitNo}. ` +
    "QR CODE column feeds the barcode printer."));

  if (!state.prism.length) {
    main.append(emptyState("Import Nursery site data first."));
    return;
  }

  const packets = state.prism.filter((r) => splitForRow(num(r.Row)) === splitNo);

  // Packets must come off the rack in the order the planter needs them, so
  // the sheet is ordered before the rack numbers are handed out: rows
  // descending, range descending on a forward row and ascending on a reverse
  // one, then grouped with spike 1 ahead of spike 2. Ported from PP2026_Run.
  const byPlot = new Map(
    packets.map((r) => [`${num(r.Range)}:${num(r.Row)}`, r]));
  const ordered = packetPrepOrder(
    packets.map((r) => [num(r.Range), num(r.Row)]));

  // Rack order then counts 1..n within each spike.
  const seen = new Map();
  const numbered = ordered.map(([rng, row]) => {
    const spike = spikeForRow(row);
    const seq = (seen.get(spike) ?? 0) + 1;
    seen.set(spike, seq);
    return { rng, row, spike, seq, record: byPlot.get(`${rng}:${row}`) ?? {} };
  });

  // Four digit columns, as the document shows: 7 -> 0 | 0 | 0 | 7. Widened
  // only if a spike ever runs past 9999, so every value stays the same width.
  const digitWidth = Math.max(4,
    ...numbered.map((p) => String(p.seq).length));
  const digitKeys = Array.from({ length: digitWidth }, (_, i) => `D${i + 1}`);

  const rows = numbered.map(({ rng, row, spike, seq, record }) => {
    const digits = rackDigits(seq, digitWidth);
    return {
      "Nursery name": state.code,
      "QR CODE": qrText(record),
      Range: rng,
      Row: row,
      Plot: `${rng}_${row}`,
      "SPIKE#": spike,
      "RACK ORDER": seq,
      ...Object.fromEntries(digitKeys.map((k, i) => [k, digits[i] ?? ""])),
      "Split no.": splitNo,
      "Material ID": record["Material ID"] ?? "",
      "Inbred Code": record["Inbred Code"] ?? "",
      "Hybrid Code": record["Hybrid Code"] ?? "",
      "Source ID": record["Source ID"] ?? "",
      "CMS reaction": record["CMS reaction"] ?? "",
      Generation: record.Generation ?? "",
    };
  });

  const spikeCounts = [...seen.entries()].sort((a, b) => a[0] - b[0]);
  main.append(el("p", { class: "sub" },
    `${rows.length} packets in this split, racked in planting order. ` +
    spikeCounts.map(([s, n]) => `Spike ${s}: ${n}`).join(", ") +
    `. Rack order split into ${digitWidth} digit columns.`));

  const node = grid({
    id: tabName,
    columns: [
      { key: "Nursery name", type: "text" },
      { key: "QR CODE", type: "text" },
      { key: "Range", type: "number" },
      { key: "Row", type: "number" },
      { key: "Plot", type: "text" },
      { key: "SPIKE#", type: "number" },
      { key: "RACK ORDER", type: "number" },
      // One column per digit, for the label template.
      ...digitKeys.map((k, i) => ({
        key: k, label: `Digit ${i + 1}`, type: "text",
      })),
      { key: "Split no.", type: "number" },
      { key: "Material ID", type: "text" },
      { key: "Inbred Code", type: "text" },
      { key: "Hybrid Code", type: "text" },
      { key: "Source ID", type: "text" },
      { key: "CMS reaction", type: "text" },
      { key: "Generation", type: "text" },
    ],
    rows,
    rowKey: (r) => r.Plot,
    onExport: (api) => runExport("Export", () =>
      exportGrid(state.code, tabName, api)),
  });

  main.append(el("div", { class: "btnrow" },
    el("button", {
      class: "action",
      onclick: () => runExport("Bartender export", () =>
        exportBartender(state.code, tabName, node.gridApi)),
    }, "Export for Bartender (Source ID Z→A)")));

  registerSheet(tabName, () => sheetOf(tabName, node.gridApi));
  main.append(node);
};

VIEWS["Nursery list"] = (main) => {
  main.append(...pageHead("Nursery list",
    "Unique materials from Nursery site, Source ID ascending."));
  if (!state.prism.length) {
    main.append(emptyState("Import Nursery site data first."));
    return;
  }

  const counts = new Map();
  for (const r of state.prism) {
    const key = r["Source ID"] || "";
    if (!key) continue;
    const rec = counts.get(key)
      || { repeats: 0, inbred: r["Inbred Code"], hybrid: r["Hybrid Code"] };
    rec.repeats += 1;
    counts.set(key, rec);
  }
  const entries = [...counts.entries()].sort((a, b) => a[0].localeCompare(b[0]));

  const dupInbred = duplicateSet(entries.map(([, v]) => v.inbred));
  const dupHybrid = duplicateSet(entries.map(([, v]) => v.hybrid));
  let showDups = false;

  const rows = entries.map(([src, v]) => ({
    "Source ID": src,
    Repeats: v.repeats,
    "Qty Required": (v.repeats * 1.4).toFixed(1),
    "Inbred Code": v.inbred ?? "",
    "Hybrid Code": v.hybrid ?? "",
  }));

  main.append(el("p", { class: "sub" },
    `${rows.length} unique Source IDs. Click a column heading to sort, or the ` +
    "▾ beside it to filter."));

  const columns = [
    { key: "Source ID", type: "text" },
    { key: "Repeats", type: "number" },
    { key: "Qty Required", type: "number" },
    { key: "Inbred Code", type: "text" },
    { key: "Hybrid Code", type: "text" },
  ];

  const holder = el("div");
  const draw = () => {
    const node = grid({
      id: "Nursery list",
      columns,
      rows,
      // Source ID is unique here by construction, so it identifies the row.
      rowKey: (r) => r["Source ID"],
      cellColour: (row, key, value) => {
        if (!showDups) return null;
        if (key === "Inbred Code" && dupInbred.has(value)) return "#ffd9d9";
        if (key === "Hybrid Code" && dupHybrid.has(value)) return "#ffd9d9";
        return null;
      },
      onExport: (api) => runExport("Export", () =>
        exportGrid(state.code, "Nursery list", api)),
    });
    // Registered while rendered so the whole-book export picks up the user's
    // own sorting, filters and added columns.
    registerSheet("Nursery list", () => sheetOf("Nursery list", node.gridApi));
    holder.replaceChildren(node);
  };

  main.append(el("div", { class: "btnrow" },
    el("button", {
      class: "action",
      onclick: () => {
        showDups = !showDups;
        draw();
        const n = dupInbred.size + dupHybrid.size;
        if (showDups) {
          alert(n
            ? `${n} duplicated code value(s) highlighted.`
            : "No duplicate Inbred or Hybrid codes found.");
        }
      },
    }, "Check duplicate codes")));

  main.append(holder);
  draw();
};

// Replacements and Planting errors are the same log with a different column
// set, so they share a renderer. `storeKey` picks which list is being edited.
// The nine fields the client's export format wants for each entry, in order.
const ENTRY_BLOCK = [
  "QR Key", "Source ID", "Material ID", "Inbred Code",
  "Experimental Hybrid Code", "Pedigree", "Generation", "CMS reaction",
  "W/column",
];

/**
 * Expand a scanned QR back into the full nine-column entry block.
 *
 * This is what lets the tab stay short. The client asked to drop Material ID,
 * Source ID and the rest from the screen, but the export format still carries
 * them — so they are looked up from the nursery data at export time rather
 * than typed twice.
 *
 * W/column has no source in PRISM yet, so it exports blank.
 */
function entryBlockFor(qr) {
  const payload = String(qr ?? "").trim();
  if (!payload) return ENTRY_BLOCK.map(() => "");
  const plot = payload.split(",")[0].trim();
  const hit = state.prism.find((r) => plotOf(r) === plot)
    || state.updatedPrism.find((r) => plotOf(r) === plot);
  if (!hit) return [payload, ...ENTRY_BLOCK.slice(1).map(() => "")];
  return [
    payload,
    hit["Source ID"] ?? "",
    hit["Material ID"] ?? "",
    hit["Inbred Code"] ?? "",
    hit["Hybrid Code"] ?? "",
    hit.Pedigree ?? "",
    hit.Generation ?? "",
    hit["CMS reaction"] ?? "",
    hit["W/column"] ?? "",
  ];
}

/** The banded sheet from Tab information.xlsx, for either log. */
function logSheet(tabName, storeKey, withReplaced) {
  const bands = withReplaced
    ? [["Original entry", 9], ["Replaced entry", 9], ["Reason", 1]]
    : [["Original entry", 9], ["Reason", 1]];
  const headers = withReplaced
    ? [...ENTRY_BLOCK, ...ENTRY_BLOCK, "Reason"]
    : [...ENTRY_BLOCK, "Reason"];
  const rows = state[storeKey].map((rec) => {
    const original = entryBlockFor(rec[QR_ORIGINAL]);
    const reason = rec.Reason ?? "";
    return withReplaced
      ? [...original, ...entryBlockFor(rec[QR_REPLACED]), reason]
      : [...original, reason];
  });
  return { name: tabName, bands, headers, rows };
}

// Replacements and Planting errors are the same log with a different column
// set, so they share a renderer.
function logView(main, { tabName, title, subtitle, storeKey, columns,
  withReplaced }) {
  main.append(...pageHead(title, subtitle));

  const blankRow = () => Object.fromEntries(columns.map((c) => [
    c.key,
    c.key === "Stage" ? STAGE_OPTIONS[0]
      : c.key === "Status" ? STATUS_DEFAULT : "",
  ]));

  main.append(el("div", { class: "btnrow" },
    el("button", {
      class: "action",
      onclick: () => { state[storeKey].push(blankRow()); save(); render(); },
    }, "Add row"),
    el("button", {
      class: "action ghost",
      onclick: () => resolvePlots(storeKey),
    }, "Read QR → Plot"),
    el("button", {
      class: "action ghost",
      onclick: () => runExport("Export", async () => {
        const sheet = logSheet(tabName, storeKey, withReplaced);
        if (!sheet.rows.length) {
          alert(`Nothing logged on ${title} yet.`);
          return;
        }
        await download(`${state.code || "nursery"} - ${tabName}`, [sheet]);
      }),
    }, "Export (client format)")));

  // The export is always available from the button above, even when the tab
  // is empty, so it is registered before the early return.
  registerSheet(tabName, () => logSheet(tabName, storeKey, withReplaced));

  if (!state[storeKey].length) {
    main.append(emptyState(`Nothing logged on ${title}.`));
    return;
  }

  main.append(grid({
    id: tabName,
    columns,
    rows: state[storeKey],
    // These rows have no natural key, so position is the identity — safe
    // because the tab owns its data and nothing rebuilds it underneath.
    rowKey: (_row, i) => `row:${i}`,
    owned: true,
    onEdit: () => save(),
    onAddRow: () => { state[storeKey].push(blankRow()); save(); render(); },
    onDeleteRow: (row) => {
      const i = state[storeKey].indexOf(row);
      if (i >= 0) state[storeKey].splice(i, 1);
      save();
      render();
    },
  }));
}

VIEWS.Replacements = (main) => logView(main, {
  tabName: "Replacements",
  title: "Replacements",
  subtitle: "Scan the original entry and its replacement. The export expands " +
    "each QR into the full nine-column entry block.",
  storeKey: "replacements",
  columns: REPLACEMENT_COLUMNS,
  withReplaced: true,
});

VIEWS["Planting errors"] = (main) => logView(main, {
  tabName: "Planting errors",
  title: "Planting errors",
  subtitle: "Original entry only — nothing was planted in its place.",
  storeKey: "plantingErrors",
  columns: PLANTING_ERROR_COLUMNS,
  withReplaced: false,
});

// The QR payload leads with the plot, so a scan can fill the Plot column.
// A value matching no plot is left blank and counted rather than guessed at.
function resolvePlots(storeKey) {
  let resolved = 0;
  let unresolved = 0;
  for (const rec of state[storeKey]) {
    const payload = String(rec[QR_ORIGINAL] ?? "").trim();
    if (!payload) continue;
    const plot = payload.split(",")[0].trim();
    const hit = state.prism.find((r) => plotOf(r) === plot);
    if (hit) {
      rec.Plot = plotOf(hit);
      resolved += 1;
    } else {
      unresolved += 1;
    }
  }
  save();
  render();
  alert(unresolved
    ? `${resolved} scan(s) resolved.\n${unresolved} QR value(s) matched no ` +
      "plot — those rows were left blank."
    : `${resolved} scan(s) resolved to a plot.`);
}

/** Fieldbook rows in serpentine order, shared with Date recording. */
function fieldbookRows() {
  const src = state.updatedPrism;
  const plots = src.map((r) => [num(r.Range), num(r.Row)]);
  const byPlot = new Map(src.map((r) => [`${num(r.Range)}:${num(r.Row)}`, r]));
  return serpentineByRange(plots).map(([rng, row]) => {
    const r = byPlot.get(`${rng}:${row}`) || {};
    return {
      Range: rng,
      Row: row,
      "Material ID": r["Material ID"] ?? "",
      "Source ID": r["Source ID"] ?? "",
      "Inbred Code": r["Inbred Code"] ?? "",
      "Hybrid Code": r["Hybrid Code"] ?? "",
      Gen: r.Generation ?? "",
      CMS: r["CMS reaction"] ?? "",
      Plot: `${rng}_${row}`,
      Comments: r.Comments ?? "",
    };
  });
}

VIEWS.Fieldbook = (main) => {
  main.append(...pageHead("Fieldbook",
    "Built from Updated nursery site, in serpentine order."));

  if (!state.updatedPrism.length) {
    main.append(emptyState(
      "Import the updated PRISM export on the ‘Updated nursery site’ tab first."));
    return;
  }

  // Print sits at the top, as asked — it is the reason this tab is opened.
  main.append(el("div", { class: "btnrow" },
    el("button", { class: "action", onclick: () => window.print() }, "Print"),
    el("button", {
      class: "action ghost",
      onclick: () => captureFieldbook(),
    }, "Capture"),
    isType("AB")
      ? el("button", {
        class: "action ghost",
        onclick: () => { activeTab = "Date recording"; render(); },
      }, "AB selection map")
      : null));

  const all = fieldbookRows();
  let rangeQuery = "";
  let rowQuery = "";
  const holder = el("div");

  // Range and Row are searched separately, and an empty box means "all" —
  // so one box alone narrows to a whole bay or a whole row.
  const draw = () => {
    const rows = all.filter((r) =>
      (!rangeQuery || String(r.Range) === rangeQuery)
      && (!rowQuery || String(r.Row) === rowQuery));

    const node = grid({
      id: "Fieldbook",
      columns: [
        { key: "Range", type: "number" },
        { key: "Row", type: "number" },
        { key: "Material ID", type: "text" },
        { key: "Source ID", type: "text" },
        { key: "Inbred Code", type: "text" },
        { key: "Hybrid Code", type: "text" },
        { key: "Gen", type: "text" },
        { key: "CMS", type: "text" },
        { key: "Plot", type: "text" },
        { key: "Comments", type: "multiline" },
      ],
      rows,
      rowKey: (r) => r.Plot,
      cellColour: (row) =>
        (isType("AB") && String(row.CMS).toUpperCase() === "B"
          ? "#e2f0d9" : null),
      onExport: (api) => runExport("Export", () =>
        exportGrid(state.code, "Fieldbook", api)),
    });
    registerSheet("Fieldbook", () => sheetOf("Fieldbook", node.gridApi));

    holder.replaceChildren(
      el("p", { class: "sub" },
        `${rows.length} of ${all.length} plots` +
        (isType("AB") ? ". B lines shaded green." : ".")),
      node);
  };

  const rangeBox = el("input", {
    class: "f", type: "number", placeholder: "All ranges",
    oninput: (e) => { rangeQuery = e.target.value.trim(); draw(); },
  });
  const rowBox = el("input", {
    class: "f", type: "number", placeholder: "All rows",
    oninput: (e) => { rowQuery = e.target.value.trim(); draw(); },
  });

  main.append(el("div", { class: "nursery-bar" },
    el("label", {}, "Range"), rangeBox,
    el("label", {}, "Row"), rowBox,
    el("button", {
      class: "action ghost",
      onclick: () => {
        rangeBox.value = "";
        rowBox.value = "";
        rangeQuery = "";
        rowQuery = "";
        draw();
      },
    }, "Show all")));

  main.append(holder);
  draw();

  main.append(el("div", { class: "note" },
    "Printing: landscape A4, duplex flipped on the short edge, page numbers " +
    "centred, file name top right. Duplex is a printer-driver setting — " +
    "choose it in the print dialog."));
};

/**
 * Capture the fieldbook as it stands.
 *
 * Read as a timestamped snapshot of what is on screen, so a walk of the
 * nursery can be kept and compared later. Confirm with the client if they
 * meant something else by "capture".
 */
function captureFieldbook() {
  const rows = fieldbookRows();
  if (!rows.length) {
    alert("Nothing to capture yet.");
    return;
  }
  const label = prompt("Name this capture:",
    `Fieldbook ${todayISO()}`);
  if (!label) return;
  state.snapshots.push({
    label,
    takenAt: new Date().toISOString(),
    rows,
  });
  save();
  alert(`Captured ${rows.length} plots as “${label}”.\n\n` +
    `${state.snapshots.length} capture(s) stored for this nursery.`);
}

// Four passes through the nursery. The month is not recorded per cell — the
// crew writes a bare day — so it is derived from a start date, as in the VBA.
const SELECTION_COLUMNS = ["S 1", "S 2", "S 3", "S 4"];
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/** A-lines out of the fieldbook, which is what date recording works from. */
function aLineRows() {
  return fieldbookRows().filter((r) => String(r.CMS).toUpperCase() === "A");
}

VIEWS["Date recording"] = (main) => {
  main.append(...pageHead("Date recording",
    "A-lines from the Fieldbook. S 1 to S 4 hold day-of-month numbers; " +
    "the month is worked out from the recording start date."));

  if (!state.updatedPrism.length) {
    main.append(emptyState(
      "Import the updated PRISM export on the ‘Updated nursery site’ tab first."));
    return;
  }

  const src = aLineRows();
  if (!src.length) {
    main.append(emptyState(
      "No A-lines in the Fieldbook — nothing to record selections against."));
    return;
  }

  let mode = "vertical";
  const holder = el("div");

  const draw = () => {
    const plots = src.map((r) => [num(r.Range), num(r.Row)]);
    // Vertical goes between two rows at a time; horizontal walks bay by bay.
    const order = mode === "vertical"
      ? serpentineTwoRowBands(plots)
      : serpentineByRange(plots);
    const byPlot = new Map(src.map((r) => [`${r.Range}:${r.Row}`, r]));

    const rows = order.map(([rng, row]) => {
      const r = byPlot.get(`${rng}:${row}`) || {};
      const stored = state.dateRecording?.[`${rng}_${row}`] ?? {};
      const out = {
        Group: bandForRow(row),
        Range: rng,
        Row: row,
        "O/E": parityForRow(row),
        "Material ID": r["Material ID"] ?? "",
        "Source ID": r["Source ID"] ?? "",
        Gen: r.Gen ?? "",
        CMS: r.CMS ?? "",
        "In. Code": r["Inbred Code"] ?? "",
      };
      for (const c of SELECTION_COLUMNS) {
        out[c] = stored[c] ?? "";
        out[`${c.replace(" ", "")} Month`] = monthLabel(out[c]);
      }
      return out;
    });

    const columns = [
      { key: "Group", type: "number", readOnly: true },
      { key: "Range", type: "number", readOnly: true },
      { key: "Row", type: "number", readOnly: true },
      { key: "O/E", type: "text", readOnly: true },
      ...SELECTION_COLUMNS.flatMap((c) => [
        { key: c, type: "number" },
        // Derived, so read-only — editing it would not change the day it
        // came from and the two would silently disagree.
        { key: `${c.replace(" ", "")} Month`, type: "text", readOnly: true },
      ]),
      { key: "Material ID", type: "text", readOnly: true },
      { key: "Source ID", type: "text", readOnly: true },
      { key: "Gen", type: "text", readOnly: true },
      { key: "CMS", type: "text", readOnly: true },
      { key: "In. Code", type: "text", readOnly: true },
    ];

    const node = grid({
      id: "Date recording",
      columns,
      rows,
      rowKey: (r) => `${r.Range}_${r.Row}`,
      owned: true,
      onEdit: (row, key, value) => {
        if (!SELECTION_COLUMNS.includes(key)) return;
        const plot = `${row.Range}_${row.Row}`;
        state.dateRecording[plot] = {
          ...(state.dateRecording[plot] ?? {}), [key]: value,
        };
        save();
      },
      onExport: (api) => runExport("Export", () =>
        exportGrid(state.code, "Date recording", api)),
    });
    registerSheet("Date recording", () =>
      sheetOf("Date recording", node.gridApi));

    holder.replaceChildren(
      el("p", { class: "sub" },
        (mode === "vertical"
          ? "Serpentine – Vertical: two rows together, snaking by range."
          : "Serpentine – Horizontal: bay by bay.") +
        ` ${rows.length} A-line plot(s).` +
        (state.recordingStart
          ? ` Recording started ${state.recordingStart}.`
          : " Set the recording start date to fill in the months.")),
      node);
  };

  main.append(el("div", { class: "btnrow" },
    el("button", {
      class: "action",
      onclick: () => { mode = "vertical"; draw(); },
    }, "Serpentine – Vertical"),
    el("button", {
      class: "action",
      onclick: () => { mode = "horizontal"; draw(); },
    }, "Serpentine – Horizontal"),
    el("button", {
      class: "action ghost",
      onclick: () => pullOutBags(main),
    }, "Pull out bags"),
    el("button", {
      class: "action ghost",
      onclick: () => generateTrend(main),
    }, "Generate graph")));

  main.append(el("div", { class: "btnrow" },
    el("button", {
      class: "action ghost",
      onclick: async () => {
        const start = await pickDate("Recording start date",
          state.recordingStart);
        if (!start) return;
        state.recordingStart = start;
        save();
        render();
      },
    }, state.recordingStart
      ? `Recording start: ${state.recordingStart}`
      : "Set recording start date")));

  main.append(holder);
  draw();
};

/** The month a bare day falls in, given the recording start date. */
function monthLabel(day) {
  if (!state.recordingStart) return "";
  const [y, m, d] = state.recordingStart.split("-").map(Number);
  const stamp = assignMonth(day, d, m, y);
  return stamp ? MONTH_NAMES[stamp[1] - 1] : "";
}

/**
 * Pull out bags: which B-lines to collect for a chosen set of days.
 *
 * Selections are recorded against the A-line, but the bags hang on its paired
 * B-line — the row beside it in the same range. So the A-line's S 1..S 4 are
 * copied across, then the days the user ticks are matched and the resulting
 * B-lines listed in horizontal serpentine order, which is the order someone
 * actually walks the field.
 */
function pullOutBags(main) {
  const aLines = aLineRows();
  if (!aLines.length) {
    alert("No A-lines to pull bags for.");
    return;
  }

  pickDays("Which recording days are you pulling bags for?").then((days) => {
    if (!days || !days.length) return;
    const wanted = new Set(days.map(Number));

    // The B-line paired with an A-line is its neighbour in the same range.
    const all = fieldbookRows();
    const byPlot = new Map(all.map((r) => [`${r.Range}:${r.Row}`, r]));
    const partnerOf = (a) =>
      byPlot.get(`${a.Range}:${a.Row + 1}`)?.CMS?.toUpperCase() === "B"
        ? byPlot.get(`${a.Range}:${a.Row + 1}`)
        : (byPlot.get(`${a.Range}:${a.Row - 1}`)?.CMS?.toUpperCase() === "B"
          ? byPlot.get(`${a.Range}:${a.Row - 1}`)
          : null);

    const hits = [];
    const unpaired = [];
    for (const a of aLines) {
      const stored = state.dateRecording?.[`${a.Range}_${a.Row}`] ?? {};
      const picked = SELECTION_COLUMNS
        .filter((c) => wanted.has(Number(stored[c])))
        .map((c) => `${c} = ${stored[c]}`);
      if (!picked.length) continue;

      const b = partnerOf(a);
      if (!b) {
        unpaired.push(`${a.Range}_${a.Row}`);
        continue;
      }
      hits.push({ b, a, picked: picked.join(", ") });
    }

    if (!hits.length) {
      alert("No selections recorded on those days.");
      return;
    }

    // Horizontal serpentine, as the client asked — bay by bay.
    const order = serpentineByRange(hits.map((h) => [h.b.Range, h.b.Row]));
    const byB = new Map(hits.map((h) => [`${h.b.Range}:${h.b.Row}`, h]));
    const rows = order
      .map(([rng, row]) => byB.get(`${rng}:${row}`))
      .filter(Boolean)
      .map((h) => ({
        Range: h.b.Range,
        Row: h.b.Row,
        Plot: h.b.Plot,
        "Material ID": h.b["Material ID"],
        "Source ID": h.b["Source ID"],
        Gen: h.b.Gen,
        "A-line plot": h.a.Plot,
        Selection: h.picked,
      }));

    const node = grid({
      id: "Pull out bags",
      columns: [
        { key: "Range", type: "number", readOnly: true },
        { key: "Row", type: "number", readOnly: true },
        { key: "Plot", type: "text", readOnly: true },
        { key: "Material ID", type: "text", readOnly: true },
        { key: "Source ID", type: "text", readOnly: true },
        { key: "Gen", type: "text", readOnly: true },
        { key: "A-line plot", type: "text", readOnly: true },
        { key: "Selection", type: "text", readOnly: true },
      ],
      rows,
      rowKey: (r) => r.Plot,
      onExport: (api) => runExport("Export", () =>
        exportGrid(state.code, "Pull out bags", api)),
    });

    main.append(el("h1", { style: "margin-top:26px;font-size:16px" },
      `Pull out bags — days ${days.join(", ")}`));
    main.append(el("p", { class: "sub" },
      `${rows.length} B-line bag(s) to pull, in walking order.` +
      (unpaired.length
        ? ` ${unpaired.length} A-line(s) had no paired B-line and were ` +
          `skipped: ${unpaired.slice(0, 5).join(", ")}` +
          (unpaired.length > 5 ? "…" : "")
        : "")));
    main.append(node);
    node.scrollIntoView({ behavior: "smooth", block: "start" });
  });
}

/** Tick days 1-31, as the client asked, rather than typing a range. */
function pickDays(label) {
  return new Promise((resolve) => {
    const chosen = new Set();
    const close = (value) => { backdrop.remove(); resolve(value); };

    const boxes = Array.from({ length: 31 }, (_, i) => i + 1).map((day) => {
      const box = el("input", { type: "checkbox" });
      box.addEventListener("change", () => {
        if (box.checked) chosen.add(day); else chosen.delete(day);
      });
      return el("label", { class: "day-box" }, box, el("span", {}, day));
    });

    const backdrop = el("div", { class: "modal-backdrop" },
      el("div", { class: "modal", style: "max-width:420px" },
        el("h3", {}, label),
        el("div", { class: "day-grid" }, boxes),
        el("div", { class: "btnrow", style: "margin:14px 0 0" },
          el("button", {
            class: "action",
            onclick: () => close([...chosen].sort((a, b) => a - b)),
          }, "Show bags"),
          el("button", { class: "action ghost", onclick: () => close(null) },
            "Cancel"))));

    backdrop.addEventListener("mousedown", (e) => {
      if (e.target === backdrop) close(null);
    });
    document.body.append(backdrop);
  });
}

/**
 * Generate graph: recordings per day, as a line per selection pass.
 *
 * Ported from GenerateS1S2TrendAnalysis. The VBA asks for the start year,
 * month and day as three separate prompts; one calendar is the same answer.
 */
function generateTrend(main) {
  if (!state.recordingStart) {
    alert("Set the recording start date first — the months are worked out " +
      "from it.");
    return;
  }
  const [year, month, day] = state.recordingStart.split("-").map(Number);
  if (year < 1900 || year > 2100) {
    alert("Recording start year must be between 1900 and 2100.");
    return;
  }

  const rows = aLineRows().map((r) =>
    state.dateRecording?.[`${r.Range}_${r.Row}`] ?? {});
  const series = trendCounts(rows, SELECTION_COLUMNS, day, month, year);

  const recorded = series.reduce((n, [, counts]) =>
    n + Object.values(counts).reduce((a, b) => a + b, 0), 0);
  if (!recorded) {
    alert("No selections recorded yet, so there is nothing to plot.");
    return;
  }

  main.append(el("h1", { style: "margin-top:26px;font-size:16px" },
    "S 1 – S 4 recording trend"));
  main.append(trendChart(series));

  const node = grid({
    id: "Trend Analysis",
    columns: [
      { key: "Date", type: "text", readOnly: true },
      ...SELECTION_COLUMNS.map((c) => ({
        key: `${c} Count`, type: "number", readOnly: true,
      })),
    ],
    rows: series.map(([[y, m, d], counts]) => ({
      Date: `${String(d).padStart(2, "0")}-${MONTH_NAMES[m - 1]}-${y}`,
      ...Object.fromEntries(
        SELECTION_COLUMNS.map((c) => [`${c} Count`, counts[c]])),
    })),
    rowKey: (r) => r.Date,
    onExport: (api) => runExport("Export", () =>
      exportGrid(state.code, "Trend Analysis", api)),
  });
  registerSheet("Trend Analysis", () => sheetOf("Trend Analysis", node.gridApi));
  main.append(node);
  main.lastChild.scrollIntoView({ behavior: "smooth", block: "center" });
}

/**
 * The trend as inline SVG.
 *
 * Drawn by hand rather than pulled in as a chart library — the app ships
 * offline with no build step, and this is one line per series.
 */
function trendChart(series) {
  const W = 900;
  const H = 320;
  const PAD = { top: 16, right: 16, bottom: 54, left: 52 };
  const plotW = W - PAD.left - PAD.right;
  const plotH = H - PAD.top - PAD.bottom;

  const maxCount = Math.max(1, ...series.flatMap(([, c]) =>
    Object.values(c)));
  const x = (i) => PAD.left
    + (series.length > 1 ? (i / (series.length - 1)) * plotW : plotW / 2);
  // Value axis starts at zero, as the VBA sets MinimumScale = 0.
  const y = (v) => PAD.top + plotH - (v / maxCount) * plotH;

  const svgEl = (tag, attrs = {}, ...kids) => {
    const node = document.createElementNS("http://www.w3.org/2000/svg", tag);
    for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, v);
    for (const kid of kids.flat()) {
      if (kid === null || kid === undefined) continue;
      node.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
    }
    return node;
  };

  const colours = ["#1f4e79", "#c00000", "#548235", "#bf8f00"];
  const parts = [];

  // Horizontal gridlines and the value axis.
  const ticks = Math.min(maxCount, 5);
  for (let t = 0; t <= ticks; t++) {
    const v = Math.round((maxCount / ticks) * t);
    parts.push(svgEl("line", {
      x1: PAD.left, x2: W - PAD.right, y1: y(v), y2: y(v),
      stroke: "#e5e7eb",
    }));
    parts.push(svgEl("text", {
      x: PAD.left - 8, y: y(v) + 4, "text-anchor": "end",
      "font-size": "11", fill: "#6b7280",
    }, v));
  }

  // Date labels, thinned so they stay readable over a long season.
  const step = Math.max(1, Math.ceil(series.length / 12));
  series.forEach(([[, m, d]], i) => {
    if (i % step) return;
    parts.push(svgEl("text", {
      x: x(i), y: H - PAD.bottom + 18, "text-anchor": "middle",
      "font-size": "10", fill: "#6b7280",
      transform: `rotate(-45 ${x(i)} ${H - PAD.bottom + 18})`,
    }, `${d} ${MONTH_NAMES[m - 1]}`));
  });

  SELECTION_COLUMNS.forEach((column, ci) => {
    const points = series.map(([, counts], i) => [x(i), y(counts[column])]);
    if (!points.length) return;
    parts.push(svgEl("polyline", {
      fill: "none", stroke: colours[ci], "stroke-width": "2",
      points: points.map(([px, py]) => `${px},${py}`).join(" "),
    }));
    // Markers, as the VBA uses xlLineMarkers.
    for (const [px, py] of points) {
      parts.push(svgEl("circle", { cx: px, cy: py, r: "2.5", fill: colours[ci] }));
    }
  });

  const legend = el("div", { class: "chart-legend" },
    SELECTION_COLUMNS.map((c, i) => el("span", { class: "legend-chip" },
      el("span", {
        class: "legend-swatch",
        style: `background:${colours[i]}`,
      }), c)));

  const svg = svgEl("svg", {
    viewBox: `0 0 ${W} ${H}`, width: "100%", role: "img",
    "aria-label": "Recordings per day for each selection pass",
  },
  svgEl("line", {
    x1: PAD.left, x2: PAD.left, y1: PAD.top, y2: H - PAD.bottom,
    stroke: "#9ca3af",
  }),
  svgEl("line", {
    x1: PAD.left, x2: W - PAD.right, y1: H - PAD.bottom, y2: H - PAD.bottom,
    stroke: "#9ca3af",
  }),
  svgEl("text", {
    x: W / 2, y: H - 6, "text-anchor": "middle", "font-size": "12",
    fill: "#374151",
  }, "Date"),
  svgEl("text", {
    x: 14, y: H / 2, "text-anchor": "middle", "font-size": "12",
    fill: "#374151", transform: `rotate(-90 14 ${H / 2})`,
  }, "Number of Records"),
  parts);

  return el("div", { class: "chart-wrap" }, svg, legend);
}

// Operations and Comments are the same shape: growth stages down the side,
// groups across. Cells take multiple lines, and extra group columns can be
// added from the grid toolbar.
function groupedView(main, tabName, title, storeKey) {
  main.append(...pageHead(title,
    "Cells take multiple lines. Add group columns from the toolbar."));

  // Stored as "<stage>|<group>", so the grid's row objects are built from the
  // map on the way in and written back on the way out.
  const rows = GROWTH_STAGES.map(([stage]) => {
    const row = { Stage: stage };
    for (const g of [1, 2, 3]) {
      row[`Group ${g}`] = state[storeKey][`${stage}|${g}`] ?? "";
    }
    return row;
  });

  const columns = [
    // The heading is deliberately blank — the stage names speak for themselves.
    { key: "Stage", label: "", type: "text", readOnly: true },
    { key: "Group 1", type: "multiline" },
    { key: "Group 2", type: "multiline" },
    { key: "Group 3", type: "multiline" },
  ];

  const node = grid({
    id: tabName,
    columns,
    rows,
    rowKey: (r) => r.Stage,
    owned: true,
    onEdit: (row, key, value) => {
      const group = key.startsWith("Group ") ? key.slice(6) : null;
      if (group) state[storeKey][`${row.Stage}|${group}`] = value;
      save();
    },
    cellColour: (row, key) =>
      (key === "Stage" ? STAGE_COLOURS.get(row.Stage) ?? null : null),
    onExport: (api) => runExport("Export", () =>
      exportGrid(state.code, tabName, api)),
  });

  registerSheet(tabName, () => sheetOf(tabName, node.gridApi));
  main.append(node);
}

VIEWS.Operations = (main) =>
  groupedView(main, "Operations", "Operations Overview", "operations");
VIEWS.Comments = (main) =>
  groupedView(main, "Comments", "Field Comments", "comments");

function viewMissing(main, name) {
  main.append(...pageHead(name, "Nursery-type extra."));
  main.append(emptyState(
    `The ${name} tab is generated in the Excel workbook. It is not yet ` +
    "implemented in the desktop app."));
}

// ---------------------------------------------------------------- actions

function generateAll() {
  if (!state.prism.length) {
    alert("Import the PRISM export on the Nursery site tab first.");
    return;
  }
  save();
  alert(`${state.prism.length} packets read from PRISM.\n\n` +
    "Material Map, Field Map, Packet Prep and Nursery list are built from " +
    "this data and update live.\n\n" +
    "Fieldbook is built later, from the Updated nursery site tab.");
  activeTab = "Material Map";
  render();
}

// Initialising always opens a *new* nursery. Editing the open one is what the
// Rename action is for — before, this overwrote whatever was already loaded.
function initNursery() {
  const code = prompt("Enter the nursery code (e.g. AUGT1-26S-IMI):", "");
  if (!code) return;
  const types = prompt(
    "Nursery type(s) — comma separated.\nOptions: " + SPEC.nursery_types.join(", "),
    "Selection");
  if (!types) return;
  const fileName = prompt("File name for this workbook:", code);
  if (fileName === null) return;

  store.createNursery({
    code: code.trim(),
    types: types.split(",").map((s) => s.trim()).filter(Boolean),
    fileName: fileName.trim(),
  });
  reloadActive();
}

$("#btnInit").addEventListener("click", initNursery);
$("#btnReset").addEventListener("click", () => {
  const name = state.code || "this untitled nursery";
  if (!confirm(
    `Clear the data in ${name}? Other nurseries on this machine are kept.`)) {
    return;
  }
  store.deleteNursery(store.activeId());
  reloadActive();
});

render();
