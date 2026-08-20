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
} from "./nursery-algos.js";
import * as store from "./store.js";
import { grid } from "./grid.js";

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

const GROWTH_STAGES = [
  ["Seedling", "#fec000"], ["Vegetative", "#a9d08e"], ["Heading", "#ff6b6b"],
  ["Flowering", "#add8e6"], ["Grain filling", "#ffff99"],
  ["Harvest", "#ccffcc"], ["Post Harvest", "#ccccff"],
];

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

function isRecurrent(gen) {
  const g = String(gen ?? "").toUpperCase();
  return g.startsWith("BC") || /^F\d/.test(g);
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

VIEWS.Home = (main) => {
  main.append(...pageHead("Nursery Workflow",
    "Run each step in order. The same steps as the workbook's Home tab."));

  main.append(nurseryPicker());

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

function buildMapGrid(valueKeys) {
  const maxRange = Math.max(...state.prism.map((r) => num(r.Range)));
  const maxRow = Math.max(...state.prism.map((r) => num(r.Row)));
  const cell = new Map();
  for (const r of state.prism) {
    cell.set(`${num(r.Range)}:${num(r.Row)}`,
      valueKeys.map((k) => String(r[k] ?? "").slice(0, 14))
        .filter(Boolean).join("\n"));
  }

  const rowNums = Array.from({ length: maxRow }, (_, i) => i + 1);
  const head = ["Rng \\ Row", ...rowNums, "Rng"];
  const body = [];
  // Ranges descend so the top range prints at the top, as in the workbook.
  for (let rng = maxRange; rng >= 1; rng--) {
    body.push([rng, ...rowNums.map((w) => cell.get(`${rng}:${w}`) ?? ""), rng]);
  }
  // Row numbers repeat along the foot, as the client asked.
  body.push(["Rng \\ Row", ...rowNums, "Rng"]);

  const tbl = table(head, body, {
    cellClass: (i, j) =>
      (j === 0 || j === maxRow + 1 || i === body.length - 1 ? "hdr" : ""),
  });
  tbl.querySelector("table").classList.add("grid");
  return tbl;
}

VIEWS["Material Map"] = (main) => {
  main.append(...pageHead("Material Map",
    "Range numbers down both sides, field rows across top and bottom."));
  if (!state.prism.length) {
    main.append(emptyState("Import Nursery site data first."));
    return;
  }
  main.append(buildMapGrid(["Material ID", "Inbred Code", "Hybrid Code"]));
};

VIEWS["Field Map"] = (main) => {
  main.append(...pageHead("Field Map",
    "Same grid as Material Map, without material information."));

  main.append(el("div", { class: "btnrow" },
    el("button", { class: "action", onclick: fieldMapWizard },
      "Set planting dates and rows"),
    state.fieldMap.length
      ? el("button", {
        class: "action ghost",
        onclick: () => {
          state.fieldMap = [];
          state.plantingDates = 1;
          save();
          render();
        },
      }, "Clear")
      : null));

  if (state.fieldMap.length) {
    const dops = [...new Set(state.fieldMap.map((e) => e.dop))];
    main.append(el("p", { class: "sub" },
      `${dops.length} planting date(s), ${state.fieldMap.length} rows assigned. ` +
      `Seed qty/plot: ${state.seedQty || "—"}`));
    main.append(table(
      ["Planting date", "Row", "Spike", "Run", "Split no.", "Seed qty/plot"],
      state.fieldMap.map((e) => [e.dop, e.row, e.spike, e.run,
        dops.indexOf(e.dop) + 1, e.qty]),
      { cellClass: (i) => (state.fieldMap[i].run === "forward" ? "fwd" : "rev") }));
  } else {
    main.append(emptyState(
      "No planting dates set. Spike numbers follow the 1,2,2,1 cycle and runs " +
      "alternate every two rows, starting forward."));
  }

  if (state.prism.length) {
    main.append(el("h1", { style: "margin-top:24px;font-size:16px" }, "Grid"));
    main.append(buildMapGrid([]));
  }
};

function fieldMapWizard() {
  if (!state.prism.length) {
    alert("Import Nursery site data first.");
    return;
  }
  const n = num(prompt("How many planting dates?",
    String(state.plantingDates || 2)));
  if (n < 1) return;

  const qty = prompt("Seed quantity per plot:", state.seedQty || "");
  if (qty === null) return;

  const dopRows = [];
  const dops = [];
  for (let i = 1; i <= n; i++) {
    const dop = prompt(`Date of planting ${i} (e.g. 25-Feb-2026):`);
    if (!dop) return;
    const rowsCsv = prompt(`Rows for planting date ${i} (comma separated):`);
    if (!rowsCsv) return;
    dops.push(dop);
    dopRows.push(rowsCsv.split(",").map((s) => num(s)).filter((r) => r > 0));
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
  main.append(table(["", ""],
    SPEC.nursery_data_labels.map((label) => [label, ""])));
};

VIEWS["Packet Prep"] = (main, tabName) => {
  const splitNo = num(tabName.split(" ").pop()) || 1;
  main.append(...pageHead(tabName,
    `Packets for planting date ${splitNo}. ` +
    "QR CODE column feeds the barcode printer."));

  if (!state.prism.length) {
    main.append(emptyState("Import Nursery site data first."));
    return;
  }

  const rows = state.prism
    .filter((r) => splitForRow(num(r.Row)) === splitNo)
    .map((r) => {
      const rng = num(r.Range);
      const row = num(r.Row);
      return [qrText(r), rng, row, plotOf(r), rng, row, splitNo,
        r["Material ID"], r["Inbred Code"], r["Source ID"],
        r["CMS reaction"], r.Generation];
    });

  main.append(el("p", { class: "sub" }, `${rows.length} packets in this split.`));
  main.append(table(
    ["QR CODE", "Range", "Row", "Plot", "SPIKE#", "RACK ORDER", "Split no.",
      "Material ID", "Inbred Code", "Source ID", "CMS reaction", "Generation"],
    rows));
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

  const holder = el("div");
  const draw = () => {
    holder.replaceChildren(grid({
      id: "Nursery list",
      columns: [
        { key: "Source ID", type: "text" },
        { key: "Repeats", type: "number" },
        { key: "Qty Required", type: "number" },
        { key: "Inbred Code", type: "text" },
        { key: "Hybrid Code", type: "text" },
      ],
      rows,
      // Source ID is unique here by construction, so it identifies the row.
      rowKey: (r) => r["Source ID"],
      cellColour: (row, key, value) => {
        if (!showDups) return null;
        if (key === "Inbred Code" && dupInbred.has(value)) return "#ffd9d9";
        if (key === "Hybrid Code" && dupHybrid.has(value)) return "#ffd9d9";
        return null;
      },
    }));
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
function logView(main, { title, subtitle, storeKey, columns }) {
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
    }, "Read QR → Plot")));

  if (!state[storeKey].length) {
    main.append(emptyState(`Nothing logged on ${title}.`));
    return;
  }

  const head = el("tr", {},
    columns.map((c) => el("th", {}, c.key)).concat([el("th", {}, "")]));

  const body = state[storeKey].map((rec, i) => {
    const cells = columns.map((c) => {
      if (c.type === "select") {
        return el("td", {}, el("select", {
          class: "f",
          onchange: (e) => { rec[c.key] = e.target.value; save(); },
        }, c.options.map((o) => el("option",
          o === rec[c.key] ? { value: o, selected: "selected" } : { value: o },
          o))));
      }
      return el("td", {}, el("input", {
        class: "f",
        value: rec[c.key] ?? "",
        // The QR payload is long; give it room so the column fits its contents.
        style: c.wide ? "min-width:340px" : "",
        oninput: (e) => { rec[c.key] = e.target.value; save(); },
      }));
    });
    cells.push(el("td", {}, el("button", {
      class: "action ghost",
      onclick: () => { state[storeKey].splice(i, 1); save(); render(); },
    }, "Remove")));
    return el("tr", {}, cells);
  });

  main.append(el("div", { class: "scroll" },
    el("table", { class: "fit" },
      el("thead", {}, head), el("tbody", {}, body))));
}

VIEWS.Replacements = (main) => logView(main, {
  title: "Replacements",
  subtitle: "Scan the original entry and its replacement. The export expands " +
    "each QR into the full entry block.",
  storeKey: "replacements",
  columns: REPLACEMENT_COLUMNS,
});

VIEWS["Planting errors"] = (main) => logView(main, {
  title: "Planting errors",
  subtitle: "Original entry only — nothing was planted in its place.",
  storeKey: "plantingErrors",
  columns: PLANTING_ERROR_COLUMNS,
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

VIEWS.Fieldbook = (main) => {
  main.append(...pageHead("Fieldbook",
    "Built from Updated nursery site. Range, Row, Material ID, Source ID, " +
    "Gen, CMS, Plot, Comments."));

  const src = state.updatedPrism;
  if (!src.length) {
    main.append(emptyState(
      "Import the updated PRISM export on the ‘Updated nursery site’ tab first."));
    return;
  }

  const plots = src.map((r) => [num(r.Range), num(r.Row)]);
  const order = serpentineByRange(plots);
  const byPlot = new Map(src.map((r) => [`${num(r.Range)}:${num(r.Row)}`, r]));

  const rows = order.map(([rng, row]) => {
    const r = byPlot.get(`${rng}:${row}`) || {};
    return [rng, row, r["Material ID"], r["Source ID"], r.Generation,
      r["CMS reaction"], `${rng}_${row}`, r.Comments];
  });

  const ab = isType("AB");
  main.append(el("p", { class: "sub" },
    `${rows.length} plots, serpentine order (rows descend on even ranges).` +
    (ab ? " B lines shaded green." : "")));
  main.append(table(
    ["Range", "Row", "Material ID", "Source ID", "Gen", "CMS", "Plot", "Comments"],
    rows,
    {
      cellClass: (i) =>
        (ab && String(rows[i][5]).toUpperCase() === "B" ? "bline" : ""),
    }));

  main.append(el("div", { class: "note" },
    "Printing: landscape A4, duplex flipped on the short edge, page numbers " +
    "centred, file name top right. Duplex is a printer-driver setting — " +
    "choose it in the print dialog."));
  main.append(el("div", { class: "btnrow" },
    el("button", { class: "action", onclick: () => window.print() }, "Print")));
};

VIEWS["Date recording"] = (main) => {
  main.append(...pageHead("Date recording",
    "AB nurseries only. S 1 and S 2 hold day-of-month numbers; " +
    "the month is applied separately."));

  const src = state.prism.filter((r) => isRecurrent(r.Generation));
  if (!src.length) {
    main.append(emptyState("No recurrent (BC*/Fn) packets in Nursery site."));
    return;
  }

  let mode = "range";
  const holder = el("div");
  const draw = () => {
    holder.replaceChildren();
    const plots = src.map((r) => [num(r.Range), num(r.Row)]);
    const order = mode === "band"
      ? serpentineTwoRowBands(plots)
      : serpentineByRange(plots);
    const byPlot = new Map(src.map((r) => [`${num(r.Range)}:${num(r.Row)}`, r]));
    const rows = order.map(([rng, row]) => {
      const r = byPlot.get(`${rng}:${row}`) || {};
      return [bandForRow(row), rng, row, parityForRow(row), "", "",
        r["Material ID"], r["Source ID"], r.Generation, r["CMS reaction"],
        r["Inbred Code"]];
    });
    holder.append(el("p", { class: "sub" },
      mode === "band"
        ? "Ordered 2 rows together, snaking by range within each band."
        : "Ordered by range, snaking through rows."));
    holder.append(table(
      ["Group", "Range", "Row", "O/E", "S 1", "S 2", "Material ID",
        "Source ID", "Gen", "CMS", "In. Code"], rows));
  };

  main.append(el("div", { class: "btnrow" },
    el("button", {
      class: "action",
      onclick: () => { mode = "range"; draw(); },
    }, "Record by range and pull out bags"),
    el("button", {
      class: "action ghost",
      onclick: () => { mode = "band"; draw(); },
    }, "Record 2-rows together")));
  main.append(holder);
  draw();
};

function groupedView(main, title, store) {
  main.append(...pageHead(title, ""));
  const head = el("tr", {},
    ["Stage", "Group 1", "Group 2", "Group 3"].map((h) => el("th", {}, h)));
  const body = GROWTH_STAGES.map(([stage, colour]) =>
    el("tr", {},
      el("td",
        { style: `background:${colour};font-weight:700;text-align:center` },
        stage),
      [1, 2, 3].map((g) => el("td", {},
        el("input", {
          class: "f",
          value: state[store][`${stage}|${g}`] ?? "",
          oninput: (e) => { state[store][`${stage}|${g}`] = e.target.value; save(); },
        })))));
  main.append(el("div", { class: "scroll" },
    el("table", {}, el("thead", {}, head), el("tbody", {}, body))));
}

VIEWS.Operations = (main) => groupedView(main, "Operations Overview", "operations");
VIEWS.Comments = (main) => groupedView(main, "Field Comments", "comments");

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
