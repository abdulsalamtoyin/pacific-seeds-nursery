// The editable grid every tab uses.
//
// The client asked for the same handful of things on every table — edit a cell,
// sort a column, filter it like Excel, colour cells, add rows and columns,
// download it. Doing that per tab would be twelve implementations that drift.
// This is the one implementation; each view describes its columns and hands
// over its rows.
//
// Two kinds of table use it:
//
//   owned     — the rows ARE the data (Replacements, Operations). An edit
//               writes straight into the row object and the caller saves.
//   computed  — the rows are rebuilt from PRISM on every render (Packet Prep,
//               Fieldbook, Material Map). An edit cannot be written back into
//               a row that is about to be thrown away, so it is stored as an
//               overlay keyed by a stable row identity and re-applied after
//               each rebuild. Overlaid cells are marked, and can be reverted.
//
// That overlay is why `rowKey` matters: it must identify the row by something
// that survives a re-import (the plot), never by position.

import * as store from "./store.js";
import {
  compare, defaultedValue, effectiveValue, filterRecords, isEditedCell,
  isOverlaid, nextSort, orphanKeys, sortRecords, todayISO,
} from "./grid-core.js";

function el(tag, attrs = {}, ...kids) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") node.className = v;
    else if (k === "style") node.setAttribute("style", v);
    else if (k.startsWith("on")) node.addEventListener(k.slice(2), v);
    else if (v !== null && v !== undefined) node.setAttribute(k, v);
  }
  for (const kid of kids.flat()) {
    if (kid === null || kid === undefined || kid === false) continue;
    node.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
  }
  return node;
}

let userColumnSeq = 0;

/**
 * Build a grid.
 *
 * @param {object}   cfg
 * @param {string}   cfg.id       persistence key — the tab name
 * @param {Array}    cfg.columns  [{key, label?, type?, options?, width?, readOnly?}]
 * @param {Array}    cfg.rows     row objects
 * @param {Function} cfg.rowKey   (row, index) => stable identity string
 * @param {boolean}  cfg.owned    true when rows are the source of truth
 * @param {Function} cfg.onEdit   owned mode: (row, columnKey, value) => void
 * @param {Function} cfg.onAddRow owned mode: () => void, appends a row
 * @param {Function} cfg.onDeleteRow owned mode: (row) => void
 * @param {Function} cfg.onExport optional; shows an Export button
 * @param {Function} cfg.onRerender called when the grid needs the view redrawn
 * @param {boolean}  cfg.toolbar  default true
 */
export function grid(cfg) {
  const {
    id,
    columns: baseColumns,
    rows: sourceRows,
    rowKey = (_row, i) => String(i),
    owned = false,
    onEdit,
    onAddRow,
    onDeleteRow,
    onExport,
    onRerender,
    cellColour,
    toolbar = true,
  } = cfg;

  const gs = store.gridState(id);
  const root = el("div", { class: "grid-wrap" });

  // Selection is a set of "rowKey|columnKey" ids; colour and style act on it.
  let selected = new Set();
  let anchor = null;

  const persist = (patch) => {
    Object.assign(gs, patch);
    store.setGridState(id, patch);
  };

  // ------------------------------------------------------------ columns

  function allColumns() {
    const extras = gs.extraColumns.map((c) => ({ ...c, userAdded: true }));
    return [...baseColumns, ...extras];
  }

  function visibleColumns() {
    return allColumns().filter((c) => !gs.hidden.includes(c.key));
  }

  // ------------------------------------------------------------ rows

  // Wrap each row so display code never has to care where its value lives.
  function records() {
    const base = sourceRows.map((row, i) => ({
      key: String(rowKey(row, i)),
      src: row,
      userAdded: false,
    }));
    const extra = gs.extraRows.map((row) => ({
      key: row.__key,
      src: row,
      userAdded: true,
    }));
    return [...base, ...extra];
  }

  // A user-added row and an owned row hold their own values; a computed row's
  // hand-edits live in the overlay. The rules themselves are in grid-core.js.
  const overlaid = (rec) => isOverlaid(rec, owned);
  const isEdited = (rec, key) =>
    isEditedCell(rec, key, { owned, edits: gs.edits });

  // Blank date cells read as today. Resolving that here — rather than only
  // when painting the input — keeps the value shown, filtered, sorted and
  // exported identical.
  const today = todayISO();
  const columnFor = (key) => allColumns().find((c) => c.key === key);
  const valueOf = (rec, key) => defaultedValue(
    effectiveValue(rec, key, { owned, edits: gs.edits }), columnFor(key), today);

  function setValue(rec, key, value) {
    if (overlaid(rec)) {
      const edits = {
        ...gs.edits,
        [rec.key]: { ...(gs.edits[rec.key] ?? {}), [key]: value },
      };
      persist({ edits });
      return;
    }
    rec.src[key] = value;
    if (rec.userAdded) persist({ extraRows: gs.extraRows });
    else onEdit?.(rec.src, key, value);
  }

  function revert(rec, key) {
    const patch = { ...(gs.edits[rec.key] ?? {}) };
    delete patch[key];
    const edits = { ...gs.edits };
    if (Object.keys(patch).length) edits[rec.key] = patch;
    else delete edits[rec.key];
    persist({ edits });
    redraw();
  }

  // An edit whose row has vanished — usually a re-import that dropped a plot.
  // Keeping it silent would lose work, so it is counted and shown.
  function orphanedEdits() {
    return orphanKeys(gs.edits, records().map((r) => r.key));
  }

  // ------------------------------------------------------------ filter/sort

  function distinct(key) {
    const seen = new Set(records().map((rec) => String(valueOf(rec, key) ?? "")));
    return [...seen].sort(compare);
  }

  function displayRecords() {
    return sortRecords(
      filterRecords(records(), gs.filters, valueOf), gs.sort, valueOf);
  }

  function cycleSort(key) {
    persist({ sort: nextSort(gs.sort, key) });
    redraw();
  }

  // Excel-style dropdown: every distinct value as a checkbox, plus a search
  // box for columns with hundreds of values.
  function filterMenu(column, anchorEl) {
    root.querySelectorAll(".filter-menu").forEach((m) => m.remove());

    const values = distinct(column.key);
    const allowed = new Set(gs.filters[column.key] ?? values);
    const menu = el("div", { class: "filter-menu" });
    const list = el("div", { class: "filter-list" });

    const draw = (needle = "") => {
      list.replaceChildren();
      for (const v of values.filter((x) =>
        x.toLowerCase().includes(needle.toLowerCase()))) {
        const box = el("input", { type: "checkbox" });
        box.checked = allowed.has(v);
        box.addEventListener("change", () => {
          if (box.checked) allowed.add(v); else allowed.delete(v);
        });
        list.append(el("label", { class: "filter-row" }, box,
          el("span", {}, v === "" ? "(blank)" : v)));
      }
    };

    const search = el("input", {
      class: "f",
      placeholder: "Search…",
      oninput: (e) => draw(e.target.value),
    });

    menu.append(
      search,
      list,
      el("div", { class: "filter-actions" },
        el("button", {
          class: "action ghost",
          onclick: () => { values.forEach((v) => allowed.add(v)); draw(search.value); },
        }, "All"),
        el("button", {
          class: "action ghost",
          onclick: () => { allowed.clear(); draw(search.value); },
        }, "None"),
        el("button", {
          class: "action",
          onclick: () => {
            const filters = { ...gs.filters };
            // Everything ticked is the same as no filter — keep state clean so
            // the marker only shows when a column is really filtered.
            if (allowed.size === values.length) delete filters[column.key];
            else filters[column.key] = [...allowed];
            persist({ filters });
            menu.remove();
            redraw();
          },
        }, "Apply")));

    draw();
    anchorEl.append(menu);
  }

  // ------------------------------------------------------------ selection

  function cellId(rec, key) { return `${rec.key}|${key}`; }

  // A user's fill always beats the view's own shading (duplicate codes,
  // B lines, planting-date bands), so colouring a cell by hand sticks.
  function applySelectionStyle(node, cid, autoColour) {
    const style = gs.styles[cid] ?? {};
    node.style.background = gs.colours[cid] ?? autoColour ?? "";
    node.style.fontWeight = style.bold ? "700" : "";
    node.style.fontStyle = style.italic ? "italic" : "";
  }

  function selectCell(rec, key, ev) {
    const cid = cellId(rec, key);
    const recs = displayRecords();
    const cols = visibleColumns();

    if (ev?.shiftKey && anchor) {
      // Rectangle between the anchor and this cell.
      const rowIdx = recs.findIndex((r) => r.key === rec.key);
      const colIdx = cols.findIndex((c) => c.key === key);
      const aRow = recs.findIndex((r) => r.key === anchor.rowKey);
      const aCol = cols.findIndex((c) => c.key === anchor.colKey);
      if (rowIdx >= 0 && colIdx >= 0 && aRow >= 0 && aCol >= 0) {
        selected = new Set();
        for (let r = Math.min(rowIdx, aRow); r <= Math.max(rowIdx, aRow); r++) {
          for (let c = Math.min(colIdx, aCol); c <= Math.max(colIdx, aCol); c++) {
            selected.add(cellId(recs[r], cols[c].key));
          }
        }
      }
    } else if (ev?.metaKey || ev?.ctrlKey) {
      if (selected.has(cid)) selected.delete(cid); else selected.add(cid);
      anchor = { rowKey: rec.key, colKey: key };
    } else {
      selected = new Set([cid]);
      anchor = { rowKey: rec.key, colKey: key };
    }
    markSelection();
  }

  function markSelection() {
    root.querySelectorAll("td[data-cid]").forEach((td) => {
      td.classList.toggle("sel", selected.has(td.dataset.cid));
    });
  }

  function styleSelection(patch) {
    if (!selected.size) {
      alert("Select one or more cells first.");
      return;
    }
    const colours = { ...gs.colours };
    const styles = { ...gs.styles };
    for (const cid of selected) {
      if ("colour" in patch) {
        if (patch.colour) colours[cid] = patch.colour;
        else delete colours[cid];
      }
      if ("bold" in patch || "italic" in patch) {
        const next = { ...(styles[cid] ?? {}), ...patch };
        delete next.colour;
        if (next.bold || next.italic) styles[cid] = next;
        else delete styles[cid];
      }
    }
    persist({ colours, styles });
    redraw();
  }

  // ------------------------------------------------------------ toolbar

  function addColumn() {
    const label = prompt("Name for the new column:");
    if (!label) return;
    userColumnSeq += 1;
    const key = `user:${label}:${Date.now().toString(36)}${userColumnSeq}`;
    persist({ extraColumns: [...gs.extraColumns, { key, label, type: "text" }] });
    redraw();
  }

  function addRow() {
    if (owned && onAddRow) { onAddRow(); return; }
    const row = {
      __key: `extra:${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
    };
    persist({ extraRows: [...gs.extraRows, row] });
    redraw();
  }

  function deleteSelectedRows() {
    if (!selected.size) { alert("Select a cell in the row(s) to delete."); return; }
    const keys = new Set([...selected].map((cid) => cid.split("|")[0]));
    const recs = records().filter((r) => keys.has(r.key));
    const fromSource = recs.filter((r) => !r.userAdded);

    if (fromSource.length && !owned) {
      alert(
        `${fromSource.length} of the selected row(s) are built from imported ` +
        "data and cannot be deleted here. Rows you added yourself were removed.");
    }
    if (owned && onDeleteRow) fromSource.forEach((r) => onDeleteRow(r.src));

    persist({ extraRows: gs.extraRows.filter((r) => !keys.has(r.__key)) });
    selected = new Set();
    redraw();
  }

  function deleteSelectedColumns() {
    if (!selected.size) { alert("Select a cell in the column(s) to delete."); return; }
    const keys = new Set([...selected].map((cid) => cid.split("|").slice(1).join("|")));
    const userKeys = gs.extraColumns.filter((c) => keys.has(c.key)).map((c) => c.key);
    const builtIn = [...keys].filter((k) => !userKeys.includes(k));

    // Built-in columns are hidden rather than destroyed — the data behind them
    // is still needed by exports and by other tabs.
    if (builtIn.length) {
      persist({ hidden: [...new Set([...gs.hidden, ...builtIn])] });
    }
    if (userKeys.length) {
      persist({
        extraColumns: gs.extraColumns.filter((c) => !userKeys.includes(c.key)),
      });
    }
    selected = new Set();
    redraw();
  }

  function buildToolbar() {
    return el("div", { class: "grid-toolbar" },
      el("button", { class: "action ghost", onclick: addRow }, "+ Row"),
      el("button", { class: "action ghost", onclick: addColumn }, "+ Column"),
      el("button", { class: "action ghost", onclick: deleteSelectedRows },
        "Delete row"),
      el("button", { class: "action ghost", onclick: deleteSelectedColumns },
        "Hide/delete column"),
      el("span", { class: "sep" }),
      el("label", { class: "swatch", title: "Fill colour" },
        el("input", {
          type: "color",
          value: "#ffff00",
          oninput: (e) => styleSelection({ colour: e.target.value }),
        }), "Fill"),
      el("button", {
        class: "action ghost",
        onclick: () => styleSelection({ colour: null }),
      }, "No fill"),
      el("button", {
        class: "action ghost",
        style: "font-weight:700",
        onclick: () => styleSelection({ bold: !(gs.styles[[...selected][0]]?.bold) }),
      }, "B"),
      el("button", {
        class: "action ghost",
        style: "font-style:italic",
        onclick: () => styleSelection({ italic: !(gs.styles[[...selected][0]]?.italic) }),
      }, "I"),
      el("span", { class: "sep" }),
      gs.hidden.length
        ? el("button", {
          class: "action ghost",
          onclick: () => { persist({ hidden: [] }); redraw(); },
        }, `Show ${gs.hidden.length} hidden column(s)`)
        : null,
      el("button", {
        class: "action ghost",
        onclick: () => {
          if (!confirm(
            "Reset this table's sorting, filters, colours, added rows and " +
            "columns, and hand-edits?")) return;
          store.clearGridState(id);
          Object.assign(gs, store.gridState(id));
          selected = new Set();
          redraw();
        },
      }, "Reset table"),
      onExport
        ? el("button", { class: "action", onclick: () => onExport(api) }, "Export")
        : null);
  }

  // ------------------------------------------------------------ render

  const body = el("div");

  function redraw() {
    body.replaceChildren();

    const cols = visibleColumns();
    const recs = displayRecords();
    const total = records().length;

    const head = el("tr", {}, cols.map((c) => {
      const arrow = gs.sort?.key === c.key
        ? (gs.sort.dir === "asc" ? " ▲" : " ▼")
        : "";
      const th = el("th", {
        style: gs.widths[c.key] ? `width:${gs.widths[c.key]}px` : "",
      },
        el("span", {
          class: "th-label",
          title: "Click to sort",
          onclick: () => cycleSort(c.key),
        }, `${c.label ?? c.key}${arrow}`),
        el("span", {
          class: `th-filter${gs.filters[c.key] ? " on" : ""}`,
          title: "Filter",
          onclick: (e) => { e.stopPropagation(); filterMenu(c, th); },
        }, "▾"));
      return th;
    }));

    const rows = recs.map((rec) => el("tr", {}, cols.map((c) => {
      const cid = cellId(rec, c.key);
      const value = valueOf(rec, c.key);
      const edited = isEdited(rec, c.key);

      const td = el("td", {
        "data-cid": cid,
        class: edited ? "edited" : "",
        title: edited
          ? "Hand-edited. Right-click to revert to the computed value."
          : "",
        onmousedown: (e) => selectCell(rec, c.key, e),
        oncontextmenu: (e) => {
          if (!edited) return;
          e.preventDefault();
          if (confirm("Revert this cell to its computed value?")) revert(rec, c.key);
        },
      });

      let field;
      if (c.readOnly) {
        field = el("span", { class: "ro" }, value);
      } else if (c.type === "select") {
        field = el("select", {
          class: "f",
          onchange: (e) => { setValue(rec, c.key, e.target.value); redraw(); },
        }, (c.options ?? []).map((o) => el("option",
          String(o) === String(value) ? { value: o, selected: "selected" } : { value: o },
          o)));
      } else if (c.type === "multiline") {
        field = el("textarea", {
          class: "f",
          rows: "2",
          oninput: (e) => setValue(rec, c.key, e.target.value),
        });
        field.value = value;
      } else {
        field = el("input", {
          class: "f",
          type: c.type === "date" ? "date" : (c.type === "number" ? "number" : "text"),
          value,
          oninput: (e) => setValue(rec, c.key, e.target.value),
        });
      }
      td.append(field);
      applySelectionStyle(td, cid, cellColour?.(rec.src, c.key, value));
      return td;
    })));

    const orphans = orphanedEdits();
    body.append(
      el("div", { class: "grid-status" },
        `${recs.length} of ${total} row(s)`,
        recs.length !== total
          ? el("button", {
            class: "link",
            onclick: () => { persist({ filters: {} }); redraw(); },
          }, "clear filters")
          : null,
        orphans.length
          ? el("span", { class: "warn" },
            ` ${orphans.length} hand-edit(s) belong to rows no longer in this ` +
            "table — they are kept, not deleted.")
          : null),
      el("div", { class: "scroll" },
        el("table", { class: "grid-table" },
          el("thead", {}, head), el("tbody", {}, rows))));

    markSelection();
  }

  const api = {
    /** Columns as displayed, left to right. */
    columns: () => visibleColumns(),
    /** Rows as displayed — filtered and sorted — as plain objects. */
    rows: () => displayRecords().map((rec) => Object.fromEntries(
      visibleColumns().map((c) => [c.key, valueOf(rec, c.key)]))),
    /** Cell fills, keyed "<row index>,<column index>" over the displayed grid. */
    colours: () => {
      const out = {};
      displayRecords().forEach((rec, r) => visibleColumns().forEach((c, i) => {
        const colour = gs.colours[cellId(rec, c.key)];
        if (colour) out[`${r},${i}`] = colour;
      }));
      return out;
    },
    styles: () => {
      const out = {};
      displayRecords().forEach((rec, r) => visibleColumns().forEach((c, i) => {
        const style = gs.styles[cellId(rec, c.key)];
        if (style) out[`${r},${i}`] = style;
      }));
      return out;
    },
    redraw,
    rerender: () => onRerender?.(),
  };

  if (toolbar) root.append(buildToolbar());
  root.append(body);
  redraw();

  // Clicking away closes any open filter menu.
  document.addEventListener("mousedown", (e) => {
    if (!root.contains(e.target)) {
      root.querySelectorAll(".filter-menu").forEach((m) => m.remove());
    }
  });

  root.gridApi = api;
  return root;
}
