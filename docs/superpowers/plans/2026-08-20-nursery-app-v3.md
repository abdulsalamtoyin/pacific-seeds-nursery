# Nursery app v3 — implementation plan

**Goal:** Apply the ~60 changes in `Toyin- Updated app changes.docx` and the
export formats in `Tab information.xlsx` to the desktop app.

**Architecture:** Three shared modules absorb the recurring asks — an editable
grid (`pwa/grid.js`), an xlsx exporter (`pwa/exporter.js` + `POST /export/xlsx`),
and a multi-nursery store (`pwa/store.js`). Tab-structure changes land in an
`app_overrides` block that only the JS generator merges, leaving the Excel
workbook and its tests untouched.

**Tech stack:** vanilla ES modules (no framework, no build step), FastAPI,
openpyxl, pytest.

**Spec:** `docs/superpowers/specs/2026-08-20-nursery-app-v3-design.md`

## Global constraints

- Desktop app only. Do not modify `excel_workflow/vba/**` or
  `excel_workflow/build_workbooks.py`.
- The existing 72 tests stay green at every commit.
- No build step, no new runtime dependencies in the browser. ES modules only.
- `pwa/nursery-spec.js` is generated — never hand-edit. Change
  `excel_workflow/spec/nursery_spec.json` and re-run
  `python -m excel_workflow.gen_spec_constants`.
- `tabsFor()` (JS) and `tabs_for()` (Python) must stay behaviourally identical;
  `tests/test_js_parity.py` enforces it.
- Nursery data defaults, verbatim: Planter `Almaco Precision Planter`,
  Seeds/side `34`, Plot length `4.5m`, Alley way spacing `0.75m`.
- Nine-column entry block order, verbatim: QR Key, Source ID, Material ID,
  Inbred Code, Experimental Hybrid Code, Pedigree, Generation, CMS reaction,
  W/column.

---

## Phase 1 — Spec overrides and multi-nursery store

### Task 1.1: `after` placement rule

- Files: `excel_workflow/spec/loader.py`, `excel_workflow/gen_spec_constants.py`
  (the `render_js` template), `tests/test_spec.py`, `tests/test_js_parity.py`
- Behaviour: a conditional rule carries exactly one of `before` / `after`.
  `before: X` emits the tab immediately before `X`; `after: X` immediately
  after. A rule with both, or neither, raises.
- Tests: placement before, placement after, both-keys raises, neither raises,
  JS/Python parity across the type × planting-date matrix.

### Task 1.2: `app_overrides` merge

- Files: `excel_workflow/spec/nursery_spec.json`,
  `excel_workflow/gen_spec_constants.py`, `tests/test_spec.py`,
  `tests/test_gen_constants.py`
- `render_js()` deep-merges `spec["app_overrides"]` over the base before
  emitting; `render_constants()` and `load_spec()` consumers ignore it.
- Overrides: `default_tabs` swaps `Replacements and Errors` →
  `Replacements`, `Planting errors`; `conditional["Date recording"]` becomes
  `{"types": ["AB"], "after": "Fieldbook"}`; `date_recording_columns` gains
  `S 3`, `S 4`, `S1 Month`..`S4 Month`; new `nursery_data_defaults`.
- Tests: base spec unchanged for the workbook builder, generated JS contains the
  overridden tab list, `test_workbook_structure.py` still green.

### Task 1.3: multi-nursery store

- Create `pwa/store.js`. Modify `pwa/app.js` (state access), `pwa/index.html`.
- API: `activeNursery()`, `listNurseries()`, `createNursery(code, types,
  fileName)`, `switchNursery(id)`, `renameNursery(id, code)`,
  `duplicateNursery(id)`, `deleteNursery(id)`, `save()`, `gridState(tabName)`.
- Migration: a legacy single-nursery payload under `ps-nursery-workbook` is
  lifted into `nurseries[<generated id>]` and becomes active. Idempotent.
- Home gains a nursery picker.

---

## Phase 2 — `pwa/grid.js`

### Task 2.1: render and inline edit

- `grid({ id, columns, rows, rowKey, opts })`. Column types `text | number |
  date | select | multiline`. Edits persist through `store.gridState(id).edits`
  keyed by `rowKey(row)`, never row index.

### Task 2.2: sort and filter

- Header click cycles asc → desc → none. Header filter dropdown lists distinct
  values as checkboxes plus a search box. Both persist per grid.

### Task 2.3: colour, styles, add/remove rows and columns

- Toolbar: add row, add column, delete row, delete column, fill colour, bold,
  italic. User columns stored in `extraColumns`, user rows in `extraRows`.

### Task 2.4: computed-cell overlay

- Cells whose computed value is overridden get a marker and a context-menu
  "revert to computed". Overlay entries whose `rowKey` no longer resolves after
  a re-import are kept and reported as orphaned, never dropped.
- Tests (`tests/test_grid.py`, via node): overlay survives re-sort, re-filter and
  re-import; orphan detection; padding of `extraColumns` on export.

---

## Phase 3 — export pipeline

### Task 3.1: `POST /export/xlsx`

- Modify `backend/app.py`. Payload `{filename, sheets:[{name, bands, headers,
  rows, colours, widths}]}` → streamed `.xlsx` via openpyxl. `bands` renders the
  merged band header row above `headers`.
- Reject empty `sheets` and sheets with no rows with a 400 and a clear message.
- Sheet names sanitised to Excel's 31-char / reserved-character rules.
- Tests (`tests/test_export.py`): band merge geometry, colour fill round-trip,
  sanitised sheet names, empty payload rejected.

### Task 3.2: `pwa/exporter.js`

- `exportGrid(id)`, `exportWorkbook()`, `exportIndividually()`,
  `exportBartender(tabName)`. Bartender sorts Source ID Z→A.
- Home gains "Export whole book" and "Export each tab".

---

## Phase 4 — tab changes

Each is its own commit.

- **4.1 Material Map** — field checkboxes above the grid, persisted; export.
- **4.2 Field Map** — calendar picker, row checkbox list, qty in grams,
  colour-coded map showing spike and run, editable, downloadable.
- **4.3 Nursery data** — date cells with wheel picker defaulting to today,
  Comments column, Capture GPS, DOP rows auto-filled from Field Map (appending
  `4th`, `5th`… when there are more than three planting dates), the four
  defaults.
- **4.4 Packet Prep** — radix-padded rack order (3 digits, 4 when a spike
  exceeds 999), Hybrid Code column, Nursery name column, Bartender export.
  Tests: padding width boundary at 999/1000, Z→A export order.
- **4.5 Nursery list** — grid with filters.
- **4.6 Replacements** — columns reduce to Stage · Plot · QR code (original
  entry) · QR code (replaced entry) · Reason · Status. Export expands each QR
  into the nine-column block under `Original entry ‖ Replaced entry ‖ Reason`.
  Tests: QR expansion, unmatched QR left blank and counted.
- **4.7 Planting errors** — new view, same as 4.6 without the replaced-entry QR;
  export is Original entry + Reason.
- **4.8 Fieldbook** — filters, Hybrid Code and Inbred Code columns, Print button
  at top, Range/Row search (blank = all), Generate AB selection map, Capture
  snapshot.
- **4.9 Operations and Comments** — Paperbags & labels row above Seedling, Post
  Harvest removed, Stage heading blanked, multiline cells, editable columns and
  text styles.

---

## Phase 5 — Date recording

### Task 5.1: month assignment

- `pwa/nursery-algos.js`: `assignMonth(day, startDay, startMonth, startYear)`
  returning `{month, year}` or `null` for blank / non-numeric / zero.
  Rule: `day >= startDay` → start month/year, else next month rolling past
  December into January of the following year.
- Mirror in `excel_workflow/nursery_algos.py`; pin with `test_js_parity.py`.

### Task 5.2: grid and the two serpentine modes

- Source: Fieldbook rows with `CMS reaction == "A"`. Columns per the spec.
  Buttons reuse `serpentineTwoRowBands()` and `serpentineByRange()`.

### Task 5.3: pull out bags

- Copy each A-line's S 1..S 4 onto its paired B-line; 1–31 checkbox grid for the
  date range; list matching B-lines in horizontal serpentine order.

### Task 5.4: trend graph

- Counts per calendar date from the start date to the latest assigned date →
  Date / S1..S4 Count. Inline SVG line chart with markers, bottom legend, axis
  titles "Date" and "Number of Records", value axis minimum 0. Exportable.
- Test counts against the `Trend Analysis` sheet of `Tab information.xlsx`.
- Validation matching the VBA: year 1900–2100, month 1–12, day 1–31.

---

## Phase 6 — package

- Full test run, manual pass through every tab in the browser, rebuild the
  Windows installer, re-run `scripts/make_release.py`.

---

## Self-review

**Spec coverage:** every docx bullet maps to a task — General → 1.2/1.3/3.2,
Material Map → 4.1, Field Map → 4.2, Nursery data → 4.3, Packet Prep → 4.4,
Nursery list → 4.5, Replacements → 4.6, Planting errors → 4.7, Fieldbook → 4.8,
Date recording → 5.1–5.4, Operations/Comments → 4.9. "Every table editable"
→ Phase 2, applied per tab in Phase 4.

**Naming consistency:** `gridState(tabName)` is the single accessor used by
`grid.js` and `exporter.js`; `rowKey` is the overlay key everywhere; grid ids are
tab names throughout.
