# Nursery app v3 — design

**Date:** 2026-08-20
**Branch:** nursery-workbook-v2
**Scope:** desktop app only (`poc/backend/`, `poc/pwa/`). The Excel/VBA workbook is
not modified.

## Sources

Two client documents drive this work, both at `poc/`:

- `Toyin- Updated app changes.docx` — ~60 changes across 12 tabs, plus an
  appendix containing `GenerateS1S2TrendAnalysis`, a VBA routine that is the
  reference implementation for the Date recording trend graph.
- `Tab information.xlsx` — four sheets that are *target formats*, not
  instructions:
  - `Replacements` — export layout: **Original entry** (9 columns) ‖
    **Replaced entry** (9 columns) ‖ **Reason**.
  - `Planting errors` — the Original entry block plus **Reason**.
  - `Date recording` — 396 sample rows showing the target columns, including
    `S1 Month` and `S2 Month`.
  - `Trend Analysis` — expected output of the appendix VBA: Date / S1 Count /
    S2 Count, with a line chart.

The nine-column entry block is, in order: QR Key, Source ID, Material ID,
Inbred Code, Experimental Hybrid Code, Pedigree, Generation, CMS reaction,
W/column.

A further requirement was given in conversation and supersedes the narrower
per-tab wording in the docx:

> every table in the app should be editable, users can sort, colour, add
> columns and rows etc

## Decisions taken

| Question | Decision |
|---|---|
| Does this touch the Excel/VBA workbook? | No. Desktop app only. VBA references in the docx are read as pointers to logic worth porting, not files to edit. |
| Multiple nurseries | Named nurseries, switchable. Create / switch / rename / duplicate / delete, each with its own data. |
| "Recording will happen 3-4 times" | Fixed **S 1 – S 4** selection columns, each with a derived Month column. |
| Hand-edits to computed tables | Edits win and are flagged. Stored as an overlay keyed to stable row identity, re-applied after every rebuild, marked in the UI, with revert-to-computed. Never lost silently. |

## Architecture

### Current shape

`pwa/app.js` holds a `VIEWS` registry — `VIEWS["Tab name"] = (main, tabName) => {}`
— dispatched by `render()`. The visible tab list comes from `tabsFor()` in the
generated `pwa/nursery-spec.js`, which mirrors `tabs_for()` in
`excel_workflow/spec/loader.py`; `tests/test_js_parity.py` pins the two to
identical output. All state is a single `BLANK` object in `localStorage`.

Tables are rendered by a stateless `table()` helper. Tabs divide into:

- **Computed** — Material Map, Field Map grid, Packet Prep, Nursery list,
  Fieldbook, Date recording. Rebuilt from `state.prism` / `state.updatedPrism`
  on every render.
- **Owned** — Replacements, Operations, Comments. Store their own rows.

### Three new modules

Nearly every requirement is one of four verbs — edit, filter, colour, export.
Three modules provide them once rather than twelve times.

#### `pwa/grid.js` — editable grid

Replaces the role of the `table()` helper. Signature:

```js
grid({ id, columns, rows, rowKey, opts }) -> HTMLElement
```

- `id` — stable persistence key, the tab name (e.g. `"Packet Prep 1"`).
- `columns` — `[{ key, label, type, options, width, readOnly, computed }]`
  where `type` is `text | number | date | select | multiline`.
- `rows` — array of row objects.
- `rowKey(row)` — returns stable row identity. For computed tables this is the
  plot (`"12_3"`), never the row index, so overlays survive re-sort and
  re-import.

Features, all available on every tab:

- Click a header to sort ascending / descending / none.
- Header filter dropdown: distinct values as checkboxes plus a search box.
- Column resize by drag; hide/show columns.
- Inline cell editing typed per column.
- Toolbar: add row, add column, delete row/column, fill colour, bold, italic.
- Export button (see exporter).
- Computed cells that carry a user edit are marked, with a context-menu
  "revert to computed".

#### `pwa/exporter.js` + `POST /export/xlsx`

The grid serialises to JSON; the backend writes a genuine `.xlsx` using
openpyxl (already a dependency via `requirements.txt`). Payload:

```json
{ "filename": "...",
  "sheets": [ { "name": "...", "bands": [["Original entry", 9], ["Reason", 1]],
                "headers": ["..."], "rows": [["..."]],
                "colours": {"r,c": "FFFF00"}, "widths": [12.0] } ] }
```

`bands` produces the merged header row seen in `Tab information.xlsx`.

Three export modes:

- **Per tab** — one sheet.
- **Whole book** — every visible tab as its own sheet, one file. Also offered
  as individual files.
- **Bartender** — Packet Prep only; flat sheet, Source ID sorted **Z→A**.

#### `pwa/store.js` — multi-nursery state

```js
state = { activeNursery: "<id>",
          nurseries: { "<id>": { code, fileName, types, plantingDates, seedQty,
                                 prism, updatedPrism, fieldMap, nurseryData,
                                 replacements, plantingErrors, dateRecording,
                                 operations, comments, grids } } }
```

`grids[tabName] = { sort, filters, colours, styles, widths, hidden,
extraColumns, extraRows, edits }`, where `edits` is
`{ [rowKey]: { [columnKey]: value } }`.

A one-time migration lifts an existing single-nursery `localStorage` payload
into `nurseries` under a generated id and sets `activeNursery`, so upgrading
loses nothing.

Home gains a nursery picker.

### Spec changes without touching the workbook

Three requirements change tab *structure*, and `nursery_spec.json` is shared
with the Excel builder. Editing it directly would restructure the workbook —
out of scope — and break `tests/test_workbook_structure.py`.

The JSON therefore gains an **`app_overrides`** block that only `render_js()`
in `excel_workflow/gen_spec_constants.py` merges when generating
`pwa/nursery-spec.js`. `render_constants()`, `loader.py` and the openpyxl
builder keep reading the base keys unchanged. The divergence is explicit and
collapses to nothing when the workbook later catches up.

Overrides:

- `default_tabs` — replace `"Replacements and Errors"` with `"Replacements"`
  and `"Planting errors"`.
- `conditional` — `"Date recording": {"types": ["AB"], "after": "Fieldbook"}`.
- `date_recording_columns` — add `S 3`, `S 4` and `S1..S4 Month`.
- `nursery_data_defaults` — the Nursery data default values.

`tabsFor()` and `tabs_for()` currently understand only `rule.before`. Both gain
`after` support in the same change so `tests/test_js_parity.py` stays green.
Placement rule: a conditional with `before: X` is emitted immediately before
`X`; one with `after: X` immediately after `X`. A rule must specify exactly one
of the two.

## Tab-by-tab requirements

Every bullet from the docx, with its implementation.

### General

| Requirement | Implementation |
|---|---|
| Initialize nursery: ask for nursery code and type | Already implemented in `initNursery()`. |
| File name? | Already prompted. Retained. |
| Export whole book (in individual tabs) and individually | Exporter: whole-book multi-sheet file, plus per-tab files. |
| How to create for multiple nurseries | `store.js` + Home nursery picker. |
| If nursery type is AB, add AB date recording | `app_overrides.conditional`, anchored after Fieldbook. |
| In-built filter for all columns like Excel | `grid.js` header filters, every tab. |

### Nursery site

No change.

### Material Map

- Checkbox row above the grid choosing which fields appear in each cell.
  Replaces the hardcoded `["Material ID", "Inbred Code", "Hybrid Code"]`.
  Selection persists per nursery.
- Downloadable via exporter.

### Field Map

- Planting date entered with a **calendar picker**, replacing the free-text
  prompt.
- Seed quantity labelled **grams**.
- Rows for each planting date chosen from a **checkbox list** of available
  rows, replacing comma-separated text entry. The existing `assignSplits()`
  guard against a row appearing in two planting dates is retained.
- The result renders as a **colour-coded map**, one colour per planting date,
  each cell showing spike number and run direction — replacing the current
  flat table.
- Grid editable and downloadable.

### Nursery data

- Any label containing "date" gets a calendar cell, defaulting to the current
  date when first opened.
- New **Comments** column to the right.
- **Capture GPS** button using browser geolocation, writing into the GPS
  coordinates row; the value stays editable.
- `1st/2nd/3rd DOP and rows` auto-filled from Field Map. The base spec provides
  three such labels; if Field Map holds more than three planting dates, further
  `4th DOP and rows`, `5th …` rows are appended so no planting date is dropped.
- Defaults, all editable: Planter = `Almaco Precision Planter`,
  Seeds/side = `34`, Plot length = `4.5m`, Alley way spacing = `0.75m`.
- Date cells use a wheel-style day/month/year picker.

### Packet Prep 1..n

Applies to every fanned-out Packet Prep tab.

- **Radix-sorted rack order**: zero-padded to 3 digits, widening to 4 when any
  spike holds more than 999 entries. Padding width is computed per tab from the
  largest spike count, so all values in a tab share one width.
- Add the missing **Hybrid Code** column.
- Add a **Nursery name** column carrying the nursery code.
- **Export to Excel** for Bartender.
- Export always sorts **Source ID Z→A**. On-screen order is unaffected.

### Nursery list

- Column filters.

### Replacements

Columns reduce to: **Stage · Plot · QR code (original entry) · QR code
(replaced entry) · Reason · Status**.

Dropped from the tab: Timestamp, Material ID, Source ID, Technician, Split no.

- Columns auto-fit their contents.
- All columns editable.
- Existing "Read QR" action populates entry fields by matching the scanned
  plot against nursery data.
- **Export** expands each QR value into the full nine-column block, producing
  the `Tab information.xlsx` layout: Original entry ‖ Replaced entry ‖ Reason,
  with a merged band header. This resolves the apparent conflict between "remove
  material ID and source ID" and an export format that contains them: the tab is
  short, the export is expanded by lookup.

`W/column` does not exist in the app or in the current PRISM import. It ships as
a blank editable column until a PRISM source for it is identified.

### Planting errors (new tab)

Identical to Replacements without the replaced-entry QR column. Export is the
Original entry block plus Reason.

### Updated nursery site

No change.

### Fieldbook

- Per-column filters.
- Add **Hybrid Code** and **Inbred Code** columns.
- **Print button moves to the top** of the tab.
- **Search** boxes for Range and Row separately, filtering to matching rows. An
  empty box means all ranges / all rows.
- **Generate AB selection map** from the supplied date range, copying A-line
  selections onto their paired B-lines.
- **Capture** button — captures a timestamped snapshot of the current filtered
  fieldbook, listed and re-openable. See open questions.
- Date recording is anchored immediately after this tab for AB nurseries.

### Date recording (new build)

Source: **Fieldbook rows, A-lines only** (`CMS reaction == "A"`). This replaces
the current `isRecurrent()` BC*/Fn filter, which selects a different set.

Columns: Group · Range · Row · O/E · S 1 · S1 Month · S 2 · S2 Month · S 3 ·
S3 Month · S 4 · S4 Month · Material ID · Source ID · Gen · CMS · In. Code.

Four buttons:

1. **Serpentine – Vertical** — two rows together, snaking by range. Uses the
   existing `serpentineTwoRowBands()`.
2. **Serpentine – Horizontal** — bay by bay. Uses the existing
   `serpentineByRange()`.
3. **Pull out bags** — copies each A-line's S 1..S 4 values onto its paired
   B-line, prompts for a date range as a **1–31 checkbox grid**, and lists the
   B-lines whose selection values match, in horizontal serpentine order.
4. **Generate graph** — port of `GenerateS1S2TrendAnalysis`.

#### Month assignment rule

Ported verbatim from the appendix VBA. Given a start year, month and day, for
each selection day value `d`:

- if `d >= startDay` → month = startMonth, year = startYear
- else → month = startMonth + 1, rolling to January of the next year past 12

The full date is `DateSerial(assignedYear, assignedMonth, d)`. Blank,
non-numeric and zero values produce an empty Month cell and are excluded from
counts.

#### Trend analysis

Counts records per calendar date from the earliest start date to the latest
assigned date, producing Date / S1 Count / S2 Count / S3 Count / S4 Count, and
renders a line chart with markers, legend at the bottom, axis titles "Date" and
"Number of Records", and a value axis minimum of 0 — matching the VBA. Table and
chart are exportable, reproducing the `Trend Analysis` sheet.

Extension beyond the VBA: the appendix handles S1 and S2 only; S3 and S4 follow
the identical rule.

### Operations and Comments

Both tabs, identically:

- New **"Paperbags & labels"** row above Seedling.
- **"Post Harvest" row removed.**
- The **"Stage" column heading is blanked**; the coloured stage cells remain.
- Cells accept **multiple lines**.
- Columns and text styles editable.

Resulting rows: Paperbags & labels, Seedling, Vegetative, Heading, Flowering,
Grain filling, Harvest.

## Error handling

- A row assigned to two planting dates is refused by `assignSplits()`, as
  today.
- QR values matching no plot are reported by count and left blank rather than
  guessed.
- Export requests with no rows return a clear message rather than an empty
  file.
- Geolocation denial or unavailability leaves the GPS cell editable by hand
  with an explanatory message.
- Grid overlay edits whose `rowKey` no longer exists after a re-import are
  retained but listed as orphaned, so re-importing corrected PRISM data cannot
  silently discard work.
- Trend analysis rejects years outside 1900–2100 and out-of-range months and
  days, matching the VBA's validation.

## Testing

The four existing suites (72 tests) must stay green.

New and extended coverage:

- `test_js_parity.py` — extended for the `after` placement rule and the
  `app_overrides` merge, keeping `tabsFor()` and `tabs_for()` identical.
- `test_spec.py` — `app_overrides` does not alter the base spec consumed by the
  workbook builder; a conditional rule may not set both `before` and `after`.
- Rack order radix padding, including the 999→4-digit boundary.
- Export Source ID Z→A ordering.
- Month assignment, including the day-below-start rollover and the December→
  January year roll.
- Trend counts against the `Trend Analysis` sheet in `Tab information.xlsx`.
- QR → nine-column expansion for both Replacements and Planting errors.
- Overlay edits surviving re-sort, re-filter and PRISM re-import.
- Multi-nursery migration from the single-nursery payload.

## Sequencing

Six phases, each independently shippable:

1. Spec `app_overrides` + `after` support + multi-nursery store and migration.
2. `grid.js` — editing, sort, filter, colour, add/remove rows and columns.
3. `exporter.js` + `POST /export/xlsx` — per-tab, whole-book, Bartender.
4. Tab changes: Material Map, Field Map, Nursery data, Packet Prep, Nursery
   list, Replacements, Planting errors, Fieldbook, Operations, Comments.
5. Date recording: four buttons, month assignment, trend chart.
6. Rebuild and re-verify the Windows installer.

## Open questions

Assumptions are recorded; correction is cheap at any point.

1. **Fieldbook "Add capture button"** — read as capturing a timestamped
   snapshot of the current filtered fieldbook. Could instead mean photo capture
   or capturing observations into a column.
2. **`W/column`** — absent from the app and from the current PRISM import.
   Ships as a blank editable column pending a source.
3. **"scrollable dates (like in iphone)"** — read as a wheel-style
   day/month/year picker on Nursery data date cells.
4. **"Read QR – Split no?"** on Replacements — read as a question rather than a
   requirement. Split no. is removed from the tab per the explicit removal
   instruction; the QR read action is kept.
5. **"from Row 2 to penultimate row"** on Date recording — VBA phrasing for
   "all data rows". The app's Fieldbook has no footer row, so all data rows are
   taken.
