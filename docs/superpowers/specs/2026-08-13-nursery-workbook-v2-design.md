# Nursery Workbook v2 — Design

**Date:** 2026-08-13
**Status:** Approved for planning
**Scope:** Restructure the Sorghum nursery workbook to the client's revised
specification, then bring the PWA to parity against the same spec.

## Source documents

| Document | Contributes |
|---|---|
| `Workbook changes.docx` | Authoritative tab list, per-tab requirements |
| `Workbook flow.pptx` (SmartArt, `ppt/diagrams/data1.xml`) | 13-step build order |
| `File for Toyin.xlsx` | Exact `Nursery data` and `Date recording` layouts |
| `Fieldbook colouring for repeating values VBA.docx` | `frmSelectColumns` + colouring module |
| `Fieldbook template VBA code.docx` (repo root) | Serpentine sort logic |
| `Sorghum nursery prep workflow.docx` (repo root) | Current vs proposed workflow; "Additionals" grouping |

Where `Workbook changes.docx` and `Sorghum nursery prep workflow.docx`
disagree, the former wins — it is the later and more specific document. One
such conflict is recorded under Open Questions.

## Goals

1. The generated workbook contains exactly the 12 specified entries below, in
   order, plus `Home` and a hidden `Settings`. Entry 5 expands to one tab per
   planting date, so the physical sheet count is `13 + N + (1 if AB)`.
2. The five currently-generated extra tabs survive as nursery-type conditional
   extras rather than being deleted.
3. Excel and PWA implementations are behaviourally identical, verifiable by
   running one PRISM export through both.

## Non-goals

- The Date-recording "add the month and generate graph" and "pull out bags with
  ease" buttons. The client marked these "need in-person explanation."
- Spray-track rows inside Field Map. Deferred until the client supplies the
  data, but the Field Map layout must tolerate their later insertion.
- Any change to the QR payload format, so already-printed labels stay valid.

## Approach: one declarative spec, three consumers

Tab names, ordering, column sets and print settings currently live in
`build_workbooks.py`, in `SHEET_*` constants in `NurseryTemplate.bas`, and
again in `pwa/app.js`. Three copies drift, and drift would make the eventual
"which implementation is better?" comparison meaningless.

A single `excel_workflow/spec/nursery_spec.json` becomes the source of truth.
It is consumed by:

1. `build_workbooks.py` — openpyxl sheet creation, headers, widths, print setup.
2. `excel_workflow/gen_spec_constants.py` → generated `vba/SpecConstants.bas`,
   holding sheet-name constants and column indices.
3. The PWA, importing the same JSON.

Algorithms (spike pattern, serpentine, split assignment, rack order) remain
hand-written per platform. They are pinned by shared golden fixtures rather
than shared code, because the runtimes are too different to share logic and
the fixtures catch divergence just as well.

`SpecConstants.bas` is generated and must never be hand-edited; the generator
is the only writer.

## Tab structure

Default workbook, in this exact order:

1. `Nursery site`
2. `Material Map`
3. `Field Map`
4. `Nursery data`
5. `Packet Prep 1..N`
6. `Nursery list`
7. `Replacements and Errors`
8. `Updated nursery site`
9. `Fieldbook`
10. `Date recording` — AB nurseries only, inserted before `Operations`
11. `Operations`
12. `Comments`

Plus `Home` (always) and `Settings` (hidden).

### Nursery-type extras

Off by default; unlocked by the nursery-type selection at init. The mapping
below follows the "Additionals (TFMSA spray, AB date recording and AB bag
pulling)" grouping in `Sorghum nursery prep workflow.docx`:

| Tab | Nursery type |
|---|---|
| `BC0 labels` | AB |
| `BC0 TFMSA record` | AB |
| `Pulling bags` | AB |
| `TFMSA Spray plots` | Other / trait |
| `Hy Heights` | Hybrid |

Their existing builders are retained unchanged and simply not invoked unless
the matching type is selected.

## Per-tab requirements

### Init wizard (`btnInitNursery`)

Prompts for a filename, then a nursery-type multi-select: Selection, AB,
Hybrid, Other. Multiple types may be chosen. Selecting AB inserts
`Date recording`. Type selection also determines which extras are built.

### Home

One navigation button per present tab. `Workbook_Open` activates Home
unconditionally.

### Nursery site

PRISM downloads arrive as `Sheet 1`. Rename to `Nursery site`; no other
transformation.

### Material Map

Generated from `Nursery site` data. Grid keyed by Range × Row, with Range
numbers down both sides and Row numbers along both top and bottom. Populates
Material ID, Inbred code and Hybrid code. Uses the existing `BuildMap`
routine, extended from one value to three.

### Field Map

Same grid geometry as Material Map, without material information.

A button collects, per planting date: the date, its row list, and seed
quantity per plot. From that:

- **Spike numbers** follow a global period-4 cycle `1, 2, 2, 1` across the row
  sequence. Row 1→spike 1, row 2→spike 2, row 3→spike 2, row 4→spike 1, then
  repeat. This models a two-cone planter whose cones swap sides when it turns
  around.
- **Run direction** flips every two rows, starting forward. Forward and reverse
  runs are filled with distinct colours.

The button is re-runnable: it rewrites its own output block in place rather
than appending, so revised planting dates do not accumulate stale output. The
layout must tolerate spray-track rows being inserted later without the grid
breaking.

### Nursery data

Printable front page. Reproduce the attached layout exactly:

- `A1:D1` merged — "R&D Fieldbook - Grain Sorghum", bold 20pt, centered.
- `B3:C3` and `B4:C4` merged, **left blank** (filled later by hand).
- Rows 5–23: labels in column B, 12pt, thin bottom borders. Column C **left
  blank**. Labels in order: GPS coordinates, Planter, Seeds/side, Plot length,
  Alley way spacing, 1st DOP and rows, 2nd DOP and rows, 3rd DOP and rows, Tag
  rows, Tagging date, 1st bagging date, Last bagging date, 1st crossing date,
  Last crossing date, Unused A-line bags removal date, Injection dates, Opening
  bags date, Closing bags date, Harvest date.
- `A24:D24` merged, blank.
- Column widths: A 10.71, B 58.29, D 9.14, F 29.14, G 59.57, H 9.14. The empty
  F/G block is preserved.

### Packet Prep

One tab per planting date: `Packet Prep 1`, `Packet Prep 2`, … The existing
25-column format is retained, plus a new `Split no.` column whose value is the
tab index — all `1` in `Packet Prep 1`, all `2` in `Packet Prep 2`.

### Nursery list

All unique materials from `Nursery site`, Source ID ascending. Existing
structure retained (Repeats, Qty Required = 1.4 × Repeats, qty per packet
prompt).

### Replacements and Errors

Merges the current `Replacements done` and `Planting error noted` tabs.

- Drop the "Rename this tab…" row.
- New leading column `Stage`, validation list: Packeting, Planting.
- New `Split no.` column immediately left of `Status`.
- QR column widened to fit 150 characters.
- `Status` defaults to "Waiting decision", validation list adds "Changes made
  in PRISM".
- A button reads the QR value, parses the plot from it, looks that plot up
  across the Packet Prep tabs, and writes the split into `Split no.`.

Scanning assumes a keyboard-wedge barcode reader typing the payload into the
QR cell. The button parses whatever is already in the cell; it does not drive
the scanner.

### Updated nursery site

Created empty. The user re-downloads from PRISM and pastes here. Hosts the
button that builds the Fieldbook.

### Fieldbook

Built from `Updated nursery site`, only after that tab has data.

Columns, in order and no others: `Range`, `Row`, `Material ID`, `Source ID`,
`Gen`, `CMS`, `Plot`, `Comments`. This **drops** the current `Crossed bags` and
`Bagging Info` columns, and **renames** `R_R` to `Plot`, moving it from column
3 to column 7. `Plot` keeps its existing derivation, `Range & "_" & Row`.

- All borders, thin. Header row bold, centered, frozen.
- Print: landscape, A4, duplex flipped on the short edge, `PrintTitleRows =
  "$1:$1"`, `CenterFooter = "&P/&N"`, `RightHeader` = filename.
- When the nursery is AB, rows with `CMS = B` are filled green.
- Re-runnable.

**Serpentine sort**, ported from `Fieldbook template VBA code.docx`: sort by
Range ascending, then within each Range sort Row descending when the Range is
even and ascending when odd.

### Fieldbook colouring

The client's module is used essentially verbatim: `StartColouring`,
`ApplyColouring_VisibleOnly`, `GetMutedColour`. It colours only visible
(filtered) rows, groups by a chosen column, and applies an eight-colour muted
pastel cycle.

`frmSelectColumns` is authored as a checked-in `.frm` text file with three
design-time controls — `lblGroupBy`, `cmbGroupBy` (fmStyleDropDownList),
`cmdApply` — with the per-column checkboxes created at runtime in
`UserForm_Initialize`, as the client's code does.

Build path: `install_excel_macros.py` imports the `.frm` via the VBA object
model on macOS Excel (confirmed present: Excel 16 + xlwings 0.35.3), then
`extract_seed.py` re-extracts `seeds/Nursery_Template.vbaProject.bin`.

**Fallback:** if `.frm` import proves unreliable on Mac Excel, replace the form
with a worksheet control panel — a grouping dropdown plus tick cells on a
helper sheet — driving the same `ApplyColouring_VisibleOnly` routine. The
colouring logic is unchanged either way; only the picker differs.

### Date recording (AB only)

Columns from the sample: `Group`, `Range`, `Row`, `O/E`, `S 1`, `S 2`,
`Material ID`, `Source ID`, `Gen`, `CMS`, `In. Code`. Widths 6.29, 6.86, 5.43,
default, 5.14, 5.57, 21.43, 28.57, 5.14, 4.71, 8.43.

`Group` and `O/E` are **kept but hidden**. The client said they "can be removed
if possible", but decoding the sample shows they are load-bearing: `Group` is
the two-row band index (`group = row \ 2` — Group 1 covers rows 2–3, Group 13
covers rows 26–27) and `O/E` carries the odd/even parity that drives the
serpentine direction.

`S 1` and `S 2` hold day-of-month numbers only; the month is applied
separately. They are left empty for field entry.

Two buttons:

- **Record by range and pull out bags** (was "Serpentine by Range") — Range
  ascending; within each Range, Row descending when Range is even, ascending
  when odd.
- **Record 2-rows together** (was "Serpentine by 2-rows") — the by-Row variant
  banded in pairs: walk each two-row group, serpentining by Range across the
  band. This reproduces the sample's `(2,3),(2,2),(3,2),(3,3),(4,3),(4,2)…`
  ordering within Group 1.

The second button's logic is **inferred from the sample data**, since the
client's original macro-enabled workbook was not supplied. Flag for
confirmation.

### Operations

Existing structure retained. Title "Operations Overview" centered, the empty
second row removed, `Plan` / `Reminders` / `Comments` renamed to `Group 1` /
`Group 2` / `Group 3`, all borders applied.

### Comments

Identical to Operations, titled "Field Comments".

## Data flow

```
PRISM export
  └─> Nursery site
        ├─> Material Map        (Material ID, Inbred, Hybrid)
        ├─> Field Map           ─ wizard: DOPs, rows, qty → spikes, runs
        │     └─> Packet Prep 1..N   (Split no. = tab index)
        ├─> Nursery list        (unique Source IDs, ascending)
        └─> Nursery data        (layout only, values entered by hand)

Replacements and Errors ──QR plot lookup──> Packet Prep N → Split no.

PRISM re-export
  └─> Updated nursery site ──button──> Fieldbook ──> Date recording (AB)
```

## Error handling

- Every builder validates its source tab is populated before running and exits
  with a specific message naming the tab and the expected start row.
- The Fieldbook button refuses to run while `Updated nursery site` is empty.
- The Field Map wizard validates that row lists are numeric, non-overlapping
  between planting dates, and within the field's row range.
- The QR lookup reports plots it cannot resolve rather than silently leaving
  `Split no.` blank.
- Re-runnable builders clear their own output region first, so a partial
  failure cannot leave a half-old, half-new tab.

## Testing

- **Golden fixtures.** One committed PRISM export plus expected Fieldbook rows,
  Packet Prep splits, spike assignments and serpentine orderings. Both
  implementations must reproduce them.
- **Pure-function unit tests** in Python for spike pattern, serpentine
  ordering, split assignment and group/parity derivation. These are the
  algorithms most likely to drift.
- **Structural assertion** that the generated workbook's sheet names and order
  match `nursery_spec.json` exactly.
- **Manual verification** for print setup, colour fills and the UserForm, since
  these are not reachable from openpyxl.

## Sequencing

**Phase 1 — Excel.** Spec layer and generator; tab restructure; init wizard and
Home; Material Map and Field Map wizard; Packet Prep fan-out; Nursery data
layout; Replacements and Errors; Fieldbook and serpentine; colouring UserForm
and seed rebuild; Date recording; Operations and Comments.

**Phase 2 — PWA.** Same spec JSON, same fixtures, parity verified against
Phase 1 output.

## Open questions

1. **Material Map provenance.** `Workbook changes.docx` says generate it from
   `Nursery site`; `Sorghum nursery prep workflow.docx` says export it from
   PRISM alongside Nursery site. This design generates it. Confirm.
2. **Extras mapping.** The tab→nursery-type table above is inferred. Confirm,
   particularly `TFMSA Spray plots`.
3. **"Right column empty in rows 5-23"** is read as column C. The sample file
   has values there, but the instruction says leave them blank, so the sample
   is read as a filled-in example rather than the template.
4. **"Record 2-rows together"** logic is inferred from sample data. Confirm
   against the client's original workbook if it can be obtained.
5. **Nursery list extras** — `Sorghum nursery prep workflow.docx` asks for
   duplicate-check buttons on Inbred/Hybrid code and a "filter qty above X for
   bulk treatment" button. Absent from `Workbook changes.docx`. In or out?
