# Pacific Seeds — Nursery Workflow (Excel + VBA)

This folder builds two workbooks that implement the proposed Sorghum nursery
workflow end-to-end with **clickable buttons, automated dashboards and
cross-file sync** between every nursery.

```
excel_workflow/
├── install.command                 # double-click — first-time bootstrap into Excel
├── install_excel_macros.py         # legacy xlwings installer (Windows-only effective)
├── build_workbooks.py              # generates .xlsx skeletons + bakes VBA from seeds
├── bake_vba.py                     # injects a vbaProject.bin into an .xlsx → .xlsm
├── extract_seed.py                 # pulls vbaProject.bin out of a working .xlsm
├── vba/                            # canonical VBA source-of-truth (text)
│   ├── NurseryTemplate.bas         # per-nursery workflow logic (13 steps)
│   ├── NurseryHub.bas              # hub aggregation + cross-file sync
│   ├── ThisWorkbook_Template.cls   # auto-refresh on open, auto-push on save
│   ├── Sheet_Home_Template.cls     # double-click router for the workflow buttons
│   ├── Sheet_Home_Hub.cls          # double-click router for the Hub Home
│   └── bootstrap.bas               # paste-once self-installer (fallback only)
├── seeds/                          # vbaProject.bin captured from working .xlsm
│   ├── Nursery_Hub.vbaProject.bin       # ← present, baked into builds
│   └── Nursery_Template.vbaProject.bin  # ← generate once via extract_seed.py
└── output/
    ├── Nursery_Hub.xlsm            # baked automatically if seed exists
    └── Nursery_Template.xlsm       # baked automatically if seed exists
```

## How `build_workbooks.py` now works

```
1. openpyxl builds Nursery_Hub.xlsx + Nursery_Template.xlsx (sheets, styling, formulas)
2. For each, if seeds/<stem>.vbaProject.bin exists:
     → bake_vba copies the .xlsx, injects the VBA, saves as .xlsm
   Otherwise:
     → leaves the .xlsx alone; you'll see a hint to run extract_seed.py
```

This means **once you have a working .xlsm, you never have to paste VBA again**.
Future `python build_workbooks.py` runs spit out fresh, ready-to-use .xlsm files.

## How the buttons work

There is no Form Control / ActiveX button (those can't be created reliably from
Python on macOS). Instead each "button" is a **styled cell**. The VBA hooks the
`Worksheet_BeforeDoubleClick` event on the Home sheet, looks up a matching
macro name in a hidden helper column, and runs it.

This means once VBA is installed, **every button is double-click-to-run, no
right-click "Assign macro" dance needed**.

---

## Recommended workflow — bake from seeds

If both seeds already exist (`seeds/Nursery_Hub.vbaProject.bin` and
`seeds/Nursery_Template.vbaProject.bin`), you're done forever:

```bash
cd poc
.venv/bin/python excel_workflow/build_workbooks.py
```

Both `.xlsm` files appear in `excel_workflow/output/` with all VBA already
working. Open and use.

### First-time seed creation (once per workbook)

The Hub seed is already in this repo. To produce the Template seed:

1. Do the one-time manual VBA install on Nursery_Template (instructions below).
2. Confirm it works (open the .xlsm, double-click Step 1 — should ask for code).
3. Lock in the seed:

   ```bash
   cd poc
   .venv/bin/python excel_workflow/extract_seed.py \
       excel_workflow/output/Nursery_Template.xlsm
   ```

4. From now on, `python build_workbooks.py` rebuilds the Template fully baked.

If you ever edit the canonical `.bas` / `.cls` files in `vba/`, you'll need to
re-do the one-time install AND re-extract the seed, since the seed snapshots
the compiled VBA, not the source text.

---

## The fastest install — `install.command` + 4 keystrokes

> **Mac reality check:** Microsoft Excel for Mac does not allow external tools
> (Python, AppleScript, xlwings) to add VBA modules. This is a hard
> architectural limit, not a setting we can flip. The workaround is a tiny
> self-installing macro you paste **once per workbook**. From there it does
> everything automatically and deletes itself.

**One-time Excel security toggle:**

1. Open Excel → **Excel ▸ Preferences ▸ Security & Privacy ▸ Macro Security**.
2. Tick **"Trust access to the VBA project object model"**.

**Then:**

3. In Finder, navigate to `poc/excel_workflow/`.
4. Double-click **`install.command`**.
   - It regenerates the .xlsx skeletons.
   - It copies the bootstrap macro (`vba/bootstrap.bas`) to your **clipboard**.
   - It opens both `Nursery_Template.xlsx` and `Nursery_Hub.xlsx` in Excel.
5. For each of the two workbooks (Template first, then Hub):
   - **⌥F11** (or Tools ▸ Macros ▸ Visual Basic Editor)
   - **Insert ▸ Module**
   - **⌘V**  ← the bootstrap is already on your clipboard
   - **F5**  ← runs `InstallAll`; pick the `vba/` folder when prompted
   - **⌘⇧S** ▸ choose **Excel Macro-Enabled Workbook (.xlsm)** ▸ Save.
6. The bootstrap detects whether it's in Template or Hub from the workbook
   name, installs the matching modules, then deletes itself.

That's it — roughly 30 seconds per workbook, one time, ever.

From now on, if you change anything in `vba/`, re-run `install.command` and
paste the bootstrap again — you get a fresh .xlsm with the latest code.

---

## Manual fallback (if install.command doesn't suit you)

Do this once per **template**. After it's saved as `.xlsm`, copy it freely for
each new nursery (or use the Hub's "New Nursery from Template" button).

1. **Regenerate the file** (if you've changed `build_workbooks.py`):
   ```bash
   cd poc
   .venv/bin/python excel_workflow/build_workbooks.py
   ```
2. **Open** `excel_workflow/output/Nursery_Template.xlsx` in Excel.
3. Open the VBA editor:
   - macOS: **Tools → Macros → Visual Basic Editor**
   - Windows: press **Alt + F11**
4. In the VBE, **File → Import File…** and import:
   - `vba/NurseryTemplate.bas`  (becomes a regular module)
5. Paste the class-module code (these can't be `Import`-ed because they belong
   to specific objects):
   - Double-click **ThisWorkbook** in the project tree, then paste the contents
     of `vba/ThisWorkbook_Template.cls` (drop the leading `VERSION 1.0 CLASS` /
     `Attribute …` lines — just the code from `Option Explicit` onwards).
   - Double-click the **Home** sheet, paste the contents of
     `vba/Sheet_Home_Template.cls` (same — keep the code, drop the headers).
6. Back in Excel, **File → Save As…** → choose
   **Excel Macro-Enabled Workbook (.xlsm)** and save (overwriting / next to the
   .xlsx is fine).
7. Close and re-open to confirm the dashboard auto-refreshes.

Workflow buttons are now double-click-to-run, in order:

**Phase 1 — Pre-field prep**
1. Initialise from PRISM export
2. Build Nursery list (auto from Nursery site, with Repeats & Qty Required)
3. Design Field Map
4. Generate Packet Prep & QR labels (the 13-step packet workflow)
5. Sort packets for racking (LSD radix)

**Phase 2 — Field operations**
6. Record replacement (Packeting / Planting)
7. Record planting error
8. Spray track + date recording
9. AB bag pulling

**Phase 3 — Post-field & sync**
10. Pull updated Nursery site from PRISM
11. Generate Fieldbook
12. Refresh dashboard
13. Push to Hub

---

## One-time setup — Nursery_Hub

1. Open `excel_workflow/output/Nursery_Hub.xlsx`.
2. **Tools → Macros → Visual Basic Editor**.
3. **File → Import File…** → `vba/NurseryHub.bas`.
4. Paste `vba/Sheet_Home_Hub.cls` into the Hub's **Home** sheet class module.
5. **File → Save As…** → `.xlsm`.

The Hub has four buttons on Home:
- **Refresh Dashboard** — reads `registry.csv` from the shared folder and
  rebuilds the Dashboard tab.
- **Register This Hub Folder** — sets the path where each nursery workbook
  will write its summary.
- **Open Nursery Folder** — opens that folder in Finder / Explorer.
- **New Nursery from Template** — clones the template into the folder under a
  new nursery code.

---

## How the workbooks "talk to each other"

```
   ┌──────────────────────────┐                     ┌──────────────────────────┐
   │  AUGT1-26S-IMI.xlsm      │  ──Push to Hub──▶   │                          │
   ├──────────────────────────┤                     │                          │
   │  AUQDGT01-26W-IMI.xlsm   │  ──Push to Hub──▶   │  registry.csv            │  ◀── Refresh Dashboard ─── Nursery_Hub.xlsm
   ├──────────────────────────┤                     │  (in shared folder)      │
   │  AUGT3-27S-IMI.xlsm      │  ──Push to Hub──▶   │                          │
   └──────────────────────────┘                     └──────────────────────────┘
```

- Each nursery workbook has a **Hub registry folder** setting on its Settings
  tab (default `~/Documents/PacificSeeds/Nurseries/`).
- "Push to Hub" (Step 13, also runs automatically on Save) writes one row to
  `registry.csv` in that folder, replacing any prior row for the same nursery
  code. Columns: `nursery_code, season, breeder, packets, replacements,
  errors, additionals, file_path, last_update`.
- The Hub's "Refresh Dashboard" reads that CSV, populates its Dashboard tab and
  the four headline metrics on Home.

This means:
- No server, no shared spreadsheet to corrupt.
- Works offline; the registry is plain CSV.
- A new nursery joins the dashboard the first time it's saved.
- Closing one workbook never affects others.

---

## Per-nursery workflow at a glance

| Step | Button | What it does |
|---|---|---|
| 1 | Initialise from PRISM export | Asks for nursery code, stamps Nursery data |
| 2 | Build Nursery list | Unique Source IDs sorted A→Z + Repeats + Qty Required + BULK flag |
| 3 | Design Field Map | Activates the Field Map tab |
| 4 | Generate Packet Prep | Creates Plot / Spike / Rack / QR for every packet |
| 5 | Sort for racking | LSD radix sort (rack ↑ then spike ↑) |
| 6 | Add replacement | Quick form: stage, original source, replaced with |
| 7 | Add planting error | Quick form: severity, note |
| 8 | Record spray | TFMSA / IMI / HPPD + date + rate |
| 9 | AB bag pulling | Bag count + date |
| 10 | Import updated PRISM | Activates Updated nursery site tab |
| 11 | Generate Fieldbook | Reorders cols, serpentine sort, landscape print |
| 12 | Refresh dashboard | Recalculates the Home stat panel |
| 13 | Push to Hub | Writes summary row to shared registry.csv |

---

## Why this is scalable

- Every workbook is self-contained — no Excel-level cross-workbook links.
- Sync is a single CSV row, so 100 nurseries cost ~10 KB on disk.
- All formulas live on the Home dashboard, so opening the file always shows
  fresh numbers.
- The Hub's Dashboard rebuilds in seconds even with hundreds of nurseries.
- The same `.xlsm` template can be cloned for any nursery — the workflow code
  has no nursery-specific state baked in.
