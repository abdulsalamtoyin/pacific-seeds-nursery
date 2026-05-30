# Pacific Seeds Nursery Workflow

End-to-end digital workflow for the Sorghum R&D nursery program. Replaces
the paper + Excel pipeline with a phone-friendly PWA, an offline-first event
capture system, and a Hub dashboard.

Two parallel delivery formats:

1. **Web app** (FastAPI + PWA) — primary path. Browse to it; install as a
   Progressive Web App on phones for field use. Compiles to a single Windows
   `.exe` via PyInstaller (this repo's GitHub Actions workflow builds it).
2. **Excel + VBA workbook pack** — drop-in upgrade to the existing VBA-based
   workflow for breeders who prefer Excel. Lives in `poc/excel_workflow/`.

## Quick start (development)

```bash
cd poc
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/uvicorn backend.app:app --host 0.0.0.0 --port 8000 --reload
```

Open <http://localhost:8000/> for the landing page,
<http://localhost:8000/app> for the workflow PWA,
or <http://localhost:8000/dashboard> for the Hub.

## Build the Windows .exe

```bash
# Option A — Windows machine
cd poc
dist_windows\build.bat
# → poc\dist\PacificSeeds.exe

# Option B — GitHub Actions (no Windows needed)
# Push this repo to GitHub → Actions tab → "Build Windows .exe" → Run workflow
# → download artifact when it finishes
```

Full build details: `poc/dist_windows/README.md`.

## Project layout

```
.
├── .github/workflows/build-windows.yml     ← CI to build the .exe
├── poc/
│   ├── backend/app.py                      ← FastAPI server
│   ├── pwa/                                ← landing + workflow PWA
│   ├── scripts/                            ← nursery_init, fieldbook_export, map_parser
│   ├── dist_windows/                       ← PyInstaller launcher + spec + installer
│   ├── excel_workflow/                     ← parallel Excel+VBA distribution
│   └── requirements.txt
└── README.md (this file)
```
