# Building the Windows .exe

This folder contains everything needed to compile the Pacific Seeds Nursery
Fieldbook into a single-file Windows executable + a polished installer.

```
dist_windows/
├── launcher.py         ← entry point: starts FastAPI, opens browser, tray icon
├── PacificSeeds.spec   ← PyInstaller config (bundles Python + all deps + PWA)
├── installer.iss       ← Inno Setup script — produces a real installer.exe
├── build.bat           ← run this on Windows to produce dist/PacificSeeds.exe
└── README.md           ← this file
```

## How it runs at runtime

```
   ┌──── PacificSeeds.exe ────┐
   │                          │
   │  launcher.py             │
   │   ├─ chooses free port   │
   │   ├─ starts uvicorn      │       Browser
   │   │   in a thread        │ ◀──── opens at
   │   ├─ waits for /ready    │       http://localhost:PORT/
   │   ├─ opens browser       │
   │   └─ shows tray icon     │       Tray menu:
   │                          │         "Open in browser"
   │  Bundled inside:         │         "Open data folder"
   │   pwa/   backend/  scripts/         "Quit"
   │   FastAPI · openpyxl · qrcode · reportlab · uvicorn …
   │                          │
   └──────────────────────────┘

   User data (per-user, never inside Program Files):
     %LOCALAPPDATA%\PacificSeeds\
       data\nursery.sqlite
       output\<nursery>_packets.pdf
       output\<nursery>_fieldbook.xlsx
```

---

## Three ways to produce the .exe

### Option 1 — Build it yourself on Windows (10 minutes)

**Prerequisites:**
- Windows 10/11
- [Python 3.11](https://www.python.org/downloads/) (tick "Add Python to PATH")
- That's it.

**Steps:**

```cmd
git clone <your-repo-url>
cd <repo>\poc
dist_windows\build.bat
```

When it finishes you'll have `poc\dist\PacificSeeds.exe` (~100 MB).
Double-click it — your default browser opens at
`http://localhost:8765/` showing the landing page.

### Option 2 — Build the installer too (extra 5 minutes)

1. Do Option 1 first.
2. Install [Inno Setup](https://jrsoftware.org/isinfo.php) (free).
3. Open `dist_windows\installer.iss` in Inno Setup Compiler.
4. **Build → Compile** (or `F9`).
5. Output: `poc\dist\installer\PacificSeedsSetup-1.0.0.exe` — that's the
   distributable installer. Email it; double-click; standard Windows installer
   flow with Start Menu + optional Desktop shortcut.

### Option 3 — Have GitHub Actions do it for you (zero local setup)

The repo has `.github/workflows/build-windows.yml`. On every git tag like
`v1.0.0`, GitHub will spin up a clean Windows machine, build the .exe + the
installer, and upload both as release artifacts. To trigger manually:

1. Go to your repo on GitHub → **Actions** tab.
2. Pick **Build Windows .exe** → **Run workflow** → Run.
3. After ~5 minutes, download `PacificSeeds-Windows` and
   `PacificSeeds-Windows-Installer` artifacts from the run.

---

## What the technician sees

After running the installer:

1. Start Menu has "Pacific Seeds Nursery Fieldbook" (and Desktop shortcut if
   they ticked the box).
2. Clicking it shows nothing for ~3 seconds (uvicorn starting), then their
   default browser opens at the landing page.
3. A PS-blue tray icon appears in the system tray. Right-click → menu:
   - **Open in browser** — re-opens the tab if they closed it
   - **Open data folder** — opens `%LOCALAPPDATA%\PacificSeeds\` in Explorer
   - **Quit** — gracefully stops the server

Data persists across launches (SQLite file lives in the data folder).

---

## Sanity checks before distributing

1. Build, install, launch on a clean Windows VM (no Python pre-installed).
2. Confirm the landing page loads.
3. Click **▶ Open the app** → should go to `/app`.
4. Initialise a nursery using a real PRISM file.
5. Quit via tray menu. Re-launch. Confirm SQLite data persisted.
6. Run the uninstaller. Confirm data folder is left intact (user data is
   preserved unless they delete it themselves).

---

## Troubleshooting build issues

| Symptom | Cause / Fix |
|---|---|
| `pyinstaller: command not found` | Activate the venv first (`.venv-win\Scripts\activate`) |
| `.exe runs but Browser tab shows ERR_CONNECTION_REFUSED` | Antivirus is blocking the bundled Python. Add `PacificSeeds.exe` to the AV's allowlist. |
| `ModuleNotFoundError: No module named ‘X'` at runtime | Add X to `hiddenimports` in `PacificSeeds.spec`, rebuild |
| Tray icon doesn't appear | `pystray` failed to import — `pip install pystray pillow` then rebuild |
| `.exe` is enormous (~250 MB) | Add modules you don't need to `excludes=[…]` in the spec |
| Inno Setup says "icon not found" | Place a `ps-logo.ico` in `pwa/` (export the SVG via [Online Convert](https://image.online-convert.com/convert-to-ico)) |

---

## Notes & gotchas

- **Antivirus false positives.** PyInstaller-built .exe files frequently trip
  AV heuristics (especially on small companies' machines). The proper fix is
  to **sign the .exe** with a code-signing certificate. For internal
  distribution, allowlisting the .exe on each tech's machine works in a pinch.
- **First launch is slow.** ~3-5 seconds while uvicorn warms up. Subsequent
  launches are quicker because Windows caches the unpacked PyInstaller
  bundle.
- **No admin needed.** The installer uses `PrivilegesRequired=lowest` so
  techs can install without IT involvement.
- **The .exe must run on the tech's machine** (not network drive). The
  PyInstaller bootloader unpacks to a temp folder per run.
- **One user = one database.** SQLite lives in `%LOCALAPPDATA%` so each
  Windows user account gets its own data. To share across techs, point them
  at a network folder (manual step — change `PS_DATA_DIR` env before
  launching, or modify `launcher.py`).
