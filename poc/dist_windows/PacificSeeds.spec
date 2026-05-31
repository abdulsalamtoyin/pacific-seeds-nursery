# PyInstaller spec for the Pacific Seeds desktop launcher.
# Build with:  pyinstaller dist_windows/PacificSeeds.spec --clean
# Output:      dist/PacificSeeds.exe (a single file ~100 MB)

import os
from pathlib import Path

PROJECT = Path(SPECPATH).parent

# Bundle the PWA static files + backend + scripts as data.
datas = [
    (str(PROJECT / "pwa"),     "pwa"),
    (str(PROJECT / "backend"), "backend"),
    (str(PROJECT / "scripts"), "scripts"),
]

# Hidden imports for things PyInstaller can't auto-detect (string-loaded modules).
hiddenimports = [
    "uvicorn",
    "uvicorn.logging",
    "uvicorn.config",
    "uvicorn.loops",
    "uvicorn.loops.auto",
    "uvicorn.loops.asyncio",
    "uvicorn.protocols",
    "uvicorn.protocols.http",
    "uvicorn.protocols.http.auto",
    "uvicorn.protocols.http.h11_impl",
    "uvicorn.protocols.websockets",
    "uvicorn.protocols.websockets.auto",
    "uvicorn.protocols.websockets.wsproto_impl",
    "uvicorn.lifespan",
    "uvicorn.lifespan.on",
    "logging.config",
    "fastapi",
    "fastapi.staticfiles",
    "fastapi.responses",
    "starlette",
    "starlette.routing",
    "starlette.responses",
    "starlette.staticfiles",
    "starlette.middleware",
    "starlette.middleware.cors",
    "multipart",
    "python_multipart",
    "anyio",
    "anyio._backends",
    "anyio._backends._asyncio",
    "sniffio",
    "openpyxl",
    "openpyxl.cell._writer",
    "qrcode",
    "qrcode.image.pil",
    "reportlab",
    "reportlab.pdfgen",
    "reportlab.pdfgen.canvas",
    "reportlab.lib",
    "reportlab.lib.pagesizes",
    "reportlab.platypus",
    "PIL",
    "PIL.Image",
    "PIL.ImageDraw",
    "PIL._tkinter_finder",
    "pystray",
    "pystray._win32",
    "nursery_init",
    "map_parser",
    "fieldbook_export",
    "backend.app",
]

a = Analysis(
    [str(PROJECT / "dist_windows" / "launcher.py")],
    pathex=[str(PROJECT), str(PROJECT / "backend"), str(PROJECT / "scripts")],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["matplotlib", "tkinter", "jupyter", "notebook", "IPython"],
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="PacificSeeds",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                # no terminal window on launch
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(PROJECT / "pwa" / "ps-logo.svg")
        if (PROJECT / "pwa" / "ps-logo.ico").exists() else None,
)
