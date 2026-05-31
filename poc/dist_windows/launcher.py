"""Pacific Seeds — Nursery Fieldbook desktop launcher (debug-friendly).

What this does on double-click:
  1. Set up a log file at %LOCALAPPDATA%\\PacificSeeds\\launcher.log
  2. Locate (or create) writable per-user data dirs.
  3. Start FastAPI in a background thread.
  4. Wait for it, open the default browser, show a tray icon.
  5. If ANYTHING fails, write the traceback to launcher.log AND pop up
     a Windows MessageBox so the user (and we) can see what went wrong —
     no more silent failures.
"""
from __future__ import annotations

import logging
import os
import socket
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path


# ───────────── Per-user data folder + early logging setup ─────────────

def _data_dir() -> Path:
    if sys.platform.startswith("win"):
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") \
               or str(Path.home() / "AppData" / "Local")
        p = Path(base) / "PacificSeeds"
    elif sys.platform == "darwin":
        p = Path.home() / "Library" / "Application Support" / "PacificSeeds"
    else:
        p = Path.home() / ".local" / "share" / "PacificSeeds"
    p.mkdir(parents=True, exist_ok=True)
    (p / "data").mkdir(exist_ok=True)
    (p / "output").mkdir(exist_ok=True)
    return p


USER_ROOT = _data_dir()
LOG_FILE = USER_ROOT / "launcher.log"

# Send everything to launcher.log AND stderr (when there's a console).
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("ps-launcher")
log.info("=== Pacific Seeds launcher start ===")
log.info("python: %s", sys.version)
log.info("platform: %s", sys.platform)
log.info("executable: %s", sys.executable)
log.info("argv: %s", sys.argv)
log.info("_MEIPASS: %s", getattr(sys, "_MEIPASS", "(not set — not frozen)"))
log.info("USER_ROOT: %s", USER_ROOT)

# Tell the FastAPI app where to put data + outputs.
os.environ.setdefault("PS_DATA_DIR", str(USER_ROOT / "data"))
os.environ.setdefault("PS_OUTPUT_DIR", str(USER_ROOT / "output"))


def show_error_popup(title: str, message: str) -> None:
    """Show a native Windows message box. No-op on other platforms."""
    log.error("Popup: %s — %s", title, message)
    if sys.platform.startswith("win"):
        try:
            import ctypes
            MB_ICONERROR = 0x10
            ctypes.windll.user32.MessageBoxW(None, message, title, MB_ICONERROR)
        except Exception as e:
            log.error("MessageBoxW failed: %s", e)


# ───────────── Bundle resource discovery ─────────────

def _resource_path(*parts: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base.joinpath(*parts)


_ROOT = _resource_path()
for sub in ("", "backend", "scripts", "pwa"):
    p = _ROOT / sub if sub else _ROOT
    if p.exists() and str(p) not in sys.path:
        sys.path.insert(0, str(p))
log.info("bundle root: %s", _ROOT)
log.info("sys.path entries (first 6): %s", sys.path[:6])


# ───────────── Networking ─────────────

def _first_free_port(candidates: list[int]) -> int:
    for port in candidates:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _wait_until_ready(host: str, port: int, timeout: float = 20.0) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        try:
            with socket.create_connection((host, port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.2)
    return False


# ───────────── Server thread ─────────────

class ServerThread(threading.Thread):
    def __init__(self, host: str, port: int):
        super().__init__(daemon=True, name="uvicorn")
        self.host = host
        self.port = port
        self._server = None
        self.startup_error: Exception | None = None

    def run(self) -> None:
        try:
            log.info("Importing uvicorn + backend.app …")
            import uvicorn
            from backend.app import app
            log.info("Imports OK. Building uvicorn Config (host=%s port=%s) …",
                     self.host, self.port)
            config = uvicorn.Config(app, host=self.host, port=self.port,
                                    log_level="warning")
            self._server = uvicorn.Server(config)
            log.info("Starting uvicorn.Server.run() …")
            self._server.run()
            log.info("uvicorn.Server.run() returned cleanly")
        except Exception as e:
            self.startup_error = e
            log.exception("Server thread crashed: %s", e)

    def stop(self) -> None:
        if self._server is not None:
            self._server.should_exit = True


# ───────────── Tray icon ─────────────

def _run_tray(server: ServerThread, url: str) -> None:
    try:
        from PIL import Image, ImageDraw
        import pystray
    except Exception as e:
        log.warning("pystray/PIL import failed (%s) — falling back to console loop", e)
        try:
            while server.is_alive():
                time.sleep(1)
        except KeyboardInterrupt:
            pass
        return

    img = Image.new("RGBA", (64, 64), (9, 42, 64, 255))
    d = ImageDraw.Draw(img)
    d.ellipse((10, 10, 54, 54), fill=(6, 120, 205, 255))
    d.text((22, 18), "PS", fill=(255, 255, 255, 255))

    def on_open(_icon, _item): webbrowser.open(url)

    def on_open_data(_icon, _item):
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(USER_ROOT))
            elif sys.platform == "darwin":
                os.system(f'open "{USER_ROOT}"')
            else:
                os.system(f'xdg-open "{USER_ROOT}"')
        except Exception as e:
            log.error("Open data folder failed: %s", e)

    def on_open_log(_icon, _item):
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(LOG_FILE))
        except Exception as e:
            log.error("Open log failed: %s", e)

    def on_quit(icon, _item):
        log.info("Quit requested via tray menu")
        icon.stop()
        server.stop()

    menu = pystray.Menu(
        pystray.MenuItem(f"Open in browser ({url})", on_open, default=True),
        pystray.MenuItem("Open data folder", on_open_data),
        pystray.MenuItem("Open launcher.log", on_open_log),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Quit", on_quit),
    )
    log.info("Showing tray icon")
    icon = pystray.Icon("PacificSeeds", img, "Pacific Seeds — Nursery", menu)
    icon.run()


# ───────────── Main ─────────────

def main() -> int:
    try:
        host = "127.0.0.1"
        port = _first_free_port([8765, 8766, 8767, 8000, 8080])
        url = f"http://{host}:{port}/"
        log.info("chosen url: %s", url)

        server = ServerThread(host, port)
        server.start()

        if not _wait_until_ready(host, port, timeout=20.0):
            # Did the server thread crash, or just hang?
            if server.startup_error:
                show_error_popup(
                    "Pacific Seeds — couldn't start",
                    f"The backend failed to start.\n\n"
                    f"Error: {type(server.startup_error).__name__}: {server.startup_error}\n\n"
                    f"Full log:\n{LOG_FILE}",
                )
            else:
                show_error_popup(
                    "Pacific Seeds — couldn't start",
                    f"The backend didn't open port {port} within 20 seconds.\n\n"
                    f"Possible causes:\n"
                    f"  • Antivirus / firewall blocking the bundled Python\n"
                    f"  • Another instance is already running\n"
                    f"  • Slow first launch — try double-clicking again\n\n"
                    f"Full log:\n{LOG_FILE}",
                )
            return 1

        log.info("Server is up. Opening browser …")
        try:
            opened = webbrowser.open(url)
            log.info("webbrowser.open returned %s", opened)
        except Exception as e:
            log.error("webbrowser.open failed: %s", e)

        _run_tray(server, url)
        server.stop()
        server.join(timeout=2)
        log.info("Launcher exiting cleanly")
        return 0
    except SystemExit:
        raise
    except Exception as e:
        log.exception("Unexpected launcher crash: %s", e)
        show_error_popup(
            "Pacific Seeds — unexpected error",
            f"{type(e).__name__}: {e}\n\nFull traceback in:\n{LOG_FILE}",
        )
        return 1


if __name__ == "__main__":
    sys.exit(main())
