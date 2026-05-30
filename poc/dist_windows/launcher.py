"""Pacific Seeds — Nursery Fieldbook desktop launcher.

This is the entry point that PyInstaller bundles into a single .exe.
On double-click:
  1. Locate (or create) a writable data folder under %LOCALAPPDATA%.
  2. Start FastAPI in a background thread on the first free port from a
     short list (8765, 8766, …).
  3. Open the user's default browser to the landing page.
  4. Show a tray icon so the user can stop the server gracefully.
"""
from __future__ import annotations

import os
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path


# ───── Locate bundled resources & writable data folder ─────

def resource_path(*parts: str) -> Path:
    """Return absolute path to a bundled resource (works in dev and PyInstaller)."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base.joinpath(*parts)


def data_dir() -> Path:
    """Writable per-user data folder. On Windows: %LOCALAPPDATA%\\PacificSeeds."""
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


# Wire the bundled app to use the user's data folder before any module imports it.
USER_ROOT = data_dir()
os.environ.setdefault("PS_DATA_DIR", str(USER_ROOT / "data"))
os.environ.setdefault("PS_OUTPUT_DIR", str(USER_ROOT / "output"))

# Add bundled source dirs to sys.path so `import backend.app` etc. work.
_ROOT = resource_path()
for sub in ("", "backend", "scripts", "pwa"):
    p = _ROOT / sub if sub else _ROOT
    if p.exists() and str(p) not in sys.path:
        sys.path.insert(0, str(p))


# ───── Find a free port ─────

def first_free_port(candidates: list[int]) -> int:
    for port in candidates:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    # Last resort: ask the OS for an ephemeral port.
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


# ───── Server thread ─────

class ServerThread(threading.Thread):
    def __init__(self, host: str, port: int):
        super().__init__(daemon=True, name="uvicorn")
        self.host = host
        self.port = port
        self._server = None

    def run(self):
        import uvicorn
        from backend.app import app
        config = uvicorn.Config(app, host=self.host, port=self.port,
                                log_level="warning")
        self._server = uvicorn.Server(config)
        try:
            self._server.run()
        except Exception as e:
            print(f"Server error: {e}", file=sys.stderr)

    def stop(self):
        if self._server is not None:
            self._server.should_exit = True


def wait_until_ready(host: str, port: int, timeout: float = 12.0) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        try:
            with socket.create_connection((host, port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.2)
    return False


# ───── Tray icon (optional; falls back to console if pystray missing) ─────

def run_tray(server: ServerThread, url: str):
    try:
        from PIL import Image, ImageDraw
        import pystray
    except Exception:
        print(f"\nPacific Seeds running at {url}")
        print("Close this window to stop the server.\n")
        try:
            while server.is_alive():
                time.sleep(1)
        except KeyboardInterrupt:
            pass
        return

    # Build a simple PS-blue icon at runtime
    img = Image.new("RGBA", (64, 64), (9, 42, 64, 255))
    d = ImageDraw.Draw(img)
    d.ellipse((10, 10, 54, 54), fill=(6, 120, 205, 255))
    d.text((22, 18), "PS", fill=(255, 255, 255, 255))

    def on_open(_icon, _item):
        webbrowser.open(url)

    def on_open_data(_icon, _item):
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(USER_ROOT))
            elif sys.platform == "darwin":
                os.system(f'open "{USER_ROOT}"')
            else:
                os.system(f'xdg-open "{USER_ROOT}"')
        except Exception:
            pass

    def on_quit(icon, _item):
        icon.stop()
        server.stop()

    menu = pystray.Menu(
        pystray.MenuItem(f"Open in browser ({url})", on_open, default=True),
        pystray.MenuItem("Open data folder", on_open_data),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Quit", on_quit),
    )
    icon = pystray.Icon("PacificSeeds", img, "Pacific Seeds — Nursery", menu)
    icon.run()


# ───── Main ─────

def main() -> int:
    host = "127.0.0.1"
    port = first_free_port([8765, 8766, 8767, 8000, 8080])
    url = f"http://{host}:{port}/"

    print("Pacific Seeds — Nursery Fieldbook")
    print(f"  data: {USER_ROOT}")
    print(f"  url : {url}")

    server = ServerThread(host, port)
    server.start()

    if not wait_until_ready(host, port):
        print("Server failed to start in time.", file=sys.stderr)
        return 1

    webbrowser.open(url)
    run_tray(server, url)
    server.stop()
    server.join(timeout=2)
    return 0


if __name__ == "__main__":
    sys.exit(main())
