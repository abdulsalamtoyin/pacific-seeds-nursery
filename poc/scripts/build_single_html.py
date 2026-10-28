"""Pack the browser app into one HTML file that opens by double-clicking.

No server, no Azure: the file carries every module, ExcelJS and the logos.
Nurseries are kept in the browser and, from Chrome or Edge, in a save file
the user picks in their OneDrive folder (pwa/file-save.js).

Why a bundler of our own: browsers refuse to load `import` from a file://
page, so the modules have to become one inline script. The app's modules
only use the plain forms of import/export, so each is wrapped in a function
and its imports become lookups — and anything fancier fails the build
loudly rather than producing a page that opens blank.

Usage:
    python scripts/build_single_html.py
    python scripts/build_single_html.py --out ~/Desktop/Nursery.html
"""
from __future__ import annotations

import argparse
import base64
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PWA = ROOT / "pwa"
DEFAULT_OUT = ROOT / "dist" / "single" / "Nursery Workbook.html"
ENTRY = "app.js"

# OneDrive sync needs MSAL and a hosted redirect page, neither of which a
# local file can have — so the single file is always built with it off.
STUB_MODULES = {
    "sync-config.js":
        'export const SYNC_CONFIG = { clientId: "", tenantId: "", scopes: [] };\n',
}

IMPORT_NAMED = re.compile(
    r'^import\s*\{([^}]*)\}\s*from\s*"\./([\w.-]+\.js)";', re.M)
IMPORT_STAR = re.compile(
    r'^import\s*\*\s*as\s+(\w+)\s+from\s*"\./([\w.-]+\.js)";', re.M)
EXPORT_DECL = re.compile(
    r'^export\s+((?:async\s+)?function\*?|class|const|let|var)\s+(\w+)', re.M)
LEFTOVER = re.compile(r'^\s*(import|export)\b', re.M)
MODULE_TAG = re.compile(
    r'(?:\s*<!--(?:(?!-->).)*-->)?\s*<script type="module" src="app.js"></script>', re.S)


class BundleError(Exception):
    pass


def source(name: str) -> str:
    if name in STUB_MODULES:
        return STUB_MODULES[name]
    path = PWA / name
    if not path.exists():
        raise BundleError(f"{name} is imported but pwa/{name} does not exist")
    return path.read_text(encoding="utf-8")


def deps(text: str) -> list[str]:
    return [m.group(2) for m in IMPORT_NAMED.finditer(text)] + \
        [m.group(2) for m in IMPORT_STAR.finditer(text)]


def module_order(entry: str) -> list[str]:
    """Dependencies before dependants. A cycle cannot be wrapped this way."""
    order: list[str] = []
    visiting: set[str] = set()

    def visit(name: str) -> None:
        if name in order:
            return
        if name in visiting:
            raise BundleError(f"import cycle through {name}")
        visiting.add(name)
        for dep in deps(source(name)):
            visit(dep)
        visiting.discard(name)
        order.append(name)

    visit(entry)
    return order


def named_bindings(spec: str) -> str:
    """`a, b as c` → `a, b: c` — the same names, as destructuring."""
    parts = [p.strip() for p in spec.split(",") if p.strip()]
    return ", ".join(re.sub(r"\s+as\s+", ": ", p) for p in parts)


def wrap(name: str) -> str:
    text = source(name)
    text = IMPORT_NAMED.sub(
        lambda m: f"const {{ {named_bindings(m.group(1))} }} = __mod[{m.group(2)!r}];",
        text)
    text = IMPORT_STAR.sub(
        lambda m: f"const {m.group(1)} = __mod[{m.group(2)!r}];", text)
    exported = [m.group(2) for m in EXPORT_DECL.finditer(text)]
    text = EXPORT_DECL.sub(lambda m: f"{m.group(1)} {m.group(2)}", text)
    # Only reached for vendored scripts and the sign-in redirect, both of which
    # the single file never loads; the page's own address keeps them valid.
    text = text.replace("import.meta.url", "document.baseURI")

    leftover = LEFTOVER.search(text)
    if leftover:
        line = text.count("\n", 0, leftover.start()) + 1
        raise BundleError(
            f"pwa/{name}:{line}: an import/export form the bundler does not "
            'handle — use `import {…} from "./x.js"`, `import * as x`, or '
            "`export function/class/const`")
    return (f"__mod[{name!r}] = (() => {{\n{text}\n"
            f"return {{ {', '.join(exported)} }};\n}})();\n")


def inline_script(js: str) -> str:
    # The HTML parser ends a script at the first "</script", wherever it is.
    return js.replace("</script", "<\\/script")


def data_uri(path: Path) -> str:
    return "data:image/svg+xml;base64," + base64.b64encode(path.read_bytes()).decode()


def build() -> str:
    bundle = '"use strict";\nconst __mod = {};\n' + \
        "".join(wrap(name) for name in module_order(ENTRY))

    html = (PWA / "index.html").read_text(encoding="utf-8")
    for svg in ("ps-logo.svg", "ps-logo-white.svg"):
        html = html.replace(f'"{svg}"', f'"{data_uri(PWA / svg)}"')

    tag = MODULE_TAG.search(html)
    if not tag:
        raise BundleError('index.html no longer loads <script type="module" src="app.js">')

    exceljs = (PWA / "vendor" / "exceljs.min.js").read_text(encoding="utf-8")
    scripts = (
        "\n  <!-- Built by scripts/build_single_html.py — edit pwa/, not this file. -->\n"
        f"  <script>{inline_script(exceljs)}</script>\n"
        f"  <script>\n{inline_script(bundle)}</script>\n"
    )
    html = html[:tag.start()] + scripts + html[tag.end():]

    local = re.search(r'(?:src|href)="(?!data:|#|https?:)[^"]+"', html)
    if local:
        raise BundleError(f"index.html still references a local file: {local.group(0)}")
    return html


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT)
    args = parser.parse_args()
    try:
        html = build()
    except BundleError as err:
        print(f"build failed: {err}", file=sys.stderr)
        return 1
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(html, encoding="utf-8")
    print(f"wrote {args.out} ({len(html.encode()) / 1e6:.1f} MB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
