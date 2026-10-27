"""The single-file build must open by double-clicking.

A file:// page cannot load modules, vendored scripts or images beside it, so
the build has to leave nothing outside the one file, and the inlined bundle
has to be valid script once every import/export has been rewritten.
"""
from __future__ import annotations

import importlib.util
import re
import shutil
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location(
    "build_single_html", ROOT / "scripts" / "build_single_html.py")
build_single_html = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(build_single_html)


@pytest.fixture(scope="module")
def html() -> str:
    return build_single_html.build()


def scripts(html: str) -> list[str]:
    return re.findall(r"<script>(.*?)</script>", html, re.S)


def test_nothing_is_loaded_from_beside_the_file(html):
    assert 'type="module"' not in html
    assert not re.search(r'(?:src|href)="(?!data:|#|https?:)', html)
    assert len(scripts(html)) == 2          # ExcelJS, then the app


def test_every_module_is_bundled_once_dependencies_first():
    order = build_single_html.module_order("app.js")
    assert order[-1] == "app.js"
    assert len(order) == len(set(order))
    assert order.index("store.js") < order.index("grid.js")
    assert "file-save.js" in order


def test_sync_is_always_off_in_the_single_file(html):
    assert 'clientId: ""' in scripts(html)[1]


def test_bundle_is_valid_script(html, tmp_path):
    if shutil.which("node") is None:
        pytest.skip("node not installed")
    bundle = tmp_path / "bundle.js"
    bundle.write_text(scripts(html)[1], encoding="utf-8")
    # Parse only — `new Function` compiles the body without running it.
    proc = subprocess.run(
        ["node", "-e", "new Function(require('fs').readFileSync(process.argv[1], 'utf8'))",
         str(bundle)], capture_output=True, text=True, timeout=30)
    assert proc.returncode == 0, proc.stderr
