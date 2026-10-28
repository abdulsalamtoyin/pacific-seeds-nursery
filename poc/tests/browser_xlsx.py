"""Run the browser's spreadsheet modules under node, for the parity tests.

The app reads PRISM files and writes exports in the browser so it can be
hosted as static files. Its Python twins stay as the reference: each test
hands the same input to both and compares what comes out. This helper is the
browser half — one node process per call, the vendored ExcelJS bundle loaded
onto globalThis exactly as the page's script tag would.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
from pathlib import Path

import pytest

PWA = Path(__file__).resolve().parent.parent / "pwa"

DRIVER = """
import {{ createRequire }} from "node:module";
import {{ readFileSync, writeFileSync }} from "node:fs";
globalThis.ExcelJS = createRequire(import.meta.url)({bundle!r});
const writer = await import({writer!r});
const reader = await import({reader!r});

const job = JSON.parse(readFileSync(process.argv[2], "utf8"));
try {{
  let result = null;
  if (job.op === "write") {{
    writeFileSync(job.out, Buffer.from(await writer.workbookBytes(job.payload)));
  }} else if (job.op === "sanitise") {{
    result = writer.sanitiseSheetName(
      job.name, job.taken === null ? null : new Set(job.taken));
  }} else if (job.op === "read") {{
    const bytes = readFileSync(job.path);
    result = await reader.readPrism(new File([bytes], job.name));
  }}
  console.log(JSON.stringify({{ ok: true, result }}));
}} catch (err) {{
  console.log(JSON.stringify({{ ok: false, error: err.message }}));
}}
"""


class BrowserError(Exception):
    """The browser module threw; carries its message."""


def _run(job: dict):
    if shutil.which("node") is None:
        pytest.skip("node not installed — cannot exercise the browser modules")

    with tempfile.TemporaryDirectory() as tmp:
        driver = Path(tmp) / "driver.mjs"
        driver.write_text(DRIVER.format(
            bundle=str(PWA / "vendor" / "exceljs.min.js"),
            writer=str(PWA / "xlsx-writer.js"),
            reader=str(PWA / "prism-reader.js"),
        ), encoding="utf-8")
        job_file = Path(tmp) / "job.json"
        out = Path(tmp) / "out.xlsx"
        job_file.write_text(json.dumps({**job, "out": str(out)}),
                            encoding="utf-8")

        proc = subprocess.run(["node", str(driver), str(job_file)],
                              capture_output=True, text=True, timeout=60)
        if proc.returncode != 0:
            pytest.fail(f"node driver failed:\n{proc.stderr}")
        reply = json.loads(proc.stdout.strip().splitlines()[-1])
        if not reply["ok"]:
            raise BrowserError(reply["error"])
        if job["op"] == "write":
            return out.read_bytes()
        return reply["result"]


def workbook_bytes(payload: dict) -> bytes:
    return _run({"op": "write", "payload": payload})


def sanitise_sheet_name(name: str, taken: set[str] | None = None) -> str:
    return _run({"op": "sanitise", "name": name,
                 "taken": None if taken is None else sorted(taken)})


def read_prism(path: Path, name: str | None = None) -> dict:
    return _run({"op": "read", "path": str(path), "name": name or path.name})
