"""Saving to a file in OneDrive must never lose a nursery.

The file is shared between computers by the OneDrive desktop app, so the risk
is one computer overwriting what another saved. Each case runs the real store
and file-save engine under node; "two computers" are two store instances
writing one fake file, whose handles behave like the browser's
FileSystemFileHandle as far as the engine can see.
"""
from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

import pytest

PWA = Path(__file__).resolve().parent.parent / "pwa"

PREAMBLE = """
import {{ pathToFileURL }} from "node:url";
const url = (p) => pathToFileURL(p).href;
const fs = await import(url({filesave!r}));

const mem = () => {{
  const m = new Map();
  return {{ getItem: (k) => (m.has(k) ? m.get(k) : null),
           setItem: (k, v) => m.set(k, String(v)), removeItem: (k) => m.delete(k) }};
}};

// One file on "OneDrive", seen through a handle per computer.
function fakeFile(name = "nurseries.json", text = "") {{
  const file = {{ text, writes: 0 }};
  file.handle = (permission = "granted") => ({{
    name,
    queryPermission: async () => permission,
    requestPermission: async () => (permission = "granted"),
    getFile: async () => ({{ text: async () => file.text }}),
    createWritable: async () => {{
      let buf = "";
      return {{ write: async (t) => {{ buf += t; }},
               close: async () => {{ file.text = buf; file.writes++; }} }};
    }},
  }});
  return file;
}}

// A computer: its own store instance (the ?tag makes node load a fresh copy)
// and its own remembered handle.
async function computer(tag, {{ pick, saved = null, confirmReplace }} = {{}}) {{
  const store = await import(url({store!r}) + "?" + tag);
  store.useBackend(mem());
  let record = saved;
  const meta = {{ load: async () => record, remember: async (m) => {{ record = m; }},
                 forget: async () => {{ record = null; }} }};
  const engine = fs.createFileSave({{
    store, meta, debounceMs: 0, confirmReplace,
    pickers: {{ pickNew: async () => pick(), pickExisting: async () => pick() }},
  }});
  return {{ store, engine, meta: () => record }};
}}
const codes = (c) => Object.values(c.store.snapshot()).map((n) => n.code).sort();
const onDisk = (f) => Object.values(JSON.parse(f.text).nurseries).map((n) => n.code).sort();
const tick = () => new Promise((r) => setTimeout(r, 5));
const out = {{}};
"""


def run(body: str) -> dict:
    if shutil.which("node") is None:
        pytest.skip("node not installed — cannot exercise file-save.js")
    script = PREAMBLE.format(filesave=str(PWA / "file-save.js"),
                             store=str(PWA / "store.js")) + body + \
        "\nconsole.log(JSON.stringify(out));\n"
    proc = subprocess.run(["node", "--input-type=module", "-e", script],
                          capture_output=True, text=True, timeout=30)
    if proc.returncode != 0:
        pytest.fail(f"node driver failed:\n{proc.stderr}")
    return json.loads(proc.stdout)


def test_new_file_gets_every_nursery_and_follows_edits():
    out = run("""
const file = fakeFile();
const a = await computer("a", { pick: () => file.handle() });
a.store.createNursery({ code: "A-1" });
await a.engine.chooseNew();
out.first = onDisk(file);
a.store.createNursery({ code: "A-2" });
await tick(); await a.engine.saveNow();
out.second = onDisk(file);
out.state = a.engine.status().state;
""")
    assert out["first"] == ["A-1"]
    assert out["second"] == ["A-1", "A-2"]
    assert out["state"] == "saved"


def test_second_computer_opening_the_file_gets_the_nurseries():
    out = run("""
const file = fakeFile();
const a = await computer("a", { pick: () => file.handle() });
a.store.createNursery({ code: "OFFICE" });
await a.engine.chooseNew();

const b = await computer("b", { pick: () => file.handle() });
b.store.createNursery({ code: "FIELD" });
out.opened = await b.engine.openExisting();
out.b = codes(b);
out.file = onDisk(file);       // and B's own nursery went into the file
""")
    assert out["opened"] is True
    assert out["b"] == ["FIELD", "OFFICE"]
    assert out["file"] == ["FIELD", "OFFICE"]


def test_save_from_another_computer_is_never_overwritten_silently():
    out = run("""
const file = fakeFile();
const a = await computer("a", { pick: () => file.handle() });
a.store.createNursery({ code: "X" });
await a.engine.chooseNew();
const b = await computer("b", { pick: () => file.handle() });
await b.engine.openExisting();

// B saves; A has not seen it and edits too.
b.store.createNursery({ code: "FROM-B" });
await tick(); await b.engine.saveNow();
a.store.createNursery({ code: "FROM-A" });
await tick(); await a.engine.saveNow();
out.stateA = a.engine.status().state;
out.fileAfterA = onDisk(file);

await a.engine.resolve("file");
out.aAfterResolve = codes(a);
out.fileAfterResolve = onDisk(file);
out.stateAfter = a.engine.status().state;
""")
    assert out["stateA"] == "changed"
    assert out["fileAfterA"] == ["FROM-B", "X"]          # A wrote nothing
    # Taking the file's copy keeps A's own new nursery and saves both.
    assert out["aAfterResolve"] == ["FROM-A", "FROM-B", "X"]
    assert out["fileAfterResolve"] == ["FROM-A", "FROM-B", "X"]
    assert out["stateAfter"] == "saved"


def test_keep_mine_overwrites_the_file():
    out = run("""
const file = fakeFile();
const a = await computer("a", { pick: () => file.handle() });
a.store.createNursery({ code: "MINE" });
await a.engine.chooseNew();
const b = await computer("b", { pick: () => file.handle() });
await b.engine.openExisting();
b.store.createNursery({ code: "THEIRS" });
await tick(); await b.engine.saveNow();
await a.engine.saveNow();
await a.engine.resolve("mine");
out.file = onDisk(file);
""")
    assert out["file"] == ["MINE"]


def test_picking_an_existing_file_as_new_does_not_wipe_it():
    out = run("""
const file = fakeFile();
const a = await computer("a", { pick: () => file.handle() });
a.store.createNursery({ code: "SEASON" });
await a.engine.chooseNew();

// A fresh browser picks the same file with "Save to a new file".
const b = await computer("b", { pick: () => file.handle() });
await b.engine.chooseNew();
out.state = b.engine.status().state;
out.file = onDisk(file);
""")
    assert out["state"] == "changed"
    assert out["file"] == ["SEASON"]


def test_reload_needs_a_click_before_writing_again():
    out = run("""
const file = fakeFile();
const a = await computer("a", { pick: () => file.handle() });
a.store.createNursery({ code: "A" });
await a.engine.chooseNew();
const writes = file.writes;

// Same computer after a reload: the handle comes back without permission.
const saved = { ...a.meta(), handle: file.handle("prompt") };
const again = await computer("a2", { saved });
again.store.createNursery({ code: "B" });
await again.engine.start();
out.state = again.engine.status().state;
out.wroteWithoutPermission = file.writes !== writes;
await again.engine.reconnect();
out.after = again.engine.status().state;
out.file = onDisk(file);
""")
    assert out["state"] == "reconnect"
    assert out["wroteWithoutPermission"] is False
    assert out["after"] == "saved"
    assert out["file"] == ["B"]


def test_declining_replacement_leaves_everything_alone():
    out = run("""
const file = fakeFile();
const a = await computer("a", { pick: () => file.handle(),
                                confirmReplace: async () => false });
a.store.createNursery({ code: "SAME" });
await a.engine.chooseNew();
await a.engine.stop();
a.store.activeNursery().seedQty = "edited here";
a.store.save();
const before = file.text;
out.opened = await a.engine.openExisting();
out.local = a.store.activeNursery().seedQty;
out.fileUnchanged = file.text === before;
out.state = a.engine.status().state;
""")
    assert out["opened"] is False
    assert out["local"] == "edited here"
    assert out["fileUnchanged"] is True
    assert out["state"] == "off"


def test_rejects_files_that_are_not_ours():
    out = run("""
for (const [k, text] of Object.entries({ junk: "not json", other: '{"a":1}',
    newer: '{"app":"ps-nursery-workbook","schema":99,"nurseries":{}}' })) {
  try { fs.parseFile(text); out[k] = "accepted"; }
  catch (e) { out[k] = e instanceof fs.NotOurFileError ? "rejected" : e.message; }
}
out.empty = fs.parseFile("  ");
""")
    assert out == {"junk": "rejected", "other": "rejected", "newer": "rejected",
                   "empty": None}
