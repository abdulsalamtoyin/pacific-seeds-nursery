"""A team sharing nurseries through a synced folder must never lose one.

Each case runs the real store, sync engine and folder drive under node. The
team folder is an in-memory stand-in for the browser's directory handle;
"two computers" are two store instances pointed at it, which is what the
OneDrive desktop app gives two people syncing the same SharePoint folder
(minus the delay, which only makes the conflicts below rarer).
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
const {{ createFolderDrive }} = await import(url({drive!r}));
const {{ createSync, fileName }} = await import(url({sync!r}));
const backup = await import(url({backup!r}));

const mem = () => {{
  const m = new Map();
  return {{ getItem: (k) => (m.has(k) ? m.get(k) : null),
           setItem: (k, v) => m.set(k, String(v)), removeItem: (k) => m.delete(k) }};
}};

// The synced folder: name -> {{text, lastModified}}. `broken` makes a file
// unreadable, as an online-only OneDrive file is while offline.
const files = new Map();
const broken = new Set();
let clock = 1;
const notFound = () => Object.assign(new Error("not found"), {{ name: "NotFoundError" }});
const fileHandle = (name) => ({{
  kind: "file", name,
  getFile: async () => {{
    if (broken.has(name)) throw new Error("cloud file not available");
    const f = files.get(name);
    if (!f) throw notFound();
    return {{ lastModified: f.lastModified, size: f.text.length, text: async () => f.text }};
  }},
  createWritable: async () => {{
    let buf = "";
    return {{ write: async (t) => {{ buf += t; }},
             close: async () => {{ files.set(name, {{ text: buf, lastModified: clock++ }}); }} }};
  }},
}});
const dir = {{
  name: "Nursery team",
  async getFileHandle(name, {{ create = false }} = {{}}) {{
    if (!files.has(name)) {{
      if (!create) throw notFound();
      files.set(name, {{ text: "", lastModified: clock++ }});
    }}
    return fileHandle(name);
  }},
  async removeEntry(name) {{ if (!files.delete(name)) throw notFound(); }},
  async *entries() {{ for (const name of [...files.keys()]) yield [name, fileHandle(name)]; }},
}};

async function computer(tag, author) {{
  const store = await import(url({store!r}) + "?" + tag);
  store.useBackend(mem());
  const sync = createSync({{ store, drive: createFolderDrive(dir), metaBackend: mem(), author }});
  return {{ store, sync }};
}}
const codes = (c) => Object.values(c.store.snapshot()).map((n) => n.code).sort();
const onDisk = (id) => JSON.parse(files.get(fileName(id)).text);
const out = {{}};
"""


def run(body: str) -> dict:
    if shutil.which("node") is None:
        pytest.skip("node not installed — cannot exercise the team folder")
    script = PREAMBLE.format(drive=str(PWA / "folder-drive.js"),
                             sync=str(PWA / "sync.js"),
                             backup=str(PWA / "backup.js"),
                             store=str(PWA / "store.js")) + body + \
        "\nconsole.log(JSON.stringify(out));\n"
    proc = subprocess.run(["node", "--input-type=module", "-e", script],
                          capture_output=True, text=True, timeout=30)
    if proc.returncode != 0:
        pytest.fail(f"node driver failed:\n{proc.stderr}")
    return json.loads(proc.stdout)


def test_a_nursery_saved_by_one_person_reaches_the_other():
    out = run("""
const a = await computer("a", "Ana");
const b = await computer("b", "Ben");
const id = a.store.createNursery({ code: "AUGT1-26S" });
await a.sync.sync();
await b.sync.sync();
out.b = codes(b);
out.author = onDisk(id).author;
out.hasSavedAt = Boolean(onDisk(id).savedAt);
""")
    assert out["b"] == ["AUGT1-26S"]
    assert out["author"] == "Ana"
    assert out["hasSavedAt"] is True


def test_edits_flow_both_ways_and_record_who_saved_last():
    out = run("""
const a = await computer("a", "Ana");
const b = await computer("b", "Ben");
const id = a.store.createNursery({ code: "N1" });
await a.sync.sync(); await b.sync.sync();
b.store.switchNursery(id);
b.store.activeNursery().operations["Seedling|1"] = "thinned";
b.store.save();
await b.sync.sync(); await a.sync.sync();
out.aSees = a.store.snapshot()[id].operations["Seedling|1"];
out.author = onDisk(id).author;
""")
    assert out["aSees"] == "thinned"
    assert out["author"] == "Ben"


def test_two_people_editing_one_nursery_get_a_choice_not_a_loss():
    out = run("""
const a = await computer("a", "Ana");
const b = await computer("b", "Ben");
const id = a.store.createNursery({ code: "N1" });
await a.sync.sync(); await b.sync.sync();
b.store.switchNursery(id);

a.store.activeNursery().seedQty = "from Ana";  a.store.save();
b.store.activeNursery().seedQty = "from Ben";  b.store.save();
await a.sync.sync();
await b.sync.sync();
out.bState = b.sync.status().state;
out.onDisk = onDisk(id).nursery.seedQty;
out.bLocal = b.store.snapshot()[id].seedQty;

await b.sync.resolve(id, "both");
out.bCodes = codes(b);
""")
    assert out["bState"] == "conflict"
    assert out["onDisk"] == "from Ana"       # B did not overwrite
    assert out["bLocal"] == "from Ben"       # nor lose its own edit
    assert out["bCodes"] == ["N1", "N1 (this computer)"]


def test_deleting_a_nursery_removes_it_for_the_team():
    out = run("""
const a = await computer("a", "Ana");
const b = await computer("b", "Ben");
const id = a.store.createNursery({ code: "GONE" });
a.store.createNursery({ code: "KEEP" });
await a.sync.sync(); await b.sync.sync();
a.store.deleteNursery(id);
await a.sync.sync(); await b.sync.sync();
out.b = codes(b);
out.fileLeft = files.has(fileName(id));
""")
    assert out["b"] == ["KEEP"]
    assert out["fileLeft"] is False


def test_an_unreadable_file_stops_sync_without_deleting_anything():
    out = run("""
const a = await computer("a", "Ana");
const b = await computer("b", "Ben");
const id = a.store.createNursery({ code: "OFFLINE" });
await a.sync.sync(); await b.sync.sync();
broken.add(fileName(id));
await b.sync.sync();
out.state = b.sync.status().state;
out.b = codes(b);
out.fileLeft = files.has(fileName(id));
""")
    assert out["state"] == "offline"
    assert out["b"] == ["OFFLINE"]
    assert out["fileLeft"] is True


def test_a_onedrive_conflict_copy_shows_up_rather_than_vanishing():
    out = run("""
const a = await computer("a", "Ana");
const id = a.store.createNursery({ code: "N1" });
await a.sync.sync();
// OneDrive keeps both sides of a clash as a renamed copy beside the file.
const copy = fileName(id).replace(".json", "-LAPTOP.json");
files.set(copy, { text: files.get(fileName(id)).text, lastModified: clock++ });
const b = await computer("b", "Ben");
await b.sync.sync();
out.b = codes(b);
""")
    assert out["b"] == ["N1", "N1"]


def test_hostile_file_names_in_the_shared_folder_are_ignored():
    out = run("""
// Raw text: in a JS literal "__proto__:" would set the prototype, not a key.
const body = (id) => `{"schema":1,"id":${JSON.stringify(id)},"author":"x",` +
  `"nursery":{"code":"EVIL","__proto__":{"polluted":true}}}`;
for (const id of ["__proto__", "constructor", "toString",
                  "hasOwnProperty", "valueOf", "../x"]) {
  files.set(`nursery-${id}.json`, { text: body(id), lastModified: clock++ });
}
files.set("nursery-n-ok.json", { text: body("n-ok"), lastModified: clock++ });
const b = await computer("b", "Ben");
await b.sync.sync();
out.state = b.sync.status().state;
out.codes = codes(b);
out.polluted = ({}).code !== undefined || ({}).polluted !== undefined;
out.ownProto = Object.getPrototypeOf(b.store.snapshot()["n-ok"]) === Object.prototype;
""")
    assert out["state"] == "idle"            # sync carried on, not "error"
    assert out["codes"] == ["EVIL"]          # only the well-named file
    assert out["polluted"] is False
    assert out["ownProto"] is True


def test_backup_rejects_files_that_are_not_ours():
    out = run("""
for (const [k, text] of Object.entries({ junk: "not json", other: '{"a":1}',
    newer: '{"app":"ps-nursery-workbook","schema":99,"nurseries":{}}' })) {
  try { backup.parseFile(text); out[k] = "accepted"; }
  catch (e) { out[k] = e instanceof backup.NotOurFileError ? "rejected" : e.message; }
}
out.roundTrip = Object.keys(backup.parseFile(backup.serialise({ x: { code: "X" } })).nurseries);
""")
    assert out == {"junk": "rejected", "other": "rejected", "newer": "rejected",
                   "roundTrip": ["x"]}
