"""OneDrive sync must never lose a nursery.

Each case runs the real store, sync engine and Graph adapter under node
against tests/fake_graph.mjs, which enforces Graph's eTag rules. "Two
computers" are two separate instances of store.js sharing one fake OneDrive,
which is exactly the situation sync exists for: the same person on the field
laptop and the office PC.
"""
from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
PWA = ROOT / "pwa"

PREAMBLE = """
import {{ pathToFileURL }} from "node:url";
const url = (p) => pathToFileURL(p).href;
const {{ FakeGraph }} = await import(url({fake!r}));
const {{ createDrive }} = await import(url({onedrive!r}));
const {{ createSync, fileName }} = await import(url({sync!r}));

const graph = new FakeGraph({opts});
const mem = () => {{
  const m = new Map();
  return {{ getItem: (k) => (m.has(k) ? m.get(k) : null),
           setItem: (k, v) => m.set(k, String(v)), removeItem: (k) => m.delete(k) }};
}};

// A computer: its own store instance (the ?tag makes node load a fresh copy),
// its own sync records, and the shared fake OneDrive.
async function computer(tag, {{ token = "tok-1", refreshed = "tok-1", confirmAdopt }} = {{}}) {{
  const store = await import(url({store!r}) + "?" + tag);
  store.useBackend(mem());
  const waits = [];
  const drive = createDrive({{
    getToken: async ({{ force }} = {{}}) => (force ? refreshed : token),
    fetchImpl: (u, init) => graph.fetch(u, init),   // late-bound: tests swap it
    wait: async (ms) => {{ waits.push(ms); }},
  }});
  const states = [];
  const sync = createSync({{ store, drive, metaBackend: mem(), confirmAdopt,
                            onEvent: (e) => states.push(e.state) }});
  return {{ store, sync, drive, waits, states }};
}}
const codes = (c) => Object.values(c.store.snapshot()).map((n) => n.code).sort();
const out = {{}};
"""


def run(body: str, opts: dict | None = None) -> dict:
    if shutil.which("node") is None:
        pytest.skip("node not installed — cannot exercise the sync engine")
    script = PREAMBLE.format(
        fake=str(ROOT / "tests" / "fake_graph.mjs"),
        onedrive=str(PWA / "onedrive.js"),
        sync=str(PWA / "sync.js"),
        store=str(PWA / "store.js"),
        opts=json.dumps(opts or {}),
    ) + body + "\nout.violations = graph.violations;\nconsole.log(JSON.stringify(out));\n"
    proc = subprocess.run(["node", "--input-type=module", "-e", script],
                          capture_output=True, text=True, timeout=60)
    if proc.returncode != 0:
        pytest.fail(f"node driver failed:\n{proc.stderr}")
    result = json.loads(proc.stdout.strip().splitlines()[-1])
    assert result["violations"] == [], result["violations"]
    return result


# ------------------------------------------------------------- first sync


def test_first_sync_uploads_every_nursery():
    out = run("""
      const a = await computer("a");
      a.store.createNursery({ code: "AUQDGT01-27S", types: ["AB"] });
      a.store.createNursery({ code: "AUGT1-26S" });
      out.result = await a.sync.sync();
      out.codes = [...graph.files.values()].map((f) => JSON.parse(f.content).nursery.code).sort();
    """)
    assert out["result"]["uploaded"] == 2
    assert out["codes"] == ["AUGT1-26S", "AUQDGT01-27S"]


def test_an_untouched_blank_nursery_is_not_uploaded():
    """The app makes an empty "untitled" nursery on first open; every computer
    would otherwise add one more to OneDrive."""
    out = run("""
      const a = await computer("a");
      a.store.activeNursery();
      out.result = await a.sync.sync();
      out.files = graph.files.size;
    """)
    assert out["files"] == 0



# ------------------------------------------------------ shared computers


def test_nurseries_already_here_are_not_uploaded_if_declined():
    """Someone else may have used this browser; their nurseries must not be
    copied into the account that signs in next."""
    out = run("""
      const asked = [];
      const b = await computer("b", { confirmAdopt: async (codes) => { asked.push(codes); return false; } });
      b.store.createNursery({ code: "PREVIOUS-USER" });
      await b.sync.sync();
      out.filesAfterDecline = graph.files.size;
      b.store.createNursery({ code: "MINE-27S" });
      await b.sync.sync();
      out.uploaded = [...graph.files.values()].map((f) => JSON.parse(f.content).nursery.code);
      out.asked = asked;
      out.removable = b.sync.safeToRemove().map((id) => b.store.snapshot()[id].code);
    """)
    assert out["asked"] == [["PREVIOUS-USER"]]          # once, with the codes
    assert out["filesAfterDecline"] == 0
    assert out["uploaded"] == ["MINE-27S"]
    assert out["removable"] == ["MINE-27S"]             # the declined one is never wiped


def test_nurseries_already_here_are_uploaded_if_accepted():
    out = run("""
      const b = await computer("b", { confirmAdopt: async () => true });
      b.store.createNursery({ code: "AUQDGT01-27S" });
      out.result = await b.sync.sync();
    """)
    assert out["result"]["uploaded"] == 1


def test_unsynced_edits_are_not_safe_to_remove():
    out = run("""
      const a = await computer("a");
      const id = a.store.createNursery({ code: "AUQDGT01-27S" });
      await a.sync.sync();
      out.before = a.sync.safeToRemove().length;
      a.store.renameNursery(id, { code: "EDITED" });
      out.after = a.sync.safeToRemove().length;
    """)
    assert out["before"] == 1
    assert out["after"] == 0

def test_a_second_computer_receives_the_nurseries():
    out = run("""
      const a = await computer("a");
      a.store.createNursery({ code: "AUQDGT01-27S" });
      await a.sync.sync();
      const b = await computer("b");
      out.result = await b.sync.sync();
      out.codes = codes(b);
    """)
    assert out["result"]["downloaded"] == 1
    assert out["codes"] == ["AUQDGT01-27S"]


# ------------------------------------------------------------- editing


def test_an_edit_on_one_computer_reaches_the_other():
    out = run("""
      const a = await computer("a"), b = await computer("b");
      const id = a.store.createNursery({ code: "N1" });
      await a.sync.sync(); await b.sync.sync();
      a.store.activeNursery().seedQty = "40"; a.store.save();
      await a.sync.sync();
      out.result = await b.sync.sync();
      out.seed = b.store.snapshot()[id].seedQty;
    """)
    assert out["seed"] == "40"


def test_nothing_is_sent_when_nothing_changed():
    out = run("""
      const a = await computer("a");
      a.store.createNursery({ code: "N1" });
      await a.sync.sync();
      graph.log.length = 0;
      out.result = await a.sync.sync();
      out.writes = graph.log.filter((l) => !l.startsWith("GET")).length;
    """)
    assert out["writes"] == 0


def test_typing_during_an_upload_is_not_marked_as_synced():
    """The upload carries the content from before the keystroke; the
    keystroke must still go up on the next sync."""
    out = run("""
      const a = await computer("a");
      const id = a.store.createNursery({ code: "N1" });
      const realFetch = graph.fetch;
      let typed = false;
      graph.fetch = async (u, init) => {
        if (init?.method === "PUT" && !typed) {
          typed = true;
          a.store.activeNursery().seedQty = "typed mid-upload"; a.store.save();
        }
        return realFetch(u, init);
      };
      await a.sync.sync();
      out.before = graph.read(fileName(id)).nursery.seedQty;
      await a.sync.sync();
      out.after = graph.read(fileName(id)).nursery.seedQty;
    """)
    assert out["before"] == ""
    assert out["after"] == "typed mid-upload"


# ------------------------------------------------------------- conflicts


CONFLICT = """
  const a = await computer("a"), b = await computer("b");
  const id = a.store.createNursery({ code: "N1" });
  await a.sync.sync(); await b.sync.sync();
  b.store.switchNursery(id);
  a.store.activeNursery().seedQty = "from A"; a.store.save();
  b.store.activeNursery().seedQty = "from B"; b.store.save();
  await a.sync.sync();
  out.result = await b.sync.sync();
  out.conflicts = b.sync.status().conflicts.map((c) => c.id);
  out.id = id;
"""


def test_edits_on_both_computers_become_a_conflict_not_an_overwrite():
    out = run(CONFLICT + """
      out.remote = graph.read(fileName(id)).nursery.seedQty;
      out.local = b.store.snapshot()[id].seedQty;
      out.state = b.sync.status().state;
    """)
    assert out["conflicts"] == [out["id"]]
    assert out["remote"] == "from A"      # A's save was not overwritten
    assert out["local"] == "from B"       # nor was B's typing
    assert out["state"] == "conflict"


def test_conflict_keep_mine():
    out = run(CONFLICT + """
      await b.sync.resolve(id, "mine");
      await a.sync.sync();
      out.remote = graph.read(fileName(id)).nursery.seedQty;
      out.a = a.store.snapshot()[id].seedQty;
    """)
    assert out["remote"] == "from B"
    assert out["a"] == "from B"


def test_conflict_keep_theirs():
    out = run(CONFLICT + """
      await b.sync.resolve(id, "theirs");
      out.b = b.store.snapshot()[id].seedQty;
      out.left = b.sync.status().conflicts.length;
    """)
    assert out["b"] == "from A"
    assert out["left"] == 0


def test_conflict_keep_both_saves_this_computers_copy_separately():
    out = run(CONFLICT + """
      await b.sync.resolve(id, "both");
      await a.sync.sync();
      out.b = Object.values(b.store.snapshot()).map((n) => [n.code, n.seedQty]);
      out.a = Object.values(a.store.snapshot()).map((n) => [n.code, n.seedQty]);
    """)
    expected = [["N1", "from A"], ["N1 (this computer)", "from B"]]
    assert sorted(out["b"]) == expected
    assert sorted(out["a"]) == expected


# ------------------------------------------------------------- deleting


def test_a_deletion_reaches_the_other_computer():
    out = run("""
      const a = await computer("a"), b = await computer("b");
      const id = a.store.createNursery({ code: "N1" });
      a.store.createNursery({ code: "N2" });
      await a.sync.sync(); await b.sync.sync();
      a.store.deleteNursery(id);
      await a.sync.sync();
      await b.sync.sync();
      out.b = codes(b);
      out.files = graph.files.size;
    """)
    assert out["b"] == ["N2"]
    assert out["files"] == 1


def test_edits_win_over_a_deletion_elsewhere():
    """Deleting on one computer must not wipe what was typed on the other."""
    out = run("""
      const a = await computer("a"), b = await computer("b");
      const id = a.store.createNursery({ code: "N1" });
      await a.sync.sync(); await b.sync.sync();
      b.store.switchNursery(id);
      a.store.deleteNursery(id);
      await a.sync.sync();
      b.store.activeNursery().seedQty = "kept"; b.store.save();
      await b.sync.sync();
      await a.sync.sync();
      out.remote = graph.read(fileName(id))?.nursery.seedQty;
      out.a = a.store.snapshot()[id]?.seedQty;
    """)
    assert out["remote"] == "kept"
    assert out["a"] == "kept"


# ------------------------------------------------------------- the network


def test_offline_keeps_working_and_catches_up():
    out = run("""
      const a = await computer("a");
      a.store.createNursery({ code: "N1" });
      graph.offline = true;
      await a.sync.sync();
      out.state = a.sync.status().state;
      out.local = codes(a);
      graph.offline = false;
      out.result = await a.sync.sync();
      out.files = graph.files.size;
    """)
    assert out["state"] == "offline"
    assert out["local"] == ["N1"]
    assert out["files"] == 1


def test_an_expired_token_is_renewed_once():
    out = run("""
      graph.validTokens = new Set(["tok-2"]);
      const a = await computer("a", { token: "tok-old", refreshed: "tok-2" });
      a.store.createNursery({ code: "N1" });
      out.result = await a.sync.sync();
      out.state = a.sync.status().state;
    """)
    assert out["state"] == "idle"
    assert out["result"]["uploaded"] == 1


def test_throttling_waits_as_told_then_retries():
    out = run("""
      const a = await computer("a");
      graph.queued.push(() => new Response("{}", { status: 429,
        headers: { "Retry-After": "3" } }));
      a.store.createNursery({ code: "N1" });
      await a.sync.sync();
      out.waits = a.waits;
      out.files = graph.files.size;
    """)
    assert out["waits"] == [3000]
    assert out["files"] == 1


def test_long_listings_are_followed_across_pages():
    out = run("""
      const a = await computer("a");
      for (let i = 0; i < 5; i++) a.store.createNursery({ code: "N" + i });
      await a.sync.sync();
      const b = await computer("b");
      await b.sync.sync();
      out.b = codes(b).length;
    """, opts={"pageSize": 2})
    assert out["b"] == 5


def test_a_broken_file_is_skipped_not_fatal():
    out = run("""
      graph.seed("nursery-bad.json", "{\\"not\\": \\"ours\\"}");
      graph.seed("notes.json", { hello: 1 });
      const a = await computer("a");
      a.store.createNursery({ code: "N1" });
      out.result = await a.sync.sync();
      out.state = a.sync.status().state;
      out.error = a.sync.status().error;
    """)
    assert out["result"]["uploaded"] == 1
    assert out["state"] == "idle"
    assert "nursery-bad.json" in out["error"]


def test_removing_from_this_computer_never_deletes_from_onedrive():
    """Sign-out's "remove from this computer": signing in again later must
    bring the nurseries back, not read their absence as a delete."""
    out = run("""
      const a = await computer("a");
      a.store.createNursery({ code: "AUQDGT01-27S" });
      await a.sync.sync();
      a.sync.removeLocal(a.sync.safeToRemove());
      out.localAfterRemove = codes(a);
      out.result = await a.sync.sync();            // the next sign-in
      out.onedrive = graph.files.size;
      out.localAfterSync = codes(a);
    """)
    assert out["localAfterRemove"] == []
    assert out["onedrive"] == 1
    assert out["result"]["deleted"] == 0
    assert out["localAfterSync"] == ["AUQDGT01-27S"]
