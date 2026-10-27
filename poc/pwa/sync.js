// Keeping the browser's nurseries and the user's OneDrive in step.
//
// The browser copy stays the working copy — the app reads and writes it as
// before, and keeps working offline. This copies changes up and down in the
// background, one JSON file per nursery.
//
// For each nursery it remembers the OneDrive version (eTag) it last agreed on
// and a fingerprint of the local content at that moment. Comparing those with
// what is there now gives the whole decision:
//
//   changed here only        upload, guarded by the old eTag
//   changed on OneDrive only take the OneDrive copy
//   changed in both places   a conflict — the user picks; nothing is dropped
//   deleted on one side      delete on the other, unless the other side has
//                            edits since, in which case the edits win
//
// The rule throughout is that no one's typing is ever discarded silently.

import { ConflictError, OfflineError } from "./onedrive.js";

const PREFIX = "nursery-";
const SCHEMA = 1;

/** A file in the app folder that is not a nursery this app wrote. */
class BadFileError extends Error {}

export const fileName = (id) => `${PREFIX}${id}.json`;
const idFromName = (name) =>
  name.startsWith(PREFIX) && name.endsWith(".json")
    ? name.slice(PREFIX.length, -".json".length) : null;

/** cyrb53 — a fast 53-bit string hash; enough to tell "changed" apart. */
export function fingerprint(nursery) {
  const s = JSON.stringify(nursery);
  let h1 = 0xdeadbeef;
  let h2 = 0x41c6ce57;
  for (let i = 0; i < s.length; i++) {
    const ch = s.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}

/**
 * @param store        pwa/store.js (or a second instance of it, in tests)
 * @param drive        pwa/onedrive.js createDrive() result
 * @param metaBackend  localStorage-like; holds what was last agreed
 * @param metaKey      per signed-in account, so two people sharing a browser
 *                     never compare against each other's OneDrive
 * @param onEvent      called with the status after every change of it
 */
export function createSync({ store, drive, metaBackend, metaKey = "ps-sync-meta",
  onEvent = () => {}, now = () => new Date() }) {
  const blankPrint = fingerprint(structuredClone(store.BLANK_NURSERY));
  const conflicts = new Map();
  let meta = loadMeta();
  let status = { state: "idle", lastSync: meta.lastSync, error: null };
  let running = null;
  let again = false;

  function loadMeta() {
    try {
      const parsed = JSON.parse(metaBackend.getItem(metaKey) ?? "null");
      if (parsed && parsed.files) return parsed;
    } catch { /* fall through */ }
    return { files: {}, lastSync: null };
  }
  const saveMeta = () => metaBackend.setItem(metaKey, JSON.stringify(meta));

  function emit(patch) {
    status = { ...status, ...patch };
    onEvent({ ...status, conflicts: [...conflicts.values()] });
  }

  const doc = (id, nursery) => ({
    schema: SCHEMA, id, savedAt: now().toISOString(), nursery,
  });

  /**
   * Agree on a version: record its eTag and the local content it matches.
   *
   * After an upload, `hash` must be what was *sent* — anything typed while
   * the upload was in flight is not on OneDrive yet and has to stay dirty.
   * After a download it defaults to the store's copy, read straight away.
   */
  function agree(id, eTag, hash = null) {
    if (hash === null) {
      const local = store.snapshot()[id];
      hash = local ? fingerprint(local) : null;
    }
    meta.files[id] = { eTag, hash };
    saveMeta();
  }

  async function fetchDoc(file) {
    const body = await drive.download(file);
    if (!body || body.schema !== SCHEMA || typeof body.nursery !== "object"
        || body.nursery === null) {
      throw new BadFileError(`${file.name} is not a nursery file this app wrote.`);
    }
    return body.nursery;
  }

  // One pass over every nursery known locally, remotely, or in the meta.
  // Returns whether another pass is needed (a write raced another device).
  async function pass(report) {
    const remote = new Map();
    for (const file of await drive.list()) {
      const id = idFromName(file.name);
      if (id) remote.set(id, file);
    }
    const local = store.snapshot();
    const ids = new Set([...Object.keys(local), ...remote.keys(),
      ...Object.keys(meta.files)]);
    let raced = false;

    for (const id of ids) {
      if (conflicts.has(id)) continue;          // waiting on the user
      const L = local[id];
      const R = remote.get(id);
      const M = meta.files[id];
      const lHash = L ? fingerprint(L) : null;
      const dirty = Boolean(L) && (!M || lHash !== M.hash);

      try {
        if (L && !R) {
          if (M && !dirty) {
            // Deleted on another device and untouched here: follow suit.
            store.deleteNursery(id);
            delete meta.files[id];
            saveMeta();
            report.changedLocal = true;
          } else if (M || lHash !== blankPrint) {
            // New here, or edited here after another device deleted it —
            // the edits win. An untouched blank nursery is not worth a file.
            const { eTag } = await drive.upload(fileName(id), doc(id, L));
            agree(id, eTag, lHash);
            report.uploaded++;
          }
        } else if (!L && R) {
          if (M && M.eTag === R.eTag) {
            // Deleted here; OneDrive still has the version we agreed on.
            await drive.remove(R.name, R.eTag);
            delete meta.files[id];
            saveMeta();
            report.deleted++;
          } else {
            // New from another device — or changed there after we deleted
            // it, in which case their edits win.
            store.putNursery(id, await fetchDoc(R));
            agree(id, R.eTag);
            report.downloaded++;
            report.changedLocal = true;
          }
        } else if (L && R) {
          if (M && M.eTag === R.eTag) {
            if (dirty) {
              const { eTag } = await drive.upload(fileName(id), doc(id, L), M.eTag);
              agree(id, eTag, lHash);
              report.uploaded++;
            }
          } else {
            const theirs = await fetchDoc(R);
            if (fingerprint(theirs) === lHash) {
              agree(id, R.eTag);                // same content both sides
            } else if (!dirty) {
              store.putNursery(id, theirs);
              agree(id, R.eTag);
              report.downloaded++;
              report.changedLocal = true;
            } else {
              conflicts.set(id, {
                id, eTag: R.eTag, local: L, remote: theirs,
                code: L.code || theirs.code || "(untitled)",
              });
            }
          }
        } else if (M) {
          delete meta.files[id];                // gone everywhere
          saveMeta();
        }
      } catch (err) {
        if (err instanceof ConflictError) raced = true;
        // One unreadable file must not stop every other nursery syncing.
        else if (err instanceof BadFileError) report.skipped.push(err.message);
        else throw err;
      }
    }
    return raced;
  }

  async function run() {
    emit({ state: "syncing", error: null, changedLocal: false });
    const report = {
      uploaded: 0, downloaded: 0, deleted: 0, skipped: [], changedLocal: false,
    };
    try {
      // A race means OneDrive moved under us mid-pass; the next pass sees it
      // as a remote change and handles it properly. Bounded, not a loop.
      for (let i = 0; i < 3 && await pass(report); i++) { /* re-check */ }
      meta.lastSync = now().toISOString();
      saveMeta();
      emit({
        state: conflicts.size ? "conflict" : "idle",
        lastSync: meta.lastSync,
        changedLocal: report.changedLocal,
        error: report.skipped.length ? report.skipped.join(" ") : null,
      });
    } catch (err) {
      emit({
        state: err instanceof OfflineError ? "offline"
          : err?.name === "NeedsSignInError" ? "signin" : "error",
        error: err.message,
        changedLocal: report.changedLocal,
      });
    }
    return { ...report, conflicts: conflicts.size };
  }

  /** Sync now. Calls made while one is running fold into one follow-up. */
  function sync() {
    if (running) { again = true; return running; }
    running = (async () => {
      let result;
      do { again = false; result = await run(); } while (again);
      running = null;
      return result;
    })();
    return running;
  }

  /**
   * Settle a conflict.
   *   "mine"   — this computer's version replaces OneDrive's
   *   "theirs" — OneDrive's version replaces this computer's
   *   "both"   — keep OneDrive's under the original, and this computer's as
   *              a separate nursery marked "(this computer)"
   */
  async function resolve(id, choice) {
    const c = conflicts.get(id);
    if (!c) throw new Error(`no conflict for ${id}`);

    if (choice === "mine") {
      const current = store.snapshot()[id] ?? c.local;
      try {
        const { eTag } = await drive.upload(fileName(id), doc(id, current), c.eTag);
        agree(id, eTag, fingerprint(current));
      } catch (err) {
        if (!(err instanceof ConflictError)) throw err;
        // Changed yet again elsewhere; re-detect on the next pass.
      }
    } else if (choice === "theirs" || choice === "both") {
      if (choice === "both") {
        const copyId = `${id}-${now().getTime().toString(36)}`;
        store.putNursery(copyId, {
          ...(store.snapshot()[id] ?? c.local),
          code: `${c.local.code || "(untitled)"} (this computer)`,
        });
      }
      store.putNursery(id, c.remote);
      agree(id, c.eTag);
    } else {
      throw new Error(`unknown choice: ${choice}`);
    }
    conflicts.delete(id);
    emit({ changedLocal: choice !== "mine" });
    return sync();
  }

  /** Sync after edits settle, when back online, and every few minutes. */
  function start({ debounceMs = 2000, intervalMs = 5 * 60_000 } = {}) {
    let timer = null;
    const soon = () => { clearTimeout(timer); timer = setTimeout(sync, debounceMs); };
    const unsubscribe = store.subscribe(soon);
    const tick = setInterval(sync, intervalMs);
    const win = typeof window !== "undefined" ? window : null;
    win?.addEventListener("online", sync);
    sync();
    return () => {
      unsubscribe();
      clearTimeout(timer);
      clearInterval(tick);
      win?.removeEventListener("online", sync);
    };
  }

  return {
    sync, resolve, start,
    status: () => ({ ...status, conflicts: [...conflicts.values()] }),
  };
}
