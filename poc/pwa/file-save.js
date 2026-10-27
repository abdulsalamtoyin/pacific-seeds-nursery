// Keeping a copy of the nurseries in a file the user picks — the no-Azure
// alternative to OneDrive sync.
//
// The user picks (or creates) a file inside their OneDrive folder on this
// computer; the OneDrive desktop app then uploads it like any other file. The
// browser copy stays the working copy, and every change is written out to the
// file a moment later. Chrome and Edge only: other browsers have no way to
// write to a chosen file, so the app offers download/restore instead.
//
// The danger is two computers sharing one file through OneDrive: each would
// overwrite the other's work. So every save stamps the file with a version,
// and before writing we check the file still carries the version we last
// wrote. If not, someone else saved it, and nothing is written until the user
// picks whose copy to keep.

const APP = "ps-nursery-workbook";
const SCHEMA = 1;
const DEBOUNCE_MS = 1000;

/** The file is not one this app wrote. */
export class NotOurFileError extends Error {}

const newVersion = () =>
  `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

/** The whole store as file text, stamped with `version`. */
export function serialise(nurseries, version, now = new Date()) {
  return JSON.stringify({
    app: APP, schema: SCHEMA, version, savedAt: now.toISOString(), nurseries,
  }, null, 1);
}

/** File text back to `{version, savedAt, nurseries}`; null for an empty file. */
export function parseFile(text) {
  if (!text.trim()) return null;
  let doc;
  try {
    doc = JSON.parse(text);
  } catch {
    throw new NotOurFileError("This file is not a Nursery Workbook save file.");
  }
  if (!doc || doc.app !== APP || !doc.nurseries || typeof doc.nurseries !== "object") {
    throw new NotOurFileError("This file is not a Nursery Workbook save file.");
  }
  if (doc.schema > SCHEMA) {
    throw new NotOurFileError(
      "This file was saved by a newer version of the app. Use that version to open it.");
  }
  return doc;
}

/**
 * Nurseries in `doc` that would replace a different copy already here —
 * what the user is asked about before a file is loaded.
 */
export function overwrittenCodes(store, doc) {
  const here = store.snapshot();
  return Object.entries(doc.nurseries)
    .filter(([id, n]) => here[id] && JSON.stringify(here[id]) !== JSON.stringify(n))
    .map(([, n]) => n.code || "untitled");
}

/** Put every nursery in `doc` into the store; ones only here are kept. */
export function loadInto(store, doc) {
  for (const [id, nursery] of Object.entries(doc.nurseries)) {
    store.putNursery(id, nursery);
  }
}

// ---------------------------------------------------------------- browser
//
// The file handle survives a reload only through IndexedDB. Chrome still asks
// again for permission to write after a reload, and only from a click — the
// "reconnect" state below.

export const fileSaveSupported = () =>
  typeof globalThis.showSaveFilePicker === "function";

const PICKER_TYPES = [{
  description: "Nursery Workbook save file",
  accept: { "application/json": [".json"] },
}];

export const browserPickers = {
  pickNew: () => globalThis.showSaveFilePicker({
    suggestedName: "Nursery Workbook nurseries.json", types: PICKER_TYPES,
  }),
  pickExisting: async () =>
    (await globalThis.showOpenFilePicker({ types: PICKER_TYPES }))[0],
};

/** `{load, remember, forget}` over one IndexedDB record. */
export function idbMeta(dbName = "ps-file-save") {
  const open = () => new Promise((resolve, reject) => {
    const req = indexedDB.open(dbName, 1);
    req.onupgradeneeded = () => req.result.createObjectStore("meta");
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  const run = async (mode, fn) => {
    const db = await open();
    try {
      return await new Promise((resolve, reject) => {
        const tx = db.transaction("meta", mode);
        const req = fn(tx.objectStore("meta"));
        tx.oncomplete = () => resolve(req.result);
        tx.onerror = () => reject(tx.error);
      });
    } finally {
      db.close();
    }
  };
  return {
    load: () => run("readonly", (s) => s.get("current")),
    remember: (meta) => run("readwrite", (s) => s.put(meta, "current")),
    forget: () => run("readwrite", (s) => s.delete("current")),
  };
}

// ---------------------------------------------------------------- engine

/**
 * @param store    pwa/store.js
 * @param meta     `{load, remember, forget}` — where the handle and the last
 *                 version written are kept (idbMeta() in the browser)
 * @param pickers  `{pickNew, pickExisting}` returning file handles
 * @param onEvent  called with the status after every change of it:
 *                 {state, fileName, lastSaved, error, changedLocal}
 *                 state: off | reconnect | saving | saved | changed | error
 * @param confirmReplace  asked with nursery codes before a file's copies
 *                 replace different ones here; resolve false to cancel
 */
export function createFileSave({ store, meta, pickers, onEvent = () => {},
  confirmReplace = async () => true, debounceMs = DEBOUNCE_MS }) {
  let handle = null;
  let lastVersion = null;
  let status = { state: "off" };
  let timer = null;
  let queue = Promise.resolve();
  let unsubscribe = null;

  const emit = (patch) => {
    status = { state: status.state, fileName: handle?.name, lastSaved: status.lastSaved,
      ...patch };
    onEvent(status);
  };

  // One file operation at a time, so a debounced save never interleaves with
  // a load and writes half-old content.
  const serial = (fn) => {
    const next = queue.then(fn, fn);
    queue = next.catch(() => {});
    return next;
  };

  const readDoc = async (h = handle) => parseFile(await (await h.getFile()).text());

  async function write() {
    if (!handle) return;
    try {
      const onDisk = await readDoc();
      if (onDisk && onDisk.version !== lastVersion) {
        emit({ state: "changed", fileSavedAt: onDisk.savedAt });
        return;
      }
      emit({ state: "saving" });
      const version = newVersion();
      const out = await handle.createWritable();
      await out.write(serialise(store.snapshot(), version));
      await out.close();
      lastVersion = version;
      await meta.remember({ handle, lastVersion });
      emit({ state: "saved", lastSaved: new Date().toISOString() });
    } catch (err) {
      emit({ state: "error", error: err.message });
    }
  }

  const saveNow = () => {
    clearTimeout(timer);
    timer = null;
    return serial(write);
  };

  function onStoreChange() {
    // Only an active file is written to; "changed" waits for the user.
    if (!handle || !["saved", "saving", "error"].includes(status.state)) return;
    clearTimeout(timer);
    timer = setTimeout(saveNow, debounceMs);
  }

  async function writable(h, ask) {
    const opts = { mode: "readwrite" };
    if (await h.queryPermission(opts) === "granted") return true;
    return ask && await h.requestPermission(opts) === "granted";
  }

  /** Restore the file from the last session. Call once at boot. */
  async function start() {
    unsubscribe ??= store.subscribe(onStoreChange);
    const saved = await meta.load();
    if (!saved?.handle) { emit({ state: "off" }); return; }
    handle = saved.handle;
    lastVersion = saved.lastVersion ?? null;
    if (await writable(handle, false)) await saveNow();
    else emit({ state: "reconnect" });
  }

  /** Ask again for write permission — call from a click. */
  async function reconnect() {
    if (!handle) return;
    if (await writable(handle, true)) await saveNow();
    else emit({ state: "reconnect" });
  }

  /** Save to a new file — call from a click. */
  async function chooseNew() {
    const h = await pickers.pickNew();
    handle = h;
    lastVersion = null;
    await meta.remember({ handle, lastVersion });
    // A file picked to be replaced may hold another computer's nurseries:
    // write() sees a version it never wrote and stops at "changed".
    await saveNow();
  }

  /**
   * Load nurseries from an existing file and keep saving to it — call from a
   * click. Resolves false if the user cancelled.
   */
  async function openExisting() {
    const h = await pickers.pickExisting();
    if (!await writable(h, true)) {
      throw new Error("The browser was not allowed to save to that file.");
    }
    return serial(async () => {
      const doc = await readDoc(h);
      if (doc) {
        const clash = overwrittenCodes(store, doc);
        if (clash.length && !await confirmReplace(clash)) return false;
      }
      handle = h;
      lastVersion = doc?.version ?? null;
      await meta.remember({ handle, lastVersion });
      if (doc) {
        loadInto(store, doc);
        emit({ changedLocal: true });
      }
      await write();   // the file also gets nurseries only this browser had
      return true;
    });
  }

  /**
   * After "changed": "file" loads the file's nurseries over these, "mine"
   * overwrites the file with this browser's.
   */
  function resolve(choice) {
    return serial(async () => {
      const doc = await readDoc();
      if (choice === "file" && doc) {
        loadInto(store, doc);
        emit({ changedLocal: true });
      }
      lastVersion = doc?.version ?? null;
      await write();
    });
  }

  /** Stop saving to the file. The file itself is left alone. */
  async function stop() {
    clearTimeout(timer);
    timer = null;
    handle = null;
    lastVersion = null;
    await meta.forget();
    emit({ state: "off" });
  }

  return {
    start, reconnect, chooseNew, openExisting, resolve, stop, saveNow,
    status: () => status,
  };
}
