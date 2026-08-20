// Nursery storage — many nurseries, one active at a time.
//
// The app used to hold exactly one nursery, so starting a new one silently
// overwrote the last. State is now a keyed collection with an active pointer,
// and a legacy payload is migrated in on first load rather than discarded.
//
// Grid state (sort, filters, colours, hand-edits, user columns) lives here too,
// per nursery per tab, so it survives a tab switch and a reload.

const LEGACY_KEY = "ps-nursery-workbook";
const STORE_KEY = "ps-nursery-workbook-v3";
const SCHEMA = 3;

// Fields every nursery carries. Cloned per nursery so two nurseries never
// share an array or map by reference.
export const BLANK_NURSERY = {
  code: "",
  fileName: "",
  types: [],
  plantingDates: 1,
  seedQty: "",
  prism: [],           // rows from Nursery site
  updatedPrism: [],    // rows from Updated nursery site
  fieldMap: [],        // {dop, row, spike, run, qty}
  nurseryData: {},     // label -> {value, comment}
  mapFields: [],       // which fields Material Map shows in each cell
  dateRecording: {},   // "<range>_<row>" -> {"S 1": "28", ...}
  recordingStart: "",  // ISO date the selection passes started from
  replacements: [],
  plantingErrors: [],
  operations: {},      // "stage|group" -> text
  comments: {},
  snapshots: [],       // Fieldbook captures
  grids: {},           // tab name -> grid state
};

export const BLANK_GRID = {
  sort: null,          // {key, dir: "asc"|"desc"}
  filters: {},         // columnKey -> [allowed values]
  colours: {},         // "rowKey|columnKey" -> css colour
  styles: {},          // "rowKey|columnKey" -> {bold, italic}
  widths: {},          // columnKey -> px
  hidden: [],          // columnKey[]
  extraColumns: [],    // {key, label, type}
  extraRows: [],       // row objects the user added
  edits: {},           // rowKey -> {columnKey: value}
};

// ---------------------------------------------------------------- backend
//
// localStorage under the browser, an in-memory stand-in under node so the
// store can be tested without a DOM.

function memoryBackend() {
  const map = new Map();
  return {
    getItem: (k) => (map.has(k) ? map.get(k) : null),
    setItem: (k, v) => map.set(k, String(v)),
    removeItem: (k) => map.delete(k),
  };
}

let backend = typeof localStorage !== "undefined"
  ? localStorage
  : memoryBackend();

let data = null;

/** Swap the storage backend. Test seam — resets the cache. */
export function useBackend(next) {
  backend = next ?? memoryBackend();
  data = null;
}

// ---------------------------------------------------------------- ids

function newId() {
  return `n-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function blankNursery(patch = {}) {
  return { ...structuredClone(BLANK_NURSERY), ...patch };
}

// ---------------------------------------------------------------- load/save

function emptyStore() {
  return { schema: SCHEMA, activeId: null, nurseries: {} };
}

// A pre-v3 payload is one nursery's fields at the top level. Lifting it in
// keeps whatever the user already had; dropping it would lose a season's work.
function migrateLegacy() {
  let legacy;
  try {
    legacy = JSON.parse(backend.getItem(LEGACY_KEY) ?? "null");
  } catch {
    return null;
  }
  if (!legacy || typeof legacy !== "object" || Array.isArray(legacy)) {
    return null;
  }

  const store = emptyStore();
  const id = newId();
  store.nurseries[id] = blankNursery({
    ...legacy,
    // The old shape had no per-tab grid state and no split error log.
    grids: {},
    plantingErrors: legacy.plantingErrors ?? [],
    nurseryData: legacy.nurseryData ?? {},
    snapshots: [],
  });
  store.activeId = id;
  return store;
}

function read() {
  if (data) return data;

  let parsed = null;
  try {
    parsed = JSON.parse(backend.getItem(STORE_KEY) ?? "null");
  } catch {
    parsed = null;
  }

  if (parsed && parsed.schema === SCHEMA && parsed.nurseries) {
    data = parsed;
  } else {
    data = migrateLegacy() ?? emptyStore();
    if (Object.keys(data.nurseries).length) persist();
  }

  // An active pointer at a deleted nursery would strand every view.
  if (!data.activeId || !data.nurseries[data.activeId]) {
    data.activeId = Object.keys(data.nurseries)[0] ?? null;
  }
  return data;
}

function persist() {
  backend.setItem(STORE_KEY, JSON.stringify(data));
}

/** Write the current state back to storage. */
export function save() {
  read();
  persist();
}

// ---------------------------------------------------------------- nurseries

/** `[{id, code, fileName, types}]`, creation order. */
export function listNurseries() {
  const store = read();
  return Object.entries(store.nurseries).map(([id, n]) => ({
    id,
    code: n.code,
    fileName: n.fileName,
    types: n.types,
  }));
}

export function activeId() {
  return read().activeId;
}

/**
 * The active nursery, creating an untitled one if the store is empty so
 * callers never have to null-check.
 */
export function activeNursery() {
  const store = read();
  if (!store.activeId) {
    const id = newId();
    store.nurseries[id] = blankNursery();
    store.activeId = id;
    persist();
  }
  return store.nurseries[store.activeId];
}

export function createNursery({ code = "", types = [], fileName = "" } = {}) {
  const store = read();
  const id = newId();
  store.nurseries[id] = blankNursery({ code, types, fileName });
  store.activeId = id;
  persist();
  return id;
}

export function switchNursery(id) {
  const store = read();
  if (!store.nurseries[id]) throw new Error(`no such nursery: ${id}`);
  store.activeId = id;
  persist();
}

export function renameNursery(id, { code, fileName } = {}) {
  const store = read();
  const nursery = store.nurseries[id];
  if (!nursery) throw new Error(`no such nursery: ${id}`);
  if (code !== undefined) nursery.code = code;
  if (fileName !== undefined) nursery.fileName = fileName;
  persist();
}

/** Deep copy, so editing the copy cannot reach back into the original. */
export function duplicateNursery(id) {
  const store = read();
  const source = store.nurseries[id];
  if (!source) throw new Error(`no such nursery: ${id}`);
  const copy = structuredClone(source);
  copy.code = `${source.code} (copy)`;
  const copyId = newId();
  store.nurseries[copyId] = copy;
  store.activeId = copyId;
  persist();
  return copyId;
}

export function deleteNursery(id) {
  const store = read();
  if (!store.nurseries[id]) throw new Error(`no such nursery: ${id}`);
  delete store.nurseries[id];
  if (store.activeId === id) {
    store.activeId = Object.keys(store.nurseries)[0] ?? null;
  }
  persist();
}

// ---------------------------------------------------------------- grid state

/** Per-nursery, per-tab grid state, created on first use. */
export function gridState(tabName) {
  const nursery = activeNursery();
  if (!nursery.grids) nursery.grids = {};
  if (!nursery.grids[tabName]) {
    nursery.grids[tabName] = structuredClone(BLANK_GRID);
    persist();
  }
  // Grids saved by an older build predate later keys; fill them in rather
  // than let a missing key throw halfway through a render.
  return Object.assign(structuredClone(BLANK_GRID), nursery.grids[tabName]);
}

/** Merge a patch into a tab's grid state and save. */
export function setGridState(tabName, patch) {
  const nursery = activeNursery();
  if (!nursery.grids) nursery.grids = {};
  const current = nursery.grids[tabName] ?? structuredClone(BLANK_GRID);
  nursery.grids[tabName] = { ...current, ...patch };
  persist();
  return nursery.grids[tabName];
}

/** Drop every stored tweak for one tab — the grid's "reset" action. */
export function clearGridState(tabName) {
  const nursery = activeNursery();
  if (nursery.grids) delete nursery.grids[tabName];
  persist();
}
