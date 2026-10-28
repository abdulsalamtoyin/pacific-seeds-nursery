// Download and restore of every nursery as one file — for browsers that
// cannot work in a team folder (folder-drive.js), and as a manual backup.

const APP = "ps-nursery-workbook";
const SCHEMA = 1;

/** The file is not one this app wrote. */
export class NotOurFileError extends Error {}

/** Every nursery as file text. */
export function serialise(nurseries, now = new Date()) {
  return JSON.stringify({
    app: APP, schema: SCHEMA, savedAt: now.toISOString(), nurseries,
  }, null, 1);
}

/** File text back to `{savedAt, nurseries}`; null for an empty file. */
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
