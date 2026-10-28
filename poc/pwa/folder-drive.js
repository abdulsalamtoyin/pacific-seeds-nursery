// A shared folder on this computer, as a drive for sync.js — the no-Azure way
// for a team to work on the same nurseries.
//
// The team syncs one SharePoint/Teams folder to their computers with the
// OneDrive desktop app; each person points the app at their local copy. Each
// nursery is one JSON file there, exactly as in the OneDrive app folder, and
// the OneDrive app carries the files between computers. Chrome and Edge only:
// other browsers cannot write to a chosen folder.
//
// Same contract as onedrive.js. A file's version ("eTag") is its modified time
// and size, so a file the OneDrive app replaced with someone else's save is
// refused (ConflictError) instead of overwritten, and sync.js asks the user.

import { ConflictError, OfflineError } from "./onedrive.js";

// A file the OneDrive app has not downloaded (online-only, while offline) or
// is busy replacing cannot be read. Treated as offline: sync stops and tries
// again later, and nothing is read as deleted.
const unreadable = (name, err) => new OfflineError(
  `${name} in the team folder could not be read (${err.message}). If OneDrive ` +
  "shows it as online-only, right-click the folder and choose “Always keep on " +
  "this device”.");

const versionOf = (file) => `${file.lastModified}:${file.size}`;

/** @param dir  a FileSystemDirectoryHandle with readwrite permission */
export function createFolderDrive(dir) {
  async function stat(name) {
    let handle;
    try {
      handle = await dir.getFileHandle(name);
    } catch (err) {
      if (err.name === "NotFoundError") return null;
      throw unreadable(name, err);
    }
    try {
      return { handle, file: await handle.getFile() };
    } catch (err) {
      throw unreadable(name, err);
    }
  }

  return {
    /** `[{name, eTag}]` for every .json file in the folder. */
    async list() {
      const files = [];
      for await (const [name, handle] of dir.entries()) {
        if (handle.kind !== "file" || !name.endsWith(".json")) continue;
        try {
          files.push({ name, eTag: versionOf(await handle.getFile()) });
        } catch (err) {
          throw unreadable(name, err);
        }
      }
      return files;
    },

    async download(file) {
      const found = await stat(file.name);
      if (!found) throw new ConflictError(`${file.name} was removed meanwhile.`);
      try {
        return JSON.parse(await found.file.text());
      } catch (err) {
        if (err instanceof SyntaxError) return null;   // sync.js reports it
        throw unreadable(file.name, err);
      }
    },

    /** As onedrive.js: with `eTag` it must still be that version; without, it must not exist. */
    async upload(name, body, eTag = null) {
      const found = await stat(name);
      if (eTag ? !found || versionOf(found.file) !== eTag : found) {
        throw new ConflictError(`${name} was changed on another computer.`);
      }
      const handle = found?.handle ?? await dir.getFileHandle(name, { create: true });
      const out = await handle.createWritable();
      await out.write(JSON.stringify(body));
      await out.close();
      return { eTag: versionOf(await handle.getFile()) };
    },

    async remove(name, eTag) {
      const found = await stat(name);
      if (!found) return;
      if (eTag && versionOf(found.file) !== eTag) {
        throw new ConflictError(`${name} was changed on another computer.`);
      }
      await dir.removeEntry(name);
    },
  };
}

// ---------------------------------------------------------------- browser
//
// The folder handle survives a reload only through IndexedDB. Chrome still
// asks again for permission after a reload, and only from a click.

export const folderSupported = () =>
  typeof globalThis.showDirectoryPicker === "function";

export const pickFolder = () =>
  globalThis.showDirectoryPicker({ id: "nursery-team-folder", mode: "readwrite" });

/** Ask for (or, with `ask` false, only check) permission to write. */
export async function canWrite(dir, ask) {
  const opts = { mode: "readwrite" };
  if (await dir.queryPermission(opts) === "granted") return true;
  return ask && await dir.requestPermission(opts) === "granted";
}

/** `{load, remember, forget}` over one IndexedDB record. */
export function idbMeta(dbName = "ps-team-folder") {
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
