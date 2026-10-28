// The user's OneDrive app folder, through Microsoft Graph.
//
// Graph's "approot" special folder is OneDrive › Apps › <app name>, created by
// OneDrive the first time it is touched, and the Files.ReadWrite.AppFolder
// permission reaches nothing outside it. Each nursery is one JSON file there.
//
// Every write carries the eTag it last saw, so a file changed on another
// computer is refused (412) instead of silently overwritten. sync.js decides
// what to do about that; this module only reports it.

const APPROOT = "https://graph.microsoft.com/v1.0/me/drive/special/approot";
const RETRIES = 3;

/** Another device changed or created the file since we last saw it. */
export class ConflictError extends Error {}
/** The network is down; try again later. */
export class OfflineError extends Error {}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * @param {() => Promise<string>} getToken  access token; called with
 *        {force: true} after a 401 so an expired token is renewed once.
 * @param {typeof fetch} fetchImpl           seam for tests.
 */
export function createDrive({ getToken, fetchImpl = globalThis.fetch,
  wait = sleep } = {}) {
  async function send(url, init = {}, { auth = true } = {}) {
    for (let attempt = 0, forced = false; ; attempt++) {
      const headers = { ...(init.headers || {}) };
      if (auth) headers.Authorization = `Bearer ${await getToken({ force: forced })}`;

      let res;
      try {
        res = await fetchImpl(url, { ...init, headers });
      } catch (err) {
        throw new OfflineError(`OneDrive could not be reached: ${err.message}`);
      }

      if (res.status === 401 && auth && !forced) { forced = true; continue; }
      // Throttled or briefly unavailable: Graph says how long to wait.
      if ([429, 503, 504].includes(res.status) && attempt < RETRIES) {
        const seconds = Number(res.headers.get("Retry-After")) || 2 ** attempt;
        await wait(Math.min(seconds, 30) * 1000);
        continue;
      }
      return res;
    }
  }

  async function fail(res, what) {
    let detail = `${res.status}`;
    try { detail = (await res.json()).error?.message || detail; } catch { /* not JSON */ }
    throw new Error(`OneDrive ${what} failed: ${detail}`);
  }

  const itemUrl = (name) => `${APPROOT}:/${encodeURIComponent(name)}`;

  return {
    /** `[{name, eTag, downloadUrl}]` for every .json file in the app folder. */
    async list() {
      const files = [];
      let url = `${APPROOT}/children?$top=200`;
      while (url) {
        const res = await send(url);
        if (res.status === 404) return files;     // folder not made yet
        if (!res.ok) await fail(res, "listing");
        const body = await res.json();
        for (const item of body.value || []) {
          if (item.file && item.name.endsWith(".json")) {
            files.push({
              name: item.name,
              eTag: item.eTag,
              downloadUrl: item["@microsoft.graph.downloadUrl"],
            });
          }
        }
        url = body["@odata.nextLink"] || null;
      }
      return files;
    },

    /** A listed file's JSON. The download URL is pre-signed: no token. */
    async download(file) {
      const res = file.downloadUrl
        ? await send(file.downloadUrl, {}, { auth: false })
        : await send(`${itemUrl(file.name)}:/content`);
      if (!res.ok) await fail(res, `download of ${file.name}`);
      return res.json();
    },

    /**
     * Write a file. With `eTag` it must still be that version; without, it
     * must not exist yet. Either way a mismatch is a ConflictError.
     * @returns {Promise<{eTag: string}>}
     */
    async upload(name, body, eTag = null) {
      const url = `${itemUrl(name)}:/content` +
        (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail");
      const headers = { "Content-Type": "application/json" };
      if (eTag) headers["If-Match"] = eTag;
      const res = await send(url, {
        method: "PUT", headers, body: JSON.stringify(body),
      });
      if (res.status === 412 || res.status === 409) {
        throw new ConflictError(`${name} was changed on another device.`);
      }
      if (!res.ok) await fail(res, `upload of ${name}`);
      return { eTag: (await res.json()).eTag };
    },

    /** Delete a file, if it is still the version we saw. Gone already is fine. */
    async remove(name, eTag) {
      const res = await send(itemUrl(name), {
        method: "DELETE", headers: eTag ? { "If-Match": eTag } : {},
      });
      if (res.status === 412) {
        throw new ConflictError(`${name} was changed on another device.`);
      }
      if (!res.ok && res.status !== 404) await fail(res, `delete of ${name}`);
    },
  };
}
