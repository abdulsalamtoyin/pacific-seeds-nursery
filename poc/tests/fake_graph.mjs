// A stand-in for the slice of Microsoft Graph that pwa/onedrive.js uses.
//
// It enforces the rules the sync engine depends on — eTags that change on
// every write, If-Match refusing a stale write with 412, conflictBehavior=fail
// refusing to overwrite with 409, listings that page — so the tests exercise
// the real adapter, not a mock of it. Faults (offline, throttling, expired
// tokens) can be switched on per test.

const APPROOT = "https://graph.microsoft.com/v1.0/me/drive/special/approot";
const DL = "https://download.example/";

const json = (status, body, headers = {}) => new Response(
  body === undefined ? null : JSON.stringify(body),
  { status, headers: { "Content-Type": "application/json", ...headers } });

export class FakeGraph {
  constructor({ pageSize = 200 } = {}) {
    this.files = new Map();            // name -> {content, version}
    this.pageSize = pageSize;
    this.validTokens = new Set(["tok-1"]);
    this.offline = false;
    this.queued = [];                  // canned responses served first
    this.log = [];                     // "METHOD url" per request
    this.violations = [];              // rules the client broke
  }

  eTag(name) {
    const f = this.files.get(name);
    return f ? `"{${name}},${f.version}"` : null;
  }

  /** Put a file there directly, as another device or a person would. */
  seed(name, body) {
    const prev = this.files.get(name);
    this.files.set(name, {
      content: typeof body === "string" ? body : JSON.stringify(body),
      version: (prev?.version ?? 0) + 1,
    });
  }

  read(name) {
    const f = this.files.get(name);
    return f ? JSON.parse(f.content) : null;
  }

  fetch = async (url, init = {}) => {
    const method = init.method || "GET";
    this.log.push(`${method} ${url}`);
    if (this.offline) throw new TypeError("Failed to fetch");
    if (this.queued.length) return this.queued.shift()();

    const headers = Object.fromEntries(
      Object.entries(init.headers || {}).map(([k, v]) => [k.toLowerCase(), v]));

    if (url.startsWith(DL)) {
      if (headers.authorization) {
        this.violations.push("token sent to a pre-signed download URL");
      }
      const name = decodeURIComponent(url.slice(DL.length).split("?")[0]);
      const f = this.files.get(name);
      return f ? new Response(f.content, { status: 200 }) : json(404, {});
    }

    const token = (headers.authorization || "").replace(/^Bearer /, "");
    if (!this.validTokens.has(token)) {
      return json(401, { error: { message: "InvalidAuthenticationToken" } });
    }

    const u = new URL(url);
    const path = decodeURIComponent(u.pathname);
    const base = "/v1.0/me/drive/special/approot";

    if (method === "GET" && path === `${base}/children`) {
      const skip = Number(u.searchParams.get("$skiptoken") || 0);
      const names = [...this.files.keys()].sort();
      const page = names.slice(skip, skip + this.pageSize);
      const body = {
        value: page.map((name) => ({
          name, eTag: this.eTag(name), file: { mimeType: "application/json" },
          "@microsoft.graph.downloadUrl": `${DL}${encodeURIComponent(name)}?sig=x`,
        })),
      };
      if (skip + this.pageSize < names.length) {
        body["@odata.nextLink"] =
          `${APPROOT}/children?$top=${this.pageSize}&$skiptoken=${skip + this.pageSize}`;
      }
      return json(200, body);
    }

    const m = path.match(/^\/v1\.0\/me\/drive\/special\/approot:\/(.+?)(:\/content)?$/);
    if (!m) return json(400, { error: { message: `unexpected ${method} ${path}` } });
    const name = m[1];
    const exists = this.files.has(name);
    const ifMatch = headers["if-match"];

    if (method === "PUT" && m[2]) {
      if (ifMatch && ifMatch !== this.eTag(name)) {
        return json(412, { error: { message: "eTag mismatch" } });
      }
      if (!ifMatch && exists
          && u.searchParams.get("@microsoft.graph.conflictBehavior") === "fail") {
        return json(409, { error: { message: "nameAlreadyExists" } });
      }
      if (!ifMatch && !u.searchParams.has("@microsoft.graph.conflictBehavior")) {
        this.violations.push(`unguarded overwrite of ${name}`);
      }
      this.seed(name, init.body);
      return json(exists ? 200 : 201, { name, eTag: this.eTag(name) });
    }

    if (method === "DELETE" && !m[2]) {
      if (!exists) return json(404, { error: { message: "itemNotFound" } });
      if (ifMatch && ifMatch !== this.eTag(name)) {
        return json(412, { error: { message: "eTag mismatch" } });
      }
      this.files.delete(name);
      return new Response(null, { status: 204 });
    }

    if (method === "GET" && m[2]) {
      return exists ? new Response(this.files.get(name).content) : json(404, {});
    }
    return json(400, { error: { message: `unexpected ${method} ${path}` } });
  };
}
