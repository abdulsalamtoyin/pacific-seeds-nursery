// Sorghum Nursery PWA — single-screen replacement capture.
// Storage: IndexedDB (manifest cache + offline event outbox).
// Scanning: native BarcodeDetector when available; falls back to jsQR via dynamic import.

const $ = sel => document.querySelector(sel);
const uuid = () =>
  (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));

// ---------- IndexedDB helpers ----------
const DB_NAME = "sorghum-nursery";
const DB_VERSION = 2;

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (ev) => {
      const db = req.result;
      if (!db.objectStoreNames.contains("manifest"))
        db.createObjectStore("manifest", { keyPath: "uuid" });
      if (!db.objectStoreNames.contains("meta"))
        db.createObjectStore("meta", { keyPath: "key" });
      if (!db.objectStoreNames.contains("outbox"))
        db.createObjectStore("outbox", { keyPath: "uuid" });
      if (!db.objectStoreNames.contains("sent"))
        db.createObjectStore("sent", { keyPath: "uuid" });
      // v2: index manifest by nursery_code for fast per-nursery deletion.
      const tx = ev.target.transaction;
      const ms = tx.objectStore("manifest");
      if (!ms.indexNames.contains("by_nursery"))
        ms.createIndex("by_nursery", "nursery_code", { unique: false });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbByIndex(store, indexName, key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const idx = db.transaction(store).objectStore(store).index(indexName);
    const req = idx.getAll(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbDeleteByIndex(store, indexName, key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, "readwrite");
    const idx = tx.objectStore(store).index(indexName);
    const cur = idx.openCursor(key);
    cur.onsuccess = () => {
      const c = cur.result;
      if (c) { c.delete(); c.continue(); }
    };
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function idbAll(store) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store).objectStore(store).getAll();
    tx.onsuccess = () => resolve(tx.result);
    tx.onerror = () => reject(tx.error);
  });
}
async function idbGet(store, key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store).objectStore(store).get(key);
    tx.onsuccess = () => resolve(tx.result);
    tx.onerror = () => reject(tx.error);
  });
}
async function idbPut(store, value) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, "readwrite");
    tx.objectStore(store).put(value);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}
async function idbClear(store) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, "readwrite");
    tx.objectStore(store).clear();
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}
async function idbDelete(store, key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, "readwrite");
    tx.objectStore(store).delete(key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

// ---------- UI status ----------
function setOnline(online) {
  const el = $("#onlineStatus");
  el.textContent = online ? "online" : "offline";
  el.classList.toggle("warn", !online);
}
window.addEventListener("online", () => { setOnline(true); trySync(); });
window.addEventListener("offline", () => setOnline(false));

async function refreshOutboxCount() {
  const items = await idbAll("outbox");
  $("#outboxStatus").textContent = `outbox: ${items.length}`;
  $("#outboxStatus").classList.toggle("warn", items.length > 0);
}

// ---------- Nursery loading ----------
async function loadedNurseryCodes() {
  const all = await idbAll("manifest");
  const codes = new Set();
  for (const p of all) if (p.nursery_code) codes.add(p.nursery_code);
  return [...codes].sort();
}

async function loadNurseryList() {
  try {
    const r = await fetch("/nurseries");
    const codes = await r.json();
    const sel = $("#nurseryPick");
    sel.innerHTML = '<option value="">— choose nursery —</option>' +
      codes.map(c => `<option>${c}</option>`).join("");
  } catch {
    $("#manifestHint").innerHTML = '<span class="err">Could not reach backend. Are you offline?</span>';
  }
  const meta = await idbGet("meta", "current");
  if (meta && meta.tech_id) $("#techId").value = meta.tech_id;
  await renderLoadedList();
  if (meta && meta.nursery_code) {
    await activate(meta.nursery_code);
  }
}

async function renderLoadedList() {
  const codes = await loadedNurseryCodes();
  const meta = await idbGet("meta", "current");
  const list = $("#loadedList");
  if (codes.length === 0) {
    list.innerHTML = '<div class="hint">None cached yet. Pick one below and tap Load.</div>';
    return;
  }
  list.innerHTML = "";
  for (const code of codes) {
    const items = await idbByIndex("manifest", "by_nursery", code);
    const isActive = meta && meta.nursery_code === code;
    const el = document.createElement("div");
    el.className = "nurs" + (isActive ? " active" : "");
    el.innerHTML = `<b>${code}</b> <span class="count">${items.length} pkts</span>`;
    const useBtn = document.createElement("button");
    useBtn.textContent = isActive ? "✓ active" : "use";
    useBtn.disabled = isActive;
    useBtn.onclick = () => activate(code);
    el.appendChild(useBtn);
    const dropBtn = document.createElement("button");
    dropBtn.textContent = "✕";
    dropBtn.title = "Drop from device cache";
    dropBtn.onclick = async () => {
      if (!confirm(`Drop cached packets for ${code}?`)) return;
      await idbDeleteByIndex("manifest", "by_nursery", code);
      if (isActive) await idbDelete("meta", "current");
      await renderLoadedList();
      await refreshActiveBar();
    };
    el.appendChild(dropBtn);
    list.appendChild(el);
  }
}

async function activate(code) {
  const meta = (await idbGet("meta", "current")) || {};
  await idbPut("meta", { key: "current",
    nursery_code: code,
    tech_id: $("#techId").value.trim() || meta.tech_id || "",
  });
  await renderLoadedList();
  await refreshActiveBar();
  await refreshServerFeed();
  $("#scanCard").classList.remove("hidden");
}

async function refreshActiveBar() {
  const meta = await idbGet("meta", "current");
  if (!meta || !meta.nursery_code) {
    $("#activeBar").classList.add("hidden");
    $("#scanCard").classList.add("hidden");
    $("#coachBar").classList.add("hidden");
    return;
  }
  $("#activeBar").classList.remove("hidden");
  $("#activeNursery").textContent = meta.nursery_code;
  $("#activeTech").textContent = meta.tech_id || "—";
  await coachRender();
}

$("#switchBtn").addEventListener("click", () => {
  $("#setupCard").scrollIntoView({ behavior: "smooth" });
});

$("#loadBtn").addEventListener("click", async () => {
  const code = $("#nurseryPick").value;
  const tech = $("#techId").value.trim();
  if (!code) return alert("Pick a nursery first.");
  if (!tech) return alert("Enter your name or initials.");
  $("#manifestHint").textContent = "Downloading manifest…";
  try {
    const r = await fetch(`/nursery/${encodeURIComponent(code)}/manifest`);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const data = await r.json();
    await idbDeleteByIndex("manifest", "by_nursery", code);
    for (const p of data.packets) await idbPut("manifest", { ...p, nursery_code: code });
    $("#manifestHint").innerHTML =
      `<span class="ok">Cached ${data.packet_count} packets for ${code}. You can go offline now.</span>`;
    await activate(code);
  } catch (err) {
    $("#manifestHint").innerHTML = `<span class="err">Failed: ${err.message}</span>`;
  }
});

$("#resetBtn").addEventListener("click", async () => {
  if (!confirm("Forget all cached nurseries and outbox? Unsynced events will be lost.")) return;
  await idbClear("manifest"); await idbClear("meta"); await idbClear("outbox"); await idbClear("sent");
  // Forgetting nursery also returns the user to the landing/intro page.
  location.href = "/";
});

// ---------- QR scanning ----------
let scanStream = null;
let scanLoop = null;

$("#scanBtn").addEventListener("click", startScan);
$("#stopScanBtn").addEventListener("click", stopScan);
$("#findBtn").addEventListener("click", () => {
  $("#findArea").classList.toggle("hidden");
  $("#scanArea").classList.add("hidden");
  stopScan();
  if (!$("#findArea").classList.contains("hidden")) {
    $("#plotInput").focus();
    renderFindResults("");
  }
});
$("#plotInput").addEventListener("input", (e) => renderFindResults(e.target.value.trim()));
$("#manualGoBtn").addEventListener("click", () => handlePayload($("#manualInput").value.trim()));

async function renderFindResults(q) {
  const meta = await idbGet("meta", "current");
  if (!meta) return;
  const items = await idbByIndex("manifest", "by_nursery", meta.nursery_code);
  const ql = q.toLowerCase();
  const hits = !q ? items.slice(0, 12)
    : items.filter(p =>
        (p.plot || "").toLowerCase().includes(ql) ||
        (p.material_id || "").toLowerCase().includes(ql) ||
        (p.source_id || "").toLowerCase().includes(ql)
      ).slice(0, 30);
  const box = $("#findResults");
  if (hits.length === 0) { box.innerHTML = '<div class="hint">No matches.</div>'; return; }
  box.innerHTML = "";
  for (const p of hits) {
    const div = document.createElement("div");
    div.className = "hit";
    div.innerHTML = `<b>Plot ${p.plot}</b> · Spike ${p.spike} · Rack ${p.rack_order}
                     <div class="hint">${p.material_id || "—"} · ${p.source_id || "—"}</div>`;
    div.onclick = () => { $("#findArea").classList.add("hidden"); showPacket(p); };
    box.appendChild(div);
  }
}

function explainCameraError(err) {
  // 1. Insecure context: no mediaDevices at all.
  const insecure = !window.isSecureContext ||
                   !navigator.mediaDevices ||
                   !navigator.mediaDevices.getUserMedia;
  if (insecure) {
    const host = location.hostname;
    return (
      `The browser is blocking the camera because this page isn't on HTTPS.\n\n` +
      `You're at: ${location.protocol}//${host}\n\n` +
      `Two fixes:\n` +
      `  • On a phone: ask IT/dev to start the server with HTTPS\n` +
      `    (see poc/scripts/make_cert.sh in the project), OR\n` +
      `  • On your laptop, open the app via http://localhost:8000/\n` +
      `    (localhost is exempt from the HTTPS rule).\n\n` +
      `Meanwhile, use "Find by plot" instead — it works without the camera.`
    );
  }
  // 2. Specific DOMException names.
  switch (err && err.name) {
    case "NotAllowedError":
    case "PermissionDeniedError":
      return "Camera permission was denied. Open your browser's site settings " +
             "for this page and re-allow camera access, then refresh.";
    case "NotFoundError":
    case "DevicesNotFoundError":
      return "No camera found on this device. Use 'Find by plot' instead.";
    case "NotReadableError":
    case "TrackStartError":
      return "Camera is in use by another app. Close other apps using the camera " +
             "(FaceTime, Zoom, Snap…), then try again.";
    case "OverconstrainedError":
      return "The requested camera (rear/environment) isn't available. Trying any camera.";
    default:
      return `Camera error: ${(err && err.message) || err}`;
  }
}

async function startScan() {
  $("#scanArea").classList.remove("hidden");
  // Pre-flight: secure context required.
  if (!window.isSecureContext || !navigator.mediaDevices?.getUserMedia) {
    alert(explainCameraError(null));
    $("#scanArea").classList.add("hidden");
    return;
  }
  try {
    scanStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: "environment" } }, audio: false,
    });
  } catch (err) {
    if (err.name === "OverconstrainedError") {
      try {
        scanStream = await navigator.mediaDevices.getUserMedia({ video: true, audio: false });
      } catch (err2) {
        alert(explainCameraError(err2));
        $("#scanArea").classList.add("hidden");
        return;
      }
    } else {
      alert(explainCameraError(err));
      $("#scanArea").classList.add("hidden");
      return;
    }
  }

  const video = $("#scanner");
  video.srcObject = scanStream;
  try { await video.play(); } catch {}

  if ("BarcodeDetector" in window) {
    const detector = new BarcodeDetector({ formats: ["qr_code"] });
    scanLoop = setInterval(async () => {
      try {
        const codes = await detector.detect(video);
        if (codes && codes[0]) { stopScan(); handlePayload(codes[0].rawValue); }
      } catch {}
    }, 250);
  } else {
    try {
      const jsQR = (await import("https://cdn.jsdelivr.net/npm/jsqr@1.4.0/+esm")).default;
      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d");
      scanLoop = setInterval(() => {
        if (video.readyState < 2) return;
        canvas.width = video.videoWidth; canvas.height = video.videoHeight;
        ctx.drawImage(video, 0, 0);
        const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const code = jsQR(img.data, img.width, img.height);
        if (code) { stopScan(); handlePayload(code.data); }
      }, 250);
    } catch (err) {
      alert("This browser can't decode QR codes and the jsQR fallback failed to load (offline?). Use 'Find by plot'.");
      stopScan();
    }
  }
}

function stopScan() {
  if (scanLoop) { clearInterval(scanLoop); scanLoop = null; }
  if (scanStream) { scanStream.getTracks().forEach(t => t.stop()); scanStream = null; }
  $("#scanArea").classList.add("hidden");
}

// ---------- Packet lookup ----------
async function handlePayload(raw) {
  if (!raw) return;
  const m = raw.match(/^SNUR:([^:]+):([a-zA-Z0-9-]+)$/);
  if (!m) return alert(`Not a Sorghum packet QR:\n${raw}`);
  const [_, nurseryCode, packetUuid] = m;
  const meta = await idbGet("meta", "current");
  if (!meta) return alert("No active nursery. Load one first.");
  if (meta.nursery_code !== nurseryCode) {
    // Offer to auto-switch if the QR's nursery is one we've cached.
    const cached = await loadedNurseryCodes();
    if (cached.includes(nurseryCode)) {
      if (!confirm(`This QR is for ${nurseryCode}. Switch to it?`)) return;
      await activate(nurseryCode);
    } else {
      return alert(`This QR is for nursery '${nurseryCode}' but you haven't loaded it.`);
    }
  }
  const packet = await idbGet("manifest", packetUuid);
  if (!packet || packet.nursery_code !== nurseryCode) {
    return alert(`Packet ${packetUuid} not found in cached manifest.`);
  }
  showPacket(packet);
}

// ---------- Event-type forms ----------
// Each entry: { label, fields: [{name, label, type, options?, required?, default?}], saveLabel }
const EVENT_TYPES = {
  replacement: {
    label: "Replacement",
    saveLabel: "Save replacement",
    fields: [
      { name: "replaced_with", label: "Replaced with (new Source ID)",
        type: "text", required: true, placeholder: "e.g. AUGT1-25W-AB-0011-B-0012" },
      { name: "stage", label: "Stage", type: "select",
        options: ["Packeting", "Planting"], default: "Packeting" },
    ],
  },
  planting_error: {
    label: "Planting error",
    saveLabel: "Save planting error",
    fields: [
      { name: "error_kind", label: "What went wrong?", type: "select",
        options: ["Missed plot", "Wrong packet planted", "Damaged packet",
                  "Mixed packets", "Other"] },
      { name: "severity", label: "Severity", type: "select",
        options: ["Low", "Medium", "High"], default: "Medium" },
    ],
  },
  spray: {
    label: "Spray applied",
    saveLabel: "Save spray",
    fields: [
      { name: "product", label: "Product", type: "select",
        options: ["TFMSA", "IMI", "HPPD", "Other"] },
      { name: "applied_on", label: "Date applied", type: "date",
        default: () => new Date().toISOString().slice(0, 10) },
      { name: "rate", label: "Rate (optional)", type: "text", placeholder: "e.g. 200 mL/ha" },
    ],
  },
  ab_pull: {
    label: "AB bag pulled",
    saveLabel: "Save AB pull",
    fields: [
      { name: "bag_count", label: "Bags pulled", type: "number", default: 1 },
      { name: "pulled_on", label: "Date pulled", type: "date",
        default: () => new Date().toISOString().slice(0, 10) },
    ],
  },
  date_record: {
    label: "Date record",
    saveLabel: "Save date record",
    fields: [
      { name: "event_label", label: "What happened?", type: "select",
        options: ["Heading", "Flowering", "Bagging", "Crossing", "Harvest"] },
      { name: "occurred_on", label: "Date", type: "date",
        default: () => new Date().toISOString().slice(0, 10) },
    ],
  },
  note: {
    label: "Note",
    saveLabel: "Save note",
    fields: [],
  },
};

function renderForm(typeKey) {
  const def = EVENT_TYPES[typeKey];
  const root = $("#formFields");
  root.innerHTML = "";
  for (const f of def.fields) {
    const id = `f_${f.name}`;
    const label = document.createElement("label");
    label.htmlFor = id;
    label.textContent = f.label;
    root.appendChild(label);
    let el;
    if (f.type === "select") {
      el = document.createElement("select");
      el.innerHTML = f.options.map(o => `<option>${o}</option>`).join("");
    } else if (f.type === "date") {
      el = document.createElement("input");
      el.type = "date";
    } else if (f.type === "number") {
      el = document.createElement("input");
      el.type = "number"; el.inputMode = "numeric";
    } else {
      el = document.createElement("input");
      el.type = "text";
      if (f.placeholder) el.placeholder = f.placeholder;
    }
    el.id = id;
    el.name = f.name;
    el.dataset.required = f.required ? "1" : "";
    if (f.default !== undefined) {
      el.value = typeof f.default === "function" ? f.default() : f.default;
    }
    root.appendChild(el);
    const sp = document.createElement("div");
    sp.style.height = "10px";
    root.appendChild(sp);
  }
  $("#saveBtn").textContent = def.saveLabel;
}

function showPacket(p) {
  $("#packetCard").classList.remove("hidden");
  $("#packetCard").scrollIntoView({ behavior: "smooth" });
  $("#packetPlot").textContent = `Plot ${p.plot}`;
  $("#packetMeta").innerHTML =
    `<b>Material:</b> ${p.material_id || "—"}<br/>` +
    `<b>Source:</b> ${p.source_id || "—"}<br/>` +
    `<b>Pedigree:</b> ${p.pedigree || "—"}`;
  $("#packetPills").innerHTML = [
    `Spike ${p.spike}`, `Rack ${p.rack_order}`,
    p.generation && `Gen ${p.generation}`, p.cms && `CMS ${p.cms}`,
  ].filter(Boolean).map(t => `<span>${t}</span>`).join("");
  $("#packetCard").dataset.uuid = p.uuid;
  $("#packetCard").dataset.plot = p.plot;
  $("#packetCard").dataset.source = p.source_id || "";
  $("#packetCard").dataset.type = "";
  $("#formArea").classList.add("hidden");
  $("#formFields").innerHTML = "";
  $("#noteInput").value = "";
  $("#saveHint").textContent = "";
  document.querySelectorAll(".type-btn").forEach(b => b.classList.remove("active"));
}

// Wire up the event-type chooser.
document.querySelectorAll(".type-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".type-btn").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    const t = btn.dataset.type;
    $("#packetCard").dataset.type = t;
    renderForm(t);
    $("#formArea").classList.remove("hidden");
    const first = $("#formFields").querySelector("input,select");
    if (first) setTimeout(() => first.focus(), 50);
  });
});

$("#cancelBtn").addEventListener("click", () => {
  $("#packetCard").classList.add("hidden");
});

$("#saveBtn").addEventListener("click", async () => {
  const card = $("#packetCard");
  const typeKey = card.dataset.type;
  if (!typeKey) return alert("Pick an event type first.");
  const def = EVENT_TYPES[typeKey];

  // Collect & validate fields.
  const payload = { plot: card.dataset.plot };
  if (typeKey === "replacement") payload.original_source = card.dataset.source;
  for (const f of def.fields) {
    const el = document.getElementById(`f_${f.name}`);
    const v = (el.value ?? "").toString().trim();
    if (f.required && !v) return alert(`Please fill in: ${f.label}`);
    payload[f.name] = v || null;
  }
  const noteVal = $("#noteInput").value.trim();
  if (noteVal) payload.note = noteVal;

  const meta = await idbGet("meta", "current");
  const ev = {
    uuid: uuid(),
    packet_uuid: card.dataset.uuid,
    nursery_code: meta.nursery_code,
    type: typeKey,
    payload,
    tech_id: meta.tech_id,
    captured_at: new Date().toISOString(),
  };
  await idbPut("outbox", ev);
  await refreshOutboxCount();
  await renderRecent();
  $("#saveHint").innerHTML = '<span class="ok">Saved locally. Will sync when online.</span>';
  card.classList.add("hidden");
  trySync();
  await coachAutoDetect("event_" + typeKey);
});

// ---------- Sync ----------
$("#syncBtn").addEventListener("click", trySync);

async function trySync() {
  const items = await idbAll("outbox");
  if (items.length === 0) {
    $("#syncBtn").textContent = "Sync now (nothing to send)";
    setTimeout(() => $("#syncBtn").textContent = "Sync now", 1500);
    return;
  }
  try {
    const r = await fetch("/sync", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ events: items }),
    });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const result = await r.json();
    for (const ev of items) { await idbPut("sent", ev); await idbDelete("outbox", ev.uuid); }
    await refreshOutboxCount();
    $("#syncBtn").textContent =
      `Synced ${result.accepted} (${result.duplicates} dup)`;
    setTimeout(() => $("#syncBtn").textContent = "Sync now", 2000);
    await refreshServerFeed();
    if (result.accepted > 0) await coachAutoDetect("sync_success");
  } catch (err) {
    $("#syncBtn").textContent = `Sync failed: ${err.message}`;
    setTimeout(() => $("#syncBtn").textContent = "Sync now", 2500);
  }
}

// ---------- Recent ----------
function summariseEvent(e) {
  const p = e.payload || {};
  switch (e.type) {
    case "replacement":    return `${p.original_source || "?"} → ${p.replaced_with || "?"} (${p.stage || ""})`;
    case "planting_error": return `${p.error_kind || "Error"} · ${p.severity || ""}`;
    case "spray":          return `${p.product || "Spray"} on ${p.applied_on || "?"}${p.rate ? " · " + p.rate : ""}`;
    case "ab_pull":        return `${p.bag_count || 1} bag(s) pulled ${p.pulled_on || ""}`;
    case "date_record":    return `${p.event_label || "Event"} on ${p.occurred_on || "?"}`;
    case "note":           return p.note || "(no text)";
    default:               return JSON.stringify(p);
  }
}

let feedMode = "all"; // "all" | "device"

$("#feedTabDevice").addEventListener("click", () => { feedMode = "device";
  $("#feedTabDevice").classList.add("active"); $("#feedTabAll").classList.remove("active");
  renderRecent(); });
$("#feedTabAll").addEventListener("click", () => { feedMode = "all";
  $("#feedTabAll").classList.add("active"); $("#feedTabDevice").classList.remove("active");
  renderRecent(); });

function renderEventLine(e) {
  const def = EVENT_TYPES[e.type];
  const typeLabel = def ? def.label : e.type;
  const plot = e.plot || e.payload?.plot || "?";
  return `<div class="ev">
    <b>Plot ${plot}</b> · ${typeLabel}
    <div class="hint">${summariseEvent(e)}</div>
    <div class="hint">${new Date(e.captured_at).toLocaleString()} · ${e.tech_id || ""}</div>
  </div>`;
}

async function refreshServerFeed() {
  if (feedMode === "all") await renderRecent();
}

async function renderRecent() {
  const list = $("#recentList");
  const meta = await idbGet("meta", "current");
  let items = [];
  if (feedMode === "device") {
    items = [...(await idbAll("outbox")), ...(await idbAll("sent"))]
      .filter(e => !meta || e.nursery_code === meta.nursery_code)
      .sort((a, b) => b.captured_at.localeCompare(a.captured_at))
      .slice(0, 20);
  } else {
    if (!meta) { list.innerHTML = '<div class="hint">Pick an active nursery.</div>'; return; }
    try {
      const r = await fetch(`/nursery/${encodeURIComponent(meta.nursery_code)}/events?limit=20`);
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      items = await r.json();
    } catch {
      // Offline fallback: show device-side.
      list.innerHTML = '<div class="hint">Offline — showing device only.</div>';
      items = [...(await idbAll("outbox")), ...(await idbAll("sent"))]
        .filter(e => !meta || e.nursery_code === meta.nursery_code)
        .sort((a, b) => b.captured_at.localeCompare(a.captured_at))
        .slice(0, 20);
    }
  }
  if (items.length === 0) { list.innerHTML = '<div class="hint">Nothing yet.</div>'; return; }
  list.innerHTML = items.map(renderEventLine).join("");
}

// ==================== Admin panel ====================
$("#admToggle").addEventListener("click", () => {
  const btn = $("#admToggle");
  const open = btn.getAttribute("aria-expanded") === "true";
  btn.setAttribute("aria-expanded", String(!open));
  $("#admBody").classList.toggle("hidden", open);
  if (!open) {
    // Lazy-load nursery list + map when panel opens
    loadNurseryListAdmin();
    loadFieldMap();
    const meta = idbGet("meta", "current");
    meta.then(m => { if (m?.nursery_code) $("#initCode").value = m.nursery_code; });
  }
});

// ----- Step 1: Initialise from PRISM upload -----
// When a file is picked, inspect it and populate the sheet dropdown.
$("#initFile").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  const sel = $("#initSheet");
  const hint = $("#initSheetHint");
  if (!file) {
    sel.innerHTML = '<option value="">— pick a file first —</option>';
    sel.disabled = true;
    hint.textContent = "";
    return;
  }
  $("#initFileHint").innerHTML = `<span class="ok">${file.name} — reading sheets…</span>`;
  sel.innerHTML = '<option>scanning…</option>';
  sel.disabled = true;
  hint.textContent = "";
  try {
    const fd = new FormData();
    fd.append("file", file);
    const r = await fetch("/admin/inspect", { method: "POST", body: fd });
    if (!r.ok) throw new Error(`HTTP ${r.status}: ${(await r.text()).slice(0,180)}`);
    const data = await r.json();
    sel.innerHTML = "";
    for (const s of data.sheets) {
      const opt = document.createElement("option");
      opt.value = s.name;
      const tag = s.has_prism_headers ? "  ✓ PRISM headers" : "";
      opt.textContent = `${s.name}  (${s.rows} rows × ${s.cols} cols)${tag}`;
      sel.appendChild(opt);
    }
    if (data.suggested) sel.value = data.suggested;
    sel.disabled = false;
    $("#initFileHint").innerHTML =
      `<span class="ok">${file.name} — ${data.sheets.length} sheet(s) found.</span>`;
    const picked = data.sheets.find(s => s.name === sel.value);
    if (picked && picked.headers_preview.length) {
      hint.innerHTML =
        `<span class="hint">Headers: <i>${picked.headers_preview.join(" · ")}</i></span>`;
    }
  } catch (err) {
    sel.innerHTML = '<option value="">(could not read sheets)</option>';
    sel.disabled = false;
    $("#initFileHint").innerHTML = `<span class="err">${err.message}</span>`;
  }
});

// When the user picks a different sheet, update the headers preview line.
$("#initSheet").addEventListener("change", async () => {
  const file = $("#initFile").files[0];
  if (!file) return;
  const name = $("#initSheet").value;
  if (!name) return;
  // Reuse the result of the last inspect (cached on the select element).
  // We re-fetch here only if no cached info; cheaper to just call again.
});

$("#initBtn").addEventListener("click", async () => {
  const code = $("#initCode").value.trim();
  const sheet = ($("#initSheet").value || "").trim() || "Sheet1";
  const file = $("#initFile").files[0];
  if (!code) return alert("Enter a nursery code first.");
  if (!file) return alert("Pick a PRISM export .xlsx file.");
  if (!sheet) return alert("Pick which sheet inside the file holds the PRISM data.");
  $("#initStatus").innerHTML = '<span class="ok">Uploading and processing…</span>';
  try {
    const fd = new FormData();
    fd.append("nursery_code", code);
    fd.append("sheet", sheet);
    fd.append("file", file);
    const r = await fetch("/admin/init", { method: "POST", body: fd });
    if (!r.ok) throw new Error(`HTTP ${r.status}: ${(await r.text()).slice(0,200)}`);
    const data = await r.json();
    $("#initStatus").innerHTML =
      `<span class="ok">✓ ${data.packet_count} packets imported. ` +
      `${data.spike_count} spike(s) from "${data.map_source}".</span>`;
    // Refresh the nursery list so this one shows up immediately
    await loadNurseryList();
    await loadNurseryListAdmin();
    await loadFieldMap();
    await coachAutoDetect("init_success");
  } catch (err) {
    $("#initStatus").innerHTML = `<span class="err">Failed: ${err.message}</span>`;
  }
});

// ----- Step 2: Nursery list -----
async function loadNurseryListAdmin() {
  const meta = await idbGet("meta", "current");
  if (!meta) { $("#listResults").textContent = "Pick a nursery first."; return; }
  try {
    const r = await fetch(`/nursery/${encodeURIComponent(meta.nursery_code)}/list`);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const data = await r.json();
    if (data.items.length === 0) {
      $("#listResults").textContent = "No source IDs found.";
      return;
    }
    let html = `<div class="hint" style="margin-bottom:6px">` +
               `${data.items.length} unique source IDs. ` +
               `Qty/packet = ${data.qty_per_packet}. ` +
               `Bulk-flagged when reps > ${data.bulk_threshold}.</div>`;
    html += `<table><thead><tr><th>Source ID</th><th>Reps</th><th>Qty req.</th>` +
            `<th>Inbred</th><th>Hybrid</th></tr></thead><tbody>`;
    for (const it of data.items) {
      html += `<tr class="${it.bulk ? 'bulk' : ''}">` +
              `<td><b>${it.source_id}</b></td><td>${it.reps}</td>` +
              `<td>${it.qty_required}</td>` +
              `<td>${it.inbred_code}</td><td>${it.hybrid_code}</td></tr>`;
    }
    html += "</tbody></table>";
    $("#listResults").innerHTML = html;
    $("#listResults")._items = data.items; // stash for CSV
    $("#listResults")._meta = data;
    await coachAutoDetect("list_loaded");
  } catch (err) {
    $("#listResults").innerHTML = `<span class="err">${err.message}</span>`;
  }
}
$("#listLoadBtn").addEventListener("click", loadNurseryListAdmin);
$("#listCsvBtn").addEventListener("click", async () => {
  const items = $("#listResults")._items;
  if (!items) return alert("Click Refresh list first.");
  const csv = "Source ID,Repeats,Qty Required,Inbred,Hybrid,Bulk\n" +
    items.map(i => [i.source_id, i.reps, i.qty_required,
                    i.inbred_code, i.hybrid_code, i.bulk ? "BULK" : ""].join(",")).join("\n");
  try { await navigator.clipboard.writeText(csv);
    alert("CSV copied to clipboard."); }
  catch { alert("Couldn't copy. CSV:\n\n" + csv.slice(0, 500)); }
});

// ----- Step 4 / 11: downloads -----
$("#pdfBtn").addEventListener("click", async () => {
  const meta = await idbGet("meta", "current");
  if (!meta) return alert("Pick a nursery first.");
  const url = `/nursery/${encodeURIComponent(meta.nursery_code)}/packets.pdf`;
  $("#downloadHint").textContent = "Opening packets PDF…";
  window.open(url, "_blank");
  await coachAutoDetect("pdf_downloaded");
});
$("#fbBtn").addEventListener("click", async () => {
  const meta = await idbGet("meta", "current");
  if (!meta) return alert("Pick a nursery first.");
  $("#downloadHint").innerHTML = '<span class="ok">Generating fieldbook…</span>';
  try {
    const r = await fetch(`/admin/fieldbook/${encodeURIComponent(meta.nursery_code)}`, {
      method: "POST"
    });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${meta.nursery_code}_fieldbook.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    $("#downloadHint").innerHTML =
      `<span class="ok">Downloaded ${meta.nursery_code}_fieldbook.xlsx</span>`;
    await coachAutoDetect("fieldbook_downloaded");
  } catch (err) {
    $("#downloadHint").innerHTML = `<span class="err">${err.message}</span>`;
  }
});

// ----- Field Map editor -----
async function loadFieldMap() {
  const meta = await idbGet("meta", "current");
  if (!meta) { $("#mapEditor").innerHTML = '<div class="hint">Pick a nursery first.</div>'; return; }
  try {
    const r = await fetch(`/nursery/${encodeURIComponent(meta.nursery_code)}/map`);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const data = await r.json();
    renderMapEditor(data.spikes);
    $("#mapStatus").innerHTML =
      data.source === "saved"
      ? `<span class="hint">Saved layout (${data.updated_at})</span>`
      : `<span class="hint">Derived from current packet data — save to lock in.</span>`;
  } catch (err) {
    $("#mapEditor").innerHTML = `<span class="err">${err.message}</span>`;
  }
}
function renderMapEditor(spikes) {
  const root = $("#mapEditor");
  root.innerHTML = "";
  // Header row
  const hdr = document.createElement("div");
  hdr.className = "spike-row";
  hdr.innerHTML = `<span class="lbl">Spike</span><span class="lbl">Row from</span>` +
                  `<span class="lbl">Row to</span><span></span>`;
  root.appendChild(hdr);
  for (const s of spikes) {
    addSpikeRow(s.spike, s.row_min, s.row_max);
  }
  if (spikes.length === 0) addSpikeRow(1, 1, 10);
}
function addSpikeRow(spike, rmin, rmax) {
  const row = document.createElement("div");
  row.className = "spike-row";
  row.innerHTML = `
    <input type="number" min="1" value="${spike}" data-k="spike" />
    <input type="number" min="1" value="${rmin}" data-k="row_min" />
    <input type="number" min="1" value="${rmax}" data-k="row_max" />
    <button class="rm" title="Remove">×</button>
  `;
  row.querySelector(".rm").onclick = () => row.remove();
  $("#mapEditor").appendChild(row);
}
$("#mapAddBtn").addEventListener("click", () => {
  // Default the new spike's range to start one above the highest existing row_max
  const rows = [...$("#mapEditor").querySelectorAll(".spike-row")].slice(1);
  let nextSpike = 1, nextStart = 1;
  rows.forEach(r => {
    const s = parseInt(r.querySelector('[data-k="spike"]').value, 10) || 0;
    const mx = parseInt(r.querySelector('[data-k="row_max"]').value, 10) || 0;
    if (s >= nextSpike) nextSpike = s + 1;
    if (mx >= nextStart) nextStart = mx + 1;
  });
  addSpikeRow(nextSpike, nextStart, nextStart + 9);
});
$("#mapSaveBtn").addEventListener("click", async () => {
  const meta = await idbGet("meta", "current");
  if (!meta) return alert("Pick a nursery first.");
  const rows = [...$("#mapEditor").querySelectorAll(".spike-row")].slice(1);
  const spikes = rows.map(r => ({
    spike: parseInt(r.querySelector('[data-k="spike"]').value, 10),
    row_min: parseInt(r.querySelector('[data-k="row_min"]').value, 10),
    row_max: parseInt(r.querySelector('[data-k="row_max"]').value, 10),
  })).filter(s => s.spike > 0 && s.row_min > 0 && s.row_max >= s.row_min);
  if (spikes.length === 0) return alert("Add at least one spike.");
  $("#mapStatus").innerHTML = '<span class="ok">Saving…</span>';
  try {
    const r = await fetch(`/nursery/${encodeURIComponent(meta.nursery_code)}/map`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ spikes }),
    });
    if (!r.ok) throw new Error(`HTTP ${r.status}: ${(await r.text()).slice(0,200)}`);
    const data = await r.json();
    $("#mapStatus").innerHTML =
      `<span class="ok">✓ Saved. ${data.packets_reassigned} packets reassigned across ` +
      `${data.spike_count} spike(s).` +
      (data.off_map > 0 ? ` ${data.off_map} packets are off-map (Spike 0).` : "") +
      `</span>`;
    await coachAutoDetect("map_saved");
    // Refresh the cached manifest so scans reflect the new spike numbers
    try {
      const m = await fetch(`/nursery/${encodeURIComponent(meta.nursery_code)}/manifest`);
      if (m.ok) {
        const md = await m.json();
        await idbDeleteByIndex("manifest", "by_nursery", meta.nursery_code);
        for (const p of md.packets) await idbPut("manifest", { ...p, nursery_code: meta.nursery_code });
      }
    } catch {}
  } catch (err) {
    $("#mapStatus").innerHTML = `<span class="err">${err.message}</span>`;
  }
});

// ====================================================================
// Workflow Coach — guides users through the 13-step nursery workflow
// ====================================================================
const COACH_STEPS = [
  // Phase 1 — Pre-field prep (Admin / breeder)
  { n: 1, phase: 1, title: "Initialise from PRISM export",
    body: "Open the Admin panel, paste the nursery code, choose the PRISM export .xlsx, and click Initialise. This creates the packet database, runs the Map parser and prepares the QR-coded packet PDF.",
    target: "#admInitHdr", openAdmin: true, autoDone: m => !!m && m.packet_count > 0 },
  { n: 2, phase: 1, title: "Build / review the Nursery list",
    body: "Click Refresh list in the Admin panel. You'll see every unique Source ID with Repeats and Qty Required (1.4 × reps). BULK-flagged rows have > 10 reps.",
    target: "#admListHdr", openAdmin: true },
  { n: 3, phase: 1, title: "Design the Field Map",
    body: "In the Field Map editor below, set how field rows map to spikes (e.g. Spike 1 = rows 1–10, Spike 2 = rows 11–26). Click Save to recompute Spike + Rack Order for every packet.",
    target: "#admMapHdr", openAdmin: true },
  { n: 4, phase: 1, title: "Print the Packet PDF",
    body: "Click Download Packet PDF in the Admin panel. Print the QR-coded labels (4×8 per A4 page) — one per packet — for the packeting team.",
    target: "#admDownloadsHdr", openAdmin: true },
  { n: 5, phase: 1, title: "Sort packets for racking",
    body: "Packets are already in LSD radix (rack ↑, spike ↑) order in the Packet prep tab of the downloaded workbook. Print or sort physically by rack order; that's the pickup sequence for racking.",
    target: "#admDownloadsHdr", openAdmin: true, manual: true },
  // Phase 2 — Field operations (Tech in the paddock)
  { n: 6, phase: 2, title: "Record replacements",
    body: "In the field: scan a packet QR (or find by plot), tap the 🔄 Replacement tile, fill stage + original source + new source, save. Works offline.",
    target: "#scanCard", manual: true },
  { n: 7, phase: 2, title: "Record planting errors",
    body: "Scan / find a packet, tap the ⚠️ Planting error tile, pick severity, note what happened, save. Use during planting whenever something doesn't go to plan.",
    target: "#scanCard", manual: true },
  { n: 8, phase: 2, title: "Record spray applications",
    body: "Scan / find a packet, tap 💧 Spray applied, pick product (TFMSA / IMI / HPPD), date, and rate. Run for each spray track.",
    target: "#scanCard", manual: true },
  { n: 9, phase: 2, title: "Record AB bag pulling",
    body: "Scan / find a packet, tap 📦 AB bag pulled, enter bag count and date. Run during weekly AB bag pull rounds.",
    target: "#scanCard", manual: true },
  // Phase 3 — Post-field & sync
  { n: 10, phase: 3, title: "Pull the updated PRISM export",
    body: "Once the breeder has updated PRISM with replacements + errors, download the refreshed Nursery site .xlsx from PRISM. Then come back here and re-run Step 1 with the new file. Everything refreshes.",
    target: "#admInitHdr", openAdmin: true, manual: true },
  { n: 11, phase: 3, title: "Generate the Fieldbook",
    body: "Click Download Fieldbook .xlsx in the Admin panel. You get a multi-tab workbook (Replacements, Planting errors, Spray log, AB bag pulling, Date recording, Notes) ready for printing or PRISM upload.",
    target: "#admDownloadsHdr", openAdmin: true },
  { n: 12, phase: 3, title: "Refresh the dashboard",
    body: "Open /dashboard in another tab to see all nurseries aggregated. It auto-refreshes every page hit. Use the All-techs feed below to see today's activity for the active nursery.",
    target: "#recentList" },
  { n: 13, phase: 3, title: "Push to Hub",
    body: "Already done — every event you record via the PWA is published instantly to the Hub. No manual push step. Cycle complete; you can now restart with a new nursery or move on to the next season.",
    target: "#recentList" },
];

const COACH_KEY_PREFIX = "coach:done:";

async function coachCurrentNursery() {
  const meta = await idbGet("meta", "current");
  return meta?.nursery_code || "_none_";
}

async function coachGetState() {
  const code = await coachCurrentNursery();
  const raw = localStorage.getItem(COACH_KEY_PREFIX + code);
  try { return raw ? JSON.parse(raw) : { current: 1, done: [] }; }
  catch { return { current: 1, done: [] }; }
}
async function coachSetState(s) {
  const code = await coachCurrentNursery();
  localStorage.setItem(COACH_KEY_PREFIX + code, JSON.stringify(s));
}

async function coachRender() {
  const state = await coachGetState();
  const bar = $("#coachBar");
  // Show coach only when a nursery is active.
  const meta = await idbGet("meta", "current");
  if (!meta?.nursery_code) { bar.classList.add("hidden"); return; }
  bar.classList.remove("hidden");

  const step = COACH_STEPS.find(s => s.n === state.current) || COACH_STEPS[0];
  $("#coachPhase").textContent = `Phase ${step.phase}`;
  $("#coachNum").textContent = `Step ${step.n} of ${COACH_STEPS.length}`;
  $("#coachTitle").textContent = step.title;
  $("#coachBody").textContent = step.body;
  const pct = ((state.done.length) / COACH_STEPS.length) * 100;
  $("#coachBarFill").style.width = pct + "%";

  $("#coachBack").disabled = state.current <= 1;
  $("#coachNext").disabled = state.current >= COACH_STEPS.length;

  // If stepper open, re-render too
  if (!$("#coachStepper").classList.contains("hidden")) coachRenderStepper(state);
}

function coachRenderStepper(state) {
  const root = $("#coachStepper");
  root.innerHTML = "";
  for (const s of COACH_STEPS) {
    const div = document.createElement("div");
    let cls = "step";
    if (state.done.includes(s.n)) cls += " done";
    if (s.n === state.current) cls += " current active";
    div.className = cls;
    const mark = state.done.includes(s.n) ? "✓" : s.n;
    div.innerHTML = `<span class="step-marker">${mark}</span>` +
                    `<div class="step-body"><div class="step-title">${s.title}</div>` +
                    `<div class="step-hint">Phase ${s.phase}</div></div>`;
    div.onclick = () => coachJump(s.n);
    root.appendChild(div);
  }
}

async function coachJump(stepNo) {
  const state = await coachGetState();
  state.current = stepNo;
  await coachSetState(state);
  await coachRender();
  await coachScrollToTarget(stepNo);
}

async function coachScrollToTarget(stepNo) {
  const step = COACH_STEPS.find(s => s.n === stepNo);
  if (!step) return;
  // Open the admin panel if the step lives there
  if (step.openAdmin) {
    const btn = $("#admToggle");
    if (btn && btn.getAttribute("aria-expanded") !== "true") btn.click();
  }
  if (!step.target) return;
  const el = document.querySelector(step.target);
  if (!el) return;
  // Wait a frame for the admin body to expand
  await new Promise(r => setTimeout(r, 60));
  el.scrollIntoView({ behavior: "smooth", block: "start" });
  const host = el.closest(".card") || el;
  host.classList.remove("coach-glow");
  // restart animation
  // eslint-disable-next-line no-unused-expressions
  void host.offsetWidth;
  host.classList.add("coach-glow");
}

async function coachAdvance(direction = +1) {
  const state = await coachGetState();
  if (direction > 0 && state.current < COACH_STEPS.length) state.current += 1;
  if (direction < 0 && state.current > 1) state.current -= 1;
  await coachSetState(state);
  await coachRender();
  await coachScrollToTarget(state.current);
}

async function coachMarkDone() {
  const state = await coachGetState();
  if (!state.done.includes(state.current)) state.done.push(state.current);
  if (state.current < COACH_STEPS.length) state.current += 1;
  await coachSetState(state);
  await coachRender();
  await coachScrollToTarget(state.current);
}

// Wire buttons
$("#coachBack").addEventListener("click", () => coachAdvance(-1));
$("#coachNext").addEventListener("click", () => coachAdvance(+1));
$("#coachDone").addEventListener("click", coachMarkDone);
$("#coachExpand").addEventListener("click", async () => {
  const stepper = $("#coachStepper");
  const open = !stepper.classList.contains("hidden");
  stepper.classList.toggle("hidden", open);
  $("#coachExpand").textContent = open ? "All steps ▾" : "All steps ▴";
  if (!open) coachRenderStepper(await coachGetState());
});

// Auto-detect completion on certain user actions (best-effort)
async function coachAutoDetect(actionType) {
  const state = await coachGetState();
  const map = {
    init_success: 1,
    list_loaded: 2,
    map_saved: 3,
    pdf_downloaded: 4,
    event_replacement: 6,
    event_planting_error: 7,
    event_spray: 8,
    event_ab_pull: 9,
    fieldbook_downloaded: 11,
    sync_success: 13,
  };
  const stepNo = map[actionType];
  if (!stepNo) return;
  if (!state.done.includes(stepNo)) state.done.push(stepNo);
  // If this is the current step or earlier, advance the cursor.
  if (state.current <= stepNo && state.current < COACH_STEPS.length) {
    state.current = stepNo + 1;
  }
  await coachSetState(state);
  await coachRender();
}

// ---------- Boot ----------
if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("/static/sw.js").catch(() => {});
}
setOnline(navigator.onLine);
loadNurseryList();
refreshOutboxCount();
renderRecent();
