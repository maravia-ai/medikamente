"use strict";

/* Medikamenten-Inventur – Web-App mit Excel (OneDrive) als Datenquelle.
 * Berechnung identisch zur Excel-Tabelle "Inventur":
 *   reicht bis  = gezählt am + Menge / (morgens + abends)
 *   heute übrig = Menge - (heute - gezählt am) * (morgens + abends)
 * Einlagern schreibt neue "Menge" (= heutiger Rest + Zugang) und "gezählt am" (= heute).
 */

const C = window.APP_CONFIG;
const SCOPES = ["Files.ReadWrite", "Calendars.ReadWrite"];
const GRAPH = "https://graph.microsoft.com/v1.0";
// Kennzeichnung der von der App angelegten Kalendertermine (eigene, feste GUID)
const EXT_PROP = "String {6f2c1b3e-8d4a-4c7e-9a51-2b7d3e9f0a14} Name MedInventurErinnerung";
const DEMO = new URLSearchParams(location.search).has("demo");
const LS_HIDDEN = "medapp.hidden";
const LS_ITEM = "medapp.itemId";
const LS_PACK = "medapp.packungsgroesse";
let priceColumnOk = false;

// Spalten im Blatt "Inventur" (0-basiert): A..K
const COL = { name: 0, dosis: 1, morgens: 2, abends: 3, menge: 4, gezaehlt: 5, preis: 8, notiz: 10 };

let pca = null;
let account = null;
let state = { sections: [], meds: [], loadedAt: null, error: null };

const $app = document.getElementById("app");
const $backdrop = document.getElementById("backdrop");
const $sheet = document.getElementById("sheet");
const $toast = document.getElementById("toast");

/* ---------- Hilfsfunktionen: Datum (Excel-Seriennummern) ---------- */

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86400000;

function todaySerial() {
  const n = new Date();
  return Math.round((Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()) - EXCEL_EPOCH) / DAY);
}
function serialToUTCDate(s) { return new Date(EXCEL_EPOCH + Math.floor(s) * DAY); }
function fmtDate(s) {
  return serialToUTCDate(s).toLocaleDateString("de-DE", { timeZone: "UTC", day: "2-digit", month: "2-digit", year: "numeric" });
}
function fmtDateLong(s) {
  return serialToUTCDate(s).toLocaleDateString("de-DE", { timeZone: "UTC", weekday: "short", day: "numeric", month: "long", year: "numeric" });
}
function isoDate(s) { return serialToUTCDate(s).toISOString().slice(0, 10); }
function fmtEuro(v) { return v.toLocaleString("de-DE", { style: "currency", currency: "EUR", minimumFractionDigits: 2, maximumFractionDigits: 4 }); }
function esc(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function plural(n, one, many) { return `${n} ${Math.abs(n) === 1 ? one : many}`; }

function getHidden() {
  try { return new Set(JSON.parse(localStorage.getItem(LS_HIDDEN) || "[]")); } catch { return new Set(); }
}
function setHidden(set) {
  try { localStorage.setItem(LS_HIDDEN, JSON.stringify([...set])); } catch { /* ignorieren */ }
}

function toast(msg, ms = 2600) {
  $toast.textContent = msg;
  $toast.classList.add("show");
  clearTimeout(toast.t);
  toast.t = setTimeout(() => $toast.classList.remove("show"), ms);
}

/* ---------- Berechnung ---------- */

function compute(med) {
  const T = todaySerial();
  const daily = med.morgens + med.abends;
  if (daily <= 0) {
    return { ...med, daily, rest: med.menge, reichtBis: null, tageUebrig: Infinity, status: "off" };
  }
  const rest = Math.round(med.menge - (T - med.gezaehlt) * daily);
  const reichtBis = med.gezaehlt + med.menge / daily;
  const tageUebrig = Math.floor(reichtBis) - T;
  let status = "ok";
  if (tageUebrig <= 0) status = "bad";
  else if (tageUebrig <= C.vorlaufTage) status = "warn";
  return { ...med, daily, rest: Math.max(0, rest), restRaw: rest, reichtBis: Math.floor(reichtBis), tageUebrig, status };
}

/* ---------- Tabelle parsen ---------- */

function parseSheet(values, startRow) {
  const headerIdx = values.findIndex(r => String(r[COL.name]).trim() === "Medikament");
  if (headerIdx < 0) throw new Error(`Kopfzeile "Medikament" im Blatt "${C.sheetName}" nicht gefunden.`);
  const h = values[headerIdx].map(v => String(v).trim().toLowerCase());
  if (!h[COL.menge].startsWith("menge") || !h[COL.gezaehlt].startsWith("gezählt")) {
    throw new Error(`Spaltenaufbau unerwartet. Erwartet: Spalte E = "Menge", Spalte F = "gezählt am".\nGefunden: E = "${values[headerIdx][COL.menge]}", F = "${values[headerIdx][COL.gezaehlt]}"`);
  }
  const sections = [];
  let current = { title: "", meds: [] };
  sections.push(current);
  for (let i = headerIdx + 1; i < values.length; i++) {
    const r = values[i];
    const name = String(r[COL.name] ?? "").trim();
    if (!name) continue;
    const menge = r[COL.menge], gez = r[COL.gezaehlt];
    if (typeof menge === "number" && typeof gez === "number") {
      current.meds.push({
        row: startRow + i,
        name,
        dosis: String(r[COL.dosis] ?? "").trim(),
        morgens: Number(r[COL.morgens]) || 0,
        abends: Number(r[COL.abends]) || 0,
        menge,
        gezaehlt: gez,
        preis: typeof r[COL.preis] === "number" ? r[COL.preis] : null,
        notiz: String(r[COL.notiz] ?? "").trim()
      });
    } else if (menge === "" && gez === "") {
      current = { title: name, meds: [] };
      sections.push(current);
    }
  }
  priceColumnOk = String(h[COL.preis] ?? "").startsWith("preis");
  return sections.filter(s => s.meds.length);
}

/* ---------- Microsoft Graph ---------- */

async function getToken() {
  if (!account) throw new Error("Nicht angemeldet.");
  try {
    const r = await pca.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError || e.errorCode === "no_tokens_found") {
      await pca.acquireTokenRedirect({ scopes: SCOPES, account });
      return new Promise(() => {}); // Seite wird umgeleitet
    }
    throw e;
  }
}

async function graph(method, path, body) {
  const token = await getToken();
  const res = await fetch(GRAPH + path, {
    method,
    headers: { Authorization: "Bearer " + token, ...(body ? { "Content-Type": "application/json" } : {}) },
    body: body ? JSON.stringify(body) : undefined
  });
  if (res.status === 204) return null;
  const data = await res.json().catch(() => null);
  if (!res.ok) {
    const msg = data?.error?.message || res.statusText;
    const err = new Error(`Microsoft Graph (${res.status}): ${msg}`);
    err.status = res.status;
    throw err;
  }
  return data;
}

function sheetPath(itemId) {
  const ws = encodeURIComponent(C.sheetName.replace(/'/g, "''"));
  return `/me/drive/items/${itemId}/workbook/worksheets('${ws}')`;
}

async function getItemId(force) {
  let id = null;
  try { id = force ? null : localStorage.getItem(LS_ITEM); } catch { /* ignorieren */ }
  if (id) return id;
  const p = C.filePath.split("/").map(encodeURIComponent).join("/");
  const item = await graph("GET", `/me/drive/root:/${p}?$select=id,name`);
  try { localStorage.setItem(LS_ITEM, item.id); } catch { /* ignorieren */ }
  return item.id;
}

async function readUsedRange() {
  let id = await getItemId(false);
  let data;
  try {
    data = await graph("GET", `${sheetPath(id)}/usedRange?$select=address,values`);
  } catch (e) {
    if (e.status !== 404) throw e;
    id = await getItemId(true); // Datei evtl. verschoben/ersetzt
    data = await graph("GET", `${sheetPath(id)}/usedRange?$select=address,values`);
  }
  const m = /!\$?[A-Z]+\$?(\d+)/.exec(data.address);
  const startRow = m ? Number(m[1]) : 1;
  return { id, values: data.values, startRow };
}

async function readRow(id, row) {
  const data = await graph("GET", `${sheetPath(id)}/range(address='A${row}:F${row}')?$select=values`);
  return data.values[0];
}

async function writePrice(id, row, preis) {
  const col = String.fromCharCode(65 + COL.preis);
  await graph("PATCH", `${sheetPath(id)}/range(address='${col}${row}')`, { values: [[preis]] });
}

function getPackSize(name) {
  try { return JSON.parse(localStorage.getItem(LS_PACK) || "{}")[name] || null; } catch { return null; }
}
function setPackSize(name, n) {
  try {
    const all = JSON.parse(localStorage.getItem(LS_PACK) || "{}");
    all[name] = n;
    localStorage.setItem(LS_PACK, JSON.stringify(all));
  } catch { /* ignorieren */ }
}
function parseDecimal(s) {
  const t = String(s).trim().replace(/\s|€/g, "").replace(",", ".");
  if (t === "") return null;
  const v = Number(t);
  return Number.isFinite(v) && v >= 0 ? v : NaN;
}

async function writeStock(id, row, menge, serial) {
  await graph("PATCH", `${sheetPath(id)}/range(address='E${row}:F${row}')`, { values: [[menge, serial]] });
}

/* ---------- Kalender (Outlook / Microsoft 365) ---------- */

function reminderSerial(m) {
  const T = todaySerial();
  return Math.max(m.reichtBis - C.vorlaufTage, T + 1);
}

async function syncReminder(m, hidden) {
  const val = m.name.replace(/'/g, "''");
  const filter = `singleValueExtendedProperties/Any(ep: ep/id eq '${EXT_PROP}' and ep/value eq '${val}')`;
  const existing = await graph("GET", `/me/events?$filter=${encodeURIComponent(filter)}&$select=id&$top=50`);
  for (const ev of existing.value || []) {
    await graph("DELETE", `/me/events/${ev.id}`);
  }
  if (hidden.has(m.name) || m.daily <= 0) return false;

  const s = reminderSerial(m);
  const [hh, mm] = C.erinnerungUhrzeit.split(":").map(Number);
  const start = `${isoDate(s)}T${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}:00`;
  const endMin = hh * 60 + mm + 15;
  const end = `${isoDate(s)}T${String(Math.floor(endMin / 60)).padStart(2, "0")}:${String(endMin % 60).padStart(2, "0")}:00`;
  const restAmTag = Math.max(0, Math.round(m.menge - (s - m.gezaehlt) * m.daily));

  await graph("POST", "/me/events", {
    subject: `Nachbestellen: ${m.name}${m.dosis ? " " + m.dosis : ""}`,
    body: {
      contentType: "text",
      content: `Vorrat reicht bis ${fmtDate(m.reichtBis)}.\nVoraussichtlicher Bestand an diesem Tag: ${restAmTag} Stück.\nEinnahme: ${m.morgens}-${m.abends} (${m.daily} pro Tag).\n\nAutomatisch erstellt von der Medikamenten-App.`
    },
    start: { dateTime: start, timeZone: C.zeitzone },
    end: { dateTime: end, timeZone: C.zeitzone },
    isReminderOn: true,
    reminderMinutesBeforeStart: 0,
    showAs: "free",
    singleValueExtendedProperties: [{ id: EXT_PROP, value: m.name }]
  });
  return true;
}

async function syncAllReminders() {
  const hidden = getHidden();
  let n = 0;
  for (const m of state.meds) {
    if (await syncReminder(m, hidden)) n++;
  }
  return n;
}

/* ---------- Demo-Modus (ohne Microsoft-Konto, Daten nur im Speicher) ---------- */

const demo = {
  startRow: 1,
  values: [
    ["Medikament", "Dosis", "morgens", "abends", "Menge", "gezählt am", "reicht bis", "Anzahl noch übrig", "Preis/Stck", "Wert", ""],
    ["Ohne Rezept (Medikationsplan)", "", "", "", "", "", "", "", "", "", ""],
    ["Tromcardin complex", "", 1, 1, 179, 46267, "", "", 0.18, "", ""],
    ["Vigantol", "1000 I.E.", 1, 0, 150, 46249, "", "", 0.06, "", ""],
    ["Mit Rezept", "", "", "", "", "", "", "", "", "", ""],
    ["ASS", "100mg", 1, 0, 106, 46304, "", "", 0.03, "", "Blutverdünner"],
    ["Valsartan", "160mg", 1, 0, 125, 46304, "", "", 0.23, "", "Senkung Blutdruck"],
    ["Prasugrel", "10mg", 1, 0, 84, 45519, "", "", 1.16, "", "letzte Einnahme 08.09.2024"],
    ["Amlodipin", "7.5mg", 1, 0, 105, 46304, "", "", 0.15, "", "Senkung Blutdruck"],
    ["Bisoprolol", "5mg", 1, 0, 98, 46249, "", "", 0.14, "", "Senkung Blutdruck"],
    ["Jardiance (Empagliflozin)", "10 mg", 1, 0, 120, 46304, "", "", 2.44, "", "Senkung Blutzucker"],
    ["Nilemdo", "180mg", 1, 0, 41, 46304, "", "", 2.94, "", "Senkung Cholesterin"],
    ["Ezetimib", "10mg", 0, 1, 100, 46304, "", "", 0.30, "", "Senkung Cholesterin"],
    ["Rosuvastatin", "40mg", 0, 1, 130, 46249, "", "", 0.32, "", "Senkung Cholesterin"],
    ["", "", "", "", "", "", "", "", 7.95, 398.01, ""]
  ]
};

/* ---------- Laden ---------- */

async function load() {
  state.error = null;
  render();
  try {
    let values, startRow, id = null;
    if (DEMO) ({ values, startRow } = demo);
    else ({ id, values, startRow } = await readUsedRange());
    state.itemId = id;
    const sections = parseSheet(values, startRow).map(s => ({ ...s, meds: s.meds.map(compute) }));
    state.sections = sections;
    state.meds = sections.flatMap(s => s.meds);
    state.loadedAt = new Date();
  } catch (e) {
    console.error(e);
    state.error = e.message || String(e);
  }
  render();
}

/* ---------- Darstellung ---------- */

function rowHtml(m) {
  let right, badge = "";
  if (m.status === "off") right = `<div class="days">pausiert</div><div class="meta">${m.rest} Stk.</div>`;
  else {
    const days = m.tageUebrig <= 0 ? "aufgebraucht" : plural(m.tageUebrig, "Tag", "Tage");
    right = `<div class="days">${days}</div><div class="meta">bis ${fmtDate(m.reichtBis)}</div>`;
    if (m.status === "warn") badge = `<span class="badge">nachbestellen</span>`;
    if (m.status === "bad") badge = `<span class="badge">leer</span>`;
  }
  const pct = m.status === "off" ? 0 : Math.max(0, Math.min(100, (m.tageUebrig / 120) * 100));
  return `<button class="row s-${m.status}" data-name="${esc(m.name)}">
    <div><div class="name">${esc(m.name)}${badge}</div>
      <div class="meta">${esc(m.dosis || "–")} · ${m.morgens}-${m.abends} · ${m.rest} Stk. vorrätig</div></div>
    <div class="right">${right}</div>
    <div class="bar"><i style="width:${pct}%"></i></div>
  </button>`;
}

function render() {
  if (!DEMO && (!C.clientId || C.clientId.startsWith("HIER"))) {
    $app.innerHTML = `<div class="center"><h1>Einrichtung fehlt</h1>
      <p>Bitte in <b>config.js</b> die Client-ID und Tenant-ID aus der Entra-App-Registrierung eintragen.</p>
      <p><a href="?demo">Demo ansehen</a></p></div>`;
    return;
  }
  if (!DEMO && !account) {
    $app.innerHTML = `<div class="center"><h1>Medikamente</h1>
      <p>Anmelden, um die Inventur aus deiner Excel-Datei zu laden.</p>
      <div class="actions"><button class="btn primary" id="login">Mit Microsoft anmelden</button></div></div>`;
    document.getElementById("login").onclick = () => pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
    return;
  }

  const hidden = getHidden();
  const visible = state.meds.filter(m => !hidden.has(m.name));
  const nWarn = visible.filter(m => m.status === "warn").length;
  const nBad = visible.filter(m => m.status === "bad").length;
  const wert = visible.reduce((a, m) => a + (m.preis ? m.preis * m.rest : 0), 0);

  let html = `<header><div><h1>Medikamente</h1>
    <div class="sub">${DEMO ? "Demo-Daten · " : ""}${state.loadedAt ? "Stand " + state.loadedAt.toLocaleString("de-DE", { dateStyle: "short", timeStyle: "short" }) : "Lade ..."}</div></div>
    <button class="iconbtn" id="reload">Aktualisieren</button></header>`;

  if (state.error) html += `<div class="error">${esc(state.error)}</div>`;

  if (state.meds.length) {
    html += `<div class="summary">
      <div class="tile ${nBad ? "bad" : "ok"}"><b>${nBad}</b><span>aufgebraucht</span></div>
      <div class="tile ${nWarn ? "warn" : "ok"}"><b>${nWarn}</b><span>nachbestellen (≤ ${C.vorlaufTage} T.)</span></div>
      <div class="tile"><b>${Math.round(wert)} €</b><span>Vorratswert</span></div></div>`;

    const urgent = visible.filter(m => m.status === "warn" || m.status === "bad").sort((a, b) => a.tageUebrig - b.tageUebrig);
    if (urgent.length) html += `<h2>Bald nachbestellen</h2><div class="list">${urgent.map(rowHtml).join("")}</div>`;

    for (const s of state.sections) {
      const ms = s.meds.filter(m => !hidden.has(m.name)).sort((a, b) => a.tageUebrig - b.tageUebrig);
      if (ms.length) html += `<h2>${esc(s.title || "Medikamente")}</h2><div class="list">${ms.map(rowHtml).join("")}</div>`;
    }

    const hid = state.meds.filter(m => hidden.has(m.name));
    if (hid.length) html += `<details><summary>Ausgeblendet (${hid.length})</summary><div class="list">${hid.map(rowHtml).join("")}</div></details>`;

    html += `<div class="actions">
      <button class="btn" id="sync">Kalender-Erinnerungen aktualisieren</button>
      ${DEMO ? "" : `<button class="btn danger" id="logout">Abmelden</button>`}
    </div>
    <p class="sub" style="margin:12px 4px">Erinnerung jeweils ${C.vorlaufTage} Tage vor Ende des Vorrats, ${C.erinnerungUhrzeit} Uhr, im Microsoft-365-Kalender.</p>`;
  } else if (!state.error) {
    html += `<div class="center">Lade Inventur ...</div>`;
  }

  $app.innerHTML = html;
  document.getElementById("reload").onclick = load;
  $app.querySelectorAll(".row").forEach(b => b.onclick = () => openDetail(b.dataset.name));
  const sync = document.getElementById("sync");
  if (sync) sync.onclick = async () => {
    if (DEMO) return toast("Im Demo-Modus nicht verfügbar.");
    sync.disabled = true; sync.textContent = "Aktualisiere Kalender ...";
    try { const n = await syncAllReminders(); toast(`${plural(n, "Erinnerung", "Erinnerungen")} im Kalender eingetragen.`); }
    catch (e) { state.error = e.message; render(); }
    finally { sync.disabled = false; sync.textContent = "Kalender-Erinnerungen aktualisieren"; }
  };
  const logout = document.getElementById("logout");
  if (logout) logout.onclick = () => pca.logoutRedirect({ account });
}

/* ---------- Detailansicht / Einlagern ---------- */

function closeDetail() { $backdrop.classList.remove("open"); }
$backdrop.addEventListener("click", e => { if (e.target === $backdrop) closeDetail(); });

function openDetail(name) {
  const m = state.meds.find(x => x.name === name);
  if (!m) return;
  const hidden = getHidden();
  let mode = "add";

  $sheet.innerHTML = `<div class="grab"></div>
    <h3>${esc(m.name)}</h3><div class="sub">${esc(m.dosis || "")}${m.notiz ? " · " + esc(m.notiz) : ""}</div>
    <div class="kv">
      <div><span>Bestand heute</span><span>${m.rest} Stück</span></div>
      <div><span>Einnahme</span><span>${m.morgens}-${m.abends} (${m.daily} pro Tag)</span></div>
      <div><span>Reicht bis</span><span>${m.reichtBis ? fmtDateLong(m.reichtBis) : "–"}</span></div>
      <div><span>Erinnerung am</span><span>${m.reichtBis ? fmtDate(reminderSerial(m)) : "–"}</span></div>
      <div><span>Zuletzt gezählt</span><span>${fmtDate(m.gezaehlt)} (${m.menge} Stk.)</span></div>
      ${m.preis != null ? `<div><span>Preis pro Stück</span><span>${fmtEuro(m.preis)}</span></div>` : ""}
    </div>
    <div class="seg"><button data-mode="add" class="on">Eingelagert</button><button data-mode="set">Gezählt</button></div>
    <label class="field" id="lbl" for="qty">Neu hinzugekommene Tabletten</label>
    <input type="number" id="qty" inputmode="numeric" min="0" step="1" placeholder="z. B. 100">
    <div class="preview" id="pv"></div>
    <div id="pricebox"${priceColumnOk ? "" : ' style="display:none"'}>
      <label class="field">Preis aktualisieren (optional)</label>
      <div class="pricegrid">
        <div><label class="field" for="ppack">Packungspreis €</label>
          <input type="text" id="ppack" inputmode="decimal" autocomplete="off" placeholder="z. B. 14,30"></div>
        <div><label class="field" for="psize">Stück pro Packung</label>
          <input type="number" id="psize" inputmode="numeric" min="1" step="1" placeholder="z. B. 100"></div>
      </div>
      <div class="preview" id="ppv"></div>
    </div>
    <div class="actions" style="margin-top:0">
      <button class="btn primary" id="save" disabled>Speichern</button>
      <button class="btn" id="hide">${hidden.has(m.name) ? "Wieder einblenden" : "Ausblenden (nicht mehr eingenommen)"}</button>
      <button class="btn" id="close">Schließen</button>
    </div>`;

  const $qty = $sheet.querySelector("#qty"), $pv = $sheet.querySelector("#pv"), $save = $sheet.querySelector("#save");
  const $pbox = $sheet.querySelector("#pricebox"), $ppack = $sheet.querySelector("#ppack");
  const $psize = $sheet.querySelector("#psize"), $ppv = $sheet.querySelector("#ppv");
  const savedPack = getPackSize(m.name);
  let psizeTouched = false;
  if (savedPack) { $psize.value = savedPack; psizeTouched = true; }

  function newMenge() {
    const v = Number($qty.value);
    if ($qty.value === "" || !Number.isFinite(v) || v < 0) return null;
    return mode === "add" ? m.rest + Math.round(v) : Math.round(v);
  }
  // Ergebnis: null = kein Preis eingegeben, NaN = ungültig, sonst Preis pro Stück
  function newPrice() {
    if (mode !== "add" || !priceColumnOk) return null;
    const p = parseDecimal($ppack.value);
    if (p === null) return null;
    const s = Number($psize.value);
    if (Number.isNaN(p) || !Number.isInteger(s) || s < 1) return NaN;
    return Math.round((p / s) * 10000) / 10000;
  }
  function update() {
    if (mode === "add" && !psizeTouched && $qty.value !== "") $psize.value = Math.round(Number($qty.value)) || "";
    const price = newPrice();
    if (price === null) $ppv.textContent = m.preis != null ? `Bisher: ${fmtEuro(m.preis)} pro Stück` : "";
    else if (Number.isNaN(price)) $ppv.textContent = "Bitte gültigen Packungspreis und Stückzahl eingeben.";
    else {
      const diff = m.preis != null ? ` (bisher ${fmtEuro(m.preis)})` : "";
      $ppv.textContent = `Neuer Preis: ${price.toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 4 })} € pro Stück${diff}`;
    }
    const n = newMenge();
    $save.disabled = n === null || (mode === "add" && Number($qty.value) === 0) || Number.isNaN(price);
    if (n === null) { $pv.textContent = ""; return; }
    if (m.daily > 0) {
      const bis = todaySerial() + Math.floor(n / m.daily);
      $pv.textContent = `Neuer Bestand: ${n} Stück · reicht bis ${fmtDate(bis)}`;
    } else $pv.textContent = `Neuer Bestand: ${n} Stück`;
  }
  $sheet.querySelectorAll(".seg button").forEach(b => b.onclick = () => {
    mode = b.dataset.mode;
    $sheet.querySelectorAll(".seg button").forEach(x => x.classList.toggle("on", x === b));
    $sheet.querySelector("#lbl").textContent = mode === "add" ? "Neu hinzugekommene Tabletten" : "Tatsächlich gezählter Bestand";
    $pbox.style.display = mode === "add" && priceColumnOk ? "" : "none";
    update();
  });
  $qty.oninput = update;
  $ppack.oninput = update;
  $psize.oninput = () => { psizeTouched = $psize.value !== ""; update(); };
  update();
  $sheet.querySelector("#close").onclick = closeDetail;

  $sheet.querySelector("#hide").onclick = async () => {
    const h = getHidden();
    h.has(m.name) ? h.delete(m.name) : h.add(m.name);
    setHidden(h);
    closeDetail();
    render();
    if (!DEMO) {
      try { await syncReminder(m, h); } catch (e) { toast("Kalender: " + e.message, 5000); }
    }
  };

  $save.onclick = async () => {
    $save.disabled = true; $save.textContent = "Speichere ...";
    try {
      const T = todaySerial();
      let menge = newMenge();
      const price = newPrice();
      if (DEMO) {
        const r = demo.values[m.row - demo.startRow];
        r[COL.menge] = menge; r[COL.gezaehlt] = T;
        if (price !== null) r[COL.preis] = price;
      } else {
        // Zeile vor dem Schreiben frisch lesen und prüfen, ob es noch dasselbe Medikament ist
        const cur = await readRow(state.itemId, m.row);
        if (String(cur[COL.name]).trim() !== m.name) {
          throw new Error("Die Tabelle wurde zwischenzeitlich verändert. Bitte aktualisieren und erneut versuchen.");
        }
        if (mode === "add") {
          const fresh = compute({ ...m, menge: cur[COL.menge], gezaehlt: cur[COL.gezaehlt] });
          menge = fresh.rest + Math.round(Number($qty.value));
        }
        await writeStock(state.itemId, m.row, menge, T);
        if (price !== null) await writePrice(state.itemId, m.row, price);
      }
      if (price !== null) setPackSize(m.name, Number($psize.value));
      closeDetail();
      await load();
      toast(`${m.name}: Bestand ${menge} Stück${price !== null ? " und Preis" : ""} gespeichert.`);
      if (!DEMO) {
        const updated = state.meds.find(x => x.name === m.name);
        if (updated) {
          try { await syncReminder(updated, getHidden()); }
          catch (e) { toast("Gespeichert, aber Kalender-Fehler: " + e.message, 5000); }
        }
      }
    } catch (e) {
      $save.disabled = false; $save.textContent = "Speichern";
      $pv.textContent = e.message;
    }
  };

  $backdrop.classList.add("open");
}

/* ---------- Start ---------- */

async function start() {
  if (DEMO) { render(); return load(); }
  if (!C.clientId || C.clientId.startsWith("HIER")) return render();
  pca = new msal.PublicClientApplication({
    auth: {
      clientId: C.clientId,
      authority: `https://login.microsoftonline.com/${C.tenantId}`,
      redirectUri: location.origin + location.pathname
    },
    cache: { cacheLocation: "localStorage" }
  });
  try {
    await pca.initialize();
    const res = await pca.handleRedirectPromise();
    account = res?.account || pca.getAllAccounts()[0] || null;
  } catch (e) {
    $app.innerHTML = `<div class="error">Anmeldung fehlgeschlagen:\n${esc(e.message || e)}</div>`;
    return;
  }
  render();
  if (account) load();
}

start();
