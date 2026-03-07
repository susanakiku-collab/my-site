/* ======================================================
THEMIS AI Dispatch
Part1
Core / Geo / Maps / State
====================================================== */

const STORAGE_KEY = "themis_dispatch_v3";
const GEO_CACHE_KEY = "themis_geo_cache";

/* ======================================================
BASIC
====================================================== */

function $(id){
  return document.getElementById(id);
}

function esc(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

function uid(prefix="id"){
  return prefix + "_" + Date.now().toString(36) + Math.random().toString(36).slice(2);
}

function toNum(v){
  if(v === null || v === undefined || v === "") return 0;
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function toNullableNum(v){
  if(v === null || v === undefined || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function todayStr(){
  return new Date().toISOString().slice(0,10);
}

/* ======================================================
GOOGLE MAP ROUTE
====================================================== */

function openRouteFromMatsudo(address){
  const origin = encodeURIComponent("松戸駅");
  const dest = encodeURIComponent(address);
  const url = `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${dest}&travelmode=driving`;
  window.open(url, "_blank");
}

/* ======================================================
MATSUDO GEO
====================================================== */

const MATSUDO_STATION = {
  lat: 35.7846,
  lon: 139.9003
};

function deg2rad(d){
  return d * Math.PI / 180;
}

function haversineKm(lat1, lon1, lat2, lon2){
  const R = 6371;
  const dLat = deg2rad(lat2 - lat1);
  const dLon = deg2rad(lon2 - lon1);

  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(deg2rad(lat1)) *
    Math.cos(deg2rad(lat2)) *
    Math.sin(dLon / 2) ** 2;

  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

/* ======================================================
FREE GEOCODE
====================================================== */

function loadGeoCache(){
  try{
    return JSON.parse(localStorage.getItem(GEO_CACHE_KEY) || "{}");
  }catch{
    return {};
  }
}

function saveGeoCache(cache){
  localStorage.setItem(GEO_CACHE_KEY, JSON.stringify(cache));
}

const geoCache = loadGeoCache();
let lastGeoRequest = 0;

async function geocodeAddress(address){
  const q = String(address || "").trim();
  if(!q) return null;

  if(geoCache[q]){
    return geoCache[q];
  }

  const now = Date.now();
  const wait = Math.max(0, 1100 - (now - lastGeoRequest));
  if(wait > 0){
    await new Promise(r => setTimeout(r, wait));
  }

  lastGeoRequest = Date.now();

  const url =
    "https://nominatim.openstreetmap.org/search?" +
    new URLSearchParams({
      q,
      format: "jsonv2",
      limit: "1",
      countrycodes: "jp"
    });

  const res = await fetch(url, {
    headers: { "Accept": "application/json" }
  });

  if(!res.ok){
    throw new Error("geocode failed");
  }

  const data = await res.json();
  if(!Array.isArray(data) || !data.length){
    return null;
  }

  const geo = {
    lat: Number(data[0].lat),
    lon: Number(data[0].lon)
  };

  geoCache[q] = geo;
  saveGeoCache(geoCache);

  return geo;
}

/* ======================================================
STATE
====================================================== */

function createDefaultState(){
  return {
    casts: [],
    vehicles: [],
    plans: [],
    actuals: [],
    dispatchResult: {},
    activeVehicleIds: [],
    settings: {
      balanceStrength: 1,
      lastAreaPenalty: 10
    }
  };
}

function normalizeState(raw){
  const s = raw && typeof raw === "object" ? raw : createDefaultState();

  s.casts = Array.isArray(s.casts) ? s.casts : [];
  s.vehicles = Array.isArray(s.vehicles) ? s.vehicles : [];
  s.plans = Array.isArray(s.plans) ? s.plans : [];
  s.actuals = Array.isArray(s.actuals) ? s.actuals : [];
  s.dispatchResult = s.dispatchResult && typeof s.dispatchResult === "object" ? s.dispatchResult : {};
  s.activeVehicleIds = Array.isArray(s.activeVehicleIds) ? s.activeVehicleIds : [];
  s.settings = s.settings && typeof s.settings === "object" ? s.settings : {};

  if(typeof s.settings.balanceStrength !== "number"){
    s.settings.balanceStrength = 1;
  }
  if(typeof s.settings.lastAreaPenalty !== "number"){
    s.settings.lastAreaPenalty = 10;
  }

  s.casts = s.casts.map(c => ({
    id: String(c.id || uid("cast")),
    name: String(c.name || "").trim(),
    address: String(c.address || "").trim(),
    memo: String(c.memo || "").trim(),
    area: String(c.area || "").trim(),
    km: toNum(c.km),
    lat: c.lat === null || c.lat === undefined || c.lat === "" ? null : Number(c.lat),
    lon: c.lon === null || c.lon === undefined || c.lon === "" ? null : Number(c.lon)
  }));

  s.vehicles = s.vehicles.map(v => ({
    id: String(v.id || "").trim().toUpperCase(),
    area: String(v.area || "").trim(),
    homeArea: String(v.homeArea || v.area || "").trim(),
    capacity: Math.max(4, Math.floor(toNum(v.capacity) || 4)),
    driver: String(v.driver || "").trim(),
    monthKm: toNum(v.monthKm),
    workDays: Math.max(0, Math.floor(toNum(v.workDays)))
  })).filter(v => v.id);

  s.plans = s.plans.map(p => ({
    id: String(p.id || uid("plan")),
    date: String(p.date || ""),
    castId: String(p.castId || ""),
    hour: Math.max(0, Math.min(5, Math.floor(toNum(p.hour)))),
    address: String(p.address || "").trim(),
    area: String(p.area || "").trim(),
    memo: String(p.memo || "").trim(),
    kmOverride: toNum(p.kmOverride),
    status: String(p.status || "active")
  }));

  s.actuals = s.actuals.map(a => ({
    id: String(a.id || uid("actual")),
    date: String(a.date || ""),
    castId: String(a.castId || ""),
    hour: Math.max(0, Math.min(5, Math.floor(toNum(a.hour)))),
    address: String(a.address || "").trim(),
    area: String(a.area || "").trim(),
    memo: String(a.memo || "").trim(),
    kmOverride: toNum(a.kmOverride),
    status: ["active", "done", "canceled"].includes(a.status) ? a.status : "active"
  }));

  return s;
}

function loadState(){
  try{
    const raw = localStorage.getItem(STORAGE_KEY);
    if(!raw) return createDefaultState();
    return normalizeState(JSON.parse(raw));
  }catch{
    return createDefaultState();
  }
}

function saveState(){
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

const state = loadState();
/* ======================================================
THEMIS AI Dispatch
Part2
Lookup / Cast Master / Geocode / Area / Route
====================================================== */

/* ======================================================
COMMON LOOKUP
====================================================== */

function getCastById(id){
  return state.casts.find(c => c.id === id) || null;
}

function getVehicleById(id){
  return state.vehicles.find(v => v.id === id) || null;
}

function getPlanById(id){
  return state.plans.find(p => p.id === id) || null;
}

function getActualById(id){
  return state.actuals.find(a => a.id === id) || null;
}

function resolveCastName(castId){
  return getCastById(castId)?.name || "";
}

function resolveCastAddress(castId){
  return getCastById(castId)?.address || "";
}

function resolveCastArea(castId){
  return getCastById(castId)?.area || "";
}

/* ======================================================
AREA LOGIC
====================================================== */

function normalizeArea(a){
  return String(a || "未分類").trim() || "未分類";
}

function areaFromAddress(address){
  const a = String(address || "");

  if(/柏|流山|我孫子|印西|白井|野田|守谷|取手/.test(a)) return "柏方面";
  if(/船橋|市川|浦安|習志野|八千代|鎌ケ谷|鎌ヶ谷/.test(a)) return "船橋方面";
  if(/松戸|金町|柴又|亀有/.test(a)) return "松戸近郊";
  if(/千代田|中央|港|新宿|渋谷|品川|目黒|世田谷|大田|中野|杉並/.test(a)) return "都内";
  if(/埼玉|川口|草加|越谷|三郷|八潮/.test(a)) return "埼玉";

  return "未分類";
}

function areaFromCoords(lat, lon, address=""){
  const a = String(address || "");

  if(/柏|流山|我孫子|印西|白井|野田|守谷|取手/.test(a)) return "柏方面";
  if(/船橋|市川|浦安|習志野|八千代|鎌ケ谷|鎌ヶ谷/.test(a)) return "船橋方面";
  if(/松戸|金町|柴又|亀有/.test(a)) return "松戸近郊";
  if(/千代田|中央|港|新宿|渋谷|品川|目黒|世田谷|大田|中野|杉並/.test(a)) return "都内";
  if(/埼玉|川口|草加|越谷|三郷|八潮/.test(a)) return "埼玉";

  const km = haversineKm(MATSUDO_STATION.lat, MATSUDO_STATION.lon, lat, lon);
  if(km <= 8) return "松戸近郊";

  const dLat = lat - MATSUDO_STATION.lat;
  const dLon = lon - MATSUDO_STATION.lon;

  if(dLat >= 0 && dLon >= 0) return "柏方面";
  if(dLat < 0 && dLon >= 0) return "船橋方面";
  if(dLat < 0 && dLon < 0) return "都内";
  if(dLat >= 0 && dLon < 0) return "埼玉";

  return "未分類";
}

/* ======================================================
CAST FORM
====================================================== */

function resetCastForm(){
  if(!$("castId")) return;

  $("castId").value = "";
  $("castName").value = "";
  $("castAddress").value = "";
  $("castArea").value = "";
  $("castKm").value = "";
  $("castMemo").value = "";
}

async function fillCastGeoAndAreaFromAddress(){
  const address = $("castAddress")?.value?.trim() || "";
  if(!address) return null;

  let geo = null;

  try{
    geo = await geocodeAddress(address);
  }catch(err){
    console.error(err);
  }

  if(geo){
    const area = areaFromCoords(geo.lat, geo.lon, address);

    if($("castArea")){
      $("castArea").value = area;
    }

    return {
      lat: geo.lat,
      lon: geo.lon,
      area
    };
  }

  const area = areaFromAddress(address);

  if($("castArea")){
    $("castArea").value = area;
  }

  return {
    lat: null,
    lon: null,
    area
  };
}

function bindCastAutoArea(){
  const btn = $("btnCastAutoArea");
  if(!btn) return;

  btn.onclick = async () => {
    const address = $("castAddress").value.trim();
    if(!address) return;
    await fillCastGeoAndAreaFromAddress();
  };
}

function bindCastReset(){
  const btn = $("btnCastReset");
  if(!btn) return;

  btn.onclick = () => {
    resetCastForm();
  };
}

function bindCastForm(){
  const form = $("castForm");
  if(!form) return;

  form.onsubmit = async e => {
    e.preventDefault();

    const id = $("castId").value.trim();
    const name = $("castName").value.trim();
    const address = $("castAddress").value.trim();
    const memo = $("castMemo").value.trim();
    const km = toNum($("castKm").value);

    if(!name || !address){
      alert("氏名と住所は必須です");
      return;
    }

    let lat = null;
    let lon = null;
    let area = $("castArea").value.trim() || areaFromAddress(address);

    try{
      const geo = await geocodeAddress(address);
      if(geo){
        lat = geo.lat;
        lon = geo.lon;
        area = areaFromCoords(lat, lon, address);
      }
    }catch(err){
      console.error("geocode error:", err);
    }

    if(id){
      const c = getCastById(id);
      if(c){
        c.name = name;
        c.address = address;
        c.memo = memo;
        c.km = km;
        c.lat = lat;
        c.lon = lon;
        c.area = area;
      }
    }else{
      state.casts.push({
        id: uid("cast"),
        name,
        address,
        memo,
        km,
        lat,
        lon,
        area
      });
    }

    saveState();
    renderCasts();
    resetCastForm();
  };
}

/* ======================================================
CAST TABLE
====================================================== */

function renderCasts(){
  const table = $("castsTable");
  if(!table) return;

  const tb = table.querySelector("tbody");
  tb.innerHTML = "";

  const list = [...state.casts].sort((a,b) => a.name.localeCompare(b.name, "ja"));

  list.forEach(c => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${esc(c.name)}</td>
      <td>${esc(c.address)}</td>
      <td>${esc(c.area || "")}</td>
      <td>${Number(c.km || 0) || ""}</td>
      <td>${c.lat ?? ""}</td>
      <td>${c.lon ?? ""}</td>
      <td>${esc(c.memo || "")}</td>
      <td>
        <div class="actions">
          <button class="btn ghost" data-act="edit" data-id="${esc(c.id)}">編集</button>
          <button class="btn ghost" data-act="route" data-id="${esc(c.id)}">ルート</button>
          <button class="btn danger ghost" data-act="del" data-id="${esc(c.id)}">削除</button>
        </div>
      </td>
    `;

    tb.appendChild(tr);
  });
}

function bindCastTable(){
  const table = $("castsTable");
  if(!table) return;

  table.onclick = e => {
    const btn = e.target.closest("button");
    if(!btn) return;

    const id = btn.dataset.id;
    const act = btn.dataset.act;

    const c = getCastById(id);
    if(!c) return;

    if(act === "edit"){
      $("castId").value = c.id;
      $("castName").value = c.name || "";
      $("castAddress").value = c.address || "";
      $("castArea").value = c.area || "";
      $("castMemo").value = c.memo || "";
      $("castKm").value = c.km || "";

      showView("casts");
      return;
    }

    if(act === "route"){
      if(!c.address){
        alert("住所がありません");
        return;
      }
      openRouteFromMatsudo(c.address);
      return;
    }

    if(act === "del"){
      if(!confirm("削除しますか？")) return;

      state.casts = state.casts.filter(x => x.id !== id);
      saveState();
      renderCasts();
    }
  };
}

/* ======================================================
CSV IMPORT
====================================================== */

function parseCsvLine(line){
  const result = [];
  let current = "";
  let inQuotes = false;

  for(let i = 0; i < line.length; i++){
    const ch = line[i];
    const next = line[i + 1];

    if(ch === '"'){
      if(inQuotes && next === '"'){
        current += '"';
        i++;
      }else{
        inQuotes = !inQuotes;
      }
      continue;
    }

    if(ch === "," && !inQuotes){
      result.push(current);
      current = "";
      continue;
    }

    current += ch;
  }

  result.push(current);
  return result.map(v => String(v).trim());
}

async function enrichMissingCastCoords(){
  const targets = state.casts.filter(c =>
    (c.lat == null || c.lon == null) && c.address
  );

  for(const c of targets){
    try{
      const geo = await geocodeAddress(c.address);

      if(geo){
        c.lat = geo.lat;
        c.lon = geo.lon;
        c.area = areaFromCoords(c.lat, c.lon, c.address);
        saveState();
      }
    }catch(err){
      console.error("geocode failed:", c.address, err);
    }
  }

  renderCasts();
}

function bindCastCsvImport(){
  const fileInput = $("fileImportCastsCsv");
  if(!fileInput) return;

  fileInput.onchange = e => {
    const file = e.target.files?.[0];
    if(!file) return;

    const reader = new FileReader();

    reader.onload = async () => {
      const text = String(reader.result || "");
      const rows = text.split(/\r?\n/).filter(Boolean);

      if(rows.length <= 1){
        alert("CSVデータがありません");
        return;
      }

      const header = parseCsvLine(rows[0]).map(h => h.toLowerCase());

      const idxName = header.findIndex(h => h.includes("name") || h.includes("氏名") || h.includes("名前"));
      const idxAddress = header.findIndex(h => h.includes("address") || h.includes("住所"));
      const idxMemo = header.findIndex(h => h.includes("memo") || h.includes("メモ"));
      const idxKm = header.findIndex(h => h.includes("km") || h.includes("距離"));
      const idxLat = header.findIndex(h => h.includes("lat") || h.includes("緯度"));
      const idxLon = header.findIndex(h => h.includes("lon") || h.includes("lng") || h.includes("経度"));

      for(let i = 1; i < rows.length; i++){
        const cols = parseCsvLine(rows[i]);

        const name = idxName >= 0 ? (cols[idxName] || "").trim() : (cols[0] || "").trim();
        const address = idxAddress >= 0 ? (cols[idxAddress] || "").trim() : (cols[1] || "").trim();
        const memo = idxMemo >= 0 ? (cols[idxMemo] || "").trim() : "";
        const km = idxKm >= 0 ? toNum(cols[idxKm]) : 0;
        const lat = idxLat >= 0 ? toNullableNum(cols[idxLat]) : null;
        const lon = idxLon >= 0 ? toNullableNum(cols[idxLon]) : null;

        if(!name || !address) continue;

        const area =
          lat !== null && lon !== null
            ? areaFromCoords(lat, lon, address)
            : areaFromAddress(address);

        state.casts.push({
          id: uid("cast"),
          name,
          address,
          memo,
          km,
          lat,
          lon,
          area
        });
      }

      saveState();
      renderCasts();
      await enrichMissingCastCoords();

      fileInput.value = "";
    };

    reader.readAsText(file);
  };
}

/* ======================================================
CAST MODULE INIT
====================================================== */

function initCastModule(){
  bindCastForm();
  bindCastAutoArea();
  bindCastReset();
  bindCastTable();
  bindCastCsvImport();
}
/* ======================================================
THEMIS AI Dispatch
Part3
Vehicle / Active Vehicle / View / Home
====================================================== */

/* ======================================================
VEHICLE HELPERS
====================================================== */

function vehicleAvgPerDay(v){
  const km = Number(v.monthKm || 0);
  const days = Number(v.workDays || 0);
  if(!days) return 0;
  return km / days;
}

function ensureActiveVehicleState(){
  state.activeVehicleIds = Array.isArray(state.activeVehicleIds)
    ? state.activeVehicleIds
    : [];

  const validIds = new Set(state.vehicles.map(v => v.id));
  state.activeVehicleIds = state.activeVehicleIds.filter(id => validIds.has(id));

  if(state.activeVehicleIds.length === 0 && state.vehicles.length > 0){
    state.activeVehicleIds = state.vehicles.map(v => v.id);
  }
}

function getActiveVehicles(){
  ensureActiveVehicleState();
  const ids = new Set(state.activeVehicleIds || []);
  return state.vehicles.filter(v => ids.has(v.id));
}

/* ======================================================
VEHICLE FORM
====================================================== */

function resetVehicleForm(){
  if(!$("vehicleEditId")) return;

  $("vehicleEditId").value = "";
  $("vehicleId").value = "";
  $("vehicleArea").value = "";
  $("vehicleHomeArea").value = "";
  $("vehicleCap").value = "";
  $("vehicleDriver").value = "";
}

function bindVehicleForm(){
  const form = $("vehicleForm");
  if(!form) return;

  form.onsubmit = e => {
    e.preventDefault();

    const editId = $("vehicleEditId").value.trim();

    const id = $("vehicleId").value.trim().toUpperCase();
    const area = $("vehicleArea").value.trim();
    const homeArea = $("vehicleHomeArea").value.trim() || area;
    const capacity = Math.max(4, Math.floor(toNum($("vehicleCap").value) || 4));
    const driver = $("vehicleDriver").value.trim();

    if(!id || !area){
      alert("車両IDと担当方面は必須です");
      return;
    }

    const duplicate = state.vehicles.find(v => v.id === id && v.id !== editId);
    if(duplicate){
      alert("同じ車両IDが既に存在します");
      return;
    }

    if(editId){
      const v = getVehicleById(editId);

      if(v){
        v.id = id;
        v.area = area;
        v.homeArea = homeArea;
        v.capacity = capacity;
        v.driver = driver;
      }

      if(editId !== id){
        state.dispatchResult = state.dispatchResult || {};

        Object.keys(state.dispatchResult).forEach(actualId => {
          if(state.dispatchResult[actualId] === editId){
            state.dispatchResult[actualId] = id;
          }
        });

        state.activeVehicleIds = state.activeVehicleIds.map(x => x === editId ? id : x);
      }

    }else{
      state.vehicles.push({
        id,
        area,
        homeArea,
        capacity,
        driver,
        monthKm: 0,
        workDays: 0
      });

      if(!state.activeVehicleIds.includes(id)){
        state.activeVehicleIds.push(id);
      }
    }

    saveState();
    renderVehicles();
    renderTodayVehicles();
    renderHome();
    resetVehicleForm();
  };
}

function bindVehicleReset(){
  const btn = $("btnVehicleReset");
  if(!btn) return;

  btn.onclick = () => {
    resetVehicleForm();
  };
}

/* ======================================================
VEHICLE TABLE
====================================================== */

function renderVehicles(){
  const table = $("vehiclesTable");
  if(!table) return;

  const tb = table.querySelector("tbody");
  tb.innerHTML = "";

  const list = [...state.vehicles].sort((a,b) => a.id.localeCompare(b.id));

  list.forEach(v => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td><b>${esc(v.id)}</b></td>
      <td>${esc(v.area)}</td>
      <td>${esc(v.homeArea || v.area)}</td>
      <td>${Number(v.capacity || 4)}</td>
      <td>${esc(v.driver || "")}</td>
      <td>${Number(v.monthKm || 0).toFixed(1)}</td>
      <td>${Number(v.workDays || 0)}</td>
      <td>${vehicleAvgPerDay(v).toFixed(1)}</td>
      <td>
        <div class="actions">
          <button class="btn ghost" data-act="edit" data-id="${esc(v.id)}">編集</button>
          <button class="btn danger ghost" data-act="del" data-id="${esc(v.id)}">削除</button>
        </div>
      </td>
    `;

    tb.appendChild(tr);
  });
}

function bindVehicleTable(){
  const table = $("vehiclesTable");
  if(!table) return;

  table.onclick = e => {
    const btn = e.target.closest("button");
    if(!btn) return;

    const id = btn.dataset.id;
    const act = btn.dataset.act;

    const v = getVehicleById(id);
    if(!v) return;

    if(act === "edit"){
      $("vehicleEditId").value = v.id;
      $("vehicleId").value = v.id;
      $("vehicleArea").value = v.area || "";
      $("vehicleHomeArea").value = v.homeArea || v.area || "";
      $("vehicleCap").value = Number(v.capacity || 4);
      $("vehicleDriver").value = v.driver || "";

      showView("vehicles");
      return;
    }

    if(act === "del"){
      if(!confirm("削除しますか？")) return;

      state.vehicles = state.vehicles.filter(x => x.id !== id);
      state.activeVehicleIds = state.activeVehicleIds.filter(x => x !== id);

      if(state.dispatchResult){
        Object.keys(state.dispatchResult).forEach(actualId => {
          if(state.dispatchResult[actualId] === id){
            delete state.dispatchResult[actualId];
          }
        });
      }

      saveState();
      renderVehicles();
      renderTodayVehicles();
      renderHome();
    }
  };
}

/* ======================================================
ACTIVE VEHICLES UI
====================================================== */

function renderTodayVehicles(){
  const box = $("todayVehiclesBox");
  if(!box) return;

  ensureActiveVehicleState();

  if(!state.vehicles.length){
    box.innerHTML = `<div class="muted">車両がありません</div>`;
    return;
  }

  box.innerHTML = state.vehicles
    .slice()
    .sort((a,b) => a.id.localeCompare(b.id))
    .map(v => {
      const checked = state.activeVehicleIds.includes(v.id) ? "checked" : "";

      return `
        <label style="display:flex;align-items:center;gap:10px;margin:6px 0;flex-wrap:wrap;">
          <input type="checkbox" class="vehTodayChk" data-id="${esc(v.id)}" ${checked} />
          <span>
            <b>車${esc(v.id)}</b>
            （${esc(v.area)} / 帰宅:${esc(v.homeArea || v.area)} / 定員${Number(v.capacity || 4)}${v.driver ? " / " + esc(v.driver) : ""}）
          </span>
        </label>
      `;
    })
    .join("");

  box.querySelectorAll(".vehTodayChk").forEach(chk => {
    chk.addEventListener("change", () => {
      const id = chk.dataset.id;

      if(chk.checked){
        if(!state.activeVehicleIds.includes(id)){
          state.activeVehicleIds.push(id);
        }
      }else{
        state.activeVehicleIds = state.activeVehicleIds.filter(x => x !== id);
      }

      saveState();
      renderHome();
      if(typeof renderDispatch === "function") renderDispatch();
    });
  });
}

function bindTodayVehicleButtons(){
  const btnAll = $("btnSelectAllVehicles");
  const btnClear = $("btnClearVehiclesToday");

  if(btnAll){
    btnAll.onclick = () => {
      state.activeVehicleIds = state.vehicles.map(v => v.id);
      saveState();
      renderTodayVehicles();
      renderHome();
      if(typeof renderDispatch === "function") renderDispatch();
    };
  }

  if(btnClear){
    btnClear.onclick = () => {
      state.activeVehicleIds = [];
      saveState();
      renderTodayVehicles();
      renderHome();
      if(typeof renderDispatch === "function") renderDispatch();
    };
  }
}

/* ======================================================
MONTH RESET
====================================================== */

function bindMonthlyReset(){
  const btn = $("btnResetMonth");
  if(!btn) return;

  btn.onclick = () => {
    if(!confirm("月間距離と出勤日数をリセットしますか？")) return;

    state.vehicles.forEach(v => {
      v.monthKm = 0;
      v.workDays = 0;
    });

    saveState();
    renderVehicles();
    renderHome();
  };
}

/* ======================================================
VIEW / NAV / HOME
====================================================== */

function showView(name){
  document.querySelectorAll(".view").forEach(v => {
    v.classList.add("hidden");
  });

  const el = $("view-" + name);
  if(el) el.classList.remove("hidden");

  document.querySelectorAll(".tab").forEach(t => {
    t.classList.remove("active");
  });

  const activeTab = document.querySelector(`.tab[data-view="${name}"]`);
  if(activeTab) activeTab.classList.add("active");

  if(name === "casts" && typeof renderCasts === "function") renderCasts();
  if(name === "vehicles" && typeof renderVehicles === "function") renderVehicles();
  if(name === "plan" && typeof renderPlans === "function") renderPlans();
  if(name === "actual" && typeof renderActuals === "function") renderActuals();
  if(name === "daily" && typeof renderDispatch === "function") renderDispatch();
  if(name === "home") renderHome();
}

function bindTabs(){
  document.querySelectorAll(".tab").forEach(tab => {
    tab.addEventListener("click", () => {
      const view = tab.dataset.view;
      showView(view);
    });
  });

  document.querySelectorAll("[data-nav]").forEach(btn => {
    btn.addEventListener("click", () => {
      const view = btn.dataset.nav;
      showView(view);
    });
  });
}

function bindGlobalButtons(){
  const btnExport = $("btnExport");
  if(btnExport){
    btnExport.onclick = () => {
      const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "themis_dispatch_v3.json";
      a.click();
      URL.revokeObjectURL(a.href);
    };
  }

  const fileImport = $("fileImport");
  if(fileImport){
    fileImport.onchange = e => {
      const file = e.target.files?.[0];
      if(!file) return;

      const reader = new FileReader();

      reader.onload = () => {
        try{
          const data = normalizeState(JSON.parse(reader.result));
          Object.keys(state).forEach(k => delete state[k]);
          Object.assign(state, data);
          saveState();
          location.reload();
        }catch{
          alert("読み込みに失敗しました");
        }finally{
          fileImport.value = "";
        }
      };

      reader.readAsText(file);
    };
  }

  const btnResetAll = $("btnResetAll");
  if(btnResetAll){
    btnResetAll.onclick = () => {
      if(!confirm("すべて削除しますか？")) return;
      localStorage.removeItem(STORAGE_KEY);
      location.reload();
    };
  }
}

function syncDateInputs(sourceId){
  const src = $(sourceId);
  if(!src) return;

  const value = src.value || todayStr();

  ["planDate","planDate_plan","planDate_actual"].forEach(id => {
    if($(id) && id !== sourceId){
      $(id).value = value;
    }
  });
}

function bindDateSync(){
  ["planDate","planDate_plan","planDate_actual"].forEach(id => {
    const el = $(id);
    if(!el) return;

    el.addEventListener("change", () => {
      syncDateInputs(id);
      if(typeof renderPlans === "function") renderPlans();
      if(typeof renderActuals === "function") renderActuals();
      if(typeof renderDispatch === "function") renderDispatch();
      renderHome();
    });
  });
}

function renderHome(){
  const homeToday = $("homeToday");
  const homeMonth = $("homeMonth");

  if(homeToday){
    const done = state.actuals.filter(x => x.status === "done").length;
    const canceled = state.actuals.filter(x => x.status === "canceled").length;

    homeToday.innerHTML = `
      <div class="row gap">
        <span class="pill">キャスト ${state.casts.length}人</span>
        <span class="pill">車両 ${state.vehicles.length}台</span>
        <span class="pill">Plan ${state.plans.length}件</span>
        <span class="pill">Actual ${state.actuals.length}件</span>
        <span class="pill">完了 ${done}件</span>
        <span class="pill">キャンセル ${canceled}件</span>
      </div>
    `;
  }

  if(homeMonth){
    if(!state.vehicles.length){
      homeMonth.innerHTML = `<div class="muted">車両がありません</div>`;
    }else{
      homeMonth.innerHTML = state.vehicles
        .slice()
        .sort((a,b) => a.id.localeCompare(b.id))
        .map(v => `
          <div class="row gap" style="margin:6px 0;">
            <span class="pill">車${esc(v.id)}</span>
            <span class="pill">${esc(v.area)}</span>
            <span class="pill">帰宅:${esc(v.homeArea || v.area)}</span>
            <span class="pill">月間:${Number(v.monthKm || 0).toFixed(1)}km</span>
            <span class="pill">出勤:${v.workDays || 0}日</span>
            <span class="pill">平均:${vehicleAvgPerDay(v).toFixed(1)}km</span>
          </div>
        `)
        .join("");
    }
  }
}

function renderAll(){
  if($("planDate") && !$("planDate").value) $("planDate").value = todayStr();
  if($("planDate_plan") && !$("planDate_plan").value) $("planDate_plan").value = $("planDate")?.value || todayStr();
  if($("planDate_actual") && !$("planDate_actual").value) $("planDate_actual").value = $("planDate")?.value || todayStr();

  if($("balanceStrength")) $("balanceStrength").value = String(state.settings?.balanceStrength ?? 1);
  if($("lastAreaPenalty")) $("lastAreaPenalty").value = String(state.settings?.lastAreaPenalty ?? 10);

  if(typeof renderCasts === "function") renderCasts();
  renderVehicles();
  renderTodayVehicles();
  renderHome();
  if(typeof renderPlans === "function") renderPlans();
  if(typeof renderActuals === "function") renderActuals();
  if(typeof renderDispatch === "function") renderDispatch();
}

/* ======================================================
VEHICLE MODULE INIT
====================================================== */

function initVehicleModule(){
  bindVehicleForm();
  bindVehicleReset();
  bindVehicleTable();
  bindTodayVehicleButtons();
  bindMonthlyReset();
}
/* ======================================================
THEMIS AI Dispatch
Part4
Plan / Actual / Status / Matrix
====================================================== */

/* ======================================================
DATE HELPERS
====================================================== */

function currentPlanDate(){
  return $("planDate_plan")?.value || $("planDate")?.value || todayStr();
}

function currentActualDate(){
  return $("planDate_actual")?.value || $("planDate")?.value || todayStr();
}

function currentDailyDate(){
  return $("planDate")?.value || todayStr();
}

/* ======================================================
NORMALIZE
====================================================== */

function normalizePlanItem(p){
  return {
    id: String(p.id || uid("plan")),
    date: String(p.date || currentPlanDate()),
    castId: String(p.castId || ""),
    hour: Math.max(0, Math.min(5, Math.floor(toNum(p.hour)))),
    address: String(p.address || "").trim(),
    area: String(p.area || "").trim(),
    memo: String(p.memo || "").trim(),
    kmOverride: toNum(p.kmOverride),
    status: String(p.status || "active")
  };
}

function normalizeActualItem(a){
  const status = ["active","done","canceled"].includes(a.status) ? a.status : "active";

  return {
    id: String(a.id || uid("actual")),
    date: String(a.date || currentActualDate()),
    castId: String(a.castId || ""),
    hour: Math.max(0, Math.min(5, Math.floor(toNum(a.hour)))),
    address: String(a.address || "").trim(),
    area: String(a.area || "").trim(),
    memo: String(a.memo || "").trim(),
    kmOverride: toNum(a.kmOverride),
    status
  };
}

function getPlansByDate(date){
  return state.plans.filter(p => (p.date || "") === date);
}

function getActualsByDate(date){
  return state.actuals.filter(a => (a.date || "") === date);
}

function isCastDoneToday(castId, date){
  return state.actuals.some(a =>
    a.date === date &&
    a.castId === castId &&
    a.status === "done"
  );
}

/* ======================================================
SELECT HELPERS
====================================================== */

function populateCastSelectWithValue(selectId, value){
  const sel = $(selectId);
  if(!sel) return;

  sel.innerHTML = "";

  const empty = document.createElement("option");
  empty.value = "";
  empty.textContent = "選択してください";
  sel.appendChild(empty);

  const date =
    selectId === "actualCastSelect"
      ? currentActualDate()
      : currentPlanDate();

  state.casts
    .slice()
    .sort((a,b) => a.name.localeCompare(b.name,"ja"))
    .forEach(c => {
      const doneToday = isCastDoneToday(c.id, date);

      const op = document.createElement("option");
      op.value = c.id;
      op.textContent = doneToday ? `${c.name} ✅完了` : c.name;

      if(String(value || "") === c.id){
        op.selected = true;
      }

      sel.appendChild(op);
    });
}

function populatePlanCastSelect(){
  populateCastSelectWithValue("planCastSelect", $("planCastSelect")?.value || "");
}

function populateActualCastSelect(){
  populateCastSelectWithValue("actualCastSelect", $("actualCastSelect")?.value || "");
}

function populateActualFromPlanSelect(){
  const sel = $("actualFromPlanSelect");
  if(!sel) return;

  const date = currentActualDate();
  const currentValue = sel.value || "";

  sel.innerHTML = "";

  const empty = document.createElement("option");
  empty.value = "";
  empty.textContent = "予定から選択";
  sel.appendChild(empty);

  getPlansByDate(date)
    .slice()
    .sort((a,b) => a.hour - b.hour)
    .forEach(p => {
      const cast = getCastById(p.castId);
      const doneToday = isCastDoneToday(p.castId, date);

      const op = document.createElement("option");
      op.value = p.id;
      op.textContent = doneToday
        ? `${p.hour}時 ${cast?.name || ""} ✅完了`
        : `${p.hour}時 ${cast?.name || ""}`;

      if(currentValue === p.id){
        op.selected = true;
      }

      sel.appendChild(op);
    });
}

/* ======================================================
PLAN FORM
====================================================== */

function resetPlanForm(){
  if(!$("planId")) return;

  $("planId").value = "";
  $("planCastSelect").value = "";
  $("planHour").value = "0";
  $("planAddress").value = "";
  $("planArea").value = "";
  $("planMemo").value = "";
  $("planKmOverride").value = "";
}

function bindPlanHelpers(){
  const castSel = $("planCastSelect");
  if(castSel){
    castSel.addEventListener("change", () => {
      const cast = getCastById(castSel.value);
      if(!cast) return;

      if($("planAddress") && !$("planAddress").value){
        $("planAddress").value = cast.address || "";
      }

      if($("planArea") && !$("planArea").value){
        $("planArea").value = cast.area || areaFromAddress(cast.address);
      }
    });
  }

  const btnAuto = $("btnPlanAutoArea");
  if(btnAuto){
    btnAuto.onclick = async () => {
      const addr = $("planAddress").value.trim();
      if(!addr) return;

      try{
        const geo = await geocodeAddress(addr);
        if(geo){
          $("planArea").value = areaFromCoords(geo.lat, geo.lon, addr);
          return;
        }
      }catch(err){
        console.error(err);
      }

      $("planArea").value = areaFromAddress(addr);
    };
  }

  const btnReset = $("btnPlanReset");
  if(btnReset){
    btnReset.onclick = () => resetPlanForm();
  }

  const btnClear = $("btnClearPlan");
  if(btnClear){
    btnClear.onclick = () => {
      const date = currentPlanDate();

      if(!confirm(`予定表 ${date} を全削除しますか？`)) return;

      state.plans = state.plans.filter(p => p.date !== date);
      saveState();
      renderPlans();
      populateActualFromPlanSelect();
      renderHome();
    };
  }
}

function bindPlanForm(){
  const form = $("planForm");
  if(!form) return;

  form.onsubmit = e => {
    e.preventDefault();

    const id = $("planId").value.trim();
    const castId = $("planCastSelect").value;
    const hour = Number($("planHour").value);

    const cast = getCastById(castId);
    if(!castId || !cast){
      alert("キャスト選択必須です");
      return;
    }

    const payload = normalizePlanItem({
      id: id || uid("plan"),
      date: currentPlanDate(),
      castId,
      hour,
      address: $("planAddress").value.trim() || cast.address,
      area: $("planArea").value.trim() || cast.area || areaFromAddress(cast.address),
      memo: $("planMemo").value.trim(),
      kmOverride: $("planKmOverride").value,
      status: "active"
    });

    if(id){
      const idx = state.plans.findIndex(x => x.id === id);
      if(idx >= 0){
        state.plans[idx] = payload;
      }
    }else{
      state.plans.push(payload);
    }

    saveState();
    renderPlans();
    populateActualFromPlanSelect();
    resetPlanForm();
    renderHome();
  };
}

/* ======================================================
PLAN TABLE
====================================================== */

function renderPlans(){
  populatePlanCastSelect();

  const table = $("planTable");
  if(!table) return;

  const tb = table.querySelector("tbody");
  tb.innerHTML = "";

  const date = currentPlanDate();

  const list = getPlansByDate(date)
    .slice()
    .sort((a,b) => {
      if(a.hour !== b.hour) return a.hour - b.hour;

      const areaA = normalizeArea(a.area);
      const areaB = normalizeArea(b.area);
      if(areaA !== areaB){
        return areaA.localeCompare(areaB, "ja");
      }

      const castA = resolveCastName(a.castId);
      const castB = resolveCastName(b.castId);
      return castA.localeCompare(castB, "ja");
    });

  let lastHour = null;
  let lastArea = null;

  list.forEach(p => {
    const cast = getCastById(p.castId);
    const castName = cast?.name || "";
    const area = normalizeArea(p.area);
    const doneToday = isCastDoneToday(p.castId, date);

    if(lastHour !== p.hour){
      const headTr = document.createElement("tr");
      headTr.innerHTML = `
        <td colspan="7" style="background:rgba(255,255,255,.06);font-weight:700;">
          ${p.hour}時
        </td>
      `;
      tb.appendChild(headTr);
      lastHour = p.hour;
      lastArea = null;
    }

    if(lastArea !== area){
      const areaTr = document.createElement("tr");
      areaTr.innerHTML = `
        <td colspan="7" style="background:rgba(255,255,255,.03);padding-left:18px;">
          <span class="pill">${esc(area)}</span>
        </td>
      `;
      tb.appendChild(areaTr);
      lastArea = area;
    }

    const tr = document.createElement("tr");

    if(p.status === "canceled") tr.classList.add("tr-canceled");
    if(doneToday) tr.classList.add("tr-done");

    tr.innerHTML = `
      <td>${p.hour}</td>
      <td>
        <b>${esc(castName)}</b>
        ${doneToday ? `<div class="small mt"><span class="badge done">送り完了</span></div>` : ``}
      </td>
      <td>${esc(p.address)}</td>
      <td>${esc(area)}</td>
      <td>${p.kmOverride || ""}</td>
      <td>
        ${
          p.status === "canceled"
            ? `<span class="badge canceled">キャンセル</span>`
            : doneToday
              ? `<span class="badge done">予定 / 完了</span>`
              : `<span class="badge active">予定</span>`
        }
      </td>
      <td>
        <div class="actions">
          <button class="btn ghost" data-act="edit" data-id="${p.id}">編集</button>
          <button class="btn ghost" data-act="route" data-id="${p.id}">ルート</button>
          <button class="btn danger ghost" data-act="del" data-id="${p.id}">削除</button>
        </div>
      </td>
    `;

    tb.appendChild(tr);
  });
}

function bindPlanTable(){
  const table = $("planTable");
  if(!table) return;

  table.onclick = e => {
    const btn = e.target.closest("button");
    if(!btn) return;

    const id = btn.dataset.id;
    const act = btn.dataset.act;

    const p = getPlanById(id);
    if(!p) return;

    if(act === "edit"){
      $("planId").value = p.id;
      populatePlanCastSelect();
      $("planCastSelect").value = p.castId;
      $("planHour").value = String(p.hour);
      $("planAddress").value = p.address || "";
      $("planArea").value = p.area || "";
      $("planMemo").value = p.memo || "";
      $("planKmOverride").value = p.kmOverride || "";

      showView("plan");
      return;
    }

    if(act === "route"){
      if(!p.address){
        alert("住所がありません");
        return;
      }
      openRouteFromMatsudo(p.address);
      return;
    }

    if(act === "del"){
      if(!confirm("削除しますか？")) return;

      state.plans = state.plans.filter(x => x.id !== id);
      saveState();
      renderPlans();
      populateActualFromPlanSelect();
      renderHome();
    }
  };
}

/* ======================================================
ACTUAL FORM
====================================================== */

function resetActualForm(){
  if(!$("actualId")) return;

  $("actualId").value = "";
  $("actualFromPlanSelect").value = "";
  $("actualCastSelect").value = "";
  $("actualHour").value = "0";
  $("actualAddress").value = "";
  $("actualArea").value = "";
  $("actualMemo").value = "";
  $("actualKmOverride").value = "";
  $("actualStatus").value = "active";
}

function bindActualHelpers(){
  const castSel = $("actualCastSelect");
  if(castSel){
    castSel.addEventListener("change", () => {
      const cast = getCastById(castSel.value);
      if(!cast) return;

      if($("actualAddress") && !$("actualAddress").value){
        $("actualAddress").value = cast.address || "";
      }

      if($("actualArea") && !$("actualArea").value){
        $("actualArea").value = cast.area || areaFromAddress(cast.address);
      }
    });
  }

  const fromPlanSel = $("actualFromPlanSelect");
  if(fromPlanSel){
    fromPlanSel.addEventListener("change", () => {
      const plan = getPlanById(fromPlanSel.value);
      if(!plan) return;

      $("actualCastSelect").value = plan.castId || "";
      $("actualHour").value = String(plan.hour ?? 0);
      $("actualAddress").value = plan.address || "";
      $("actualArea").value = plan.area || "";
      $("actualMemo").value = plan.memo || "";
      $("actualKmOverride").value = plan.kmOverride || "";
    });
  }

  const btnAuto = $("btnActualAutoArea");
  if(btnAuto){
    btnAuto.onclick = async () => {
      const addr = $("actualAddress").value.trim();
      if(!addr) return;

      try{
        const geo = await geocodeAddress(addr);
        if(geo){
          $("actualArea").value = areaFromCoords(geo.lat, geo.lon, addr);
          return;
        }
      }catch(err){
        console.error(err);
      }

      $("actualArea").value = areaFromAddress(addr);
    };
  }

  const btnReset = $("btnActualReset");
  if(btnReset){
    btnReset.onclick = () => resetActualForm();
  }

  const btnAddFromPlan = $("btnAddFromPlan");
  if(btnAddFromPlan){
    btnAddFromPlan.onclick = () => {
      const planId = $("actualFromPlanSelect").value;
      const plan = getPlanById(planId);

      if(!plan){
        alert("予定を選択してください");
        return;
      }

      $("actualId").value = "";
      $("actualCastSelect").value = plan.castId || "";
      $("actualHour").value = String(plan.hour ?? 0);
      $("actualAddress").value = plan.address || "";
      $("actualArea").value = plan.area || "";
      $("actualMemo").value = plan.memo || "";
      $("actualKmOverride").value = plan.kmOverride || "";
      $("actualStatus").value = "active";

      showView("actual");
    };
  }

  const btnClear = $("btnClearActual");
  if(btnClear){
    btnClear.onclick = () => {
      const date = currentActualDate();

      if(!confirm(`Actual ${date} を全削除しますか？`)) return;

      state.actuals = state.actuals.filter(a => a.date !== date);
      state.dispatchResult = {};

      saveState();
      renderActuals();
      renderPlans();
      if(typeof renderDispatch === "function") renderDispatch();
      renderHome();
    };
  }
}

function bindActualForm(){
  const form = $("actualForm");
  if(!form) return;

  form.onsubmit = e => {
    e.preventDefault();

    const id = $("actualId").value.trim();
    const castId = $("actualCastSelect").value;
    const hour = Number($("actualHour").value);

    const cast = getCastById(castId);
    if(!castId || !cast){
      alert("キャスト必須です");
      return;
    }

    const payload = normalizeActualItem({
      id: id || uid("actual"),
      date: currentActualDate(),
      castId,
      hour,
      address: $("actualAddress").value.trim() || cast.address,
      area: $("actualArea").value.trim() || cast.area || areaFromAddress(cast.address),
      memo: $("actualMemo").value.trim(),
      kmOverride: $("actualKmOverride").value,
      status: $("actualStatus").value
    });

    if(id){
      const idx = state.actuals.findIndex(x => x.id === id);
      if(idx >= 0){
        state.actuals[idx] = payload;
      }
    }else{
      state.actuals.push(payload);
    }

    saveState();
    renderActuals();
    renderPlans();
    resetActualForm();
    if(typeof renderDispatch === "function") renderDispatch();
    renderHome();
  };
}

/* ======================================================
ACTUAL TABLE
====================================================== */

function renderActuals(){
  populateActualCastSelect();
  populateActualFromPlanSelect();

  const table = $("actualTable");
  if(!table) return;

  const tb = table.querySelector("tbody");
  tb.innerHTML = "";

  const date = currentActualDate();
  const list = getActualsByDate(date)
    .slice()
    .sort((a,b) => a.hour - b.hour);

  list.forEach(a => {
    const cast = getCastById(a.castId);

    const tr = document.createElement("tr");
    if(a.status === "done") tr.classList.add("tr-done");
    if(a.status === "canceled") tr.classList.add("tr-canceled");

    tr.innerHTML = `
      <td>${a.hour}</td>
      <td>${esc(cast?.name || "")}</td>
      <td>${esc(a.address)}</td>
      <td>${esc(a.area)}</td>
      <td>${a.kmOverride || ""}</td>

      <td>
        <div class="actions">
          ${
            a.status === "done"
              ? `<button class="btn ghost" data-act="set-active" data-id="${a.id}">未完了に戻す</button>`
              : `<button class="btn" data-act="set-done" data-id="${a.id}">完了</button>`
          }

          ${
            a.status === "canceled"
              ? `<button class="btn ghost" data-act="set-active" data-id="${a.id}">キャンセル解除</button>`
              : `<button class="btn danger ghost" data-act="set-canceled" data-id="${a.id}">キャンセル</button>`
          }
        </div>

        <div class="small mt">
          ${
            a.status === "done"
              ? `<span class="badge done">完了</span>`
              : a.status === "canceled"
                ? `<span class="badge canceled">キャンセル</span>`
                : `<span class="badge active">未完了</span>`
          }
        </div>
      </td>

      <td>
        <div class="actions">
          <button class="btn ghost" data-act="edit" data-id="${a.id}">編集</button>
          <button class="btn ghost" data-act="route" data-id="${a.id}">ルート</button>
          <button class="btn danger ghost" data-act="del" data-id="${a.id}">削除</button>
        </div>
      </td>
    `;

    tb.appendChild(tr);
  });

  renderActualMatrix();
}

function bindActualTable(){
  const table = $("actualTable");
  if(!table) return;

  table.onclick = e => {
    const btn = e.target.closest("button");
    if(!btn) return;

    const id = btn.dataset.id;
    const act = btn.dataset.act;

    const a = getActualById(id);
    if(!a) return;

    if(act === "edit"){
      $("actualId").value = a.id;
      populateActualCastSelect();
      $("actualCastSelect").value = a.castId;
      $("actualHour").value = String(a.hour);
      $("actualAddress").value = a.address || "";
      $("actualArea").value = a.area || "";
      $("actualMemo").value = a.memo || "";
      $("actualKmOverride").value = a.kmOverride || "";
      $("actualStatus").value = a.status || "active";

      showView("actual");
      return;
    }

    if(act === "route"){
      if(!a.address){
        alert("住所がありません");
        return;
      }
      openRouteFromMatsudo(a.address);
      return;
    }

    if(act === "set-done"){
      a.status = "done";
      saveState();
      renderActuals();
      renderPlans();
      if(typeof renderDispatch === "function") renderDispatch();
      renderHome();
      return;
    }

    if(act === "set-canceled"){
      a.status = "canceled";
      saveState();
      renderActuals();
      renderPlans();
      if(typeof renderDispatch === "function") renderDispatch();
      renderHome();
      return;
    }

    if(act === "set-active"){
      a.status = "active";
      saveState();
      renderActuals();
      renderPlans();
      if(typeof renderDispatch === "function") renderDispatch();
      renderHome();
      return;
    }

    if(act === "del"){
      if(!confirm("削除しますか？")) return;

      state.actuals = state.actuals.filter(x => x.id !== id);

      if(state.dispatchResult && state.dispatchResult[id]){
        delete state.dispatchResult[id];
      }

      saveState();
      renderActuals();
      renderPlans();
      if(typeof renderDispatch === "function") renderDispatch();
      renderHome();
    }
  };
}

/* ======================================================
ACTUAL MATRIX
====================================================== */

function renderActualMatrix(){
  const table = $("actualMatrixTable");
  if(!table) return;

  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  const date = currentActualDate();
  const list = getActualsByDate(date);

  if(!list.length){
    thead.innerHTML = "";
    tbody.innerHTML = `<tr><td class="muted">データなし</td></tr>`;
    return;
  }

  const hours = [0,1,2,3,4,5];
  const areas = [...new Set(list.map(x => x.area || "未分類"))];

  thead.innerHTML = `
    <tr>
      <th>時間</th>
      ${areas.map(a => `<th>${esc(a)}</th>`).join("")}
    </tr>
  `;

  tbody.innerHTML = hours.map(h => {
    const cols = areas.map(area => {
      const items = list.filter(x =>
        x.hour === h &&
        (x.area || "未分類") === area
      );

      if(!items.length){
        return `<td class="muted small">-</td>`;
      }

      const kmSum = items.reduce((s,x) => s + Number(x.kmOverride || 0), 0);

      const names = items.map(x => {
        const cast = getCastById(x.castId);

        const badge =
          x.status === "done"
            ? `<span class="badge done">完了</span>`
            : x.status === "canceled"
              ? `<span class="badge canceled">キャンセル</span>`
              : `<span class="badge active">未完了</span>`;

        return `${badge} ${esc(cast?.name || "")} (${Number(x.kmOverride || 0).toFixed(1)}km)`;
      }).join("<br>");

      return `
        <td>
          <div class="small"><b>${items.length}人</b> / ${kmSum.toFixed(1)}km</div>
          <div class="muted small" style="margin-top:6px;line-height:1.45;">
            ${names}
          </div>
        </td>
      `;
    }).join("");

    return `<tr><th>${h}時</th>${cols}</tr>`;
  }).join("");
}

/* ======================================================
PLAN / ACTUAL MODULE INIT
====================================================== */

function initPlanActualModule(){
  bindPlanHelpers();
  bindPlanForm();
  bindPlanTable();

  bindActualHelpers();
  bindActualForm();
  bindActualTable();
}
/* ======================================================
THEMIS AI Dispatch
Part5
Dispatch AI / Reassign / Finalize / Init
====================================================== */

/* ======================================================
SETTINGS
====================================================== */

function bindSettingsForm(){
  const balance = $("balanceStrength");
  const penalty = $("lastAreaPenalty");

  if(balance){
    balance.onchange = () => {
      state.settings.balanceStrength = Number(balance.value || 1);
      saveState();
      renderDispatch();
      renderHome();
    };
  }

  if(penalty){
    penalty.onchange = () => {
      state.settings.lastAreaPenalty = Number(penalty.value || 10);
      saveState();
      renderDispatch();
      renderHome();
    };
  }
}

/* ======================================================
KM RESOLVE
====================================================== */

function resolveKm(item){
  if(Number(item.kmOverride || 0) > 0){
    return Number(item.kmOverride || 0);
  }

  const cast = getCastById(item.castId);

  if(cast){
    if(Number(cast.km || 0) > 0){
      return Number(cast.km || 0);
    }

    if(
      cast.lat !== null &&
      cast.lat !== undefined &&
      cast.lon !== null &&
      cast.lon !== undefined
    ){
      return Number(
        haversineKm(
          MATSUDO_STATION.lat,
          MATSUDO_STATION.lon,
          Number(cast.lat),
          Number(cast.lon)
        ).toFixed(1)
      );
    }
  }

  return 12;
}

/* ======================================================
AREA CLUSTER
====================================================== */

function clusterByArea(list){
  const map = {};

  list.forEach(x => {
    const k = normalizeArea(x.area);
    if(!map[k]) map[k] = [];
    map[k].push(x);
  });

  return map;
}

/* ======================================================
DISPATCH TARGETS
完了したキャストは次の自動配車から外す
====================================================== */

function getDispatchActuals(){
  const date = currentDailyDate();

  return getActualsByDate(date)
    .filter(x => x.status === "active");
}

function makeHourLoadMap(vehicles){
  const map = {};

  vehicles.forEach(v => {
    map[v.id] = { 0:0, 1:0, 2:0, 3:0, 4:0, 5:0 };
  });

  return map;
}

function makeVehicleKmMap(vehicles){
  const map = {};

  vehicles.forEach(v => {
    map[v.id] = 0;
  });

  return map;
}

/* ======================================================
SCORING
====================================================== */

function areaMatchScore(vehicle, item){
  const itemArea = normalizeArea(item.area);
  const vArea = normalizeArea(vehicle.area);
  const vHome = normalizeArea(vehicle.homeArea || vehicle.area);

  let score = 0;

  if(vArea !== itemArea){
    score += 5;
  }

  if(Number(item.hour) === 5 && vArea !== itemArea){
    score += Number(state.settings.lastAreaPenalty || 10);
  }

  if(vHome === itemArea){
    score -= 15;
  }

  if(Number(item.hour) === 5 && vHome === itemArea){
    score -= 10;
  }

  return score;
}

function balanceScore(vehicle){
  const avg = vehicleAvgPerDay(vehicle);
  const strength = Number(state.settings.balanceStrength || 1);
  return avg * strength;
}

function dispatchScore(vehicle, item, hourLoadMap, vehicleKmMap, areaRideMap){
  const baseKm = resolveKm(item);
  const hour = Number(item.hour || 0);

  let score = baseKm;

  score += areaMatchScore(vehicle, item);
  score += balanceScore(vehicle);
  score += (hourLoadMap[vehicle.id]?.[hour] || 0) * 8;
  score += Number(vehicleKmMap[vehicle.id] || 0) * 0.08;

  const sameAreaCount =
    areaRideMap[vehicle.id]?.[normalizeArea(item.area)] || 0;

  score -= sameAreaCount * 10;

  if(hour === 5){
    score -= sameAreaCount * 8;
  }

  return score;
}

/* ======================================================
DISPATCH CORE
====================================================== */

function chooseVehicleForRide(item, vehicles, hourLoadMap, vehicleKmMap, areaRideMap){
  let bestVehicle = null;
  let bestScore = Infinity;

  vehicles.forEach(vehicle => {
    const hour = Number(item.hour || 0);
    const currentLoad = hourLoadMap[vehicle.id]?.[hour] || 0;
    const cap = Number(vehicle.capacity || 4);

    if(currentLoad >= cap) return;

    const score = dispatchScore(
      vehicle,
      item,
      hourLoadMap,
      vehicleKmMap,
      areaRideMap
    );

    if(score < bestScore){
      bestScore = score;
      bestVehicle = vehicle;
    }
  });

  return bestVehicle;
}

function computeDispatch(){
  const vehicles = getActiveVehicles();
  const rides = getDispatchActuals();

  if(!vehicles.length || !rides.length){
    return {};
  }

  const assignments = {};
  const hourLoadMap = makeHourLoadMap(vehicles);
  const vehicleKmMap = makeVehicleKmMap(vehicles);
  const areaRideMap = {};

  vehicles.forEach(v => {
    areaRideMap[v.id] = {};
  });

  const byHour = {};

  rides.forEach(r => {
    const hour = Number(r.hour || 0);
    if(!byHour[hour]) byHour[hour] = [];
    byHour[hour].push(r);
  });

  Object.keys(byHour).forEach(hourKey => {
    const hour = Number(hourKey);
    const clustered = clusterByArea(byHour[hour]);

    Object.keys(clustered).forEach(area => {
      clustered[area].sort((a,b) => resolveKm(b) - resolveKm(a));
    });

    const merged = Object.keys(clustered)
      .sort((a,b) => {
        const ak = clustered[a].reduce((s,x) => s + resolveKm(x), 0);
        const bk = clustered[b].reduce((s,x) => s + resolveKm(x), 0);
        return bk - ak;
      })
      .flatMap(area => clustered[area]);

    merged.forEach(item => {
      const chosen = chooseVehicleForRide(
        item,
        vehicles,
        hourLoadMap,
        vehicleKmMap,
        areaRideMap
      );

      if(!chosen) return;

      assignments[item.id] = chosen.id;
      hourLoadMap[chosen.id][hour] += 1;
      vehicleKmMap[chosen.id] += resolveKm(item);

      const key = normalizeArea(item.area);
      areaRideMap[chosen.id][key] = (areaRideMap[chosen.id][key] || 0) + 1;
    });
  });

  return assignments;
}

/* ======================================================
MANUAL REASSIGN
====================================================== */

function getEffectiveAssignments(){
  const base = computeDispatch();

  const saved =
    state.dispatchResult &&
    typeof state.dispatchResult === "object"
      ? state.dispatchResult
      : {};

  return {
    ...base,
    ...saved
  };
}

function getVehicleHourLoad(assignments, hour){
  const load = {};

  getActiveVehicles().forEach(v => {
    load[v.id] = 0;
  });

  getDispatchActuals().forEach(a => {
    if(Number(a.hour) !== Number(hour)) return;

    const vId = assignments[a.id];
    if(!vId) return;

    load[vId] = (load[vId] || 0) + 1;
  });

  return load;
}

function canMoveRideToVehicle(actualId, toVehicleId, assignments){
  const actual = getActualById(actualId);
  const vehicle = getVehicleById(toVehicleId);

  if(!actual || !vehicle) return false;

  const hour = Number(actual.hour);
  const load = getVehicleHourLoad(assignments, hour);

  const currentVehicleId = assignments[actualId];
  if(currentVehicleId === toVehicleId) return true;

  const currentLoad = load[toVehicleId] || 0;
  return currentLoad < Number(vehicle.capacity || 4);
}

function updateDispatchAssignment(actualId, toVehicleId){
  const assignments = getEffectiveAssignments();

  if(!canMoveRideToVehicle(actualId, toVehicleId, assignments)){
    alert("その車両はその時間帯で定員オーバーになります");
    return;
  }

  state.dispatchResult = state.dispatchResult || {};
  state.dispatchResult[actualId] = toVehicleId;

  saveState();
  renderDispatch();
}

/* ======================================================
PLAN BUILD
====================================================== */

function buildPlan(assignments){
  const vehiclesMap = {};

  getActiveVehicles().forEach(v => {
    vehiclesMap[v.id] = {
      id: v.id,
      area: v.area,
      homeArea: v.homeArea || v.area,
      capacity: v.capacity,
      driver: v.driver || "",
      rides: [],
      totalKm: 0
    };
  });

  getDispatchActuals().forEach(a => {
    const vId = assignments[a.id];
    if(!vId || !vehiclesMap[vId]) return;

    const cast = getCastById(a.castId);

    const ride = {
      actualId: a.id,
      castId: a.castId,
      hour: a.hour,
      name: cast?.name || "",
      area: a.area,
      km: resolveKm(a),
      status: a.status || "active"
    };

    vehiclesMap[vId].rides.push(ride);
    vehiclesMap[vId].totalKm += ride.km;
  });

  Object.values(vehiclesMap).forEach(v => {
    v.rides.sort((a,b) => {
      if(a.hour !== b.hour) return a.hour - b.hour;
      return b.km - a.km;
    });

    let lastHour = null;
    let order = 0;

    v.rides.forEach(r => {
      if(lastHour !== r.hour){
        lastHour = r.hour;
        order = 1;
      }else{
        order += 1;
      }
      r.dropOrder = order;
    });
  });

  return vehiclesMap;
}

function saveDispatchResult(assignments){
  state.dispatchResult = assignments || {};
  saveState();
}

/* ======================================================
RENDER DISPATCH
====================================================== */

function renderDispatch(){
  renderTodayVehicles();
  renderPlans();

  const box = $("planBox");
  if(!box) return;

  const rides = getDispatchActuals();
  const vehicles = getActiveVehicles();

  if(!vehicles.length){
    box.innerHTML = `<div class="muted">出勤車がありません</div>`;
    return;
  }

  if(!rides.length){
    box.innerHTML = `<div class="muted">未完了のActualがありません</div>`;
    return;
  }

  const assignments = getEffectiveAssignments();
  const plan = buildPlan(assignments);

  box.innerHTML = "";

  Object.values(plan)
    .sort((a,b) => a.id.localeCompare(b.id))
    .forEach(v => {
      const card = document.createElement("div");
      card.className = "card";

      let html = `
        <div class="cardHead">
          <div>
            <h3>車両 ${esc(v.id)}</h3>
            <div class="muted small">
              ${esc(v.area)} / 帰宅:${esc(v.homeArea)} / 定員${v.capacity}${v.driver ? " / " + esc(v.driver) : ""}
            </div>
          </div>
          <div class="row gap">
            <span class="pill">人数 ${v.rides.length}</span>
            <span class="pill">距離 ${Number(v.totalKm || 0).toFixed(1)}km</span>
          </div>
        </div>
      `;

      if(!v.rides.length){
        html += `<div class="muted small">送りなし</div>`;
      }else{
        v.rides.forEach(r => {
          const isLastMismatch =
            Number(r.hour) === 5 &&
            normalizeArea(r.area) !== normalizeArea(v.area);

          const activeVehicles = getActiveVehicles();

          const options = activeVehicles
            .slice()
            .sort((a,b) => a.id.localeCompare(b.id))
            .map(av => {
              const selected = av.id === v.id ? "selected" : "";
              return `<option value="${esc(av.id)}" ${selected}>車両 ${esc(av.id)}</option>`;
            })
            .join("");

          html += `
            <div class="row" style="justify-content:space-between;padding:8px 0;border-top:1px solid rgba(255,255,255,.06);align-items:flex-start;">
              <div class="row gap" style="flex:1;min-width:280px;">
                <span class="pill">${r.hour}時</span>
                <span class="pill">順番 ${r.dropOrder}</span>
                <span><b>${esc(r.name)}</b></span>
                <span class="muted">${esc(r.area)}</span>
                ${isLastMismatch ? `<span class="badge canceled">5時方面ズレ</span>` : ``}
              </div>

              <div style="display:flex;flex-direction:column;gap:6px;min-width:190px;">
                <div class="muted small" style="text-align:right;">${Number(r.km || 0).toFixed(1)}km</div>
                <select class="dispatchVehicleSelect" data-actual-id="${esc(r.actualId)}">
                  ${options}
                </select>
              </div>
            </div>
          `;
        });
      }

      card.innerHTML = html;
      box.appendChild(card);
    });

  box.querySelectorAll(".dispatchVehicleSelect").forEach(sel => {
    sel.addEventListener("change", () => {
      const actualId = sel.dataset.actualId;
      const toVehicleId = sel.value;
      updateDispatchAssignment(actualId, toVehicleId);
    });
  });
}

/* ======================================================
COPY
====================================================== */

function bindCopyDispatch(){
  const btn = $("btnCopy");
  if(!btn) return;

  btn.onclick = async () => {
    const assignments = getEffectiveAssignments();
    const plan = buildPlan(assignments);

    let text = `【THEMIS AI Dispatch】\n`;

    Object.values(plan)
      .sort((a,b) => a.id.localeCompare(b.id))
      .forEach(v => {
        text += `\n車両 ${v.id} (${v.area} / 帰宅:${v.homeArea})\n`;

        if(!v.rides.length){
          text += `  送りなし\n`;
        }else{
          v.rides.forEach(r => {
            text += `  ${r.hour}時 / ${r.dropOrder}件目 / ${r.name} / ${r.area} / ${Number(r.km || 0).toFixed(1)}km\n`;
          });
        }
      });

    try{
      await navigator.clipboard.writeText(text);
      alert("コピーしました");
    }catch{
      alert("コピーに失敗しました");
    }
  };
}

function bindCopyMatrix(){
  const btn = $("btnCopyMatrix");
  if(!btn) return;

  btn.onclick = async () => {
    const date = currentActualDate();
    const list = getActualsByDate(date);

    if(!list.length){
      alert("Actualがありません");
      return;
    }

    const hours = [0,1,2,3,4,5];
    const areas = [...new Set(list.map(x => x.area || "未分類"))];

    let text = `【Actual一覧 ${date}】\n`;

    hours.forEach(h => {
      text += `\n${h}時\n`;

      areas.forEach(area => {
        const items = list.filter(x =>
          x.hour === h &&
          (x.area || "未分類") === area
        );

        if(!items.length) return;

        const kmSum = items.reduce((s,x) => s + resolveKm(x), 0);

        const names = items.map(x => {
          const cast = getCastById(x.castId);
          return `${cast?.name || ""}(${resolveKm(x)}km/${x.status})`;
        }).join(" / ");

        text += `  ${area}: ${items.length}人 / ${kmSum.toFixed(1)}km / ${names}\n`;
      });
    });

    try{
      await navigator.clipboard.writeText(text);
      alert("コピーしました");
    }catch{
      alert("コピーに失敗しました");
    }
  };
}

/* ======================================================
BUTTONS
====================================================== */

function bindAutoPlanButton(){
  const btn = $("btnAutoPlan");
  if(!btn) return;

  btn.onclick = () => {
    const assignments = computeDispatch();
    state.dispatchResult = { ...assignments };
    saveState();
    renderDispatch();
    alert("自動配車しました");
  };
}

function bindFinalizeButton(){
  const btn = $("btnFinalize");
  if(!btn) return;

  btn.onclick = () => {
    const assignments = getEffectiveAssignments();
    const plan = buildPlan(assignments);

    if(!confirm("本日の距離を月間実績に反映しますか？")) return;

    Object.values(plan).forEach(v => {
      const veh = getVehicleById(v.id);
      if(!veh) return;

      const sum = v.rides.reduce((s,r) => s + Number(r.km || 0), 0);
      veh.monthKm += sum;

      if(sum > 0){
        veh.workDays += 1;
      }
    });

    saveDispatchResult(assignments);
    saveState();

    renderVehicles();
    renderDispatch();
    renderHome();

    alert("月間距離を更新しました");
  };
}

/* ======================================================
INIT
====================================================== */

function initModules(){
  initCastModule();
  initVehicleModule();
  initPlanActualModule();

  bindSettingsForm();
  bindCopyDispatch();
  bindCopyMatrix();
  bindAutoPlanButton();
  bindFinalizeButton();

  renderAll();
}

document.addEventListener("DOMContentLoaded", () => {
  const d = todayStr();

  if($("planDate")) $("planDate").value = d;
  if($("planDate_plan")) $("planDate_plan").value = d;
  if($("planDate_actual")) $("planDate_actual").value = d;

  bindTabs();
  bindGlobalButtons();
  bindDateSync();
  initModules();

  showView("home");
});
