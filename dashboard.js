const {
  SUPABASE_URL,
  SUPABASE_ANON_KEY,
  ORIGIN_LABEL,
  ORIGIN_LAT,
  ORIGIN_LNG
} = window.APP_CONFIG;

const supabaseClient = window.supabase.createClient(
  SUPABASE_URL,
  SUPABASE_ANON_KEY
);

let currentUser = null;
let currentDispatchId = null;

let editingCastId = null;
let editingVehicleId = null;
let editingPlanId = null;
let editingActualId = null;

let allCastsCache = [];
let allVehiclesCache = [];
let currentPlansCache = [];
let currentActualsCache = [];
let currentDailyReportsCache = [];
let currentMileageExportRows = [];
let activeVehicleIdsForToday = new Set();
let lastAutoDispatchRunAtMinutes = null;

const SPECIAL_LATE_NIGHT_DATES = [
  // "2026-03-20",
  // "2026-04-28"
];

const els = {
  plansTimeAreaMatrix: document.getElementById("plansTimeAreaMatrix"),
  userEmail: document.getElementById("userEmail"),
  sendLineBtn: document.getElementById("sendLineBtn"),
  originLabelText: document.getElementById("originLabelText"),
  logoutBtn: document.getElementById("logoutBtn"),

  exportAllBtn: document.getElementById("exportAllBtn"),
  importAllBtn: document.getElementById("importAllBtn"),
  importAllFileInput: document.getElementById("importAllFileInput"),
  exportCsvBtnHeader: document.getElementById("exportCsvBtnHeader"),
  openManualBtn: document.getElementById("openManualBtn"),
  dangerResetBtn: document.getElementById("dangerResetBtn"),
  resetCastsBtn: document.getElementById("resetCastsBtn"),
  resetVehiclesBtn: document.getElementById("resetVehiclesBtn"),

  homeCastCount: document.getElementById("homeCastCount"),
  homeVehicleCount: document.getElementById("homeVehicleCount"),
  homePlanCount: document.getElementById("homePlanCount"),
  homeActualCount: document.getElementById("homeActualCount"),
  homeDoneCount: document.getElementById("homeDoneCount"),
  homeCancelCount: document.getElementById("homeCancelCount"),
  homeMonthlyVehicleList: document.getElementById("homeMonthlyVehicleList"),
  resetMonthlySummaryBtn: document.getElementById("resetMonthlySummaryBtn"),

  castName: document.getElementById("castName"),
  castDistanceKm: document.getElementById("castDistanceKm"),
  castAddress: document.getElementById("castAddress"),
  castArea: document.getElementById("castArea"),
  castMemo: document.getElementById("castMemo"),
  castLatLngText: document.getElementById("castLatLngText"),
  castPhone: document.getElementById("castPhone"),
  castGeoStatus: document.getElementById("castGeoStatus"),
  castLat: document.getElementById("castLat"),
  castLng: document.getElementById("castLng"),
  saveCastBtn: document.getElementById("saveCastBtn"),
  guessAreaBtn: document.getElementById("guessAreaBtn"),
  openGoogleMapBtn: document.getElementById("openGoogleMapBtn"),
  cancelEditBtn: document.getElementById("cancelEditBtn"),
  importCsvBtn: document.getElementById("importCsvBtn"),
  exportCsvBtn: document.getElementById("exportCsvBtn"),
  csvFileInput: document.getElementById("csvFileInput"),
  castsTableBody: document.getElementById("castsTableBody"),
  castSearchName: document.getElementById("castSearchName"),
  castSearchArea: document.getElementById("castSearchArea"),
  castSearchAddress: document.getElementById("castSearchAddress"),
  castSearchPhone: document.getElementById("castSearchPhone"),
  castSearchRunBtn: document.getElementById("castSearchRunBtn"),
  castSearchResetBtn: document.getElementById("castSearchResetBtn"),
  castSearchCount: document.getElementById("castSearchCount"),
  castSearchResultWrap: document.getElementById("castSearchResultWrap"),

  vehiclePlateNumber: document.getElementById("vehiclePlateNumber"),
  vehicleArea: document.getElementById("vehicleArea"),
  vehicleHomeArea: document.getElementById("vehicleHomeArea"),
  vehicleSeatCapacity: document.getElementById("vehicleSeatCapacity"),
  vehicleDriverName: document.getElementById("vehicleDriverName"),
  vehicleLineId: document.getElementById("vehicleLineId"),
  vehicleStatus: document.getElementById("vehicleStatus"),
  vehicleMemo: document.getElementById("vehicleMemo"),
  saveVehicleBtn: document.getElementById("saveVehicleBtn"),
  cancelVehicleEditBtn: document.getElementById("cancelVehicleEditBtn"),
  importVehicleCsvBtn: document.getElementById("importVehicleCsvBtn"),
  exportVehicleCsvBtn: document.getElementById("exportVehicleCsvBtn"),
  vehicleCsvFileInput: document.getElementById("vehicleCsvFileInput"),
  vehiclesTableBody: document.getElementById("vehiclesTableBody"),
  mileageReportStartDate: document.getElementById("mileageReportStartDate"),
  mileageReportEndDate: document.getElementById("mileageReportEndDate"),
  previewMileageReportBtn: document.getElementById("previewMileageReportBtn"),
  exportMileageReportBtn: document.getElementById("exportMileageReportBtn"),
  mileageReportTableWrap: document.getElementById("mileageReportTableWrap"),

  dispatchDate: document.getElementById("dispatchDate"),
  optimizeBtn: document.getElementById("optimizeBtn"),
  confirmDailyBtn: document.getElementById("confirmDailyBtn"),
  clearActualBtn: document.getElementById("clearActualBtn"),
  checkAllVehiclesBtn: document.getElementById("checkAllVehiclesBtn"),
  uncheckAllVehiclesBtn: document.getElementById("uncheckAllVehiclesBtn"),
  clearManualLastVehicleBtn: document.getElementById("clearManualLastVehicleBtn"),
  dailyVehicleChecklist: document.getElementById("dailyVehicleChecklist"),
  manualLastVehicleInfo: document.getElementById("manualLastVehicleInfo"),
  dailyMileageInputs: document.getElementById("dailyMileageInputs"),
  saveDailyMileageBtn: document.getElementById("saveDailyMileageBtn"),
  copyResultBtn: document.getElementById("copyResultBtn"),
  dailyDispatchResult: document.getElementById("dailyDispatchResult"),

  planDate: document.getElementById("planDate"),
  exportPlansCsvBtn: document.getElementById("exportPlansCsvBtn"),
  importPlansCsvBtn: document.getElementById("importPlansCsvBtn"),
  plansCsvFileInput: document.getElementById("plansCsvFileInput"),
  clearPlansBtn: document.getElementById("clearPlansBtn"),
  planCastSelect: document.getElementById("planCastSelect"),
  planHour: document.getElementById("planHour"),
  planDistanceKm: document.getElementById("planDistanceKm"),
  planAddress: document.getElementById("planAddress"),
  planArea: document.getElementById("planArea"),
  planNote: document.getElementById("planNote"),
  savePlanBtn: document.getElementById("savePlanBtn"),
  guessPlanAreaBtn: document.getElementById("guessPlanAreaBtn"),
  cancelPlanEditBtn: document.getElementById("cancelPlanEditBtn"),
  plansGroupedTable: document.getElementById("plansGroupedTable"),
  planCastSuggest: document.getElementById("planCastSuggest"),

  actualDate: document.getElementById("actualDate"),
  addSelectedPlanBtn: document.getElementById("addSelectedPlanBtn"),
  copyActualTableBtn: document.getElementById("copyActualTableBtn"),
  planSelect: document.getElementById("planSelect"),
  castSelect: document.getElementById("castSelect"),
  castSuggest: document.getElementById("castSuggest"),
  actualHour: document.getElementById("actualHour"),
  actualDistanceKm: document.getElementById("actualDistanceKm"),
  actualStatus: document.getElementById("actualStatus"),
  actualAddress: document.getElementById("actualAddress"),
  actualArea: document.getElementById("actualArea"),
  actualNote: document.getElementById("actualNote"),
  saveActualBtn: document.getElementById("saveActualBtn"),
  guessActualAreaBtn: document.getElementById("guessActualAreaBtn"),
  cancelActualEditBtn: document.getElementById("cancelActualEditBtn"),
  actualTableWrap: document.getElementById("actualTableWrap"),
  actualTimeAreaMatrix: document.getElementById("actualTimeAreaMatrix"),

  historyList: document.getElementById("historyList")
};

function todayStr() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}


function getCurrentDispatchDateStr() {
  return els.dispatchDate?.value || els.actualDate?.value || els.planDate?.value || todayStr();
}

function getManualLastVehicleStorageKey(dateStr = getCurrentDispatchDateStr()) {
  return `THEMIS_MANUAL_LAST_VEHICLE_${dateStr}`;
}

function getManualLastVehicleState(dateStr = getCurrentDispatchDateStr()) {
  try {
    const raw = window.localStorage.getItem(getManualLastVehicleStorageKey(dateStr));
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    console.error(e);
    return null;
  }
}

function getManualLastVehicleId(dateStr = getCurrentDispatchDateStr()) {
  return Number(getManualLastVehicleState(dateStr)?.vehicle_id || 0);
}

function saveManualLastVehicleState(vehicle, dateStr = getCurrentDispatchDateStr()) {
  const payload = {
    vehicle_id: Number(vehicle?.id || 0),
    driver_name: vehicle?.driver_name || vehicle?.plate_number || "",
    home_area: normalizeAreaLabel(vehicle?.home_area || ""),
    vehicle_area: normalizeAreaLabel(vehicle?.vehicle_area || "")
  };
  window.localStorage.setItem(getManualLastVehicleStorageKey(dateStr), JSON.stringify(payload));
  renderManualLastVehicleInfo();
}

function clearManualLastVehicleState(dateStr = getCurrentDispatchDateStr()) {
  window.localStorage.removeItem(getManualLastVehicleStorageKey(dateStr));
  renderManualLastVehicleInfo();
}

function renderManualLastVehicleInfo() {
  const state = getManualLastVehicleState();
  const text = state?.vehicle_id
    ? `ラスト便車両: ${state.driver_name || "-"} / 帰宅:${normalizeAreaLabel(state.home_area || "無し")}`
    : "ラスト便車両: なし";
  if (els.manualLastVehicleInfo) els.manualLastVehicleInfo.textContent = text;
}

function setManualLastVehicle(vehicleId) {
  const vehicle = allVehiclesCache.find(v => Number(v.id) === Number(vehicleId));
  if (!vehicle) {
    alert("車両が見つかりません");
    return;
  }
  saveManualLastVehicleState(vehicle);
  renderDailyDispatchResult();
  alert(`ラスト便車両に設定しました: ${vehicle.driver_name || vehicle.plate_number || "-"}`);
}

function clearManualLastVehicle() {
  clearManualLastVehicleState();
  renderManualLastVehicleInfo();
  renderDailyDispatchResult();
  alert("ラスト便車両を解除しました");
}

function isManualLastVehicle(vehicleId, dateStr = getCurrentDispatchDateStr()) {
  return Number(vehicleId) === Number(getManualLastVehicleId(dateStr));
}



function isManualLastTripItem() {
  return false;
}

function moveManualLastItemsToEnd(rows) {
  return Array.isArray(rows) ? rows : [];
}


function formatDateTimeJa(value) {
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value || "");
  return d.toLocaleString("ja-JP");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function toNullableNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function normalizeStatus(status) {
  if (status === "done") return "done";
  if (status === "cancel") return "cancel";
  if (status === "assigned") return "assigned";
  return "pending";
}

function getStatusText(status) {
  const s = normalizeStatus(status);
  if (s === "done") return "完了";
  if (s === "cancel") return "キャンセル";
  if (s === "assigned") return "配車済";
  return "未完了";
}

function getHourLabel(hour) {
  const n = Number(hour);
  return `${n}時`;
}

function getMonthKey(dateStr) {
  const d = new Date(dateStr);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function getMonthStartStr(dateStr) {
  const d = new Date(dateStr);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-01`;
}

function getDayOfWeek(dateStr) {
  const d = new Date(dateStr);
  return d.getDay();
}

function isFridayOrSaturday(dateStr) {
  const day = getDayOfWeek(dateStr);
  return day === 5 || day === 6;
}

function isSpecialLateNightEve(dateStr) {
  return SPECIAL_LATE_NIGHT_DATES.includes(dateStr);
}

function getDefaultLastHour(dateStr) {
  const day = getDayOfWeek(dateStr);
  // 火曜-金曜は4時、土曜-月曜は5時
  if (day >= 2 && day <= 5) return 4;
  return 5;
}

function isValidLatLng(lat, lng) {
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return false;
  if (lat < -90 || lat > 90) return false;
  if (lng < -180 || lng > 180) return false;
  return true;
}

const GOOGLE_GEOCODE_CACHE_KEY = "themis_google_geocode_cache_v1";
const GOOGLE_ROUTE_DISTANCE_CACHE_KEY = "themis_google_route_distance_cache_v1";
let lastCastGeocodeKey = "";
let castGeocodeSeq = 0;
let googleMapsApiPromise = null;

function normalizeGeocodeAddressKey(address) {
  return String(address || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function setCastGeoStatus(state = "idle", message = "") {
  if (!els.castGeoStatus) return;
  els.castGeoStatus.className = `geo-status ${state}`;
  els.castGeoStatus.textContent = message || "住所入力後 Enter で座標取得 / 未取得時は座標貼り付けから手動反映できます";
}

function scheduleCastAutoGeocode() {
  const address = String(els.castAddress?.value || "").trim();
  if (!address) {
    if (els.castLat) els.castLat.value = "";
    if (els.castLng) els.castLng.value = "";
    if (els.castLatLngText) els.castLatLngText.value = "";
    lastCastGeocodeKey = "";
    castGeocodeSeq++;
    setCastGeoStatus("idle", "住所入力後 Enter で座標取得 / 未取得時は座標貼り付けから手動反映できます");
    return;
  }
  setCastGeoStatus("idle", "住所入力後 Enter で座標取得");
}

async function triggerCastAddressGeocodeNow() {
  const address = String(els.castAddress?.value || "").trim();
  if (!address) {
    setCastGeoStatus("idle", "住所を入力してください");
    return null;
  }
  const runSeq = ++castGeocodeSeq;
  const currentKey = normalizeGeocodeAddressKey(address);
  setCastGeoStatus("loading", "取得中…");
  const result = await fillCastLatLngFromAddress({ silent: true, force: currentKey !== lastCastGeocodeKey });
  if (runSeq !== castGeocodeSeq) return result;
  if (result) {
    const sourceText = result.source === "cache" ? "キャッシュから取得" : result.source === "existing" ? "入力済み座標" : "Google Geocoding";
    setCastGeoStatus("success", `✔ 座標取得済 (${sourceText})`);
  } else {
    setCastGeoStatus("error", "座標を自動取得できませんでした。住所を確認して Enter で再試行するか、座標貼り付けで手動入力してください");
  }
  return result;
}

function loadGeocodeCache() {
  try {
    return JSON.parse(localStorage.getItem(GOOGLE_GEOCODE_CACHE_KEY) || "{}");
  } catch (_) {
    return {};
  }
}

function saveGeocodeCache(cache) {
  try {
    localStorage.setItem(GOOGLE_GEOCODE_CACHE_KEY, JSON.stringify(cache || {}));
  } catch (_) {}
}


function loadRouteDistanceCache() {
  try {
    return JSON.parse(localStorage.getItem(GOOGLE_ROUTE_DISTANCE_CACHE_KEY) || "{}");
  } catch (_) {
    return {};
  }
}

function saveRouteDistanceCache(cache) {
  try {
    localStorage.setItem(GOOGLE_ROUTE_DISTANCE_CACHE_KEY, JSON.stringify(cache || {}));
  } catch (_) {}
}

function makeRouteDistanceCacheKey(address, lat, lng) {
  const latNum = toNullableNumber(lat);
  const lngNum = toNullableNumber(lng);
  if (isValidLatLng(latNum, lngNum)) return `latlng:${latNum},${lngNum}`;
  return `addr:${normalizeGeocodeAddressKey(address)}`;
}

async function getGoogleDrivingDistanceKmFromOrigin(address, lat, lng) {
  const cacheKey = makeRouteDistanceCacheKey(address, lat, lng);
  if (!cacheKey || cacheKey === 'addr:') return null;

  const cache = loadRouteDistanceCache();
  const cached = cache[cacheKey];
  if (Number.isFinite(Number(cached))) return Number(cached);

  await loadGoogleMapsApi();
  if (!window.google?.maps?.DirectionsService) return null;

  const destinationLat = toNullableNumber(lat);
  const destinationLng = toNullableNumber(lng);
  const destination = isValidLatLng(destinationLat, destinationLng)
    ? { lat: destinationLat, lng: destinationLng }
    : String(address || '').trim();

  if (!destination) return null;

  const runOnce = () => new Promise((resolve, reject) => {
    const service = new google.maps.DirectionsService();
    service.route({
      origin: { lat: ORIGIN_LAT, lng: ORIGIN_LNG },
      destination,
      travelMode: google.maps.TravelMode.DRIVING,
      region: 'JP'
    }, (result, status) => {
      if (status === 'OK') {
        const leg = result?.routes?.[0]?.legs?.[0];
        const meters = Number(leg?.distance?.value || 0);
        if (meters > 0) {
          resolve(Number((meters / 1000).toFixed(1)));
          return;
        }
        resolve(null);
        return;
      }
      if (status === 'ZERO_RESULTS' || status === 'NOT_FOUND') {
        resolve(null);
        return;
      }
      reject(new Error(`Google directions error: ${status}`));
    });
  });

  let km = null;
  try {
    km = await runOnce();
  } catch (error) {
    const transient = /UNKNOWN_ERROR|ERROR|OVER_QUERY_LIMIT/.test(String(error?.message || ""));
    if (!transient) throw error;
    await new Promise(r => setTimeout(r, 700));
    km = await runOnce();
  }

  if (Number.isFinite(Number(km)) && Number(km) > 0) {
    cache[cacheKey] = Number(km);
    saveRouteDistanceCache(cache);
    return Number(km);
  }

  return null;
}

async function resolveDistanceKmFromOrigin(address, lat, lng) {
  const latNum = toNullableNumber(lat);
  const lngNum = toNullableNumber(lng);

  try {
    const routeKm = await getGoogleDrivingDistanceKmFromOrigin(address, latNum, lngNum);
    if (Number.isFinite(Number(routeKm)) && Number(routeKm) > 0) return Number(routeKm);
  } catch (error) {
    console.warn('resolveDistanceKmFromOrigin fallback:', error);
  }

  if (isValidLatLng(latNum, lngNum)) {
    return estimateRoadKmFromStation(latNum, lngNum);
  }
  return null;
}

async function ensureCastDistanceAutoFilled(lat, lng, address, force = false) {
  if (!els.castDistanceKm) return null;
  const current = toNullableNumber(els.castDistanceKm.value);
  if (current !== null && !force) return current;
  const autoKm = await resolveDistanceKmFromOrigin(address, lat, lng);
  if (autoKm !== null) els.castDistanceKm.value = String(autoKm);
  return autoKm;
}

async function resolveDistanceKmForCastRecord(cast, overrideAddress = '') {
  let distance = toNullableNumber(cast?.distance_km);
  if (distance !== null) return distance;

  const lat = toNullableNumber(cast?.latitude);
  const lng = toNullableNumber(cast?.longitude);
  const address = String(overrideAddress || cast?.address || '').trim();

  if (!address && !isValidLatLng(lat, lng)) return null;
  return await resolveDistanceKmFromOrigin(address, lat, lng);
}

function loadGoogleMapsApi() {
  if (window.google?.maps?.Geocoder) return Promise.resolve(window.google.maps);
  if (googleMapsApiPromise) return googleMapsApiPromise;

  const apiKey = String(window.APP_CONFIG?.GOOGLE_MAPS_API_KEY || "").trim();
  if (!apiKey) {
    return Promise.reject(new Error("GOOGLE_MAPS_API_KEY が未設定です"));
  }

  googleMapsApiPromise = new Promise((resolve, reject) => {
    const callbackName = "__themisGoogleMapsInit";
    const timeout = window.setTimeout(() => {
      cleanup();
      reject(new Error("Google Maps API の読み込みがタイムアウトしました"));
    }, 15000);

    function cleanup() {
      window.clearTimeout(timeout);
      try { delete window[callbackName]; } catch (_) { window[callbackName] = undefined; }
    }

    if (window.google?.maps?.Geocoder) {
      cleanup();
      resolve(window.google.maps);
      return;
    }

    const existing = document.getElementById("googleMapsApiScript");
    window[callbackName] = () => {
      cleanup();
      if (window.google?.maps?.Geocoder) {
        resolve(window.google.maps);
      } else {
        reject(new Error("Google Maps API は読み込まれましたが Geocoder が使えません"));
      }
    };

    if (existing) {
      existing.addEventListener("error", () => {
        cleanup();
        reject(new Error("Google Maps API の読み込みに失敗しました"));
      }, { once: true });
      return;
    }

    const script = document.createElement("script");
    script.id = "googleMapsApiScript";
    script.async = true;
    script.defer = true;
    script.onerror = () => {
      cleanup();
      reject(new Error("Google Maps API の読み込みに失敗しました"));
    };
    script.src = `https://maps.googleapis.com/maps/api/js?key=${encodeURIComponent(apiKey)}&language=ja&region=JP&callback=${callbackName}`;
    document.head.appendChild(script);
  }).catch((error) => {
    googleMapsApiPromise = null;
    throw error;
  });

  return googleMapsApiPromise;
}

async function geocodeAddressGoogle(address) {
  const normalizedAddress = String(address || "").trim();
  if (!normalizedAddress) return null;

  const key = normalizeGeocodeAddressKey(normalizedAddress);
  const cache = loadGeocodeCache();
  const cached = cache[key];
  if (cached && isValidLatLng(Number(cached.lat), Number(cached.lng))) {
    return {
      lat: Number(cached.lat),
      lng: Number(cached.lng),
      source: "cache"
    };
  }

  await loadGoogleMapsApi();

  async function runOnce() {
    return await new Promise((resolve, reject) => {
      const geocoder = new google.maps.Geocoder();
      geocoder.geocode({
        address: normalizedAddress,
        region: "JP",
        componentRestrictions: { country: "JP" }
      }, (results, status) => {
        if (status === "OK") {
          const first = Array.isArray(results) ? results[0] : null;
          if (!first?.geometry?.location) {
            resolve(null);
            return;
          }
          const loc = first.geometry.location;
          resolve({
            lat: Number(loc.lat()),
            lng: Number(loc.lng()),
            source: "api"
          });
          return;
        }
        if (status === "ZERO_RESULTS") {
          resolve(null);
          return;
        }
        reject(new Error(`Google geocode error: ${status}`));
      });
    });
  }

  let result = null;
  try {
    result = await runOnce();
  } catch (error) {
    const transient = /OVER_QUERY_LIMIT|UNKNOWN_ERROR|ERROR/.test(String(error?.message || ""));
    if (!transient) throw error;
    await new Promise(r => setTimeout(r, 700));
    result = await runOnce();
  }

  if (!result || !isValidLatLng(result.lat, result.lng)) return null;

  cache[key] = { lat: result.lat, lng: result.lng, ts: Date.now() };
  saveGeocodeCache(cache);
  return result;
}

async function fillCastLatLngFromAddress(options = {}) {
  const silent = Boolean(options.silent);
  const address = String(els.castAddress?.value || "").trim();

  if (!address) {
    if (!silent) alert("住所を入力してください");
    return null;
  }

  const currentKey = normalizeGeocodeAddressKey(address);
  const existingLat = toNullableNumber(els.castLat?.value);
  const existingLng = toNullableNumber(els.castLng?.value);
  if (isValidLatLng(existingLat, existingLng) && !options.force && currentKey === lastCastGeocodeKey) {
    return { lat: existingLat, lng: existingLng, source: "existing" };
  }

  try {
    const result = await geocodeAddressGoogle(address);
    if (!result) {
      if (!silent) alert("住所から座標を取得できませんでした。手動入力してください");
      return null;
    }

    if (els.castLat) els.castLat.value = String(result.lat);
    if (els.castLng) els.castLng.value = String(result.lng);
    if (els.castLatLngText) els.castLatLngText.value = `${result.lat},${result.lng}`;
    if (els.castArea) {
      els.castArea.value = normalizeAreaLabel(guessArea(result.lat, result.lng, address));
    }
    if (els.castDistanceKm && !String(els.castDistanceKm.value || "").trim()) {
      const autoKm = await resolveDistanceKmFromOrigin(address, result.lat, result.lng);
      if (autoKm !== null) els.castDistanceKm.value = String(autoKm);
    }

    lastCastGeocodeKey = currentKey;
    if (!silent) {
      const label = result.source === "cache" ? "キャッシュ" : "Google Geocoding";
      alert(`${label}で座標を取得しました`);
    }
    return result;
  } catch (error) {
    console.error("fillCastLatLngFromAddress error:", error);
    if (!silent) alert(`住所から座標取得できませんでした。${error.message || "時間をおいて再試行するか、手動で入力してください"}`);
    return null;
  }
}

function parseLatLngText(text) {
  const raw = String(text || "").trim();
  if (!raw) return null;

  const normalized = raw
    .replace(/[　 ]+/g, "")
    .replace(/，/g, ",")
    .replace(/、/g, ",")
    .replace(/緯度[:：]/g, "")
    .replace(/経度[:：]/g, "")
    .replace(/latitude[:=]/gi, "")
    .replace(/longitude[:=]/gi, "");

  const patterns = [
    /@(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/,
    /q=(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/,
    /ll=(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/,
    /(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/
  ];

  for (const pattern of patterns) {
    const match = normalized.match(pattern);
    if (!match) continue;
    const lat = Number(match[1]);
    const lng = Number(match[2]);
    if (isValidLatLng(lat, lng)) return { lat, lng };
  }
  return null;
}

function haversineKm(lat1, lng1, lat2, lng2) {
  const R = 6371;
  const dLat = ((lat2 - lat1) * Math.PI) / 180;
  const dLng = ((lng2 - lng1) * Math.PI) / 180;
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos((lat1 * Math.PI) / 180) *
      Math.cos((lat2 * Math.PI) / 180) *
      Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

function estimateRoadKmFromStation(lat, lng) {
  const straight = haversineKm(ORIGIN_LAT, ORIGIN_LNG, lat, lng);
  return Number((straight * 1.22).toFixed(1));
}

function estimateRoadKmBetweenPoints(lat1, lng1, lat2, lng2) {
  if (!isValidLatLng(lat1, lng1) || !isValidLatLng(lat2, lng2)) return 0;
  const straight = haversineKm(lat1, lng1, lat2, lng2);
  return Number((straight * 1.22).toFixed(1));
}

function getItemLatLng(item) {
  const lat = toNullableNumber(item?.casts?.latitude);
  const lng = toNullableNumber(item?.casts?.longitude);
  if (isValidLatLng(lat, lng)) return { lat, lng };
  return null;
}

function sortItemsByNearestRoute(items) {
  const remaining = [...items];
  const sorted = [];

  let currentLat = ORIGIN_LAT;
  let currentLng = ORIGIN_LNG;

  while (remaining.length) {
    let bestIndex = 0;
    let bestScore = Infinity;

    remaining.forEach((item, index) => {
      const point = getItemLatLng(item);

      let score;
      if (point) {
        score = estimateRoadKmBetweenPoints(currentLat, currentLng, point.lat, point.lng);
      } else {
        score = Number(item.distance_km || 999999);
      }

      if (score < bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });

    const picked = remaining.splice(bestIndex, 1)[0];
    sorted.push(picked);

    const pickedPoint = getItemLatLng(picked);
    if (pickedPoint) {
      currentLat = pickedPoint.lat;
      currentLng = pickedPoint.lng;
    }
  }

  return sorted;
}

function calculateRouteDistance(items) {
  if (!items.length) return 0;

  let total = 0;
  let currentLat = ORIGIN_LAT;
  let currentLng = ORIGIN_LNG;

  items.forEach(item => {
    const point = getItemLatLng(item);

    if (point) {
      total += estimateRoadKmBetweenPoints(currentLat, currentLng, point.lat, point.lng);
      currentLat = point.lat;
      currentLng = point.lng;
    } else {
      total += Number(item.distance_km || 0);
    }
  });

  return Number(total.toFixed(1));
}


function normalizeAddressText(address) {
  return String(address || "")
    .trim()
    .replace(/[　\s]+/g, "")
    .replace(/ヶ/g, "ケ")
    .replace(/之/g, "の")
    .replace(/−/g, "-")
    .replace(/ー/g, "-");
}

function detectPrefecture(address) {
  const a = normalizeAddressText(address);

  if (a.includes("東京都")) return "東京";
  if (a.includes("埼玉県")) return "埼玉";
  if (a.includes("千葉県")) return "千葉";
  if (a.includes("茨城県")) return "茨城";

  return "";
}

function detectCityTownArea(address) {
  const raw = String(address || "").trim();
  const a = normalizeAddressText(raw);
  if (!a) return "";

  const tokyoMatch = a.match(
    /(足立区|葛飾区|江戸川区|墨田区|江東区|荒川区|台東区|中央区|千代田区|港区|新宿区|文京区|品川区|目黒区|大田区|世田谷区|渋谷区|中野区|杉並区|豊島区|北区|板橋区|練馬区)(.+)/
  );
  if (tokyoMatch) {
    const wardRaw = tokyoMatch[1];
    let townRaw = tokyoMatch[2] || "";

    townRaw = townRaw
      .replace(/丁目.*$/, "")
      .replace(/番地.*$/, "")
      .replace(/番.*$/, "")
      .replace(/号.*$/, "")
      .replace(/[0-9０-９]+/g, "")
      .replace(/[-‐-‒–—―ー－]+/g, "")
      .replace(/日本$/, "")
      .trim();

    const ward = wardRaw.replace("区", "");
    return townRaw ? `${ward} ${townRaw}方面` : `${ward}方面`;
  }

  const chibaWardMatch = a.match(
    /(千葉市中央区|千葉市花見川区|千葉市稲毛区|千葉市若葉区|千葉市緑区|千葉市美浜区)(.+)/
  );
  if (chibaWardMatch) {
    const wardRaw = chibaWardMatch[1];
    let townRaw = chibaWardMatch[2] || "";

    townRaw = townRaw
      .replace(/丁目.*$/, "")
      .replace(/番地.*$/, "")
      .replace(/番.*$/, "")
      .replace(/号.*$/, "")
      .replace(/[0-9０-９]+/g, "")
      .replace(/[-‐-‒–—―ー－]+/g, "")
      .replace(/日本$/, "")
      .trim();

    const ward = wardRaw.replace("千葉市", "");

    const chibaWardAliasMap = [
      { ward: "中央区", from: ["蘇我", "今井"], to: "蘇我" },
      { ward: "中央区", from: ["千葉寺", "末広", "青葉町"], to: "千葉寺" },
      { ward: "中央区", from: ["長洲", "港町", "本千葉町"], to: "本千葉" },

      { ward: "花見川区", from: ["幕張本郷", "幕張町"], to: "幕張本郷" },
      { ward: "花見川区", from: ["検見川町", "新検見川"], to: "検見川" },

      { ward: "稲毛区", from: ["稲毛東", "稲毛台町", "小仲台"], to: "稲毛" },
      { ward: "稲毛区", from: ["穴川", "天台"], to: "穴川" },

      { ward: "若葉区", from: ["都賀", "西都賀"], to: "都賀" },
      { ward: "若葉区", from: ["桜木", "加曽利"], to: "桜木" },

      { ward: "緑区", from: ["おゆみ野", "おゆみ野南", "おゆみ野中央"], to: "おゆみ野" },
      { ward: "緑区", from: ["誉田町", "誉田"], to: "誉田" },

      { ward: "美浜区", from: ["ひび野", "打瀬", "中瀬"], to: "海浜幕張" },
      { ward: "美浜区", from: ["幕張西"], to: "幕張" },
      { ward: "美浜区", from: ["稲毛海岸", "高洲", "高浜"], to: "稲毛海岸" }
    ];

    for (const row of chibaWardAliasMap) {
      if (row.ward !== ward) continue;
      for (const key of row.from) {
        if (townRaw.includes(key)) {
          return `${ward} ${row.to}方面`;
        }
      }
    }

    return townRaw ? `${ward} ${townRaw}方面` : `${ward}方面`;
  }

  const cityMatch = a.match(
    /(松戸市|柏市|流山市|我孫子市|野田市|市川市|鎌ケ谷市|鎌ヶ谷市|船橋市|三郷市|八潮市|草加市|越谷市|吉川市|守谷市|取手市|つくば市|牛久市|龍ケ崎市|龍ヶ崎市)(.+)/
  );
  if (!cityMatch) return "";

  const cityRaw = cityMatch[1];
  let townRaw = cityMatch[2] || "";

  townRaw = townRaw
    .replace(/丁目.*$/, "")
    .replace(/番地.*$/, "")
    .replace(/番.*$/, "")
    .replace(/号.*$/, "")
    .replace(/[0-9０-９]+/g, "")
    .replace(/[-‐-‒–—―ー－]+/g, "")
    .replace(/日本$/, "")
    .replace(/^千葉県|^埼玉県|^東京都|^茨城県/, "")
    .trim();

  const city = cityRaw.replace(/(市|区)$/, "");

  const aliasMap = [
    { city: "柏", from: ["若柴"], to: "若柴" },
    { city: "流山", from: ["おおたかの森西", "おおたかの森東", "おおたかの森南", "おおたかの森北"], to: "おおたかの森" },

    { city: "市川", from: ["八幡", "東菅野"], to: "八幡" },
    { city: "市川", from: ["南八幡"], to: "本八幡" },
    { city: "市川", from: ["鬼高", "高石神", "中山"], to: "下総中山" },
    { city: "市川", from: ["行徳駅前", "末広", "宝", "湊"], to: "行徳" },
    { city: "市川", from: ["妙典", "塩焼"], to: "妙典" },
    { city: "市川", from: ["南行徳", "香取", "相之川"], to: "南行徳" },
    { city: "市川", from: ["堀之内", "中国分", "北国分"], to: "堀之内" },
    { city: "市川", from: ["国府台", "真間", "菅野"], to: "国府台" },

    { city: "船橋", from: ["西船", "西船橋", "本郷町"], to: "西船橋" },
    { city: "船橋", from: ["東船橋", "市場"], to: "東船橋" },
    { city: "船橋", from: ["前原西", "前原東", "津田沼"], to: "津田沼" },
    { city: "船橋", from: ["習志野台", "高根台"], to: "習志野台" },
    { city: "船橋", from: ["北習志野"], to: "北習志野" },
    { city: "船橋", from: ["浜町", "若松", "高瀬町"], to: "南船橋" },
    { city: "船橋", from: ["二和東", "二和西", "咲が丘"], to: "二和" },

    { city: "三郷", from: ["中央"], to: "中央" },
    { city: "草加", from: ["谷塚仲町", "谷塚上町", "谷塚町"], to: "谷塚" },
    { city: "守谷", from: ["百合ケ丘", "百合ヶ丘"], to: "百合ヶ丘" },
    { city: "つくば", from: ["研究学園"], to: "研究学園" },
    { city: "牛久", from: ["ひたち野西", "ひたち野東"], to: "ひたち野うしく" },

    { city: "松戸", from: ["八ケ崎", "八ヶ崎"], to: "八ヶ崎" },
    { city: "松戸", from: ["日暮", "金ケ作", "常盤平西窪町"], to: "八柱" },
    { city: "松戸", from: ["常盤平", "常盤平双葉町"], to: "常盤平" },
    { city: "松戸", from: ["新松戸", "幸谷"], to: "新松戸" },
    { city: "松戸", from: ["馬橋", "西馬橋", "中和倉"], to: "馬橋" },
    { city: "松戸", from: ["五香", "五香西", "五香南"], to: "五香" },
    { city: "松戸", from: ["小金原", "栗ケ沢", "根木内"], to: "小金原" },
    { city: "松戸", from: ["松飛台", "串崎新田"], to: "松飛台" }
  ];

  for (const row of aliasMap) {
    if (row.city !== city) continue;
    for (const key of row.from) {
      if (townRaw.includes(key)) {
        return `${city} ${row.to}方面`;
      }
    }
  }

  if (townRaw) {
    return `${city} ${townRaw}方面`;
  }

  return `${city}方面`;
}

function classifyAreaByAddress(address) {
  const cityTown = detectCityTownArea(address);
  if (cityTown) return cityTown;

  const a = normalizeAddressText(address);
  if (!a) return "";

  const areaMap = [
    { area: "柏方面", keywords: ["南柏", "柏", "北柏"] },
    { area: "柏の葉方面", keywords: ["柏の葉", "柏たなか"] },
    { area: "流山方面", keywords: ["流山", "南流山"] },
    { area: "野田方面", keywords: ["野田", "梅郷", "運河", "川間", "愛宕", "清水公園"] },
    { area: "おおたかの森方面", keywords: ["流山おおたかの森", "おおたかの森"] },
    { area: "我孫子方面", keywords: ["我孫子", "天王台", "湖北"] },
    { area: "鎌ヶ谷方面", keywords: ["新鎌ケ谷", "新鎌ヶ谷", "鎌ケ谷", "鎌ヶ谷"] },
    { area: "船橋方面", keywords: ["船橋", "西船橋"] },
    { area: "市川方面", keywords: ["市川", "本八幡", "下総中山", "妙典", "行徳"] },
    { area: "千葉方面", keywords: ["千葉市", "千葉", "蘇我", "稲毛", "幕張", "海浜幕張", "都賀"] },

    { area: "足立方面", keywords: ["足立区"] },
    { area: "葛飾方面", keywords: ["葛飾区"] },
    { area: "江戸川方面", keywords: ["江戸川区"] },
    { area: "墨田方面", keywords: ["墨田区"] },
    { area: "江東方面", keywords: ["江東区"] },
    { area: "荒川方面", keywords: ["荒川区"] },
    { area: "台東方面", keywords: ["台東区"] },

    { area: "三郷方面", keywords: ["三郷中央", "三郷"] },
    { area: "吉川方面", keywords: ["吉川"] },
    { area: "八潮方面", keywords: ["八潮"] },
    { area: "谷塚方面", keywords: ["谷塚"] },
    { area: "草加方面", keywords: ["草加"] },
    { area: "越谷方面", keywords: ["新越谷", "越谷"] },

    { area: "藤代方面", keywords: ["藤代"] },
    { area: "取手方面", keywords: ["取手"] },
    { area: "守谷方面", keywords: ["守谷"] },
    { area: "つくば方面", keywords: ["研究学園", "つくば"] },
    { area: "牛久方面", keywords: ["ひたち野うしく", "牛久"] }
  ];

  for (const row of areaMap) {
    if (row.keywords.some(keyword => a.includes(keyword))) {
      return row.area;
    }
  }

  return "";
}

function getDirection8(lat, lng) {
  if (!isValidLatLng(lat, lng)) return "";

  const dLat = lat - ORIGIN_LAT;
  const dLng = lng - ORIGIN_LNG;
  const angle = Math.atan2(dLat, dLng) * 180 / Math.PI;

  if (angle >= -22.5 && angle < 22.5) return "東";
  if (angle >= 22.5 && angle < 67.5) return "北東";
  if (angle >= 67.5 && angle < 112.5) return "北";
  if (angle >= 112.5 && angle < 157.5) return "北西";
  if (angle >= -67.5 && angle < -22.5) return "南東";
  if (angle >= -112.5 && angle < -67.5) return "南";
  if (angle >= -157.5 && angle < -112.5) return "南西";
  return "西";
}

function classifyAreaByLatLng(lat, lng) {
  if (!isValidLatLng(lat, lng)) return "";

  if (lat >= 35.79 && lat <= 35.86 && lng >= 139.90 && lng <= 139.97) {
    return "松戸方面";
  }

  if (lat >= 35.79 && lat <= 35.88 && lng >= 139.76 && lng <= 139.85) {
    if (lat >= 35.84) return "草加方面";
    if (lat >= 35.81) return "谷塚方面";
    return "八潮方面";
  }

  if (lat >= 35.84 && lat <= 35.92 && lng >= 139.84 && lng <= 139.92) {
    return "三郷方面";
  }

  if (lat >= 35.86 && lat <= 35.91 && lng >= 139.82 && lng <= 139.87) {
    return "吉川方面";
  }

  if (lat >= 35.84 && lat <= 35.89 && lng >= 139.92 && lng <= 139.98) {
    return "柏方面";
  }

  if (lat >= 35.85 && lat <= 35.91 && lng > 139.98 && lng <= 140.05) {
    return "柏の葉方面";
  }

  if (lat >= 35.92 && lat <= 35.99 && lng >= 139.84 && lng <= 139.91) {
    return "野田方面";
  }

  if (lat >= 35.84 && lat <= 35.90 && lng >= 139.88 && lng <= 139.95) {
    return "流山方面";
  }

  if (lat >= 35.85 && lat <= 35.89 && lng > 140.00 && lng <= 140.08) {
    return "我孫子方面";
  }

  if (lat >= 35.70 && lat <= 35.78 && lng >= 139.78 && lng <= 139.86) {
    return "墨田方面";
  }
  if (lat >= 35.73 && lat <= 35.80 && lng >= 139.80 && lng <= 139.88) {
    return "足立方面";
  }
  if (lat >= 35.75 && lat <= 35.79 && lng >= 139.84 && lng <= 139.89) {
    return "葛飾方面";
  }
  if (lat >= 35.67 && lat <= 35.72 && lng >= 139.80 && lng <= 139.86) {
    return "江東方面";
  }

  if (lat >= 35.90 && lat <= 36.02 && lng >= 140.00 && lng <= 140.08) {
    return "藤代方面";
  }
  if (lat >= 35.88 && lat <= 35.95 && lng >= 140.03 && lng <= 140.10) {
    return "取手方面";
  }
  if (lat >= 35.93 && lat <= 36.02 && lng >= 139.97 && lng <= 140.05) {
    return "守谷方面";
  }
  if (lat >= 36.02 && lat <= 36.10 && lng >= 140.05 && lng <= 140.15) {
    return "つくば方面";
  }
  if (lat >= 35.95 && lat <= 36.02 && lng >= 140.12 && lng <= 140.20) {
    return "牛久方面";
  }

  return "";
}

function guessArea(lat, lng, address = "") {
  const byAddress = classifyAreaByAddress(address);
  if (byAddress) return byAddress;

  const byLatLng = classifyAreaByLatLng(lat, lng);
  if (byLatLng) return byLatLng;

  const pref = detectPrefecture(address);
  const dir = getDirection8(lat, lng);

  if (pref && dir) return `${pref}${dir}方面`;
  if (pref) return `${pref}方面`;
  if (dir) return `${dir}方面`;

  return "周辺";
}


function normalizeCastInputValue(value) {
  return String(value || "").trim();
}

function findCastByInputValue(value) {
  const normalized = normalizeCastInputValue(value);
  if (!normalized) return null;

  return (
    allCastsCache.find(c => String(c.name || "").trim() === normalized) ||
    allCastsCache.find(
      c => `${String(c.name || "").trim()} | ${normalizeAreaLabel(c.area || "-")}` === normalized
    ) ||
    null
  );
}

function normalizeAreaLabel(area) {
  const value = String(area || "").trim();
  if (!value) return "無し";
  return value;
}

function getAreaDisplayGroup(area) {
  const a = normalizeAreaLabel(area);

  if (!a || a === "無し") return "その他";

  if (
    a.includes("松戸") || a.includes("新松戸") || a.includes("馬橋") ||
    a.includes("八柱") || a.includes("北松戸") || a.includes("小金") ||
    a.includes("常盤平") || a.includes("みのり台")
  ) return "松戸方面";

  if (a.includes("柏")) return "柏方面";

  if (
    a.includes("我孫子") || a.includes("取手") || a.includes("藤代") ||
    a.includes("守谷") || a.includes("牛久") || a.includes("つくば")
  ) return "我孫子・取手方面";

  if (
    a.includes("葛飾") || a.includes("足立") || a.includes("江戸川") ||
    a.includes("墨田") || a.includes("江東") || a.includes("荒川") ||
    a.includes("台東")
  ) return "東京東方面";

  if (
    a.includes("市川") || a.includes("船橋") || a.includes("鎌ヶ谷") ||
    a.includes("鎌ケ谷") || a.includes("千葉")
  ) return "市川・船橋方面";

  if (
    a.includes("三郷") || a.includes("八潮") || a.includes("草加") ||
    a.includes("吉川") || a.includes("越谷")
  ) return "埼玉方面";

  if (a.includes("流山") || a.includes("野田") || a.includes("柏の葉")) return "流山・野田方面";

  return "その他";
}


function getGroupedAreaHeaderHtml(area) {
  const detailArea = normalizeAreaLabel(area || "無し");
  const displayGroup = getAreaDisplayGroup(detailArea);

  if (!detailArea || detailArea === "無し") {
    return `<span class="group-main">${escapeHtml(displayGroup)}</span>`;
  }

  if (displayGroup === detailArea) {
    return `<span class="group-main">${escapeHtml(displayGroup)}</span>`;
  }

  return `
    <div class="group-area-stack">
      <span class="group-main">${escapeHtml(displayGroup)}</span>
      <span class="group-sub">${escapeHtml(detailArea)}</span>
    </div>
  `;
}

function getGroupedAreasByDisplay(items, areaGetter) {
  const ordered = [];
  const seen = new Set();

  for (const item of items) {
    const detailArea = normalizeAreaLabel(areaGetter(item));
    const key = `${getAreaDisplayGroup(detailArea)}__${detailArea}`;
    if (seen.has(key)) continue;
    seen.add(key);
    ordered.push({
      displayGroup: getAreaDisplayGroup(detailArea),
      detailArea
    });
  }

  ordered.sort((a, b) => {
    const groupCompare = a.displayGroup.localeCompare(b.displayGroup, "ja");
    if (groupCompare !== 0) return groupCompare;
    return a.detailArea.localeCompare(b.detailArea, "ja");
  });

  return ordered;
}

const AREA_CANONICAL_PATTERNS = [
  ["松戸近郊", ["松戸近郊", "松戸方面", "松戸"]],
  ["葛飾方面", ["葛飾方面", "葛飾区", "葛飾"]],
  ["足立方面", ["足立方面", "足立区", "足立"]],
  ["江戸川方面", ["江戸川方面", "江戸川区", "江戸川"]],
  ["市川方面", ["市川方面", "市川", "本八幡", "妙典", "行徳", "下総中山"]],
  ["船橋方面", ["船橋方面", "船橋", "習志野", "西船橋"]],
  ["鎌ヶ谷方面", ["鎌ヶ谷方面", "鎌ケ谷方面", "鎌ヶ谷", "鎌ケ谷", "新鎌ヶ谷", "新鎌ケ谷"]],
  ["我孫子方面", ["我孫子方面", "我孫子", "天王台", "湖北"]],
  ["取手方面", ["取手方面", "取手", "藤代方面", "藤代", "桐木"]],
  ["藤代方面", ["藤代方面", "藤代", "取手", "桐木"]],
  ["守谷方面", ["守谷方面", "守谷"]],
  ["柏方面", ["柏方面", "柏", "南柏", "北柏"]],
  ["柏の葉方面", ["柏の葉方面", "柏の葉", "柏たなか"]],
  ["流山方面", ["流山方面", "流山", "南流山", "おおたかの森"]],
  ["野田方面", ["野田方面", "野田", "運河", "梅郷", "川間"]],
  ["三郷方面", ["三郷方面", "三郷"]],
  ["八潮方面", ["八潮方面", "八潮"]],
  ["草加方面", ["草加方面", "草加", "谷塚方面", "谷塚"]],
  ["吉川方面", ["吉川方面", "吉川"]],
  ["越谷方面", ["越谷方面", "越谷", "新越谷"]],
  ["千葉方面", ["千葉方面", "千葉", "幕張", "蘇我", "稲毛", "都賀"]]
];

const AREA_AFFINITY_MAP = {
  "松戸近郊": { "葛飾方面": 80, "市川方面": 72, "柏方面": 60, "三郷方面": 62, "足立方面": 58 },
  "葛飾方面": { "松戸近郊": 80, "足立方面": 62, "江戸川方面": 58, "市川方面": 55 },
  "足立方面": { "葛飾方面": 62, "松戸近郊": 58, "八潮方面": 55, "草加方面": 52 },
  "江戸川方面": { "葛飾方面": 58, "市川方面": 54, "船橋方面": 46 },
  "市川方面": { "葛飾方面": 55, "松戸近郊": 72, "船橋方面": 68, "鎌ヶ谷方面": 66, "江戸川方面": 54 },
  "船橋方面": { "市川方面": 68, "鎌ヶ谷方面": 76, "千葉方面": 58, "江戸川方面": 46 },
  "鎌ヶ谷方面": { "船橋方面": 76, "市川方面": 66, "柏方面": 56, "我孫子方面": 42 },
  "我孫子方面": { "取手方面": 88, "藤代方面": 84, "柏方面": 70, "守谷方面": 60, "鎌ヶ谷方面": 42 },
  "取手方面": { "我孫子方面": 88, "藤代方面": 92, "守谷方面": 72, "柏方面": 44 },
  "藤代方面": { "我孫子方面": 84, "取手方面": 92, "守谷方面": 70 },
  "守谷方面": { "取手方面": 72, "藤代方面": 70, "我孫子方面": 60, "つくば方面": 62 },
  "柏方面": { "我孫子方面": 70, "流山方面": 66, "柏の葉方面": 64, "野田方面": 56, "鎌ヶ谷方面": 56, "松戸近郊": 60, "取手方面": 44 },
  "柏の葉方面": { "柏方面": 64, "流山方面": 62, "野田方面": 58 },
  "流山方面": { "柏方面": 66, "柏の葉方面": 62, "野田方面": 58, "三郷方面": 56, "吉川方面": 54 },
  "野田方面": { "柏方面": 56, "柏の葉方面": 58, "流山方面": 58, "吉川方面": 52 },
  "三郷方面": { "松戸近郊": 62, "流山方面": 56, "八潮方面": 62, "吉川方面": 54 },
  "八潮方面": { "三郷方面": 62, "草加方面": 56, "足立方面": 55 },
  "草加方面": { "八潮方面": 56, "足立方面": 52, "越谷方面": 52 },
  "吉川方面": { "流山方面": 54, "野田方面": 52, "三郷方面": 54, "越谷方面": 52 },
  "越谷方面": { "草加方面": 52, "吉川方面": 52 },
  "千葉方面": { "船橋方面": 58 }
};

const AREA_DIRECTION_MAP = {
  "松戸近郊": "CENTER",
  "葛飾方面": "W",
  "足立方面": "W",
  "江戸川方面": "SW",
  "市川方面": "S",
  "船橋方面": "SE",
  "鎌ヶ谷方面": "SE",
  "我孫子方面": "NE",
  "取手方面": "NE",
  "藤代方面": "E",
  "守谷方面": "E",
  "柏方面": "E",
  "柏の葉方面": "NE",
  "流山方面": "N",
  "野田方面": "N",
  "三郷方面": "NW",
  "八潮方面": "W",
  "草加方面": "NW",
  "吉川方面": "N",
  "越谷方面": "NW",
  "千葉方面": "SE"
};

const DIRECTION_RING = ["E", "NE", "N", "NW", "W", "SW", "S", "SE"];

const HOME_ALLOWED_AREA_MAP = {
  "葛飾方面": ["葛飾方面", "松戸近郊", "足立方面", "江戸川方面", "市川方面", "八潮方面", "三郷方面"],
  "足立方面": ["足立方面", "葛飾方面", "八潮方面", "草加方面", "松戸近郊"],
  "江戸川方面": ["江戸川方面", "葛飾方面", "市川方面", "船橋方面"],
  "市川方面": ["市川方面", "江戸川方面", "船橋方面", "鎌ヶ谷方面", "松戸近郊"],
  "船橋方面": ["船橋方面", "鎌ヶ谷方面", "市川方面", "千葉方面"],
  "鎌ヶ谷方面": ["鎌ヶ谷方面", "船橋方面", "市川方面", "柏方面", "松戸近郊"],
  "我孫子方面": ["我孫子方面", "取手方面", "藤代方面", "守谷方面", "柏方面"],
  "取手方面": ["取手方面", "藤代方面", "我孫子方面", "守谷方面", "柏方面"],
  "藤代方面": ["藤代方面", "取手方面", "守谷方面", "我孫子方面"],
  "守谷方面": ["守谷方面", "取手方面", "藤代方面", "我孫子方面", "柏方面"],
  "柏方面": ["柏方面", "柏の葉方面", "流山方面", "我孫子方面", "鎌ヶ谷方面", "松戸近郊"],
  "柏の葉方面": ["柏の葉方面", "柏方面", "流山方面", "野田方面"],
  "流山方面": ["流山方面", "柏方面", "野田方面", "三郷方面", "吉川方面", "松戸近郊"],
  "野田方面": ["野田方面", "流山方面", "柏方面", "吉川方面", "柏の葉方面"],
  "三郷方面": ["三郷方面", "八潮方面", "松戸近郊", "吉川方面", "流山方面"],
  "八潮方面": ["八潮方面", "三郷方面", "足立方面", "葛飾方面", "草加方面"],
  "草加方面": ["草加方面", "足立方面", "八潮方面", "越谷方面"],
  "吉川方面": ["吉川方面", "流山方面", "野田方面", "三郷方面", "越谷方面"],
  "越谷方面": ["越谷方面", "草加方面", "吉川方面"],
  "松戸近郊": ["松戸近郊", "葛飾方面", "市川方面", "柏方面", "三郷方面", "足立方面", "流山方面"],
  "千葉方面": ["千葉方面", "船橋方面"]
};

const ROUTE_FLOW_MAP = {
  "松戸近郊": ["流山方面", "吉川方面", "三郷方面", "柏方面", "我孫子方面", "市川方面", "葛飾方面"],
  "流山方面": ["松戸近郊", "吉川方面", "野田方面", "柏方面", "柏の葉方面"],
  "吉川方面": ["松戸近郊", "流山方面", "三郷方面", "野田方面"],
  "三郷方面": ["松戸近郊", "吉川方面", "八潮方面", "流山方面"],
  "柏方面": ["松戸近郊", "我孫子方面", "流山方面", "柏の葉方面", "取手方面"],
  "柏の葉方面": ["柏方面", "流山方面", "野田方面"],
  "我孫子方面": ["松戸近郊", "柏方面", "取手方面", "藤代方面", "守谷方面"],
  "取手方面": ["我孫子方面", "藤代方面", "守谷方面", "柏方面"],
  "藤代方面": ["我孫子方面", "取手方面", "守谷方面"],
  "守谷方面": ["我孫子方面", "取手方面", "藤代方面", "柏方面"],
  "市川方面": ["松戸近郊", "鎌ヶ谷方面", "船橋方面", "葛飾方面", "江戸川方面"],
  "鎌ヶ谷方面": ["市川方面", "船橋方面", "柏方面", "松戸近郊"],
  "船橋方面": ["市川方面", "鎌ヶ谷方面", "千葉方面"],
  "葛飾方面": ["松戸近郊", "足立方面", "江戸川方面", "市川方面", "三郷方面"],
  "足立方面": ["葛飾方面", "松戸近郊", "八潮方面"],
  "江戸川方面": ["葛飾方面", "市川方面", "船橋方面"],
  "野田方面": ["流山方面", "吉川方面", "柏方面", "柏の葉方面"],
  "八潮方面": ["葛飾方面", "足立方面", "三郷方面", "草加方面"],
  "草加方面": ["八潮方面", "足立方面", "越谷方面"],
  "越谷方面": ["草加方面", "吉川方面"],
  "千葉方面": ["船橋方面", "市川方面"]
};

function getRouteFlowCompatibilityBetweenAreas(areaA, areaB) {
  const a = getCanonicalArea(areaA);
  const b = getCanonicalArea(areaB);
  if (!a || !b) return 0;
  if (a === b) return 100;

  const linkedA = ROUTE_FLOW_MAP[a] || [];
  const linkedB = ROUTE_FLOW_MAP[b] || [];
  const affinity = getAreaAffinityScore(a, b);
  const direction = getDirectionAffinityScore(a, b);

  let score = 0;
  if (linkedA.includes(b) || linkedB.includes(a)) score = 82;
  else if (affinity >= 72) score = 62;
  else if (affinity >= 54 && direction >= 28) score = 44;
  else if (direction >= 72) score = 36;
  else if (direction <= -38) score = -60;

  if ((a === "松戸近郊" && ["流山方面", "吉川方面", "三郷方面", "柏方面", "我孫子方面"].includes(b)) ||
      (b === "松戸近郊" && ["流山方面", "吉川方面", "三郷方面", "柏方面", "我孫子方面"].includes(a))) {
    score = Math.max(score, 88);
  }

  return score;
}

function getRouteFlowVehicleScore(targetArea, existingAreas = [], homeArea = "") {
  const areas = Array.isArray(existingAreas) ? existingAreas.filter(Boolean) : [];
  let best = 0;

  for (const area of areas) {
    best = Math.max(best, getRouteFlowCompatibilityBetweenAreas(targetArea, area));
  }

  if (!areas.length && homeArea) {
    best = Math.max(best, getRouteFlowCompatibilityBetweenAreas(targetArea, homeArea) * 0.35);
  }

  return best;
}

const ROUTE_CONTINUITY_HINTS = [
  { dir: "N", patterns: ["新松戸", "馬橋", "北松戸", "北小金", "小金", "幸谷", "新八柱", "八柱", "常盤平", "みのり台"] },
  { dir: "NE", patterns: ["南流山", "流山", "柏", "柏の葉", "我孫子", "取手", "藤代", "守谷"] },
  { dir: "W", patterns: ["葛飾", "足立", "綾瀬", "亀有", "金町", "八潮"] },
  { dir: "SW", patterns: ["江戸川"] },
  { dir: "S", patterns: ["市川", "本八幡", "妙典", "行徳"] },
  { dir: "SE", patterns: ["船橋", "習志野", "鎌ヶ谷", "鎌ケ谷"] },
  { dir: "NW", patterns: ["三郷", "吉川", "越谷", "草加"] }
];

function getAreaTravelDirection(area) {
  const raw = normalizeAreaLabel(area);
  if (!raw || raw === "無し") return "";

  for (const hint of ROUTE_CONTINUITY_HINTS) {
    if (hint.patterns.some(pattern => raw.includes(pattern))) return hint.dir;
  }

  return getAreaDirectionCluster(raw);
}

function getDirectionDistanceByKey(dirA, dirB) {
  if (!dirA || !dirB) return 99;
  if (dirA === dirB) return 0;
  if (dirA === "CENTER" || dirB === "CENTER") return 1;
  const indexA = DIRECTION_RING.indexOf(dirA);
  const indexB = DIRECTION_RING.indexOf(dirB);
  if (indexA < 0 || indexB < 0) return 99;
  const raw = Math.abs(indexA - indexB);
  return Math.min(raw, DIRECTION_RING.length - raw);
}

function isGatewayNearArea(area) {
  const raw = normalizeAreaLabel(area);
  if (!raw || raw === "無し") return false;
  return ["松戸", "新松戸", "馬橋", "北松戸", "北小金", "小金", "八柱", "常盤平", "みのり台"].some(keyword => raw.includes(keyword));
}

function getPairRouteContinuityPenalty(areaA, areaB) {
  const dirA = getAreaTravelDirection(areaA);
  const dirB = getAreaTravelDirection(areaB);
  const distance = getDirectionDistanceByKey(dirA, dirB);
  const routeFlow = getRouteFlowCompatibilityBetweenAreas(areaA, areaB);
  const affinity = getAreaAffinityScore(areaA, areaB);

  let penalty = 0;

  if (distance === 0) penalty += 0;
  else if (distance === 1) penalty += 10;
  else if (distance === 2) penalty += 48;
  else if (distance === 3) penalty += 130;
  else if (distance >= 4 && distance < 99) penalty += 240;

  if (routeFlow <= 0) penalty += 26;
  else if (routeFlow < 40) penalty += 14;

  if (affinity >= 72) penalty -= 26;
  else if (affinity >= 54) penalty -= 10;

  const gatewayA = isGatewayNearArea(areaA);
  const gatewayB = isGatewayNearArea(areaB);
  if (gatewayA || gatewayB) {
    if (distance >= 2) penalty += 82;
  }

  const pair = [getCanonicalArea(areaA), getCanonicalArea(areaB)].filter(Boolean).sort().join("__");
  const strongPairs = new Set([
    "吉川方面__松戸近郊",
    "三郷方面__松戸近郊",
    "我孫子方面__松戸近郊",
    "柏方面__松戸近郊",
    "流山方面__松戸近郊"
  ]);
  if (strongPairs.has(pair)) penalty -= 48;

  if ((gatewayA || gatewayB) && ((dirA === "N" && ["W", "SW", "S"].includes(dirB)) || (dirB === "N" && ["W", "SW", "S"].includes(dirA)))) {
    penalty += 120;
  }

  return Math.max(0, penalty);
}

function getRouteContinuityPenalty(targetArea, existingAreas = [], homeArea = "") {
  const areas = Array.isArray(existingAreas) ? existingAreas.filter(Boolean) : [];
  if (!areas.length) return 0;

  let penalty = 0;
  let bestPairPenalty = Infinity;

  for (const area of areas) {
    const pairPenalty = getPairRouteContinuityPenalty(targetArea, area);
    bestPairPenalty = Math.min(bestPairPenalty, pairPenalty);
    penalty += pairPenalty;
  }

  penalty = bestPairPenalty === Infinity ? penalty : penalty * 0.45 + bestPairPenalty * 0.9;

  if (homeArea) {
    const homeDir = getAreaTravelDirection(homeArea);
    const targetDir = getAreaTravelDirection(targetArea);
    const homeDistance = getDirectionDistanceByKey(targetDir, homeDir);
    if (homeDistance >= 3) penalty += 42;
  }

  return Math.max(0, penalty);
}

function getRouteFlowSortWeight(area) {
  const canonical = getCanonicalArea(area);
  if (["守谷方面", "藤代方面", "取手方面", "我孫子方面", "千葉方面"].includes(canonical)) return 100;
  if (["吉川方面", "船橋方面", "野田方面", "柏の葉方面"].includes(canonical)) return 85;
  if (["流山方面", "柏方面", "市川方面", "鎌ヶ谷方面", "三郷方面", "足立方面", "葛飾方面"].includes(canonical)) return 70;
  if (canonical === "松戸近郊") return 40;
  return 55;
}

function sortClustersForRouteFlow(clusters) {
  return [...clusters].sort((a, b) => {
    if (a.hour !== b.hour) return a.hour - b.hour;
    const aw = getRouteFlowSortWeight(a.area);
    const bw = getRouteFlowSortWeight(b.area);
    if (bw !== aw) return bw - aw;
    if (b.count !== a.count) return b.count - a.count;
    return b.totalDistance - a.totalDistance;
  });
}

function getAssignmentAreasByVehicleHour(assignments, itemsById, vehicleId, hour, excludeItemId = null) {
  return assignments
    .filter(a => Number(a.vehicle_id) === Number(vehicleId) && Number(a.actual_hour) === Number(hour) && Number(a.item_id) !== Number(excludeItemId || -1))
    .map(a => normalizeAreaLabel(itemsById.get(Number(a.item_id))?.destination_area || ""))
    .filter(Boolean);
}

function optimizeAssignmentsByRouteFlow(assignments, items, vehicles) {
  const itemById = new Map(items.map(item => [Number(item.id), item]));
  const vehicleMap = new Map(vehicles.map(v => [Number(v.id), v]));
  const working = assignments.map(a => ({ ...a }));

  for (const assignment of working) {
    const item = itemById.get(Number(assignment.item_id));
    if (!item) continue;

    const area = normalizeAreaLabel(item.destination_area || item.cluster_area || "無し");
    const currentVehicle = vehicleMap.get(Number(assignment.vehicle_id));
    const currentHome = normalizeAreaLabel(currentVehicle?.home_area || "");
    const currentAreas = getAssignmentAreasByVehicleHour(working, itemById, assignment.vehicle_id, assignment.actual_hour, assignment.item_id);
    const currentRouteScore = getRouteFlowVehicleScore(area, currentAreas, currentHome);
    const currentHomeScore = getLastTripHomePriorityWeight(area, currentHome, true, true);

    let best = null;

    for (const vehicle of vehicles) {
      if (Number(vehicle.id) === Number(assignment.vehicle_id)) continue;

      const hourAssignments = working.filter(a => Number(a.vehicle_id) === Number(vehicle.id) && Number(a.actual_hour) === Number(assignment.actual_hour));
      const seatCapacity = Number(vehicle.seat_capacity || 4);
      if (hourAssignments.length >= seatCapacity) continue;

      const homeArea = normalizeAreaLabel(vehicle.home_area || "");
      if (isHardReverseForHome(area, homeArea)) continue;

      const existingAreas = hourAssignments.map(a => normalizeAreaLabel(itemById.get(Number(a.item_id))?.destination_area || ""));
      const routeScore = getRouteFlowVehicleScore(area, existingAreas, homeArea);
      const continuityPenalty = getRouteContinuityPenalty(area, existingAreas, homeArea);
      const currentContinuityPenalty = getRouteContinuityPenalty(area, currentAreas, currentHome);
      const homeScore = getLastTripHomePriorityWeight(area, homeArea, true, true);
      const direction = Math.max(0, getDirectionAffinityScore(area, homeArea));
      const strict = getStrictHomeCompatibilityScore(area, homeArea);

      const totalScore = routeScore * 2.4 + homeScore * 0.9 + direction * 0.5 + strict * 0.6 - continuityPenalty * 1.7 - hourAssignments.length * 12;
      const currentTotal = currentRouteScore * 2.4 + currentHomeScore * 0.9 + Math.max(0, getDirectionAffinityScore(area, currentHome)) * 0.5 + getStrictHomeCompatibilityScore(area, currentHome) * 0.6 - currentContinuityPenalty * 1.7 - currentAreas.length * 12;

      if (totalScore >= currentTotal + 90) {
        if (!best || totalScore > best.totalScore) best = { vehicle, totalScore };
      }
    }

    if (best) {
      assignment.vehicle_id = best.vehicle.id;
      assignment.driver_name = best.vehicle.driver_name || "";
    }
  }

  return working;
}

function getCanonicalArea(area) {
  const normalized = normalizeAreaLabel(area);
  if (!normalized || normalized === "無し") return "";

  for (const [canonical, patterns] of AREA_CANONICAL_PATTERNS) {
    if (patterns.some(pattern => normalized.includes(pattern))) {
      return canonical;
    }
  }

  if (normalized.endsWith("方面")) return normalized;
  return normalized;
}

function getAreaAffinityScore(areaA, areaB) {
  const a = getCanonicalArea(areaA);
  const b = getCanonicalArea(areaB);
  if (!a || !b) return 0;
  if (a === b) return 100;
  return Number(AREA_AFFINITY_MAP[a]?.[b] || AREA_AFFINITY_MAP[b]?.[a] || 0);
}

function getAreaDirectionCluster(area) {
  const canonical = getCanonicalArea(area);
  return AREA_DIRECTION_MAP[canonical] || "";
}

function getDirectionDistance(areaA, areaB) {
  const dirA = getAreaDirectionCluster(areaA);
  const dirB = getAreaDirectionCluster(areaB);
  if (!dirA || !dirB) return 99;
  if (dirA === "CENTER" || dirB === "CENTER") return 1;
  const indexA = DIRECTION_RING.indexOf(dirA);
  const indexB = DIRECTION_RING.indexOf(dirB);
  if (indexA < 0 || indexB < 0) return 99;
  const raw = Math.abs(indexA - indexB);
  return Math.min(raw, DIRECTION_RING.length - raw);
}

function getDirectionAffinityScore(areaA, areaB) {
  const distance = getDirectionDistance(areaA, areaB);
  if (distance === 99) return 0;
  if (distance === 0) return 100;
  if (distance === 1) return 72;
  if (distance === 2) return 28;
  if (distance === 3) return -38;
  return -95;
}

function getStrictHomeCompatibilityScore(clusterArea, homeArea) {
  const cluster = getCanonicalArea(clusterArea);
  const home = getCanonicalArea(homeArea);
  if (!cluster || !home) return 0;
  if (cluster === home) return 100;
  const allowed = HOME_ALLOWED_AREA_MAP[home] || [];
  if (allowed.includes(cluster)) return 78;
  const directionScore = getDirectionAffinityScore(cluster, home);
  if (directionScore >= 72) return 52;
  if (directionScore >= 28) return 18;
  return 0;
}

function isHardReverseForHome(clusterArea, homeArea) {
  const affinity = getAreaAffinityScore(clusterArea, homeArea);
  const directionScore = getDirectionAffinityScore(clusterArea, homeArea);
  const strictScore = getStrictHomeCompatibilityScore(clusterArea, homeArea);
  if (directionScore <= -95) return true;
  if (directionScore <= -38 && strictScore === 0) return true;
  if (affinity <= 25 && strictScore === 0) return true;
  return false;
}

function getLastTripHomePriorityWeight(clusterArea, homeArea, isLastRun, isDefaultLastHourCluster) {
  const affinity = getAreaAffinityScore(clusterArea, homeArea);
  const directionScore = getDirectionAffinityScore(clusterArea, homeArea);
  const strictScore = getStrictHomeCompatibilityScore(clusterArea, homeArea);

  let weight = affinity * 1.1 + Math.max(directionScore, 0) * 1.15 + strictScore * 1.35;

  if (directionScore < 0) {
    weight += directionScore * 2.4;
  }

  if (isHardReverseForHome(clusterArea, homeArea)) {
    weight -= isLastRun ? 520 : (isDefaultLastHourCluster ? 320 : 90);
  }

  if (isLastRun) return weight * 2.8;
  if (isDefaultLastHourCluster) return weight * 2.2;
  return weight * 0.45;
}

function openGoogleMap(address, lat = null, lng = null) {
  const origin = encodeURIComponent(ORIGIN_LABEL);
  const latNum = toNullableNumber(lat);
  const lngNum = toNullableNumber(lng);
  let dest = "";
  if (isValidLatLng(latNum, lngNum)) dest = encodeURIComponent(`${latNum},${lngNum}`);
  else dest = encodeURIComponent(String(address || "").trim());
  if (!dest) return;
  window.open(
    `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${dest}&travelmode=driving`,
    "_blank"
  );
}




function buildMapUrlFromAddressOrLatLng(address, lat, lng) {
  const numLat = toNullableNumber(lat);
  const numLng = toNullableNumber(lng);

  if (isValidLatLng(numLat, numLng)) {
    return `https://www.google.com/maps/search/?api=1&query=${numLat},${numLng}`;
  }

  const safeAddress = String(address || "").trim();
  if (safeAddress) {
    return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(safeAddress)}`;
  }

  return "";
}

function buildCastMapUrl(cast) {
  return buildMapUrlFromAddressOrLatLng(
    cast?.address,
    cast?.latitude,
    cast?.longitude
  );
}

function buildDispatchItemMapUrl(item) {
  return buildMapUrlFromAddressOrLatLng(
    item?.destination_address || item?.casts?.address || "",
    item?.casts?.latitude,
    item?.casts?.longitude
  );
}

function buildMapLinkHtml({ name, address, lat, lng, className = "map-name-link" }) {
  const safeName = escapeHtml(name || "-");
  const mapUrl = buildMapUrlFromAddressOrLatLng(address, lat, lng);

  if (!mapUrl) return safeName;

  return `<a href="${mapUrl}" target="_blank" rel="noopener noreferrer" class="${className}">${safeName} 📍</a>`;
}

function downloadTextFile(filename, text, mimeType = "text/plain;charset=utf-8") {
  const blob = new Blob([text], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function csvEscape(value) {
  const s = String(value ?? "");
  if (s.includes('"') || s.includes(",") || s.includes("\n")) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

function parseCsvLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const next = line[i + 1];

    if (char === '"' && inQuotes && next === '"') {
      current += '"';
      i++;
      continue;
    }
    if (char === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (char === "," && !inQuotes) {
      result.push(current.trim());
      current = "";
      continue;
    }
    current += char;
  }

  result.push(current.trim());
  return result;
}

function normalizeCsvHeader(header) {
  const raw = String(header || "").trim();
  const h = raw.toLowerCase();

  const map = {
    name: "name",
    名前: "name",
    cast_name: "name",

    phone: "phone",
    tel: "phone",
    telephone: "phone",
    電話: "phone",
    電話番号: "phone",

    address: "address",
    住所: "address",

    area: "area",
    方面: "area",
    地域: "area",

    latitude: "latitude",
    lat: "latitude",
    緯度: "latitude",

    longitude: "longitude",
    lng: "longitude",
    lon: "longitude",
    経度: "longitude",

    memo: "memo",
    メモ: "memo",
    note: "memo",

    distance_km: "distance_km",
    距離: "distance_km",
    想定距離: "distance_km",

    plate_number: "plate_number",
    vehicle_id: "plate_number",
    車両id: "plate_number",
    車両ID: "plate_number",
    車両: "plate_number",

    vehicle_area: "vehicle_area",
    担当方面: "vehicle_area",

    home_area: "home_area",
    帰宅方面: "home_area",

    seat_capacity: "seat_capacity",
    定員: "seat_capacity",
    乗車可能人員: "seat_capacity",

    driver_name: "driver_name",
    driver: "driver_name",
    ドライバー: "driver_name",
    ドライバー名: "driver_name",

    line_id: "line_id",
    line: "line_id",
    lineid: "line_id",
    "line id": "line_id",
    LINEID: "line_id",
    "LINE ID": "line_id",
    line_id_: "line_id",

    status: "status",
    状態: "status"
  };

  return map[raw] || map[h] || h;
}

function parseCsv(text) {
  const lines = text
    .replace(/^\uFEFF/, "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .filter(line => line.trim() !== "");

  if (!lines.length) return [];

  const headers = parseCsvLine(lines[0]).map(h => normalizeCsvHeader(h.trim()));
  const rows = [];

  for (let i = 1; i < lines.length; i++) {
    const values = parseCsvLine(lines[i]);
    const row = {};
    headers.forEach((header, idx) => {
      row[header] = values[idx] ?? "";
    });
    rows.push(row);
  }

  return rows;
}

async function readCsvFileAsText(file) {
  const buffer = await file.arrayBuffer();
  let text = new TextDecoder("utf-8").decode(buffer);

  const mojibakeLike =
    text.includes("�") ||
    (!text.includes("name") &&
      !text.includes("address") &&
      !text.includes("名前") &&
      !text.includes("住所"));

  if (mojibakeLike) {
    try {
      text = new TextDecoder("shift_jis").decode(buffer);
    } catch (e) {
      console.warn("shift_jis decode failed:", e);
    }
  }

  return text;
}

function normalizeCsvRows(rows) {
  return rows.map(row => {
    const normalized = {};
    Object.keys(row).forEach(key => {
      const nk = normalizeCsvHeader(key);
      normalized[nk] = row[key];
    });
    return normalized;
  });
}

function getVehicleMonthlyStatsMap(reportRows, targetMonth) {
  const map = new Map();

  reportRows.forEach(row => {
    if (getMonthKey(row.report_date) !== targetMonth) return;
    const vehicleId = Number(row.vehicle_id);
    const prev = map.get(vehicleId) || {
      totalDistance: 0,
      workedDays: 0,
      avgDistance: 0
    };
    prev.totalDistance += Number(row.distance_km || 0);
    prev.workedDays += 1;
    prev.avgDistance = prev.workedDays > 0 ? prev.totalDistance / prev.workedDays : 0;
    map.set(vehicleId, prev);
  });

  return map;
}

function getDoneCastIdsInActuals() {
  const ids = new Set();
  currentActualsCache.forEach(item => {
    if (Number(item.cast_id) && normalizeStatus(item.status) === "done") {
      ids.add(Number(item.cast_id));
    }
  });
  return ids;
}

function getPlannedCastIds() {
  const ids = new Set();
  currentPlansCache.forEach(plan => {
    if (!plan.cast_id) return;
    if (["planned", "assigned", "done", "cancel"].includes(plan.status)) {
      ids.add(Number(plan.cast_id));
    }
  });
  return ids;
}

function getRemainingPlannedCastIds(dateStr) {
  const ids = new Set();

  currentPlansCache.forEach(plan => {
    if (plan.plan_date !== dateStr) return;
    if (!plan.cast_id) return;
    const status = String(plan.status || "");
    if (status === "done" || status === "cancel") return;
    ids.add(Number(plan.cast_id));
  });

  currentActualsCache.forEach(item => {
    const status = normalizeStatus(item.status);
    if (!item.cast_id) return;
    if (status === "done") {
      ids.delete(Number(item.cast_id));
    }
  });

  return ids;
}

function isLastClusterOfTheDay(cluster, dateStr) {
  const remainingIds = getRemainingPlannedCastIds(dateStr);
  cluster.items.forEach(item => {
    remainingIds.delete(Number(item.cast_id));
  });
  return remainingIds.size === 0;
}

function getVehicleAreaMatchScore(vehicle, area) {
  const normalizedArea = normalizeAreaLabel(area);
  const vehicleArea = normalizeAreaLabel(vehicle?.vehicle_area || "");
  const homeArea = normalizeAreaLabel(vehicle?.home_area || "");
  let score = 0;

  const vehicleAffinity = getAreaAffinityScore(vehicleArea, normalizedArea);
  const homeAffinity = getAreaAffinityScore(homeArea, normalizedArea);
  const vehicleDirection = Math.max(0, getDirectionAffinityScore(vehicleArea, normalizedArea));
  const homeDirection = Math.max(0, getDirectionAffinityScore(homeArea, normalizedArea));
  const strictHome = getStrictHomeCompatibilityScore(normalizedArea, homeArea);

  score += vehicleAffinity * 0.40;
  score += homeAffinity * 0.24;
  score += vehicleDirection * 0.20;
  score += homeDirection * 0.28;
  score += strictHome * 0.34;

  return score;
}

function buildDispatchClusters(items) {
  const activeItems = [...items]
    .filter(item => !["done", "cancel"].includes(normalizeStatus(item.status)))
    .map(item => ({
      ...item,
      cluster_hour: Number(item.actual_hour ?? 0),
      cluster_area: normalizeAreaLabel(item.destination_area || "無し"),
      cluster_distance: Number(item.distance_km || 0)
    }));

  const clusterMap = new Map();

  activeItems.forEach(item => {
    const key = `${item.cluster_hour}__${item.cluster_area}`;
    if (!clusterMap.has(key)) {
      clusterMap.set(key, {
        key,
        hour: item.cluster_hour,
        area: item.cluster_area,
        items: [],
        totalDistance: 0,
        count: 0
      });
    }
    const cluster = clusterMap.get(key);
    cluster.items.push(item);
    cluster.totalDistance += item.cluster_distance;
    cluster.count += 1;
  });

  return sortClustersForRouteFlow([...clusterMap.values()]);
}

async function ensureAuth() {
  const { data, error } = await supabaseClient.auth.getUser();

  if (error) {
    alert("ユーザー情報の取得に失敗しました");
    window.location.href = "index.html";
    return false;
  }

  currentUser = data.user;

  if (!currentUser) {
    window.location.href = "index.html";
    return false;
  }

  const { error: profileError } = await supabaseClient
    .from("profiles")
    .upsert({
      id: currentUser.id,
      email: currentUser.email,
      display_name: currentUser.email,
      role: "dispatcher"
    });

  if (profileError) {
    console.error(profileError);
    alert("profiles作成エラー: " + profileError.message);
    return false;
  }

  if (els.userEmail) els.userEmail.value = currentUser.email || "";
  return true;
}

async function logout() {
  await supabaseClient.auth.signOut();
  window.location.href = "index.html";
}

function openManual() {
  window.open(
    "https://drive.google.com/file/d/1LRTe2qcaef3dtItKcTAijadgHKpdmPAM/view?usp=drive_link",
    "_blank"
  );
}

function activateTab(tabId) {
  document.querySelectorAll(".main-tab").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.tab === tabId);
  });

  document.querySelectorAll(".page-panel").forEach(panel => {
    panel.classList.toggle("active", panel.id === tabId);
  });
}

function setupTabs() {
  document.querySelectorAll(".main-tab").forEach(btn => {
    btn.addEventListener("click", () => activateTab(btn.dataset.tab));
  });

  document.querySelectorAll(".go-tab-btn").forEach(btn => {
    btn.addEventListener("click", () => activateTab(btn.dataset.goTab));
  });
}

function renderHomeSummary() {
  const actualDone = currentActualsCache.filter(x => normalizeStatus(x.status) === "done").length;
  const actualCancel = currentActualsCache.filter(x => normalizeStatus(x.status) === "cancel").length;

  if (els.homeCastCount) els.homeCastCount.textContent = String(allCastsCache.length);
  if (els.homeVehicleCount) els.homeVehicleCount.textContent = String(allVehiclesCache.length);
  if (els.homePlanCount) els.homePlanCount.textContent = String(currentPlansCache.length);
  if (els.homeActualCount) els.homeActualCount.textContent = String(currentActualsCache.length);
  if (els.homeDoneCount) els.homeDoneCount.textContent = String(actualDone);
  if (els.homeCancelCount) els.homeCancelCount.textContent = String(actualCancel);
}

function renderHomeMonthlyVehicleList() {
  if (!els.homeMonthlyVehicleList) return;

  const monthKey = getMonthKey(els.dispatchDate?.value || todayStr());
  const statsMap = getVehicleMonthlyStatsMap(currentDailyReportsCache, monthKey);

  els.homeMonthlyVehicleList.innerHTML = "";

  if (!allVehiclesCache.length) {
    els.homeMonthlyVehicleList.innerHTML = `<div class="chip">車両なし</div>`;
    return;
  }

  allVehiclesCache.forEach(vehicle => {
    const stats = statsMap.get(Number(vehicle.id)) || {
      totalDistance: 0,
      workedDays: 0,
      avgDistance: 0
    };

    const row = document.createElement("div");
    row.className = "home-monthly-item";
    row.innerHTML = `
      ${escapeHtml(vehicle.driver_name || vehicle.plate_number || "-")}</span>
      <span class="chip">${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))}</span>
      <span class="chip">帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))}</span>
      <span class="chip">月間:${stats.totalDistance.toFixed(1)}km</span>
      <span class="chip">出勤:${stats.workedDays}日</span>
      <span class="chip">平均:${stats.avgDistance.toFixed(1)}km</span>
    `;
    els.homeMonthlyVehicleList.appendChild(row);
  });
}
function resetCastForm() {
  editingCastId = null;
  if (els.castName) els.castName.value = "";
  if (els.castDistanceKm) els.castDistanceKm.value = "";
  if (els.castAddress) els.castAddress.value = "";
  if (els.castArea) els.castArea.value = "";
  if (els.castMemo) els.castMemo.value = "";
  if (els.castLatLngText) els.castLatLngText.value = "";
  if (els.castPhone) els.castPhone.value = "";
  if (els.castLat) els.castLat.value = "";
  if (els.castLng) els.castLng.value = "";
  lastCastGeocodeKey = "";
  setCastGeoStatus("idle", "住所入力後 Enter で座標取得 / 未取得時は座標貼り付けから手動反映できます");
  if (els.cancelEditBtn) els.cancelEditBtn.classList.add("hidden");
}

function fillCastForm(cast) {
  editingCastId = cast.id;
  if (els.castName) els.castName.value = cast.name || "";
  if (els.castDistanceKm) els.castDistanceKm.value = cast.distance_km ?? "";
  if (els.castAddress) els.castAddress.value = cast.address || "";
  if (els.castArea) els.castArea.value = normalizeAreaLabel(cast.area || "");
  if (els.castMemo) els.castMemo.value = cast.memo || "";
  if (els.castPhone) els.castPhone.value = cast.phone || "";
  if (els.castLat) els.castLat.value = cast.latitude ?? "";
  if (els.castLng) els.castLng.value = cast.longitude ?? "";
  if (els.castLatLngText) {
    els.castLatLngText.value =
      cast.latitude != null && cast.longitude != null
        ? `${cast.latitude},${cast.longitude}`
        : "";
  }
  lastCastGeocodeKey = normalizeGeocodeAddressKey(cast.address || "");
  if (cast.latitude != null && cast.longitude != null) {
    setCastGeoStatus("success", "✔ 座標取得済");
  } else {
    setCastGeoStatus("idle", "住所入力後 Enter で座標取得 / 未取得時は座標貼り付けから手動反映できます");
  }
  if (els.cancelEditBtn) els.cancelEditBtn.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function isDuplicateCast(name, address) {
  const normalizedName = String(name || "").trim();
  const normalizedAddress = String(address || "").trim();

  return allCastsCache.find(
    c =>
      String(c.name || "").trim() === normalizedName &&
      String(c.address || "").trim() === normalizedAddress &&
      Number(c.id) !== Number(editingCastId)
  );
}

async function saveCast() {
  const name = els.castName?.value.trim();
  const address = els.castAddress?.value.trim();

  if (!name) {
    alert("氏名を入力してください");
    return;
  }

  const duplicate = isDuplicateCast(name, address);
  if (duplicate) {
    alert("このキャストは既に登録されています");
    return;
  }

  let lat = toNullableNumber(els.castLat?.value);
  let lng = toNullableNumber(els.castLng?.value);

  const addressKey = normalizeGeocodeAddressKey(address);
  if (address && (!isValidLatLng(lat, lng) || addressKey !== lastCastGeocodeKey)) {
    const geocoded = await fillCastLatLngFromAddress({ silent: true, force: addressKey !== lastCastGeocodeKey });
    lat = geocoded?.lat ?? toNullableNumber(els.castLat?.value);
    lng = geocoded?.lng ?? toNullableNumber(els.castLng?.value);
  }

  const manualArea = els.castArea?.value.trim() || "";
  const autoArea = guessArea(lat, lng, address);
  const autoDistance = await resolveDistanceKmFromOrigin(address, lat, lng);

  const payload = {
    name,
    phone: els.castPhone?.value.trim() || "",
    address,
    area: normalizeAreaLabel(manualArea || autoArea || ""),
    distance_km: toNullableNumber(els.castDistanceKm?.value) ?? autoDistance,
    latitude: lat,
    longitude: lng,
    memo: els.castMemo?.value.trim() || "",
    is_active: true
  };

  let error;
  if (editingCastId) {
    ({ error } = await supabaseClient.from("casts").update(payload).eq("id", editingCastId));
  } else {
    payload.created_by = currentUser.id;
    ({ error } = await supabaseClient.from("casts").insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(
    null,
    null,
    editingCastId ? "update_cast" : "create_cast",
    editingCastId ? "キャストを更新" : "キャストを作成"
  );

  resetCastForm();
  await loadCasts();
}

async function deleteCast(castId) {
  if (!window.confirm("このキャストを削除しますか？")) return;

  const { error } = await supabaseClient
    .from("casts")
    .update({ is_active: false })
    .eq("id", castId);

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "delete_cast", `キャストID ${castId} を削除`);
  await loadCasts();
}

async function loadCasts() {
  const { data, error } = await supabaseClient
    .from("casts")
    .select("*")
    .eq("is_active", true);

  if (error) {
    console.error(error);
    return;
  }

  allCastsCache = [...(data || [])].sort((a, b) =>
    String(a.name || "").localeCompare(String(b.name || ""), "ja")
  );

  renderCastsTable();
  renderCastSelects();
  renderCastSearchResults();
  renderHomeSummary();
}

function renderCastsTable() {
  if (!els.castsTableBody) return;

  els.castsTableBody.innerHTML = "";

  if (!allCastsCache.length) {
    els.castsTableBody.innerHTML = `<tr><td colspan="6" class="muted">キャストがありません</td></tr>`;
    return;
  }

  allCastsCache.forEach(cast => {
  const tr = document.createElement("tr");
  tr.innerHTML = `
  <td>
    ${
      buildCastMapUrl(cast)
        ? `<a href="${buildCastMapUrl(cast)}" target="_blank" rel="noopener noreferrer" class="cast-name-link">${escapeHtml(cast.name || "")} 📍</a>`
        : `${escapeHtml(cast.name || "")}`
    }
  </td>
  <td>${escapeHtml(cast.address || "")}</td>
  <td>${escapeHtml(normalizeAreaLabel(cast.area || ""))}</td>
  <td>${cast.distance_km ?? ""}</td>
  <td>${escapeHtml(cast.memo || "")}</td>
  <td class="actions-cell">
    <button class="btn ghost cast-edit-btn" data-id="${cast.id}">編集</button>
    <button class="btn ghost cast-route-btn" data-address="${escapeHtml(cast.address || "")}">ルート</button>
    <button class="btn danger cast-delete-btn" data-id="${cast.id}">削除</button>
  </td>
`;
  els.castsTableBody.appendChild(tr);
});

  els.castsTableBody.querySelectorAll(".cast-edit-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const cast = allCastsCache.find(x => Number(x.id) === Number(btn.dataset.id));
      if (cast) fillCastForm(cast);
    });
  });

  els.castsTableBody.querySelectorAll(".cast-route-btn").forEach(btn => {
    btn.addEventListener("click", () => openGoogleMap(btn.dataset.address || ""));
  });

  els.castsTableBody.querySelectorAll(".cast-delete-btn").forEach(btn => {
    btn.addEventListener("click", async () => deleteCast(Number(btn.dataset.id)));
  });
}

function exportCastsCsv() {
  const headers = [
    "name",
    "phone",
    "address",
    "area",
    "distance_km",
    "latitude",
    "longitude",
    "memo"
  ];

  const rows = allCastsCache.map(cast => [
    cast.name || "",
    cast.phone || "",
    cast.address || "",
    normalizeAreaLabel(cast.area || ""),
    cast.distance_km ?? "",
    cast.latitude ?? "",
    cast.longitude ?? "",
    cast.memo || ""
  ]);

  const csv = [headers.join(","), ...rows.map(r => r.map(csvEscape).join(","))].join("\n");
  downloadTextFile(`casts_${todayStr()}.csv`, csv, "text/csv;charset=utf-8");
}

async function importCastCsvFile() {
  const file = els.csvFileInput?.files?.[0];
  if (!file) {
    alert("CSVファイルを選択してください");
    return;
  }

  try {
    const text = await readCsvFileAsText(file);
    let rows = parseCsv(text);
    rows = normalizeCsvRows(rows);

    if (!rows.length) {
      alert("CSVデータが空です");
      return;
    }

    const uniqueMap = new Map();

    for (const row of rows) {
      const name = String(row.name || "").trim();
      const address = String(row.address || "").trim();
      if (!name || !address) continue;

      const key = `${name}__${address}`;
      if (!uniqueMap.has(key)) {
        uniqueMap.set(key, row);
      }
    }

    const mergedRows = [...uniqueMap.values()];
    const payloads = [];

    for (const row of mergedRows) {
      const name = String(row.name || "").trim();
      const address = String(row.address || "").trim();
      if (!name || !address) continue;

      const lat = toNullableNumber(row.latitude);
      const lng = toNullableNumber(row.longitude);
      const autoArea = guessArea(lat, lng, address);

      payloads.push({
        name,
        phone: String(row.phone || "").trim(),
        address,
        area: normalizeAreaLabel(String(row.area || "").trim() || autoArea || ""),
        distance_km:
          toNullableNumber(row.distance_km) ??
          (isValidLatLng(lat, lng) ? estimateRoadKmFromStation(lat, lng) : null),
        latitude: lat,
        longitude: lng,
        memo: String(row.memo || "").trim(),
        is_active: true,
        created_by: currentUser.id
      });
    }

    if (!payloads.length) {
      alert("取り込めるデータがありません");
      els.csvFileInput.value = "";
      return;
    }

    const { error } = await supabaseClient
      .from("casts")
      .upsert(payloads, { onConflict: "name,address" });

    if (error) {
      console.error("CSV import supabase error:", error);
      alert("CSV取込エラー: " + error.message);
      return;
    }

    els.csvFileInput.value = "";
    await addHistory(null, null, "import_csv", `${payloads.length}件のキャストをCSV取込/更新`);
    alert(`${payloads.length}件のキャストをCSV取込/更新しました`);
    await loadCasts();
  } catch (error) {
    console.error("importCastCsvFile error:", error);
    alert("CSV取込中にエラーが発生しました");
  }
}

function applyCastLatLng() {
  const parsed = parseLatLngText(els.castLatLngText?.value || "");
  if (!parsed) {
    alert("座標形式が正しくありません");
    return;
  }

  if (els.castLat) els.castLat.value = parsed.lat;
  if (els.castLng) els.castLng.value = parsed.lng;

  if (els.castArea) {
    els.castArea.value = normalizeAreaLabel(
      guessArea(parsed.lat, parsed.lng, els.castAddress?.value || "")
    );
  }

  if (els.castDistanceKm) {
    els.castDistanceKm.value = String(estimateRoadKmFromStation(parsed.lat, parsed.lng));
  }
  lastCastGeocodeKey = normalizeGeocodeAddressKey(els.castAddress?.value || "");
  setCastGeoStatus("manual", "✔ 座標反映済（手動入力）");
}

function guessCastArea() {
  const lat = toNullableNumber(els.castLat?.value);
  const lng = toNullableNumber(els.castLng?.value);
  if (els.castArea) {
    els.castArea.value = normalizeAreaLabel(guessArea(lat, lng, els.castAddress?.value || ""));
  }
}


function getFilteredCastsForSearch() {
  const nameQ = String(els.castSearchName?.value || "").trim().toLowerCase();
  const areaQ = String(els.castSearchArea?.value || "").trim().toLowerCase();
  const addressQ = String(els.castSearchAddress?.value || "").trim().toLowerCase();
  const phoneQ = String(els.castSearchPhone?.value || "").trim().toLowerCase();

  return allCastsCache.filter(cast => {
    const name = String(cast.name || "").toLowerCase();
    const area = String(normalizeAreaLabel(cast.area || "")).toLowerCase();
    const address = String(cast.address || "").toLowerCase();
    const phone = String(cast.phone || "").toLowerCase();

    if (nameQ && !name.includes(nameQ)) return false;
    if (areaQ && !area.includes(areaQ)) return false;
    if (addressQ && !address.includes(addressQ)) return false;
    if (phoneQ && !phone.includes(phoneQ)) return false;

    return true;
  });
}

function renderCastSearchResults() {
  if (!els.castSearchResultWrap) return;

  const rows = getFilteredCastsForSearch();
  if (els.castSearchCount) els.castSearchCount.textContent = String(rows.length);

  if (!rows.length) {
    els.castSearchResultWrap.innerHTML =
      `<div class="muted" style="padding:14px;">該当するキャストがありません</div>`;
    return;
  }

  let html = `
    <table class="data-table">
      <thead>
        <tr>
          <th>氏名</th>
          <th>住所</th>
          <th>方面</th>
          <th>想定距離(km)</th>
          <th>電話</th>
          <th>メモ</th>
          <th>操作</th>
        </tr>
      </thead>
      <tbody>
  `;

  rows.forEach(cast => {
    html += `
      <tr>
        <td>
          ${
            buildCastMapUrl(cast)
              ? `<a href="${buildCastMapUrl(cast)}" target="_blank" rel="noopener noreferrer" class="cast-name-link">${escapeHtml(cast.name || "")} 📍</a>`
              : `${escapeHtml(cast.name || "")}`
          }
        </td>
        <td>${escapeHtml(cast.address || "")}</td>
        <td>${escapeHtml(normalizeAreaLabel(cast.area || ""))}</td>
        <td>${cast.distance_km ?? ""}</td>
        <td>${escapeHtml(cast.phone || "")}</td>
        <td>${escapeHtml(cast.memo || "")}</td>
        <td class="actions-cell">
          <button class="btn ghost cast-search-map-btn" data-id="${cast.id}">地図</button>
          <button class="btn ghost cast-search-route-btn" data-address="${escapeHtml(cast.address || "")}">ルート</button>
          <button class="btn ghost cast-search-edit-btn" data-id="${cast.id}">編集へ</button>
        </td>
      </tr>
    `;
  });

  html += `</tbody></table>`;
  els.castSearchResultWrap.innerHTML = html;

  els.castSearchResultWrap.querySelectorAll(".cast-search-map-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const cast = allCastsCache.find(x => Number(x.id) === Number(btn.dataset.id));
      const url = buildCastMapUrl(cast);
      if (url) window.open(url, "_blank");
    });
  });

  els.castSearchResultWrap.querySelectorAll(".cast-search-route-btn").forEach(btn => {
    btn.addEventListener("click", () => openGoogleMap(btn.dataset.address || ""));
  });

  els.castSearchResultWrap.querySelectorAll(".cast-search-edit-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const cast = allCastsCache.find(x => Number(x.id) === Number(btn.dataset.id));
      if (!cast) return;
      activateTab("castsTab");
      fillCastForm(cast);
    });
  });
}

function resetCastSearchFilters() {
  if (els.castSearchName) els.castSearchName.value = "";
  if (els.castSearchArea) els.castSearchArea.value = "";
  if (els.castSearchAddress) els.castSearchAddress.value = "";
  if (els.castSearchPhone) els.castSearchPhone.value = "";
  renderCastSearchResults();
}

function resetVehicleForm() {
  editingVehicleId = null;
  if (els.vehiclePlateNumber) els.vehiclePlateNumber.value = "";
  if (els.vehicleArea) els.vehicleArea.value = "";
  if (els.vehicleHomeArea) els.vehicleHomeArea.value = "";
  if (els.vehicleSeatCapacity) els.vehicleSeatCapacity.value = "";
  if (els.vehicleDriverName) els.vehicleDriverName.value = "";
  if (els.vehicleLineId) els.vehicleLineId.value = "";
  if (els.vehicleStatus) els.vehicleStatus.value = "waiting";
  if (els.vehicleMemo) els.vehicleMemo.value = "";
  if (els.cancelVehicleEditBtn) els.cancelVehicleEditBtn.classList.add("hidden");
}

function fillVehicleForm(vehicle) {
  editingVehicleId = vehicle.id;
  if (els.vehiclePlateNumber) els.vehiclePlateNumber.value = vehicle.plate_number || "";
  if (els.vehicleArea) els.vehicleArea.value = normalizeAreaLabel(vehicle.vehicle_area || "");
  if (els.vehicleHomeArea) els.vehicleHomeArea.value = normalizeAreaLabel(vehicle.home_area || "");
  if (els.vehicleSeatCapacity) els.vehicleSeatCapacity.value = vehicle.seat_capacity ?? "";
  if (els.vehicleDriverName) els.vehicleDriverName.value = vehicle.driver_name || "";
  if (els.vehicleLineId) els.vehicleLineId.value = vehicle.line_id || "";
  if (els.vehicleStatus) els.vehicleStatus.value = vehicle.status || "waiting";
  if (els.vehicleMemo) els.vehicleMemo.value = vehicle.memo || "";
  if (els.cancelVehicleEditBtn) els.cancelVehicleEditBtn.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function isDuplicateVehicle(plateNumber) {
  const normalizedPlate = String(plateNumber || "").trim();
  return allVehiclesCache.find(
    v =>
      String(v.plate_number || "").trim() === normalizedPlate &&
      Number(v.id) !== Number(editingVehicleId)
  );
}

async function saveVehicle() {
  const plateNumber = els.vehiclePlateNumber?.value.trim();
  if (!plateNumber) {
    alert("車両IDを入力してください");
    return;
  }

  const duplicate = isDuplicateVehicle(plateNumber);
  if (duplicate) {
    alert("この車両IDは既に登録されています");
    return;
  }

  const payload = {
    plate_number: plateNumber,
    vehicle_area: normalizeAreaLabel(els.vehicleArea?.value.trim() || ""),
    home_area: normalizeAreaLabel(els.vehicleHomeArea?.value.trim() || ""),
    seat_capacity: Number(els.vehicleSeatCapacity?.value || 4),
    driver_name: els.vehicleDriverName?.value.trim() || "",
    line_id: els.vehicleLineId?.value.trim() || "",
    status: els.vehicleStatus?.value || "waiting",
    memo: els.vehicleMemo?.value.trim() || "",
    is_active: true
  };

  let error;
  if (editingVehicleId) {
    ({ error } = await supabaseClient.from("vehicles").update(payload).eq("id", editingVehicleId));
  } else {
    payload.created_by = currentUser.id;
    ({ error } = await supabaseClient.from("vehicles").insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(
    null,
    null,
    editingVehicleId ? "update_vehicle" : "create_vehicle",
    editingVehicleId ? "車両を更新" : "車両を登録"
  );

  resetVehicleForm();
  await loadVehicles();
}

async function deleteVehicle(vehicleId) {
  if (!window.confirm("この車両を削除しますか？")) return;

  const { error } = await supabaseClient
    .from("vehicles")
    .update({ is_active: false })
    .eq("id", vehicleId);

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "delete_vehicle", `車両ID ${vehicleId} を削除`);
  await loadVehicles();
}

async function loadVehicles() {
  const { data, error } = await supabaseClient
    .from("vehicles")
    .select("*")
    .eq("is_active", true)
    .order("id", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  allVehiclesCache = data || [];
  renderVehiclesTable();
  renderDailyVehicleChecklist();
  renderHomeSummary();
}

function renderVehiclesTable() {
  if (!els.vehiclesTableBody) return;

  const monthKey = getMonthKey(els.dispatchDate?.value || todayStr());
  const statsMap = getVehicleMonthlyStatsMap(currentDailyReportsCache, monthKey);

  els.vehiclesTableBody.innerHTML = "";

  if (!allVehiclesCache.length) {
    els.vehiclesTableBody.innerHTML = `<tr><td colspan="9" class="muted">車両がありません</td></tr>`;
    return;
  }

  allVehiclesCache.forEach(vehicle => {
    const stats = statsMap.get(Number(vehicle.id)) || {
      totalDistance: 0,
      workedDays: 0,
      avgDistance: 0
    };

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <tr>
      <td>${escapeHtml(vehicle.driver_name || "-")}</td>
      <td>${escapeHtml(vehicle.plate_number || "-")}</td>
      <td>${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))}</td>
      <td>${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))}</td>
      <td>${vehicle.seat_capacity ?? "-"}</td>
      <td>${stats.totalDistance.toFixed(1)}</td>
      <td>${stats.workedDays}</td>
      <td>${stats.avgDistance.toFixed(1)}</td>
      <td class="actions-cell">
        <button class="btn ghost vehicle-edit-btn" data-id="${vehicle.id}">編集</button>
        <button class="btn danger vehicle-delete-btn" data-id="${vehicle.id}">削除</button>
      </td>
    `;
    els.vehiclesTableBody.appendChild(tr);
  });

  els.vehiclesTableBody.querySelectorAll(".vehicle-edit-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const vehicle = allVehiclesCache.find(x => Number(x.id) === Number(btn.dataset.id));
      if (vehicle) fillVehicleForm(vehicle);
    });
  });

  els.vehiclesTableBody.querySelectorAll(".vehicle-delete-btn").forEach(btn => {
    btn.addEventListener("click", async () => deleteVehicle(Number(btn.dataset.id)));
  });
}

async function fetchDriverMileageRows(startDate, endDate) {
  const { data, error } = await supabaseClient
    .from("vehicle_daily_reports")
    .select(`
      *,
      vehicles (
        id,
        plate_number,
        driver_name
      )
    `)
    .gte("report_date", startDate)
    .lte("report_date", endDate)
    .order("report_date", { ascending: true })
    .order("vehicle_id", { ascending: true });

  if (error) {
    console.error(error);
    alert("走行実績の取得に失敗しました: " + error.message);
    return [];
  }

  return data || [];
}

function normalizeMileageExportRows(rows) {
  return rows.map(row => ({
    report_date: row.report_date || "",
    driver_name: row.driver_name || row.vehicles?.driver_name || "-",
    plate_number: row.vehicles?.plate_number || "-",
    distance_km: Number(row.distance_km || 0),
    worked_flag: Number(row.distance_km || 0) > 0 ? 1 : 0,
    note: row.note || ""
  }));
}

function renderMileageReportTable(rows) {
  if (!els.mileageReportTableWrap) return;

  if (!rows.length) {
    els.mileageReportTableWrap.innerHTML =
      `<div class="muted" style="padding:14px;">対象期間の走行実績はありません</div>`;
    return;
  }

  const grouped = new Map();

  rows.forEach(row => {
    const driver = row.driver_name || "-";
    if (!grouped.has(driver)) grouped.set(driver, []);
    grouped.get(driver).push(row);
  });

  let html = `<div class="grouped-plan-list">`;

  [...grouped.entries()].forEach(([driver, driverRows]) => {
    const totalDistance = driverRows.reduce((sum, row) => sum + Number(row.distance_km || 0), 0);
    html += `
      <div class="grouped-section">
        <div class="grouped-hour-title">
          ${escapeHtml(driver)} / ${driverRows.length}日 / 合計 ${totalDistance.toFixed(1)}km / 1日平均 ${(driverRows.length ? totalDistance / driverRows.length : 0).toFixed(1)}km
        </div>
    `;

    driverRows.forEach(row => {
      html += `
        <div class="grouped-row">
          <div>${escapeHtml(row.report_date || "")}</div>
          <div><strong>${escapeHtml(row.driver_name || "-")}</strong></div>
          <div>${Number(row.distance_km || 0).toFixed(1)}km</div>
        </div>
      `;
    });

    html += `</div>`;
  });

  html += `</div>`;
  els.mileageReportTableWrap.innerHTML = html;
}

async function previewDriverMileageReport() {
  const startDate = els.mileageReportStartDate?.value;
  const endDate = els.mileageReportEndDate?.value;

  if (!startDate || !endDate) {
    alert("開始日と終了日を選択してください");
    return;
  }

  if (startDate > endDate) {
    alert("開始日は終了日以前にしてください");
    return;
  }

  const rawRows = await fetchDriverMileageRows(startDate, endDate);
  currentMileageExportRows = normalizeMileageExportRows(rawRows);
  renderMileageReportTable(currentMileageExportRows);
}

function buildMileageSummaryRows(rows) {
  const map = new Map();

  rows.forEach(row => {
    const key = row.driver_name || "-";
    const prev = map.get(key) || {
      driver_name: key,
      days: 0,
      total_distance_km: 0,
      avg_distance_km: 0
    };
    prev.days += 1;
    prev.total_distance_km += Number(row.distance_km || 0);
    prev.avg_distance_km = prev.days ? prev.total_distance_km / prev.days : 0;
    map.set(key, prev);
  });

  return [...map.values()].map(row => ({
    driver_name: row.driver_name,
    days: row.days,
    total_distance_km: Number(row.total_distance_km.toFixed(1)),
    avg_distance_km: Number(row.avg_distance_km.toFixed(1))
  }));
}


function buildMileageDateRange(startDate, endDate) {
  const dates = [];
  const current = new Date(`${startDate}T00:00:00`);
  const end = new Date(`${endDate}T00:00:00`);

  while (current <= end) {
    dates.push(new Date(current));
    current.setDate(current.getDate() + 1);
  }

  return dates;
}

function formatMileageSheetDate(date) {
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${m}/${d}`;
}

function buildMileageMatrixRows(rows, startDate, endDate) {
  const dates = buildMileageDateRange(startDate, endDate);
  const grouped = new Map();

  rows.forEach(row => {
    const driver = row.driver_name || "-";
    if (!grouped.has(driver)) {
      grouped.set(driver, {
        driver_name: driver,
        byDate: new Map(),
        total_distance_km: 0,
        days: 0
      });
    }

    const item = grouped.get(driver);
    const key = row.report_date || "";
    const distance = Number(row.distance_km || 0);

    item.byDate.set(key, distance);
    item.total_distance_km += distance;
    if (Number(row.worked_flag || 0)) item.days += 1;
  });

  const sortedDrivers = [...grouped.values()].sort((a, b) =>
    String(a.driver_name || "").localeCompare(String(b.driver_name || ""), "ja")
  );

  const aoa = [];
  aoa.push([`${startDate.replaceAll("-", "/")}〜${endDate.replaceAll("-", "/")}`]);

  const header = ["No", "名前", ...dates.map(formatMileageSheetDate), "月間走行距離", "出勤日数", "1日平均走行距離"];
  aoa.push(header);

  let grandTotalDistance = 0;
  let grandWorkedDays = 0;
  const dailyTotals = new Map();

  sortedDrivers.forEach((row, index) => {
    const daily = dates.map(date => {
      const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
      const value = row.byDate.get(key);

      if (value != null && value !== 0) {
        dailyTotals.set(key, Number((dailyTotals.get(key) || 0) + Number(value || 0)));
        return Number(Number(value).toFixed(1));
      }
      return "";
    });

    const avg = row.days ? row.total_distance_km / row.days : 0;

    grandTotalDistance += row.total_distance_km;
    grandWorkedDays += row.days;

    aoa.push([
      index + 1,
      row.driver_name,
      ...daily,
      Number(row.total_distance_km.toFixed(1)),
      row.days,
      Number(avg.toFixed(1))
    ]);
  });

  const grandAvg = grandWorkedDays ? grandTotalDistance / grandWorkedDays : 0;
  const overallDaily = dates.map(date => {
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
    const total = Number(dailyTotals.get(key) || 0);
    return total ? Number(total.toFixed(1)) : "";
  });

  aoa.push([]);
  aoa.push([
    "",
    "全体",
    ...overallDaily,
    Number(grandTotalDistance.toFixed(1)),
    grandWorkedDays,
    Number(grandAvg.toFixed(1))
  ]);

  return aoa;
}

function applyMileageMatrixSheetStyle(sheet, startDate, endDate) {
  if (!sheet["!ref"] || !window.XLSX) return;

  const range = window.XLSX.utils.decode_range(sheet["!ref"]);
  sheet["!merges"] = sheet["!merges"] || [];
  sheet["!merges"].push({
    s: { r: 0, c: 0 },
    e: { r: 0, c: range.e.c }
  });

  sheet["!cols"] = [];
  for (let c = 0; c <= range.e.c; c++) {
    let wch = 8;
    if (c === 0) wch = 6;
    if (c === 1) wch = 12;
    if (c >= range.e.c - 2) wch = 14;
    sheet["!cols"].push({ wch });
  }
}

async function exportDriverMileageReportXlsx() {
  if (!window.XLSX) {
    alert("Excel出力ライブラリの読み込みに失敗しました");
    return;
  }

  if (!currentMileageExportRows.length) {
    await previewDriverMileageReport();
    if (!currentMileageExportRows.length) return;
  }

  const startDate = els.mileageReportStartDate?.value || todayStr();
  const endDate = els.mileageReportEndDate?.value || todayStr();

  const matrixAoa = buildMileageMatrixRows(currentMileageExportRows, startDate, endDate);
  const matrixSheet = window.XLSX.utils.aoa_to_sheet(matrixAoa);
  applyMileageMatrixSheetStyle(matrixSheet, startDate, endDate);

  const dailyRowsJa = currentMileageExportRows.map(row => ({
    日付: row.report_date || "",
    ドライバー名: row.driver_name || "-",
    走行距離_km: Number(row.distance_km || 0),
    出勤: Number(row.worked_flag || 0) ? "出勤" : "",
    メモ: row.note || ""
  }));

  const summaryRowsJa = buildMileageSummaryRows(currentMileageExportRows).map((row, index) => ({
    No: index + 1,
    名前: row.driver_name,
    月間走行距離: Number(row.total_distance_km || 0),
    出勤日数: Number(row.days || 0),
    "1日平均走行距離": Number(row.avg_distance_km || 0)
  }));

  const workbook = window.XLSX.utils.book_new();
  const dailySheet = window.XLSX.utils.json_to_sheet(dailyRowsJa);
  const summarySheet = window.XLSX.utils.json_to_sheet(summaryRowsJa);

  window.XLSX.utils.book_append_sheet(workbook, matrixSheet, "原本形式");
  window.XLSX.utils.book_append_sheet(workbook, dailySheet, "日別一覧");
  window.XLSX.utils.book_append_sheet(workbook, summarySheet, "ドライバー別集計");

  window.XLSX.writeFile(workbook, `driver_mileage_${startDate}_${endDate}.xlsx`);
}

function exportVehiclesCsv() {
  const headers = [
    "plate_number",
    "vehicle_area",
    "home_area",
    "seat_capacity",
    "driver_name",
    "line_id",
    "status",
    "memo"
  ];

  const rows = allVehiclesCache.map(vehicle => [
    vehicle.plate_number || "",
    normalizeAreaLabel(vehicle.vehicle_area || ""),
    normalizeAreaLabel(vehicle.home_area || ""),
    vehicle.seat_capacity ?? "",
    vehicle.driver_name || "",
    vehicle.line_id || "",
    vehicle.status || "waiting",
    vehicle.memo || ""
  ]);

  const csv = [headers.join(","), ...rows.map(r => r.map(csvEscape).join(","))].join("\n");
  downloadTextFile(`vehicles_${todayStr()}.csv`, csv, "text/csv;charset=utf-8");
}


function normalizePlanImportMode(input) {
  const value = String(input || "").trim();
  if (["1", "add", "append", "追加"].includes(value)) return "append";
  if (["2", "replace", "置換", "上書き"].includes(value)) return "replace";
  if (["3", "skip", "重複スキップ", "skip-duplicates"].includes(value)) return "skip";
  return "";
}

function getPlanDuplicateKey(row) {
  return [
    row.plan_date || "",
    Number(row.plan_hour || 0),
    Number(row.cast_id || 0),
    String(row.destination_address || "").trim(),
    normalizeAreaLabel(row.planned_area || ""),
    String(row.note || "").trim()
  ].join("|");
}

function exportPlansCsv() {
  const planDate = els.planDate?.value || todayStr();
  const rows = [...currentPlansCache]
    .sort((a, b) => Number(a.plan_hour || 0) - Number(b.plan_hour || 0) || Number(a.id || 0) - Number(b.id || 0))
    .map(plan => ({
      plan_date: planDate,
      plan_hour: Number(plan.plan_hour || 0),
      cast_id: Number(plan.cast_id || plan.casts?.id || 0),
      cast_name: plan.casts?.name || "",
      destination_address: plan.destination_address || plan.casts?.address || "",
      planned_area: normalizeAreaLabel(plan.planned_area || plan.casts?.area || ""),
      distance_km: plan.distance_km ?? "",
      note: plan.note || "",
      status: plan.status || "planned",
      vehicle_group: plan.vehicle_group || ""
    }));

  const headers = [
    "plan_date",
    "plan_hour",
    "cast_id",
    "cast_name",
    "destination_address",
    "planned_area",
    "distance_km",
    "note",
    "status",
    "vehicle_group"
  ];

  const csv = [headers.join(","), ...rows.map(row => headers.map(key => csvEscape(row[key] ?? "")).join(","))].join("\n");
  downloadTextFile(`plans_${planDate}.csv`, csv, "text/csv;charset=utf-8");
}

async function triggerImportPlansCsv() {
  els.plansCsvFileInput?.click();
}

async function importPlansCsvFile() {
  const file = els.plansCsvFileInput?.files?.[0];
  if (!file) return;

  const selectedDate = els.planDate?.value || todayStr();
  const modeInput = window.prompt(
    "予定CSVの取込方法を選んでください\n1: 追加\n2: 同日データを置換\n3: 重複をスキップ",
    "3"
  );
  const mode = normalizePlanImportMode(modeInput);
  if (!mode) {
    alert("取込を中止しました");
    els.plansCsvFileInput.value = "";
    return;
  }

  try {
    const text = await readCsvFileAsText(file);
    let rows = parseCsv(text);
    rows = normalizeCsvRows(rows);

    if (!rows.length) {
      alert("CSVにデータがありません");
      els.plansCsvFileInput.value = "";
      return;
    }

    const { data: existingRows, error: existingError } = await supabaseClient
      .from("dispatch_plans")
      .select("id, plan_date, plan_hour, cast_id, destination_address, planned_area, note")
      .eq("plan_date", selectedDate)
      .order("plan_hour", { ascending: true });

    if (existingError) {
      alert(existingError.message);
      els.plansCsvFileInput.value = "";
      return;
    }

    const existingList = existingRows || [];
    const existingKeys = new Set(existingList.map(getPlanDuplicateKey));

    if (mode === "replace" && existingList.length) {
      const { error: deleteError } = await supabaseClient
        .from("dispatch_plans")
        .delete()
        .eq("plan_date", selectedDate);
      if (deleteError) {
        alert(deleteError.message);
        els.plansCsvFileInput.value = "";
        return;
      }
      existingKeys.clear();
    }

    const inserts = [];
    const skipped = [];
    const missingCasts = [];

    for (const row of rows) {
      const castIdValue = Number(row.cast_id || 0);
      let cast = null;
      if (castIdValue) {
        cast = allCastsCache.find(x => Number(x.id) === castIdValue) || null;
      }
      if (!cast && row.cast_name) {
        const name = String(row.cast_name).trim();
        cast = allCastsCache.find(x => String(x.name || "").trim() === name) || null;
      }
      if (!cast) {
        missingCasts.push(String(row.cast_name || row.cast_id || "不明"));
        continue;
      }

      const payload = {
        plan_date: selectedDate,
        plan_hour: Number(row.plan_hour || 0),
        cast_id: Number(cast.id),
        destination_address: String(row.destination_address || cast.address || "").trim(),
        planned_area: normalizeAreaLabel(String(row.planned_area || cast.area || "無し")),
        distance_km: toNullableNumber(row.distance_km),
        note: String(row.note || "").trim(),
        status: String(row.status || "planned").trim() || "planned",
        vehicle_group: String(row.vehicle_group || "").trim(),
        created_by: currentUser?.id || null
      };

      const key = getPlanDuplicateKey(payload);
      if (mode === "skip" && existingKeys.has(key)) {
        skipped.push(`${getHourLabel(payload.plan_hour)} / ${cast.name}`);
        continue;
      }
      existingKeys.add(key);
      inserts.push(payload);
    }

    if (!inserts.length) {
      let msg = "取り込める予定がありませんでした。";
      if (missingCasts.length) msg += `\n未登録キャスト: ${[...new Set(missingCasts)].join(", ")}`;
      if (skipped.length) msg += `\n重複スキップ: ${skipped.length}件`;
      alert(msg);
      els.plansCsvFileInput.value = "";
      await loadPlansByDate(selectedDate);
      return;
    }

    const { error: insertError } = await supabaseClient
      .from("dispatch_plans")
      .insert(inserts);

    if (insertError) {
      alert(insertError.message);
      els.plansCsvFileInput.value = "";
      return;
    }

    let summary = `${inserts.length}件の予定をCSV取込しました`;
    if (mode === "replace") summary += "（同日置換）";
    if (skipped.length) summary += ` / 重複スキップ ${skipped.length}件`;
    if (missingCasts.length) summary += ` / 未登録キャスト ${[...new Set(missingCasts)].length}件`;

    await addHistory(null, null, "import_plans_csv", summary);
    alert(summary + `\n取込日付: ${selectedDate}`);
    els.plansCsvFileInput.value = "";
    await loadPlansByDate(selectedDate);
  } catch (error) {
    console.error("importPlansCsvFile error:", error);
    alert("予定CSV取込に失敗しました");
    els.plansCsvFileInput.value = "";
  }
}

async function importVehicleCsvFile() {
  const file = els.vehicleCsvFileInput?.files?.[0];
  if (!file) {
    alert("CSVファイルを選択してください");
    return;
  }

  try {
    const text = await readCsvFileAsText(file);
    let rows = parseCsv(text);
    rows = normalizeCsvRows(rows);

    if (!rows.length) {
      alert("CSVデータが空です");
      return;
    }

    const inserts = [];

    for (const row of rows) {
      const plateNumber = String(row.plate_number || "").trim();
      if (!plateNumber) continue;

      const exists = allVehiclesCache.find(
        v => String(v.plate_number || "").trim() === plateNumber
      );
      if (exists) {
        console.log("車両重複スキップ:", plateNumber);
        continue;
      }

      inserts.push({
        plate_number: plateNumber,
        vehicle_area: normalizeAreaLabel(String(row.vehicle_area || "").trim() || ""),
        home_area: normalizeAreaLabel(String(row.home_area || "").trim() || ""),
        seat_capacity: Number(row.seat_capacity || 4),
        driver_name: String(row.driver_name || "").trim(),
        line_id: String(row.line_id || "").trim(),
        status: String(row.status || "waiting").trim() || "waiting",
        memo: String(row.memo || "").trim(),
        is_active: true,
        created_by: currentUser.id
      });
    }

    if (!inserts.length) {
      alert("新規車両はありません");
      els.vehicleCsvFileInput.value = "";
      return;
    }

    const { error } = await supabaseClient.from("vehicles").insert(inserts);
    if (error) {
      console.error("Vehicle CSV import error:", error);
      alert("車両CSV取込エラー: " + error.message);
      return;
    }

    els.vehicleCsvFileInput.value = "";
    await addHistory(null, null, "import_vehicle_csv", `${inserts.length}件の車両をCSV取込`);
    alert(`${inserts.length}件の車両を取り込みました`);
    await loadVehicles();
  } catch (error) {
    console.error("importVehicleCsvFile error:", error);
    alert("車両CSV取込中にエラーが発生しました");
  }
}

function clearPlanCastDerivedFields() {
  if (els.planAddress) els.planAddress.value = "";
  if (els.planArea) els.planArea.value = "";
  if (els.planDistanceKm) els.planDistanceKm.value = "";
}

function clearActualCastDerivedFields() {
  if (els.actualAddress) els.actualAddress.value = "";
  if (els.actualArea) els.actualArea.value = "";
  if (els.actualDistanceKm) els.actualDistanceKm.value = "";
}

async function syncPlanFieldsFromCastInput(forceFill = false) {
  const cast = findCastByInputValue(els.planCastSelect?.value || "");
  if (!cast) {
    clearPlanCastDerivedFields();
    return null;
  }

  let distance = await resolveDistanceKmForCastRecord(cast);

  if (els.planAddress) els.planAddress.value = cast.address || "";
  if (els.planArea) {
    els.planArea.value = normalizeAreaLabel(
      cast.area ||
        guessArea(
          toNullableNumber(cast.latitude),
          toNullableNumber(cast.longitude),
          cast.address || ""
        )
    );
  }
  if (els.planDistanceKm) els.planDistanceKm.value = distance ?? "";

  return cast;
}

async function syncActualFieldsFromCastInput(forceFill = false) {
  const cast = findCastByInputValue(els.castSelect?.value || "");
  if (!cast) {
    clearActualCastDerivedFields();
    return null;
  }

  let distance = await resolveDistanceKmForCastRecord(cast);

  if (els.actualAddress) els.actualAddress.value = cast.address || "";
  if (els.actualArea) {
    els.actualArea.value = normalizeAreaLabel(
      cast.area ||
        guessArea(
          toNullableNumber(cast.latitude),
          toNullableNumber(cast.longitude),
          cast.address || ""
        )
    );
  }
  if (els.actualDistanceKm) els.actualDistanceKm.value = distance ?? "";

  return cast;
}

function resetPlanForm() {
  editingPlanId = null;
  if (els.planCastSelect) els.planCastSelect.value = "";
  if (els.planHour) els.planHour.value = "0";
  clearPlanCastDerivedFields();
  if (els.planNote) els.planNote.value = "";
}

function fillPlanForm(plan) {
  editingPlanId = plan.id;
  if (els.planCastSelect) els.planCastSelect.value = plan.casts?.name || "";
  if (els.planHour) els.planHour.value = String(plan.plan_hour ?? 0);
  if (els.planDistanceKm) els.planDistanceKm.value = plan.distance_km ?? "";
  if (els.planAddress) els.planAddress.value = plan.destination_address || plan.casts?.address || "";
  if (els.planArea) els.planArea.value = normalizeAreaLabel(plan.planned_area || "");
  if (els.planNote) els.planNote.value = plan.note || "";
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function fillPlanFormFromSelectedCast() {
  const cast = findCastByInputValue(els.planCastSelect?.value || "");
  if (!cast) return;

  if (els.planAddress && !els.planAddress.value.trim()) {
    els.planAddress.value = cast.address || "";
  }

  if (els.planArea && !els.planArea.value.trim()) {
    els.planArea.value = normalizeAreaLabel(
      cast.area ||
        guessArea(
          toNullableNumber(cast.latitude),
          toNullableNumber(cast.longitude),
          cast.address || ""
        )
    );
  }

  if (els.planDistanceKm && !els.planDistanceKm.value.trim()) {
    const distance = await resolveDistanceKmForCastRecord(cast);
    els.planDistanceKm.value = distance ?? "";
  }
}

async function loadPlansByDate(dateStr) {
  const { data, error } = await supabaseClient
    .from("dispatch_plans")
    .select(`
      *,
      casts (
        id,
        name,
        phone,
        address,
        area,
        distance_km,
        latitude,
        longitude
      )
    `)
    .eq("plan_date", dateStr)
    .order("plan_hour", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  currentPlansCache = data || [];
  renderPlanGroupedTable();
  renderPlansTimeAreaMatrix();
  renderPlanSelect();
  renderPlanCastSelect();
  renderHomeSummary();
}


function getPlanSelectableCasts() {
  const plannedIds = getPlannedCastIds();
  const editingPlan = editingPlanId
    ? currentPlansCache.find(x => Number(x.id) === Number(editingPlanId))
    : null;
  const editingCastIdForPlan = Number(editingPlan?.cast_id || 0);

  return allCastsCache.filter(
    cast => Number(cast.id) === editingCastIdForPlan || !plannedIds.has(Number(cast.id))
  );
}

function getActualSelectableCasts() {
  const usedCastIds = new Set();
  const doneCastIds = getDoneCastIdsInActuals();

  currentActualsCache.forEach(item => {
    if (item.cast_id && normalizeStatus(item.status) !== "cancel") {
      usedCastIds.add(Number(item.cast_id));
    }
  });

  const editingActual = editingActualId
    ? currentActualsCache.find(x => Number(x.id) === Number(editingActualId))
    : null;
  const editingCastIdForActual = Number(editingActual?.cast_id || 0);

  return allCastsCache.filter(
    cast =>
      Number(cast.id) === editingCastIdForActual ||
      (!usedCastIds.has(Number(cast.id)) && !doneCastIds.has(Number(cast.id)))
  );
}

function getCastSearchText(cast) {
  return [
    String(cast.name || "").trim(),
    normalizeAreaLabel(cast.area || "-"),
    String(cast.address || "").trim()
  ].join(" / ");
}

function filterCastCandidates(casts, query) {
  const q = String(query || "").trim().toLowerCase();

  const sorted = [...casts].sort((a, b) =>
    String(a.name || "").localeCompare(String(b.name || ""), "ja")
  );

  if (!q) return sorted;

  return sorted.filter(cast => {
    const hay = [
      cast.name || "",
      cast.address || "",
      cast.area || "",
      cast.phone || "",
      cast.memo || ""
    ]
      .join(" ")
      .toLowerCase();

    return hay.includes(q);
  });
}

function renderCastSearchSuggest(container, casts, onPick) {
  if (!container) return;

  if (!casts.length) {
    container.innerHTML = "";
    container.classList.add("hidden");
    return;
  }

  container.innerHTML = casts
    .map(
      cast => `
        <button type="button" class="cast-search-item" data-id="${cast.id}">
          <span>${escapeHtml(cast.name || "-")}</span>
          <small>${escapeHtml(normalizeAreaLabel(cast.area || "-"))} / ${escapeHtml(cast.address || "")}</small>
        </button>
      `
    )
    .join("");

  container.classList.remove("hidden");

  container.querySelectorAll(".cast-search-item").forEach(btn => {
    btn.addEventListener("mousedown", event => {
      event.preventDefault();
      const cast = allCastsCache.find(x => Number(x.id) === Number(btn.dataset.id));
      if (cast) onPick(cast);
      container.classList.add("hidden");
    });
  });
}

function setupSearchableCastInput(input, suggest, getCandidates, onPick) {
  if (!input || !suggest) return;
  if (input.dataset.searchBound === "1") return;
  input.dataset.searchBound = "1";

  const openSuggest = () => {
    const casts = filterCastCandidates(getCandidates(), input.value || "");
    renderCastSearchSuggest(suggest, casts, onPick);
  };

  input.addEventListener("focus", openSuggest);
  input.addEventListener("input", openSuggest);

  input.addEventListener("keydown", event => {
    if (event.key === "Escape") {
      suggest.classList.add("hidden");
      return;
    }

    if (event.key === "Enter") {
      const exact = findCastByInputValue(input.value || "");
      if (exact) {
        onPick(exact);
        suggest.classList.add("hidden");
        return;
      }

      const candidates = filterCastCandidates(getCandidates(), input.value || "");
      if (candidates.length === 1) {
        onPick(candidates[0]);
        suggest.classList.add("hidden");
      }
    }
  });

  input.addEventListener("blur", () => {
    window.setTimeout(() => {
      suggest.classList.add("hidden");
    }, 150);
  });
}

function setupSearchableCastInputs() {
  setupSearchableCastInput(
    els.planCastSelect,
    els.planCastSuggest,
    getPlanSelectableCasts,
    cast => {
      if (els.planCastSelect) els.planCastSelect.value = cast.name || "";
      fillPlanFormFromSelectedCast();
    }
  );

  setupSearchableCastInput(
    els.castSelect,
    els.castSuggest,
    getActualSelectableCasts,
    cast => {
      if (els.castSelect) els.castSelect.value = cast.name || "";
      fillActualFormFromSelectedCast();
    }
  );
}

function renderPlanCastSelect() {
  const input = els.planCastSelect;
  const list = document.getElementById("planCastList");
  if (!input || !list) return;

  const editingPlan = editingPlanId
    ? currentPlansCache.find(x => Number(x.id) === Number(editingPlanId))
    : null;

  list.innerHTML = "";

  getPlanSelectableCasts().forEach(cast => {
    const option = document.createElement("option");
    option.value = String(cast.name || "").trim();
    option.label = getCastSearchText(cast);
    list.appendChild(option);
  });

  if (editingPlan?.casts?.name) {
    input.value = editingPlan.casts.name;
  }
}

function isPlanAlreadyAddedToActual(plan, excludeActualId = null) {
  if (!plan) return false;

  const targetDate = String(plan.plan_date || els.actualDate?.value || todayStr()).trim();
  const targetCastId = Number(plan.cast_id || 0);
  const targetHour = Number(plan.plan_hour || 0);
  const targetAddress = String(plan.destination_address || plan.casts?.address || "").trim();

  return currentActualsCache.some(item => {
    if (excludeActualId !== null && Number(item.id) === Number(excludeActualId)) return false;

    const itemDate = String(item.plan_date || els.actualDate?.value || todayStr()).trim();
    const itemCastId = Number(item.cast_id || 0);
    const itemHour = Number(item.actual_hour || 0);
    const itemAddress = String(item.destination_address || item.casts?.address || "").trim();

    if (itemDate !== targetDate) return false;

    const sameCastHour = itemCastId === targetCastId && itemHour === targetHour;
    const sameAddressHour = !!targetAddress && itemAddress === targetAddress && itemHour === targetHour;
    const sameCastAddress = itemCastId === targetCastId && !!targetAddress && itemAddress === targetAddress;

    return sameCastHour || sameAddressHour || sameCastAddress;
  });
}

function getLinkedPlanForActual(actualItem) {
  if (!actualItem) return null;

  const actualDate = String(actualItem.plan_date || els.actualDate?.value || todayStr()).trim();
  const actualCastId = Number(actualItem.cast_id || 0);
  const actualHour = Number(actualItem.actual_hour || 0);
  const actualAddress = String(actualItem.destination_address || actualItem.casts?.address || "").trim();

  return currentPlansCache.find(plan => {
    const planDate = String(plan.plan_date || "").trim();
    const planCastId = Number(plan.cast_id || 0);
    const planHour = Number(plan.plan_hour || 0);
    const planAddress = String(plan.destination_address || plan.casts?.address || "").trim();

    if (planDate !== actualDate) return false;

    const sameCastHour = planCastId === actualCastId && planHour === actualHour;
    const sameAddressHour = !!actualAddress && planAddress === actualAddress && planHour === actualHour;
    const sameCastAddress = planCastId === actualCastId && !!actualAddress && planAddress === actualAddress;

    return sameCastHour || sameAddressHour || sameCastAddress;
  }) || null;
}

function renderPlanSelect() {
  if (!els.planSelect) return;

  const targetDate = els.actualDate?.value || todayStr();
  const doneCastIds = getDoneCastIdsInActuals();
  const editingActual = editingActualId
    ? currentActualsCache.find(x => Number(x.id) === Number(editingActualId))
    : null;
  const editingPlan = getLinkedPlanForActual(editingActual);
  const selectedValueBefore = String(els.planSelect.value || "");
  const appendedPlanIds = new Set();

  els.planSelect.innerHTML = `<option value="">予定から選択</option>`;

  const appendOption = plan => {
    if (!plan || appendedPlanIds.has(Number(plan.id))) return;
    appendedPlanIds.add(Number(plan.id));

    const option = document.createElement("option");
    option.value = plan.id;
    option.textContent = `${getHourLabel(plan.plan_hour)} / ${plan.casts?.name || "-"} / ${normalizeAreaLabel(plan.planned_area || "-")}`;
    if (editingPlan && Number(plan.id) === Number(editingPlan.id) && editingActualId) {
      option.textContent += " [編集中]";
    }
    els.planSelect.appendChild(option);
  };

  currentPlansCache
    .filter(plan => plan.plan_date === targetDate)
    .filter(plan => plan.status === "planned" || (editingPlan && Number(plan.id) === Number(editingPlan.id)))
    .filter(plan => !doneCastIds.has(Number(plan.cast_id)) || (editingPlan && Number(plan.id) === Number(editingPlan.id)))
    .filter(plan => !isPlanAlreadyAddedToActual(plan, editingActualId || null) || (editingPlan && Number(plan.id) === Number(editingPlan.id)))
    .forEach(appendOption);

  if (editingPlan) {
    appendOption(editingPlan);
  }

  if (editingPlan) {
    els.planSelect.value = String(editingPlan.id);
  } else if (selectedValueBefore && appendedPlanIds.has(Number(selectedValueBefore))) {
    els.planSelect.value = selectedValueBefore;
  } else {
    els.planSelect.value = "";
  }
}

async function savePlan() {
  const cast = findCastByInputValue(els.planCastSelect?.value || "");
  const castId = Number(cast?.id || 0);
  if (!castId) {
    alert("キャストを選択または入力してください");
    return;
  }

  const planDate = els.planDate?.value || todayStr();
  const hour = Number(els.planHour?.value || 0);
  const address = els.planAddress?.value.trim() || "";
  let distanceKm = toNullableNumber(els.planDistanceKm?.value);
  if (distanceKm === null) {
    distanceKm = await resolveDistanceKmForCastRecord(cast, address);
    if (distanceKm !== null && els.planDistanceKm) els.planDistanceKm.value = String(distanceKm);
  }
  const area = els.planArea?.value.trim() || "";
  const note = els.planNote?.value.trim() || "";

  const payload = {
    plan_date: planDate,
    plan_hour: hour,
    cast_id: castId,
    destination_address: address,
    planned_area: normalizeAreaLabel(area || "無し"),
    distance_km: distanceKm,
    note,
    status: "planned"
  };

  let error;
  if (editingPlanId) {
    ({ error } = await supabaseClient.from("dispatch_plans").update(payload).eq("id", editingPlanId));
  } else {
    payload.created_by = currentUser.id;
    ({ error } = await supabaseClient.from("dispatch_plans").insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(
    null,
    null,
    editingPlanId ? "update_plan" : "create_plan",
    editingPlanId ? "予定を更新" : "予定を作成"
  );

  resetPlanForm();
  await loadPlansByDate(planDate);
}

async function deletePlan(planId) {
  if (!window.confirm("この予定を削除しますか？")) return;

  const { error } = await supabaseClient.from("dispatch_plans").delete().eq("id", planId);
  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "delete_plan", `予定ID ${planId} を削除`);
  await loadPlansByDate(els.planDate?.value || todayStr());
}

function renderPlanGroupedTable() {
  if (!els.plansGroupedTable) return;

  if (!currentPlansCache.length) {
    els.plansGroupedTable.innerHTML = `<div class="muted" style="padding:14px;">予定がありません</div>`;
    return;
  }

  const hours = [...new Set(currentPlansCache.map(x => Number(x.plan_hour)))].sort((a, b) => a - b);
  let html = `<div class="grouped-plan-list">`;

  hours.forEach(hour => {
    const hourItems = currentPlansCache.filter(x => Number(x.plan_hour) === hour);
    const groupedAreas = getGroupedAreasByDisplay(hourItems, x => x.planned_area || "無し");

    html += `<div class="grouped-section">`;
    html += `<div class="grouped-hour-title">${getHourLabel(hour)}</div>`;

    groupedAreas.forEach(({ detailArea }) => {
      const areaItems = hourItems.filter(
        x => normalizeAreaLabel(x.planned_area || "無し") === detailArea
      );

      html += `<div class="grouped-area-title">${getGroupedAreaHeaderHtml(detailArea)}</div>`;

      areaItems.forEach(plan => {
        html += `
          <div class="grouped-row">
            <div>${getHourLabel(hour)}</div>
            <div><strong>${buildMapLinkHtml({
             name: plan.casts?.name,
             address: plan.destination_address || plan.casts?.address,
             lat: plan.casts?.latitude,
             lng: plan.casts?.longitude
             })}</strong></div>
             <div>${escapeHtml(normalizeAreaLabel(plan.planned_area || "無し"))}</div>
             <div>${plan.distance_km ?? ""}</div>
             <div class="op-cell">
              <span class="badge-status ${normalizeStatus(plan.status)}">${escapeHtml(getStatusText(plan.status))}</span>
              <button class="btn ghost plan-edit-btn" data-id="${plan.id}">編集</button>
              <button class="btn ghost plan-route-btn" data-address="${escapeHtml(plan.destination_address || plan.casts?.address || "")}">ルート</button>
              <button class="btn danger plan-delete-btn" data-id="${plan.id}">削除</button>
            </div>
          </div>
        `;
      });
    });

    html += `</div>`;
  });

  html += `</div>`;
  els.plansGroupedTable.innerHTML = html;

  els.plansGroupedTable.querySelectorAll(".plan-edit-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const plan = currentPlansCache.find(x => Number(x.id) === Number(btn.dataset.id));
      if (plan) fillPlanForm(plan);
    });
  });

  els.plansGroupedTable.querySelectorAll(".plan-route-btn").forEach(btn => {
    btn.addEventListener("click", () => openGoogleMap(btn.dataset.address || ""));
  });

  els.plansGroupedTable.querySelectorAll(".plan-delete-btn").forEach(btn => {
    btn.addEventListener("click", async () => deletePlan(Number(btn.dataset.id)));
  });
}

function renderPlansTimeAreaMatrix() {
  if (!els.plansTimeAreaMatrix) return;

  const hours = [0, 1, 2, 3, 4, 5];
  const areas = [
    ...new Set(
      currentPlansCache.map(x => getAreaDisplayGroup(normalizeAreaLabel(x.planned_area || "無し")))
    )
  ];

  if (!areas.length) {
    els.plansTimeAreaMatrix.innerHTML =
      `<div class="muted" style="padding:14px;">一覧がありません</div>`;
    return;
  }

  let html = `
    <table class="matrix-table">
      <thead>
        <tr>
          <th>時間</th>
          ${areas.map(area => `<th>${escapeHtml(area)}</th>`).join("")}
        </tr>
      </thead>
      <tbody>
  `;

  hours.forEach(hour => {
    html += `<tr><td>${getHourLabel(hour)}</td>`;

    areas.forEach(area => {
      const rows = currentPlansCache.filter(
        plan =>
          Number(plan.plan_hour ?? 0) === hour &&
          getAreaDisplayGroup(normalizeAreaLabel(plan.planned_area || "無し")) === area
      );

      if (!rows.length) {
        html += `<td>-</td>`;
      } else {
        const totalDistance = rows.reduce(
          (sum, row) => sum + Number(row.distance_km || 0),
          0
        );

        html += `
          <td>
            <div class="matrix-card">
              <div class="matrix-summary">${rows.length}人 / ${totalDistance.toFixed(1)}km</div>
              ${rows.map(row => `
                <div class="matrix-item">
                  <span class="badge-status ${normalizeStatus(row.status)}">${escapeHtml(getStatusText(row.status))}</span>
                  <span>${buildMapLinkHtml({
                   name: row.casts?.name,
                   address: row.destination_address || row.casts?.address,
                   lat: row.casts?.latitude,
                   lng: row.casts?.longitude
                  })} (${Number(row.distance_km || 0).toFixed(1)}km)</span>
                </div>
              `).join("")}
            </div>
          </td>
        `;
      }
    });

    html += `</tr>`;
  });

  html += `</tbody></table>`;
  els.plansTimeAreaMatrix.innerHTML = html;
}

function guessPlanArea() {
  if (els.planArea) {
    els.planArea.value = normalizeAreaLabel(
      classifyAreaByAddress(els.planAddress?.value || "") || "無し"
    );
  }
}

async function clearAllPlans() {
  if (!window.confirm("この日の予定を全消去しますか？")) return;

  const planDate = els.planDate?.value || todayStr();
  const { error } = await supabaseClient
    .from("dispatch_plans")
    .delete()
    .eq("plan_date", planDate);

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "clear_plans", `${planDate} の予定を全削除`);
  await loadPlansByDate(planDate);
}

function resetActualForm() {
  editingActualId = null;
  if (els.planSelect) els.planSelect.value = "";
  if (els.castSelect) els.castSelect.value = "";
  if (els.actualHour) els.actualHour.value = "0";
  if (els.actualStatus) els.actualStatus.value = "pending";
  clearActualCastDerivedFields();
  if (els.actualNote) els.actualNote.value = "";
}

function fillActualForm(item) {
  editingActualId = item.id;
  renderPlanSelect();
  if (els.castSelect) els.castSelect.value = item.casts?.name || "";
  if (els.actualHour) els.actualHour.value = String(item.actual_hour ?? 0);
  if (els.actualDistanceKm) els.actualDistanceKm.value = item.distance_km ?? "";
  if (els.actualStatus) els.actualStatus.value = item.status || "pending";
  if (els.actualAddress) els.actualAddress.value = item.destination_address || item.casts?.address || "";
  if (els.actualArea) els.actualArea.value = normalizeAreaLabel(item.destination_area || item.casts?.area || "");
  if (els.actualNote) els.actualNote.value = item.note || "";
  const linkedPlan = getLinkedPlanForActual(item);
  if (els.planSelect) els.planSelect.value = linkedPlan ? String(linkedPlan.id) : "";
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function fillActualFormFromSelectedCast() {
  const cast = findCastByInputValue(els.castSelect?.value || "");
  if (!cast) return;

  if (els.actualAddress && !els.actualAddress.value.trim()) {
    els.actualAddress.value = cast.address || "";
  }

  if (els.actualArea && !els.actualArea.value.trim()) {
    els.actualArea.value = normalizeAreaLabel(
      cast.area ||
        guessArea(
          toNullableNumber(cast.latitude),
          toNullableNumber(cast.longitude),
          cast.address || ""
        )
    );
  }

  if (els.actualDistanceKm && !els.actualDistanceKm.value.trim()) {
    let distance = toNullableNumber(cast.distance_km);
    if (distance === null) {
      const lat = toNullableNumber(cast.latitude);
      const lng = toNullableNumber(cast.longitude);
      if (isValidLatLng(lat, lng)) {
        distance = estimateRoadKmFromStation(lat, lng);
      }
    }
    els.actualDistanceKm.value = distance ?? "";
  }
}

function fillActualFormFromSelectedPlan() {
  const planId = Number(els.planSelect?.value || 0);
  if (!planId) return;

  const plan = currentPlansCache.find(p => Number(p.id) === Number(planId));
  if (!plan) return;

  if (els.castSelect) els.castSelect.value = plan.casts?.name || "";
  if (els.actualHour) els.actualHour.value = String(plan.plan_hour ?? 0);
  if (els.actualAddress) els.actualAddress.value = plan.destination_address || plan.casts?.address || "";
  if (els.actualArea) els.actualArea.value = normalizeAreaLabel(plan.planned_area || plan.casts?.area || "無し");
  if (els.actualDistanceKm) els.actualDistanceKm.value = plan.distance_km ?? plan.casts?.distance_km ?? "";
  if (els.actualNote) els.actualNote.value = plan.note || "";
}

function renderCastSelects() {
  const editingActual = editingActualId
    ? currentActualsCache.find(x => Number(x.id) === Number(editingActualId))
    : null;

  const input = els.castSelect;
  const list = document.getElementById("castList");

  if (input && list) {
    list.innerHTML = "";

    getActualSelectableCasts().forEach(cast => {
      const option = document.createElement("option");
      option.value = String(cast.name || "").trim();
      option.label = getCastSearchText(cast);
      list.appendChild(option);
    });

    if (editingActual?.casts?.name) {
      input.value = editingActual.casts.name;
    }
  }

  renderPlanCastSelect();
  setupSearchableCastInputs();
}

async function loadActualsByDate(dateStr) {
  const { data: dispatches, error: dispatchError } = await supabaseClient
    .from("dispatches")
    .select("*")
    .eq("dispatch_date", dateStr)
    .order("id", { ascending: false })
    .limit(1);

  if (dispatchError) {
    console.error(dispatchError);
    return;
  }

  if (dispatches?.length) {
    currentDispatchId = dispatches[0].id;
  } else {
    const { data: inserted, error: createError } = await supabaseClient
      .from("dispatches")
      .insert({
        dispatch_date: dateStr,
        status: "draft",
        created_by: currentUser.id
      })
      .select()
      .single();

    if (createError) {
      console.error(createError);
      return;
    }
    currentDispatchId = inserted.id;
  }

  const { data, error } = await supabaseClient
    .from("dispatch_items")
    .select(`
      *,
      casts (
        id,
        name,
        phone,
        address,
        area,
        distance_km,
        latitude,
        longitude
      )
    `)
    .eq("dispatch_id", currentDispatchId)
    .order("actual_hour", { ascending: true })
    .order("stop_order", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  currentActualsCache = data || [];
  renderActualTable();
  renderActualTimeAreaMatrix();
  renderHomeSummary();
  renderCastSelects();
  renderManualLastVehicleInfo();
}

async function saveActual() {
  const cast = findCastByInputValue(els.castSelect?.value || "");
  const castId = Number(cast?.id || 0);
  if (!castId) {
    alert("キャストを選択または入力してください");
    return;
  }

  const dateStr = els.actualDate?.value || todayStr();
  const hour = Number(els.actualHour?.value || 0);
  const address = els.actualAddress?.value.trim() || "";
  const area = normalizeAreaLabel(els.actualArea?.value.trim() || "無し");
  let distanceKm = toNullableNumber(els.actualDistanceKm?.value);
  if (distanceKm === null) {
    distanceKm = await resolveDistanceKmForCastRecord(cast, address);
    if (distanceKm !== null && els.actualDistanceKm) els.actualDistanceKm.value = String(distanceKm);
  }
  const status = els.actualStatus?.value || "pending";
  const note = els.actualNote?.value.trim() || "";

  const existingActual = editingActualId
    ? currentActualsCache.find(x => Number(x.id) === Number(editingActualId))
    : null;

  const stopOrder = existingActual
    ? Number(existingActual.stop_order || 1)
    : currentActualsCache.filter(
        x =>
          Number(x.actual_hour) === hour &&
          Number(x.id) !== Number(editingActualId || 0)
      ).length + 1;

  const payload = {
    dispatch_id: currentDispatchId,
    cast_id: castId,
    actual_hour: hour,
    stop_order: stopOrder,
    pickup_label: ORIGIN_LABEL,
    destination_address: address,
    destination_area: area,
    distance_km: distanceKm,
    status,
    note,
    plan_date: dateStr
  };

  let error;
  if (editingActualId) {
    ({ error } = await supabaseClient.from("dispatch_items").update(payload).eq("id", editingActualId));
  } else {
    ({ error } = await supabaseClient.from("dispatch_items").insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(
    currentDispatchId,
    editingActualId || null,
    editingActualId ? "update_actual" : "create_actual",
    editingActualId ? "実際の送りを更新" : "実際の送りを追加"
  );

  resetActualForm();
  await loadActualsByDate(dateStr);
  if (!editingActualId) {
    try {
      await assignUnassignedActualsForToday();
      await loadActualsByDate(dateStr);
    } catch (assignError) {
      console.error("assignUnassignedActualsForToday error:", assignError);
    }
  }
  await loadPlansByDate(els.planDate?.value || dateStr);
}

async function deleteActual(itemId) {
  if (!window.confirm("このActualを削除しますか？")) return;

  const { error } = await supabaseClient.from("dispatch_items").delete().eq("id", itemId);
  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(currentDispatchId, itemId, "delete_actual", `Actual ID ${itemId} を削除`);
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
}

async function updateActualStatus(itemId, status) {
  const item = currentActualsCache.find(x => Number(x.id) === Number(itemId));
  if (!item) {
    alert("対象のActualが見つかりません");
    return;
  }

  const { error } = await supabaseClient
    .from("dispatch_items")
    .update({ status })
    .eq("id", itemId);

  if (error) {
    alert(error.message);
    return;
  }

  const targetPlan = currentPlansCache.find(
    plan =>
      Number(plan.cast_id) === Number(item.cast_id) &&
      plan.plan_date === (els.actualDate?.value || todayStr()) &&
      Number(plan.plan_hour) === Number(item.actual_hour ?? -1)
  );

  if (targetPlan) {
    let nextPlanStatus = targetPlan.status;
    if (status === "done") nextPlanStatus = "done";
    else if (status === "cancel") nextPlanStatus = "planned";
    else if (status === "pending") nextPlanStatus = "assigned";

    const { error: planError } = await supabaseClient
      .from("dispatch_plans")
      .update({ status: nextPlanStatus })
      .eq("id", targetPlan.id);

    if (planError) console.error(planError);
  }

  await addHistory(currentDispatchId, itemId, "update_actual_status", `Actual状態を ${status} に変更`);
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
}

async function addPlanToActual() {
  const planId = Number(els.planSelect?.value || 0);
  if (!planId) {
    alert("予定を選択してください");
    return;
  }

  const plan = currentPlansCache.find(x => Number(x.id) === Number(planId));
  if (!plan) {
    alert("予定が見つかりません");
    return;
  }

  if (isPlanAlreadyAddedToActual(plan)) {
    alert("その予定はすでにActualへ追加されています");
    renderPlanSelect();
    return;
  }

  if (currentActualsCache.some(x => Number(x.cast_id) === Number(plan.cast_id) && normalizeStatus(x.status) !== "cancel")) {
    alert("そのキャストはすでにActualにあります");
    renderPlanSelect();
    return;
  }

  const doneCastIds = getDoneCastIdsInActuals();
  if (doneCastIds.has(Number(plan.cast_id))) {
    alert("このキャストはすでに送り完了です");
    renderPlanSelect();
    return;
  }

  const payload = {
    dispatch_id: currentDispatchId,
    cast_id: plan.cast_id,
    actual_hour: Number(plan.plan_hour || 0),
    stop_order:
      currentActualsCache.filter(x => Number(x.actual_hour) === Number(plan.plan_hour || 0)).length + 1,
    pickup_label: ORIGIN_LABEL,
    destination_address: plan.destination_address || plan.casts?.address || "",
    destination_area: normalizeAreaLabel(plan.planned_area || "無し"),
    distance_km: plan.distance_km ?? plan.casts?.distance_km ?? null,
    status: "pending",
    note: plan.note || "",
    plan_date: plan.plan_date
  };

  const { error } = await supabaseClient.from("dispatch_items").insert(payload);
  if (error) {
    alert(error.message);
    return;
  }

  await supabaseClient
    .from("dispatch_plans")
    .update({ status: "assigned" })
    .eq("id", plan.id);

  await addHistory(currentDispatchId, null, "add_plan_to_actual", `予定ID ${plan.id} をActualへ追加`);
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
  if (els.planSelect) els.planSelect.value = "";
  renderPlanSelect();
}
function renderActualTable() {
  if (!els.actualTableWrap) return;

  if (!currentActualsCache.length) {
    els.actualTableWrap.innerHTML = `<div class="muted" style="padding:14px;">Actualがありません</div>`;
    return;
  }

  const hours = [...new Set(currentActualsCache.map(x => Number(x.actual_hour ?? 0)))].sort((a, b) => a - b);
  let html = `<div class="grouped-actual-list">`;

  hours.forEach(hour => {
    const hourItems = currentActualsCache.filter(x => Number(x.actual_hour ?? 0) === hour);
    const groupedAreas = getGroupedAreasByDisplay(hourItems, x => x.destination_area || "無し");

    html += `<div class="grouped-section">`;
    html += `<div class="grouped-hour-title">${getHourLabel(hour)}</div>`;

    groupedAreas.forEach(({ detailArea }) => {
      html += `<div class="grouped-area-title">${getGroupedAreaHeaderHtml(detailArea)}</div>`;

      hourItems
        .filter(item => normalizeAreaLabel(item.destination_area || "無し") === detailArea)
        .forEach(item => {
          html += `
            <div class="grouped-row">
              <div>${getHourLabel(hour)}</div>
              <div><strong>${buildMapLinkHtml({
               name: item.casts?.name,
               address: item.destination_address || item.casts?.address,
               lat: item.casts?.latitude,
               lng: item.casts?.longitude
               })}</strong></div>
              <div>${escapeHtml(normalizeAreaLabel(item.destination_area || "無し"))}</div>
              <div>${item.distance_km ?? ""}</div>
              <div class="op-cell">
                <div class="state-stack">
                  <button class="btn primary actual-done-btn" data-id="${item.id}">完了</button>
                  <button class="btn danger actual-cancel-btn" data-id="${item.id}">キャンセル</button>
                  <span class="badge-status ${normalizeStatus(item.status)}">${escapeHtml(getStatusText(item.status))}</span>
                </div>
                <button class="btn ghost actual-edit-btn" data-id="${item.id}">編集</button>
                <button class="btn ghost actual-route-btn" data-address="${escapeHtml(item.destination_address || item.casts?.address || "")}">ルート</button>
                <button class="btn danger actual-delete-btn" data-id="${item.id}">削除</button>
              </div>
            </div>
          `;
        });
    });

    html += `</div>`;
  });

  html += `</div>`;
  els.actualTableWrap.innerHTML = html;

  els.actualTableWrap.querySelectorAll(".actual-edit-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const item = currentActualsCache.find(x => Number(x.id) === Number(btn.dataset.id));
      if (item) fillActualForm(item);
    });
  });

  els.actualTableWrap.querySelectorAll(".actual-route-btn").forEach(btn => {
    btn.addEventListener("click", () => openGoogleMap(btn.dataset.address || ""));
  });

  els.actualTableWrap.querySelectorAll(".actual-delete-btn").forEach(btn => {
    btn.addEventListener("click", async () => deleteActual(Number(btn.dataset.id)));
  });

  els.actualTableWrap.querySelectorAll(".actual-done-btn").forEach(btn => {
    btn.addEventListener("click", async () => updateActualStatus(Number(btn.dataset.id), "done"));
  });

  els.actualTableWrap.querySelectorAll(".actual-cancel-btn").forEach(btn => {
    btn.addEventListener("click", async () => updateActualStatus(Number(btn.dataset.id), "cancel"));
  });
}

function renderActualTimeAreaMatrix() {
  if (!els.actualTimeAreaMatrix) return;

  const hours = [0, 1, 2, 3, 4, 5];
  const areas = [
    ...new Set(
      currentActualsCache.map(x => getAreaDisplayGroup(normalizeAreaLabel(x.destination_area || "無し")))
    )
  ];

  if (!areas.length) {
    els.actualTimeAreaMatrix.innerHTML = `<div class="muted" style="padding:14px;">一覧がありません</div>`;
    return;
  }

  let html = `
    <table class="matrix-table">
      <thead>
        <tr>
          <th>時間</th>
          ${areas.map(a => `<th>${escapeHtml(a)}</th>`).join("")}
        </tr>
      </thead>
      <tbody>
  `;

  hours.forEach(hour => {
    html += `<tr><td>${getHourLabel(hour)}</td>`;

    areas.forEach(area => {
      const rows = currentActualsCache.filter(
        x =>
          Number(x.actual_hour ?? 0) === hour &&
          getAreaDisplayGroup(normalizeAreaLabel(x.destination_area || "無し")) === area
      );

      if (!rows.length) {
        html += `<td>-</td>`;
      } else {
        const totalDistance = rows.reduce((sum, row) => sum + Number(row.distance_km || 0), 0);

        html += `
          <td>
            <div class="matrix-card">
              <div class="matrix-summary">${rows.length}人 / ${totalDistance.toFixed(1)}km</div>
              ${rows
                .map(
                  row => `
                    <div class="matrix-item">
                      <span class="badge-status ${normalizeStatus(row.status)}">${escapeHtml(getStatusText(row.status))}</span>
                      <span>${buildMapLinkHtml({
                      name: row.casts?.name,
                      address: row.destination_address || row.casts?.address,
                      lat: row.casts?.latitude,
                      lng: row.casts?.longitude
                      })} (${Number(row.distance_km || 0).toFixed(1)}km)</span>
                    </div>
                  `
                )
                .join("")}
            </div>
          </td>
        `;
      }
    });

    html += `</tr>`;
  });

  html += `</tbody></table>`;
  els.actualTimeAreaMatrix.innerHTML = html;
}

function guessActualArea() {
  if (els.actualArea) {
    els.actualArea.value = normalizeAreaLabel(
      classifyAreaByAddress(els.actualAddress?.value || "") || "無し"
    );
  }
}

function renderDailyVehicleChecklist() {
  if (!els.dailyVehicleChecklist) return;

  els.dailyVehicleChecklist.innerHTML = "";

  if (!allVehiclesCache.length) {
    els.dailyVehicleChecklist.innerHTML = `<div class="muted">車両がありません</div>`;
    return;
  }

  allVehiclesCache.forEach(vehicle => {
    const row = document.createElement("div");
    row.className = "vehicle-check-item";
    row.innerHTML = `
      <div class="vehicle-check-label">
       ${escapeHtml(vehicle.driver_name || "-")}
      （${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))} / 帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))} / 定員${vehicle.seat_capacity ?? "-"}）
     </div>
      <input class="vehicle-check-input" type="checkbox" data-id="${vehicle.id}" ${activeVehicleIdsForToday.has(Number(vehicle.id)) ? "checked" : ""} />
    `;
    els.dailyVehicleChecklist.appendChild(row);
  });

  els.dailyVehicleChecklist.querySelectorAll(".vehicle-check-input").forEach(input => {
  input.addEventListener("change", () => {
    const id = Number(input.dataset.id);

    if (input.checked) activeVehicleIdsForToday.add(id);
    else activeVehicleIdsForToday.delete(id);

    renderDailyMileageInputs();   // ←追加①
    renderDailyDispatchResult();
  });
});
}

function getSelectedVehiclesForToday() {
  return allVehiclesCache.filter(v => activeVehicleIdsForToday.has(Number(v.id)));
}

function toggleAllVehicles(checked) {
  if (checked) {
    activeVehicleIdsForToday = new Set(allVehiclesCache.map(v => Number(v.id)));
  } else {
    activeVehicleIdsForToday = new Set();
  }
  renderDailyVehicleChecklist();
  renderDailyMileageInputs();
  renderDailyDispatchResult();
}


function getActiveDispatchItemsForAutoAssign() {
  return currentActualsCache.filter(
    item => !["done", "cancel"].includes(normalizeStatus(item.status))
  );
}

function sortItemsForFallbackDispatch(items) {
  return [...items].sort((a, b) => {
    const ah = Number(a.actual_hour ?? 0);
    const bh = Number(b.actual_hour ?? 0);
    if (ah !== bh) return ah - bh;

    const aa = normalizeAreaLabel(a.destination_area || "");
    const ba = normalizeAreaLabel(b.destination_area || "");
    if (aa !== ba) return aa.localeCompare(ba, "ja");

    return Number(a.stop_order || 0) - Number(b.stop_order || 0);
  });
}

function buildFallbackAssignments(items, vehicles) {
  const orderedItems = sortItemsForFallbackDispatch(items);
  if (!orderedItems.length || !vehicles.length) return [];

  const assignments = [];
  const seatLoads = new Map();

  function getLoad(vehicleId, hour) {
    return Number(seatLoads.get(`${vehicleId}_${hour}`) || 0);
  }

  function addLoad(vehicleId, hour) {
    const key = `${vehicleId}_${hour}`;
    seatLoads.set(key, getLoad(vehicleId, hour) + 1);
  }

  orderedItems.forEach((item, index) => {
    const hour = Number(item.actual_hour ?? 0);
    let selectedVehicle =
      vehicles.find(v => getLoad(v.id, hour) < Number(v.seat_capacity || 4)) ||
      vehicles[index % vehicles.length];

    if (!selectedVehicle) selectedVehicle = vehicles[0];

    assignments.push({
      item_id: item.id,
      actual_hour: hour,
      vehicle_id: selectedVehicle.id,
      driver_name: selectedVehicle.driver_name || "",
      distance_km: Number(item.distance_km || 0)
    });

    addLoad(selectedVehicle.id, hour);
  });

  return assignments;
}

function applyManualLastVehicleToAssignments(assignments, vehicles) {
  if (!assignments.length || !vehicles.length) return assignments;

  const dateStr = els.actualDate?.value || todayStr();
  const defaultLastHour = getDefaultLastHour(dateStr);
  const manualLastVehicleId = getManualLastVehicleId();
  const manualVehicle = vehicles.find(v => Number(v.id) === Number(manualLastVehicleId)) || null;
  const vehicleMap = new Map(vehicles.map(v => [Number(v.id), v]));
  const itemMap = new Map(currentActualsCache.map(item => [Number(item.id), item]));

  const hourCounts = new Map();
  assignments.forEach(a => {
    const key = `${Number(a.vehicle_id)}__${Number(a.actual_hour ?? 0)}`;
    hourCounts.set(key, Number(hourCounts.get(key) || 0) + 1);
  });

  const lastHourAssignments = assignments.filter(a => Number(a.actual_hour ?? 0) === Number(defaultLastHour));
  if (lastHourAssignments.length) {
    const optimizedLastHour = assignments.map(a => ({ ...a }));

    for (const target of optimizedLastHour) {
      if (Number(target.actual_hour ?? 0) !== Number(defaultLastHour)) continue;

      const item = itemMap.get(Number(target.item_id));
      if (!item) continue;
      const area = normalizeAreaLabel(item.destination_area || item.cluster_area || '無し');

      let bestCandidate = null;
      for (const vehicle of vehicles) {
        const seatCapacity = Number(vehicle.seat_capacity || 4);
        const currentKey = `${Number(vehicle.id)}__${Number(defaultLastHour)}`;
        const currentLoad = Number(hourCounts.get(currentKey) || 0);
        const isSameVehicle = Number(vehicle.id) === Number(target.vehicle_id);
        if (!isSameVehicle && currentLoad >= seatCapacity) continue;

        const currentVehicle = vehicleMap.get(Number(target.vehicle_id));
        const currentHome = normalizeAreaLabel(currentVehicle?.home_area || '');
        const candidateHome = normalizeAreaLabel(vehicle.home_area || '');
        const candidateAffinity = getAreaAffinityScore(area, candidateHome);
        const currentAffinity = getAreaAffinityScore(area, currentHome);
        const candidateDirection = getDirectionAffinityScore(area, candidateHome);
        const currentDirection = getDirectionAffinityScore(area, currentHome);
        const candidateStrict = getStrictHomeCompatibilityScore(area, candidateHome);
        const currentStrict = getStrictHomeCompatibilityScore(area, currentHome);
        const candidateReverse = isHardReverseForHome(area, candidateHome);
        const vehicleAreaScore = getVehicleAreaMatchScore(vehicle, area);
        const manualBonus = manualVehicle && Number(vehicle.id) === Number(manualVehicle.id) ? 6 : 0;
        const reversePenalty = candidateReverse ? 240 : 0;
        const score = candidateAffinity * 8 + Math.max(candidateDirection, 0) * 4 + candidateStrict * 5 + vehicleAreaScore + manualBonus - currentLoad * 8 - reversePenalty;

        if (!bestCandidate || score > bestCandidate.score) {
          bestCandidate = { vehicle, score, candidateAffinity, currentAffinity };
        }
      }

      if (!bestCandidate) continue;
      const currentVehicle = vehicleMap.get(Number(target.vehicle_id));
      const currentHome = normalizeAreaLabel(currentVehicle?.home_area || '');
      const currentAffinity = getAreaAffinityScore(area, currentHome);
      const currentDirection = getDirectionAffinityScore(area, currentHome);
      const currentStrict = getStrictHomeCompatibilityScore(area, currentHome);
      const bestDirection = getDirectionAffinityScore(area, bestCandidate.vehicle.home_area || '');
      const bestStrict = getStrictHomeCompatibilityScore(area, bestCandidate.vehicle.home_area || '');
      const shouldMove =
        Number(bestCandidate.vehicle.id) !== Number(target.vehicle_id) &&
        (
          bestStrict > currentStrict ||
          (bestStrict === currentStrict && bestDirection > currentDirection) ||
          (bestStrict === currentStrict && bestDirection === currentDirection && bestCandidate.candidateAffinity > currentAffinity) ||
          (bestStrict === currentStrict && bestDirection === currentDirection && bestCandidate.candidateAffinity === currentAffinity && manualVehicle && Number(bestCandidate.vehicle.id) === Number(manualVehicle.id))
        );

      if (!shouldMove) continue;

      const fromKey = `${Number(target.vehicle_id)}__${Number(defaultLastHour)}`;
      const toKey = `${Number(bestCandidate.vehicle.id)}__${Number(defaultLastHour)}`;
      hourCounts.set(fromKey, Math.max(0, Number(hourCounts.get(fromKey) || 0) - 1));
      hourCounts.set(toKey, Number(hourCounts.get(toKey) || 0) + 1);

      target.vehicle_id = bestCandidate.vehicle.id;
      target.driver_name = bestCandidate.vehicle.driver_name || '';
      target.manual_last_vehicle = manualVehicle && Number(bestCandidate.vehicle.id) === Number(manualVehicle.id);
    }

    return optimizedLastHour;
  }

  if (!manualVehicle) return assignments;

  const sorted = [...assignments].sort((a, b) => {
    const ah = Number(a.actual_hour ?? 0);
    const bh = Number(b.actual_hour ?? 0);
    if (ah !== bh) return ah - bh;
    return Number(a.item_id || 0) - Number(b.item_id || 0);
  });

  const last = sorted[sorted.length - 1];
  if (!last) return assignments;

  const lastItem = itemMap.get(Number(last.item_id));
  const lastArea = normalizeAreaLabel(lastItem?.destination_area || '');
  const manualAffinity = getAreaAffinityScore(lastArea, manualVehicle.home_area || '');
  const currentVehicle = vehicleMap.get(Number(last.vehicle_id));
  const currentAffinity = getAreaAffinityScore(lastArea, currentVehicle?.home_area || '');
  const manualDirection = getDirectionAffinityScore(lastArea, manualVehicle.home_area || '');
  const currentDirection = getDirectionAffinityScore(lastArea, currentVehicle?.home_area || '');
  const manualStrict = getStrictHomeCompatibilityScore(lastArea, manualVehicle.home_area || '');
  const currentStrict = getStrictHomeCompatibilityScore(lastArea, currentVehicle?.home_area || '');

  if (isHardReverseForHome(lastArea, manualVehicle.home_area || '')) return assignments;
  if (manualStrict < currentStrict) return assignments;
  if (manualStrict === currentStrict && manualDirection < currentDirection) return assignments;
  if (manualStrict === currentStrict && manualDirection === currentDirection && manualAffinity < currentAffinity) return assignments;

  return assignments.map(a =>
    Number(a.item_id) === Number(last.item_id)
      ? {
          ...a,
          vehicle_id: manualVehicle.id,
          driver_name: manualVehicle.driver_name || '',
          manual_last_vehicle: true
        }
      : a
  );
}


function buildMonthlyDistanceMapForCurrentMonth() {
  const monthKey = getMonthKey(els.dispatchDate?.value || todayStr());
  return getVehicleMonthlyStatsMap(currentDailyReportsCache, monthKey);
}


function getCurrentClockMinutes() {
  const now = new Date();
  return now.getHours() * 60 + now.getMinutes();
}

function formatClockTimeFromMinutesGlobal(totalMinutes) {
  const safe = Math.max(0, Math.round(Number(totalMinutes || 0)));
  const h = Math.floor(safe / 60) % 24;
  const m = safe % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

function getDistanceZoneInfoGlobal(distanceKm) {
  const km = Number(distanceKm || 0);
  if (km <= 10) return { key: "short", label: "近距離", speedKmh: 25 };
  if (km <= 25) return { key: "middle", label: "中距離", speedKmh: 30 };
  return { key: "long", label: "長距離", speedKmh: 35 };
}

function estimateTravelMinutesByDistanceGlobal(distanceKm) {
  const km = Math.max(0, Number(distanceKm || 0));
  const zone = getDistanceZoneInfoGlobal(km);
  const speed = Number(zone.speedKmh || 30);
  if (!speed) return 0;
  return Math.round((km / speed) * 60);
}

function calculateRouteDistanceGlobal(items) {
  if (!items || !items.length) return 0;
  let total = 0;
  let currentLat = ORIGIN_LAT;
  let currentLng = ORIGIN_LNG;

  items.forEach(item => {
    const point = getItemLatLng(item);
    if (point) {
      total += estimateRoadKmBetweenPoints(currentLat, currentLng, point.lat, point.lng);
      currentLat = point.lat;
      currentLng = point.lng;
    } else {
      total += Number(item.distance_km || 0);
    }
  });

  return Number(total.toFixed(1));
}

function calcVehicleRotationForecastGlobal(vehicle, orderedRows) {
  const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
  if (!rows.length) {
    return {
      routeDistanceKm: 0,
      returnDistanceKm: 0,
      zoneLabel: "-",
      predictedReturnTime: "-",
      predictedReadyTime: "-",
      predictedReturnMinutes: 0,
      extraSharedDelayMinutes: 0,
      stopCount: 0,
      returnAfterLabel: "-"
    };
  }

  const firstHour = rows.reduce((min, row) => {
    const val = Number(row.actual_hour ?? row.plan_hour ?? 0);
    return Number.isFinite(val) ? Math.min(min, val) : min;
  }, 99);

  const baseHour = firstHour === 99 ? 0 : firstHour;
  const routeDistanceKm = Number(calculateRouteDistanceGlobal(rows) || 0);
  const lastRow = rows[rows.length - 1] || {};
  const returnDistanceKm = Number(lastRow.distance_km || 0);
  const primaryZone = getDistanceZoneInfoGlobal(Math.max(routeDistanceKm, returnDistanceKm));

  let departDelayMinutes = 20;
  if (baseHour === 3) departDelayMinutes = 18;
  else if (baseHour === 4) departDelayMinutes = 12;
  else if (baseHour >= 5) departDelayMinutes = 8;

  const outboundMinutes = estimateTravelMinutesByDistanceGlobal(routeDistanceKm);
  const returnMinutes = estimateTravelMinutesByDistanceGlobal(returnDistanceKm);
  const dropoffMinutes = rows.length * 1;

  const baseStartMinutes = Number.isFinite(lastAutoDispatchRunAtMinutes) && lastAutoDispatchRunAtMinutes !== null
    ? lastAutoDispatchRunAtMinutes
    : (baseHour * 60 + departDelayMinutes);

  const predictedReturnMinutes = Math.round(outboundMinutes + dropoffMinutes + returnMinutes);
  const predictedReturnAbs = baseStartMinutes + predictedReturnMinutes;
  const predictedReadyAbs = predictedReturnAbs + 1;

  let extraSharedDelayMinutes = 0;
  if (rows.length >= 2) {
    const firstOnly = [rows[0]];
    const singleRouteDistanceKm = Number(calculateRouteDistanceGlobal(firstOnly) || rows[0].distance_km || 0);
    const singleReturnDistanceKm = Number(rows[0].distance_km || 0);
    const singleOutbound = estimateTravelMinutesByDistanceGlobal(singleRouteDistanceKm);
    const singleReturn = estimateTravelMinutesByDistanceGlobal(singleReturnDistanceKm);
    const singlePredictedReturnMinutes = Math.round(singleOutbound + 1 + singleReturn);
    extraSharedDelayMinutes = Math.max(0, predictedReturnMinutes - singlePredictedReturnMinutes);
  }

  return {
    routeDistanceKm,
    returnDistanceKm,
    zoneLabel: primaryZone.label,
    predictedReturnTime: formatClockTimeFromMinutesGlobal(predictedReturnAbs),
    predictedReadyTime: formatClockTimeFromMinutesGlobal(predictedReadyAbs),
    predictedReturnMinutes,
    extraSharedDelayMinutes: Math.round(extraSharedDelayMinutes),
    stopCount: rows.length,
    returnAfterLabel: `${predictedReturnMinutes}分後`
  };
}

function calcVehicleDailyStatsGlobal(vehicleId, items) {
  const rows = (items || []).filter(i => Number(i.vehicle_id) === Number(vehicleId));
  if (!rows.length) {
    return {
      totalKm: 0,
      sendKm: 0,
      returnKm: 0,
      driveMinutes: 0,
      count: 0
    };
  }

  const orderedRows = moveManualLastItemsToEnd(
    sortItemsByNearestRoute(
      [...rows].sort((a, b) => Number(a.actual_hour ?? 0) - Number(b.actual_hour ?? 0))
    )
  );

  const sendKm = Number(calculateRouteDistanceGlobal(orderedRows) || 0);
  const lastRow = orderedRows[orderedRows.length - 1] || {};
  const returnKm = Number(lastRow.distance_km || 0);
  const totalKm = Number((sendKm + returnKm).toFixed(1));

  let driveMinutes =
    estimateTravelMinutesByDistanceGlobal(sendKm) +
    estimateTravelMinutesByDistanceGlobal(returnKm) +
    orderedRows.length * 1;

  return {
    totalKm,
    sendKm: Number(sendKm.toFixed(1)),
    returnKm: Number(returnKm.toFixed(1)),
    driveMinutes: Math.round(driveMinutes),
    count: orderedRows.length
  };
}

function optimizeAssignments(items, vehicles, monthlyMap) {
  const workingVehicles = vehicles.filter(v => v.status !== "maintenance");
  const manualLastVehicleId = getManualLastVehicleId();
  const clusters = buildDispatchClusters(items);
  const assignments = [];

  if (!workingVehicles.length || !clusters.length) return assignments;

  const vehicleUsage = new Map();

  function getVehicleState(vehicleId) {
    if (!vehicleUsage.has(vehicleId)) {
      vehicleUsage.set(vehicleId, {
        totalAssigned: 0,
        totalDistance: 0,
        hourLoads: new Map(),
        hourAreas: new Map()
      });
    }
    return vehicleUsage.get(vehicleId);
  }

  function getHourLoad(vehicleId, hour) {
    const state = getVehicleState(vehicleId);
    return Number(state.hourLoads.get(hour) || 0);
  }

  function getHourAreas(vehicleId, hour) {
    const state = getVehicleState(vehicleId);
    return [...(state.hourAreas.get(hour) || [])];
  }

  function addHourLoad(vehicleId, hour, count, distance, area) {
    const state = getVehicleState(vehicleId);
    state.totalAssigned += count;
    state.totalDistance += distance;
    state.hourLoads.set(hour, getHourLoad(vehicleId, hour) + count);
    if (area) {
      const existing = state.hourAreas.get(hour) || [];
      state.hourAreas.set(hour, [...existing, normalizeAreaLabel(area)]);
    }
  }

  function getIdleVehicleCountForHour(hour) {
    return workingVehicles.filter(vehicle => getHourLoad(vehicle.id, hour) === 0).length;
  }

  function shouldPreferSpread(cluster, routeFlowScore, routeContinuityPenalty, sameHourLoad) {
    if (cluster.count <= 1) return true;
    if (sameHourLoad <= 0) return true;
    const strongRouteLink = routeFlowScore >= 90 && routeContinuityPenalty <= 45;
    const reasonableRideShare = routeFlowScore >= 70 && routeContinuityPenalty <= 60;
    if (strongRouteLink) return false;
    if (cluster.count <= 2 && reasonableRideShare) return false;
    return true;
  }

  function getDistanceZoneForAi(distanceKm) {
    const km = Number(distanceKm || 0);
    if (km <= 10) return "near";
    if (km <= 25) return "mid";
    return "far";
  }

  function getZoneSpeedKmh(zone) {
    if (zone === "near") return 25;
    if (zone === "mid") return 30;
    return 35;
  }

  function getBaseDispatchDelayMinutes(hour) {
    const h = Number(hour || 0);
    if (h <= 1) return 20;
    if (h === 2) return 20;
    if (h === 3) return 18;
    if (h === 4) return 12;
    return 8;
  }

  function estimateTravelMinutesByDistance(distanceKm) {
    const km = Math.max(0, Number(distanceKm || 0));
    const zone = getDistanceZoneForAi(km);
    const speed = getZoneSpeedKmh(zone);
    if (!speed) return 0;
    return Math.round((km / speed) * 60);
  }

  function estimateDropoffMinutes(stopCount) {
    return Math.max(1, Number(stopCount || 1)) * 1;
  }

  function estimateRotationReadyMinutes(hour, distanceKm, stopCount) {
    const travelOut = estimateTravelMinutesByDistance(distanceKm);
    const dropoff = estimateDropoffMinutes(stopCount);
    const returnTrip = estimateTravelMinutesByDistance(distanceKm);
    return getBaseDispatchDelayMinutes(hour) + travelOut + dropoff + returnTrip;
  }

  function estimateRideShareExtraDelay(distanceKm, stopCount, sameHourLoad, routeFlowScore, routeContinuityPenalty) {
    if (Number(sameHourLoad || 0) <= 0) return 0;

    const zone = getDistanceZoneForAi(distanceKm);
    const zoneBase =
      zone === "near" ? 2 :
      zone === "mid" ? 5 : 8;

    const raw =
      zoneBase +
      Number(sameHourLoad || 0) * 2 +
      Number(stopCount || 1) * 1 +
      Number(routeContinuityPenalty || 0) / 30 -
      Number(routeFlowScore || 0) / 40;

    return Math.max(0, Math.round(raw));
  }

  function getRotationPredictionScore(hour, distanceKm, stopCount, sameHourLoad, routeFlowScore, routeContinuityPenalty, idleVehicleCount) {
    const predictedReadyMinutes = estimateRotationReadyMinutes(hour, distanceKm, stopCount);
    const extraDelay = estimateRideShareExtraDelay(
      distanceKm,
      stopCount,
      sameHourLoad,
      routeFlowScore,
      routeContinuityPenalty
    );
    const canShare = extraDelay <= 8;

    let score = predictedReadyMinutes * 0.65 + extraDelay * 1.8;

    if (Number(sameHourLoad || 0) === 0 && Number(idleVehicleCount || 0) > 0) {
      score -= 24;
    }

    if (Number(sameHourLoad || 0) > 0 && !canShare) {
      score += 42;
    } else if (Number(sameHourLoad || 0) > 0 && canShare) {
      score -= 10;
    }

    return {
      predictedReadyMinutes,
      extraDelay,
      canShare,
      score
    };
  }

  function getNormalRunReturnPenalty(hour, addedDistance, stopCount, sameHourLoad, routeFlowScore, routeContinuityPenalty, idleVehicleCount) {
    const rotation = getRotationPredictionScore(
      hour,
      addedDistance,
      stopCount,
      sameHourLoad,
      routeFlowScore,
      routeContinuityPenalty,
      idleVehicleCount
    );
    return rotation.score;
  }

  function formatClockTimeFromMinutes(totalMinutes) {
    const safe = Math.max(0, Math.round(Number(totalMinutes || 0)));
    const h = Math.floor(safe / 60) % 24;
    const m = safe % 60;
    return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
  }

  function getDistanceZoneInfo(distanceKm) {
    const km = Number(distanceKm || 0);
    if (km <= 10) return { key: "short", label: "近距離", speedKmh: 25 };
    if (km <= 25) return { key: "middle", label: "中距離", speedKmh: 30 };
    return { key: "long", label: "長距離", speedKmh: 35 };
  }

  function getExpectedDepartureDelayMinutes(baseHour) {
    const hour = Number(baseHour || 0);
    if (hour <= 2) return 20;
    if (hour === 3) return 18;
    if (hour === 4) return 12;
    if (hour >= 5) return 8;
    return 20;
  }

  function calcVehicleRotationForecast(vehicle, orderedRows) {
    const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
    if (!rows.length) {
      return {
        routeDistanceKm: 0,
        returnDistanceKm: 0,
        zoneLabel: "-",
        predictedDepartureTime: "-",
        predictedReturnTime: "-",
        predictedReadyTime: "-",
        predictedReturnMinutes: 0,
        extraSharedDelayMinutes: 0,
        stopCount: 0
      };
    }
  
    const firstHour = rows.reduce((min, row) => {
      const val = Number(row.actual_hour ?? row.plan_hour ?? 0);
      return Number.isFinite(val) ? Math.min(min, val) : min;
    }, 99);
  
    const baseHour = firstHour === 99 ? 0 : firstHour;
    const routeDistanceKm = Number(calculateRouteDistance(rows) || 0);
    const lastRow = rows[rows.length - 1] || {};
    const returnDistanceKm = Number(lastRow.distance_km || 0);
    const primaryZone = getDistanceZoneInfo(Math.max(routeDistanceKm, returnDistanceKm));
  
    const departDelayMinutes = getExpectedDepartureDelayMinutes(baseHour);
    const outboundMinutes = estimateTravelMinutesByDistance(routeDistanceKm);
    const returnMinutes = estimateTravelMinutesByDistance(returnDistanceKm);
    const dropoffMinutes = rows.length * 1;
  
    const predictedDepartureAbs = baseHour * 60 + departDelayMinutes;
    const predictedReturnAbs = predictedDepartureAbs + outboundMinutes + dropoffMinutes + returnMinutes;
    const predictedReadyAbs = predictedReturnAbs + 1;
  
    let extraSharedDelayMinutes = 0;
    if (rows.length >= 2) {
      const firstOnly = [rows[0]];
      const singleRouteDistanceKm = Number(calculateRouteDistance(firstOnly) || rows[0].distance_km || 0);
      const singleReturnDistanceKm = Number(rows[0].distance_km || 0);
      const singleOutbound = estimateTravelMinutesByDistance(singleRouteDistanceKm);
      const singleReturn = estimateTravelMinutesByDistance(singleReturnDistanceKm);
      const singleDropoff = 1;
      const singlePredictedReturnAbs = predictedDepartureAbs + singleOutbound + singleDropoff + singleReturn;
      extraSharedDelayMinutes = Math.max(0, predictedReturnAbs - singlePredictedReturnAbs);
    }
  
    return {
      routeDistanceKm,
      returnDistanceKm,
      zoneLabel: primaryZone.label,
      predictedDepartureTime: formatClockTimeFromMinutes(predictedDepartureAbs),
      predictedReturnTime: formatClockTimeFromMinutes(predictedReturnAbs),
      predictedReadyTime: formatClockTimeFromMinutes(predictedReadyAbs),
      predictedReturnMinutes: Math.round(predictedReturnAbs - predictedDepartureAbs),
      extraSharedDelayMinutes: Math.round(extraSharedDelayMinutes),
      stopCount: rows.length
    };
  }

  function buildRotationTimelineHtml(vehicles, activeItems) {
    const timeline = vehicles
      .map(vehicle => {
        const rows = moveManualLastItemsToEnd(
          sortItemsByNearestRoute(
            activeItems
              .filter(item => Number(item.vehicle_id) === Number(vehicle.id))
              .sort((a, b) => Number(a.actual_hour ?? 0) - Number(b.actual_hour ?? 0))
          )
        );
        const forecast = calcVehicleRotationForecast(vehicle, rows);
        return {
          name: vehicle.driver_name || vehicle.plate_number || "-",
          readyTime: forecast.predictedReadyTime,
          returnTime: forecast.predictedReturnTime,
          rotationMinutes: forecast.predictedReturnMinutes,
          hasRows: rows.length > 0
        };
      })
      .filter(x => x.hasRows)
      .sort((a, b) => a.readyTime.localeCompare(b.readyTime));
  
    if (!timeline.length) return "";
  
    return `
      <div class="panel-card" style="margin-bottom:16px;">
        <h3 style="margin-bottom:10px;">車両稼働タイムライン</h3>
        <div style="display:flex; flex-wrap:wrap; gap:8px;">
          ${timeline.map(item => `
            <div class="chip" style="padding:8px 12px;">
              <strong>${escapeHtml(item.name)}</strong>
              / 戻り ${escapeHtml(item.returnTime)}
              / 次便可能 ${escapeHtml(item.readyTime)}
              / 回転 ${item.rotationMinutes}分
            </div>
          `).join("")}
        </div>
      </div>
    `;
  }

  function isDirectionSplitPair(baseArea, compareArea) {
    const a = normalizeAreaLabel(baseArea);
    const b = normalizeAreaLabel(compareArea);
    if (!a || !b) return false;
    const affinity = getDirectionAffinityScore(a, b);
    return affinity <= -38;
  }

  function getDirectionSplitGuardScore(clusterArea, existingAreas, distanceKm, isLastRun, isDefaultLastHourCluster) {
    if (isLastRun || isDefaultLastHourCluster) return 0;

    const zone = getDistanceZoneForAi(distanceKm);
    const areas = (existingAreas || []).map(normalizeAreaLabel).filter(Boolean);
    if (!areas.length) return 0;

    let score = 0;
    for (const area of areas) {
      if (!isDirectionSplitPair(clusterArea, area)) continue;

      if (zone === 'long') score += 260;
      else if (zone === 'mid') score += 120;
      else score += 40;
    }

    return score;
  }

  function isHardDirectionSplitBlocked(clusterArea, existingAreas, distanceKm, isLastRun, isDefaultLastHourCluster) {
    if (isLastRun || isDefaultLastHourCluster) return false;

    const zone = getDistanceZoneForAi(distanceKm);
    if (zone !== 'long') return false;

    const areas = (existingAreas || []).map(normalizeAreaLabel).filter(Boolean);
    if (!areas.length) return false;

    return areas.some(area => isDirectionSplitPair(clusterArea, area));
  }

  for (const cluster of clusters) {
    const dateStr = els.actualDate?.value || todayStr();
    const isLastRun = isLastClusterOfTheDay(cluster, dateStr);
    const defaultLastHour = getDefaultLastHour(dateStr);
    const isDefaultLastHourCluster = Number(cluster.hour) === Number(defaultLastHour);

    const candidateScores = workingVehicles
      .map(vehicle => {
        const seatCapacity = Number(vehicle.seat_capacity || 4);
        const sameHourLoad = getHourLoad(vehicle.id, cluster.hour);

        if (sameHourLoad >= seatCapacity) return null;
        if (sameHourLoad + cluster.count > seatCapacity) return null;

        const monthly = monthlyMap.get(Number(vehicle.id)) || {
          totalDistance: 0,
          workedDays: 0,
          avgDistance: 0
        };

        const normalizedClusterArea = normalizeAreaLabel(cluster.area);
        const homeArea = normalizeAreaLabel(vehicle?.home_area || "");

        const projectedWorkedDays = Math.max(Number(monthly.workedDays || 0), 1);
        const projectedAvg =
          (Number(monthly.totalDistance || 0) + Number(cluster.totalDistance || 0)) /
          projectedWorkedDays;

        let score = 1000;

        const vehicleAreaMatch = getVehicleAreaMatchScore(vehicle, cluster.area);
        const directionAffinity = getDirectionAffinityScore(normalizedClusterArea, homeArea);
        const strictHomeScore = getStrictHomeCompatibilityScore(normalizedClusterArea, homeArea);
        const hardReverse = isHardReverseForHome(normalizedClusterArea, homeArea);
        const existingHourAreas = getHourAreas(vehicle.id, cluster.hour);
        const routeFlowScore = getRouteFlowVehicleScore(normalizedClusterArea, existingHourAreas, homeArea);
        const routeContinuityPenalty = getRouteContinuityPenalty(normalizedClusterArea, existingHourAreas, homeArea);
        const directionSplitGuardScore = getDirectionSplitGuardScore(
          normalizedClusterArea,
          existingHourAreas,
          Number(cluster.totalDistance || 0),
          isLastRun,
          isDefaultLastHourCluster
        );
        const hardDirectionSplitBlocked = isHardDirectionSplitBlocked(
          normalizedClusterArea,
          existingHourAreas,
          Number(cluster.totalDistance || 0),
          isLastRun,
          isDefaultLastHourCluster
        );

        // 方面一致 + 方向クラスタ + 帰宅適合 + 経由ルート相性を優先
        score -= vehicleAreaMatch;
        score -= routeFlowScore * (isLastRun || isDefaultLastHourCluster ? 2.05 : 1.35);
        score += routeContinuityPenalty * (isLastRun || isDefaultLastHourCluster ? 1.75 : 1.1);
        score -= Math.max(directionAffinity, 0) * (isLastRun || isDefaultLastHourCluster ? 2.2 : 0.45);
        score -= strictHomeScore * (isLastRun || isDefaultLastHourCluster ? 2.8 : 0.55);

        // 同時間帯の過積載を避ける
        score += sameHourLoad * 35;

        // 今日の割当数の偏りを抑える
        score += getVehicleState(vehicle.id).totalAssigned * 8;

        // 今日の車両負荷が高すぎる車両を少し避ける
        score += getVehicleState(vehicle.id).totalDistance * 0.18;

        // 月間平均距離が高い車両に積みすぎない
        score += projectedAvg * 0.55;

        // 通常便は、空き車両があれば分散優先で松戸駅へ早く戻れるようにする
        if (!(isLastRun || isDefaultLastHourCluster)) {
          const idleVehicleCount = getIdleVehicleCountForHour(cluster.hour);
          const rotationPrediction = getRotationPredictionScore(
            cluster.hour,
            Number(cluster.totalDistance || 0),
            Number(cluster.count || 1),
            sameHourLoad,
            routeFlowScore,
            routeContinuityPenalty,
            idleVehicleCount
          );
          const preferSpread =
            shouldPreferSpread(cluster, routeFlowScore, routeContinuityPenalty, sameHourLoad) ||
            !rotationPrediction.canShare;

          if (sameHourLoad === 0) {
            score -= idleVehicleCount > 1 ? 150 : 95;
          } else if (preferSpread && idleVehicleCount > 0) {
            score += 165 + sameHourLoad * 38;
          } else if (!preferSpread && rotationPrediction.canShare) {
            score -= Math.min(routeFlowScore, 120) * 0.40;
            score += routeContinuityPenalty * 0.15;
            score -= 18;
          }

          score += getNormalRunReturnPenalty(
            cluster.hour,
            Number(cluster.totalDistance || 0),
            Number(cluster.count || 1),
            sameHourLoad,
            routeFlowScore,
            routeContinuityPenalty,
            idleVehicleCount
          );
        }

        // ラスト便で逆方向は強く除外
        if ((isLastRun || isDefaultLastHourCluster) && hardReverse) {
          score += isLastRun ? 560 : 320;
        }

        // ラスト便時間帯では適合の弱い車両を避ける
        if ((isLastRun || isDefaultLastHourCluster) && strictHomeScore < 50) {
          score += 110;
        }

        // 長距離ゾーンで方向が割れる組み合わせを強く避ける
        score += directionSplitGuardScore;
        if (hardDirectionSplitBlocked) {
          score += 420;
        }

        // 同一ルートの折れ返しが大きい組み合わせを避ける
        if (routeContinuityPenalty >= 170) {
          score += isLastRun || isDefaultLastHourCluster ? 180 : 95;
        } else if (routeContinuityPenalty >= 95) {
          score += isLastRun || isDefaultLastHourCluster ? 90 : 42;
        }

        // 手動ラスト便車両は適合している時だけ優遇
        if (manualLastVehicleId && (isLastRun || isDefaultLastHourCluster)) {
          if (Number(vehicle.id) === Number(manualLastVehicleId) && !hardReverse && strictHomeScore >= 50) score -= 180;
          else if (Number(vehicle.id) === Number(manualLastVehicleId) && hardReverse) score += 220;
          else score += 35;
        }

        // ラスト便は帰宅方面の近接スコアを強く優遇
        const homePriorityWeight = getLastTripHomePriorityWeight(
          normalizedClusterArea,
          homeArea,
          isLastRun,
          isDefaultLastHourCluster
        );
        score -= homePriorityWeight;

        if ((isLastRun || isDefaultLastHourCluster) && homePriorityWeight <= 0) {
          score += isLastRun ? 80 : 28;
        }

        return { vehicle, score };
      })
      .filter(Boolean)
      .sort((a, b) => a.score - b.score);

    if (candidateScores.length) {
      const bestVehicle = candidateScores[0].vehicle;
      const sortedItems = sortItemsByNearestRoute(cluster.items);
      const bestHourLoad = getHourLoad(bestVehicle.id, cluster.hour);
      const bestExistingAreas = getHourAreas(bestVehicle.id, cluster.hour);
      const bestRouteFlowScore = getRouteFlowVehicleScore(normalizeAreaLabel(cluster.area), bestExistingAreas, normalizeAreaLabel(bestVehicle?.home_area || ""));
      const bestRouteContinuityPenalty = getRouteContinuityPenalty(normalizeAreaLabel(cluster.area), bestExistingAreas, normalizeAreaLabel(bestVehicle?.home_area || ""));
      const idleVehicleCount = getIdleVehicleCountForHour(cluster.hour);
      const keepTogether = (isLastRun || isDefaultLastHourCluster)
        ? true
        : !shouldPreferSpread(cluster, bestRouteFlowScore, bestRouteContinuityPenalty, bestHourLoad) || idleVehicleCount <= 1;

      if (keepTogether) {
        sortedItems.forEach(item => {
          assignments.push({
            item_id: item.id,
            actual_hour: cluster.hour,
            vehicle_id: bestVehicle.id,
            driver_name: bestVehicle.driver_name || "",
            distance_km: Number(item.distance_km || 0)
          });
        });

        addHourLoad(
          bestVehicle.id,
          cluster.hour,
          cluster.count,
          calculateRouteDistance(sortedItems),
          cluster.area
        );
        continue;
      }
    }

    const splitItems = sortItemsByNearestRoute(cluster.items);

    for (const item of splitItems) {
      const perItemCandidates = workingVehicles
        .map(vehicle => {
          const seatCapacity = Number(vehicle.seat_capacity || 4);
          const sameHourLoad = getHourLoad(vehicle.id, cluster.hour);

          if (sameHourLoad >= seatCapacity) return null;

          const monthly = monthlyMap.get(Number(vehicle.id)) || {
            totalDistance: 0,
            workedDays: 0,
            avgDistance: 0
          };

          const normalizedClusterArea = normalizeAreaLabel(cluster.area);
          const homeArea = normalizeAreaLabel(vehicle?.home_area || "");
          const existingHourAreas = getHourAreas(vehicle.id, cluster.hour);
          const routeFlowScore = getRouteFlowVehicleScore(normalizedClusterArea, existingHourAreas, homeArea);
          const routeContinuityPenalty = getRouteContinuityPenalty(normalizedClusterArea, existingHourAreas, homeArea);
          const directionSplitGuardScore = getDirectionSplitGuardScore(
            normalizedClusterArea,
            existingHourAreas,
            Number(item.distance_km || 0),
            isLastRun,
            isDefaultLastHourCluster
          );
          const hardDirectionSplitBlocked = isHardDirectionSplitBlocked(
            normalizedClusterArea,
            existingHourAreas,
            Number(item.distance_km || 0),
            isLastRun,
            isDefaultLastHourCluster
          );

          const projectedWorkedDays = Math.max(Number(monthly.workedDays || 0), 1);
          const projectedAvg =
            (Number(monthly.totalDistance || 0) + Number(item.distance_km || 0)) /
            projectedWorkedDays;

          let score = 1000;

          // 方面一致 + 経由ルート相性を優先
          score -= getVehicleAreaMatchScore(vehicle, cluster.area);
          score -= routeFlowScore * (isLastRun || isDefaultLastHourCluster ? 1.85 : 1.25);
          score += routeContinuityPenalty * (isLastRun || isDefaultLastHourCluster ? 1.55 : 1.0);

          // 同時間帯負荷
          score += sameHourLoad * 35;

          // 今日の割当数偏り
          score += getVehicleState(vehicle.id).totalAssigned * 8;

          // 今日の累積ルート距離偏り
          score += getVehicleState(vehicle.id).totalDistance * 0.18;

          // 月間平均の高い車両へ積みすぎない
          score += projectedAvg * 0.55;

          // 長距離ゾーンで方向が割れる組み合わせを強く避ける
          score += directionSplitGuardScore;
          if (hardDirectionSplitBlocked) {
            score += 420;
          }

          // 通常便は空いている車両を使い、戻り時間が短くなるよう分散を優先
          if (!(isLastRun || isDefaultLastHourCluster)) {
            const idleVehicleCount = getIdleVehicleCountForHour(cluster.hour);
            const rotationPrediction = getRotationPredictionScore(
              cluster.hour,
              Number(item.distance_km || 0),
              1,
              sameHourLoad,
              routeFlowScore,
              routeContinuityPenalty,
              idleVehicleCount
            );
            const preferSpread =
              sameHourLoad > 0 &&
              (!(routeFlowScore >= 85 && routeContinuityPenalty <= 45) || !rotationPrediction.canShare);

            if (sameHourLoad === 0) {
              score -= idleVehicleCount > 1 ? 135 : 84;
            } else if (preferSpread && idleVehicleCount > 0) {
              score += 130 + sameHourLoad * 30;
            } else if (rotationPrediction.canShare) {
              score -= Math.min(routeFlowScore, 120) * 0.26;
              score -= 10;
            }

            score += getNormalRunReturnPenalty(
              cluster.hour,
              Number(item.distance_km || 0),
              1,
              sameHourLoad,
              routeFlowScore,
              routeContinuityPenalty,
              idleVehicleCount
            );
          }

          const homePriorityWeight = getLastTripHomePriorityWeight(
            normalizedClusterArea,
            homeArea,
            isLastRun,
            isDefaultLastHourCluster
          );
          score -= homePriorityWeight;

          if ((isLastRun || isDefaultLastHourCluster) && homePriorityWeight <= 0) {
            score += isLastRun ? 36 : 10;
          }

          if (routeContinuityPenalty >= 170) {
            score += isLastRun || isDefaultLastHourCluster ? 160 : 80;
          } else if (routeContinuityPenalty >= 95) {
            score += isLastRun || isDefaultLastHourCluster ? 72 : 30;
          }

          return { vehicle, score };
        })
        .filter(Boolean)
        .sort((a, b) => a.score - b.score);

      if (!perItemCandidates.length) continue;

      const bestVehicle = perItemCandidates[0].vehicle;

      assignments.push({
        item_id: item.id,
        actual_hour: cluster.hour,
        vehicle_id: bestVehicle.id,
        driver_name: bestVehicle.driver_name || "",
        distance_km: Number(item.distance_km || 0)
      });

      addHourLoad(
        bestVehicle.id,
        cluster.hour,
        1,
        Number(item.distance_km || 0),
        cluster.area
      );
    }
  }

  return assignments;
}

async function applyAutoDispatchAssignments(assignments) {
  const groupedOrderMap = new Map();

  for (const a of assignments) {
    const key = `${a.vehicle_id}_${a.actual_hour}`;
    const nextOrder = (groupedOrderMap.get(key) || 0) + 1;
    groupedOrderMap.set(key, nextOrder);

    const { error } = await supabaseClient
      .from("dispatch_items")
      .update({
        vehicle_id: a.vehicle_id,
        driver_name: a.driver_name,
        stop_order: nextOrder,
        status: "pending"
      })
      .eq("id", a.item_id);

    if (error) {
      throw error;
    }
  }
}

async function assignUnassignedActualsForToday() {
  const selectedVehicles = getSelectedVehiclesForToday();
  if (!selectedVehicles.length) return;

  const unassignedItems = currentActualsCache.filter(item => {
    const status = normalizeStatus(item.status);
    if (status === "cancel") return false;
    if (Number(item.vehicle_id || 0) > 0) return false;
    return true;
  });

  if (!unassignedItems.length) return;

  const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
  let assignments = optimizeAssignments(unassignedItems, selectedVehicles, monthlyMap);
  assignments = optimizeAssignmentsByRouteFlow(assignments, unassignedItems, selectedVehicles);

  if (!assignments.length) {
    assignments = buildFallbackAssignments(unassignedItems, selectedVehicles);
  }

  assignments = optimizeAssignmentsByDistanceBalance(assignments, unassignedItems, selectedVehicles, monthlyMap);
  assignments = applyLastTripDistanceCorrectionToAssignments(assignments, unassignedItems, selectedVehicles, monthlyMap);
  assignments = applyManualLastVehicleToAssignments(assignments, selectedVehicles);

  if (!assignments.length) return;

  await applyAutoDispatchAssignments(assignments);
}

async function runAutoDispatch() {
  const selectedVehicles = getSelectedVehiclesForToday();
  if (!selectedVehicles.length) {
    alert("本日使用する車両を選択してください");
    return;
  }

  const activeItems = getActiveDispatchItemsForAutoAssign();
  if (!activeItems.length) {
    alert("自動配車対象のActualがありません");
    return;
  }

  lastAutoDispatchRunAtMinutes = getCurrentClockMinutes();

  const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
  let assignments = optimizeAssignments(activeItems, selectedVehicles, monthlyMap);
  assignments = optimizeAssignmentsByRouteFlow(assignments, activeItems, selectedVehicles);

  if (!assignments.length) {
    assignments = buildFallbackAssignments(activeItems, selectedVehicles);
  }

  assignments = optimizeAssignmentsByDistanceBalance(assignments, activeItems, selectedVehicles, monthlyMap);
  assignments = applyLastTripDistanceCorrectionToAssignments(assignments, activeItems, selectedVehicles, monthlyMap);
  assignments = applyManualLastVehicleToAssignments(assignments, selectedVehicles);

  if (!assignments.length) {
    alert("自動配車に失敗しました");
    return;
  }

  try {
    await applyAutoDispatchAssignments(assignments);
  } catch (error) {
    console.error(error);
    alert(`配車更新エラー: ${error.message}`);
    return;
  }

  await addHistory(currentDispatchId, null, "auto_dispatch", "自動配車を実行");
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
  renderDailyDispatchResult();
}


function getVehicleRotationForecastSafe(vehicle, orderedRows) {
  try {
    if (typeof calcVehicleRotationForecastGlobal === "function") {
      return calcVehicleRotationForecastGlobal(vehicle, orderedRows);
    }
  } catch (e) {
    console.warn("calcVehicleRotationForecastGlobal fallback:", e);
  }
  try {
    if (typeof calcVehicleRotationForecast === "function") {
      return calcVehicleRotationForecast(vehicle, orderedRows);
    }
  } catch (e) {
    console.warn("calcVehicleRotationForecast fallback:", e);
  }

  const totalDistance = Array.isArray(orderedRows)
    ? orderedRows.reduce((sum, row) => sum + Number(row?.distance_km || 0), 0)
    : 0;

  return {
    routeDistanceKm: totalDistance,
    returnDistanceKm: 0,
    zoneLabel: "-",
    predictedDepartureTime: "-",
    predictedReturnTime: "-",
    predictedReadyTime: "-",
    predictedReturnMinutes: 0,
    extraSharedDelayMinutes: 0,
    returnAfterLabel: "0分後",
    stopCount: Array.isArray(orderedRows) ? orderedRows.length : 0,
    totalKm: totalDistance,
    dailyDistanceKm: totalDistance,
    jobCount: Array.isArray(orderedRows) ? orderedRows.length : 0,
    count: Array.isArray(orderedRows) ? orderedRows.length : 0
  };
}

function buildRotationTimelineHtmlSafe(vehicles, activeItems) {
  try {
    const timeline = (Array.isArray(vehicles) ? vehicles : [])
      .map(vehicle => {
        const rows = (Array.isArray(activeItems) ? activeItems : []).filter(
          item => Number(item.vehicle_id) === Number(vehicle.id)
        );

        if (!rows.length) return null;

        const orderedRows = (typeof moveManualLastItemsToEnd === "function" && typeof sortItemsByNearestRoute === "function")
          ? moveManualLastItemsToEnd(sortItemsByNearestRoute(rows))
          : rows;

        const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
        const summary = getVehicleDailySummary(vehicle, orderedRows);

        return {
          name: vehicle?.driver_name || vehicle?.plate_number || "-",
          lineId: vehicle?.line_id || "",
          returnAfterLabel: forecast?.returnAfterLabel || `${Number(forecast?.predictedReturnMinutes || 0)}分後`,
          nextRunTime: forecast?.predictedReadyTime || "-",
          totalKm: summary.totalKm,
          totalJobs: summary.jobCount
        };
      })
      .filter(Boolean);

    if (!timeline.length) return "";

    return `
      <div class="panel-card" style="margin-bottom:16px;">
        <h3 style="margin-bottom:10px;">車両稼働タイムライン</h3>
        <div style="display:flex; flex-wrap:wrap; gap:8px;">
          ${timeline.map(item => `
            <div class="chip" style="padding:8px 12px;">
              <strong>${escapeHtml(item.name)}</strong>
              ${item.lineId ? `/ LINE ${escapeHtml(item.lineId)}` : ""}
              / 戻り ${escapeHtml(item.returnAfterLabel)}
              / 次便可能 ${escapeHtml(item.nextRunTime)}
              / 累計 ${Number(item.totalKm || 0).toFixed(1)}km
              / ${Number(item.totalJobs || 0)}件
            </div>
          `).join("")}
        </div>
      </div>
    `;
  } catch (e) {
    console.error("buildRotationTimelineHtmlSafe error:", e);
    return "";
  }
}



function formatMinutesAsJa(totalMinutes) {
  const safe = Math.max(0, Math.round(Number(totalMinutes || 0)));
  const hours = Math.floor(safe / 60);
  const minutes = safe % 60;
  if (hours <= 0) return `${minutes}分`;
  if (minutes === 0) return `${hours}時間`;
  return `${hours}時間${minutes}分`;
}

function getVehiclePersistentDailyStats(vehicleId, orderedRows) {
  const numericVehicleId = Number(vehicleId || 0);
  const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
  const reportDate = els.dispatchDate?.value || els.actualDate?.value || todayStr();

  const reportedRow = Array.isArray(currentDailyReportsCache)
    ? currentDailyReportsCache.find(
        row =>
          String(row.report_date || "") === String(reportDate || "") &&
          Number(row.vehicle_id || 0) === numericVehicleId
      )
    : null;

  const actualRows = Array.isArray(currentActualsCache)
    ? currentActualsCache.filter(
        item =>
          Number(item?.vehicle_id || 0) === numericVehicleId &&
          normalizeStatus(item?.status) !== "cancel"
      )
    : [];

  const baseRows = actualRows.length
    ? moveManualLastItemsToEnd(
        sortItemsByNearestRoute(
          [...actualRows].sort((a, b) => {
            const ah = Number(a?.actual_hour ?? a?.plan_hour ?? 0);
            const bh = Number(b?.actual_hour ?? b?.plan_hour ?? 0);
            if (ah !== bh) return ah - bh;
            return Number(a?.stop_order || 0) - Number(b?.stop_order || 0);
          })
        )
      )
    : rows;

  if (reportedRow && Number.isFinite(Number(reportedRow.distance_km))) {
    const reportedDistance = Number(Number(reportedRow.distance_km || 0).toFixed(1));
    const jobCount = actualRows.length || rows.length || 0;
    const driveMinutes = Math.round(
      estimateTravelMinutesByDistanceGlobal(reportedDistance) + jobCount
    );

    return {
      sendKm: reportedDistance,
      returnKm: 0,
      totalKm: reportedDistance,
      driveMinutes,
      jobCount,
      hasFixedReport: true
    };
  }

  if (!baseRows.length) {
    return {
      sendKm: 0,
      returnKm: 0,
      totalKm: 0,
      driveMinutes: 0,
      jobCount: 0,
      hasFixedReport: false
    };
  }

  const sendKm = Number(calculateRouteDistanceGlobal(baseRows) || 0);
  const lastRow = baseRows[baseRows.length - 1] || {};
  const returnKm = Number(lastRow.distance_km || 0);
  const totalKm = Number((sendKm + returnKm).toFixed(1));
  const driveMinutes = Math.round(
    estimateTravelMinutesByDistanceGlobal(sendKm) +
      estimateTravelMinutesByDistanceGlobal(returnKm) +
      baseRows.length
  );

  return {
    sendKm: Number(sendKm.toFixed(1)),
    returnKm: Number(returnKm.toFixed(1)),
    totalKm,
    driveMinutes,
    jobCount: baseRows.length,
    hasFixedReport: false
  };
}

function getVehicleDailySummary(vehicle, orderedRows) {
  const summary = getVehiclePersistentDailyStats(Number(vehicle?.id || 0), orderedRows);
  return {
    sendKm: Number(summary.sendKm || 0),
    returnKm: Number(summary.returnKm || 0),
    totalKm: Number(summary.totalKm || 0),
    driveMinutes: Math.round(Number(summary.driveMinutes || 0)),
    jobCount: Number(summary.jobCount || 0),
    hasFixedReport: Boolean(summary.hasFixedReport)
  };
}

function getVehicleProjectedMonthlyDistance(vehicleId, monthlyMap, orderedRows) {
  const currentMonth = monthlyMap?.get(Number(vehicleId)) || { totalDistance: 0 };
  const todaySummary = getVehicleDailySummary({ id: vehicleId }, orderedRows);
  if (todaySummary.hasFixedReport) {
    return Number(Number(currentMonth.totalDistance || 0).toFixed(1));
  }
  return Number(Number(currentMonth.totalDistance || 0) + Number(todaySummary.totalKm || 0));
}

function getVehicleCardSortValue(vehicle, orderedRows, monthlyMap) {
  const hasRows = Array.isArray(orderedRows) && orderedRows.length > 0;
  const maxHour = hasRows
    ? Math.max(...orderedRows.map(row => Number(row.actual_hour ?? 0)))
    : -1;
  const manualPriority = isManualLastVehicle(vehicle?.id) ? -1000 : 0;
  const projectedMonthly = getVehicleProjectedMonthlyDistance(vehicle?.id, monthlyMap, orderedRows);
  return manualPriority + maxHour * 100000 + projectedMonthly;
}

function buildVehicleLineLabel(vehicle) {
  return vehicle?.line_id ? `LINE ${vehicle.line_id}` : "";
}

function buildDailyDispatchVehicleCards(vehicles, activeItems, monthlyMap) {
  return vehicles.map(vehicle => {
    const rows = activeItems
      .filter(item => Number(item.vehicle_id) === Number(vehicle.id))
      .sort((a, b) => {
        const ah = Number(a.actual_hour ?? 0);
        const bh = Number(b.actual_hour ?? 0);
        if (ah !== bh) return ah - bh;

        const aa = normalizeAreaLabel(a.destination_area || "");
        const ba = normalizeAreaLabel(b.destination_area || "");
        if (aa !== ba) return aa.localeCompare(ba, "ja");

        return Number(a.stop_order || 0) - Number(b.stop_order || 0);
      });

    const orderedRows = moveManualLastItemsToEnd(sortItemsByNearestRoute(rows));
    return { vehicle, rows, orderedRows };
  });
}

function buildLineResultText() {
  const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
  const vehicles = getSelectedVehiclesForToday();
  const activeItems = currentActualsCache.filter(
    x => normalizeStatus(x.status) !== "done" && normalizeStatus(x.status) !== "cancel"
  );

  const cards = buildDailyDispatchVehicleCards(vehicles, activeItems, monthlyMap);
  const lines = [];

  cards.forEach(({ vehicle, orderedRows }) => {
    if (!orderedRows.length) return;

    const summary = getVehicleDailySummary(vehicle, orderedRows);
    const uniqueAreas = [...new Set(
      orderedRows
        .map(row => normalizeAreaLabel(row.destination_area || row.casts?.area || "無し"))
        .filter(Boolean)
    )];
    const areaLabel = uniqueAreas.length
      ? uniqueAreas.join("・")
      : normalizeAreaLabel(vehicle.vehicle_area || "無し");
    const lineId = String(vehicle?.line_id || "").trim() || "LINE未設定";
    const driverName = vehicle?.driver_name || vehicle?.plate_number || "-";
    const estimatedRoundTripKm = `${Number(summary.totalKm || 0).toFixed(1)}km`;
    const estimatedRoundTripTime = formatMinutesAsJa(summary.driveMinutes);

    lines.push(`${lineId} ${driverName} ${areaLabel} ${estimatedRoundTripKm} ${estimatedRoundTripTime}`);

    orderedRows.forEach(row => {
      const mapUrl = buildDispatchItemMapUrl(row) || buildMapUrlFromAddressOrLatLng(
        row?.destination_address || row?.casts?.address || "",
        row?.casts?.latitude,
        row?.casts?.longitude
      );
      const castName = row?.casts?.name || "-";
      lines.push(`${castName}：${mapUrl || "地図URLなし"}`);
    });

    lines.push("");
  });

  return lines.join("\n").trim();
}


function optimizeAssignmentsByDistanceBalance(assignments, items, vehicles, monthlyMap) {
  const working = assignments.map(a => ({ ...a }));
  if (!working.length || !vehicles.length) return working;

  const itemMap = new Map((items || []).map(item => [Number(item.id), item]));
  const hourLoads = new Map();
  const assignedDistance = new Map();

  const getDistanceForAssignment = assignment => {
    const item = itemMap.get(Number(assignment.item_id));
    return Number(item?.distance_km ?? assignment?.distance_km ?? 0);
  };

  const rebuild = () => {
    hourLoads.clear();
    assignedDistance.clear();

    working.forEach(a => {
      const hourKey = `${Number(a.vehicle_id)}__${Number(a.actual_hour ?? 0)}`;
      hourLoads.set(hourKey, Number(hourLoads.get(hourKey) || 0) + 1);
      assignedDistance.set(
        Number(a.vehicle_id),
        Number(assignedDistance.get(Number(a.vehicle_id)) || 0) + getDistanceForAssignment(a)
      );
    });
  };

  const getProjectedDistance = vehicleId => {
    const monthly = monthlyMap?.get(Number(vehicleId)) || { totalDistance: 0 };
    return Number(monthly.totalDistance || 0) + Number(assignedDistance.get(Number(vehicleId)) || 0);
  };

  rebuild();

  for (const assignment of working) {
    const currentVehicleId = Number(assignment.vehicle_id);
    const currentProjected = getProjectedDistance(currentVehicleId);
    const item = itemMap.get(Number(assignment.item_id));
    if (!item) continue;

    const area = normalizeAreaLabel(item.destination_area || item.cluster_area || "無し");
    const dist = Number(item.distance_km || assignment.distance_km || 0);
    const currentVehicle = vehicles.find(v => Number(v.id) === currentVehicleId);
    const currentCompat =
      getStrictHomeCompatibilityScore(area, currentVehicle?.home_area || "") * 1.4 +
      Math.max(0, getDirectionAffinityScore(area, currentVehicle?.home_area || "")) * 0.7 +
      getAreaAffinityScore(area, currentVehicle?.home_area || "");

    let bestMove = null;

    for (const vehicle of vehicles) {
      const vehicleId = Number(vehicle.id);
      if (vehicleId === currentVehicleId) continue;

      const hourKey = `${vehicleId}__${Number(assignment.actual_hour ?? 0)}`;
      const seatCapacity = Number(vehicle.seat_capacity || 4);
      const hourLoad = Number(hourLoads.get(hourKey) || 0);
      if (hourLoad >= seatCapacity) continue;

      const candidateProjected = getProjectedDistance(vehicleId);
      const projectedGapImprove = currentProjected - candidateProjected;
      if (projectedGapImprove < 10) continue;

      const compat =
        getStrictHomeCompatibilityScore(area, vehicle.home_area || "") * 1.4 +
        Math.max(0, getDirectionAffinityScore(area, vehicle.home_area || "")) * 0.7 +
        getAreaAffinityScore(area, vehicle.home_area || "");
      if (isHardReverseForHome(area, vehicle.home_area || "")) continue;

      const score = projectedGapImprove * 2.2 + (compat - currentCompat) * 1.1 - dist * 0.12;
      if (!bestMove || score > bestMove.score) {
        bestMove = { vehicle, score };
      }
    }

    if (bestMove && bestMove.score >= 28) {
      assignment.vehicle_id = bestMove.vehicle.id;
      assignment.driver_name = bestMove.vehicle.driver_name || "";
      rebuild();
    }
  }

  return working;
}

function applyLastTripDistanceCorrectionToAssignments(assignments, items, vehicles, monthlyMap) {
  const working = assignments.map(a => ({ ...a }));
  if (!working.length || !vehicles.length) return working;

  const itemMap = new Map((items || []).map(item => [Number(item.id), item]));
  const dateStr = els.actualDate?.value || todayStr();
  const defaultLastHour = getDefaultLastHour(dateStr);
  const targetHour = working.some(a => Number(a.actual_hour ?? 0) === Number(defaultLastHour))
    ? Number(defaultLastHour)
    : Math.max(...working.map(a => Number(a.actual_hour ?? 0)));

  const targetRows = working.filter(a => Number(a.actual_hour ?? 0) === Number(targetHour));
  if (!targetRows.length) return working;

  const manualVehicleId = getManualLastVehicleId();
  const projectedDistanceByVehicle = new Map();

  working.forEach(a => {
    const item = itemMap.get(Number(a.item_id));
    projectedDistanceByVehicle.set(
      Number(a.vehicle_id),
      Number(projectedDistanceByVehicle.get(Number(a.vehicle_id)) || 0) +
        Number(item?.distance_km ?? a.distance_km ?? 0)
    );
  });

  const evaluate = (vehicle, item) => {
    const area = normalizeAreaLabel(item?.destination_area || item?.cluster_area || "無し");
    const itemDistance = Number(item?.distance_km || 0);
    const strict = getStrictHomeCompatibilityScore(area, vehicle?.home_area || "");
    const direction = Math.max(0, getDirectionAffinityScore(area, vehicle?.home_area || ""));
    const affinity = getAreaAffinityScore(area, vehicle?.home_area || "");
    const vehicleMatch = getVehicleAreaMatchScore(vehicle, area);
    const projectedMonthly =
      Number(monthlyMap?.get(Number(vehicle?.id))?.totalDistance || 0) +
      Number(projectedDistanceByVehicle.get(Number(vehicle?.id)) || 0);
    const hardReverse = isHardReverseForHome(area, vehicle?.home_area || "");

    let score =
      strict * 8 +
      direction * 4 +
      affinity * 3 +
      vehicleMatch * 0.6 +
      itemDistance * (strict >= 78 ? 1.4 : strict >= 52 ? 0.8 : 0.2);

    score -= projectedMonthly * 0.12;
    if (hardReverse) score -= 9999;
    if (Number(vehicle?.id) === Number(manualVehicleId) && strict >= 52 && !hardReverse) score += 110;
    return score;
  };

  for (const target of targetRows) {
    const item = itemMap.get(Number(target.item_id));
    if (!item) continue;

    const currentVehicle = vehicles.find(v => Number(v.id) === Number(target.vehicle_id));
    const currentScore = evaluate(currentVehicle, item);
    let best = { vehicle: currentVehicle, score: currentScore };

    for (const vehicle of vehicles) {
      const seatCapacity = Number(vehicle.seat_capacity || 4);
      const load = working.filter(
        a =>
          Number(a.vehicle_id) === Number(vehicle.id) &&
          Number(a.actual_hour ?? 0) === Number(targetHour) &&
          Number(a.item_id) !== Number(target.item_id)
      ).length;
      if (load >= seatCapacity) continue;

      const score = evaluate(vehicle, item);
      if (score > best.score) best = { vehicle, score };
    }

    if (
      best.vehicle &&
      Number(best.vehicle.id) !== Number(target.vehicle_id) &&
      best.score >= currentScore + 24
    ) {
      target.vehicle_id = best.vehicle.id;
      target.driver_name = best.vehicle.driver_name || "";
      target.manual_last_vehicle = Number(best.vehicle.id) === Number(manualVehicleId);
    }
  }

  return working;
}

function renderDailyDispatchResult() {
  if (!els.dailyDispatchResult) return;

  const vehicles = getSelectedVehiclesForToday();
  if (!vehicles.length) {
    els.dailyDispatchResult.innerHTML = `<div class="muted">使用車両が未選択です</div>`;
    return;
  }

  const activeItems = currentActualsCache.filter(
    x => normalizeStatus(x.status) !== "done" && normalizeStatus(x.status) !== "cancel"
  );

  try {
    const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
    const timelineHtml = buildRotationTimelineHtmlSafe(vehicles, activeItems);
    const cards = buildDailyDispatchVehicleCards(vehicles, activeItems, monthlyMap);

    const cardsHtml = cards
      .map(({ vehicle, rows, orderedRows }) => {
        const summary = getVehicleDailySummary(vehicle, orderedRows);
        const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
        const lineLabel = buildVehicleLineLabel(vehicle);
        const projectedMonthly = getVehicleProjectedMonthlyDistance(vehicle.id, monthlyMap, orderedRows);

        const body = orderedRows.length
          ? orderedRows
              .map(
                (row, index) => `
                  <div class="dispatch-row">
                    <div class="dispatch-left">
                      <span class="badge-time">${escapeHtml(getHourLabel(row.actual_hour))}</span>
                      <span class="badge-order">順番 ${index + 1}</span>
                      <span class="dispatch-name">${buildMapLinkHtml({
                        name: row.casts?.name,
                        address: row.destination_address || row.casts?.address,
                        lat: row.casts?.latitude,
                        lng: row.casts?.longitude,
                        className: "dispatch-name-link"
                      })}</span>
                      <span class="dispatch-area">${escapeHtml(normalizeAreaLabel(row.destination_area || "-"))}</span>
                      ${isManualLastTripItem(row) ? `<span class="badge-status assigned">ラスト便</span>` : ""}
                    </div>
                    <div class="dispatch-right">
                      <div class="dispatch-distance">${Number(row.distance_km || 0).toFixed(1)}km</div>
                      <select class="dispatch-vehicle-select" data-item-id="${row.id}">
                        ${vehicles
                          .map(
                            v => `
                              <option value="${v.id}" ${Number(v.id) === Number(vehicle.id) ? "selected" : ""}>
                                 ${escapeHtml(v.driver_name || v.plate_number || "-")}
                              </option>
                            `
                          )
                          .join("")}
                      </select>
                    </div>
                  </div>
                `
              )
              .join("")
          : `<div class="empty-vehicle-text">送りなし</div>`;

        return `
          <div class="vehicle-result-card">
            <div class="vehicle-result-head">
              <div class="vehicle-result-title">
                <h4>
                  ${escapeHtml(vehicle.driver_name || vehicle.plate_number || "-")}
                  ${isManualLastVehicle(vehicle.id) ? `<span class="badge-status assigned" style="margin-left:8px;">手動ラスト便車両</span>` : ""}
                </h4>
                <div class="vehicle-result-meta">
                  ${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))}
                  / 帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))}
                  / 定員${vehicle.seat_capacity ?? "-"}
                  ${lineLabel ? `/ ${escapeHtml(lineLabel)}` : ""}
                </div>
              </div>
              <div class="vehicle-result-badges">
                <span class="metric-badge">人数 ${rows.length}</span>
                <span class="metric-badge">累計距離 ${summary.totalKm.toFixed(1)}km</span>
                <span class="metric-badge">累計時間 ${escapeHtml(formatMinutesAsJa(summary.driveMinutes))}</span>
                <span class="metric-badge">累計件数 ${summary.jobCount}件</span>
              </div>
            </div>
            <div class="vehicle-result-body">${body}</div>
            ${orderedRows.length ? `
              <div class="dispatch-meta" style="margin-top:10px; font-size:12px; color:#9aa3b2; line-height:1.8;">
                戻り ${escapeHtml(forecast.returnAfterLabel)}
                / 次便可能 ${escapeHtml(forecast.predictedReadyTime)}
                / 累計距離 ${summary.totalKm.toFixed(1)}km
                / 累計時間 ${escapeHtml(formatMinutesAsJa(summary.driveMinutes))}
                / 累計件数 ${summary.jobCount}件
                / 月間見込 ${projectedMonthly.toFixed(1)}km
                ${forecast.extraSharedDelayMinutes > 0 ? `/ 同乗追加遅延 ${forecast.extraSharedDelayMinutes}分` : ""}
              </div>
            ` : `
              <div class="dispatch-meta" style="margin-top:10px; font-size:12px; color:#9aa3b2; line-height:1.8;">
                累計距離 ${summary.totalKm.toFixed(1)}km
                / 累計時間 ${escapeHtml(formatMinutesAsJa(summary.driveMinutes))}
                / 累計件数 ${summary.jobCount}件
                / 月間見込 ${projectedMonthly.toFixed(1)}km
              </div>
            `}
          </div>
        `;
      })
      .join("");

    els.dailyDispatchResult.innerHTML = timelineHtml + cardsHtml;

    els.dailyDispatchResult.querySelectorAll(".dispatch-vehicle-select").forEach(select => {
      select.addEventListener("change", async () => {
        const itemId = Number(select.dataset.itemId);
        const vehicleId = Number(select.value);
        const vehicle = allVehiclesCache.find(v => Number(v.id) === vehicleId);

        const { error } = await supabaseClient
          .from("dispatch_items")
          .update({
            vehicle_id: vehicleId,
            driver_name: vehicle?.driver_name || null
          })
          .eq("id", itemId);

        if (error) {
          alert(error.message);
          return;
        }

        await addHistory(currentDispatchId, itemId, "change_vehicle", "車両を変更");
        await loadActualsByDate(els.actualDate?.value || todayStr());
        renderDailyDispatchResult();
      });
    });
  } catch (error) {
    console.error("renderDailyDispatchResult error:", error);
    els.dailyDispatchResult.innerHTML = `<div class="muted">配車結果の表示でエラーが発生しました</div>`;
  }
}

function buildCopyResultText() {
  return buildLineResultText();
}

async function copyDispatchResult() {
  const text = buildCopyResultText();
  try {
    await navigator.clipboard.writeText(text);
    alert("結果をコピーしました");
  } catch (error) {
    console.error(error);
    alert("コピーに失敗しました");
  }
}

async function clearAllActuals() {
  if (!window.confirm("この日のActualを全消去しますか？")) return;
  if (!currentDispatchId) return;

  const { error } = await supabaseClient
    .from("dispatch_items")
    .delete()
    .eq("dispatch_id", currentDispatchId);

  if (error) {
    alert(error.message);
    return;
  }

  await supabaseClient
    .from("dispatch_plans")
    .update({ status: "planned" })
    .eq("plan_date", els.actualDate?.value || todayStr())
    .eq("status", "assigned");

  await addHistory(currentDispatchId, null, "clear_actual", "Actualを全消去");
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
}

async function loadDailyReports(dateStr) {
  const monthKey = getMonthKey(dateStr);
  const monthStart = `${monthKey}-01`;
  const start = new Date(monthStart);
  const next = new Date(start.getFullYear(), start.getMonth() + 1, 1);
  const nextStr = `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, "0")}-01`;

  const { data, error } = await supabaseClient
    .from("vehicle_daily_reports")
    .select(`
      *,
      vehicles (
        id,
        plate_number,
        driver_name
      )
    `)
    .gte("report_date", monthStart)
    .lt("report_date", nextStr)
    .order("report_date", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  currentDailyReportsCache = data || [];
  renderHomeMonthlyVehicleList();
  renderVehiclesTable();
}

async function confirmDailyToMonthly() {
  const doneRows = currentActualsCache.filter(x => normalizeStatus(x.status) === "done");
  if (!doneRows.length) {
    alert("完了したActualがありません");
    return;
  }

  const grouped = new Map();

  doneRows.forEach(row => {
    const vehicleId = Number(row.vehicle_id);
    if (!vehicleId) return;

    const prev = grouped.get(vehicleId) || {
      distance: 0,
      driver_name: row.driver_name || ""
    };
    prev.distance += Number(row.distance_km || 0);
    if (!prev.driver_name && row.driver_name) prev.driver_name = row.driver_name;
    grouped.set(vehicleId, prev);
  });

  const reportDate = els.dispatchDate?.value || todayStr();

  for (const [vehicleId, info] of grouped.entries()) {
    const { data: existing, error: selectError } = await supabaseClient
      .from("vehicle_daily_reports")
      .select("id, distance_km")
      .eq("report_date", reportDate)
      .eq("vehicle_id", vehicleId)
      .maybeSingle();

    if (selectError) {
      console.error(selectError);
      continue;
    }

    if (existing) {
      const { error: updateError } = await supabaseClient
        .from("vehicle_daily_reports")
        .update({
          driver_name: info.driver_name || null,
          distance_km: Number(info.distance.toFixed(1)),
          note: "当日運用の完了データから更新",
          created_by: currentUser.id
        })
        .eq("id", existing.id);

      if (updateError) console.error(updateError);
    } else {
      const { error: insertError } = await supabaseClient
        .from("vehicle_daily_reports")
        .insert({
          report_date: reportDate,
          vehicle_id: vehicleId,
          driver_name: info.driver_name || null,
          distance_km: Number(info.distance.toFixed(1)),
          note: "当日運用の完了データから自動反映",
          created_by: currentUser.id
        });

      if (insertError) console.error(insertError);
    }
  }

  await addHistory(currentDispatchId, null, "confirm_daily", "完了データを月間へ反映");
  await loadDailyReports(reportDate);
  await loadVehicles();
  await loadHomeAndAll();
}

async function resetMonthlySummary() {
  if (!window.confirm("今月の走行記録を削除しますか？")) return;

  const dateStr = els.dispatchDate?.value || todayStr();
  const monthKey = getMonthKey(dateStr);
  const monthStart = `${monthKey}-01`;
  const start = new Date(monthStart);
  const next = new Date(start.getFullYear(), start.getMonth() + 1, 1);
  const nextStr = `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, "0")}-01`;

  const { error } = await supabaseClient
    .from("vehicle_daily_reports")
    .delete()
    .gte("report_date", monthStart)
    .lt("report_date", nextStr);

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "reset_monthly_reports", `${monthKey} の月間距離/出勤日数をリセット`);
  await loadDailyReports(dateStr);
  renderManualLastVehicleInfo();
}

async function addHistory(dispatchId, itemId, action, message) {
  const { error } = await supabaseClient
    .from("dispatch_history")
    .insert({
      dispatch_id: dispatchId,
      item_id: itemId,
      action,
      message,
      acted_by: currentUser.id
    });

  if (error) console.error(error);
}

async function loadHistory() {
  const { data, error } = await supabaseClient
    .from("dispatch_history")
    .select("*")
    .order("id", { ascending: false })
    .limit(50);

  if (error) {
    console.error(error);
    return;
  }

  if (!els.historyList) return;
  els.historyList.innerHTML = "";

  if (!data?.length) {
    els.historyList.innerHTML = `<div class="muted">履歴はありません</div>`;
    return;
  }

  data.forEach(row => {
    const div = document.createElement("div");
    div.className = "history-item";
    div.innerHTML = `
      <h4>${escapeHtml(row.action)}</h4>
      <p>${escapeHtml(row.message || "")}</p>
      <p class="muted">${escapeHtml(formatDateTimeJa(row.created_at))}</p>
    `;
    els.historyList.appendChild(div);
  });
}

async function fetchAllTableRows(tableName, orderColumn = "id") {
  const pageSize = 1000;
  let from = 0;
  let allRows = [];

  while (true) {
    const { data, error } = await supabaseClient
      .from(tableName)
      .select("*")
      .order(orderColumn, { ascending: true })
      .range(from, from + pageSize - 1);

    if (error) throw error;

    const rows = data || [];
    allRows = allRows.concat(rows);

    if (rows.length < pageSize) break;
    from += pageSize;
  }

  return allRows;
}

function stripMetaForInsert(row, extraRemoveKeys = []) {
  const clone = { ...row };
  const removeKeys = [
    "id",
    "created_at",
    "updated_at",
    ...extraRemoveKeys
  ];

  removeKeys.forEach(key => {
    if (key in clone) delete clone[key];
  });

  return clone;
}

async function exportAllData() {
  try {
    const [
      casts,
      vehicles,
      dispatches,
      dispatchPlans,
      dispatchItems,
      vehicleDailyReports,
      dispatchHistory
    ] = await Promise.all([
      fetchAllTableRows("casts"),
      fetchAllTableRows("vehicles"),
      fetchAllTableRows("dispatches"),
      fetchAllTableRows("dispatch_plans"),
      fetchAllTableRows("dispatch_items"),
      fetchAllTableRows("vehicle_daily_reports"),
      fetchAllTableRows("dispatch_history")
    ]);

    const payload = {
      app: "THEMIS AI Dispatch",
      version: 2,
      exported_at: new Date().toISOString(),
      origin: {
        label: ORIGIN_LABEL,
        lat: ORIGIN_LAT,
        lng: ORIGIN_LNG
      },
      data: {
        casts,
        vehicles,
        dispatches,
        dispatch_plans: dispatchPlans,
        dispatch_items: dispatchItems,
        vehicle_daily_reports: vehicleDailyReports,
        dispatch_history: dispatchHistory
      }
    };

    downloadTextFile(
      `themis_full_backup_${todayStr()}.json`,
      JSON.stringify(payload, null, 2),
      "application/json;charset=utf-8"
    );

    await addHistory(null, null, "export_all", "全体バックアップを出力");
  } catch (error) {
    console.error("exportAllData error:", error);
    alert("全体エクスポートに失敗しました: " + error.message);
  }
}

function triggerImportAll() {
  els.importAllFileInput?.click();
}

async function importAllDataFromFile() {
  const file = els.importAllFileInput?.files?.[0];
  if (!file) {
    alert("JSONファイルを選択してください");
    return;
  }

  try {
    const text = await file.text();
    const json = JSON.parse(text);

    if (!json?.data) {
      alert("バックアップJSONの形式が正しくありません");
      return;
    }

    const backup = json.data;

    const proceed = window.confirm(
      "現在のデータを全消去して、バックアップJSONから復元しますか？"
    );
    if (!proceed) {
      els.importAllFileInput.value = "";
      return;
    }

    const proceed2 = window.confirm(
      "本当に実行しますか？現在のデータは上書きではなく、一度削除してから復元します。"
    );
    if (!proceed2) {
      els.importAllFileInput.value = "";
      return;
    }

    const deleteTargets = [
      "dispatch_items",
      "dispatch_plans",
      "vehicle_daily_reports",
      "dispatch_history",
      "dispatches",
      "casts",
      "vehicles"
    ];

    for (const table of deleteTargets) {
      const { error } = await supabaseClient
        .from(table)
        .delete()
        .neq("id", 0);

      if (error) {
        console.error(`${table} delete error:`, error);
        alert(`${table} の削除に失敗しました: ${error.message}`);
        return;
      }
    }

    const castIdMap = new Map();
    const vehicleIdMap = new Map();
    const dispatchIdMap = new Map();
    const planIdMap = new Map();
    const itemIdMap = new Map();

    // 1. casts
    for (const oldRow of backup.casts || []) {
      const row = stripMetaForInsert(oldRow);
      if (!row.created_by) row.created_by = currentUser.id;

      const { data, error } = await supabaseClient
        .from("casts")
        .insert(row)
        .select()
        .single();

      if (error) throw error;
      castIdMap.set(Number(oldRow.id), Number(data.id));
    }

    // 2. vehicles
    for (const oldRow of backup.vehicles || []) {
      const row = stripMetaForInsert(oldRow);
      if (!row.created_by) row.created_by = currentUser.id;

      const { data, error } = await supabaseClient
        .from("vehicles")
        .insert(row)
        .select()
        .single();

      if (error) throw error;
      vehicleIdMap.set(Number(oldRow.id), Number(data.id));
    }

    // 3. dispatches
    for (const oldRow of backup.dispatches || []) {
      const row = stripMetaForInsert(oldRow);
      if (!row.created_by) row.created_by = currentUser.id;

      const { data, error } = await supabaseClient
        .from("dispatches")
        .insert(row)
        .select()
        .single();

      if (error) throw error;
      dispatchIdMap.set(Number(oldRow.id), Number(data.id));
    }

    // 4. dispatch_plans
    for (const oldRow of backup.dispatch_plans || []) {
      const row = stripMetaForInsert(oldRow);
      row.cast_id = castIdMap.get(Number(oldRow.cast_id)) || null;
      if (!row.created_by) row.created_by = currentUser.id;

      const { data, error } = await supabaseClient
        .from("dispatch_plans")
        .insert(row)
        .select()
        .single();

      if (error) throw error;
      planIdMap.set(Number(oldRow.id), Number(data.id));
    }

    // 5. dispatch_items
    for (const oldRow of backup.dispatch_items || []) {
      const row = stripMetaForInsert(oldRow);
      row.dispatch_id = dispatchIdMap.get(Number(oldRow.dispatch_id)) || null;
      row.cast_id = castIdMap.get(Number(oldRow.cast_id)) || null;
      row.vehicle_id = oldRow.vehicle_id != null
        ? (vehicleIdMap.get(Number(oldRow.vehicle_id)) || null)
        : null;

      const { data, error } = await supabaseClient
        .from("dispatch_items")
        .insert(row)
        .select()
        .single();

      if (error) throw error;
      itemIdMap.set(Number(oldRow.id), Number(data.id));
    }

    // 6. vehicle_daily_reports
    for (const oldRow of backup.vehicle_daily_reports || []) {
      const row = stripMetaForInsert(oldRow);
      row.vehicle_id = vehicleIdMap.get(Number(oldRow.vehicle_id)) || null;
      if (!row.created_by) row.created_by = currentUser.id;

      const { error } = await supabaseClient
        .from("vehicle_daily_reports")
        .insert(row);

      if (error) throw error;
    }

    // 7. dispatch_history
    for (const oldRow of backup.dispatch_history || []) {
      const row = stripMetaForInsert(oldRow);
      row.dispatch_id =
        oldRow.dispatch_id != null
          ? (dispatchIdMap.get(Number(oldRow.dispatch_id)) || null)
          : null;
      row.item_id =
        oldRow.item_id != null
          ? (itemIdMap.get(Number(oldRow.item_id)) || null)
          : null;
      if (!row.acted_by) row.acted_by = currentUser.id;

      const { error } = await supabaseClient
        .from("dispatch_history")
        .insert(row);

      if (error) throw error;
    }

    els.importAllFileInput.value = "";

    await addHistory(null, null, "import_all", "全体バックアップから復元");
    alert("全体インポートが完了しました");
    await loadHomeAndAll();
    renderManualLastVehicleInfo();
  } catch (error) {
    console.error("importAllDataFromFile error:", error);
    alert("全体インポートに失敗しました: " + error.message);
  }
}

async function resetAllCastsDanger() {
  if (!window.confirm("本当にキャスト全データを消去しますか？この操作は元に戻せません。")) return;

  const { error } = await supabaseClient
    .from("casts")
    .update({ is_active: false })
    .eq("is_active", true);

  if (error) {
    console.error(error);
    alert("キャスト全データ消去に失敗しました: " + error.message);
    return;
  }

  await addHistory(null, null, "reset_casts", "キャスト全データを消去");
  await loadCasts();
  alert("キャスト全データを消去しました");
}

async function resetAllVehiclesDanger() {
  if (!window.confirm("本当に車両全データを消去しますか？この操作は元に戻せません。")) return;

  const { error } = await supabaseClient
    .from("vehicles")
    .update({ is_active: false })
    .eq("is_active", true);

  if (error) {
    console.error(error);
    alert("車両全データ消去に失敗しました: " + error.message);
    return;
  }

  await addHistory(null, null, "reset_vehicles", "車両全データを消去");
  await loadVehicles();
  alert("車両全データを消去しました");
}

async function resetAllDataDanger() {
  if (!window.confirm("本当に全データを消去しますか？この操作は元に戻せません。")) return;

  try {
    const deleteTargets = [
      "dispatch_items",
      "dispatch_plans",
      "vehicle_daily_reports",
      "casts",
      "vehicles"
    ];

    for (const table of deleteTargets) {
      const { error } = await supabaseClient.from(table).delete().neq("id", 0);
      if (error) {
        console.error(`${table} delete error:`, error);
        alert(`${table} の削除でエラー: ${error.message}`);
        return;
      }
    }

    const { error: historyDeleteError } = await supabaseClient
      .from("dispatch_history")
      .delete()
      .neq("id", 0);

    if (historyDeleteError) {
      console.error("dispatch_history delete error:", historyDeleteError);
      alert(`dispatch_history の削除でエラー: ${historyDeleteError.message}`);
      return;
    }

    currentDispatchId = null;
    activeVehicleIdsForToday = new Set();

    resetCastForm();
    resetVehicleForm();
    resetPlanForm();
    resetActualForm();

    alert("全データを削除しました");
    await loadHomeAndAll();
    renderManualLastVehicleInfo();
  } catch (err) {
    console.error("resetAllDataDanger error:", err);
    alert("全消去中にエラーが発生しました");
  }
}

async function loadHomeAndAll() {
  const dateStr = els.dispatchDate?.value || todayStr();

  if (els.dispatchDate) els.dispatchDate.value = dateStr;
  if (els.planDate) els.planDate.value = dateStr;
  if (els.actualDate) els.actualDate.value = dateStr;
  if (els.mileageReportStartDate && !els.mileageReportStartDate.value) els.mileageReportStartDate.value = getMonthStartStr(dateStr);
  if (els.mileageReportEndDate && !els.mileageReportEndDate.value) els.mileageReportEndDate.value = dateStr;

  await loadCasts();
  await loadVehicles();
  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
  await loadHistory();

  renderDailyVehicleChecklist();
  renderDailyMileageInputs();
  renderDailyDispatchResult();
  renderHomeSummary();
  renderHomeMonthlyVehicleList();
}

function renderDailyMileageInputs() {
  if (!els.dailyMileageInputs) return;

  const defaultDate = els.dispatchDate?.value || todayStr();
  const selectedVehicles = getSelectedVehiclesForToday();

  els.dailyMileageInputs.innerHTML = "";

  if (!selectedVehicles.length) {
    els.dailyMileageInputs.innerHTML = `<div class="muted">出勤車両を選択すると入力欄が表示されます</div>`;
    return;
  }

  selectedVehicles.forEach(vehicle => {
    const existing = currentDailyReportsCache.find(
      r =>
        Number(r.vehicle_id) === Number(vehicle.id)
    );

    const row = document.createElement("div");
    row.className = "daily-mileage-row";
    row.innerHTML = `
      <div>
        <div class="daily-mileage-label">${escapeHtml(vehicle.plate_number || "-")}</div>
        <div class="daily-mileage-sub">
          ${escapeHtml(vehicle.driver_name || "-")} / 帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))}
        </div>
      </div>

      <div class="field">
        <label>入力日</label>
        <input
          type="date"
          class="daily-mileage-date-input"
          data-vehicle-id="${vehicle.id}"
          value="${existing?.report_date || defaultDate}"
        />
      </div>

      <div class="field">
        <label>実績走行距離(km)</label>
        <input
          type="number"
          step="0.1"
          min="0"
          class="daily-mileage-input"
          data-vehicle-id="${vehicle.id}"
          value="${existing?.distance_km ?? ""}"
          placeholder="例：72.5"
        />
      </div>

      <div class="field">
        <label>メモ</label>
        <input
          type="text"
          class="daily-mileage-note-input"
          data-vehicle-id="${vehicle.id}"
          value="${escapeHtml(existing?.note || "")}"
          placeholder="任意"
        />
      </div>
    `;
    els.dailyMileageInputs.appendChild(row);
  });
}

async function saveDailyMileageReports() {
  const selectedVehicles = getSelectedVehiclesForToday();

  if (!selectedVehicles.length) {
    alert("先に出勤車両を選択してください");
    return;
  }

  const mileageInputs = [...document.querySelectorAll(".daily-mileage-input")];
  const noteInputs = [...document.querySelectorAll(".daily-mileage-note-input")];
  const dateInputs = [...document.querySelectorAll(".daily-mileage-date-input")];

  for (const vehicle of selectedVehicles) {
    const mileageInput = mileageInputs.find(
      input => Number(input.dataset.vehicleId) === Number(vehicle.id)
    );
    const noteInput = noteInputs.find(
      input => Number(input.dataset.vehicleId) === Number(vehicle.id)
    );
    const dateInput = dateInputs.find(
      input => Number(input.dataset.vehicleId) === Number(vehicle.id)
    );

    const reportDate = dateInput?.value || (els.dispatchDate?.value || todayStr());
    const distanceKm = toNullableNumber(mileageInput?.value);
    const note = noteInput?.value.trim() || "日次報告入力";

    if (!reportDate) continue;
    if (distanceKm === null) continue;

    const { data: existing, error: selectError } = await supabaseClient
      .from("vehicle_daily_reports")
      .select("id")
      .eq("report_date", reportDate)
      .eq("vehicle_id", vehicle.id)
      .maybeSingle();

    if (selectError) {
      console.error(selectError);
      alert("日次報告の確認でエラーが発生しました");
      return;
    }

    if (existing) {
      const { error: updateError } = await supabaseClient
        .from("vehicle_daily_reports")
        .update({
          driver_name: vehicle.driver_name || null,
          distance_km: distanceKm,
          note,
          created_by: currentUser.id
        })
        .eq("id", existing.id);

      if (updateError) {
        console.error(updateError);
        alert("日次報告の更新に失敗しました: " + updateError.message);
        return;
      }
    } else {
      const { error: insertError } = await supabaseClient
        .from("vehicle_daily_reports")
        .insert({
          report_date: reportDate,
          vehicle_id: vehicle.id,
          driver_name: vehicle.driver_name || null,
          distance_km: distanceKm,
          note,
          created_by: currentUser.id
        });

      if (insertError) {
        console.error(insertError);
        alert("日次報告の保存に失敗しました: " + insertError.message);
        return;
      }
    }
  }

  await addHistory(null, null, "save_daily_mileage", `日次走行距離を保存`);
  alert("日次走行距離を保存しました");

  await loadDailyReports(els.dispatchDate?.value || todayStr());
  renderDailyMileageInputs();
  renderHomeMonthlyVehicleList();
  renderVehiclesTable();
}

async function syncDateAndReloadFromDispatchDate() {
  const dateStr = els.dispatchDate?.value || todayStr();
  if (els.planDate) els.planDate.value = dateStr;
  if (els.actualDate) els.actualDate.value = dateStr;

  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
  renderManualLastVehicleInfo();
  renderDailyDispatchResult();
}

async function syncDateAndReloadFromPlanDate() {
  const dateStr = els.planDate?.value || todayStr();
  if (els.dispatchDate) els.dispatchDate.value = dateStr;
  if (els.actualDate) els.actualDate.value = dateStr;

  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
  renderManualLastVehicleInfo();
}

async function syncDateAndReloadFromActualDate() {
  const dateStr = els.actualDate?.value || todayStr();
  if (els.dispatchDate) els.dispatchDate.value = dateStr;
  if (els.planDate) els.planDate.value = dateStr;

  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
}

function bindPlanAndActualFormEvents() {
  if (els.planCastSelect) {
    els.planCastSelect.addEventListener("change", () => syncPlanFieldsFromCastInput(true));
    els.planCastSelect.addEventListener("input", () => syncPlanFieldsFromCastInput(false));
  }
  if (els.castSelect) {
    els.castSelect.addEventListener("change", () => syncActualFieldsFromCastInput(true));
    els.castSelect.addEventListener("input", () => syncActualFieldsFromCastInput(false));
  }
  if (els.planSelect) els.planSelect.addEventListener("change", fillActualFormFromSelectedPlan);
  if (els.cancelPlanEditBtn) els.cancelPlanEditBtn.addEventListener("click", resetPlanForm);
  if (els.cancelActualEditBtn) els.cancelActualEditBtn.addEventListener("click", resetActualForm);
  if (els.addSelectedPlanBtn) els.addSelectedPlanBtn.addEventListener("click", addPlanToActual);
}

function bindDispatchEvents() {
  if (els.optimizeBtn) els.optimizeBtn.addEventListener("click", runAutoDispatch);
}

function bindPostDispatchEvents() {
  if (els.copyResultBtn) els.copyResultBtn.addEventListener("click", copyDispatchResult);
  if (els.confirmDailyBtn) els.confirmDailyBtn.addEventListener("click", confirmDailyToMonthly);
  if (els.clearActualBtn) els.clearActualBtn.addEventListener("click", clearAllActuals);
}

function setupEvents() {
  els.logoutBtn?.addEventListener("click", logout);
  els.exportAllBtn?.addEventListener("click", exportAllData);
  els.importAllBtn?.addEventListener("click", triggerImportAll);
  els.importAllFileInput?.addEventListener("change", importAllDataFromFile);
  els.exportCsvBtnHeader?.addEventListener("click", exportCastsCsv);
  els.openManualBtn?.addEventListener("click", openManual);
  els.dangerResetBtn?.addEventListener("click", resetAllDataDanger);
  els.resetCastsBtn?.addEventListener("click", resetAllCastsDanger);
  els.resetVehiclesBtn?.addEventListener("click", resetAllVehiclesDanger);

  els.saveCastBtn?.addEventListener("click", saveCast);
  els.guessAreaBtn?.addEventListener("click", guessCastArea);
  els.castAddress?.addEventListener("input", () => {
    const nextKey = normalizeGeocodeAddressKey(els.castAddress?.value || "");
    if (nextKey && nextKey !== lastCastGeocodeKey) {
      if (els.castLat) els.castLat.value = "";
      if (els.castLng) els.castLng.value = "";
      if (els.castLatLngText) els.castLatLngText.value = "";
      if (els.castDistanceKm) els.castDistanceKm.value = "";
    }
    scheduleCastAutoGeocode();
  });
  els.castAddress?.addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    triggerCastAddressGeocodeNow();
  });
  els.castLatLngText?.addEventListener("change", () => {
    const hasText = String(els.castLatLngText?.value || "").trim();
    if (!hasText) return;
    applyCastLatLng();
  });
  els.openGoogleMapBtn?.addEventListener("click", () => openGoogleMap(els.castAddress?.value || "", els.castLat?.value, els.castLng?.value));
  els.cancelEditBtn?.addEventListener("click", resetCastForm);
  els.importCsvBtn?.addEventListener("click", () => els.csvFileInput?.click());
  els.exportCsvBtn?.addEventListener("click", exportCastsCsv);
  els.csvFileInput?.addEventListener("change", importCastCsvFile);
  els.castSearchRunBtn?.addEventListener("click", renderCastSearchResults);
  els.castSearchResetBtn?.addEventListener("click", resetCastSearchFilters);
  els.castSearchName?.addEventListener("input", renderCastSearchResults);
  els.castSearchArea?.addEventListener("input", renderCastSearchResults);
  els.castSearchAddress?.addEventListener("input", renderCastSearchResults);
  els.castSearchPhone?.addEventListener("input", renderCastSearchResults);

  els.saveVehicleBtn?.addEventListener("click", saveVehicle);
  els.cancelVehicleEditBtn?.addEventListener("click", resetVehicleForm);
  els.importVehicleCsvBtn?.addEventListener("click", () => els.vehicleCsvFileInput?.click());
  els.exportVehicleCsvBtn?.addEventListener("click", exportVehiclesCsv);
  els.vehicleCsvFileInput?.addEventListener("change", importVehicleCsvFile);
  els.exportPlansCsvBtn?.addEventListener("click", exportPlansCsv);
  els.importPlansCsvBtn?.addEventListener("click", triggerImportPlansCsv);
  els.plansCsvFileInput?.addEventListener("change", importPlansCsvFile);
  els.previewMileageReportBtn?.addEventListener("click", previewDriverMileageReport);
  els.exportMileageReportBtn?.addEventListener("click", exportDriverMileageReportXlsx);

  els.savePlanBtn?.addEventListener("click", savePlan);
  els.guessPlanAreaBtn?.addEventListener("click", guessPlanArea);
  els.clearPlansBtn?.addEventListener("click", clearAllPlans);

  els.saveActualBtn?.addEventListener("click", saveActual);
  els.guessActualAreaBtn?.addEventListener("click", guessActualArea);

  bindPlanAndActualFormEvents();
  setupSearchableCastInputs();
  bindDispatchEvents();
  bindPostDispatchEvents();

  els.checkAllVehiclesBtn?.addEventListener("click", () => toggleAllVehicles(true));
  els.uncheckAllVehiclesBtn?.addEventListener("click", () => toggleAllVehicles(false));
  els.clearManualLastVehicleBtn?.addEventListener("click", clearManualLastVehicle);
  els.resetMonthlySummaryBtn?.addEventListener("click", resetMonthlySummary);

  els.dispatchDate?.addEventListener("change", syncDateAndReloadFromDispatchDate);
  els.planDate?.addEventListener("change", syncDateAndReloadFromPlanDate);
  els.actualDate?.addEventListener("change", syncDateAndReloadFromActualDate);

  els.copyResultBtn?.addEventListener("click", copyDispatchResult);
  els.sendLineBtn?.addEventListener("click", sendDispatchResultToLine);
  els.saveDailyMileageBtn?.addEventListener("click", saveDailyMileageReports);

  els.copyActualTableBtn?.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(els.actualTableWrap?.innerText || "");
      alert("表をコピーしました");
    } catch (e) {
      console.error(e);
      alert("コピーに失敗しました");
    }
  });
}

function sendDispatchResultToLine() {
  const text = buildCopyResultText();
  const url = "https://line.me/R/msg/text/?" + encodeURIComponent(text);
  window.open(url, "_blank");
}

document.addEventListener("DOMContentLoaded", async () => {
  try {
    console.log("SUPABASE_URL:", SUPABASE_URL);

    const ok = await ensureAuth();
    if (!ok) return;

    setupTabs();
    setupEvents();

    resetCastForm();
    resetVehicleForm();
    resetPlanForm();
    resetActualForm();

    if (els.originLabelText) {
      els.originLabelText.value = ORIGIN_LABEL || "松戸駅";
    }

    const today = todayStr();
    if (els.dispatchDate) els.dispatchDate.value = today;
    if (els.planDate) els.planDate.value = today;
    if (els.actualDate) els.actualDate.value = today;

    await loadHomeAndAll();
    renderManualLastVehicleInfo();
  } catch (err) {
    console.error("dashboard init error:", err);
    alert("初期化中にエラーが発生しました。Console を確認してください。");
  }
});


/* ===== THEMIS v3.7 配車AI強化版 patch start ===== */
const THEMIS_V37_LEARN_KEY = "themis_v37_dispatch_learning_v1";

function getThemisV37LearningStore() {
  try {
    return JSON.parse(window.localStorage.getItem(THEMIS_V37_LEARN_KEY) || "{}") || {};
  } catch (e) {
    console.error(e);
    return {};
  }
}

function saveThemisV37LearningStore(store) {
  try {
    window.localStorage.setItem(THEMIS_V37_LEARN_KEY, JSON.stringify(store || {}));
  } catch (e) {
    console.error(e);
  }
}

function normalizeMunicipalityLabel(value) {
  return String(value || "").trim().replace(/[　\s]+/g, "");
}

function extractMunicipalityFromAddress(address) {
  const normalized = normalizeAddressText(address || "");
  if (!normalized) return "";

  const patterns = [
    /(東京都[^0-9\-]{1,12}?区)/,
    /(東京都[^0-9\-]{1,16}?市)/,
    /((?:北海道|大阪府|京都府|[^都道府県]{2,6}県)[^0-9\-]{1,16}?市)/,
    /((?:北海道|大阪府|京都府|[^都道府県]{2,6}県)[^0-9\-]{1,16}?郡[^0-9\-]{1,16}町)/,
    /((?:北海道|大阪府|京都府|[^都道府県]{2,6}県)[^0-9\-]{1,16}?郡[^0-9\-]{1,16}村)/,
    /((?:北海道|大阪府|京都府|[^都道府県]{2,6}県)[^0-9\-]{1,16}?町)/,
    /((?:北海道|大阪府|京都府|[^都道府県]{2,6}県)[^0-9\-]{1,16}?村)/
  ];

  for (const pattern of patterns) {
    const matched = normalized.match(pattern);
    if (matched && matched[1]) return normalizeMunicipalityLabel(matched[1]);
  }
  return "";
}

const THEMIS_V37_MUNICIPALITY_AREA_HINTS = [
  { area: "葛飾方面", keys: ["東京都葛飾区"] },
  { area: "足立方面", keys: ["東京都足立区"] },
  { area: "江戸川方面", keys: ["東京都江戸川区"] },
  { area: "墨田方面", keys: ["東京都墨田区"] },
  { area: "江東方面", keys: ["東京都江東区"] },
  { area: "荒川方面", keys: ["東京都荒川区"] },
  { area: "台東方面", keys: ["東京都台東区"] },
  { area: "市川方面", keys: ["千葉県市川市"] },
  { area: "船橋方面", keys: ["千葉県船橋市", "千葉県習志野市"] },
  { area: "鎌ヶ谷方面", keys: ["千葉県鎌ケ谷市", "千葉県鎌ヶ谷市"] },
  { area: "我孫子方面", keys: ["千葉県我孫子市"] },
  { area: "柏方面", keys: ["千葉県柏市"] },
  { area: "流山方面", keys: ["千葉県流山市"] },
  { area: "野田方面", keys: ["千葉県野田市"] },
  { area: "松戸近郊", keys: ["千葉県松戸市"] },
  { area: "三郷方面", keys: ["埼玉県三郷市"] },
  { area: "吉川方面", keys: ["埼玉県吉川市"] },
  { area: "八潮方面", keys: ["埼玉県八潮市"] },
  { area: "草加方面", keys: ["埼玉県草加市"] },
  { area: "越谷方面", keys: ["埼玉県越谷市"] },
  { area: "取手方面", keys: ["茨城県取手市"] },
  { area: "藤代方面", keys: ["茨城県取手市藤代"] },
  { area: "守谷方面", keys: ["茨城県守谷市"] },
  { area: "つくば方面", keys: ["茨城県つくば市"] },
  { area: "牛久方面", keys: ["茨城県牛久市"] }
];

function getAreaByMunicipality(address) {
  const municipality = extractMunicipalityFromAddress(address);
  if (!municipality) return "";
  for (const row of THEMIS_V37_MUNICIPALITY_AREA_HINTS) {
    if (row.keys.some(key => municipality.includes(normalizeMunicipalityLabel(key)) || normalizeMunicipalityLabel(key).includes(municipality))) {
      return row.area;
    }
  }
  return "";
}

const _THEMIS_V36_guessArea = guessArea;
// v3.7 municipality extraction is intentionally disabled.
// Keep the cleaner pre-v3.7 area labeling/display while preserving other v3.7 logic.
guessArea = function(lat, lng, address = "") {
  return _THEMIS_V36_guessArea(lat, lng, address);
};

function getThemisV37LearnedAreaScore(homeArea, destArea) {
  const store = getThemisV37LearningStore();
  const key = `${getCanonicalArea(homeArea) || normalizeAreaLabel(homeArea)}__${getCanonicalArea(destArea) || normalizeAreaLabel(destArea)}`;
  return Number(store.areaPair?.[key] || 0);
}

function learnThemisV37FromDoneRows(rows) {
  if (!Array.isArray(rows) || !rows.length) return;
  const store = getThemisV37LearningStore();
  store.areaPair = store.areaPair || {};
  store.routePair = store.routePair || {};

  const byVehicle = new Map();
  rows.forEach(row => {
    const vehicleId = Number(row.vehicle_id || 0);
    if (!vehicleId) return;
    if (!byVehicle.has(vehicleId)) byVehicle.set(vehicleId, []);
    byVehicle.get(vehicleId).push(row);

    const homeArea = normalizeAreaLabel(
      allVehiclesCache.find(v => Number(v.id) === vehicleId)?.home_area || row.driver_home_area || ""
    );
    const destArea = normalizeAreaLabel(row.destination_area || row.cluster_area || "無し");
    if (homeArea && destArea && homeArea !== "無し" && destArea !== "無し") {
      const key = `${getCanonicalArea(homeArea) || homeArea}__${getCanonicalArea(destArea) || destArea}`;
      store.areaPair[key] = Math.min(120, Number(store.areaPair[key] || 0) + 3);
    }
  });

  for (const rowsByVehicle of byVehicle.values()) {
    const ordered = [...rowsByVehicle].sort((a, b) => {
      const ah = Number(a.actual_hour ?? 0);
      const bh = Number(b.actual_hour ?? 0);
      if (ah !== bh) return ah - bh;
      return Number(a.stop_order || 0) - Number(b.stop_order || 0);
    });
    for (let i = 0; i < ordered.length - 1; i += 1) {
      const a = normalizeAreaLabel(ordered[i].destination_area || "");
      const b = normalizeAreaLabel(ordered[i + 1].destination_area || "");
      if (!a || !b || a === "無し" || b === "無し") continue;
      const key = `${getCanonicalArea(a) || a}__${getCanonicalArea(b) || b}`;
      store.routePair[key] = Math.min(80, Number(store.routePair[key] || 0) + 2);
    }
  }

  saveThemisV37LearningStore(store);
}

const _THEMIS_V36_confirmDailyToMonthly = confirmDailyToMonthly;
confirmDailyToMonthly = async function() {
  const doneRowsBefore = Array.isArray(currentActualsCache)
    ? currentActualsCache.filter(x => normalizeStatus(x.status) === "done")
    : [];
  const result = await _THEMIS_V36_confirmDailyToMonthly.apply(this, arguments);
  try {
    learnThemisV37FromDoneRows(doneRowsBefore);
  } catch (e) {
    console.error(e);
  }
  return result;
};

const _THEMIS_V36_getLastTripHomePriorityWeight = getLastTripHomePriorityWeight;
getLastTripHomePriorityWeight = function(clusterArea, homeArea, isLastRun, isDefaultLastHourCluster) {
  let weight = _THEMIS_V36_getLastTripHomePriorityWeight(clusterArea, homeArea, isLastRun, isDefaultLastHourCluster);
  const learned = getThemisV37LearnedAreaScore(homeArea, clusterArea);
  const strict = getStrictHomeCompatibilityScore(clusterArea, homeArea);
  const direction = getDirectionAffinityScore(clusterArea, homeArea);

  let returnTimeScore = 0;
  if (strict >= 78) returnTimeScore += 42;
  else if (strict >= 52) returnTimeScore += 22;
  if (direction >= 72) returnTimeScore += 18;
  else if (direction >= 28) returnTimeScore += 8;
  if (isHardReverseForHome(clusterArea, homeArea)) returnTimeScore -= (isLastRun ? 120 : 50);

  weight += learned * (isLastRun ? 1.6 : 0.8);
  weight += returnTimeScore * (isLastRun ? 1.8 : (isDefaultLastHourCluster ? 1.2 : 0.35));
  return weight;
};

function getThemisV37LearnedRoutePairScore(areaA, areaB) {
  const store = getThemisV37LearningStore();
  const key1 = `${getCanonicalArea(areaA) || normalizeAreaLabel(areaA)}__${getCanonicalArea(areaB) || normalizeAreaLabel(areaB)}`;
  const key2 = `${getCanonicalArea(areaB) || normalizeAreaLabel(areaB)}__${getCanonicalArea(areaA) || normalizeAreaLabel(areaA)}`;
  return Math.max(Number(store.routePair?.[key1] || 0), Number(store.routePair?.[key2] || 0));
}

function getThemisV37RouteSequenceScore(fromItem, toItem) {
  const pointA = getItemLatLng(fromItem);
  const pointB = getItemLatLng(toItem);
  const areaA = normalizeAreaLabel(fromItem?.destination_area || fromItem?.cluster_area || fromItem?.planned_area || "無し");
  const areaB = normalizeAreaLabel(toItem?.destination_area || toItem?.cluster_area || toItem?.planned_area || "無し");
  const routeFlow = getRouteFlowCompatibilityBetweenAreas(areaA, areaB);
  const continuityPenalty = getPairRouteContinuityPenalty(areaA, areaB);
  const learned = getThemisV37LearnedRoutePairScore(areaA, areaB);
  let score = routeFlow * 2.4 + learned * 4.2 - continuityPenalty * 1.35;

  if (pointA && pointB) {
    const leg = estimateRoadKmBetweenPoints(pointA.lat, pointA.lng, pointB.lat, pointB.lng);
    score -= leg * 3.2;
  } else {
    score -= Math.abs(Number(toItem?.distance_km || 0) - Number(fromItem?.distance_km || 0)) * 0.4;
  }

  const dirA = getAreaDirectionCluster(areaA);
  const dirB = getAreaDirectionCluster(areaB);
  if (dirA && dirB && dirA === dirB) score += 18;
  return score;
}

sortItemsByNearestRoute = function(items) {
  const remaining = [...items];
  const sorted = [];
  let currentLat = ORIGIN_LAT;
  let currentLng = ORIGIN_LNG;
  let currentItem = null;

  while (remaining.length) {
    let bestIndex = 0;
    let bestScore = -Infinity;

    remaining.forEach((item, index) => {
      const point = getItemLatLng(item);
      let score = 0;
      if (point) {
        score -= estimateRoadKmBetweenPoints(currentLat, currentLng, point.lat, point.lng) * 3.0;
      } else {
        score -= Number(item.distance_km || 999999) * 1.1;
      }

      if (currentItem) {
        score += getThemisV37RouteSequenceScore(currentItem, item);
      } else {
        const area = normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || "無し");
        score += getRouteFlowSortWeight(area) * 4.8;
        score += getAreaAffinityScore(area, "松戸近郊") * 0.12;
      }

      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });

    const picked = remaining.splice(bestIndex, 1)[0];
    sorted.push(picked);
    currentItem = picked;

    const pickedPoint = getItemLatLng(picked);
    if (pickedPoint) {
      currentLat = pickedPoint.lat;
      currentLng = pickedPoint.lng;
    }
  }

  return sorted;
};

const _THEMIS_V36_runAutoDispatch = runAutoDispatch;
runAutoDispatch = async function() {
  const result = await _THEMIS_V36_runAutoDispatch.apply(this, arguments);
  try {
    await loadActualsByDate(els.actualDate?.value || todayStr());
    renderDailyDispatchResult();
  } catch (e) {
    console.error(e);
  }
  return result;
};
/* ===== THEMIS v3.7 配車AI強化版 patch end ===== */
