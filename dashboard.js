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
let activeVehicleIdsForToday = new Set();

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
  dangerResetBtn: document.getElementById("dangerResetBtn"),

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
  castLat: document.getElementById("castLat"),
  castLng: document.getElementById("castLng"),
  saveCastBtn: document.getElementById("saveCastBtn"),
  guessAreaBtn: document.getElementById("guessAreaBtn"),
  applyLatLngBtn: document.getElementById("applyLatLngBtn"),
  openGoogleMapBtn: document.getElementById("openGoogleMapBtn"),
  cancelEditBtn: document.getElementById("cancelEditBtn"),
  importCsvBtn: document.getElementById("importCsvBtn"),
  exportCsvBtn: document.getElementById("exportCsvBtn"),
  csvFileInput: document.getElementById("csvFileInput"),
  castsTableBody: document.getElementById("castsTableBody"),

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

  dispatchDate: document.getElementById("dispatchDate"),
  optimizeBtn: document.getElementById("optimizeBtn"),
  confirmDailyBtn: document.getElementById("confirmDailyBtn"),
  clearActualBtn: document.getElementById("clearActualBtn"),
  checkAllVehiclesBtn: document.getElementById("checkAllVehiclesBtn"),
  uncheckAllVehiclesBtn: document.getElementById("uncheckAllVehiclesBtn"),
  dailyVehicleChecklist: document.getElementById("dailyVehicleChecklist"),
  copyResultBtn: document.getElementById("copyResultBtn"),
  dailyDispatchResult: document.getElementById("dailyDispatchResult"),

  planDate: document.getElementById("planDate"),
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

  actualDate: document.getElementById("actualDate"),
  addSelectedPlanBtn: document.getElementById("addSelectedPlanBtn"),
  copyActualTableBtn: document.getElementById("copyActualTableBtn"),
  planSelect: document.getElementById("planSelect"),
  castSelect: document.getElementById("castSelect"),
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
  if (isFridayOrSaturday(dateStr)) return 5;
  if (isSpecialLateNightEve(dateStr)) return 5;
  return 4;
}

function isValidLatLng(lat, lng) {
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return false;
  if (lat < -90 || lat > 90) return false;
  if (lng < -180 || lng > 180) return false;
  return true;
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

function normalizeAddressText(address) {
  return String(address || "")
    .trim()
    .replace(/[　\s]+/g, "")
    .replace(/ヶ/g, "ケ")
    .replace(/之/g, "の");
}

function classifyAreaByAddress(address) {
  const a = normalizeAddressText(address);
  if (!a) return "";
  if (a.includes("船橋")) return "船橋方面";
  if (a.includes("柏")) return "柏方面";
  if (a.includes("松戸")) return "松戸近郊";
  if (a.includes("市川")) return "市川方面";
  if (a.includes("流山")) return "流山方面";
  if (a.includes("八千代")) return "八千代方面";
  if (a.includes("三郷")) return "三郷方面";
  if (a.includes("守谷")) return "守谷方面";
  if (a.includes("葛飾")) return "葛飾方面";
  if (a.includes("江戸川")) return "都内方面";
  if (a.includes("東京")) return "都内方面";
  if (a.includes("茨城")) return "茨城方面";
  return "";
}

function classifyAreaByLatLng(lat, lng) {
  if (!isValidLatLng(lat, lng)) return "周辺";
  if (lng >= 139.99) return "都内方面";
  if (lat >= 35.84 && lng >= 139.94) return "柏方面";
  if (lat >= 35.80 && lat < 35.86 && lng >= 139.90 && lng < 139.96) return "松戸近郊";
  if (lat >= 35.69 && lat < 35.80 && lng >= 139.93 && lng < 140.09) return "船橋方面";
  if (lng > 140.05) return "茨城方面";
  return "周辺";
}

function guessArea(lat, lng, address = "") {
  return classifyAreaByAddress(address) || classifyAreaByLatLng(lat, lng);
}

function normalizeAreaLabel(area) {
  const value = String(area || "").trim();
  if (!value) return "無し";
  if (value.includes("柏")) return "柏方面";
  if (value.includes("松戸")) return "松戸近郊";
  if (value.includes("船橋")) return "船橋方面";
  if (value.includes("市川")) return "市川方面";
  if (value.includes("流山")) return "流山方面";
  if (value.includes("八千代")) return "八千代方面";
  if (value.includes("三郷")) return "三郷方面";
  if (value.includes("守谷")) return "守谷方面";
  if (value.includes("葛飾")) return "葛飾方面";
  if (value.includes("都内") || value.includes("東京") || value.includes("江戸川")) return "都内方面";
  if (value.includes("茨城")) return "茨城方面";
  return value;
}

function openGoogleMap(address) {
  if (!address) return;
  const origin = encodeURIComponent(ORIGIN_LABEL);
  const dest = encodeURIComponent(address);
  window.open(
    `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${dest}&travelmode=driving`,
    "_blank"
  );
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

  if (vehicleArea && normalizedArea && vehicleArea === normalizedArea) {
    score += 30;
  }
  if (homeArea && normalizedArea && homeArea === normalizedArea) {
    score += 12;
  }

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

  return [...clusterMap.values()].sort((a, b) => {
    if (a.hour !== b.hour) return a.hour - b.hour;
    if (b.count !== a.count) return b.count - a.count;
    return b.totalDistance - a.totalDistance;
  });
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
      <span class="chip">${escapeHtml(vehicle.plate_number || "-")}</span>
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

  const lat = toNullableNumber(els.castLat?.value);
  const lng = toNullableNumber(els.castLng?.value);
  const manualArea = els.castArea?.value.trim() || "";
  const autoArea = guessArea(lat, lng, address);
  const autoDistance = isValidLatLng(lat, lng)
    ? estimateRoadKmFromStation(lat, lng)
    : null;

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
      <td>${escapeHtml(cast.name || "")}</td>
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

    const inserts = [];

    for (const row of rows) {
      const name = String(row.name || "").trim();
      const address = String(row.address || "").trim();
      if (!name) continue;

      const exists = allCastsCache.find(
        c =>
          String(c.name || "").trim() === name &&
          String(c.address || "").trim() === address
      );
      if (exists) {
        console.log("重複スキップ:", name, address);
        continue;
      }

      const lat = toNullableNumber(row.latitude);
      const lng = toNullableNumber(row.longitude);
      const autoArea = guessArea(lat, lng, address);

      inserts.push({
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

    if (!inserts.length) {
      alert("新規キャストはありません");
      els.csvFileInput.value = "";
      return;
    }

    const { error } = await supabaseClient.from("casts").insert(inserts);
    if (error) {
      console.error("CSV import supabase error:", error);
      alert("CSV取込エラー: " + error.message);
      return;
    }

    els.csvFileInput.value = "";
    await addHistory(null, null, "import_csv", `${inserts.length}件のキャストをCSV取込`);
    alert(`${inserts.length}件のキャストを取り込みました`);
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
}

function guessCastArea() {
  const lat = toNullableNumber(els.castLat?.value);
  const lng = toNullableNumber(els.castLng?.value);
  if (els.castArea) {
    els.castArea.value = normalizeAreaLabel(guessArea(lat, lng, els.castAddress?.value || ""));
  }
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
      <td>${escapeHtml(vehicle.plate_number || "")}</td>
      <td>${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || ""))}</td>
      <td>${escapeHtml(normalizeAreaLabel(vehicle.home_area || ""))}</td>
      <td>${vehicle.seat_capacity ?? ""}</td>
      <td>${escapeHtml(vehicle.driver_name || "")}</td>
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
function resetPlanForm() {
  editingPlanId = null;
  if (els.planCastSelect) els.planCastSelect.value = "";
  if (els.planHour) els.planHour.value = "0";
  if (els.planDistanceKm) els.planDistanceKm.value = "";
  if (els.planAddress) els.planAddress.value = "";
  if (els.planArea) els.planArea.value = "";
  if (els.planNote) els.planNote.value = "";
  if (els.cancelPlanEditBtn) els.cancelPlanEditBtn.classList.add("hidden");
}

function fillPlanForm(plan) {
  editingPlanId = plan.id;
  if (els.planCastSelect) els.planCastSelect.value = String(plan.cast_id || "");
  if (els.planHour) els.planHour.value = String(plan.plan_hour ?? 0);
  if (els.planDistanceKm) els.planDistanceKm.value = plan.distance_km ?? "";
  if (els.planAddress) els.planAddress.value = plan.destination_address || plan.casts?.address || "";
  if (els.planArea) els.planArea.value = normalizeAreaLabel(plan.planned_area || "");
  if (els.planNote) els.planNote.value = plan.note || "";
  if (els.cancelPlanEditBtn) els.cancelPlanEditBtn.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function fillPlanFormFromSelectedCast() {
  const castId = Number(els.planCastSelect?.value || 0);
  if (!castId) return;

  const cast = allCastsCache.find(c => Number(c.id) === castId);
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
    let distance = toNullableNumber(cast.distance_km);
    if (distance === null) {
      const lat = toNullableNumber(cast.latitude);
      const lng = toNullableNumber(cast.longitude);
      if (isValidLatLng(lat, lng)) {
        distance = estimateRoadKmFromStation(lat, lng);
      }
    }
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

function renderPlanCastSelect() {
  if (!els.planCastSelect) return;

  const plannedIds = getPlannedCastIds();
  const editingPlan = editingPlanId
    ? currentPlansCache.find(x => Number(x.id) === Number(editingPlanId))
    : null;
  const editingCastIdForPlan = Number(editingPlan?.cast_id || 0);

  els.planCastSelect.innerHTML = `<option value="">選択してください</option>`;

  allCastsCache
    .filter(cast => Number(cast.id) === editingCastIdForPlan || !plannedIds.has(Number(cast.id)))
    .forEach(cast => {
      const option = document.createElement("option");
      option.value = cast.id;
      option.textContent = `${cast.name} | ${normalizeAreaLabel(cast.area || "-")}`;
      option.dataset.address = cast.address || "";
      option.dataset.area = normalizeAreaLabel(cast.area || "");
      option.dataset.distance = cast.distance_km ?? "";
      els.planCastSelect.appendChild(option);
    });
}

function renderPlanSelect() {
  if (!els.planSelect) return;

  const targetDate = els.actualDate?.value || todayStr();
  const doneCastIds = getDoneCastIdsInActuals();

  els.planSelect.innerHTML = `<option value="">予定から選択</option>`;

  currentPlansCache
    .filter(plan => plan.plan_date === targetDate)
    .filter(plan => plan.status === "planned")
    .filter(plan => !doneCastIds.has(Number(plan.cast_id)))
    .forEach(plan => {
      const option = document.createElement("option");
      option.value = plan.id;
      option.textContent = `${getHourLabel(plan.plan_hour)} / ${plan.casts?.name || "-"} / ${normalizeAreaLabel(plan.planned_area || "-")}`;
      els.planSelect.appendChild(option);
    });
}

async function savePlan() {
  const castId = Number(els.planCastSelect?.value);
  if (!castId) {
    alert("キャストを選択してください");
    return;
  }

  const planDate = els.planDate?.value || todayStr();
  const hour = Number(els.planHour?.value || 0);
  const address = els.planAddress?.value.trim() || "";
  const distanceKm = toNullableNumber(els.planDistanceKm?.value);
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
    const areas = [...new Set(hourItems.map(x => normalizeAreaLabel(x.planned_area || "無し")))];

    html += `<div class="grouped-section">`;
    html += `<div class="grouped-hour-title">${getHourLabel(hour)}</div>`;

    areas.forEach(area => {
      const areaItems = hourItems.filter(
        x => normalizeAreaLabel(x.planned_area || "無し") === area
      );

      html += `<div class="grouped-area-title">${escapeHtml(area)}</div>`;

      areaItems.forEach(plan => {
        html += `
          <div class="grouped-row">
            <div>${getHourLabel(hour)}</div>
            <div><strong>${escapeHtml(plan.casts?.name || "-")}</strong></div>
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
      currentPlansCache.map(x => normalizeAreaLabel(x.planned_area || "無し"))
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
          normalizeAreaLabel(plan.planned_area || "無し") === area
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
                  <span>${escapeHtml(row.casts?.name || "-")} (${Number(row.distance_km || 0).toFixed(1)}km)</span>
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
  if (els.actualDistanceKm) els.actualDistanceKm.value = "";
  if (els.actualStatus) els.actualStatus.value = "pending";
  if (els.actualAddress) els.actualAddress.value = "";
  if (els.actualArea) els.actualArea.value = "";
  if (els.actualNote) els.actualNote.value = "";
  if (els.cancelActualEditBtn) els.cancelActualEditBtn.classList.add("hidden");
}

function fillActualForm(item) {
  editingActualId = item.id;
  if (els.castSelect) els.castSelect.value = String(item.cast_id || "");
  if (els.actualHour) els.actualHour.value = String(item.actual_hour ?? 0);
  if (els.actualDistanceKm) els.actualDistanceKm.value = item.distance_km ?? "";
  if (els.actualStatus) els.actualStatus.value = item.status || "pending";
  if (els.actualAddress) els.actualAddress.value = item.destination_address || item.casts?.address || "";
  if (els.actualArea) els.actualArea.value = normalizeAreaLabel(item.destination_area || item.casts?.area || "");
  if (els.actualNote) els.actualNote.value = item.note || "";
  if (els.cancelActualEditBtn) els.cancelActualEditBtn.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function fillActualFormFromSelectedCast() {
  const castId = Number(els.castSelect?.value || 0);
  if (!castId) return;

  const cast = allCastsCache.find(c => Number(c.id) === castId);
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

  if (els.castSelect) els.castSelect.value = String(plan.cast_id || "");
  if (els.actualHour) els.actualHour.value = String(plan.plan_hour ?? 0);
  if (els.actualAddress) els.actualAddress.value = plan.destination_address || plan.casts?.address || "";
  if (els.actualArea) els.actualArea.value = normalizeAreaLabel(plan.planned_area || plan.casts?.area || "無し");
  if (els.actualDistanceKm) els.actualDistanceKm.value = plan.distance_km ?? plan.casts?.distance_km ?? "";
  if (els.actualNote) els.actualNote.value = plan.note || "";
}

function renderCastSelects() {
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

  if (els.castSelect) {
    els.castSelect.innerHTML = `<option value="">選択してください</option>`;

    allCastsCache
      .filter(
        cast =>
          Number(cast.id) === editingCastIdForActual ||
          (!usedCastIds.has(Number(cast.id)) && !doneCastIds.has(Number(cast.id)))
      )
      .forEach(cast => {
        const option = document.createElement("option");
        option.value = cast.id;
        option.textContent = `${cast.name} | ${normalizeAreaLabel(cast.area || "-")}`;
        option.dataset.address = cast.address || "";
        option.dataset.area = normalizeAreaLabel(cast.area || "");
        option.dataset.distance = cast.distance_km ?? "";
        option.dataset.lat = cast.latitude ?? "";
        option.dataset.lng = cast.longitude ?? "";
        els.castSelect.appendChild(option);
      });
  }

  renderPlanCastSelect();
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
}

async function saveActual() {
  const castId = Number(els.castSelect?.value);
  if (!castId) {
    alert("キャストを選択してください");
    return;
  }

  const dateStr = els.actualDate?.value || todayStr();
  const hour = Number(els.actualHour?.value || 0);
  const address = els.actualAddress?.value.trim() || "";
  const area = normalizeAreaLabel(els.actualArea?.value.trim() || "無し");
  const distanceKm = toNullableNumber(els.actualDistanceKm?.value);
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

  if (currentActualsCache.some(x => Number(x.cast_id) === Number(plan.cast_id))) {
    alert("そのキャストはすでにActualにあります");
    return;
  }

  const doneCastIds = getDoneCastIdsInActuals();
  if (doneCastIds.has(Number(plan.cast_id))) {
    alert("このキャストはすでに送り完了です");
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
    const areas = [...new Set(hourItems.map(x => normalizeAreaLabel(x.destination_area || "無し")))];

    html += `<div class="grouped-section">`;
    html += `<div class="grouped-hour-title">${getHourLabel(hour)}</div>`;

    areas.forEach(area => {
      html += `<div class="grouped-area-title">${escapeHtml(area)}</div>`;

      hourItems
        .filter(item => normalizeAreaLabel(item.destination_area || "無し") === area)
        .forEach(item => {
          html += `
            <div class="grouped-row">
              <div>${getHourLabel(hour)}</div>
              <div><strong>${escapeHtml(item.casts?.name || "-")}</strong></div>
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

  els.actualTableWrap.querySelectorAll(".actual-done-btn").forEach(btn => {
    btn.addEventListener("click", async () => updateActualStatus(Number(btn.dataset.id), "done"));
  });

  els.actualTableWrap.querySelectorAll(".actual-cancel-btn").forEach(btn => {
    btn.addEventListener("click", async () => updateActualStatus(Number(btn.dataset.id), "cancel"));
  });

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
}

function renderActualTimeAreaMatrix() {
  if (!els.actualTimeAreaMatrix) return;

  const hours = [0, 1, 2, 3, 4, 5];
  const areas = [...new Set(currentActualsCache.map(x => normalizeAreaLabel(x.destination_area || "無し")))];

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
          normalizeAreaLabel(x.destination_area || "無し") === area
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
                      <span>${escapeHtml(row.casts?.name || "-")} (${Number(row.distance_km || 0).toFixed(1)}km)</span>
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
        ${escapeHtml(vehicle.plate_number || "")}
        （${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))} / 帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))} / 定員${vehicle.seat_capacity ?? "-"} / ${escapeHtml(vehicle.driver_name || "-")}）
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
  renderDailyDispatchResult();
}

function buildMonthlyDistanceMapForCurrentMonth() {
  const monthKey = getMonthKey(els.dispatchDate?.value || todayStr());
  return getVehicleMonthlyStatsMap(currentDailyReportsCache, monthKey);
}

function optimizeAssignments(items, vehicles, monthlyMap) {
  const workingVehicles = vehicles.filter(v => v.status !== "maintenance");
  const clusters = buildDispatchClusters(items);
  const assignments = [];

  if (!workingVehicles.length || !clusters.length) return assignments;

  const vehicleUsage = new Map();

  function getVehicleState(vehicleId) {
    if (!vehicleUsage.has(vehicleId)) {
      vehicleUsage.set(vehicleId, {
        totalAssigned: 0,
        totalDistance: 0,
        hourLoads: new Map()
      });
    }
    return vehicleUsage.get(vehicleId);
  }

  function getHourLoad(vehicleId, hour) {
    const state = getVehicleState(vehicleId);
    return Number(state.hourLoads.get(hour) || 0);
  }

  function addHourLoad(vehicleId, hour, count, distance) {
    const state = getVehicleState(vehicleId);
    state.totalAssigned += count;
    state.totalDistance += distance;
    state.hourLoads.set(hour, getHourLoad(vehicleId, hour) + count);
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

        let score = 1000;
        score -= getVehicleAreaMatchScore(vehicle, cluster.area);
        score += sameHourLoad * 35;
        score += getVehicleState(vehicle.id).totalAssigned * 8;
        score += getVehicleState(vehicle.id).totalDistance * 0.08;
        score += monthly.avgDistance * 0.25;
        score += cluster.totalDistance * 0.15;

        if (isLastRun) {
          if (homeArea && normalizedClusterArea && homeArea === normalizedClusterArea) score -= 120;
          else score += 40;
        }

        if (!isLastRun && isDefaultLastHourCluster) {
          if (homeArea && normalizedClusterArea && homeArea === normalizedClusterArea) score -= 25;
        }

        return { vehicle, score };
      })
      .filter(Boolean)
      .sort((a, b) => a.score - b.score);

    if (candidateScores.length) {
      const bestVehicle = candidateScores[0].vehicle;

      const sortedItems = [...cluster.items].sort((a, b) => {
        const aDist = Number(a.distance_km || 0);
        const bDist = Number(b.distance_km || 0);
        return aDist - bDist;
      });

      sortedItems.forEach(item => {
        assignments.push({
          item_id: item.id,
          actual_hour: cluster.hour,
          vehicle_id: bestVehicle.id,
          driver_name: bestVehicle.driver_name || "",
          distance_km: Number(item.distance_km || 0)
        });
      });

      addHourLoad(bestVehicle.id, cluster.hour, cluster.count, cluster.totalDistance);
      continue;
    }

    const splitItems = [...cluster.items].sort((a, b) => {
      const aDist = Number(a.distance_km || 0);
      const bDist = Number(b.distance_km || 0);
      return aDist - bDist;
    });

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

          let score = 1000;
          score -= getVehicleAreaMatchScore(vehicle, cluster.area);
          score += sameHourLoad * 35;
          score += getVehicleState(vehicle.id).totalAssigned * 8;
          score += getVehicleState(vehicle.id).totalDistance * 0.08;
          score += monthly.avgDistance * 0.25;
          score += Number(item.distance_km || 0) * 0.15;

          if (isLastRun) {
            if (homeArea && normalizedClusterArea && homeArea === normalizedClusterArea) score -= 120;
            else score += 40;
          }

          if (!isLastRun && isDefaultLastHourCluster) {
            if (homeArea && normalizedClusterArea && homeArea === normalizedClusterArea) score -= 25;
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

      addHourLoad(bestVehicle.id, cluster.hour, 1, Number(item.distance_km || 0));
    }
  }

  return assignments;
}

async function runAutoDispatch() {
  const selectedVehicles = getSelectedVehiclesForToday();
  if (!selectedVehicles.length) {
    alert("本日使用する車両を選択してください");
    return;
  }

  if (!currentActualsCache.length) {
    alert("Actualがありません");
    return;
  }

  const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
  const assignments = optimizeAssignments(currentActualsCache, selectedVehicles, monthlyMap);
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
      console.error(error);
      alert(`配車更新エラー: ${error.message}`);
      return;
    }
  }

  await addHistory(currentDispatchId, null, "auto_dispatch", "自動配車を実行");
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
  renderDailyDispatchResult();
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

  const cardsHtml = vehicles
    .map(vehicle => {
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

      const totalDistance = rows.reduce((sum, row) => sum + Number(row.distance_km || 0), 0);

      const body = rows.length
        ? rows
            .map(
              (row, index) => `
                <div class="dispatch-row">
                  <div class="dispatch-left">
                    <span class="badge-time">${escapeHtml(getHourLabel(row.actual_hour))}</span>
                    <span class="badge-order">順番 ${index + 1}</span>
                    <span class="dispatch-name">${escapeHtml(row.casts?.name || "-")}</span>
                    <span class="dispatch-area">${escapeHtml(normalizeAreaLabel(row.destination_area || "-"))}</span>
                  </div>
                  <div class="dispatch-right">
                    <div class="dispatch-distance">${Number(row.distance_km || 0).toFixed(1)}km</div>
                    <select class="dispatch-vehicle-select" data-item-id="${row.id}">
                      ${vehicles
                        .map(
                          v => `
                            <option value="${v.id}" ${Number(v.id) === Number(vehicle.id) ? "selected" : ""}>
                              ${escapeHtml(v.plate_number || "")}
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
              <h4>${escapeHtml(vehicle.plate_number || "-")}</h4>
              <div class="vehicle-result-meta">
                ${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))}
                / 帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))}
                / 定員${vehicle.seat_capacity ?? "-"}
                / ${escapeHtml(vehicle.driver_name || "-")}
              </div>
            </div>
            <div class="vehicle-result-badges">
              <span class="metric-badge">人数 ${rows.length}</span>
              <span class="metric-badge">距離 ${totalDistance.toFixed(1)}km</span>
            </div>
          </div>
          <div class="vehicle-result-body">${body}</div>
        </div>
      `;
    })
    .join("");

  els.dailyDispatchResult.innerHTML = cardsHtml;

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
}

function buildCopyResultText() {
  const vehicles = getSelectedVehiclesForToday();
  const activeItems = currentActualsCache.filter(
    x => normalizeStatus(x.status) !== "done" && normalizeStatus(x.status) !== "cancel"
  );

  const lines = [];

  vehicles.forEach(vehicle => {
    const rows = activeItems
      .filter(item => Number(item.vehicle_id) === Number(vehicle.id))
      .sort((a, b) => {
        const ah = Number(a.actual_hour || 0);
        const bh = Number(b.actual_hour || 0);
        if (ah !== bh) return ah - bh;

        const aa = normalizeAreaLabel(a.destination_area || "");
        const ba = normalizeAreaLabel(b.destination_area || "");
        if (aa !== ba) return aa.localeCompare(ba, "ja");

        return Number(a.stop_order || 0) - Number(b.stop_order || 0);
      });

    lines.push(`${vehicle.line_id ? vehicle.line_id + " " : ""}${vehicle.driver_name || vehicle.plate_number || ""}`);

    if (!rows.length) {
      lines.push("送りなし");
    } else {
      rows.forEach(row => {
        lines.push(`${getHourLabel(row.actual_hour)} ${row.casts?.name || "-"} ${normalizeAreaLabel(row.destination_area || "-")}`);
      });
    }

    lines.push("");
  });

  return lines.join("\n").trim();
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
  } catch (error) {
    console.error("importAllDataFromFile error:", error);
    alert("全体インポートに失敗しました: " + error.message);
  }
}

async function resetAllDataDanger() {
  if (!window.confirm("全消去しますか？この操作は戻せません。")) return;
  if (!window.confirm("本当に全消去しますか？")) return;

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

  await loadCasts();
  await loadVehicles();
  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
  await loadHistory();

  renderDailyVehicleChecklist();
  renderDailyDispatchResult();
  renderHomeSummary();
  renderHomeMonthlyVehicleList();
}

async function syncDateAndReloadFromDispatchDate() {
  const dateStr = els.dispatchDate?.value || todayStr();
  if (els.planDate) els.planDate.value = dateStr;
  if (els.actualDate) els.actualDate.value = dateStr;

  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
  renderDailyDispatchResult();
}

async function syncDateAndReloadFromPlanDate() {
  const dateStr = els.planDate?.value || todayStr();
  if (els.dispatchDate) els.dispatchDate.value = dateStr;
  if (els.actualDate) els.actualDate.value = dateStr;

  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
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
  if (els.planCastSelect) els.planCastSelect.addEventListener("change", fillPlanFormFromSelectedCast);
  if (els.castSelect) els.castSelect.addEventListener("change", fillActualFormFromSelectedCast);
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
  els.dangerResetBtn?.addEventListener("click", resetAllDataDanger);

  els.saveCastBtn?.addEventListener("click", saveCast);
  els.guessAreaBtn?.addEventListener("click", guessCastArea);
  els.applyLatLngBtn?.addEventListener("click", applyCastLatLng);
  els.openGoogleMapBtn?.addEventListener("click", () => openGoogleMap(els.castAddress?.value || ""));
  els.cancelEditBtn?.addEventListener("click", resetCastForm);
  els.importCsvBtn?.addEventListener("click", () => els.csvFileInput?.click());
  els.exportCsvBtn?.addEventListener("click", exportCastsCsv);
  els.csvFileInput?.addEventListener("change", importCastCsvFile);

  els.saveVehicleBtn?.addEventListener("click", saveVehicle);
  els.cancelVehicleEditBtn?.addEventListener("click", resetVehicleForm);
  els.importVehicleCsvBtn?.addEventListener("click", () => els.vehicleCsvFileInput?.click());
  els.exportVehicleCsvBtn?.addEventListener("click", exportVehiclesCsv);
  els.vehicleCsvFileInput?.addEventListener("change", importVehicleCsvFile);

  els.savePlanBtn?.addEventListener("click", savePlan);
  els.guessPlanAreaBtn?.addEventListener("click", guessPlanArea);
  els.clearPlansBtn?.addEventListener("click", clearAllPlans);

  els.saveActualBtn?.addEventListener("click", saveActual);
  els.guessActualAreaBtn?.addEventListener("click", guessActualArea);

  bindPlanAndActualFormEvents();
  bindDispatchEvents();
  bindPostDispatchEvents();

  els.checkAllVehiclesBtn?.addEventListener("click", () => toggleAllVehicles(true));
  els.uncheckAllVehiclesBtn?.addEventListener("click", () => toggleAllVehicles(false));
  els.resetMonthlySummaryBtn?.addEventListener("click", resetMonthlySummary);

  els.dispatchDate?.addEventListener("change", syncDateAndReloadFromDispatchDate);
  els.planDate?.addEventListener("change", syncDateAndReloadFromPlanDate);
  els.actualDate?.addEventListener("change", syncDateAndReloadFromActualDate);

  els.copyResultBtn?.addEventListener("click", copyDispatchResult);
  els.sendLineBtn?.addEventListener("click", sendDispatchResultToLine);

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
  } catch (err) {
    console.error("dashboard init error:", err);
    alert("初期化中にエラーが発生しました。Console を確認してください。");
  }
});
