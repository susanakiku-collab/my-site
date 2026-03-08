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

let allCastsCache = [];
let allVehiclesCache = [];
let currentDispatchItemsCache = [];
let currentPlansCache = [];
let currentDailyReportsCache = [];

const BLOCKS = ["A1", "A2", "B", "C1", "D", "無し"];
const BOARD_ROWS = [
  { key: 0, label: "12" },
  { key: 1, label: "1" },
  { key: 2, label: "2" },
  { key: 3, label: "3" },
  { key: 4, label: "4" },
  { key: 5, label: "L" }
];

const els = {
  userEmail: document.getElementById("userEmail"),
  logoutBtn: document.getElementById("logoutBtn"),

  dispatchDate: document.getElementById("dispatchDate"),
  driverName: document.getElementById("driverName"),
  createDispatchBtn: document.getElementById("createDispatchBtn"),
  loadDispatchBtn: document.getElementById("loadDispatchBtn"),
  optimizeBtn: document.getElementById("optimizeBtn"),
  copyLineDispatchBtn: document.getElementById("copyLineDispatchBtn"),
  castSelect: document.getElementById("castSelect"),
  planSelect: document.getElementById("planSelect"),
  stopOrder: document.getElementById("stopOrder"),
  addCastBtn: document.getElementById("addCastBtn"),
  addPlanToDispatchBtn: document.getElementById("addPlanToDispatchBtn"),
  dispatchSummary: document.getElementById("dispatchSummary"),
  dispatchList: document.getElementById("dispatchList"),

  planDate: document.getElementById("planDate"),
  planHour: document.getElementById("planHour"),
  planCastSelect: document.getElementById("planCastSelect"),
  planArea: document.getElementById("planArea"),
  planNote: document.getElementById("planNote"),
  savePlanBtn: document.getElementById("savePlanBtn"),
  cancelPlanEditBtn: document.getElementById("cancelPlanEditBtn"),
  loadPlansBtn: document.getElementById("loadPlansBtn"),
  plansSummary: document.getElementById("plansSummary"),
  plansBoard: document.getElementById("plansBoard"),

  castName: document.getElementById("castName"),
  castPhone: document.getElementById("castPhone"),
  castAddress: document.getElementById("castAddress"),
  castLatLngText: document.getElementById("castLatLngText"),
  castArea: document.getElementById("castArea"),
  castLat: document.getElementById("castLat"),
  castLng: document.getElementById("castLng"),
  castMemo: document.getElementById("castMemo"),
  openGoogleMapBtn: document.getElementById("openGoogleMapBtn"),
  applyLatLngBtn: document.getElementById("applyLatLngBtn"),
  saveCastBtn: document.getElementById("saveCastBtn"),
  cancelEditBtn: document.getElementById("cancelEditBtn"),
  csvFileInput: document.getElementById("csvFileInput"),
  importCsvBtn: document.getElementById("importCsvBtn"),
  exportCsvBtn: document.getElementById("exportCsvBtn"),
  castsList: document.getElementById("castsList"),

  vehiclePlateNumber: document.getElementById("vehiclePlateNumber"),
  vehicleDriverName: document.getElementById("vehicleDriverName"),
  vehicleLineId: document.getElementById("vehicleLineId"),
  vehicleStatus: document.getElementById("vehicleStatus"),
  vehicleSeatCapacity: document.getElementById("vehicleSeatCapacity"),
  vehicleHomeArea: document.getElementById("vehicleHomeArea"),
  vehicleMemo: document.getElementById("vehicleMemo"),
  saveVehicleBtn: document.getElementById("saveVehicleBtn"),
  cancelVehicleEditBtn: document.getElementById("cancelVehicleEditBtn"),
  vehiclesList: document.getElementById("vehiclesList"),

  reportDate: document.getElementById("reportDate"),
  reportVehicleSelect: document.getElementById("reportVehicleSelect"),
  reportDriverName: document.getElementById("reportDriverName"),
  reportDistanceKm: document.getElementById("reportDistanceKm"),
  reportNote: document.getElementById("reportNote"),
  saveDailyReportBtn: document.getElementById("saveDailyReportBtn"),
  loadDailyReportsBtn: document.getElementById("loadDailyReportsBtn"),
  dailyReportsList: document.getElementById("dailyReportsList"),
  monthlyAverageList: document.getElementById("monthlyAverageList"),

  historyList: document.getElementById("historyList")
};

function todayStr() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeStatus(status) {
  if (status === "done") return "done";
  if (status === "cancel") return "cancel";
  if (status === "assigned") return "assigned";
  return "pending";
}

function toNullableNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function isValidLatLng(lat, lng) {
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return false;
  if (lat < -90 || lat > 90) return false;
  if (lng < -180 || lng > 180) return false;
  return true;
}

function formatDateTimeJa(value) {
  return new Date(value).toLocaleString("ja-JP");
}

function vehicleStatusLabel(status) {
  if (status === "running") return "稼働中";
  if (status === "maintenance") return "整備中";
  return "待機";
}

function vehicleStatusBadgeClass(status) {
  if (status === "running") return "done";
  if (status === "maintenance") return "cancel";
  return "";
}

function planStatusLabel(status) {
  if (status === "done") return "完了";
  if (status === "cancel") return "キャンセル";
  if (status === "assigned") return "配車済";
  return "予定";
}

function planStatusBadgeClass(status) {
  if (status === "done") return "done";
  if (status === "cancel") return "cancel";
  if (status === "assigned") return "assigned";
  return "";
}

function openRouteFromMatsudo(address) {
  const origin = encodeURIComponent(ORIGIN_LABEL);
  const dest = encodeURIComponent(address);
  const url = `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${dest}&travelmode=driving`;
  window.open(url, "_blank");
}

function openGoogleMapsForPin(address) {
  if (!address || !address.trim()) {
    alert("住所を入力してください");
    return;
  }
  const url =
    "https://www.google.com/maps/search/?api=1&query=" +
    encodeURIComponent(address.trim());
  window.open(url, "_blank");
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

function estimateMinutes(km) {
  return Math.max(8, Math.round(km * 3.2));
}

function normalizeAddressText(address) {
  return String(address || "")
    .trim()
    .replace(/[　\s]+/g, "")
    .replace(/ヶ/g, "ケ")
    .replace(/之/g, "の");
}

function getMonthKey(dateStr) {
  const d = new Date(dateStr);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function getDisplayArea(cast) {
  return String(cast?.area || "").trim();
}

/* =========================
   簡易地域推定
========================= */

function classifyAreaByAddress(address) {
  const a = normalizeAddressText(address);
  if (!a) return "";

  if (a.includes("船橋")) return "船橋";
  if (a.includes("柏")) return "柏";
  if (a.includes("松戸")) return "松戸";
  if (a.includes("市川")) return "市川";
  if (a.includes("流山")) return "流山";
  if (a.includes("八千代")) return "八千代";
  if (a.includes("三郷")) return "三郷";
  if (a.includes("守谷")) return "守谷";
  if (a.includes("葛飾")) return "葛飾";
  if (a.includes("江戸川")) return "江戸川";
  return "";
}

function classifyAreaByLatLng(lat, lng) {
  if (!isValidLatLng(lat, lng)) return "周辺";
  if (lng >= 139.99) return "東京東部";
  if (lat >= 35.84 && lng >= 139.94) return "柏";
  if (lat >= 35.81 && lat < 35.86 && lng >= 139.90 && lng < 139.96) return "松戸";
  if (lat >= 35.69 && lat < 35.80 && lng >= 139.93 && lng < 140.09) return "船橋";
  return "周辺";
}

function guessArea(lat, lng, address = "") {
  return classifyAreaByAddress(address) || classifyAreaByLatLng(lat, lng);
}

/* =========================
   AI配車
========================= */

function getVehicleAssignments(assignments, vehicleId) {
  return assignments.filter(x => Number(x.vehicle_id) === Number(vehicleId));
}

function getVehicleLoad(assignments, vehicleId) {
  return getVehicleAssignments(assignments, vehicleId).length;
}

function canAssignToVehicle(vehicle, assignments) {
  const capacity = Number(vehicle.seat_capacity || 4);
  const load = getVehicleLoad(assignments, vehicle.id);
  return load < capacity;
}

function getVehicleCurrentMaxDistance(assignments, vehicleId) {
  const rows = getVehicleAssignments(assignments, vehicleId);
  if (!rows.length) return 0;
  return Math.max(...rows.map(x => Number(x.distance_km || 0)));
}

function buildMonthlyDistanceMap(reports, targetMonth) {
  const map = new Map();

  reports.forEach(row => {
    const monthKey = getMonthKey(row.report_date);
    if (monthKey !== targetMonth) return;

    const vehicleId = Number(row.vehicle_id);
    const prev = map.get(vehicleId) || {
      worked_days: 0,
      total_distance_km: 0
    };

    prev.worked_days += 1;
    prev.total_distance_km += Number(row.distance_km || 0);
    prev.avg_distance_per_day = prev.total_distance_km / prev.worked_days;
    map.set(vehicleId, prev);
  });

  return map;
}

function getVehicleMonthlyAvgDistance(monthlyMap, vehicleId) {
  const row = monthlyMap.get(Number(vehicleId));
  if (!row) return 0;
  return Number(row.avg_distance_per_day || 0);
}

function chooseBestVehicleForCast(cast, vehicles, assignments, monthlyDistanceMap) {
  let bestVehicle = null;
  let bestScore = Infinity;

  vehicles.forEach(vehicle => {
    if (!canAssignToVehicle(vehicle, assignments)) return;

    const currentLoad = getVehicleLoad(assignments, vehicle.id);
    const maxDistance = getVehicleCurrentMaxDistance(assignments, vehicle.id);
    const avgDistance = getVehicleMonthlyAvgDistance(monthlyMap, vehicle.id);

    let score = 0;
    score += currentLoad * 20;
    score += maxDistance * 5;
    score += avgDistance * 0.15;
    score += Number(cast.distance_km || 0);

    if (score < bestScore) {
      bestScore = score;
      bestVehicle = vehicle;
    }
  });

  return bestVehicle;
}

function optimizeDispatchAssignmentsV2({ items, vehicles, monthlyDistanceMap }) {
  const activeVehicles = vehicles.filter(v =>
    v.status === "waiting" || v.status === "running"
  );

  const sorted = [...items].sort((a, b) => {
    const aBlock = String(a.destination_area || "無し");
    const bBlock = String(b.destination_area || "無し");
    if (aBlock !== bBlock) return aBlock.localeCompare(bBlock);
    return Number(a.distance_km || 0) - Number(b.distance_km || 0);
  });

  const assignments = [];

  sorted.forEach(item => {
    const vehicle = chooseBestVehicleForCast(item, activeVehicles, assignments, monthlyDistanceMap);
    if (!vehicle) return;

    assignments.push({
      item_id: item.id,
      vehicle_id: vehicle.id,
      driver_name: vehicle.driver_name || "",
      line_id: vehicle.line_id || ""
    });
  });

  return assignments;
}

/* =========================
   汎用
========================= */

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

function getUsedCastIdsInCurrentDispatch() {
  const ids = new Set();
  currentDispatchItemsCache.forEach(item => {
    if (item.cast_id) ids.add(Number(item.cast_id));
  });
  return ids;
}

function getPlannedOrBlockedCastIds() {
  const ids = new Set();
  currentPlansCache.forEach(plan => {
    if (!plan.cast_id) return;
    if (["planned", "assigned", "done", "cancel"].includes(plan.status)) {
      ids.add(Number(plan.cast_id));
    }
  });
  currentDispatchItemsCache.forEach(item => {
    if (item.cast_id && normalizeStatus(item.status) !== "cancel") {
      ids.add(Number(item.cast_id));
    }
  });
  return ids;
}

function buildLineDispatchText(items, vehicles) {
  const grouped = new Map();

  items
    .filter(item => normalizeStatus(item.status) !== "cancel")
    .sort((a, b) => {
      const aDriver = String(a.driver_name || "");
      const bDriver = String(b.driver_name || "");
      if (aDriver !== bDriver) return aDriver.localeCompare(bDriver);
      return Number(a.stop_order || 0) - Number(b.stop_order || 0);
    })
    .forEach(item => {
      const key = item.driver_name || "未設定";
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(item);
    });

  const lines = [];
  lines.push("【本日の送り配車】");
  lines.push("");

  for (const [driverName, rows] of grouped.entries()) {
    const vehicle = vehicles.find(v => (v.driver_name || "") === driverName);
    const lineId = vehicle?.line_id || "";
    lines.push(`${lineId ? lineId + " " : ""}${driverName}`);
    rows.forEach(row => {
      lines.push(`${row.casts?.name || "不明"} / ${row.destination_area || "-"}`);
    });
    lines.push("");
  }

  return lines.join("\n").trim();
}

async function copyLineDispatchText() {
  if (!currentDispatchItemsCache.length) {
    alert("配車データがありません");
    return;
  }

  const text = buildLineDispatchText(currentDispatchItemsCache, allVehiclesCache);

  try {
    await navigator.clipboard.writeText(text);
    alert("LINE貼り付け用の配車文をコピーしました");
  } catch (err) {
    console.error(err);
    alert("コピーに失敗しました");
  }
}

/* =========================
   認証
========================= */

async function ensureAuth() {
  const { data, error } = await supabaseClient.auth.getUser();

  if (error) {
    console.error(error);
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
    console.error("profiles upsert error:", profileError);
    alert("profiles作成エラー: " + profileError.message);
    return false;
  }

  els.userEmail.textContent = currentUser.email || "";
  return true;
}

async function logout() {
  await supabaseClient.auth.signOut();
  window.location.href = "index.html";
}

/* =========================
   キャスト管理
========================= */

function resetCastForm() {
  editingCastId = null;
  els.castName.value = "";
  els.castPhone.value = "";
  els.castAddress.value = "";
  els.castLatLngText.value = "";
  els.castArea.value = "";
  els.castLat.value = "";
  els.castLng.value = "";
  els.castMemo.value = "";
  els.saveCastBtn.textContent = "保存";
  els.cancelEditBtn.classList.add("hidden");
}

function fillCastForm(cast) {
  editingCastId = cast.id;
  els.castName.value = cast.name || "";
  els.castPhone.value = cast.phone || "";
  els.castAddress.value = cast.address || "";
  els.castLatLngText.value =
    cast.latitude != null && cast.longitude != null
      ? `${cast.latitude},${cast.longitude}`
      : "";
  els.castArea.value = getDisplayArea(cast) || "";
  els.castLat.value = cast.latitude ?? "";
  els.castLng.value = cast.longitude ?? "";
  els.castMemo.value = cast.memo || "";
  els.saveCastBtn.textContent = "更新";
  els.cancelEditBtn.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function applyLatLngFromText() {
  const parsed = parseLatLngText(els.castLatLngText.value);

  if (!parsed) {
    alert("座標の形式が正しくありません。例: 35.742891,139.876123");
    return;
  }

  els.castLat.value = parsed.lat;
  els.castLng.value = parsed.lng;

  if (!els.castArea.value.trim()) {
    els.castArea.value = guessArea(parsed.lat, parsed.lng, els.castAddress.value);
  }
}

async function saveCast() {
  const name = els.castName.value.trim();
  if (!name) {
    alert("名前を入力してください");
    return;
  }

  const lat = toNullableNumber(els.castLat.value);
  const lng = toNullableNumber(els.castLng.value);

  if ((lat !== null || lng !== null) && !isValidLatLng(lat, lng)) {
    alert("緯度経度が正しくありません");
    return;
  }

  const payload = {
    name,
    phone: els.castPhone.value.trim(),
    address: els.castAddress.value.trim(),
    area: els.castArea.value.trim() || guessArea(lat, lng, els.castAddress.value) || null,
    latitude: lat,
    longitude: lng,
    memo: els.castMemo.value.trim(),
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

  await addHistory(null, null, editingCastId ? "update_cast" : "create_cast", editingCastId ? "キャストを更新" : "キャストを作成");

  resetCastForm();
  await loadCasts();
  refreshPlanCastSelect();
  await loadHistory();
}

async function softDeleteCast(castId) {
  const ok = window.confirm("このキャストを削除しますか？");
  if (!ok) return;

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
  refreshPlanCastSelect();
  await loadDispatchItems(currentDispatchId);
  await loadHistory();
}

async function loadCasts() {
  const { data, error } = await supabaseClient
    .from("casts")
    .select("*")
    .eq("is_active", true)
    .order("id", { ascending: false });

  if (error) {
    console.error(error);
    return;
  }

  allCastsCache = data || [];
  renderCastSelect(allCastsCache);
  renderCastList(allCastsCache);
  refreshPlanCastSelect();
}

function renderCastSelect(casts) {
  els.castSelect.innerHTML = `<option value="">選択してください</option>`;
  const usedCastIds = getUsedCastIdsInCurrentDispatch();

  casts
    .filter(cast => !usedCastIds.has(Number(cast.id)))
    .forEach(cast => {
      const displayArea = getDisplayArea(cast);
      const option = document.createElement("option");
      option.value = cast.id;
      option.textContent = `${cast.name} | ${displayArea || "方面未設定"}`;
      option.dataset.address = cast.address || "";
      option.dataset.lat = cast.latitude ?? "";
      option.dataset.lng = cast.longitude ?? "";
      option.dataset.area = displayArea || "";
      els.castSelect.appendChild(option);
    });
}

function renderCastList(casts) {
  els.castsList.innerHTML = "";

  if (!casts.length) {
    els.castsList.innerHTML = `<div class="item"><p>キャストがまだ登録されていません。</p></div>`;
    return;
  }

  casts.forEach(cast => {
    const lat = toNullableNumber(cast.latitude);
    const lng = toNullableNumber(cast.longitude);
    const km = lat !== null && lng !== null ? estimateRoadKmFromStation(lat, lng) : null;
    const displayArea = getDisplayArea(cast);

    const div = document.createElement("div");
    div.className = "item";
    div.innerHTML = `
      <h3>${escapeHtml(cast.name)}</h3>
      <p>電話: ${escapeHtml(cast.phone || "-")}</p>
      <p>住所: ${escapeHtml(cast.address || "-")}</p>
      <p>方面: ${escapeHtml(displayArea || "-")}</p>
      <p>松戸駅から推定距離: ${km !== null ? `${km} km` : "-"}</p>
      <p>座標: ${cast.latitude ?? "-"}, ${cast.longitude ?? "-"}</p>
      <p>メモ: ${escapeHtml(cast.memo || "-")}</p>
      <div class="actions">
        <button class="btn secondary route-btn" data-address="${escapeHtml(cast.address || "")}">Googleマップ</button>
        <button class="btn edit-cast-btn" data-id="${cast.id}">編集</button>
        <button class="btn danger delete-cast-btn" data-id="${cast.id}">削除</button>
      </div>
    `;

    div.querySelector(".route-btn")?.addEventListener("click", e => {
      const address = e.currentTarget.dataset.address;
      if (address) openRouteFromMatsudo(address);
    });

    div.querySelector(".edit-cast-btn")?.addEventListener("click", () => fillCastForm(cast));
    div.querySelector(".delete-cast-btn")?.addEventListener("click", () => softDeleteCast(cast.id));

    els.castsList.appendChild(div);
  });
}

function exportCastsCsv() {
  const headers = ["name", "phone", "address", "area", "latitude", "longitude", "memo"];
  const rows = allCastsCache.map(cast => [
    cast.name || "",
    cast.phone || "",
    cast.address || "",
    cast.area || "",
    cast.latitude ?? "",
    cast.longitude ?? "",
    cast.memo || ""
  ]);

  const csv = [headers.join(","), ...rows.map(row => row.map(csvEscape).join(","))].join("\n");
  downloadTextFile(`casts_export_${todayStr()}.csv`, csv, "text/csv;charset=utf-8");
}

/* =========================
   車両管理
========================= */

function resetVehicleForm() {
  editingVehicleId = null;
  els.vehiclePlateNumber.value = "";
  els.vehicleDriverName.value = "";
  els.vehicleLineId.value = "";
  els.vehicleStatus.value = "waiting";
  els.vehicleSeatCapacity.value = "4";
  els.vehicleHomeArea.value = "";
  els.vehicleMemo.value = "";
  els.saveVehicleBtn.textContent = "保存";
  els.cancelVehicleEditBtn.classList.add("hidden");
}

function fillVehicleForm(vehicle) {
  editingVehicleId = vehicle.id;
  els.vehiclePlateNumber.value = vehicle.plate_number || "";
  els.vehicleDriverName.value = vehicle.driver_name || "";
  els.vehicleLineId.value = vehicle.line_id || "";
  els.vehicleStatus.value = vehicle.status || "waiting";
  els.vehicleSeatCapacity.value = String(vehicle.seat_capacity || 4);
  els.vehicleHomeArea.value = vehicle.home_area || "";
  els.vehicleMemo.value = vehicle.memo || "";
  els.saveVehicleBtn.textContent = "更新";
  els.cancelVehicleEditBtn.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function saveVehicle() {
  const plateNumber = els.vehiclePlateNumber.value.trim();
  if (!plateNumber) {
    alert("ナンバーを入力してください");
    return;
  }

  const payload = {
    plate_number: plateNumber,
    driver_name: els.vehicleDriverName.value.trim(),
    line_id: els.vehicleLineId.value.trim(),
    status: els.vehicleStatus.value,
    seat_capacity: Number(els.vehicleSeatCapacity.value || 4),
    home_area: els.vehicleHomeArea.value.trim(),
    memo: els.vehicleMemo.value.trim(),
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

  await addHistory(null, null, editingVehicleId ? "update_vehicle" : "create_vehicle", editingVehicleId ? "車両を更新" : "車両を登録");

  resetVehicleForm();
  await loadVehicles();
  await loadHistory();
}

async function softDeleteVehicle(vehicleId) {
  const ok = window.confirm("この車両を削除しますか？");
  if (!ok) return;

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
  await loadHistory();
}

async function loadVehicles() {
  const { data, error } = await supabaseClient
    .from("vehicles")
    .select("*")
    .eq("is_active", true)
    .order("id", { ascending: false });

  if (error) {
    console.error(error);
    return;
  }

  allVehiclesCache = data || [];
  renderVehicleList(allVehiclesCache);
  renderReportVehicleSelect(allVehiclesCache);
}

function renderReportVehicleSelect(vehicles) {
  els.reportVehicleSelect.innerHTML = `<option value="">選択してください</option>`;
  vehicles.forEach(vehicle => {
    const option = document.createElement("option");
    option.value = vehicle.id;
    option.textContent = `${vehicle.plate_number} | ${vehicle.driver_name || "-"}`;
    option.dataset.driver = vehicle.driver_name || "";
    els.reportVehicleSelect.appendChild(option);
  });
}

function renderVehicleList(vehicles) {
  els.vehiclesList.innerHTML = "";

  if (!vehicles.length) {
    els.vehiclesList.innerHTML = `<div class="item"><p>車両がまだ登録されていません。</p></div>`;
    return;
  }

  vehicles.forEach(vehicle => {
    const badgeClass = vehicleStatusBadgeClass(vehicle.status);

    const div = document.createElement("div");
    div.className = "item";
    div.innerHTML = `
      <h3>${escapeHtml(vehicle.plate_number)}</h3>
      <p>担当ドライバー: ${escapeHtml(vehicle.driver_name || "-")}</p>
      <p>LINE ID: ${escapeHtml(vehicle.line_id || "-")}</p>
      <p>状態: <span class="badge ${badgeClass}">${escapeHtml(vehicleStatusLabel(vehicle.status))}</span></p>
      <p>乗車可能人数: ${escapeHtml(vehicle.seat_capacity || 4)}人</p>
      <p>帰宅方面: ${escapeHtml(vehicle.home_area || "-")}</p>
      <p>メモ: ${escapeHtml(vehicle.memo || "-")}</p>
      <div class="actions">
        <button class="btn edit-vehicle-btn" data-id="${vehicle.id}">編集</button>
        <button class="btn danger delete-vehicle-btn" data-id="${vehicle.id}">削除</button>
      </div>
    `;

    div.querySelector(".edit-vehicle-btn")?.addEventListener("click", () => fillVehicleForm(vehicle));
    div.querySelector(".delete-vehicle-btn")?.addEventListener("click", () => softDeleteVehicle(vehicle.id));

    els.vehiclesList.appendChild(div);
  });
}

/* =========================
   予定表
========================= */

function resetPlanForm() {
  editingPlanId = null;
  els.planDate.value = els.planDate.value || todayStr();
  els.planHour.value = "0";
  els.planCastSelect.value = "";
  els.planArea.value = "";
  els.planNote.value = "";
  els.savePlanBtn.textContent = "予定保存";
  els.cancelPlanEditBtn.classList.add("hidden");
}

function refreshPlanCastSelect() {
  const blockedIds = getPlannedOrBlockedCastIds();
  const currentEditCastId = editingPlanId
    ? Number(currentPlansCache.find(x => Number(x.id) === Number(editingPlanId))?.cast_id || 0)
    : null;

  els.planCastSelect.innerHTML = `<option value="">選択してください</option>`;

  allCastsCache
    .filter(cast => !blockedIds.has(Number(cast.id)) || Number(cast.id) === currentEditCastId)
    .forEach(cast => {
      const option = document.createElement("option");
      option.value = cast.id;
      option.textContent = `${cast.name} | ${getDisplayArea(cast) || "方面未設定"}`;
      els.planCastSelect.appendChild(option);
    });
}

function fillPlanForm(plan) {
  editingPlanId = plan.id;
  els.planDate.value = plan.plan_date;
  els.planHour.value = String(plan.plan_hour);
  refreshPlanCastSelect();
  els.planCastSelect.value = String(plan.cast_id || "");
  els.planArea.value = plan.planned_area || "";
  els.planNote.value = plan.note || "";
  els.savePlanBtn.textContent = "予定更新";
  els.cancelPlanEditBtn.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function savePlan() {
  const planDate = els.planDate.value;
  const planHour = Number(els.planHour.value);
  const castId = Number(els.planCastSelect.value);

  if (!planDate) {
    alert("予定日を入力してください");
    return;
  }

  if (!Number.isFinite(planHour)) {
    alert("時間を選択してください");
    return;
  }

  if (!castId) {
    alert("キャストを選択してください");
    return;
  }

  if (!els.planArea.value) {
    alert("方面ブロックを選択してください");
    return;
  }

  const payload = {
    plan_date: planDate,
    plan_hour: planHour,
    cast_id: castId,
    planned_area: els.planArea.value,
    note: els.planNote.value.trim(),
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

  await addHistory(null, null, editingPlanId ? "update_plan" : "create_plan", editingPlanId ? "送り予定を更新" : "送り予定を作成");

  resetPlanForm();
  await loadPlansByDate(planDate);

  if (els.dispatchDate?.value === planDate) {
    renderPlanSelect(currentPlansCache);
  }

  await loadHistory();
}

async function updatePlanStatus(planId, status) {
  const { error } = await supabaseClient
    .from("dispatch_plans")
    .update({ status })
    .eq("id", planId);

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "update_plan_status", `送り予定の状態を ${status} に変更`);
  await loadPlansByDate(els.planDate.value || todayStr());
  await loadHistory();
}

async function deletePlan(planId) {
  const ok = window.confirm("この予定を削除しますか？");
  if (!ok) return;

  const { error } = await supabaseClient
    .from("dispatch_plans")
    .delete()
    .eq("id", planId);

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "delete_plan", `送り予定ID ${planId} を削除`);
  await loadPlansByDate(els.planDate.value || todayStr());
  await loadHistory();
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
  renderPlansSummary(currentPlansCache, dateStr);
  renderPlansBoard(currentPlansCache, dateStr);
  renderPlanSelect(currentPlansCache);
  refreshPlanCastSelect();
}

function renderPlanSelect(plans) {
  const dispatchDate = els.dispatchDate?.value || "";

  els.planSelect.innerHTML = `<option value="">予定を選択してください</option>`;

  plans
    .filter(plan => plan.status === "planned")
    .filter(plan => !dispatchDate || plan.plan_date === dispatchDate)
    .forEach(plan => {
      const hourLabel = BOARD_ROWS.find(r => r.key === Number(plan.plan_hour))?.label || plan.plan_hour;
      const option = document.createElement("option");
      option.value = plan.id;
      option.textContent = `${hourLabel} | ${plan.planned_area || "-"} | ${plan.casts?.name || "不明"}`;
      els.planSelect.appendChild(option);
    });
}

function renderPlansSummary(plans, dateStr) {
  if (!plans.length) {
    els.plansSummary.textContent = `${dateStr} の予定はまだありません。`;
    return;
  }

  const byStatus = {
    planned: plans.filter(x => x.status === "planned").length,
    assigned: plans.filter(x => x.status === "assigned").length,
    done: plans.filter(x => x.status === "done").length,
    cancel: plans.filter(x => x.status === "cancel").length
  };

  els.plansSummary.textContent =
    `${dateStr} / 予定 ${byStatus.planned} 件 / 配車済 ${byStatus.assigned} 件 / 完了 ${byStatus.done} 件 / キャンセル ${byStatus.cancel} 件`;
}

function buildPlanCardHtml(plan) {
  const status = normalizeStatus(plan.status);
  const statusLabel = planStatusLabel(plan.status);

  return `
    <div class="plan-name-card ${status}">
      <div class="plan-name">${escapeHtml(plan.casts?.name || "不明")}</div>
      <div class="plan-mini-status ${status}">${escapeHtml(statusLabel)}</div>
      <div class="mini-actions">
        <button class="btn edit-plan-btn" data-id="${plan.id}">編集</button>
        <button class="btn secondary plan-done-btn" data-id="${plan.id}">完了</button>
        <button class="btn danger plan-cancel-btn" data-id="${plan.id}">キャンセル</button>
      </div>
    </div>
  `;
}

function renderPlansBoard(plans, dateStr) {
  els.plansBoard.innerHTML = "";

  const wrapper = document.createElement("div");
  wrapper.className = "item";

  let html = `
    <div style="overflow:auto;">
      <table>
        <thead>
          <tr>
            <th>時間</th>
            ${BLOCKS.map(block => `<th>${block}</th>`).join("")}
          </tr>
        </thead>
        <tbody>
  `;

  BOARD_ROWS.forEach(row => {
    html += `<tr>`;
    html += `<td>${row.label}</td>`;

    BLOCKS.forEach(block => {
      const cellPlans = plans.filter(
        p => Number(p.plan_hour) === row.key && String(p.planned_area || "") === block
      );

      const cellHtml = cellPlans.length
        ? `<div class="plan-cell">${cellPlans.map(buildPlanCardHtml).join("")}</div>`
        : `<div class="plan-cell"></div>`;

      html += `<td>${cellHtml}</td>`;
    });

    html += `</tr>`;
  });

  html += `
        </tbody>
      </table>
    </div>
  `;

  wrapper.innerHTML = html;
  els.plansBoard.appendChild(wrapper);

  wrapper.querySelectorAll(".edit-plan-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const plan = currentPlansCache.find(x => Number(x.id) === Number(btn.dataset.id));
      if (plan) fillPlanForm(plan);
    });
  });

  wrapper.querySelectorAll(".plan-done-btn").forEach(btn => {
    btn.addEventListener("click", async () => {
      await updatePlanStatus(Number(btn.dataset.id), "done");
    });
  });

  wrapper.querySelectorAll(".plan-cancel-btn").forEach(btn => {
    btn.addEventListener("click", async () => {
      await updatePlanStatus(Number(btn.dataset.id), "cancel");
    });
  });
}

/* =========================
   配車管理
========================= */

async function createDispatch() {
  const dispatchDate = els.dispatchDate.value;
  const driverName = els.driverName.value.trim();

  if (!dispatchDate) {
    alert("日付を選択してください");
    return;
  }

  const { data: existing, error: existingError } = await supabaseClient
    .from("dispatches")
    .select("*")
    .eq("dispatch_date", dispatchDate)
    .order("id", { ascending: false })
    .limit(1);

  if (existingError) {
    alert(existingError.message);
    return;
  }

  if (existing?.length) {
    currentDispatchId = existing[0].id;
    els.driverName.value = existing[0].driver_name || "";

    if (els.planDate) {
      els.planDate.value = dispatchDate;
    }

    await loadDispatchItems(currentDispatchId);
    await loadPlansByDate(dispatchDate);
    return;
  }

  const { data, error } = await supabaseClient
    .from("dispatches")
    .insert({
      dispatch_date: dispatchDate,
      driver_name: driverName,
      status: "draft",
      created_by: currentUser.id
    })
    .select()
    .single();

  if (error) {
    alert(error.message);
    return;
  }

  currentDispatchId = data.id;

  if (els.planDate) {
    els.planDate.value = dispatchDate;
  }

  await addHistory(currentDispatchId, null, "create_dispatch", `配車を作成: ${dispatchDate}`);
  await loadDispatchItems(currentDispatchId);
  await loadPlansByDate(dispatchDate);
  await loadHistory();
}

async function loadDispatchByDate(dateStr) {
  const { data, error } = await supabaseClient
    .from("dispatches")
    .select("*")
    .eq("dispatch_date", dateStr)
    .order("id", { ascending: false })
    .limit(1);

  if (error) {
    console.error(error);
    return;
  }

  const dispatch = data?.[0];
  currentDispatchId = dispatch?.id || null;

  if (dispatch) {
    els.driverName.value = dispatch.driver_name || "";
    await loadDispatchItems(dispatch.id);
  } else {
    currentDispatchItemsCache = [];
    renderCastSelect(allCastsCache);
    els.dispatchSummary.textContent = "";
    els.dispatchList.innerHTML = `<div class="item"><p>この日の配車はまだありません。</p></div>`;
  }

  if (els.planDate) {
    els.planDate.value = dateStr;
  }

  await loadPlansByDate(dateStr);
}

async function addCastToDispatch() {
  if (!currentDispatchId) {
    alert("先に配車を作成または読込してください");
    return;
  }

  const castId = Number(els.castSelect.value);
  if (!castId) {
    alert("キャストを選択してください");
    return;
  }

  if (currentDispatchItemsCache.some(item => Number(item.cast_id) === castId)) {
    alert("このキャストはすでに配車に追加されています");
    return;
  }

  const selected = els.castSelect.selectedOptions[0];
  const destinationAddress = selected?.dataset.address || "";
  const lat = toNullableNumber(selected?.dataset.lat);
  const lng = toNullableNumber(selected?.dataset.lng);
  const displayArea = selected?.dataset.area || "";
  const distanceKm = lat !== null && lng !== null ? estimateRoadKmFromStation(lat, lng) : null;
  const travelMinutes = distanceKm !== null ? estimateMinutes(distanceKm) : null;

  let stopOrder = Number(els.stopOrder.value || 1);
  if (!Number.isFinite(stopOrder) || stopOrder <= 0) {
    stopOrder = currentDispatchItemsCache.length + 1;
  }

  const { data, error } = await supabaseClient
    .from("dispatch_items")
    .insert({
      dispatch_id: currentDispatchId,
      cast_id: castId,
      stop_order: stopOrder,
      pickup_label: ORIGIN_LABEL,
      destination_address: destinationAddress,
      destination_area: displayArea || "無し",
      latitude: lat,
      longitude: lng,
      distance_km: distanceKm,
      travel_minutes: travelMinutes,
      plan_date: els.dispatchDate.value || null,
      status: "pending"
    })
    .select()
    .single();

  if (error) {
    alert(error.message);
    return;
  }

  await normalizeStopOrders(currentDispatchId);
  await addHistory(currentDispatchId, data.id, "add_cast", "予定外キャストを配車に追加");
  await loadDispatchItems(currentDispatchId);
  await loadHistory();
}

async function addPlanToDispatch() {
  if (!currentDispatchId) {
    alert("先に配車を作成または読込してください");
    return;
  }

  await loadPlansByDate(els.dispatchDate.value || todayStr());

  const planId = Number(els.planSelect.value);
  if (!planId) {
    alert("予定を選択してください");
    return;
  }

  const plan = currentPlansCache.find(x => Number(x.id) === Number(planId));
  if (!plan) {
    alert("予定が見つかりません");
    return;
  }

  if (plan.status !== "planned") {
    alert("この予定は追加できません");
    return;
  }

  const cast = plan.casts;
  if (!cast) {
    alert("予定のキャスト情報がありません");
    return;
  }

  if (currentDispatchItemsCache.some(item => Number(item.cast_id) === Number(cast.id))) {
    alert("このキャストはすでに配車済みです");
    return;
  }

  const lat = toNullableNumber(cast.latitude);
  const lng = toNullableNumber(cast.longitude);
  const distanceKm = lat !== null && lng !== null ? estimateRoadKmFromStation(lat, lng) : null;
  const travelMinutes = distanceKm !== null ? estimateMinutes(distanceKm) : null;
  const stopOrder = currentDispatchItemsCache.length + 1;

  const { data, error } = await supabaseClient
    .from("dispatch_items")
    .insert({
      dispatch_id: currentDispatchId,
      cast_id: cast.id,
      stop_order: stopOrder,
      pickup_label: ORIGIN_LABEL,
      destination_address: cast.address || "",
      destination_area: plan.planned_area || "無し",
      latitude: lat,
      longitude: lng,
      distance_km: distanceKm,
      travel_minutes: travelMinutes,
      plan_hour: plan.plan_hour,
      plan_date: plan.plan_date,
      status: "pending"
    })
    .select()
    .single();

  if (error) {
    alert(error.message);
    return;
  }

  await supabaseClient
    .from("dispatch_plans")
    .update({ status: "assigned" })
    .eq("id", plan.id);

  await addHistory(currentDispatchId, data.id, "add_plan_to_dispatch", `予定ID ${plan.id} を配車に追加`);
  await loadDispatchItems(currentDispatchId);
  await loadPlansByDate(plan.plan_date);
  await loadHistory();
}

async function loadDispatchItems(dispatchId) {
  if (!dispatchId) {
    currentDispatchItemsCache = [];
    renderDispatchSummary([]);
    renderDispatchItems([]);
    renderCastSelect(allCastsCache);
    return;
  }

  const { data, error } = await supabaseClient
    .from("dispatch_items")
    .select(`
      *,
      casts (
        id,
        name,
        phone,
        area,
        address,
        latitude,
        longitude
      )
    `)
    .eq("dispatch_id", dispatchId)
    .order("stop_order", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  currentDispatchItemsCache = data || [];
  renderDispatchSummary(currentDispatchItemsCache);
  renderDispatchItems(currentDispatchItemsCache);
  renderCastSelect(allCastsCache);

  els.stopOrder.value = String(currentDispatchItemsCache.length + 1);
}

function renderDispatchSummary(items) {
  if (!items.length) {
    els.dispatchSummary.textContent = "配車先はまだありません。";
    return;
  }

  const active = items.filter(x => normalizeStatus(x.status) === "pending");
  const done = items.filter(x => normalizeStatus(x.status) === "done");
  const cancel = items.filter(x => normalizeStatus(x.status) === "cancel");

  els.dispatchSummary.textContent =
    `未完了 ${active.length} 件 / 完了 ${done.length} 件 / キャンセル ${cancel.length} 件`;
}

function renderDispatchItems(items) {
  els.dispatchList.innerHTML = "";

  if (!items.length) {
    els.dispatchList.innerHTML = `<div class="item"><p>配車先はまだありません。</p></div>`;
    return;
  }

  const grouped = new Map();

  items.forEach(item => {
    const key = item.driver_name || "未割当";
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(item);
  });

  [...grouped.entries()].forEach(([driverName, list]) => {
    const vehicle = allVehiclesCache.find(v => (v.driver_name || "") === driverName);
    const lineId = vehicle?.line_id || "";
    const wrapper = document.createElement("div");
    wrapper.className = "item";

    let html = `
      <h3>${escapeHtml(lineId ? `${lineId} ${driverName}` : driverName)}</h3>
    `;

    list.forEach(item => {
      const status = normalizeStatus(item.status);
      const badgeClass = status === "done" ? "done" : status === "cancel" ? "cancel" : "";
      const cardClass =
        status === "done" ? "done-card" :
        status === "cancel" ? "cancel-card" :
        "";

      html += `
        <div class="item ${cardClass}" style="margin-top:8px;">
          <p><strong>${escapeHtml(item.casts?.name || "不明")}</strong> / ${escapeHtml(item.destination_area || "-")}</p>
          <p><span class="badge ${badgeClass}">${escapeHtml(status)}</span></p>
          <div class="actions">
            <button class="btn done-btn" data-id="${item.id}">完了</button>
            <button class="btn danger cancel-btn" data-id="${item.id}">キャンセル</button>
            <button class="btn danger delete-item-btn" data-id="${item.id}">削除</button>
          </div>
        </div>
      `;
    });

    wrapper.innerHTML = html;

    wrapper.querySelectorAll(".done-btn").forEach(btn => {
      btn.addEventListener("click", async e => {
        await updateItemStatus(Number(e.currentTarget.dataset.id), "done");
      });
    });

    wrapper.querySelectorAll(".cancel-btn").forEach(btn => {
      btn.addEventListener("click", async e => {
        await updateItemStatus(Number(e.currentTarget.dataset.id), "cancel");
      });
    });

    wrapper.querySelectorAll(".delete-item-btn").forEach(btn => {
      btn.addEventListener("click", async e => {
        await deleteDispatchItem(Number(e.currentTarget.dataset.id));
      });
    });

    els.dispatchList.appendChild(wrapper);
  });
}

async function updateItemStatus(itemId, status) {
  const item = currentDispatchItemsCache.find(x => Number(x.id) === Number(itemId));

  const { error } = await supabaseClient
    .from("dispatch_items")
    .update({ status })
    .eq("id", itemId);

  if (error) {
    alert(error.message);
    return;
  }

  if (item?.plan_date && item?.cast_id) {
    const relatedPlan = currentPlansCache.find(
      x => Number(x.cast_id) === Number(item.cast_id) && x.plan_date === item.plan_date
    );
    if (relatedPlan) {
      await supabaseClient
        .from("dispatch_plans")
        .update({ status })
        .eq("id", relatedPlan.id);
    }
  }

  await addHistory(currentDispatchId, itemId, "update_status", `状態を ${status} に変更`);
  await loadDispatchItems(currentDispatchId);
  await loadPlansByDate(els.planDate.value || todayStr());
  await loadHistory();
}

async function deleteDispatchItem(itemId) {
  const ok = window.confirm("この配車項目を削除しますか？");
  if (!ok) return;

  const item = currentDispatchItemsCache.find(x => Number(x.id) === Number(itemId));

  const { error } = await supabaseClient
    .from("dispatch_items")
    .delete()
    .eq("id", itemId);

  if (error) {
    alert(error.message);
    return;
  }

  if (item?.plan_date && item?.cast_id) {
    const relatedPlan = currentPlansCache.find(
      x => Number(x.cast_id) === Number(item.cast_id) && x.plan_date === item.plan_date
    );
    if (relatedPlan && relatedPlan.status === "assigned") {
      await supabaseClient
        .from("dispatch_plans")
        .update({ status: "planned" })
        .eq("id", relatedPlan.id);
    }
  }

  await normalizeStopOrders(currentDispatchId);
  await addHistory(currentDispatchId, itemId, "delete_dispatch_item", "配車項目を削除");
  await loadDispatchItems(currentDispatchId);
  await loadPlansByDate(els.planDate.value || todayStr());
  await loadHistory();
}

async function normalizeStopOrders(dispatchId) {
  if (!dispatchId) return;

  const { data, error } = await supabaseClient
    .from("dispatch_items")
    .select("id, stop_order")
    .eq("dispatch_id", dispatchId)
    .order("stop_order", { ascending: true })
    .order("id", { ascending: true });

  if (error || !data) return;

  for (let i = 0; i < data.length; i++) {
    const desiredOrder = i + 1;
    if (Number(data[i].stop_order) === desiredOrder) continue;

    await supabaseClient
      .from("dispatch_items")
      .update({ stop_order: desiredOrder })
      .eq("id", data[i].id);
  }
}

async function runOptimize() {
  if (!currentDispatchId) {
    alert("先に配車を読み込んでください");
    return;
  }

  const activeItems = currentDispatchItemsCache.filter(x => normalizeStatus(x.status) === "pending");
  if (!activeItems.length) {
    alert("未完了の配車がありません");
    return;
  }

  const dateStr = els.dispatchDate.value || todayStr();
  const monthlyDistanceMap = buildMonthlyDistanceMap(currentDailyReportsCache, getMonthKey(dateStr));

  const aiResult = optimizeDispatchAssignmentsV2({
    items: activeItems,
    vehicles: allVehiclesCache,
    monthlyDistanceMap
  });

  for (let i = 0; i < aiResult.length; i++) {
    const result = aiResult[i];
    await supabaseClient
      .from("dispatch_items")
      .update({
        stop_order: i + 1,
        vehicle_id: result.vehicle_id,
        driver_name: result.driver_name
      })
      .eq("id", result.item_id);
  }

  await addHistory(currentDispatchId, null, "optimize", "AI配車を実行");
  await loadDispatchItems(currentDispatchId);
  await loadHistory();
}

/* =========================
   走行管理
========================= */

async function saveDailyReport() {
  const reportDate = els.reportDate.value;
  const vehicleId = Number(els.reportVehicleSelect.value);
  const distanceKm = Number(els.reportDistanceKm.value);
  const driverName = els.reportDriverName.value.trim();

  if (!reportDate) {
    alert("報告日を入力してください");
    return;
  }
  if (!vehicleId) {
    alert("車両を選択してください");
    return;
  }
  if (!Number.isFinite(distanceKm) || distanceKm < 0) {
    alert("走行距離を正しく入力してください");
    return;
  }

  const { error } = await supabaseClient
    .from("vehicle_daily_reports")
    .insert({
      report_date: reportDate,
      vehicle_id: vehicleId,
      driver_name: driverName || null,
      distance_km: distanceKm,
      note: els.reportNote.value.trim(),
      created_by: currentUser.id
    });

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "save_daily_report", `走行距離を保存: ${distanceKm}km`);
  els.reportDistanceKm.value = "";
  els.reportNote.value = "";
  await loadDailyReports(reportDate);
  await loadHistory();
}

async function loadDailyReports(reportDate) {
  const { data, error } = await supabaseClient
    .from("vehicle_daily_reports")
    .select(`
      *,
      vehicles (
        id,
        plate_number
      )
    `)
    .eq("report_date", reportDate)
    .order("id", { ascending: false });

  if (error) {
    console.error(error);
    return;
  }

  currentDailyReportsCache = data || [];
  renderDailyReports(currentDailyReportsCache);
  await loadMonthlyAverages(reportDate);
}

function renderDailyReports(reports) {
  els.dailyReportsList.innerHTML = "";

  if (!reports.length) {
    els.dailyReportsList.innerHTML = `<div class="item"><p>本日の報告はまだありません。</p></div>`;
    return;
  }

  reports.forEach(row => {
    const div = document.createElement("div");
    div.className = "item";
    div.innerHTML = `
      <h3>${escapeHtml(row.vehicles?.plate_number || "-")}</h3>
      <p>ドライバー: ${escapeHtml(row.driver_name || "-")}</p>
      <p>走行距離: ${escapeHtml(row.distance_km)} km</p>
      <p>メモ: ${escapeHtml(row.note || "-")}</p>
      <p>${formatDateTimeJa(row.created_at)}</p>
    `;
    els.dailyReportsList.appendChild(div);
  });
}

async function loadMonthlyAverages(reportDate) {
  const monthKey = getMonthKey(reportDate);
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
        plate_number
      )
    `)
    .gte("report_date", monthStart)
    .lt("report_date", nextStr)
    .order("report_date", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  const grouped = new Map();
  (data || []).forEach(row => {
    const id = Number(row.vehicle_id);
    const prev = grouped.get(id) || {
      vehicle_label: `${row.vehicles?.plate_number || "-"}`,
      worked_days: 0,
      total_distance_km: 0
    };
    prev.worked_days += 1;
    prev.total_distance_km += Number(row.distance_km || 0);
    grouped.set(id, prev);
  });

  const rows = [...grouped.entries()]
    .map(([vehicleId, value]) => ({
      vehicle_id: vehicleId,
      vehicle_label: value.vehicle_label,
      worked_days: value.worked_days,
      total_distance_km: value.total_distance_km,
      avg_distance_per_day: value.worked_days > 0 ? value.total_distance_km / value.worked_days : 0
    }))
    .sort((a, b) => a.avg_distance_per_day - b.avg_distance_per_day);

  renderMonthlyAverages(rows);
}

function renderMonthlyAverages(rows) {
  els.monthlyAverageList.innerHTML = "";

  if (!rows.length) {
    els.monthlyAverageList.innerHTML = `<div class="item"><p>今月の報告はまだありません。</p></div>`;
    return;
  }

  rows.forEach(row => {
    const div = document.createElement("div");
    div.className = "item";
    div.innerHTML = `
      <h3>${escapeHtml(row.vehicle_label)}</h3>
      <p>出勤日数: ${escapeHtml(row.worked_days)} 日</p>
      <p>月間合計距離: ${escapeHtml(row.total_distance_km.toFixed(1))} km</p>
      <p>1日平均距離: ${escapeHtml(row.avg_distance_per_day.toFixed(1))} km</p>
    `;
    els.monthlyAverageList.appendChild(div);
  });
}

/* =========================
   履歴
========================= */

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

  if (error) {
    console.error(error);
  }
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

  els.historyList.innerHTML = "";

  if (!data?.length) {
    els.historyList.innerHTML = `<div class="item"><p>履歴はまだありません。</p></div>`;
    return;
  }

  data.forEach(row => {
    const div = document.createElement("div");
    div.className = "item";
    div.innerHTML = `
      <h3>${escapeHtml(row.action)}</h3>
      <p>${escapeHtml(row.message || "")}</p>
      <p>${formatDateTimeJa(row.created_at)}</p>
    `;
    els.historyList.appendChild(div);
  });
}

/* =========================
   CSV
========================= */

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
  const h = String(header || "").trim().toLowerCase();
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
    note: "memo"
  };
  return map[header] || map[h] || h;
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

async function importCsv() {
  const file = els.csvFileInput.files?.[0];
  if (!file) {
    alert("CSVファイルを選択してください");
    return;
  }

  const text = await file.text();
  const rows = parseCsv(text);

  if (!rows.length) {
    alert("CSVデータが空です");
    return;
  }

  const existingNames = new Set(allCastsCache.map(c => String(c.name || "").trim().toLowerCase()));

  const payload = rows
    .map(row => {
      const lat = toNullableNumber(row.latitude);
      const lng = toNullableNumber(row.longitude);

      return {
        name: String(row.name || "").trim(),
        phone: String(row.phone || "").trim(),
        address: String(row.address || "").trim(),
        area: String(row.area || "").trim() || guessArea(lat, lng, row.address) || null,
        latitude: lat,
        longitude: lng,
        memo: String(row.memo || "").trim(),
        is_active: true,
        created_by: currentUser.id
      };
    })
    .filter(x => x.name)
    .filter(x => !existingNames.has(x.name.toLowerCase()));

  if (!payload.length) {
    alert("有効なname列がないか、すべて既存キャストと重複しています");
    return;
  }

  const { error } = await supabaseClient.from("casts").insert(payload);

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "import_csv", `${payload.length}件のキャストをCSV取込`);
  els.csvFileInput.value = "";
  await loadCasts();
  await loadHistory();
  alert(`${payload.length}件 取り込みました`);
}

/* =========================
   UI
========================= */

function setupTabs() {
  const tabs = document.querySelectorAll(".tab");
  const panels = document.querySelectorAll(".tab-panel");

  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      tabs.forEach(x => x.classList.remove("active"));
      panels.forEach(x => x.classList.remove("active"));
      tab.classList.add("active");
      const panel = document.getElementById(tab.dataset.tab);
      if (panel) panel.classList.add("active");
    });
  });
}

function setupEvents() {
  els.logoutBtn?.addEventListener("click", logout);

  els.openGoogleMapBtn?.addEventListener("click", () => openGoogleMapsForPin(els.castAddress.value));
  els.applyLatLngBtn?.addEventListener("click", applyLatLngFromText);
  els.saveCastBtn?.addEventListener("click", saveCast);
  els.cancelEditBtn?.addEventListener("click", resetCastForm);
  els.importCsvBtn?.addEventListener("click", importCsv);
  els.exportCsvBtn?.addEventListener("click", exportCastsCsv);

  els.saveVehicleBtn?.addEventListener("click", saveVehicle);
  els.cancelVehicleEditBtn?.addEventListener("click", resetVehicleForm);

  els.savePlanBtn?.addEventListener("click", savePlan);
  els.cancelPlanEditBtn?.addEventListener("click", resetPlanForm);
  els.loadPlansBtn?.addEventListener("click", async () => {
    await loadPlansByDate(els.planDate.value || todayStr());
  });

  els.planDate?.addEventListener("change", async () => {
    await loadPlansByDate(els.planDate.value || todayStr());
  });

  els.createDispatchBtn?.addEventListener("click", createDispatch);
  els.loadDispatchBtn?.addEventListener("click", async () => {
    await loadDispatchByDate(els.dispatchDate.value);
  });
  els.addCastBtn?.addEventListener("click", addCastToDispatch);
  els.addPlanToDispatchBtn?.addEventListener("click", addPlanToDispatch);
  els.optimizeBtn?.addEventListener("click", runOptimize);
  els.copyLineDispatchBtn?.addEventListener("click", copyLineDispatchText);

  els.dispatchDate?.addEventListener("change", async () => {
    if (!els.dispatchDate.value) return;
    if (els.planDate) els.planDate.value = els.dispatchDate.value;
    await loadDispatchByDate(els.dispatchDate.value);
  });

  els.castLatLngText?.addEventListener("blur", () => {
    const parsed = parseLatLngText(els.castLatLngText.value);
    if (!parsed) return;
    els.castLat.value = parsed.lat;
    els.castLng.value = parsed.lng;
    if (!els.castArea.value.trim()) {
      els.castArea.value = guessArea(parsed.lat, parsed.lng, els.castAddress.value);
    }
  });

  els.reportVehicleSelect?.addEventListener("change", () => {
    const selected = els.reportVehicleSelect.selectedOptions?.[0];
    if (!els.reportDriverName.value.trim()) {
      els.reportDriverName.value = selected?.dataset.driver || "";
    }
  });

  els.saveDailyReportBtn?.addEventListener("click", saveDailyReport);
  els.loadDailyReportsBtn?.addEventListener("click", async () => {
    await loadDailyReports(els.reportDate.value || todayStr());
  });
}

/* =========================
   起動
========================= */

document.addEventListener("DOMContentLoaded", async () => {
  try {
    const ok = await ensureAuth();
    if (!ok) return;

    setupTabs();
    setupEvents();

    resetCastForm();
    resetVehicleForm();
    resetPlanForm();

    const today = todayStr();
    if (els.dispatchDate) els.dispatchDate.value = today;
    if (els.planDate) els.planDate.value = today;
    if (els.reportDate) els.reportDate.value = today;

    await loadCasts();
    await loadVehicles();
    await loadDispatchByDate(today);
    await loadPlansByDate(today);
    await loadDailyReports(today);
    await loadHistory();

    refreshPlanCastSelect();
  } catch (err) {
    console.error("dashboard init error:", err);
    alert("初期化中にエラーが発生しました。F12 の Console を確認してください。");
  }
});
