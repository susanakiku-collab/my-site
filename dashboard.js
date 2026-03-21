
// ===== direction guard (auto inserted) =====
function __shouldBlockDirectionMerge(items, vehicles){
  try{
    const areas = new Set((items||[]).map(i=>{
      return (i.destination_area||i.cluster_area||i.area||'').toString().trim();
    }).filter(Boolean));
    return (vehicles?.length || 0) >= areas.size;
  }catch(e){return false;}
}
// ==========================================

// THEMIS AI Dispatch v6.9.12
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
let simulationSlotHour = null;
let lastSimulationResult = null;
let isRefreshingHybridUI = false;
let suppressSimulationSlotChange = false;
const ENABLE_DISTANCE_REBALANCE = false;
const ENFORCE_AREA_PRIORITY_STRICT = true;
const ENABLE_DISPLAY_GROUP_FORCE_BRANCH = false;

const els = {
  plansTimeAreaMatrix: document.getElementById("plansTimeAreaMatrix"),
  userEmail: document.getElementById("userEmail"),
  sendLineBtn: document.getElementById("sendLineBtn"),
  originLabelText: document.getElementById("originLabelText"),
  logoutBtn: document.getElementById("logoutBtn"),

  exportAllBtn: document.getElementById("exportAllBtn"),
  importAllBtn: document.getElementById("importAllBtn"),
  importAllFileInput: document.getElementById("importAllFileInput"),
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
  castTravelMinutes: document.getElementById("castTravelMinutes"),
  fetchCastTravelMinutesBtn: document.getElementById("fetchCastTravelMinutesBtn"),
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
  operationContextSummary: document.getElementById("operationContextSummary"),
  operationDiagnosis: document.getElementById("operationDiagnosis"),
  simulationSlotSelect: document.getElementById("simulationSlotSelect"),
  simulationIncludePlanInflow: document.getElementById("simulationIncludePlanInflow"),
  runSimulationBtn: document.getElementById("runSimulationBtn"),
  runSimulationDispatchBtn: document.getElementById("runSimulationDispatchBtn"),
  simulationDiagnosis: document.getElementById("simulationDiagnosis"),
  simulationPreview: document.getElementById("simulationPreview"),

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
  const checkedNames = getDriverLastTripCheckedNames();
  const extra = checkedNames.length ? ` / ラスト便チェック: ${checkedNames.join("・")}` : "";
  if (els.manualLastVehicleInfo) els.manualLastVehicleInfo.textContent = text + extra;
}function clearManualLastVehicle() {
  clearManualLastVehicleState();
  renderManualLastVehicleInfo();
  renderDailyDispatchResult();
  alert("ラスト便車両を解除しました");
}

function isManualLastVehicle(vehicleId, dateStr = getCurrentDispatchDateStr()) {
  return Number(vehicleId) === Number(getManualLastVehicleId(dateStr));
}

function getDriverLastTripStorageKey(dateStr = getCurrentDispatchDateStr()) {
  return `THEMIS_DRIVER_LAST_TRIP_FLAGS_${dateStr}`;
}

function getDriverLastTripState(dateStr = getCurrentDispatchDateStr()) {
  try {
    const raw = window.localStorage.getItem(getDriverLastTripStorageKey(dateStr));
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (error) {
    console.error(error);
    return {};
  }
}

function saveDriverLastTripState(state, dateStr = getCurrentDispatchDateStr()) {
  try {
    window.localStorage.setItem(
      getDriverLastTripStorageKey(dateStr),
      JSON.stringify(state || {})
    );
  } catch (error) {
    console.error(error);
  }
}

function isDriverLastTripChecked(vehicleId, dateStr = getCurrentDispatchDateStr()) {
  const state = getDriverLastTripState(dateStr);
  return Boolean(state[String(Number(vehicleId || 0))]);
}

function hasAnyDriverLastTripChecked(dateStr = getCurrentDispatchDateStr()) {
  const state = getDriverLastTripState(dateStr);
  return Object.values(state).some(Boolean);
}

function setDriverLastTripChecked(vehicleId, checked, dateStr = getCurrentDispatchDateStr()) {
  const key = String(Number(vehicleId || 0));
  const state = getDriverLastTripState(dateStr);

  if (checked) state[key] = true;
  else delete state[key];

  saveDriverLastTripState(state, dateStr);
}

function getDriverLastTripCheckedNames(dateStr = getCurrentDispatchDateStr()) {
  const state = getDriverLastTripState(dateStr);
  return allVehiclesCache
    .filter(vehicle => Boolean(state[String(Number(vehicle.id || 0))]))
    .map(vehicle => vehicle.driver_name || vehicle.plate_number || "-");
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
}function getDefaultLastHour(dateStr) {
  const day = getDayOfWeek(dateStr);
  // 火曜-金曜は4時、土曜-月曜は5時
  if (day >= 2 && day <= 5) return 4;
  return 5;
}

function getEffectiveLastHour(dateStr, rows = []) {
  const rowHours = (Array.isArray(rows) ? rows : [])
    .map(row => Number(row?.actual_hour ?? row?.plan_hour ?? row?.hour))
    .filter(Number.isFinite);
  if (rowHours.length) return Math.max(...rowHours);

  const cacheHours = [];
  (Array.isArray(currentActualsCache) ? currentActualsCache : []).forEach(row => {
    const status = normalizeStatus(row?.status);
    if (status === "done" || status === "cancel") return;
    const hour = Number(row?.actual_hour ?? row?.plan_hour);
    if (Number.isFinite(hour)) cacheHours.push(hour);
  });
  (Array.isArray(currentPlansCache) ? currentPlansCache : []).forEach(row => {
    const hour = Number(row?.plan_hour ?? row?.actual_hour);
    if (Number.isFinite(hour)) cacheHours.push(hour);
  });

  if (cacheHours.length) return Math.max(...cacheHours);
  return getDefaultLastHour(dateStr);
}


function getUnifiedHourSet() {
  const set = new Set();
  (Array.isArray(currentPlansCache) ? currentPlansCache : []).forEach(row => {
    const hour = Number(row?.plan_hour ?? row?.actual_hour);
    if (Number.isFinite(hour)) set.add(hour);
  });
  (Array.isArray(currentActualsCache) ? currentActualsCache : []).forEach(row => {
    const status = normalizeStatus(row?.status);
    if (status === "cancel") return;
    const hour = Number(row?.actual_hour ?? row?.plan_hour);
    if (Number.isFinite(hour)) set.add(hour);
  });
  if (!set.size) set.add(getDefaultLastHour(els.dispatchDate?.value || todayStr()));
  return [...set].sort((a, b) => a - b);
}

function getOperationBaseHour() {
  const activeHours = (Array.isArray(currentActualsCache) ? currentActualsCache : [])
    .filter(row => {
      const status = normalizeStatus(row?.status);
      return status !== "done" && status !== "cancel";
    })
    .map(row => Number(row?.actual_hour ?? row?.plan_hour))
    .filter(Number.isFinite)
    .sort((a, b) => a - b);

  if (activeHours.length) return activeHours[0];

  const plannedHours = (Array.isArray(currentPlansCache) ? currentPlansCache : [])
    .map(row => Number(row?.plan_hour ?? row?.actual_hour))
    .filter(Number.isFinite)
    .sort((a, b) => a - b);

  if (plannedHours.length) return plannedHours[0];
  return getDefaultLastHour(els.dispatchDate?.value || todayStr());
}

function getNextHourSlot(baseHour) {
  const hours = getUnifiedHourSet();
  const found = hours.find(hour => Number(hour) > Number(baseHour));
  return Number.isFinite(found) ? found : null;
}

function getLastHourSlot() {
  const hours = getUnifiedHourSet();
  return hours.length ? hours[hours.length - 1] : getDefaultLastHour(els.dispatchDate?.value || todayStr());
}

function parseClockTextToMinutes(value) {
  if (!value || value === "-") return null;
  const m = String(value).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return Number(m[1]) * 60 + Number(m[2]);
}

function buildProjectedRowsForHour(targetHour) {
  const hour = Number(targetHour);
  const actualRows = (Array.isArray(currentActualsCache) ? currentActualsCache : [])
    .filter(row => Number(row?.actual_hour ?? row?.plan_hour) === hour)
    .filter(row => {
      const status = normalizeStatus(row?.status);
      return status !== "done" && status !== "cancel";
    });

  const actualPlanIds = new Set(actualRows.map(row => Number(row?.plan_id || 0)).filter(Boolean));
  const actualCastKeys = new Set(actualRows.map(row => `${Number(row?.cast_id || 0)}:${hour}`));
  const doneOrCancelPlanIds = new Set(
    (Array.isArray(currentActualsCache) ? currentActualsCache : [])
      .filter(row => Number(row?.actual_hour ?? row?.plan_hour) === hour)
      .filter(row => {
        const status = normalizeStatus(row?.status);
        return status === "done" || status === "cancel";
      })
      .map(row => Number(row?.plan_id || 0))
      .filter(Boolean)
  );

  const planRows = (Array.isArray(currentPlansCache) ? currentPlansCache : [])
    .filter(plan => Number(plan?.plan_hour ?? -1) === hour)
    .filter(plan => !actualPlanIds.has(Number(plan?.id || 0)))
    .filter(plan => !doneOrCancelPlanIds.has(Number(plan?.id || 0)))
    .filter(plan => !actualCastKeys.has(`${Number(plan?.cast_id || 0)}:${hour}`))
    .map((plan, index) => ({
      id: -100000 - index - Number(plan?.id || 0),
      plan_id: Number(plan?.id || 0),
      cast_id: Number(plan?.cast_id || 0),
      actual_hour: hour,
      plan_hour: hour,
      status: "pending",
      distance_km: Number(plan?.distance_km || plan?.casts?.distance_km || 0),
      destination_area: plan?.planned_area || plan?.casts?.area || "",
      destination_address: plan?.destination_address || plan?.casts?.address || "",
      note: plan?.note || "",
      casts: {
        ...(plan?.casts || {}),
        area: plan?.planned_area || plan?.casts?.area || "",
        address: plan?.destination_address || plan?.casts?.address || "",
        distance_km: Number(plan?.distance_km || plan?.casts?.distance_km || 0),
        travel_minutes: Number(plan?.casts?.travel_minutes || 0),
      },
      simulated_from_plan: true
    }));

  return [...actualRows, ...planRows];
}

function getEffectiveVehiclesForHour(targetHour) {
  const slotMinutes = Number(targetHour) * 60;
  const vehicles = Array.isArray(getSelectedVehiclesForToday()) ? getSelectedVehiclesForToday().filter(Boolean) : [];
  const activeItems = (Array.isArray(currentActualsCache) ? currentActualsCache : []).filter(row => {
    const status = normalizeStatus(row?.status);
    return status !== "done" && status !== "cancel";
  });

  return vehicles.filter(vehicle => {
    const rows = activeItems.filter(item => Number(item?.vehicle_id || 0) === Number(vehicle.id));
    if (!rows.length) return true;
    const orderedRows = moveManualLastItemsToEnd(sortItemsByNearestRoute(rows));
    const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
    const readyMinutes = parseClockTextToMinutes(forecast?.predictedReadyTime);
    if (readyMinutes === null) return true;
    return readyMinutes <= slotMinutes;
  });
}

function getVehicleDeadNamesForHour(targetHour) {
  const slotMinutes = Number(targetHour) * 60;
  const vehicles = Array.isArray(getSelectedVehiclesForToday()) ? getSelectedVehiclesForToday().filter(Boolean) : [];
  const activeItems = (Array.isArray(currentActualsCache) ? currentActualsCache : []).filter(row => {
    const status = normalizeStatus(row?.status);
    return status !== "done" && status !== "cancel";
  });

  return vehicles.filter(vehicle => {
    const rows = activeItems.filter(item => Number(item?.vehicle_id || 0) === Number(vehicle.id));
    if (!rows.length) return false;
    const orderedRows = moveManualLastItemsToEnd(sortItemsByNearestRoute(rows));
    const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
    const readyMinutes = parseClockTextToMinutes(forecast?.predictedReadyTime);
    return readyMinutes !== null && readyMinutes > slotMinutes;
  }).map(vehicle => vehicle?.driver_name || vehicle?.plate_number || `車両${vehicle.id}`);
}

function buildSimulationRowsForHour(targetHour, options = {}) {
  const hour = Number(targetHour);
  const includePlanInflow = Boolean(options.includePlanInflow);
  const actualRows = Array.isArray(currentActualsCache) ? currentActualsCache : [];
  const plans = Array.isArray(currentPlansCache) ? currentPlansCache : [];

  const actualPlanIds = new Set();
  const actualCastHourKeys = new Set();
  actualRows.forEach(row => {
    const planId = Number(row?.plan_id || 0);
    const castId = Number(row?.cast_id || 0);
    const actualHour = Number(row?.actual_hour ?? row?.plan_hour);
    if (planId) actualPlanIds.add(planId);
    if (castId && Number.isFinite(actualHour)) actualCastHourKeys.add(`${castId}:${actualHour}`);
  });

  const isOpenPlan = (plan) => {
    const status = String(plan?.status || '').toLowerCase();
    return status !== 'done' && status !== 'cancel';
  };

  const toSimRow = (plan, index, sourceType) => ({
    id: -700000 - index - Number(plan?.id || 0),
    plan_id: Number(plan?.id || 0),
    cast_id: Number(plan?.cast_id || 0),
    actual_hour: hour,
    plan_hour: Number(plan?.plan_hour ?? hour),
    status: 'pending',
    distance_km: Number(plan?.distance_km || plan?.casts?.distance_km || 0),
    destination_area: plan?.planned_area || plan?.casts?.area || '',
    destination_address: plan?.destination_address || plan?.casts?.address || '',
    note: plan?.note || '',
    simulation_source: sourceType,
    casts: {
      ...(plan?.casts || {}),
      area: plan?.planned_area || plan?.casts?.area || '',
      address: plan?.destination_address || plan?.casts?.address || '',
      distance_km: Number(plan?.distance_km || plan?.casts?.distance_km || 0),
      travel_minutes: Number(plan?.casts?.travel_minutes || 0),
    }
  });

  const selectedPlans = plans
    .filter(plan => Number(plan?.plan_hour ?? -1) === hour)
    .filter(isOpenPlan)
    .filter(plan => !actualPlanIds.has(Number(plan?.id || 0)))
    .filter(plan => !actualCastHourKeys.has(`${Number(plan?.cast_id || 0)}:${hour}`));

  const inflowPlans = includePlanInflow
    ? plans
        .filter(plan => Number(plan?.plan_hour ?? -1) < hour)
        .filter(isOpenPlan)
        .filter(plan => !actualPlanIds.has(Number(plan?.id || 0)))
    : [];

  const merged = [];
  const seenPlanIds = new Set();
  const seenCastIds = new Set();

  [...selectedPlans, ...inflowPlans].forEach(plan => {
    const planId = Number(plan?.id || 0);
    const castId = Number(plan?.cast_id || 0);
    if (planId && seenPlanIds.has(planId)) return;
    if (castId && seenCastIds.has(castId)) return;
    if (planId) seenPlanIds.add(planId);
    if (castId) seenCastIds.add(castId);
    merged.push(toSimRow(plan, merged.length, Number(plan?.plan_hour ?? hour) < hour ? 'plan_inflow' : 'slot_plan'));
  });

  const summary = {
    slotPlanCount: selectedPlans.length,
    inflowPlanCount: inflowPlans.filter(plan => {
      const planId = Number(plan?.id || 0);
      const castId = Number(plan?.cast_id || 0);
      return (!planId || seenPlanIds.has(planId)) && (!castId || seenCastIds.has(castId));
    }).length
  };

  return { rows: merged, summary };
}

function diagnoseSimulationHourWindow(targetHour, options = {}) {
  const built = buildSimulationRowsForHour(targetHour, options);
  const rows = built.rows;
  const vehicles = getEffectiveVehiclesForHour(targetHour);
  const effectiveVehicleCount = vehicles.length;
  const totalCapacity = vehicles.reduce((sum, vehicle) => sum + Math.max(1, Number(vehicle?.seat_capacity || 4)), 0);
  const castCount = rows.length;
  const areaGroups = new Set(rows.map(row => getAreaDisplayGroup(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "無し"))));
  const shortageCount = Math.max(0, castCount - totalCapacity);
  let statusKey = "ok";
  let statusText = "OK";
  if (totalCapacity < castCount) {
    statusKey = "danger";
    statusText = "定員オーバー見込み";
  } else if (areaGroups.size > effectiveVehicleCount && castCount > 0) {
    statusKey = "warn";
    statusText = "車両不足見込み";
  }
  return {
    hour: Number(targetHour),
    rows,
    vehicles,
    castCount,
    effectiveVehicleCount,
    totalCapacity,
    areaGroupCount: areaGroups.size,
    shortageCount,
    statusKey,
    statusText,
    deadVehicleNames: getVehicleDeadNamesForHour(targetHour),
    slotPlanCount: built.summary.slotPlanCount,
    inflowPlanCount: built.summary.inflowPlanCount
  };
}

function diagnoseHourWindow(targetHour) {
  const rows = buildProjectedRowsForHour(targetHour);
  const vehicles = getEffectiveVehiclesForHour(targetHour);
  const effectiveVehicleCount = vehicles.length;
  const totalCapacity = vehicles.reduce((sum, vehicle) => sum + Math.max(1, Number(vehicle?.seat_capacity || 4)), 0);
  const castCount = rows.length;
  const areaGroups = new Set(rows.map(row => getAreaDisplayGroup(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "無し"))));
  const shortageCount = Math.max(0, castCount - totalCapacity);
  let statusKey = "ok";
  let statusText = "OK";
  if (totalCapacity < castCount) {
    statusKey = "danger";
    statusText = "定員オーバー見込み";
  } else if (areaGroups.size > effectiveVehicleCount && castCount > 0) {
    statusKey = "warn";
    statusText = "車両不足見込み";
  }
  return {
    hour: Number(targetHour),
    rows,
    vehicles,
    castCount,
    effectiveVehicleCount,
    totalCapacity,
    areaGroupCount: areaGroups.size,
    shortageCount,
    statusKey,
    statusText,
    deadVehicleNames: getVehicleDeadNamesForHour(targetHour)
  };
}

function getStatusPillClass(statusKey) {
  if (statusKey === "danger") return "danger";
  if (statusKey === "warn") return "warn";
  return "ok";
}

function buildStatusPill(label, value, statusKey = "ok") {
  return `<span class="status-pill ${getStatusPillClass(statusKey)}"><span class="status-pill-label">${escapeHtml(label)}</span><span class="status-pill-value">${escapeHtml(value)}</span></span>`;
}

function buildSummaryItemsHtml(items) {
  return items.map(item => `
    <div class="hybrid-summary-item ${escapeHtml(item.tone || '')}">
      <div class="hybrid-summary-icon">${escapeHtml(item.icon || '•')}</div>
      <div class="hybrid-summary-body">
        <div class="hybrid-summary-label">${escapeHtml(item.label)}</div>
        <div class="hybrid-summary-value">${escapeHtml(item.value)}</div>
      </div>
    </div>
  `).join("");
}

function buildOperationDiagnosisHtml(operationDiag, nextDiag, lastDiag) {
  const nextDead = nextDiag?.deadVehicleNames?.length ? nextDiag.deadVehicleNames.join(" / ") : "なし";
  return `
    <div class="hybrid-state-row">
      ${buildStatusPill("今", `${getHourLabel(operationDiag.hour)} / ${operationDiag.castCount}名`, operationDiag.statusKey)}
      ${buildStatusPill("次", nextDiag ? `${getHourLabel(nextDiag.hour)} / ${nextDiag.castCount}名` : '-', nextDiag?.statusKey || 'ok')}
      ${buildStatusPill("終", lastDiag ? `${getHourLabel(lastDiag.hour)} / ${lastDiag.castCount}名` : '-', lastDiag?.statusKey || 'ok')}
    </div>
    <div class="hybrid-legend-grid">
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便の状態</div>
        <div class="hybrid-legend-value">${escapeHtml(nextDiag ? nextDiag.statusText : '-')}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便NG車両</div>
        <div class="hybrid-legend-value">${escapeHtml(nextDead)}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便 実効/定員</div>
        <div class="hybrid-legend-value">${nextDiag ? `${nextDiag.effectiveVehicleCount}台 / ${nextDiag.totalCapacity}` : '-'}</div>
      </div>
    </div>
  `;
}

function buildSimulationDiagnosisHtml(diag) {
  return `
    <div class="hybrid-state-row">
      ${buildStatusPill("便", `${getHourLabel(diag.hour)}`, diag.statusKey)}
      ${buildStatusPill("人数", `${diag.castCount}名`, diag.statusKey)}
      ${buildStatusPill("車両", `${diag.effectiveVehicleCount}台`, diag.statusKey)}
      ${buildStatusPill("定員", `${diag.totalCapacity}`, diag.statusKey)}
    </div>
    <div class="hybrid-legend-grid">
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">判定</div>
        <div class="hybrid-legend-value">${escapeHtml(diag.statusText)}${diag.shortageCount > 0 ? ` / 未配車見込み ${diag.shortageCount}名` : ''}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">方面系統</div>
        <div class="hybrid-legend-value">${diag.areaGroupCount}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便NG車両</div>
        <div class="hybrid-legend-value">${escapeHtml(diag.deadVehicleNames.length ? diag.deadVehicleNames.join(' / ') : 'なし')}</div>
      </div>
    </div>
  `;
}

function renderOperationAndSimulationUI() {
  if (isRefreshingHybridUI) return;
  isRefreshingHybridUI = true;
  try {
    const operationBaseHour = getOperationBaseHour();
    const realNextHour = getNextHourSlot(operationBaseHour);
    const lastHour = getLastHourSlot();
    const operationDiag = diagnoseHourWindow(operationBaseHour);
    const nextDiag = Number.isFinite(realNextHour) ? diagnoseHourWindow(realNextHour) : null;
    const lastDiag = Number.isFinite(lastHour) ? diagnoseHourWindow(lastHour) : null;

    if (els.operationContextSummary) {
      els.operationContextSummary.innerHTML = buildSummaryItemsHtml([
        { icon: "📍", label: "今の基準", value: getHourLabel(operationBaseHour), tone: "tone-blue" },
        { icon: "⏭", label: "次に見る便", value: Number.isFinite(realNextHour) ? getHourLabel(realNextHour) : "-", tone: "tone-indigo" },
        { icon: "🏁", label: "最終便", value: Number.isFinite(lastHour) ? getHourLabel(lastHour) : "-", tone: "tone-amber" },
        { icon: "🚐", label: "次便の使える車両", value: nextDiag ? `${nextDiag.effectiveVehicleCount}台` : "-", tone: "tone-green" }
      ]);
    }

    if (els.operationDiagnosis) {
      const cls = nextDiag?.statusKey === "danger" || lastDiag?.statusKey === "danger"
        ? "danger"
        : (nextDiag?.statusKey === "warn" || lastDiag?.statusKey === "warn" ? "warn" : "ok");      els.operationDiagnosis.className = `hybrid-diagnosis ${cls}`;
      els.operationDiagnosis.innerHTML = buildOperationDiagnosisHtml(operationDiag, nextDiag, lastDiag);
    }

    const hourOptions = getUnifiedHourSet();
    if (els.simulationSlotSelect) {
      const previous = Number.isFinite(simulationSlotHour) ? simulationSlotHour : (Number.isFinite(realNextHour) ? realNextHour : operationBaseHour);
      suppressSimulationSlotChange = true;
      try {
        const optionHtml = hourOptions.map(hour => `<option value="${hour}">${getHourLabel(hour)}</option>`).join("");
        if (els.simulationSlotSelect.innerHTML !== optionHtml) {
          els.simulationSlotSelect.innerHTML = optionHtml;
        }
        const targetValue = hourOptions.includes(previous) ? previous : (hourOptions[0] ?? operationBaseHour);
        simulationSlotHour = targetValue;
        if (String(els.simulationSlotSelect.value) !== String(targetValue)) {
          els.simulationSlotSelect.value = String(targetValue);
        }
      } finally {
        setTimeout(() => { suppressSimulationSlotChange = false; }, 0);
      }
    }

    if (!lastSimulationResult && els.simulationDiagnosis) {
      els.simulationDiagnosis.className = 'hybrid-diagnosis muted';
      els.simulationDiagnosis.textContent = '試算対象便を選択してください';
    }
  } catch (error) {
    console.error(error);
    if (els.operationDiagnosis) {
      els.operationDiagnosis.className = 'hybrid-diagnosis danger';
      els.operationDiagnosis.textContent = `診断エラー: ${error?.message || error}`;
    }
  } finally {
    isRefreshingHybridUI = false;
  }
}

function runSlotDiagnosisPreview() {
  const hour = Number(els.simulationSlotSelect?.value ?? simulationSlotHour ?? getOperationBaseHour());
  simulationSlotHour = hour;
  const diag = diagnoseSimulationHourWindow(hour, { includePlanInflow: Boolean(els.simulationIncludePlanInflow?.checked) });
  lastSimulationResult = { type: 'diagnosis', hour, diag };
  if (els.simulationDiagnosis) {    els.simulationDiagnosis.className = `hybrid-diagnosis ${diag.statusKey}`;
    els.simulationDiagnosis.innerHTML = buildSimulationDiagnosisHtml(diag);
  }
  if (els.simulationPreview) {
    if (!diag.rows.length) {
      els.simulationPreview.className = 'simulation-preview muted';
      els.simulationPreview.textContent = 'この便の予定対象はありません';
    } else {
      const list = diag.rows.map(row => `
        <div class="sim-mini-chip">
          <span class="sim-mini-name">${escapeHtml(row?.casts?.name || '-')}</span>
          <span class="sim-mini-area">${escapeHtml(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || '-'))}</span>
        </div>
      `).join('');
      els.simulationPreview.className = 'simulation-preview';
      els.simulationPreview.innerHTML = `
        <div class="sim-preview-head">
          <h4 class="sim-preview-title">試算対象一覧</h4>
          <span class="chip">${getHourLabel(diag.hour)}</span>
        </div>
        <div class="sim-preview-meta">対象便予定 ${diag.slotPlanCount}名 / 前便未処理予定 ${diag.inflowPlanCount}名</div>
        <div class="sim-mini-chip-grid">${list}</div>
      `;
    }
  }
}

function getAssignmentsForPreview(items, vehicles, monthlyMap) {
  let assignments = optimizeAssignments(items, vehicles, monthlyMap);
  if (!Array.isArray(assignments)) assignments = [];

  if (__hasEnoughVehiclesForDisplayGroups(items, vehicles)) {
    assignments = __buildAssignmentsPreserveDisplayGroups(items, vehicles, monthlyMap);
  } else {
    const routeFlowAssignments = optimizeAssignmentsByRouteFlow(assignments, items, vehicles);
    assignments = Array.isArray(routeFlowAssignments) ? routeFlowAssignments : assignments;
    if (!Array.isArray(assignments) || !assignments.length) assignments = buildFallbackAssignments(items, vehicles);
    const balancedAssignments = optimizeAssignmentsByDistanceBalance(assignments, items, vehicles, monthlyMap);
    assignments = Array.isArray(balancedAssignments) ? balancedAssignments : assignments;
    const correctedAssignments = applyLastTripDistanceCorrectionToAssignments(assignments, items, vehicles, monthlyMap);
    assignments = Array.isArray(correctedAssignments) ? correctedAssignments : assignments;
    const manualAssignments = applyManualLastVehicleToAssignments(assignments, vehicles);
    assignments = Array.isArray(manualAssignments) ? manualAssignments : assignments;
  }

  return Array.isArray(assignments) ? assignments : [];
}

function runSimulationDispatchPreview() {
  const hour = Number(els.simulationSlotSelect?.value ?? simulationSlotHour ?? getOperationBaseHour());
  simulationSlotHour = hour;
  const simBuilt = buildSimulationRowsForHour(hour, { includePlanInflow: Boolean(els.simulationIncludePlanInflow?.checked) });
  const items = simBuilt.rows.map((row, index) => ({ ...row, id: Number(row?.id || -(index + 1)) }));
  const vehicles = getEffectiveVehiclesForHour(hour);
  const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
  const assignments = getAssignmentsForPreview(items, vehicles, monthlyMap);
  const applied = items.map(item => {
    const assigned = assignments.find(a => Number(a?.item_id) === Number(item.id));
    return assigned ? { ...item, vehicle_id: assigned.vehicle_id, driver_name: assigned.driver_name || '' } : { ...item };
  });
  const cards = buildDailyDispatchVehicleCards(vehicles, applied, monthlyMap);
  const html = cards.map(({ vehicle, orderedRows }) => {
    const summary = getVehicleDailySummary(vehicle, orderedRows);
    const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
    const body = orderedRows.length
      ? orderedRows.map((row, index) => `
          <div class="sim-preview-row">
            <span>順番${index + 1} ${escapeHtml(row?.casts?.name || '-')} / ${escapeHtml(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || '-'))}</span>
            <span>${Number(row?.distance_km || 0).toFixed(1)}km</span>
          </div>
        `).join('')
      : `<div class="muted">送りなし</div>`;
    return `
      <div class="sim-preview-card">
        <h4>${escapeHtml(vehicle?.driver_name || vehicle?.plate_number || '-')}</h4>
        <div class="sim-preview-meta">人数 ${orderedRows.length} / 距離 ${summary.totalKm.toFixed(1)}km / 戻り ${escapeHtml(forecast.returnAfterLabel)} / 次便可能 ${escapeHtml(forecast.predictedReadyTime)}</div>
        ${body}
      </div>
    `;
  }).join('');

  const unassigned = applied.filter(item => !Number(item?.vehicle_id || 0));
  const diag = diagnoseSimulationHourWindow(hour, { includePlanInflow: Boolean(els.simulationIncludePlanInflow?.checked) });
  lastSimulationResult = { type: 'dispatch', hour, diag, unassignedCount: unassigned.length };
  if (els.simulationDiagnosis) {
    const cls = unassigned.length > 0 || diag.statusKey === 'danger' ? 'danger' : (diag.statusKey === 'warn' ? 'warn' : 'ok');    els.simulationDiagnosis.className = `hybrid-diagnosis ${cls}`;
    els.simulationDiagnosis.innerHTML = buildSimulationDiagnosisHtml({
      ...diag,
      effectiveVehicleCount: vehicles.length,
      shortageCount: unassigned.length
    });
  }
  if (els.simulationPreview) {
    els.simulationPreview.className = 'simulation-preview';
    els.simulationPreview.innerHTML = `
      <div class="sim-preview-head">
        <h4 class="sim-preview-title">試算配車プレビュー</h4>
        <span class="chip">${getHourLabel(hour)}</span>
      </div>
      <div class="sim-preview-meta">このプレビューは保存しません。対象便予定 ${diag.slotPlanCount}名 / 前便未処理予定 ${diag.inflowPlanCount}名 / 現実の配車状態は維持されます。</div>
      <div class="sim-preview-grid">${html || '<div class="muted">配車結果なし</div>'}</div>
    `;
  }
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

function ensureCastTravelMinutesUi() {
  if (document.getElementById("castTravelMinutes")) {
    els.castTravelMinutes = document.getElementById("castTravelMinutes");
    els.fetchCastTravelMinutesBtn = document.getElementById("fetchCastTravelMinutesBtn");
    if (els.fetchCastTravelMinutesBtn) {
      els.fetchCastTravelMinutesBtn.style.display = "none";
      els.fetchCastTravelMinutesBtn.disabled = true;
      els.fetchCastTravelMinutesBtn.title = "時間取得は使用しません";
    }
    return;
  }

  const distanceField = els.castDistanceKm?.closest?.(".field");
  if (distanceField?.parentElement) {
    const wrap = document.createElement("div");
    wrap.className = "field";
    wrap.innerHTML = `
      <label for="castTravelMinutes">片道予想時間(分)</label>
      <input id="castTravelMinutes" type="number" min="0" step="1" placeholder="例：35" />
    `;
    distanceField.insertAdjacentElement("afterend", wrap);
  }

  els.castTravelMinutes = document.getElementById("castTravelMinutes");
  els.fetchCastTravelMinutesBtn = document.getElementById("fetchCastTravelMinutesBtn");
  if (els.fetchCastTravelMinutesBtn) {
    els.fetchCastTravelMinutesBtn.style.display = "none";
    els.fetchCastTravelMinutesBtn.disabled = true;
    els.fetchCastTravelMinutesBtn.title = "時間取得は使用しません";
  }
}

function getStoredTravelMinutes(value) {
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? Math.round(n) : 0;
}

function getCastTravelMinutesValue(castLike) {
  if (!castLike) return 0;
  return getStoredTravelMinutes(castLike.travel_minutes || castLike.travelMinutes);
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
      if (status === 'ZERO_RESULTS' || status === 'NOT_FOUND' || status === 'REQUEST_DENIED' || status === 'INVALID_REQUEST') {
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
  const rows = Array.isArray(items) ? items.filter(Boolean) : [];
  if (!rows.length) return [];

  const groups = new Map();
  rows.forEach((row, idx) => {
    const hourKey = Number(row?.actual_hour ?? row?.plan_hour ?? 0);
    if (!groups.has(hourKey)) groups.set(hourKey, []);
    groups.get(hourKey).push({ row, idx });
  });

  const ordered = [];
  const hourKeys = [...groups.keys()].sort((a, b) => a - b);

  function getRowDistanceValue(row) {
    return Number(row?.distance_km ?? row?.casts?.distance_km ?? 999999) || 999999;
  }
  function getRowTravelValue(row) {
    const v = Number(row?.casts?.travel_minutes ?? row?.travel_minutes ?? 999999);
    return Number.isFinite(v) && v > 0 ? v : 999999;
  }
  function getRowStableId(row, fallbackIdx) {
    return Number(row?.id ?? row?.item_id ?? row?.cast_id ?? fallbackIdx ?? 0) || 0;
  }

  hourKeys.forEach(hourKey => {
    const remaining = groups.get(hourKey).slice();
    let currentDistance = null;

    while (remaining.length) {
      let bestIndex = 0;
      let bestKey = null;

      remaining.forEach((entry, index) => {
        const row = entry.row;
        const distanceValue = getRowDistanceValue(row);
        const travelValue = getRowTravelValue(row);
        const stableId = getRowStableId(row, entry.idx);
        const primary = currentDistance == null
          ? distanceValue
          : Math.abs(distanceValue - currentDistance);
        const compareKey = [primary, distanceValue, travelValue, stableId];

        if (!bestKey) {
          bestKey = compareKey;
          bestIndex = index;
          return;
        }
        for (let i = 0; i < compareKey.length; i += 1) {
          if (compareKey[i] === bestKey[i]) continue;
          if (compareKey[i] < bestKey[i]) {
            bestKey = compareKey;
            bestIndex = index;
          }
          break;
        }
      });

      const picked = remaining.splice(bestIndex, 1)[0].row;
      ordered.push(picked);
      currentDistance = getRowDistanceValue(picked);
    }
  });

  return ordered;
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

function extractCityTownFromAddress(address = "") {
  const raw = String(address || "").trim();
  const a = normalizeAddressText(raw);
  if (!a) return { city: "", town: "" };

  const match = a.match(
    /(足立区|葛飾区|江戸川区|墨田区|江東区|荒川区|台東区|中央区|千代田区|港区|新宿区|文京区|品川区|目黒区|大田区|世田谷区|渋谷区|中野区|杉並区|豊島区|北区|板橋区|練馬区|松戸市|柏市|白井市|印西市|流山市|我孫子市|野田市|市川市|鎌ケ谷市|鎌ヶ谷市|船橋市|浦安市|習志野市|八千代市|三郷市|八潮市|草加市|越谷市|吉川市|守谷市|取手市|つくば市|つくばみらい市|牛久市|龍ケ崎市|龍ヶ崎市|利根町)(.+)?/
  );

  if (!match) return { city: "", town: "" };

  const cityRaw = match[1] || "";
  let townRaw = match[2] || "";

  townRaw = townRaw
    .replace(/丁目.*$/, "")
    .replace(/番地.*$/, "")
    .replace(/番.*$/, "")
    .replace(/号.*$/, "")
    .replace(/[0-9０-９]+/g, "")
    .replace(/[-‐-‒–—―ー－]+/g, "")
    .replace(/^千葉県|^埼玉県|^東京都|^茨城県/, "")
    .replace(/^日本/, "")
    .trim();

  const city = cityRaw.replace(/(市|区|町)$/, "");
  return { city, town: townRaw };
}

function getSevenAreaGroupByCity(city = "") {
  if (!city) return "東京方面";

  if (["松戸"].includes(city)) return "松戸方面";

  if (["柏", "白井", "印西"].includes(city)) return "柏方面";

  if (["我孫子", "取手", "守谷", "つくばみらい", "牛久", "龍ケ崎", "龍ヶ崎", "利根", "つくば"].includes(city)) {
    return "我孫子・取手方面";
  }

  if (["市川", "船橋", "鎌ケ谷", "鎌ヶ谷", "浦安", "習志野", "八千代"].includes(city)) {
    return "市川・船橋方面";
  }

  if ([
    "足立", "葛飾", "江戸川", "墨田", "江東", "荒川", "台東",
    "中央", "千代田", "港", "新宿", "文京", "品川", "目黒",
    "大田", "世田谷", "渋谷", "中野", "杉並", "豊島", "北",
    "板橋", "練馬"
  ].includes(city)) {
    return "東京方面";
  }

  if (["三郷", "八潮", "草加", "越谷", "吉川"].includes(city)) return "埼玉方面";

  if (["流山", "野田"].includes(city)) return "流山・野田方面";

  return "東京方面";
}

function detectAreaGroup(address = "") {
  const { city, town } = extractCityTownFromAddress(address);

  const allJobs = [
    ...(currentPlansCache || []),
    ...(currentActualsCache || []),
    ...(allCastsCache || [])
  ];

  const normalizedTown = String(town || "").trim();

  const sameTownGroups = allJobs
    .map(j => ({
      address: j.address || j.destination_address || "",
      area: j.area || j.area_group || j.planned_area || j.destination_area || ""
    }))
    .map(j => {
      const parsed = extractCityTownFromAddress(j.address);
      return {
        city: parsed.city,
        town: parsed.town,
        group: getAreaDisplayGroup(j.area)
      };
    })
    .filter(x => x.city === city && normalizedTown && x.town === normalizedTown && x.group);

  if (sameTownGroups.length > 0) {
    const count = {};
    sameTownGroups.forEach(x => count[x.group] = (count[x.group] || 0) + 1);
    return Object.keys(count).sort((a, b) => count[b] - count[a])[0];
  }

  const sameCityGroups = allJobs
    .map(j => ({
      address: j.address || j.destination_address || "",
      area: j.area || j.area_group || j.planned_area || j.destination_area || ""
    }))
    .map(j => {
      const parsed = extractCityTownFromAddress(j.address);
      return {
        city: parsed.city,
        group: getAreaDisplayGroup(j.area)
      };
    })
    .filter(x => x.city === city && x.group);

  if (sameCityGroups.length > 0) {
    const count = {};
    sameCityGroups.forEach(x => count[x.group] = (count[x.group] || 0) + 1);
    return Object.keys(count).sort((a, b) => count[b] - count[a])[0];
  }

  return getSevenAreaGroupByCity(city);
}function classifyAreaByAddress(address) {
  return detectAreaGroup(address);
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
  const a = normalizeAreaLabel(area || "");

  if (!a || a === "無し") return "東京方面";

  if (a.includes("松戸")) return "松戸方面";

  if (a.includes("柏") || a.includes("白井") || a.includes("印西")) return "柏方面";

  if (
    a.includes("我孫子") || a.includes("取手") || a.includes("守谷") ||
    a.includes("牛久") || a.includes("藤代") || a.includes("つくば") ||
    a.includes("龍ケ崎") || a.includes("龍ヶ崎")
  ) return "我孫子・取手方面";

  if (
    a.includes("市川") || a.includes("船橋") || a.includes("鎌ヶ谷") ||
    a.includes("鎌ケ谷") || a.includes("浦安") || a.includes("習志野") ||
    a.includes("八千代")
  ) return "市川・船橋方面";

  if (
    a.includes("足立") || a.includes("葛飾") || a.includes("江戸川") ||
    a.includes("墨田") || a.includes("江東") || a.includes("荒川") ||
    a.includes("台東") || a.includes("中央") || a.includes("千代田") ||
    a.includes("港") || a.includes("新宿") || a.includes("文京") ||
    a.includes("品川") || a.includes("目黒") || a.includes("大田") ||
    a.includes("世田谷") || a.includes("渋谷") || a.includes("中野") ||
    a.includes("杉並") || a.includes("豊島") || a.includes("北") ||
    a.includes("板橋") || a.includes("練馬") || a.includes("東京")
  ) return "東京方面";

  if (
    a.includes("三郷") || a.includes("八潮") || a.includes("草加") ||
    a.includes("越谷") || a.includes("吉川")
  ) return "埼玉方面";

  if (a.includes("流山") || a.includes("野田")) return "流山・野田方面";

  return "東京方面";
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

  if (areas.length) {
    if (hasHardReverseMix(targetArea, areas)) best -= 220;
    best += getAreaBundleBonus(targetArea, areas);
    best -= getAreaSplitPenalty(targetArea, areas);
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

const HARD_ROUTE_MIX_GROUPS = {
  NE: ["取手方面", "藤代方面", "守谷方面", "我孫子方面", "柏方面", "柏の葉方面"],
  N: ["流山方面", "野田方面", "吉川方面"],
  NW: ["三郷方面", "八潮方面", "草加方面", "越谷方面", "足立方面", "葛飾方面"],
  S: ["市川方面", "江戸川方面"],
  SE: ["船橋方面", "鎌ヶ谷方面", "千葉方面"],
  CENTER: ["松戸近郊"]
};

function getHardRouteMixGroup(area) {
  const canonical = getCanonicalArea(area);
  for (const [group, areas] of Object.entries(HARD_ROUTE_MIX_GROUPS)) {
    if (areas.includes(canonical)) return group;
  }
  return "";
}

function isHardReverseMixForRoute(areaA, areaB) {
  const a = getCanonicalArea(areaA);
  const b = getCanonicalArea(areaB);
  if (!a || !b || a === b) return false;

  const groupA = getHardRouteMixGroup(a);
  const groupB = getHardRouteMixGroup(b);
  const dirA = getAreaTravelDirection(a);
  const dirB = getAreaTravelDirection(b);
  const dirDistance = getDirectionDistanceByKey(dirA, dirB);
  const affinity = getAreaAffinityScore(a, b);

  const neLike = new Set(["NE", "N"]);
  const southLike = new Set(["S", "SE", "SW"]);
  const westLike = new Set(["W", "NW"]);

  if ((neLike.has(groupA) && southLike.has(groupB)) || (neLike.has(groupB) && southLike.has(groupA))) {
    return true;
  }
  if ((groupA === "NE" && westLike.has(groupB)) || (groupB === "NE" && westLike.has(groupA))) {
    return affinity <= 28;
  }
  if ((groupA === "N" && southLike.has(groupB)) || (groupB === "N" && southLike.has(groupA))) {
    return affinity <= 36;
  }
  if (dirDistance >= 4 && affinity <= 28) return true;
  if (dirDistance >= 3 && affinity <= 18 && ![a, b].includes("松戸近郊")) return true;
  return false;
}

function hasHardReverseMix(targetArea, existingAreas = []) {
  return (existingAreas || []).some(area => isHardReverseMixForRoute(targetArea, area));
}

function getHardReverseMixPenalty(targetArea, existingAreas = []) {
  const reverseCount = (existingAreas || []).filter(area => isHardReverseMixForRoute(targetArea, area)).length;
  return reverseCount ? 520 + (reverseCount - 1) * 180 : 0;
}

function getAreaMixGuardScore(targetArea, existingAreas = [], isLastRun = false) {
  const penalty = getHardReverseMixPenalty(targetArea, existingAreas);
  if (!penalty) return 0;
  return isLastRun ? penalty * 1.15 : penalty;
}

function getAreaBundleBonus(targetArea, existingAreas = []) {
  const areas = Array.isArray(existingAreas) ? existingAreas.filter(Boolean) : [];
  if (!areas.length) return 0;

  const targetCanonical = getCanonicalArea(targetArea);
  const targetGroup = getHardRouteMixGroup(targetArea);
  let sameCanonicalCount = 0;
  let nearGroupCount = 0;
  let highAffinityCount = 0;

  for (const area of areas) {
    const canonical = getCanonicalArea(area);
    if (canonical && canonical === targetCanonical) sameCanonicalCount += 1;
    if (targetGroup && getHardRouteMixGroup(area) === targetGroup) nearGroupCount += 1;
    if (getAreaAffinityScore(targetArea, area) >= 72) highAffinityCount += 1;
  }

  let bonus = 0;
  if (sameCanonicalCount >= 2) bonus += 86;
  else if (sameCanonicalCount >= 1) bonus += 44;

  if (nearGroupCount >= 2) bonus += 56;
  else if (nearGroupCount >= 1) bonus += 24;

  if (highAffinityCount >= 2) bonus += 40;
  else if (highAffinityCount >= 1) bonus += 18;

  return bonus;
}

function getAreaSplitPenalty(targetArea, existingAreas = []) {
  const areas = Array.isArray(existingAreas) ? existingAreas.filter(Boolean) : [];
  const groups = new Set(areas.map(getHardRouteMixGroup).filter(Boolean));
  const targetGroup = getHardRouteMixGroup(targetArea);
  if (!targetGroup || !groups.size) return 0;

  const nearGroupCount = areas.filter(area => getHardRouteMixGroup(area) === targetGroup).length;
  const highAffinityCount = areas.filter(area => getAreaAffinityScore(targetArea, area) >= 72).length;

  if (groups.size >= 2 && !groups.has(targetGroup)) return 220;
  if (groups.size >= 2 && groups.has(targetGroup)) {
    if (nearGroupCount >= 1 || highAffinityCount >= 1) return 18;
    return 68;
  }
  return 0;
}

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
  const groupA = getHardRouteMixGroup(areaA);
  const groupB = getHardRouteMixGroup(areaB);
  const sameBroadGroup = groupA && groupA === groupB;

  let penalty = 0;

  if (distance === 0) penalty += 0;
  else if (distance === 1) penalty += sameBroadGroup ? 6 : 18;
  else if (distance === 2) penalty += sameBroadGroup ? 28 : 120;
  else if (distance === 3) penalty += sameBroadGroup ? 88 : 300;
  else if (distance >= 4 && distance < 99) penalty += sameBroadGroup ? 160 : 560;

  if (routeFlow <= 0) penalty += sameBroadGroup ? 8 : 42;
  else if (routeFlow < 40) penalty += sameBroadGroup ? 4 : 22;

  if (affinity >= 88) penalty -= 72;
  else if (affinity >= 72) penalty -= 40;
  else if (affinity >= 54) penalty -= 14;

  const gatewayA = isGatewayNearArea(areaA);
  const gatewayB = isGatewayNearArea(areaB);
  if (gatewayA || gatewayB) {
    if (distance >= 2) penalty += 140;
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
    penalty += 140;
  }

  if (isHardReverseMixForRoute(areaA, areaB)) {
    penalty += 900;
  } else if (sameBroadGroup && affinity >= 72) {
    penalty -= 36;
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

  penalty += getAreaMixGuardScore(targetArea, areas);

  return Math.max(0, penalty);
}

function isStrictSameDirectionArea(targetArea, existingAreas = []) {
  const target = normalizeAreaLabel(targetArea);
  const areas = Array.isArray(existingAreas) ? existingAreas.map(normalizeAreaLabel).filter(Boolean) : [];
  if (!target || target === "無し" || !areas.length) return true;
  const targetGroup = getAreaDisplayGroup(target);
  return areas.every(area => {
    const group = getAreaDisplayGroup(area);
    if (group && targetGroup && group !== targetGroup) return false;
    return !hasHardReverseMix(target, [area]);
  });
}

function isLastRunHardAreaConstraintSatisfied(targetArea, existingAreas = [], homeArea = "") {
  const target = normalizeAreaLabel(targetArea);
  const home = normalizeAreaLabel(homeArea || "");
  if (!target || target === "無し") return false;
  if (home && isHardReverseForHome(target, home)) return false;
  if (!isStrictSameDirectionArea(target, existingAreas)) return false;
  const strict = getStrictHomeCompatibilityScore(target, home);
  const direction = getDirectionAffinityScore(target, home);
  // Last run should prioritize going toward driver's home direction.
  return strict >= 52 || direction >= 24;
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
      if (existingAreas.length && hasHardReverseMix(area, existingAreas)) continue;
      const routeScore = getRouteFlowVehicleScore(area, existingAreas, homeArea);
      const continuityPenalty = getRouteContinuityPenalty(area, existingAreas, homeArea);
      const currentContinuityPenalty = getRouteContinuityPenalty(area, currentAreas, currentHome);
      const homeScore = getLastTripHomePriorityWeight(area, homeArea, true, true);
      const direction = Math.max(0, getDirectionAffinityScore(area, homeArea));
      const strict = getStrictHomeCompatibilityScore(area, homeArea);

      const mixGuardPenalty = getAreaMixGuardScore(area, existingAreas, true);
      const currentMixGuardPenalty = getAreaMixGuardScore(area, currentAreas, true);
      const totalScore = routeScore * 2.4 + homeScore * 0.9 + direction * 0.5 + strict * 0.6 - continuityPenalty * 2.15 - mixGuardPenalty - hourAssignments.length * 12;
      const currentTotal = currentRouteScore * 2.4 + currentHomeScore * 0.9 + Math.max(0, getDirectionAffinityScore(area, currentHome)) * 0.5 + getStrictHomeCompatibilityScore(area, currentHome) * 0.6 - currentContinuityPenalty * 2.15 - currentMixGuardPenalty - currentAreas.length * 12;

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

    travel_minutes: "travel_minutes",
    travel_time: "travel_minutes",
    片道予想時間: "travel_minutes",
    "片道予想時間(分)": "travel_minutes",
    予想時間: "travel_minutes",
    時間: "travel_minutes",

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
      <span class="chip">${escapeHtml(vehicle.driver_name || vehicle.plate_number || "-")}</span>
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
  if (els.castTravelMinutes) els.castTravelMinutes.value = "";
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
  if (els.castTravelMinutes) els.castTravelMinutes.value = cast.travel_minutes ?? "";
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
    travel_minutes: getStoredTravelMinutes(els.castTravelMinutes?.value) || null,
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

  const castsTable = els.castsTableBody.closest("table");
  const castsHeaderRow = castsTable?.querySelector("thead tr");
  if (castsHeaderRow) {
    castsHeaderRow.innerHTML = `
      <th>氏名</th>
      <th>住所</th>
      <th>方面</th>
      <th>想定距離(km)</th>
      <th>片道予想時間(分)</th>
      <th>メモ</th>
      <th>操作</th>
    `;
  }

  if (!allCastsCache.length) {
    els.castsTableBody.innerHTML = `<tr><td colspan="7" class="muted">キャストがありません</td></tr>`;
    return;
  }

  allCastsCache.forEach(cast => {
  const tr = document.createElement("tr");
  const travelMinutes = getStoredTravelMinutes(cast.travel_minutes);
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
  <td>${travelMinutes ?? ""}</td>
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
    "travel_minutes",
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
    getStoredTravelMinutes(cast.travel_minutes) ?? "",
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
        travel_minutes: getStoredTravelMinutes(row.travel_minutes) || null,
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
          <th>片道予想時間(分)</th>
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
        <td>${getStoredTravelMinutes(cast.travel_minutes) ?? ""}</td>
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


function buildMileageCalendarRows(rows, startDate, endDate) {
  const start = new Date(`${startDate}T00:00:00`);
  const monthEnd = new Date(start.getFullYear(), start.getMonth() + 1, 0);

  const firstHalfDays = [];
  const secondHalfDays = [];

  for (let day = 1; day <= monthEnd.getDate(); day++) {
    const d = new Date(start.getFullYear(), start.getMonth(), day);
    const item = {
      day,
      key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`,
      label: `${day}日`
    };
    if (day <= 16) firstHalfDays.push(item);
    else secondHalfDays.push(item);
  }

  const grouped = new Map();

  (rows || []).forEach(row => {
    const vehicleId = Number(row.vehicle_id || 0);
    const driver = String(row.driver_name || "-").trim() || "-";
    const plateNumber = String(row.plate_number || "").trim();
    const key = vehicleId > 0 ? `vehicle:${vehicleId}` : `name:${driver}__${plateNumber}`;

    if (!grouped.has(key)) {
      grouped.set(key, {
        vehicle_id: vehicleId,
        plate_number: plateNumber,
        driver_name: driver,
        byDate: new Map(),
        total_distance_km: 0,
        worked_days: 0,
        avg_distance_km: 0
      });
    }

    const item = grouped.get(key);
    const dateKey = String(row.report_date || "");
    const distance = Number(row.distance_km || 0);
    item.byDate.set(dateKey, Number(((item.byDate.get(dateKey) || 0) + distance).toFixed(1)));
    item.total_distance_km += distance;
  });

  const vehicleOrderMap = new Map(
    (typeof getSortedVehiclesForDisplay === "function" ? getSortedVehiclesForDisplay() : (allVehiclesCache || []))
      .map((vehicle, index) => [Number(vehicle.id || 0), index])
  );

  const drivers = [...grouped.values()].sort((a, b) => {
    const orderA = vehicleOrderMap.has(Number(a.vehicle_id || 0)) ? vehicleOrderMap.get(Number(a.vehicle_id || 0)) : Number.MAX_SAFE_INTEGER;
    const orderB = vehicleOrderMap.has(Number(b.vehicle_id || 0)) ? vehicleOrderMap.get(Number(b.vehicle_id || 0)) : Number.MAX_SAFE_INTEGER;
    if (orderA !== orderB) return orderA - orderB;
    return String(a.driver_name || "").localeCompare(String(b.driver_name || ""), "ja", { numeric: true, sensitivity: "base" });
  });

  drivers.forEach(driver => {
    driver.worked_days = [...driver.byDate.values()].filter(v => Number(v || 0) > 0).length;
    driver.total_distance_km = Number(driver.total_distance_km.toFixed(1));
    driver.avg_distance_km = driver.worked_days ? Number((driver.total_distance_km / driver.worked_days).toFixed(1)) : 0;
  });

  return {
    firstHalfDays,
    secondHalfDays,
    drivers,
    monthLabel: `${start.getFullYear()}年${start.getMonth() + 1}月`
  };
}

function renderMileageMatrixSection(title, days, drivers, options = {}) {
  const showOverallColumns = Boolean(options.showOverallColumns);
  const headerDays = days.map(day => `<th>${escapeHtml(day.label)}</th>`).join("");
  const extraHeader = showOverallColumns ? `<th>全合計距離</th><th>全出勤日数</th><th>全平均距離</th>` : "";

  const bodyRows = drivers.map(driver => {
    const sectionValues = days.map(day => Number(driver.byDate.get(day.key) || 0));
    const sectionDistance = Number(sectionValues.reduce((sum, value) => sum + value, 0).toFixed(1));
    const sectionWorkedDays = sectionValues.filter(value => value > 0).length;
    const sectionAvgDistance = sectionWorkedDays ? Number((sectionDistance / sectionWorkedDays).toFixed(1)) : 0;
    const cells = sectionValues.map(value => `<td class="mileage-day-cell">${value > 0 ? `${value.toFixed(1)}km` : "-"}</td>`).join("");
    const extraOverallCells = showOverallColumns ? `
        <td>${Number(driver.total_distance_km || 0).toFixed(1)}km</td>
        <td>${Number(driver.worked_days || 0)}日</td>
        <td>${Number(driver.avg_distance_km || 0).toFixed(1)}km</td>` : "";

    return `<tr>
      <td class="mileage-driver-cell">${escapeHtml(driver.driver_name || "-")}</td>
      ${cells}
      <td>${sectionDistance.toFixed(1)}km</td>
      <td>${sectionWorkedDays}日</td>
      <td>${sectionAvgDistance.toFixed(1)}km</td>
      ${extraOverallCells}
    </tr>`;
  }).join("");

  const dailyTotals = days.map(day => Number(drivers.reduce((sum, driver) => sum + Number(driver.byDate.get(day.key) || 0), 0).toFixed(1)));
  const dailyWorkedCounts = days.map(day => drivers.reduce((count, driver) => count + (Number(driver.byDate.get(day.key) || 0) > 0 ? 1 : 0), 0));
  const dailyAverages = dailyTotals.map((total, index) => dailyWorkedCounts[index] > 0 ? Number((total / dailyWorkedCounts[index]).toFixed(1)) : 0);

  const sectionTotalDistance = Number(dailyTotals.reduce((sum, value) => sum + Number(value || 0), 0).toFixed(1));
  const sectionWorkedDays = dailyWorkedCounts.reduce((sum, value) => sum + Number(value || 0), 0);
  const sectionAverageDistance = sectionWorkedDays ? Number((sectionTotalDistance / sectionWorkedDays).toFixed(1)) : 0;

  const overallTotalDistance = Number(drivers.reduce((sum, driver) => sum + Number(driver.total_distance_km || 0), 0).toFixed(1));
  const overallWorkedDays = drivers.reduce((sum, driver) => sum + Number(driver.worked_days || 0), 0);
  const overallAvgDistance = overallWorkedDays ? Number((overallTotalDistance / overallWorkedDays).toFixed(1)) : 0;

  const footerTotalCells = dailyTotals.map(value => `<td class="mileage-day-cell mileage-summary-cell">${value > 0 ? `${value.toFixed(1)}km` : "-"}</td>`).join("");
  const footerAverageCells = dailyAverages.map(value => `<td class="mileage-day-cell mileage-summary-cell">${value > 0 ? `${value.toFixed(1)}km` : "-"}</td>`).join("");

  const footerOverallCells = showOverallColumns ? `
      <td class="mileage-summary-cell">${overallTotalDistance.toFixed(1)}km</td>
      <td class="mileage-summary-cell">${overallWorkedDays}日</td>
      <td class="mileage-summary-cell">${overallAvgDistance.toFixed(1)}km</td>` : "";

  const footerAverageOverallCells = showOverallColumns ? `
      <td class="mileage-summary-cell">${drivers.length ? (overallTotalDistance / drivers.length).toFixed(1) : "0.0"}km</td>
      <td class="mileage-summary-cell">${drivers.length ? (overallWorkedDays / drivers.length).toFixed(1) : "0.0"}日</td>
      <td class="mileage-summary-cell">${overallAvgDistance.toFixed(1)}km</td>` : "";

  const colspan = days.length + 4 + (showOverallColumns ? 3 : 0);

  return `<div class="mileage-matrix-section">
    <div class="grouped-hour-title">${escapeHtml(title)}</div>
    <div class="table-wrap mileage-matrix-wrap">
      <table class="matrix-table mileage-matrix-table">
        <thead>
          <tr>
            <th>ドライバー</th>
            ${headerDays}
            <th>合計距離</th>
            <th>合計出勤数</th>
            <th>1日平均距離</th>
            ${extraHeader}
          </tr>
        </thead>
        <tbody>
          ${bodyRows || `<tr><td colspan="${colspan}" class="muted">データがありません</td></tr>`}
          <tr class="mileage-summary-row">
            <td class="mileage-driver-cell">日別合計</td>
            ${footerTotalCells}
            <td class="mileage-summary-cell">${sectionTotalDistance.toFixed(1)}km</td>
            <td class="mileage-summary-cell">${sectionWorkedDays}日</td>
            <td class="mileage-summary-cell">${sectionAverageDistance.toFixed(1)}km</td>
            ${footerOverallCells}
          </tr>
          <tr class="mileage-summary-row">
            <td class="mileage-driver-cell">日別平均</td>
            ${footerAverageCells}
            <td class="mileage-summary-cell">${drivers.length ? (sectionTotalDistance / drivers.length).toFixed(1) : "0.0"}km</td>
            <td class="mileage-summary-cell">-</td>
            <td class="mileage-summary-cell">${sectionAverageDistance.toFixed(1)}km</td>
            ${footerAverageOverallCells}
          </tr>
        </tbody>
      </table>
    </div>
  </div>`;
}

function renderMileageReportTable(rows) {
  if (!els.mileageReportTableWrap) return;

  if (!rows.length) {
    els.mileageReportTableWrap.innerHTML = `<div class="muted" style="padding:14px;">対象期間の走行実績はありません</div>`;
    return;
  }

  const startDate = els.mileageReportStartDate?.value || todayStr();
  const endDate = els.mileageReportEndDate?.value || todayStr();
  const calendar = buildMileageCalendarRows(rows, startDate, endDate);

  let html = `<div class="grouped-plan-list mileage-report-grid">`;
  html += `<div class="grouped-hour-title">${escapeHtml(calendar.monthLabel)} / ${escapeHtml(startDate.replaceAll("-", "/"))} ～ ${escapeHtml(endDate.replaceAll("-", "/"))}</div>`;
  html += renderMileageMatrixSection("1日～16日", calendar.firstHalfDays, calendar.drivers);
  html += renderMileageMatrixSection("17日～月末", calendar.secondHalfDays, calendar.drivers, { showOverallColumns: true });
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
function buildMileageCsvSectionRows(title, days, drivers, options = {}) {
  const showOverallColumns = Boolean(options.showOverallColumns);
  const rows = [];

  rows.push([title]);

  const header = [
    "ドライバー",
    ...days.map(day => day.label),
    "合計距離",
    "合計出勤数",
    "1日平均距離"
  ];
  if (showOverallColumns) {
    header.push("全合計距離", "全出勤日数", "全平均距離");
  }
  rows.push(header);

  drivers.forEach(driver => {
    const sectionValues = days.map(day => Number(driver.byDate.get(day.key) || 0));
    const sectionDistance = Number(sectionValues.reduce((sum, value) => sum + value, 0).toFixed(1));
    const sectionWorkedDays = sectionValues.filter(value => value > 0).length;
    const sectionAvgDistance = sectionWorkedDays > 0 ? Number((sectionDistance / sectionWorkedDays).toFixed(1)) : 0;

    const row = [
      driver.driver_name || "-",
      ...sectionValues.map(value => value > 0 ? Number(value.toFixed(1)) : ""),
      sectionDistance,
      sectionWorkedDays,
      sectionAvgDistance
    ];

    if (showOverallColumns) {
      row.push(
        Number(Number(driver.total_distance_km || 0).toFixed(1)),
        Number(driver.worked_days || 0),
        Number(Number(driver.avg_distance_km || 0).toFixed(1))
      );
    }

    rows.push(row);
  });

  const dailyTotals = days.map(day =>
    Number(drivers.reduce((sum, driver) => sum + Number(driver.byDate.get(day.key) || 0), 0).toFixed(1))
  );
  const dailyWorkedCounts = days.map(day =>
    drivers.reduce((count, driver) => count + (Number(driver.byDate.get(day.key) || 0) > 0 ? 1 : 0), 0)
  );
  const dailyAverages = dailyTotals.map((total, index) =>
    dailyWorkedCounts[index] > 0 ? Number((total / dailyWorkedCounts[index]).toFixed(1)) : ""
  );

  const sectionTotalDistance = Number(dailyTotals.reduce((sum, value) => sum + Number(value || 0), 0).toFixed(1));
  const sectionWorkedDays = dailyWorkedCounts.reduce((sum, value) => sum + Number(value || 0), 0);
  const sectionAverageDistance = sectionWorkedDays > 0 ? Number((sectionTotalDistance / sectionWorkedDays).toFixed(1)) : 0;

  const overallTotalDistance = Number(drivers.reduce((sum, driver) => sum + Number(driver.total_distance_km || 0), 0).toFixed(1));
  const overallWorkedDays = drivers.reduce((sum, driver) => sum + Number(driver.worked_days || 0), 0);
  const overallAvgDistance = overallWorkedDays > 0 ? Number((overallTotalDistance / overallWorkedDays).toFixed(1)) : 0;

  const totalRow = [
    "日別合計",
    ...dailyTotals.map(v => v || ""),
    sectionTotalDistance,
    sectionWorkedDays,
    sectionAverageDistance
  ];
  if (showOverallColumns) totalRow.push(overallTotalDistance, overallWorkedDays, overallAvgDistance);
  rows.push(totalRow);

  const avgRow = [
    "日別平均",
    ...dailyAverages,
    drivers.length ? Number((sectionTotalDistance / drivers.length).toFixed(1)) : 0,
    "",
    sectionAverageDistance
  ];
  if (showOverallColumns) {
    avgRow.push(
      drivers.length ? Number((overallTotalDistance / drivers.length).toFixed(1)) : 0,
      drivers.length ? Number((overallWorkedDays / drivers.length).toFixed(1)) : 0,
      overallAvgDistance
    );
  }
  rows.push(avgRow);

  return rows;
}

function buildMileageMatrixCsvRows(rows, startDate, endDate) {
  const calendar = buildMileageCalendarRows(rows, startDate, endDate);
  const result = [];
  result.push([calendar.monthLabel, `${startDate.replaceAll("-", "/")} ～ ${endDate.replaceAll("-", "/")}`]);
  result.push([]);
  result.push(...buildMileageCsvSectionRows("1日～16日", calendar.firstHalfDays, calendar.drivers));
  result.push([]);
  result.push(...buildMileageCsvSectionRows("17日～月末", calendar.secondHalfDays, calendar.drivers, {
    showOverallColumns: true
  }));
  return result;
}

async function exportDriverMileageReportCsv() {
  if (!currentMileageExportRows.length) {
    await previewDriverMileageReport();
    if (!currentMileageExportRows.length) return;
  }

  const startDate = els.mileageReportStartDate?.value || todayStr();
  const endDate = els.mileageReportEndDate?.value || todayStr();

  const aoa = buildMileageMatrixCsvRows(currentMileageExportRows, startDate, endDate);
  const csv = aoa.map(row => row.map(value => csvEscape(value ?? "")).join(",")).join("\n");
  downloadTextFile(`driver_mileage_matrix_${startDate}_${endDate}.csv`, csv, "text/csv;charset=utf-8");
}

async function exportDriverMileageReportXlsx() {
  return exportDriverMileageReportCsv();
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
        travel_minutes,
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
  renderOperationAndSimulationUI();
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


function getMatrixLegendHtml() {
  return `
    <div class="matrix-legend" aria-label="色の説明">
      <span class="matrix-legend-item"><span class="matrix-dot pending"></span>黄色 = 未完了</span>
      <span class="matrix-legend-item"><span class="matrix-dot done"></span>緑 = 完了</span>
      <span class="matrix-legend-item"><span class="matrix-dot cancel"></span>赤 = キャンセル</span>
    </div>
  `;
}

function getPlanLinkedActualStatus(planRow) {
  const linked = currentActualsCache.find(
    item =>
      Number(item.cast_id) === Number(planRow.cast_id) &&
      (item.actual_date || els.actualDate?.value || todayStr()) === (planRow.plan_date || els.planDate?.value || todayStr()) &&
      Number(item.actual_hour ?? -1) === Number(planRow.plan_hour ?? -1)
  );
  if (linked) return normalizeStatus(linked.status);
  return normalizeStatus(planRow.status);
}

function buildMatrixMetaText(distanceKm, travelMinutes) {
  const parts = [];
  const distance = Number(distanceKm || 0);
  const minutes = Number(travelMinutes || 0);
  if (Number.isFinite(distance) && distance > 0) parts.push(`${distance.toFixed(1)}km`);
  if (Number.isFinite(minutes) && minutes > 0) parts.push(`片道${minutes}分`);
  return parts.length ? ` (${parts.join(" / ")})` : "";
}

function buildMatrixNameLine(row, status, addressKey = "destination_address") {
  const safeStatus = normalizeStatus(status);
  const nameHtml = buildMapLinkHtml({
    name: row.casts?.name,
    address: row[addressKey] || row.casts?.address,
    lat: row.casts?.latitude,
    lng: row.casts?.longitude,
    className: `map-name-link matrix-name status-${safeStatus}`
  });
  const metaText = buildMatrixMetaText(row.distance_km, row.casts?.travel_minutes || row.travel_minutes);
  return `<span class="matrix-line">${nameHtml}<span class="matrix-meta">${escapeHtml(metaText)}</span></span>`;
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
    els.plansTimeAreaMatrix.innerHTML = `<div class="muted" style="padding:14px;">一覧がありません</div>`;
    return;
  }

  let html = `
    ${getMatrixLegendHtml()}
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
        html += `
          <td>
            <div class="matrix-card">
              ${rows.map(row => `
                <div class="matrix-item">
                  ${buildMatrixNameLine(row, getPlanLinkedActualStatus(row), "destination_address")}
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
        travel_minutes,
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
    else if (status === "cancel") nextPlanStatus = "cancel";
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

  const actionableItems = currentActualsCache.filter(item => {
    const status = normalizeStatus(item.status);
    return status !== "done" && status !== "cancel";
  });

  if (!actionableItems.length) {
    els.actualTableWrap.innerHTML = `<div class="muted" style="padding:14px;">配車決定で対応待ちのActualはありません</div>`;
    return;
  }

  const hours = [...new Set(actionableItems.map(x => Number(x.actual_hour ?? 0)))].sort((a, b) => a - b);
  let html = `<div class="grouped-actual-list">`;

  hours.forEach(hour => {
    const hourItems = actionableItems.filter(x => Number(x.actual_hour ?? 0) === hour);
    const groupedAreas = getGroupedAreasByDisplay(hourItems, x => x.destination_area || "無し");

    html += `<div class="grouped-section">`;
    html += `<div class="grouped-hour-title">${getHourLabel(hour)}</div>`;

    groupedAreas.forEach(({ detailArea }) => {
      const areaItems = hourItems.filter(
        item => normalizeAreaLabel(item.destination_area || "無し") === detailArea
      );

      if (!areaItems.length) return;

      html += `<div class="grouped-area-title">${getGroupedAreaHeaderHtml(detailArea)}</div>`;

      areaItems.forEach(item => {
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
    ${getMatrixLegendHtml()}
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
        html += `
          <td>
            <div class="matrix-card">
              ${rows.map(row => `
                <div class="matrix-item">
                  ${buildMatrixNameLine(row, normalizeStatus(row.status), "destination_address")}
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

  const monthKey = getMonthKey(els.dispatchDate?.value || todayStr());
  const monthlyStatsMap = getVehicleMonthlyStatsMap(currentDailyReportsCache, monthKey);

  const header = document.createElement("div");
  header.className = "vehicle-check-header";
  header.innerHTML = `
    <div class="vehicle-check-header-info"></div>
    <div class="vehicle-check-header-col">可能車両</div>
    <div class="vehicle-check-header-col">ラスト便</div>
  `;
  els.dailyVehicleChecklist.appendChild(header);

  allVehiclesCache.forEach(vehicle => {
    const stats = monthlyStatsMap.get(Number(vehicle.id)) || {
      totalDistance: 0,
      workedDays: 0,
      avgDistance: 0
    };
    const avgDistanceText = `${Number(stats.avgDistance || 0).toFixed(1)}km`;

    const row = document.createElement("div");
    row.className = "vehicle-check-item";
    row.innerHTML = `
      <div class="vehicle-check-info">
        <div class="vehicle-check-name">${escapeHtml(vehicle.driver_name || "-")}</div>
        <div class="vehicle-check-car">車両 ${escapeHtml(vehicle.plate_number || "-")}</div>
        <div class="vehicle-check-meta">担当 ${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))} / 帰宅 ${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))} / 定員 ${vehicle.seat_capacity ?? "-"} / 1日平均距離 ${avgDistanceText}</div>
      </div>
      <label class="vehicle-check-toggle vehicle-check-toggle-work">
        <input class="vehicle-check-input" type="checkbox" data-id="${vehicle.id}" ${activeVehicleIdsForToday.has(Number(vehicle.id)) ? "checked" : ""} />
        <span>可能車両</span>
      </label>
      <label class="vehicle-check-toggle vehicle-check-toggle-last">
        <input class="driver-last-trip-input" type="checkbox" data-id="${vehicle.id}" ${isDriverLastTripChecked(vehicle.id) ? "checked" : ""} />
        <span>ラスト便</span>
      </label>
    `;
    els.dailyVehicleChecklist.appendChild(row);
  });

  renderOperationAndSimulationUI();
  els.dailyVehicleChecklist.querySelectorAll(".vehicle-check-input").forEach(input => {
    input.addEventListener("change", () => {
      const id = Number(input.dataset.id);

      if (input.checked) activeVehicleIdsForToday.add(id);
      else activeVehicleIdsForToday.delete(id);

      renderDailyMileageInputs();
      renderDailyDispatchResult();
      renderDailyVehicleChecklist();
    });
  });

  els.dailyVehicleChecklist.querySelectorAll(".driver-last-trip-input").forEach(input => {
    input.addEventListener("change", () => {
      const id = Number(input.dataset.id);
      setDriverLastTripChecked(id, input.checked);

      if (input.checked && !activeVehicleIdsForToday.has(id)) {
        activeVehicleIdsForToday.add(id);
      }

      renderDailyMileageInputs();
      renderDailyDispatchResult();
      renderDailyVehicleChecklist();
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
  const dateStr = els.actualDate?.value || todayStr();
  const effectiveLastHour = getEffectiveLastHour(dateStr, orderedItems);

  function getLoad(vehicleId, hour) {
    return Number(seatLoads.get(`${vehicleId}_${hour}`) || 0);
  }

  function addLoad(vehicleId, hour) {
    const key = `${vehicleId}_${hour}`;
    seatLoads.set(key, getLoad(vehicleId, hour) + 1);
  }

  orderedItems.forEach((item, index) => {
    const hour = Number(item.actual_hour ?? 0);
    const targetArea = normalizeAreaLabel(item.destination_area || item.cluster_area || item.planned_area || item?.casts?.area || "無し");
    const isLastHour = Number(hour) === Number(effectiveLastHour);
    let candidateVehicles = vehicles.filter(v => getLoad(v.id, hour) < Number(v.seat_capacity || 4));
    if (isLastHour) {
      const existingAreasByVehicle = new Map();
      assignments
        .filter(a => Number(a.actual_hour ?? 0) === Number(hour))
        .forEach(a => {
          const assignedItem = orderedItems.find(row => Number(row.id) === Number(a.item_id));
          const area = normalizeAreaLabel(assignedItem?.destination_area || assignedItem?.cluster_area || assignedItem?.planned_area || assignedItem?.casts?.area || "無し");
          const key = Number(a.vehicle_id);
          if (!existingAreasByVehicle.has(key)) existingAreasByVehicle.set(key, []);
          if (area) existingAreasByVehicle.get(key).push(area);
        });
      const strictCandidates = candidateVehicles.filter(v => {
        const existingAreas = existingAreasByVehicle.get(Number(v.id)) || [];
        return isLastRunHardAreaConstraintSatisfied(targetArea, existingAreas, v?.home_area || "");
      });
      if (strictCandidates.length) {
        candidateVehicles = strictCandidates;
      } else {
        const directionSafeCandidates = candidateVehicles.filter(v => {
          const existingAreas = existingAreasByVehicle.get(Number(v.id)) || [];
          if (!isStrictSameDirectionArea(targetArea, existingAreas)) return false;
          return !isHardReverseForHome(targetArea, v?.home_area || "");
        });
        candidateVehicles = directionSafeCandidates;
      }
    }

    if (isLastHour && !candidateVehicles.length) return;

    let selectedVehicle =
      candidateVehicles[0] ||
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
  const defaultLastHour = getEffectiveLastHour(dateStr, assignments.map(a => ({ actual_hour: a?.actual_hour })));
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
        const existingAreas = optimizedLastHour
          .filter(
            a =>
              Number(a.actual_hour ?? 0) === Number(defaultLastHour) &&
              Number(a.vehicle_id) === Number(vehicle.id) &&
              Number(a.item_id) !== Number(target.item_id)
          )
          .map(a => {
            const row = itemMap.get(Number(a.item_id));
            return normalizeAreaLabel(row?.destination_area || row?.cluster_area || row?.planned_area || row?.casts?.area || "");
          })
          .filter(Boolean);
        if (!isLastRunHardAreaConstraintSatisfied(area, existingAreas, vehicle.home_area || "")) continue;

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


const AREA_AVERAGE_SPEED_KMH = {
  "松戸近郊": 33,
  "柏方面": 33,
  "柏の葉方面": 34,
  "流山方面": 33,
  "野田方面": 34,
  "我孫子方面": 36,
  "取手方面": 39,
  "藤代方面": 39,
  "守谷方面": 39,
  "牛久方面": 39,
  "葛飾方面": 27,
  "足立方面": 27,
  "江戸川方面": 27,
  "墨田方面": 27,
  "江東方面": 27,
  "荒川方面": 27,
  "台東方面": 27,
  "市川方面": 31,
  "船橋方面": 31,
  "鎌ヶ谷方面": 31,
  "三郷方面": 34,
  "八潮方面": 34,
  "草加方面": 34,
  "吉川方面": 34,
  "越谷方面": 34,
  "千葉方面": 31
};

function normalizeAreaSpeedLookupInput(input) {
  if (Array.isArray(input)) return getRepresentativeAreaFromRows(input);
  if (input && typeof input === "object") {
    return normalizeAreaLabel(
      input.destination_area ||
      input.planned_area ||
      input.cluster_area ||
      input.area ||
      input.home_area ||
      input.vehicle_area ||
      ""
    );
  }
  return normalizeAreaLabel(input || "");
}

function getRepresentativeAreaFromRows(rows) {
  const list = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!list.length) return "";

  const counts = new Map();
  list.forEach(row => {
    const area = normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.cluster_area || row?.casts?.area || "");
    if (!area || area === "無し") return;
    counts.set(area, Number(counts.get(area) || 0) + 1);
  });

  if (!counts.size) {
    return normalizeAreaLabel(
      list[list.length - 1]?.destination_area ||
      list[list.length - 1]?.planned_area ||
      list[list.length - 1]?.cluster_area ||
      list[list.length - 1]?.casts?.area ||
      ""
    );
  }

  const sorted = [...counts.entries()].sort((a, b) => {
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0], "ja");
  });

  return normalizeAreaLabel(sorted[0]?.[0] || "");
}

function getAreaAverageSpeedKmh(areaInput, distanceKm = 0) {
  const rawArea = normalizeAreaSpeedLookupInput(areaInput);
  const canonical = getCanonicalArea(rawArea);

  if (AREA_AVERAGE_SPEED_KMH[canonical]) {
    return Number(AREA_AVERAGE_SPEED_KMH[canonical]);
  }

  if (AREA_AVERAGE_SPEED_KMH[rawArea]) {
    return Number(AREA_AVERAGE_SPEED_KMH[rawArea]);
  }

  if (canonical && /方面$/.test(canonical)) {
    if (/取手|藤代|守谷|牛久/.test(canonical)) return 39;
    if (/吉川|三郷|八潮|草加|越谷/.test(canonical)) return 34;
    if (/市川|船橋|鎌ヶ谷|鎌ケ谷|千葉/.test(canonical)) return 31;
    if (/葛飾|足立|江戸川|墨田|江東|荒川|台東/.test(canonical)) return 27;
    if (/我孫子/.test(canonical)) return 36;
    if (/柏|流山|野田|松戸/.test(canonical)) return 33;
  }

  const km = Number(distanceKm || 0);
  if (km <= 10) return 30;
  if (km <= 25) return 33;
  return 36;
}

function estimateTravelMinutesByAreaSpeed(distanceKm, areaInput) {
  const km = Math.max(0, Number(distanceKm || 0));
  const speed = getAreaAverageSpeedKmh(areaInput, km);
  if (!speed) return 0;
  return Math.round((km / speed) * 60);
}

function estimateFallbackTravelMinutes(distanceKm, areaInput = "") {
  return estimateTravelMinutesByAreaSpeed(distanceKm, areaInput);
}

function getRowOneWayTravelMinutes(row, fallbackDistanceKm = null, fallbackArea = "") {
  const stored = getCastTravelMinutesValue(row?.casts || row);
  if (stored > 0) return stored;

  const km = Number(fallbackDistanceKm ?? row?.distance_km ?? row?.casts?.distance_km ?? 0);
  const area = fallbackArea || normalizeAreaLabel(row?.destination_area || row?.casts?.area || "");
  return estimateFallbackTravelMinutes(km, area);
}

function getRowsOutboundMinutes(rows) {
  const ordered = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!ordered.length) return 0;
  const representativeArea = getRepresentativeAreaFromRows(ordered);
  const routeDistanceKm = Number(calculateRouteDistanceGlobal(ordered) || 0);
  const lastRow = ordered[ordered.length - 1] || {};
  const storedLast = getCastTravelMinutesValue(lastRow?.casts || lastRow);
  if (storedLast > 0) return storedLast;
  return estimateFallbackTravelMinutes(routeDistanceKm, representativeArea);
}

function getRowsReturnMinutes(rows) {
  const ordered = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!ordered.length) return 0;
  const lastRow = ordered[ordered.length - 1] || {};
  const area = normalizeAreaLabel(lastRow.destination_area || lastRow.casts?.area || getRepresentativeAreaFromRows(ordered) || "");
  return getRowOneWayTravelMinutes(lastRow, lastRow.distance_km || 0, area);
}

function getRowsTravelTimeSummary(rows) {
  const ordered = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!ordered.length) {
    return { outboundMinutes: 0, returnMinutes: 0, totalMinutes: 0, sendOnlyMinutes: 0, stopCount: 0 };
  }
  const outboundMinutes = Math.round(getRowsOutboundMinutes(ordered));
  const returnMinutes = Math.round(getRowsReturnMinutes(ordered));
  const stopCount = ordered.length;
  const sendOnlyMinutes = Math.round(outboundMinutes + stopCount);
  const totalMinutes = Math.round(outboundMinutes + returnMinutes + stopCount);
  return { outboundMinutes, returnMinutes, totalMinutes, sendOnlyMinutes, stopCount };
}

function formatClockTimeFromMinutesGlobal(totalMinutes) {
  const safe = Math.max(0, Math.round(Number(totalMinutes || 0)));
  const h = Math.floor(safe / 60) % 24;
  const m = safe % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

function getDistanceZoneInfoGlobal(distanceKm, areaInput = "") {
  const km = Number(distanceKm || 0);
  const speedKmh = getAreaAverageSpeedKmh(areaInput, km);
  const canonical = getCanonicalArea(normalizeAreaSpeedLookupInput(areaInput));
  return {
    key: canonical || (km <= 10 ? "short" : km <= 25 ? "middle" : "long"),
    label: canonical || "-",
    speedKmh
  };
}function calculateRouteDistanceGlobal(items) {
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
    const representativeArea = getRepresentativeAreaFromRows(rows);
  const returnArea = normalizeAreaLabel(lastRow.destination_area || lastRow.casts?.area || representativeArea || "");
  const primaryZone = getDistanceZoneInfoGlobal(Math.max(routeDistanceKm, returnDistanceKm), representativeArea);
  const timeSummary = getRowsTravelTimeSummary(rows);

  let departDelayMinutes = 20;
  if (baseHour === 3) departDelayMinutes = 18;
  else if (baseHour === 4) departDelayMinutes = 12;
  else if (baseHour >= 5) departDelayMinutes = 8;

  const outboundMinutes = timeSummary.outboundMinutes;
  const returnMinutes = timeSummary.returnMinutes;
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
    const singleArea = normalizeAreaLabel(rows[0]?.destination_area || rows[0]?.casts?.area || representativeArea || "");
    const singleOutbound = getRowOneWayTravelMinutes(rows[0], singleRouteDistanceKm, singleArea);
    const singleReturn = getRowOneWayTravelMinutes(rows[0], singleReturnDistanceKm, singleArea);
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
}function optimizeAssignments(items, vehicles, monthlyMap) {
  const workingVehicles = vehicles.filter(v => v.status !== "maintenance");
  const manualLastVehicleId = getManualLastVehicleId();
  const clusters = buildDispatchClusters(items);
  const assignments = [];
  const itemById = new Map((items || []).map(item => [Number(item.id), item]));
  const itemMetaMap = new Map((items || []).map(item => [Number(item.id), {
    distanceKm: Number(item.distance_km || 0),
    area: normalizeAreaLabel(item.destination_area || item.cluster_area || "無し")
  }]));

  if (!workingVehicles.length || !clusters.length) return assignments;

  const LONG_BUNDLE_ROUND_TRIP_MINUTES = 55;
  const actualLastHour = clusters.reduce((max, cluster) => Math.max(max, Number(cluster?.hour ?? -Infinity)), -Infinity);
  const effectiveLastHour = getEffectiveLastHour(els.actualDate?.value || todayStr(), clusters.map(c => ({ hour: c?.hour })));
  const vehicleUsage = new Map();
  const vehicleLongClusterLocks = new Map();

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

  function getVehicleLockedAreas(vehicleId) {
    return [...(vehicleLongClusterLocks.get(Number(vehicleId)) || [])];
  }

  function lockVehicleClusterArea(vehicleId, area) {
    const normalized = normalizeAreaLabel(area);
    if (!normalized || normalized === "無し") return;
    const key = Number(vehicleId);
    const existing = vehicleLongClusterLocks.get(key) || [];
    if (!existing.includes(normalized)) {
      vehicleLongClusterLocks.set(key, [...existing, normalized]);
    }
  }

  function shouldClusterLock(cluster) {
    return Number(cluster?.count || 0) >= 2 && getClusterRoundTripMinutes(cluster) >= LONG_BUNDLE_ROUND_TRIP_MINUTES;
  }

  function isClusterLockHardMismatch(clusterArea, lockedAreas = []) {
    const normalized = normalizeAreaLabel(clusterArea);
    const areas = Array.isArray(lockedAreas) ? lockedAreas.filter(Boolean).map(normalizeAreaLabel) : [];
    if (!areas.length || !normalized || normalized === "無し") return false;
    if (hasHardReverseMix(normalized, areas)) return true;

    const canonical = getCanonicalArea(normalized);
    const group = getAreaDisplayGroup(normalized);
    const sameCanonical = areas.some(area => getCanonicalArea(area) && canonical && getCanonicalArea(area) === canonical);
    const sameGroup = areas.some(area => getAreaDisplayGroup(area) === group);
    if (sameCanonical || sameGroup) return false;

    const bestAffinity = Math.max(...areas.map(area => getAreaAffinityScore(normalized, area)), 0);
    const bestDirection = Math.max(...areas.map(area => getDirectionAffinityScore(normalized, area)), -999);
    return bestAffinity < 64 && bestDirection < 24;
  }

  function getClusterLockPenalty(clusterArea, lockedAreas = []) {
    const normalized = normalizeAreaLabel(clusterArea);
    const areas = Array.isArray(lockedAreas) ? lockedAreas.filter(Boolean).map(normalizeAreaLabel) : [];
    if (!areas.length || !normalized || normalized === "無し") return 0;
    if (hasHardReverseMix(normalized, areas)) return 1200;

    const canonical = getCanonicalArea(normalized);
    const group = getAreaDisplayGroup(normalized);
    const bestAffinity = Math.max(...areas.map(area => getAreaAffinityScore(normalized, area)), 0);
    const bestDirection = Math.max(...areas.map(area => getDirectionAffinityScore(normalized, area)), -999);
    const sameCanonical = areas.some(area => getCanonicalArea(area) && canonical && getCanonicalArea(area) === canonical);
    const sameGroup = areas.some(area => getAreaDisplayGroup(area) === group);

    if (sameCanonical) return -260;
    if (sameGroup) return -140;
    if (bestAffinity >= 72 && bestDirection >= 20) return -80;
    if (bestAffinity >= 64 && bestDirection >= 18) return 120;
    if (bestAffinity >= 56) return 260;
    return 520;
  }

  function getIdleVehicleCountForHour(hour) {
    return workingVehicles.filter(vehicle => getHourLoad(vehicle.id, hour) === 0).length;
  }

  function shouldPreferSpread(cluster, routeFlowScore, routeContinuityPenalty, sameHourLoad) {
    const roundTripMinutes = getClusterRoundTripMinutes(cluster);
    if (roundTripMinutes >= LONG_BUNDLE_ROUND_TRIP_MINUTES) return false;
    if (cluster.count <= 1) {
      const strongRouteLink = routeFlowScore >= 78 && routeContinuityPenalty <= 85;
      if (strongRouteLink) return false;
      return true;
    }
    if (sameHourLoad <= 0) return true;
    const strongRouteLink = routeFlowScore >= 90 && routeContinuityPenalty <= 45;
    const reasonableRideShare = routeFlowScore >= 70 && routeContinuityPenalty <= 60;
    if (strongRouteLink) return false;
    if (cluster.count <= 2 && reasonableRideShare) return false;
    return true;
  }

  function getClusterRepresentativeTravelMinutes(cluster) {
    const items = Array.isArray(cluster?.items) ? cluster.items : [];
    const minutesList = items
      .map(item => getStoredTravelMinutes(item?.casts?.travel_minutes || item?.travel_minutes))
      .filter(val => Number(val) > 0);
    if (minutesList.length) return Math.max(...minutesList);
    return Math.max(1, estimateTravelMinutesByDistance(Number(cluster?.totalDistance || 0), normalizeAreaLabel(cluster?.area || "")));
  }

  function getClusterRoundTripMinutes(cluster) {
    return getClusterRepresentativeTravelMinutes(cluster) * 2;
  }

  function getBundleCompatibilityScore(targetArea, existingAreas = []) {
    const targetCanonical = getCanonicalArea(targetArea);
    const targetGroup = getAreaDisplayGroup(targetArea);
    const areas = Array.isArray(existingAreas) ? existingAreas.filter(Boolean) : [];
    let best = 0;

    for (const rawArea of areas) {
      const area = normalizeAreaLabel(rawArea);
      const canonical = getCanonicalArea(area);
      const group = getAreaDisplayGroup(area);
      const affinity = getAreaAffinityScore(targetArea, area);
      const direction = getDirectionAffinityScore(targetArea, area);
      let score = 0;

    if (targetCanonical && canonical && targetCanonical === canonical) score += 180;
    else if (targetGroup && group && targetGroup === group) score += 130;
    else if (affinity >= 72) score += 100;
    else if (affinity >= 54) score += 62;

      if (direction >= 72) score += 18;
      else if (direction >= 28) score += 8;
      else if (direction <= -38) score -= 60;

      best = Math.max(best, score);
    }

    return best;
  }

  function shouldForceBundleCluster(cluster, existingAreas = []) {
    const roundTripMinutes = getClusterRoundTripMinutes(cluster);
    const bundleScore = getBundleCompatibilityScore(normalizeAreaLabel(cluster?.area || ""), existingAreas);
    return roundTripMinutes >= LONG_BUNDLE_ROUND_TRIP_MINUTES && bundleScore >= 70;
  }

  function getDistanceZoneForAi(distanceKm) {
    const km = Number(distanceKm || 0);
    if (km <= 10) return "near";
    if (km <= 25) return "mid";
    return "far";
  }

  function isShortLayerDistance(distanceKm) {
    return Number(distanceKm || 0) < 15;
  }

  function isLongLayerDistance(distanceKm) {
    return Number(distanceKm || 0) > 35;
  }

  function getHourAssignedDistanceValues(vehicleId, hour) {
    return assignments
      .filter(a => Number(a.vehicle_id) === Number(vehicleId) && Number(a.actual_hour) === Number(hour))
      .map(a => Number(itemMetaMap.get(Number(a.item_id))?.distanceKm || a.distance_km || 0))
      .filter(v => v > 0);
  }

  function getHourAssignedRows(vehicleId, hour) {
    return assignments
      .filter(a => Number(a.vehicle_id) === Number(vehicleId) && Number(a.actual_hour) === Number(hour))
      .map(a => itemById.get(Number(a.item_id)))
      .filter(Boolean);
  }

  function getDistinctAreaGroupCountForHour(hour) {
    const groups = new Set(
      clusters
        .filter(cluster => Number(cluster?.hour) === Number(hour))
        .map(cluster => getAreaDisplayGroup(normalizeAreaLabel(cluster?.area || "無し")))
        .filter(Boolean)
    );
    return groups.size;
  }

  function hasAreaVehicleShortageForHour(hour) {
    return getDistinctAreaGroupCountForHour(hour) > workingVehicles.length;
  }

  function buildForecastRowsForVehicle(vehicleId, hour, extraItems = []) {
    const baseRows = getHourAssignedRows(vehicleId, hour);
    const rows = [...baseRows, ...(Array.isArray(extraItems) ? extraItems.filter(Boolean) : [])];
    return sortItemsByNearestRoute(rows);
  }

  function getVehicleCandidateReturnMetrics(vehicle, hour, extraItems = []) {
    const rows = buildForecastRowsForVehicle(vehicle?.id, hour, extraItems);
    const forecast = calcVehicleRotationForecast(vehicle, rows);
    return {
      rows,
      forecast,
      predictedReturnMinutes: Number(forecast?.predictedReturnMinutes || 0),
      extraSharedDelayMinutes: Number(forecast?.extraSharedDelayMinutes || 0),
      routeDistanceKm: Number(forecast?.routeDistanceKm || 0)
    };
  }

  function getVehicleItemProximityScore(vehicle, hour, item) {
    const point = getItemLatLng(item);
    if (!point) return 0;

    const existingRows = getHourAssignedRows(vehicle?.id, hour);
    if (!existingRows.length) {
      const homeArea = normalizeAreaLabel(vehicle?.home_area || "");
      const area = normalizeAreaLabel(item?.destination_area || item?.cluster_area || "無し");
      return Math.max(0, getAreaAffinityScore(area, homeArea)) * 0.35;
    }

    let bestKm = Infinity;
    existingRows.forEach(row => {
      const rowPoint = getItemLatLng(row);
      if (!rowPoint) return;
      const km = estimateRoadKmBetweenPoints(point.lat, point.lng, rowPoint.lat, rowPoint.lng);
      if (km < bestKm) bestKm = km;
    });

    if (!Number.isFinite(bestKm)) {
      return 0;
    }

    if (bestKm <= 4) return 140;
    if (bestKm <= 8) return 96;
    if (bestKm <= 12) return 58;
    if (bestKm <= 18) return 24;
    return -Math.min(90, Math.round((bestKm - 18) * 3.5));
  }

  function isEastNearArea(areaInput = "") {
    const area = normalizeAreaLabel(areaInput);
    return ["市川", "船橋", "習志野", "鎌ヶ谷", "鎌ケ谷", "新鎌ヶ谷"].some(k => area.includes(k));
  }

  function isNorthEastLongArea(areaInput = "") {
    const area = normalizeAreaLabel(areaInput);
    return ["牛久", "ひたち野", "取手", "藤代", "宮和田", "守谷", "我孫子"].some(k => area.includes(k));
  }

  function isKashiwaFlexibleArea(areaInput = "") {
    const area = normalizeAreaLabel(areaInput);
    return ["柏", "あけぼの"].some(k => area.includes(k));
  }

  function getLongMixedPenalty(clusterDistanceKm, existingDistanceValues = [], clusterArea = "", existingAreas = []) {
    const clusterKm = Number(clusterDistanceKm || 0);
    const allDistances = [...(existingDistanceValues || []).map(v => Number(v || 0)).filter(v => v > 0), clusterKm].filter(v => v > 0);
    const shortCount = allDistances.filter(isShortLayerDistance).length;
    const longCount = allDistances.filter(isLongLayerDistance).length;
    if (!shortCount || !longCount) return 0;

    const normalizedArea = normalizeAreaLabel(clusterArea);
    const areas = [normalizedArea, ...(existingAreas || []).map(normalizeAreaLabel)].filter(Boolean);
    const hasEast = areas.some(isEastNearArea);
    const hasNorthEastLong = areas.some(isNorthEastLongArea);
    const hasKashiwa = areas.some(isKashiwaFlexibleArea);

    let penalty = 0;
    if (longCount >= 1 && shortCount >= 2) penalty += 220;
    else if (longCount >= 1 && shortCount >= 1) penalty += 120;

    if (hasEast && hasNorthEastLong) penalty += 260;
    if (hasEast && hasKashiwa && hasNorthEastLong) penalty += 80;

    return penalty;
  }

  function getLongIsolationBonus(clusterDistanceKm, existingDistanceValues = [], clusterArea = "", existingAreas = []) {
    const clusterKm = Number(clusterDistanceKm || 0);
    if (!isLongLayerDistance(clusterKm)) return 0;
    const distances = (existingDistanceValues || []).map(v => Number(v || 0)).filter(v => v > 0);
    if (!distances.length) return 0;

    const normalizedArea = normalizeAreaLabel(clusterArea);
    const areas = (existingAreas || []).map(normalizeAreaLabel).filter(Boolean);
    const allAreas = [normalizedArea, ...areas];

    const shortCount = distances.filter(isShortLayerDistance).length;
    const hasEast = allAreas.some(isEastNearArea);
    const hasNorthEastLong = allAreas.some(isNorthEastLongArea);

    if (shortCount >= 2 && hasEast && hasNorthEastLong) return 180;
    if (shortCount >= 1 && hasEast && hasNorthEastLong) return 90;
    return 0;
  }

  function getBaseDispatchDelayMinutes(hour) {
    const h = Number(hour || 0);
    if (h <= 1) return 20;
    if (h === 2) return 20;
    if (h === 3) return 18;
    if (h === 4) return 12;
    return 8;
  }

  function estimateTravelMinutesByDistance(distanceKm, areaInput = "") {
    return estimateTravelMinutesByAreaSpeed(distanceKm, areaInput);
  }

  function estimateDropoffMinutes(stopCount) {
    return Math.max(1, Number(stopCount || 1)) * 1;
  }

  function estimateRotationReadyMinutes(hour, distanceKm, stopCount, areaInput = "") {
    const travelOut = estimateTravelMinutesByDistance(distanceKm, areaInput);
    const dropoff = estimateDropoffMinutes(stopCount);
    const returnTrip = estimateTravelMinutesByDistance(distanceKm, areaInput);
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

  function getRotationPredictionScore(hour, distanceKm, stopCount, sameHourLoad, routeFlowScore, routeContinuityPenalty, idleVehicleCount, areaInput = "") {
    const predictedReadyMinutes = estimateRotationReadyMinutes(hour, distanceKm, stopCount, areaInput);
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

  function getNormalRunReturnPenalty(hour, addedDistance, stopCount, sameHourLoad, routeFlowScore, routeContinuityPenalty, idleVehicleCount, areaInput = "") {
    const rotation = getRotationPredictionScore(
      hour,
      addedDistance,
      stopCount,
      sameHourLoad,
      routeFlowScore,
      routeContinuityPenalty,
      idleVehicleCount,
      areaInput
    );
    return rotation.score;
  }

  function formatClockTimeFromMinutes(totalMinutes) {
    const safe = Math.max(0, Math.round(Number(totalMinutes || 0)));
    const h = Math.floor(safe / 60) % 24;
    const m = safe % 60;
    return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
  }

  function getDistanceZoneInfo(distanceKm, areaInput = "") {
    const km = Number(distanceKm || 0);
    const speedKmh = getAreaAverageSpeedKmh(areaInput, km);
    const canonical = getCanonicalArea(normalizeAreaSpeedLookupInput(areaInput));
    return {
      key: canonical || (km <= 10 ? "short" : km <= 25 ? "middle" : "long"),
      label: canonical || "-",
      speedKmh
    };
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
        const representativeArea = getRepresentativeAreaFromRows(rows);
    const returnArea = normalizeAreaLabel(lastRow.destination_area || lastRow.casts?.area || representativeArea || "");
    const primaryZone = getDistanceZoneInfo(Math.max(routeDistanceKm, returnDistanceKm), representativeArea);
  
    const departDelayMinutes = getExpectedDepartureDelayMinutes(baseHour);
    const outboundMinutes = estimateTravelMinutesByDistance(routeDistanceKm, representativeArea);
    const returnMinutes = estimateTravelMinutesByDistance(returnDistanceKm, returnArea || representativeArea);
    const dropoffMinutes = rows.length * 1;
  
    const predictedDepartureAbs = baseHour * 60 + departDelayMinutes;
    const predictedReturnAbs = predictedDepartureAbs + outboundMinutes + dropoffMinutes + returnMinutes;
    const predictedReadyAbs = predictedReturnAbs + 1;
  
    let extraSharedDelayMinutes = 0;
    if (rows.length >= 2) {
      const firstOnly = [rows[0]];
      const singleRouteDistanceKm = Number(calculateRouteDistance(firstOnly) || rows[0].distance_km || 0);
      const singleReturnDistanceKm = Number(rows[0].distance_km || 0);
      const singleArea = normalizeAreaLabel(rows[0]?.destination_area || rows[0]?.casts?.area || representativeArea || "");
      const singleOutbound = estimateTravelMinutesByDistance(singleRouteDistanceKm, singleArea);
      const singleReturn = estimateTravelMinutesByDistance(singleReturnDistanceKm, singleArea);
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
    const defaultLastHour = effectiveLastHour;
    const isDefaultLastHourCluster =
      Number(cluster.hour) === Number(defaultLastHour) &&
      Number(cluster.hour) === Number(actualLastHour);

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
        const existingHourDistanceValues = getHourAssignedDistanceValues(vehicle.id, cluster.hour);
        const hardReverseMixBlocked = existingHourAreas.length && hasHardReverseMix(normalizedClusterArea, existingHourAreas);
        const lockedAreas = getVehicleLockedAreas(vehicle.id);
        const clusterLockPenalty = getClusterLockPenalty(normalizedClusterArea, lockedAreas);
        const areaMixGuardScore = getAreaMixGuardScore(normalizedClusterArea, existingHourAreas, isLastRun || isDefaultLastHourCluster);
        const routeFlowScore = getRouteFlowVehicleScore(normalizedClusterArea, existingHourAreas, homeArea);
        const routeContinuityPenalty = getRouteContinuityPenalty(normalizedClusterArea, existingHourAreas, homeArea);
        if ((isLastRun || isDefaultLastHourCluster) && !isLastRunHardAreaConstraintSatisfied(normalizedClusterArea, existingHourAreas, homeArea)) {
          return null;
        }
        if (ENFORCE_AREA_PRIORITY_STRICT) {
          const areaPriorityWeak = strictHomeScore < 50 && directionAffinity < 24 && vehicleAreaMatch < 55;
          const mixedDirectionRisk = existingHourAreas.length > 0 && routeContinuityPenalty >= 260;
          if (areaPriorityWeak || mixedDirectionRisk) return null;
        }
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
        score -= routeFlowScore * (isLastRun || isDefaultLastHourCluster ? 2.25 : 1.55);
        score += routeContinuityPenalty * (isLastRun || isDefaultLastHourCluster ? 2.45 : 1.55);
        score += areaMixGuardScore;
        score += clusterLockPenalty;
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

        // 通常便は、往復55分以上なら同方向クラスターを優先し、近距離だけ分散を許容する
        if (!(isLastRun || isDefaultLastHourCluster)) {
          const idleVehicleCount = getIdleVehicleCountForHour(cluster.hour);
          const roundTripMinutes = getClusterRoundTripMinutes(cluster);
          const bundleCompatibility = getBundleCompatibilityScore(normalizedClusterArea, existingHourAreas);
          const forceBundle = shouldForceBundleCluster(cluster, existingHourAreas);
          const rotationPrediction = getRotationPredictionScore(
            cluster.hour,
            Number(cluster.totalDistance || 0),
            Number(cluster.count || 1),
            sameHourLoad,
            routeFlowScore,
            routeContinuityPenalty,
            idleVehicleCount,
            normalizedClusterArea
          );
          const preferSpread =
            (!forceBundle) &&
            (shouldPreferSpread(cluster, routeFlowScore, routeContinuityPenalty, sameHourLoad) ||
            !rotationPrediction.canShare);

          if (sameHourLoad === 0) {
            if (roundTripMinutes < LONG_BUNDLE_ROUND_TRIP_MINUTES) {
              score -= idleVehicleCount > 1 ? 150 : 95;
            } else {
              score += idleVehicleCount > 1 ? 35 : 12;
            }
          } else if (preferSpread && idleVehicleCount > 0) {
            score += 165 + sameHourLoad * 38;
          } else if (!preferSpread && rotationPrediction.canShare) {
            score -= Math.min(routeFlowScore, 120) * 0.40;
            score += routeContinuityPenalty * 0.15;
            score -= 18;
          }
          if (sameHourLoad > 0 && routeFlowScore >= 78 && routeContinuityPenalty <= 95) {
            score -= 120;
          }

          score -= bundleCompatibility * (roundTripMinutes >= LONG_BUNDLE_ROUND_TRIP_MINUTES ? 1.45 : 0.65);
          if (forceBundle && sameHourLoad > 0) {
            score -= 180;
          }
          if (roundTripMinutes < LONG_BUNDLE_ROUND_TRIP_MINUTES && sameHourLoad > 0 && bundleCompatibility < 50 && idleVehicleCount > 0) {
            score += 75;
          }

          score += getNormalRunReturnPenalty(
            cluster.hour,
            Number(cluster.totalDistance || 0),
            Number(cluster.count || 1),
            sameHourLoad,
            routeFlowScore,
            routeContinuityPenalty,
            idleVehicleCount,
            normalizedClusterArea
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
        if (hardReverseMixBlocked) {
          score += isLastRun || isDefaultLastHourCluster ? 880 : 620;
        }

        // v6.9.6: ロング案件が近距離群へ混ざるのを強く避ける
        score += getLongMixedPenalty(
          Number(cluster.totalDistance || 0),
          existingHourDistanceValues,
          normalizedClusterArea,
          existingHourAreas
        );
        score += getLongIsolationBonus(
          Number(cluster.totalDistance || 0),
          existingHourDistanceValues,
          normalizedClusterArea,
          existingHourAreas
        );

        // 同一ルートの折れ返しが大きい組み合わせを避ける
        if (routeContinuityPenalty >= 260) {
          score += isLastRun || isDefaultLastHourCluster ? 280 : 180;
        } else if (routeContinuityPenalty >= 170) {
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
        score -= homePriorityWeight * 0.42;

        const driverLastTripChecked = isDriverLastTripChecked(vehicle.id);
        if (driverLastTripChecked && (isLastRun || isDefaultLastHourCluster)) {
          score -= 90;
          score -= Math.max(0, homePriorityWeight) * 0.28;
          if (strictHomeScore >= 78) score -= 48;
          else if (strictHomeScore >= 52) score -= 24;
          if (hardReverse) score += isLastRun ? 260 : 140;
        } else if (!driverLastTripChecked && hasAnyDriverLastTripChecked() && (isLastRun || isDefaultLastHourCluster)) {
          score += 12;
        }

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
      const forceBundle = shouldForceBundleCluster(cluster, bestExistingAreas);
      const clusterLocked = shouldClusterLock(cluster);
      const areaVehicleShortage = hasAreaVehicleShortageForHour(cluster.hour);
      const forceSameDirectionBundle =
        bestHourLoad > 0 &&
        bestRouteFlowScore >= 78 &&
        bestRouteContinuityPenalty <= 95;
      const keepTogether = clusterLocked || forceSameDirectionBundle || ((isLastRun || isDefaultLastHourCluster)
        ? true
        : (!areaVehicleShortage && (forceBundle || !shouldPreferSpread(cluster, bestRouteFlowScore, bestRouteContinuityPenalty, bestHourLoad) || idleVehicleCount <= 1)));

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
        if (clusterLocked || forceBundle) {
          lockVehicleClusterArea(bestVehicle.id, cluster.area);
        }
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
          const normalizedItemArea = normalizeAreaLabel(item?.destination_area || normalizedClusterArea || "無し");
          const homeArea = normalizeAreaLabel(vehicle?.home_area || "");
          const existingHourAreas = getHourAreas(vehicle.id, cluster.hour);
          const hardReverseMixBlocked = existingHourAreas.length && hasHardReverseMix(normalizedClusterArea, existingHourAreas);
          const lockedAreas = getVehicleLockedAreas(vehicle.id);
          const clusterLockPenalty = getClusterLockPenalty(normalizedClusterArea, lockedAreas);
          if (isClusterLockHardMismatch(normalizedClusterArea, lockedAreas)) return null;
          const areaMixGuardScore = getAreaMixGuardScore(normalizedItemArea, existingHourAreas, isLastRun || isDefaultLastHourCluster);
          const routeFlowScore = getRouteFlowVehicleScore(normalizedItemArea, existingHourAreas, homeArea);
          const routeContinuityPenalty = getRouteContinuityPenalty(normalizedItemArea, existingHourAreas, homeArea);
          if ((isLastRun || isDefaultLastHourCluster) && !isLastRunHardAreaConstraintSatisfied(normalizedItemArea, existingHourAreas, homeArea)) {
            return null;
          }
          if (ENFORCE_AREA_PRIORITY_STRICT) {
            const strictItem = getStrictHomeCompatibilityScore(normalizedItemArea, homeArea);
            const dirItem = getDirectionAffinityScore(normalizedItemArea, homeArea);
            const areaMatchItem = getVehicleAreaMatchScore(vehicle, normalizedItemArea);
            const areaPriorityWeak = strictItem < 50 && dirItem < 24 && areaMatchItem < 55;
            const mixedDirectionRisk = existingHourAreas.length > 0 && routeContinuityPenalty >= 240;
            if (areaPriorityWeak || mixedDirectionRisk) return null;
          }
          const directionSplitGuardScore = getDirectionSplitGuardScore(
            normalizedItemArea,
            existingHourAreas,
            Number(item.distance_km || 0),
            isLastRun,
            isDefaultLastHourCluster
          );
          const hardDirectionSplitBlocked = isHardDirectionSplitBlocked(
            normalizedItemArea,
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
          score -= routeFlowScore * (isLastRun || isDefaultLastHourCluster ? 2.0 : 1.4);
          score += routeContinuityPenalty * (isLastRun || isDefaultLastHourCluster ? 2.2 : 1.4);
          score += areaMixGuardScore;
          score += clusterLockPenalty;

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
          if (hardReverseMixBlocked) {
            score += isLastRun || isDefaultLastHourCluster ? 880 : 620;
          }

          // 通常便は往復55分以上なら同方向束ね、近距離だけ分散を優先
          if (!(isLastRun || isDefaultLastHourCluster)) {
            const idleVehicleCount = getIdleVehicleCountForHour(cluster.hour);
            const pseudoCluster = { ...cluster, count: 1, totalDistance: Number(item.distance_km || 0), items: [item] };
            const roundTripMinutes = getClusterRoundTripMinutes(pseudoCluster);
            const bundleCompatibility = getBundleCompatibilityScore(normalizedItemArea, existingHourAreas);
            const forceBundle = shouldForceBundleCluster(pseudoCluster, existingHourAreas);
            const rotationPrediction = getRotationPredictionScore(
              cluster.hour,
              Number(item.distance_km || 0),
              1,
              sameHourLoad,
              routeFlowScore,
              routeContinuityPenalty,
              idleVehicleCount,
              normalizedItemArea
            );
            const preferSpread =
              sameHourLoad > 0 &&
              (!forceBundle) &&
              (!(routeFlowScore >= 85 && routeContinuityPenalty <= 45) || !rotationPrediction.canShare);

            if (sameHourLoad === 0) {
              if (roundTripMinutes < LONG_BUNDLE_ROUND_TRIP_MINUTES) score -= idleVehicleCount > 1 ? 135 : 84;
              else score += idleVehicleCount > 1 ? 35 : 12;
            } else if (preferSpread && idleVehicleCount > 0) {
              score += 130 + sameHourLoad * 30;
            } else if (rotationPrediction.canShare) {
              score -= Math.min(routeFlowScore, 120) * 0.26;
              score -= 10;
            }
            if (sameHourLoad > 0 && routeFlowScore >= 78 && routeContinuityPenalty <= 95) {
              score -= 95;
            }

            score -= bundleCompatibility * (roundTripMinutes >= LONG_BUNDLE_ROUND_TRIP_MINUTES ? 1.35 : 0.55);
            if (forceBundle && sameHourLoad > 0) score -= 150;
            if (roundTripMinutes < LONG_BUNDLE_ROUND_TRIP_MINUTES && sameHourLoad > 0 && bundleCompatibility < 50 && idleVehicleCount > 0) score += 60;

            score += getNormalRunReturnPenalty(
              cluster.hour,
              Number(item.distance_km || 0),
              1,
              sameHourLoad,
              routeFlowScore,
              routeContinuityPenalty,
              idleVehicleCount,
              normalizedItemArea
            );
          }

          const homePriorityWeight = getLastTripHomePriorityWeight(
            normalizedItemArea,
            homeArea,
            isLastRun,
            isDefaultLastHourCluster
          );
          score -= homePriorityWeight * 0.42;

          if ((isLastRun || isDefaultLastHourCluster) && homePriorityWeight <= 0) {
            score += isLastRun ? 36 : 10;
          }

          if (routeContinuityPenalty >= 170) {
            score += isLastRun || isDefaultLastHourCluster ? 160 : 80;
          } else if (routeContinuityPenalty >= 95) {
            score += isLastRun || isDefaultLastHourCluster ? 72 : 30;
          }

          if (!(isLastRun || isDefaultLastHourCluster) && hasAreaVehicleShortageForHour(cluster.hour)) {
            const returnMetrics = getVehicleCandidateReturnMetrics(vehicle, cluster.hour, [item]);
            const proximityScore = getVehicleItemProximityScore(vehicle, cluster.hour, item);
            score += returnMetrics.predictedReturnMinutes * 0.85;
            score += returnMetrics.extraSharedDelayMinutes * 2.4;
            score += returnMetrics.routeDistanceKm * 0.16;
            score -= proximityScore;
            if (sameHourLoad === 0) {
              score -= 32;
            }
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
      if (shouldClusterLock({ ...cluster, count: 1, totalDistance: Number(item.distance_km || 0), items: [item] })) {
        lockVehicleClusterArea(bestVehicle.id, cluster.area);
      }
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


function __getDisplayGroupAreaLabel(item) {
  const area = normalizeAreaLabel(
    item?.destination_area ||
    item?.cluster_area ||
    item?.planned_area ||
    item?.casts?.area ||
    "無し"
  );
  return getAreaDisplayGroup(area);
}

function __hasEnoughVehiclesForDisplayGroups(items, vehicles) {
  if (!ENABLE_DISPLAY_GROUP_FORCE_BRANCH) return false;
  const safeItems = Array.isArray(items) ? items.filter(Boolean) : [];
  const safeVehicles = Array.isArray(vehicles) ? vehicles.filter(Boolean) : [];
  if (!safeItems.length || !safeVehicles.length) return false;

  const hourMap = new Map();
  safeItems.forEach(item => {
    const hour = Number(item?.actual_hour ?? item?.plan_hour ?? 0);
    const group = __getDisplayGroupAreaLabel(item);
    if (!hourMap.has(hour)) hourMap.set(hour, new Set());
    hourMap.get(hour).add(group);
  });

  for (const groups of hourMap.values()) {
    if (groups.size > safeVehicles.length) return false;
  }
  return true;
}

function __getGroupFirstVehicleScore(vehicle, rows, monthlyMap) {
  const primary = rows[0] || {};
  const area = normalizeAreaLabel(
    primary?.destination_area ||
    primary?.cluster_area ||
    primary?.planned_area ||
    primary?.casts?.area ||
    "無し"
  );
  const month = monthlyMap?.get(Number(vehicle?.id)) || { totalDistance: 0, avgDistance: 0, workedDays: 0 };
  let score = 0;
  score += getVehicleAreaMatchScore(vehicle, area) * 2.6;
  score += getStrictHomeCompatibilityScore(area, vehicle?.home_area || "") * 1.9;
  score += Math.max(0, getDirectionAffinityScore(area, vehicle?.home_area || "")) * 0.8;
  score += getAreaAffinityScore(area, vehicle?.home_area || "") * 0.7;
  score -= Number(month.totalDistance || 0) * 0.02;
  score -= Number(month.avgDistance || 0) * 0.11;
  return score;
}

function __buildAssignmentsPreserveDisplayGroups(items, vehicles, monthlyMap) {
  const safeItems = Array.isArray(items) ? items.filter(Boolean) : [];
  const safeVehicles = Array.isArray(vehicles) ? vehicles.filter(Boolean) : [];
  if (!safeItems.length || !safeVehicles.length) return [];

  const assignments = [];
  const byHour = new Map();

  safeItems.forEach(item => {
    const hour = Number(item?.actual_hour ?? item?.plan_hour ?? 0);
    if (!byHour.has(hour)) byHour.set(hour, []);
    byHour.get(hour).push(item);
  });

  const hours = [...byHour.keys()].sort((a, b) => a - b);

  for (const hour of hours) {
    const hourItems = byHour.get(hour) || [];
    const byGroup = new Map();

    hourItems.forEach(item => {
      const group = __getDisplayGroupAreaLabel(item);
      if (!byGroup.has(group)) byGroup.set(group, []);
      byGroup.get(group).push(item);
    });

    const groupEntries = [...byGroup.entries()].sort((a, b) => {
      if (b[1].length !== a[1].length) return b[1].length - a[1].length;
      return String(a[0] || "").localeCompare(String(b[0] || ""), "ja");
    });

    const vehicleState = safeVehicles.map(vehicle => ({
      vehicle,
      id: Number(vehicle.id),
      capacity: Math.max(1, Number(vehicle.seat_capacity || 4)),
      count: 0,
      group: "",
      used: false
    }));

    for (const [group, rows] of groupEntries) {
      let remaining = [...sortItemsByNearestRoute(rows)];

      while (remaining.length) {
        const candidates = vehicleState
          .filter(state => state.count < state.capacity)
          .filter(state => !state.group || state.group === group)
          .sort((a, b) => {
            const aGroupBoost = a.group === group ? 5000 : (a.used ? 0 : 1800);
            const bGroupBoost = b.group === group ? 5000 : (b.used ? 0 : 1800);
            const aScore = __getGroupFirstVehicleScore(a.vehicle, remaining, monthlyMap) + aGroupBoost - a.count * 30;
            const bScore = __getGroupFirstVehicleScore(b.vehicle, remaining, monthlyMap) + bGroupBoost - b.count * 30;
            return bScore - aScore;
          });

        const picked = candidates[0];
        if (!picked) return [];

        picked.group = group;
        picked.used = true;

        const freeSeats = Math.max(0, picked.capacity - picked.count);
        const chunk = remaining.splice(0, freeSeats);
        let stopOrder = picked.count + 1;

        for (const row of chunk) {
          assignments.push({
            item_id: Number(row?.id || 0),
            actual_hour: hour,
            vehicle_id: picked.id,
            driver_name: picked.vehicle?.driver_name || "",
            distance_km: Number(row?.distance_km || 0),
            stop_order: stopOrder
          });
          stopOrder += 1;
        }

        picked.count += chunk.length;
      }
    }
  }

  const perVehicleHour = new Map();
  assignments.forEach(a => {
    const key = `${Number(a.vehicle_id)}__${Number(a.actual_hour)}`;
    if (!perVehicleHour.has(key)) perVehicleHour.set(key, []);
    perVehicleHour.get(key).push(a);
  });
  for (const rows of perVehicleHour.values()) {
    rows.sort((a, b) => Number(a.stop_order || 0) - Number(b.stop_order || 0) || Number(a.item_id || 0) - Number(b.item_id || 0));
    rows.forEach((a, idx) => {
      a.stop_order = idx + 1;
    });
  }

  return assignments;
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

  if (__hasEnoughVehiclesForDisplayGroups(unassignedItems, selectedVehicles)) {
    assignments = __buildAssignmentsPreserveDisplayGroups(unassignedItems, selectedVehicles, monthlyMap);
  } else {
    assignments = optimizeAssignmentsByRouteFlow(assignments, unassignedItems, selectedVehicles);

    if (!assignments.length) {
      assignments = buildFallbackAssignments(unassignedItems, selectedVehicles);
    }

    assignments = optimizeAssignmentsByDistanceBalance(assignments, unassignedItems, selectedVehicles, monthlyMap);
    assignments = applyLastTripDistanceCorrectionToAssignments(assignments, unassignedItems, selectedVehicles, monthlyMap);
    assignments = applyManualLastVehicleToAssignments(assignments, selectedVehicles);
  }

  if (!assignments.length) return;

  await applyAutoDispatchAssignments(assignments);
}

async function runAutoDispatch() {
  const selectedVehicles = Array.isArray(getSelectedVehiclesForToday())
    ? getSelectedVehiclesForToday().filter(Boolean)
    : [];
  if (!selectedVehicles.length) {
    alert("可能車両を選択してください");
    return;
  }

  const activeItems = Array.isArray(getActiveDispatchItemsForAutoAssign())
    ? getActiveDispatchItemsForAutoAssign().filter(Boolean)
    : [];
  if (!activeItems.length) {
    alert("自動配車対象のActualがありません");
    return;
  }

  lastAutoDispatchRunAtMinutes = getCurrentClockMinutes();

  const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
  let assignments = optimizeAssignments(activeItems, selectedVehicles, monthlyMap);
  if (!Array.isArray(assignments)) assignments = [];

  if (__hasEnoughVehiclesForDisplayGroups(activeItems, selectedVehicles)) {
    assignments = __buildAssignmentsPreserveDisplayGroups(activeItems, selectedVehicles, monthlyMap);
  } else {
    const routeFlowAssignments = optimizeAssignmentsByRouteFlow(assignments, activeItems, selectedVehicles);
    assignments = Array.isArray(routeFlowAssignments) ? routeFlowAssignments : assignments;

    if (!Array.isArray(assignments) || !assignments.length) {
      assignments = buildFallbackAssignments(activeItems, selectedVehicles);
    }
    if (!Array.isArray(assignments)) assignments = [];

    const balancedAssignments = optimizeAssignmentsByDistanceBalance(assignments, activeItems, selectedVehicles, monthlyMap);
    assignments = Array.isArray(balancedAssignments) ? balancedAssignments : assignments;

    const correctedAssignments = applyLastTripDistanceCorrectionToAssignments(assignments, activeItems, selectedVehicles, monthlyMap);
    assignments = Array.isArray(correctedAssignments) ? correctedAssignments : assignments;

    const manualAssignments = applyManualLastVehicleToAssignments(assignments, selectedVehicles);
    assignments = Array.isArray(manualAssignments) ? manualAssignments : assignments;
  }

  if (!Array.isArray(assignments) || !assignments.length) {
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
  scrollToDispatchResult();
}



function scrollToDispatchResult() {
  try {
    const target = els.dailyDispatchResult || document.getElementById("dailyDispatchResult");
    if (!target) return;
    window.requestAnimationFrame(() => {
      target.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  } catch (error) {
    console.warn("scrollToDispatchResult error:", error);
  }
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
      getStoredTravelMinutes(baseRows[baseRows.length - 1]?.casts?.travel_minutes) ||
      estimateFallbackTravelMinutes(reportedDistance, getRepresentativeAreaFromRows(baseRows)) + jobCount
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
  const driveMinutes = Math.round(getRowsTravelTimeSummary(baseRows).totalMinutes);

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
  const vehicleId = Number(vehicle?.id || 0);
  const summary = getVehiclePersistentDailyStats(vehicleId, orderedRows);
  const isLastTripDriver = isDriverLastTripChecked(vehicleId);
  const displayTotalKm = (!summary.hasFixedReport && isLastTripDriver)
    ? Number(summary.sendKm || 0)
    : Number(summary.totalKm || 0);
  const sendOnlyMinutes = Math.round(getRowsTravelTimeSummary(orderedRows).sendOnlyMinutes);
  const displayDriveMinutes = (!summary.hasFixedReport && isLastTripDriver)
    ? sendOnlyMinutes
    : Math.round(Number(summary.driveMinutes || 0));

  return {
    sendKm: Number(summary.sendKm || 0),
    returnKm: Number(summary.returnKm || 0),
    totalKm: Number(displayTotalKm || 0),
    driveMinutes: Number(displayDriveMinutes || 0),
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
}function buildVehicleLineLabel(vehicle) {
  return "";
}


let dispatchOverviewMap = null;
let dispatchOverviewLayer = null;
let dispatchOverviewHost = null;
let lastDispatchOverviewCards = [];
const dispatchOverviewFilterState = new Map();

function getDispatchOverviewVehicleColor(index) {
  const colors = ["#3b82f6", "#22c55e", "#ef4444", "#f59e0b", "#a855f7", "#06b6d4", "#84cc16", "#ec4899"];
  return colors[Math.abs(Number(index) || 0) % colors.length];
}

function getVehicleDisplayName(vehicle) {
  return String(vehicle?.driver_name || vehicle?.plate_number || `車両${vehicle?.id || ""}` || "車両").trim() || "車両";
}

function getDispatchRowLatLng(row) {
  const lat = Number(row?.casts?.latitude ?? row?.latitude ?? row?.lat);
  const lng = Number(row?.casts?.longitude ?? row?.longitude ?? row?.lng);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
  return { lat, lng };
}

function buildVehicleRouteMapUrl(vehicle, orderedRows) {
  const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
  if (!rows.length) return "";
  const origin = `${ORIGIN_LAT},${ORIGIN_LNG}`;
  const points = rows
    .map(row => {
      const p = getDispatchRowLatLng(row);
      if (p) return `${p.lat},${p.lng}`;
      const address = String(row?.destination_address || row?.casts?.address || "").trim();
      return address || "";
    })
    .filter(Boolean);
  if (!points.length) return "";
  const destination = points[points.length - 1];
  const waypoints = points.slice(0, -1);
  const params = new URLSearchParams({
    api: "1",
    origin,
    destination,
    travelmode: "driving"
  });
  if (waypoints.length) params.set("waypoints", waypoints.join("|"));
  return `https://www.google.com/maps/dir/?${params.toString()}`;
}

function renderDispatchOverviewLegend(cards) {
  const host = document.getElementById("dispatchOverviewLegend");
  if (!host) return;
  host.innerHTML = "";
  (cards || []).forEach(({ vehicle, orderedRows }, index) => {
    if (!orderedRows?.length) return;
    const vehicleId = Number(vehicle?.id || 0);
    if (!dispatchOverviewFilterState.has(vehicleId)) dispatchOverviewFilterState.set(vehicleId, true);
    const isOn = dispatchOverviewFilterState.get(vehicleId) !== false;
    const color = getDispatchOverviewVehicleColor(index);
    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn ghost";
    button.style.cssText = `display:inline-flex;align-items:center;gap:8px;padding:6px 10px;border-radius:999px;border:1px solid ${isOn ? color : 'rgba(148,163,184,.35)'};background:${isOn ? 'rgba(15,23,42,.85)' : 'rgba(15,23,42,.45)'};color:${isOn ? '#fff' : '#94a3b8'};cursor:pointer;`;
    button.innerHTML = `<span style="display:inline-block;width:12px;height:12px;border-radius:999px;background:${color};box-shadow:0 0 0 2px rgba(255,255,255,.18) inset;"></span><span>${escapeHtml(getVehicleDisplayName(vehicle))}</span><span style="font-size:11px;opacity:.8;">${isOn ? '表示中' : '非表示'}</span>`;
    button.addEventListener("click", () => {
      dispatchOverviewFilterState.set(vehicleId, !isOn);
      renderDispatchOverviewLegend(lastDispatchOverviewCards);
      renderDispatchOverviewMap(lastDispatchOverviewCards);
    });
    host.appendChild(button);
  });
}

function renderDispatchOverviewMap(cards) {
  const host = document.getElementById("dispatchOverviewMap");
  lastDispatchOverviewCards = Array.isArray(cards) ? cards : [];
  renderDispatchOverviewLegend(lastDispatchOverviewCards);
  if (!host || !window.L) return;

  if (dispatchOverviewMap && dispatchOverviewHost !== host) {
    try { dispatchOverviewMap.remove(); } catch (_) {}
    dispatchOverviewMap = null;
    dispatchOverviewLayer = null;
  }

  dispatchOverviewHost = host;
  host.innerHTML = "";

  if (!dispatchOverviewMap) {
    dispatchOverviewMap = window.L.map(host, { preferCanvas: true, zoomControl: true });
    window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "&copy; OpenStreetMap contributors"
    }).addTo(dispatchOverviewMap);
  }

  if (dispatchOverviewLayer) {
    try { dispatchOverviewLayer.remove(); } catch (_) {}
  }
  dispatchOverviewLayer = window.L.layerGroup().addTo(dispatchOverviewMap);

  const bounds = [];
  const originMarker = window.L.circleMarker([ORIGIN_LAT, ORIGIN_LNG], {
    radius: 9,
    color: "#111827",
    fillColor: "#111827",
    fillOpacity: 1,
    weight: 2
  }).bindPopup(`<b>${escapeHtml(ORIGIN_LABEL || '起点')}</b>`);
  originMarker.addTo(dispatchOverviewLayer);
  bounds.push([ORIGIN_LAT, ORIGIN_LNG]);

  (lastDispatchOverviewCards || []).forEach(({ vehicle, orderedRows }, index) => {
    if (!orderedRows?.length) return;
    const vehicleId = Number(vehicle?.id || 0);
    if (!dispatchOverviewFilterState.has(vehicleId)) dispatchOverviewFilterState.set(vehicleId, true);
    if (dispatchOverviewFilterState.get(vehicleId) === false) return;
    const color = getDispatchOverviewVehicleColor(index);
    orderedRows.forEach((row, rowIndex) => {
      const point = getDispatchRowLatLng(row);
      if (!point) return;
      const pinNo = rowIndex + 1;
      const marker = window.L.marker([point.lat, point.lng], {
        icon: window.L.divIcon({
          className: 'dispatch-overview-pin',
          html: `<div style="width:26px;height:26px;border-radius:999px;background:${color};border:2px solid rgba(255,255,255,.96);box-shadow:0 6px 16px rgba(15,23,42,.38);display:flex;align-items:center;justify-content:center;color:#fff;font-weight:800;font-size:12px;line-height:1;">${pinNo}</div>`,
          iconSize: [26, 26],
          iconAnchor: [13, 13],
          popupAnchor: [0, -14]
        })
      });
      const vehicleName = getVehicleDisplayName(vehicle);
      const castName = String(row?.casts?.name || "-");
      const area = normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "-");
      const distanceKm = Number(row?.distance_km || 0).toFixed(1);
      marker.bindPopup(`<div style="min-width:180px;line-height:1.7;"><b>${escapeHtml(vehicleName)}-${pinNo}</b><br>キャスト: ${escapeHtml(castName)}<br>方面: ${escapeHtml(area)}<br>距離: ${escapeHtml(distanceKm)}km</div>`);
      marker.bindTooltip(`${escapeHtml(vehicleName)}-${pinNo}`, { direction: "top", opacity: 0.95 });
      marker.addTo(dispatchOverviewLayer);
      bounds.push([point.lat, point.lng]);
    });
  });

  window.setTimeout(() => {
    try {
      dispatchOverviewMap.invalidateSize();
      if (bounds.length >= 2) {
        dispatchOverviewMap.fitBounds(bounds, { padding: [24, 24] });
      } else if (bounds.length === 1) {
        dispatchOverviewMap.setView(bounds[0], 13);
      } else {
        dispatchOverviewMap.setView([ORIGIN_LAT, ORIGIN_LNG], 11);
      }
    } catch (error) {
      console.error("dispatchOverviewMap fit error:", error);
    }
  }, 50);
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

function buildLineVehicleBlock(vehicle, orderedRows) {
  const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
  if (!rows.length) return "";

  const summary = getVehicleDailySummary(vehicle, rows);
  const forecast = getVehicleRotationForecastSafe(vehicle, rows);
  const driverName = getVehicleDisplayName(vehicle);
  const lineId = String(vehicle?.line_id || "").trim();
  const lastTripTag = isDriverLastTripChecked(vehicle?.id) ? " 【ラスト便】" : "";
  const routeUrl = buildVehicleRouteMapUrl(vehicle, rows);

  const header = [lineId, `🚗 ${driverName}${lastTripTag}`].filter(Boolean).join(" ");
  const body = rows.map((row, index) => {
    const castName = String(row?.casts?.name || "-").trim() || "-";
    const areaLabel = normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "無し");
    return `${index + 1}️⃣ ${castName}　${areaLabel}`;
  });

  const footer = [
    `戻り ${forecast?.returnAfterLabel || "-"} / 次便可能 ${forecast?.predictedReadyTime || "-"}`,
    `距離 ${Number(summary?.totalKm || 0).toFixed(1)}km / 時間 ${formatMinutesAsJa(summary?.driveMinutes || 0)}`
  ];

  if (routeUrl) {
    footer.push("📍ルート");
    footer.push(routeUrl);
  }

  return [header, ...body, ...footer].join("\n");
}

function buildLineResultText() {
  const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
  const vehicles = getSelectedVehiclesForToday();
  const activeItems = currentActualsCache.filter(
    x => normalizeStatus(x.status) !== "done" && normalizeStatus(x.status) !== "cancel"
  );

  const cards = buildDailyDispatchVehicleCards(vehicles, activeItems, monthlyMap);
  return cards
    .map(({ vehicle, orderedRows }) => buildLineVehicleBlock(vehicle, orderedRows))
    .filter(Boolean)
    .join("\n\n")
    .trim();
}



function getOvernightLooseHourBucket(item) {
  const hour = Number(item?.actual_hour ?? item?.plan_hour ?? 0);
  return hour >= 0 && hour <= 5 ? "overnight" : String(hour);
}

function getAssignmentItemHourGroupKey(item) {
  const area = normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || "無し");
  const canonical = getCanonicalArea(area) || area;
  const group = getAreaDisplayGroup(area) || canonical || area;
  const hourBucket = getOvernightLooseHourBucket(item);
  return `${hourBucket}__${group}__${canonical}`;
}

function getSoftBridgeAreaScore(baseArea, compareArea) {
  const base = getCanonicalArea(normalizeAreaLabel(baseArea)) || normalizeAreaLabel(baseArea);
  const compare = getCanonicalArea(normalizeAreaLabel(compareArea)) || normalizeAreaLabel(compareArea);
  if (!base || !compare) return 0;
  if (base === compare) return 140;

  const northeast = new Set(["我孫子方面", "取手方面", "藤代方面", "守谷方面", "牛久方面"]);
  const eastUrban = new Set(["葛飾方面", "足立方面", "墨田方面", "荒川方面", "江戸川方面", "市川方面"]);
  const downtownBridge = new Set(["墨田方面", "荒川方面"]);

  if ((base === "我孫子方面" && eastUrban.has(compare)) || (compare === "我孫子方面" && eastUrban.has(base))) return 108;
  if ((northeast.has(base) && eastUrban.has(compare)) || (northeast.has(compare) && eastUrban.has(base))) return 92;
  if ((base === "我孫子方面" && northeast.has(compare)) || (compare === "我孫子方面" && northeast.has(base))) return 132;
  if ((downtownBridge.has(base) && eastUrban.has(compare)) || (downtownBridge.has(compare) && eastUrban.has(base))) return 86;

  const affinity = getAreaAffinityScore(base, compare);
  const direction = getDirectionAffinityScore(base, compare);
  if (affinity >= 72 && direction >= 0) return 74;
  if (affinity >= 58 && direction >= 18) return 54;
  return 0;
}

function getRoundTripMinutesForItem(item) {
  const area = normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || "無し");
  const travel = getStoredTravelMinutes(item?.casts?.travel_minutes || item?.travel_minutes);
  const oneWay = travel > 0 ? travel : estimateTravelMinutesByDistance(Number(item?.distance_km || 0), area);
  return Math.round(oneWay * 2);
}

function countAssignmentsByVehicle(assignments) {
  const map = new Map();
  (assignments || []).forEach(a => {
    map.set(Number(a.vehicle_id), Number(map.get(Number(a.vehicle_id)) || 0) + 1);
  });
  return map;
}

function rebundleLongDistanceDirectionalClusters(assignments, items, vehicles, monthlyMap, options = {}) {
  const working = Array.isArray(assignments) ? assignments.map(a => ({ ...a })) : [];
  if (!working.length || !Array.isArray(items) || !items.length || !Array.isArray(vehicles) || !vehicles.length) return working;

  const itemMap = new Map(items.map(item => [Number(item.id), item]));
  const vehicleMap = new Map(vehicles.map(v => [Number(v.id), v]));
  const threshold = Number(options.roundTripThreshold || 55);

  const rebuildVehicleLoadMap = () => countAssignmentsByVehicle(working);

  const byKey = new Map();
  working.forEach(a => {
    const item = itemMap.get(Number(a.item_id));
    if (!item) return;
    const key = getAssignmentItemHourGroupKey(item);
    if (!byKey.has(key)) byKey.set(key, []);
    byKey.get(key).push(a);
  });

  for (const [, clusterAssignments] of byKey.entries()) {
    if (!clusterAssignments || clusterAssignments.length < 2) continue;

    const itemsForCluster = clusterAssignments.map(a => itemMap.get(Number(a.item_id))).filter(Boolean);
    if (itemsForCluster.length < 2) continue;

    const minRoundTrip = Math.min(...itemsForCluster.map(getRoundTripMinutesForItem));
    if (minRoundTrip < threshold) continue;

    const canonical = getCanonicalArea(normalizeAreaLabel(itemsForCluster[0]?.destination_area || itemsForCluster[0]?.cluster_area || itemsForCluster[0]?.casts?.area || "無し"));
    const group = getAreaDisplayGroup(itemsForCluster[0]?.destination_area || itemsForCluster[0]?.cluster_area || itemsForCluster[0]?.casts?.area || "無し");

    let bestVehicleId = null;
    let bestScore = -Infinity;
    const vehicleLoadMap = rebuildVehicleLoadMap();

    for (const vehicle of vehicles) {
      const vehicleId = Number(vehicle.id);
      const homeArea = normalizeAreaLabel(vehicle?.home_area || "");
      const currentVehicleAssignments = working.filter(a => Number(a.vehicle_id) === vehicleId);
      const currentAreas = currentVehicleAssignments.map(a => {
        const item = itemMap.get(Number(a.item_id));
        return normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.casts?.area || "無し");
      }).filter(Boolean);

      if (currentAreas.length && currentAreas.some(area => hasHardReverseMix(area, [canonical || group]))) continue;
      if (isHardReverseForHome(canonical || group, homeArea)) continue;

      const month = monthlyMap?.get(vehicleId) || { totalDistance: 0, avgDistance: 0, workedDays: 0 };
      const existingCount = Number(vehicleLoadMap.get(vehicleId) || 0);
      const strict = getStrictHomeCompatibilityScore(canonical || group, homeArea);
      const direction = Math.max(0, getDirectionAffinityScore(canonical || group, homeArea));
      const affinity = currentAreas.length
        ? Math.max(...currentAreas.map(area => getAreaAffinityScore(canonical || group, area)))
        : getAreaAffinityScore(canonical || group, homeArea);
      const sameGroupBonus = currentAreas.length
        ? Math.max(...currentAreas.map(area => {
            const areaGroup = getAreaDisplayGroup(area);
            const areaCanonical = getCanonicalArea(area);
            return (areaCanonical && canonical && areaCanonical === canonical) ? 160 : (areaGroup === group ? 110 : 0);
          }))
        : 40;
      const currentCountInCluster = clusterAssignments.filter(a => Number(a.vehicle_id) === vehicleId).length;
      const totalClusterDistance = itemsForCluster.reduce((sum, item) => sum + Number(item?.distance_km || 0), 0);
      let score = 0;
      score += currentCountInCluster * 260;
      score += sameGroupBonus;
      score += strict * 1.3 + direction * 0.9 + affinity * 0.6;
      score -= Number(month.totalDistance || 0) * 0.02;
      score -= Number(month.avgDistance || 0) * 0.15;
      score -= existingCount * 24;
      score -= totalClusterDistance * 0.05;

      if (bestVehicleId == null || score > bestScore) {
        bestVehicleId = vehicleId;
        bestScore = score;
      }
    }

    if (bestVehicleId == null) continue;
    clusterAssignments.forEach(a => {
      const vehicle = vehicleMap.get(bestVehicleId);
      a.vehicle_id = bestVehicleId;
      a.driver_name = vehicle?.driver_name || "";
    });
  }

  const evaluateBridgeVehicle = (assignment, vehicle) => {
    const item = itemMap.get(Number(assignment.item_id));
    if (!item || !vehicle) return -Infinity;
    const targetArea = normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || "無し");
    const vehicleId = Number(vehicle.id);
    const existingAssignments = working.filter(a => Number(a.vehicle_id) === vehicleId && Number(a.item_id) !== Number(assignment.item_id));
    const existingAreas = existingAssignments
      .map(a => {
        const row = itemMap.get(Number(a.item_id));
        return normalizeAreaLabel(row?.destination_area || row?.cluster_area || row?.planned_area || row?.casts?.area || "無し");
      })
      .filter(Boolean);

    if (existingAreas.length && existingAreas.some(area => hasHardReverseMix(targetArea, [area]))) return -Infinity;
    if (isHardReverseForHome(targetArea, vehicle?.home_area || "")) return -Infinity;

    const month = monthlyMap?.get(vehicleId) || { totalDistance: 0, avgDistance: 0, workedDays: 0 };
    const existingCount = existingAssignments.length;
    const strict = getStrictHomeCompatibilityScore(targetArea, vehicle?.home_area || "");
    const direction = Math.max(0, getDirectionAffinityScore(targetArea, vehicle?.home_area || ""));
    const bestBridge = existingAreas.length
      ? Math.max(...existingAreas.map(area => getSoftBridgeAreaScore(targetArea, area)))
      : 0;
    const bestAffinity = existingAreas.length
      ? Math.max(...existingAreas.map(area => getAreaAffinityScore(targetArea, area)))
      : getAreaAffinityScore(targetArea, vehicle?.home_area || "");
    const bundleCount = existingAreas.filter(area => getSoftBridgeAreaScore(targetArea, area) >= 90).length;
    const isLooseOvernight = getOvernightLooseHourBucket(item) === "overnight";
    let score = 0;
    score += bestBridge * 3.6;
    score += bundleCount * 110;
    score += strict * 1.15 + direction * 0.7 + bestAffinity * 0.4;
    score -= Number(month.totalDistance || 0) * 0.018;
    score -= Number(month.avgDistance || 0) * 0.11;
    score -= existingCount * 14;
    if (isLooseOvernight) score += 44;
    if (!existingAreas.length) score -= 120;
    return score;
  };

  for (let pass = 0; pass < 2; pass += 1) {
    for (const assignment of [...working]) {
      const item = itemMap.get(Number(assignment.item_id));
      if (!item) continue;
      if (getOvernightLooseHourBucket(item) !== "overnight") continue;

      const targetArea = normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || "無し");
      const currentVehicle = vehicleMap.get(Number(assignment.vehicle_id));
      const currentScore = evaluateBridgeVehicle(assignment, currentVehicle);
      let bestVehicle = currentVehicle;
      let bestScore = currentScore;

      for (const vehicle of vehicles) {
        const score = evaluateBridgeVehicle(assignment, vehicle);
        if (score > bestScore) {
          bestScore = score;
          bestVehicle = vehicle;
        }
      }

      if (!bestVehicle || Number(bestVehicle.id) === Number(assignment.vehicle_id)) continue;
      if (bestScore < currentScore + 42) continue;

      const destinationAssignments = working.filter(a => Number(a.vehicle_id) === Number(bestVehicle.id) && Number(a.item_id) !== Number(assignment.item_id));
      const destinationAreas = destinationAssignments.map(a => {
        const row = itemMap.get(Number(a.item_id));
        return normalizeAreaLabel(row?.destination_area || row?.cluster_area || row?.planned_area || row?.casts?.area || "無し");
      }).filter(Boolean);
      if (destinationAreas.length && destinationAreas.every(area => getSoftBridgeAreaScore(targetArea, area) < 70)) continue;

      assignment.vehicle_id = Number(bestVehicle.id);
      assignment.driver_name = bestVehicle?.driver_name || "";
    }
  }

  return working;
}

function optimizeAssignmentsByDistanceBalance(assignments, items, vehicles, monthlyMap) {
  const working = assignments.map(a => ({ ...a }));
  if (!working.length || !vehicles.length) return working;
  if (!ENABLE_DISTANCE_REBALANCE) return working;

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
      if (ENFORCE_AREA_PRIORITY_STRICT && compat < currentCompat + 10) continue;

      const score = projectedGapImprove * 1.4 + (compat - currentCompat) * 1.6 - dist * 0.12;
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
  const defaultLastHour = getEffectiveLastHour(dateStr, working.map(a => ({ actual_hour: a?.actual_hour })));
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

  const evaluate = (vehicle, item, existingAreas = []) => {
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
    if (!isLastRunHardAreaConstraintSatisfied(area, existingAreas, vehicle?.home_area || "")) return -Infinity;

    let score =
      strict * 3.2 +
      direction * 1.6 +
      affinity * 1.2 +
      vehicleMatch * 0.25 +
      itemDistance * (strict >= 78 ? 0.35 : strict >= 52 ? 0.15 : 0.05);

    score -= projectedMonthly * 0.04;
    if (hardReverse) score -= 9999;
    if (Number(vehicle?.id) === Number(manualVehicleId) && strict >= 52 && !hardReverse) score += 38;
    return score;
  };

  for (const target of targetRows) {
    const item = itemMap.get(Number(target.item_id));
    if (!item) continue;

    const currentVehicle = vehicles.find(v => Number(v.id) === Number(target.vehicle_id));
    const currentExistingAreas = working
      .filter(
        a =>
          Number(a.vehicle_id) === Number(target.vehicle_id) &&
          Number(a.actual_hour ?? 0) === Number(targetHour) &&
          Number(a.item_id) !== Number(target.item_id)
      )
      .map(a => normalizeAreaLabel(itemMap.get(Number(a.item_id))?.destination_area || itemMap.get(Number(a.item_id))?.cluster_area || ""))
      .filter(Boolean);
    const currentScore = evaluate(currentVehicle, item, currentExistingAreas);
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

      const existingAreas = working
        .filter(
          a =>
            Number(a.vehicle_id) === Number(vehicle.id) &&
            Number(a.actual_hour ?? 0) === Number(targetHour) &&
            Number(a.item_id) !== Number(target.item_id)
        )
        .map(a => normalizeAreaLabel(itemMap.get(Number(a.item_id))?.destination_area || itemMap.get(Number(a.item_id))?.cluster_area || ""))
        .filter(Boolean);
      const score = evaluate(vehicle, item, existingAreas);
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
    renderOperationAndSimulationUI();
    return;
  }

  const activeItems = currentActualsCache.filter(
    x => normalizeStatus(x.status) !== "done" && normalizeStatus(x.status) !== "cancel"
  );

  try {
    const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
    const timelineHtml = buildRotationTimelineHtmlSafe(vehicles, activeItems);
    const cards = buildDailyDispatchVehicleCards(vehicles, activeItems, monthlyMap);

    const overviewHtml = `
      <div class="vehicle-result-card" style="margin-bottom:16px;">
        <div class="vehicle-result-head" style="align-items:flex-start;gap:12px;">
          <div class="vehicle-result-title">
            <h4>全配車 俯瞰マップ</h4>
            <div class="vehicle-result-meta">各車両は色分け表示。凡例クリックで表示ON/OFFできます。</div>
          </div>
          <div class="vehicle-result-badges">
            <span class="metric-badge">起点 黒</span>
            <span class="metric-badge">車両別 色分け</span>
            <span class="metric-badge">ピン番号=降車順</span>
          </div>
        </div>
        <div id="dispatchOverviewLegend" style="display:flex;flex-wrap:wrap;gap:8px;margin:10px 0 14px;"></div>
        <div id="dispatchOverviewMap" style="width:100%;height:420px;border-radius:18px;overflow:hidden;background:#0f172a;"></div>
      </div>
    `;

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
                  ${isDriverLastTripChecked(vehicle.id) ? `<span class="badge-status assigned" style="margin-left:8px;">ラスト便チェック</span>` : ""}
                </h4>
                <div class="vehicle-result-meta">
                  ${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))}
                  / 帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))}
                  / 定員${vehicle.seat_capacity ?? "-"}
                  ${isDriverLastTripChecked(vehicle.id) ? `/ ラスト便対象` : ""}
                </div>
              </div>
              <div class="vehicle-result-badges">
                <span class="metric-badge">人数 ${rows.length}</span>
                <span class="metric-badge">累計距離 ${summary.totalKm.toFixed(1)}km</span>
                <span class="metric-badge">累計時間 ${escapeHtml(formatMinutesAsJa(summary.driveMinutes))}</span>
                <span class="metric-badge">累計件数 ${summary.jobCount}件</span>
                ${buildVehicleRouteMapUrl(vehicle, orderedRows) ? `<a href="${escapeHtml(buildVehicleRouteMapUrl(vehicle, orderedRows))}" target="_blank" rel="noopener noreferrer" class="metric-badge" style="text-decoration:none;">Google Mapsで開く</a>` : ""}
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

    els.dailyDispatchResult.innerHTML = timelineHtml + overviewHtml + cardsHtml;
    renderOperationAndSimulationUI();
    renderDispatchOverviewMap(cards);

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
    renderOperationAndSimulationUI();
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
    els.dailyMileageInputs.innerHTML = `<div class="muted">可能車両を選択すると入力欄が表示されます</div>`;
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
    alert("先に可能車両を選択してください");
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
  if (els.simulationSlotSelect) els.simulationSlotSelect.addEventListener("change", () => {
    if (suppressSimulationSlotChange || isRefreshingHybridUI) return;
    simulationSlotHour = Number(els.simulationSlotSelect.value || getOperationBaseHour());
    runSlotDiagnosisPreview();
  });
  if (els.simulationIncludePlanInflow) els.simulationIncludePlanInflow.addEventListener("change", () => {
    if (isRefreshingHybridUI) return;
    runSlotDiagnosisPreview();
  });
  if (els.runSimulationBtn) els.runSimulationBtn.addEventListener("click", runSlotDiagnosisPreview);
  if (els.runSimulationDispatchBtn) els.runSimulationDispatchBtn.addEventListener("click", runSimulationDispatchPreview);
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
  els.openManualBtn?.addEventListener("click", openManual);
  els.dangerResetBtn?.addEventListener("click", resetAllDataDanger);
  els.resetCastsBtn?.addEventListener("click", resetAllCastsDanger);
  els.resetVehiclesBtn?.addEventListener("click", resetAllVehiclesDanger);

  ensureCastTravelMinutesUi();
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
  // 時間取得ボタンは使用しない運用のため未バインド
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

    ensureCastTravelMinutesUi();
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
];const _THEMIS_V36_guessArea = guessArea;
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



/* ===== THEMIS v5.4 配車AI強化版 patch start ===== */

function getThemisV54RowPoint(row) {
  const lat = toNullableNumber(row?.casts?.latitude ?? row?.latitude);
  const lng = toNullableNumber(row?.casts?.longitude ?? row?.longitude);
  if (isValidLatLng(lat, lng)) return { lat, lng };
  return null;
}

function getThemisV54RowArea(row) {
  return normalizeAreaLabel(
    row?.destination_area ||
    row?.planned_area ||
    row?.cluster_area ||
    row?.casts?.area ||
    ""
  );
}

function getThemisV54RowDistance(row) {
  return Number(
    row?.distance_km ??
    row?.casts?.distance_km ??
    0
  ) || 0;
}

function getThemisV54StoredTravelMinutes(row) {
  const stored = getStoredTravelMinutes(row?.casts?.travel_minutes ?? row?.travel_minutes);
  return stored > 0 ? stored : 0;
}

function getThemisV54SegmentDistanceKm(fromPoint, toRow) {
  const toPoint = getThemisV54RowPoint(toRow);
  if (fromPoint && toPoint) {
    return Number(estimateRoadKmBetweenPoints(fromPoint.lat, fromPoint.lng, toPoint.lat, toPoint.lng) || 0);
  }
  return getThemisV54RowDistance(toRow);
}

function getThemisV54SegmentMinutes(fromPoint, toRow) {
  const segKm = Math.max(0, Number(getThemisV54SegmentDistanceKm(fromPoint, toRow) || 0));
  const area = getThemisV54RowArea(toRow);
  const storedMinutes = getThemisV54StoredTravelMinutes(toRow);
  const baseDistance = Math.max(0.1, Number(getThemisV54RowDistance(toRow) || segKm || 0.1));

  if (storedMinutes > 0) {
    const derivedSpeed = Math.max(16, Math.min(60, (baseDistance / storedMinutes) * 60));
    return Math.max(1, Math.round((segKm / derivedSpeed) * 60));
  }

  return Math.max(1, Math.round(estimateFallbackTravelMinutes(segKm, area)));
}

function getThemisV54VehicleIdFromRows(rows) {
  const first = Array.isArray(rows) ? rows.find(Boolean) : null;
  return Number(first?.vehicle_id || 0);
}

function getThemisV54VehicleIsLastTrip(rows) {
  const vehicleId = getThemisV54VehicleIdFromRows(rows);
  return vehicleId > 0 ? isDriverLastTripChecked(vehicleId) : false;
}

function getThemisV54TravelSummary(rows) {
  const ordered = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!ordered.length) {
    return {
      outboundMinutes: 0,
      returnMinutes: 0,
      totalMinutes: 0,
      sendOnlyMinutes: 0,
      stopCount: 0,
      segmentMinutes: [],
      model: "v5.4"
    };
  }

  const stopCount = ordered.length;
  const segmentMinutes = [];

  let currentPoint = { lat: ORIGIN_LAT, lng: ORIGIN_LNG };
  let outboundMinutes = 0;

  ordered.forEach((row, index) => {
    let legMinutes = 0;
    if (index === 0) {
      legMinutes = getThemisV54StoredTravelMinutes(row);
      if (!(legMinutes > 0)) {
        legMinutes = Math.max(1, Math.round(
          estimateFallbackTravelMinutes(getThemisV54RowDistance(row), getThemisV54RowArea(row))
        ));
      }
    } else {
      legMinutes = getThemisV54SegmentMinutes(currentPoint, row);
    }

    legMinutes = Math.max(1, Math.round(Number(legMinutes || 0)));
    segmentMinutes.push(legMinutes);
    outboundMinutes += legMinutes;

    const point = getThemisV54RowPoint(row);
    if (point) currentPoint = point;
  });

  const dwellMinutes = stopCount;
  const sendOnlyMinutes = Math.max(0, Math.round(outboundMinutes + dwellMinutes));
  const isLastTrip = getThemisV54VehicleIsLastTrip(ordered);
  const returnMinutes = isLastTrip ? 0 : Math.max(0, Math.round(outboundMinutes));
  const totalMinutes = Math.max(0, Math.round(sendOnlyMinutes + returnMinutes));

  return {
    outboundMinutes: Math.round(outboundMinutes),
    returnMinutes,
    totalMinutes,
    sendOnlyMinutes,
    stopCount,
    segmentMinutes,
    model: "v5.4"
  };
}

getRowsOutboundMinutes = function(rows) {
  return getThemisV54TravelSummary(rows).outboundMinutes;
};

getRowsReturnMinutes = function(rows) {
  return getThemisV54TravelSummary(rows).returnMinutes;
};

getRowsTravelTimeSummary = function(rows) {
  return getThemisV54TravelSummary(rows);
};

calcVehicleRotationForecastGlobal = function(vehicle, orderedRows) {
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
  const representativeArea = getRepresentativeAreaFromRows(rows);
  const summary = getThemisV54TravelSummary(rows);
  const primaryZone = getDistanceZoneInfoGlobal(Math.max(routeDistanceKm, returnDistanceKm), representativeArea);

  const departDelayMinutes = getExpectedDepartureDelayMinutes(baseHour);
  const predictedDepartureAbs = baseHour * 60 + departDelayMinutes;
  const predictedReturnAbs = predictedDepartureAbs + Number(summary.totalMinutes || 0);
  const predictedReadyAbs = predictedReturnAbs + 1;

  let extraSharedDelayMinutes = 0;
  if (rows.length >= 2) {
    const firstOnlySummary = getThemisV54TravelSummary([rows[0]]);
    extraSharedDelayMinutes = Math.max(0, Number(summary.totalMinutes || 0) - Number(firstOnlySummary.totalMinutes || 0));
  }

  return {
    routeDistanceKm: Number(routeDistanceKm.toFixed(1)),
    returnDistanceKm: Number(returnDistanceKm.toFixed(1)),
    zoneLabel: primaryZone.label,
    predictedDepartureTime: formatClockTimeFromMinutesGlobal(predictedDepartureAbs),
    predictedReturnTime: formatClockTimeFromMinutesGlobal(predictedReturnAbs),
    predictedReadyTime: formatClockTimeFromMinutesGlobal(predictedReadyAbs),
    predictedReturnMinutes: Math.round(Number(summary.totalMinutes || 0)),
    extraSharedDelayMinutes: Math.round(extraSharedDelayMinutes),
    stopCount: rows.length,
    returnAfterLabel: `${Math.round(Number(summary.totalMinutes || 0))}分後`
  };
};

sortItemsByNearestRoute = function(items) {
  const remaining = [...(items || [])].filter(Boolean);
  const sorted = [];
  let currentPoint = { lat: ORIGIN_LAT, lng: ORIGIN_LNG };
  let currentItem = null;

  while (remaining.length) {
    let bestIndex = 0;
    let bestScore = -Infinity;

    remaining.forEach((item, index) => {
      const point = getThemisV54RowPoint(item);
      const area = getThemisV54RowArea(item);
      const storedTravel = getThemisV54StoredTravelMinutes(item);
      const baseDistance = getThemisV54RowDistance(item);

      let score = 0;

      const legKm = point
        ? estimateRoadKmBetweenPoints(currentPoint.lat, currentPoint.lng, point.lat, point.lng)
        : Math.max(0, baseDistance);

      score -= legKm * 3.2;

      if (currentItem) {
        const prevArea = getThemisV54RowArea(currentItem);
        score += getRouteFlowCompatibilityBetweenAreas(prevArea, area) * 2.1;
        score -= getPairRouteContinuityPenalty(prevArea, area) * 1.35;
        if (isHardReverseMixForRoute(prevArea, area)) score -= 460;
        score += getDirectionAffinityScore(prevArea, area) * 0.42;
      } else {
        score += getRouteFlowSortWeight(area) * 3.4;
        score += Math.min(40, storedTravel) * 0.7;
      }

      // 1件目は travel_minutes の長いキャストを先頭にしすぎないよう軽く抑える
      if (!currentItem && storedTravel > 0) {
        score -= Math.max(0, storedTravel - 25) * 0.18;
      }

      // 近いのに大きく逆流する候補は避ける
      if (currentItem) {
        const prevArea = getThemisV54RowArea(currentItem);
        if (getDirectionAffinityScore(prevArea, area) <= -38) {
          score -= 120;
        }
      }

      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });

    const picked = remaining.splice(bestIndex, 1)[0];
    sorted.push(picked);
    currentItem = picked;

    const pickedPoint = getThemisV54RowPoint(picked);
    if (pickedPoint) currentPoint = pickedPoint;
  }

  return sorted;
};

function getThemisV54VehicleProjectedAvg(vehicleId, monthlyMap, extraAssignedDistanceMap, extraAssignedCountMap) {
  const stats = monthlyMap?.get(Number(vehicleId)) || { totalDistance: 0, workedDays: 0, avgDistance: 0 };
  const reportDate = els.dispatchDate?.value || els.actualDate?.value || todayStr();

  const hasReportToday = Array.isArray(currentDailyReportsCache)
    ? currentDailyReportsCache.some(
        row => Number(row.vehicle_id || 0) === Number(vehicleId) &&
               String(row.report_date || "") === String(reportDate)
      )
    : false;

  const extraDistance = Number(extraAssignedDistanceMap.get(Number(vehicleId)) || 0);
  const extraCount = Number(extraAssignedCountMap.get(Number(vehicleId)) || 0);
  const projectedDays = Math.max(1, Number(stats.workedDays || 0) + (hasReportToday ? 0 : (extraCount > 0 ? 1 : 0)));
  return (Number(stats.totalDistance || 0) + extraDistance) / projectedDays;
}

optimizeAssignmentsByDistanceBalance = function(assignments, items, vehicles, monthlyMap) {
  const working = (assignments || []).map(a => ({ ...a }));
  if (!working.length || !Array.isArray(vehicles) || !vehicles.length) return working;

  const itemMap = new Map((items || []).map(item => [Number(item.id), item]));
  const hourLoads = new Map();
  const assignedDistance = new Map();
  const assignedCount = new Map();

  const rebuild = () => {
    hourLoads.clear();
    assignedDistance.clear();
    assignedCount.clear();

    working.forEach(a => {
      const item = itemMap.get(Number(a.item_id));
      const itemDistance = Number(item?.distance_km ?? a?.distance_km ?? 0);
      const hourKey = `${Number(a.vehicle_id)}__${Number(a.actual_hour ?? 0)}`;
      hourLoads.set(hourKey, Number(hourLoads.get(hourKey) || 0) + 1);
      assignedDistance.set(Number(a.vehicle_id), Number(assignedDistance.get(Number(a.vehicle_id)) || 0) + itemDistance);
      assignedCount.set(Number(a.vehicle_id), Number(assignedCount.get(Number(a.vehicle_id)) || 0) + 1);
    });
  };

  const scoreForVehicle = (vehicle, item, currentHour) => {
    const area = getThemisV54RowArea(item);
    const compat =
      getStrictHomeCompatibilityScore(area, vehicle?.home_area || "") * 1.7 +
      Math.max(0, getDirectionAffinityScore(area, vehicle?.home_area || "")) * 0.95 +
      getAreaAffinityScore(area, vehicle?.home_area || "") * 0.85 +
      getVehicleAreaMatchScore(vehicle, area) * 0.55;

    const projectedAvg = getThemisV54VehicleProjectedAvg(vehicle?.id, monthlyMap, assignedDistance, assignedCount);
    const currentHourLoad = Number(hourLoads.get(`${Number(vehicle?.id)}__${Number(currentHour ?? 0)}`) || 0);

    return compat - projectedAvg * 1.35 - currentHourLoad * 18;
  };

  rebuild();

  for (const assignment of working) {
    const item = itemMap.get(Number(assignment.item_id));
    if (!item) continue;

    const currentVehicle = vehicles.find(v => Number(v.id) === Number(assignment.vehicle_id));
    const currentScore = scoreForVehicle(currentVehicle, item, assignment.actual_hour);
    let best = { vehicle: currentVehicle, score: currentScore };

    for (const vehicle of vehicles) {
      if (Number(vehicle.id) === Number(assignment.vehicle_id)) continue;
      const hourKey = `${Number(vehicle.id)}__${Number(assignment.actual_hour ?? 0)}`;
      const hourLoad = Number(hourLoads.get(hourKey) || 0);
      const seatCapacity = Number(vehicle.seat_capacity || 4);
      if (hourLoad >= seatCapacity) continue;
      if (isHardReverseForHome(getThemisV54RowArea(item), vehicle.home_area || "")) continue;

      const candidateScore = scoreForVehicle(vehicle, item, assignment.actual_hour);
      if (candidateScore > best.score) {
        best = { vehicle, score: candidateScore };
      }
    }

    if (best.vehicle && Number(best.vehicle.id) !== Number(assignment.vehicle_id) && best.score >= currentScore + 18) {
      assignment.vehicle_id = best.vehicle.id;
      assignment.driver_name = best.vehicle.driver_name || "";
      rebuild();
    }
  }

  return working;
};

applyLastTripDistanceCorrectionToAssignments = function(assignments, items, vehicles, monthlyMap) {
  const working = (assignments || []).map(a => ({ ...a }));
  if (!working.length || !Array.isArray(vehicles) || !vehicles.length) return working;

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
  const projectedCountByVehicle = new Map();
  working.forEach(a => {
    const item = itemMap.get(Number(a.item_id));
    projectedDistanceByVehicle.set(
      Number(a.vehicle_id),
      Number(projectedDistanceByVehicle.get(Number(a.vehicle_id)) || 0) +
        Number(item?.distance_km ?? a.distance_km ?? 0)
    );
    projectedCountByVehicle.set(
      Number(a.vehicle_id),
      Number(projectedCountByVehicle.get(Number(a.vehicle_id)) || 0) + 1
    );
  });

  const evaluate = (vehicle, item) => {
    const area = getThemisV54RowArea(item);
    const strict = getStrictHomeCompatibilityScore(area, vehicle?.home_area || "");
    const direction = Math.max(0, getDirectionAffinityScore(area, vehicle?.home_area || ""));
    const affinity = getAreaAffinityScore(area, vehicle?.home_area || "");
    const vehicleMatch = getVehicleAreaMatchScore(vehicle, area);
    const hardReverse = isHardReverseForHome(area, vehicle?.home_area || "");
    const projectedAvg = getThemisV54VehicleProjectedAvg(vehicle?.id, monthlyMap, projectedDistanceByVehicle, projectedCountByVehicle);

    let score = strict * 9 + direction * 5 + affinity * 3.8 + vehicleMatch * 0.8;
    score -= projectedAvg * 0.95;
    if (hardReverse) score -= 9999;
    if (Number(vehicle?.id) === Number(manualVehicleId) && strict >= 52 && !hardReverse) score += 140;
    if (isDriverLastTripChecked(vehicle?.id)) score += 120;
    return score;
  };

  for (const target of targetRows) {
    const item = itemMap.get(Number(target.item_id));
    if (!item) continue;

    const currentVehicle = vehicles.find(v => Number(v.id) === Number(target.vehicle_id));
    let best = { vehicle: currentVehicle, score: evaluate(currentVehicle, item) };

    for (const vehicle of vehicles) {
      const seatCapacity = Number(vehicle.seat_capacity || 4);
      const load = working.filter(
        a =>
          Number(a.vehicle_id) === Number(vehicle.id) &&
          Number(a.actual_hour ?? 0) === Number(targetHour) &&
          Number(a.item_id) !== Number(target.item_id)
      ).length;
      if (load >= seatCapacity) continue;

      const candidateScore = evaluate(vehicle, item);
      if (candidateScore > best.score) best = { vehicle, score: candidateScore };
    }

    if (best.vehicle && Number(best.vehicle.id) !== Number(target.vehicle_id) && best.score >= evaluate(currentVehicle, item) + 12) {
      target.vehicle_id = best.vehicle.id;
      target.driver_name = best.vehicle.driver_name || "";
      target.manual_last_vehicle = Number(best.vehicle.id) === Number(manualVehicleId);
    }
  }

  return working;
};

const _THEMIS_V54_BASE_runAutoDispatch = runAutoDispatch;
runAutoDispatch = async function() {
  const result = await _THEMIS_V54_BASE_runAutoDispatch.apply(this, arguments);
  try {
    await loadActualsByDate(els.actualDate?.value || todayStr());
    renderDailyDispatchResult();
  } catch (error) {
    console.error("THEMIS v5.4 rerender error:", error);
  }
  return result;
};

/* ===== THEMIS v5.4 配車AI強化版 patch end ===== */


/* ===== THEMIS v5.5.1 patch start ===== */
(function(){
  const _THEMIS_V55_BASE_optimizeAssignments = optimizeAssignments;

  function v551RoundTripMinutesForRow(row) {
    const stored = getStoredTravelMinutes(row?.casts?.travel_minutes ?? row?.travel_minutes);
    if (stored > 0) return stored * 2;
    return Math.max(0, Math.round(estimateFallbackTravelMinutes(Number(row?.distance_km || row?.casts?.distance_km || 0), normalizeAreaLabel(row?.destination_area || row?.casts?.area || '')) * 2));
  }

  function v551AreaSignature(area) {
    const normalized = normalizeAreaLabel(area || '');
    return {
      normalized,
      canonical: getCanonicalArea(normalized) || '',
      group: getAreaDisplayGroup(normalized) || ''
    };
  }

  function v551IsFriendlyDirection(areaA, areaB) {
    const a = v551AreaSignature(areaA);
    const b = v551AreaSignature(areaB);
    if (!a.normalized || !b.normalized) return false;
    if (a.canonical && b.canonical && a.canonical === b.canonical) return true;
    if (a.group && b.group && a.group === b.group) return true;
    const affinity = getAreaAffinityScore(a.normalized, b.normalized);
    const direction = getDirectionAffinityScore(a.normalized, b.normalized);
    return affinity >= 62 || direction >= 26;
  }

  function v551IsStrongReverse(areaA, areaB) {
    const a = normalizeAreaLabel(areaA || '');
    const b = normalizeAreaLabel(areaB || '');
    if (!a || !b) return false;
    return getDirectionAffinityScore(a, b) <= -38;
  }

  function v551VehicleStateSummary(assignments, vehicles, hour) {
    const map = new Map();
    (vehicles || []).forEach(v => map.set(Number(v.id), { count:0, areas:[], vehicle:v }));
    (assignments || []).filter(r => Number(r?.actual_hour ?? 0) === Number(hour)).forEach(r => {
      const id = Number(r.vehicle_id || 0);
      if (!map.has(id)) return;
      const state = map.get(id);
      state.count += 1;
      const area = normalizeAreaLabel(r.destination_area || r.cluster_area || r.planned_area || r.casts?.area || '');
      if (area) state.areas.push(area);
    });
    return map;
  }

  function v551CanMoveIntoVehicle(row, vehicleState, vehicle) {
    const cap = Number(vehicle?.seat_capacity || 4);
    if (Number(vehicleState?.count || 0) >= cap) return false;
    const rowArea = normalizeAreaLabel(row.destination_area || row.cluster_area || row.planned_area || row.casts?.area || '');
    const areas = Array.isArray(vehicleState?.areas) ? vehicleState.areas : [];
    if (!areas.length) return true;
    if (areas.some(a => v551IsStrongReverse(rowArea, a))) return false;
    return true;
  }

  function v551PickBundleVehicle(bundleRows, vehicleStates, vehicles, monthlyMap) {
    const first = bundleRows[0];
    const rowArea = normalizeAreaLabel(first.destination_area || first.cluster_area || first.planned_area || first.casts?.area || '');
    let best = null;
    for (const [vehicleId, state] of vehicleStates.entries()) {
      const vehicle = state.vehicle || (vehicles || []).find(v => Number(v.id) === Number(vehicleId));
      if (!vehicle) continue;
      const capacity = Number(vehicle.seat_capacity || 4);
      if (state.count + bundleRows.length > capacity) continue;
      if (!bundleRows.every(r => v551CanMoveIntoVehicle(r, state, vehicle))) continue;
      let score = 0;
      const areas = state.areas || [];
      if (areas.length) {
        for (const area of areas) {
          if (v551IsFriendlyDirection(rowArea, area)) score += 120;
          if (v551IsStrongReverse(rowArea, area)) score -= 400;
        }
      } else {
        score += 15;
      }
      const monthly = monthlyMap.get(Number(vehicleId)) || {};
      const avg = Number(monthly.averageDistance || 0);
      score -= avg * 0.08;
      score -= Number(state.count || 0) * 4;
      if (!best || score > best.score) best = { vehicleId:Number(vehicleId), score };
    }
    return best?.vehicleId || null;
  }

  function v551RebundleAssignments(assignments, items, vehicles, monthlyMap) {
    let rows = Array.isArray(assignments) ? assignments.map(r => ({ ...r })) : [];
    if (!rows.length) return rows;

    const grouped = new Map();
    rows.forEach(row => {
      const hour = Number(row?.actual_hour ?? 0);
      const area = normalizeAreaLabel(row.destination_area || row.cluster_area || row.planned_area || row.casts?.area || '');
      const sig = v551AreaSignature(area);
      const key = `${hour}__${sig.canonical || sig.group || area}`;
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(row);
    });

    for (const bundleRows of grouped.values()) {
      if (bundleRows.length < 2) continue;
      const first = bundleRows[0];
      const roundTrip = v551RoundTripMinutesForRow(first);
      if (roundTrip < 60) continue;
      const hour = Number(first.actual_hour ?? 0);
      const vehicleIds = [...new Set(bundleRows.map(r => Number(r.vehicle_id || 0)).filter(Boolean))];
      if (vehicleIds.length <= 1) continue;
      const vehicleStates = v551VehicleStateSummary(rows, vehicles, hour);
      const targetVehicleId = v551PickBundleVehicle(bundleRows, vehicleStates, vehicles, monthlyMap);
      if (!targetVehicleId) continue;
      rows = rows.map(r => {
        if (bundleRows.some(br => Number(br.id) === Number(r.id))) {
          return { ...r, vehicle_id: targetVehicleId };
        }
        return r;
      });
    }

    return rows;
  }

  optimizeAssignments = function(items, vehicles, monthlyMap) {
    let rows = _THEMIS_V55_BASE_optimizeAssignments.apply(this, arguments);
    if (!Array.isArray(rows)) rows = [];
    rows = v551RebundleAssignments(rows, items, vehicles, monthlyMap || new Map());
    return rows;
  };

  const _THEMIS_V55_BASE_buildRotationTimelineHtmlSafe = typeof buildRotationTimelineHtmlSafe === 'function' ? buildRotationTimelineHtmlSafe : null;
  buildRotationTimelineHtmlSafe = function(vehicles, activeItems) {
    try {
      const timeline = (Array.isArray(vehicles) ? vehicles : [])
        .map(vehicle => {
          const rows = (Array.isArray(activeItems) ? activeItems : []).filter(item => Number(item.vehicle_id) === Number(vehicle.id));
          if (!rows.length) return null;
          const orderedRows = (typeof moveManualLastItemsToEnd === 'function' && typeof sortItemsByNearestRoute === 'function')
            ? moveManualLastItemsToEnd(sortItemsByNearestRoute(rows))
            : rows;
          const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
          const summary = getVehicleDailySummary(vehicle, orderedRows);
          return {
            name: vehicle?.driver_name || vehicle?.plate_number || '-',
            returnAfterLabel: forecast?.returnAfterLabel || '-',
            nextRunTime: forecast?.predictedReadyTime || '-',
            rotationMinutes: Number(forecast?.predictedReturnMinutes || 0),
            totalKm: Number(summary?.totalKm || 0),
            totalJobs: Number(summary?.jobCount || 0)
          };
        })
        .filter(Boolean);
      if (!timeline.length) return '';
      return `
        <div class="panel-card" style="margin-bottom:16px;">
          <h3 style="margin-bottom:10px;">車両稼働タイムライン</h3>
          <div style="display:flex; flex-wrap:wrap; gap:8px;">
            ${timeline.map(item => `
              <div class="chip" style="padding:8px 12px;">
                <strong>${escapeHtml(item.name)}</strong>
                / 戻り ${escapeHtml(item.returnAfterLabel)}
                / 次便可能 ${escapeHtml(item.nextRunTime)}
                / 回転時間 ${Number(item.rotationMinutes || 0)}分
                / 累計 ${Number(item.totalKm || 0).toFixed(1)}km
                / ${Number(item.totalJobs || 0)}件
              </div>
            `).join('')}
          </div>
        </div>
      `;
    } catch (e) {
      console.error('buildRotationTimelineHtmlSafe v5.5.1 error:', e);
      return _THEMIS_V55_BASE_buildRotationTimelineHtmlSafe ? _THEMIS_V55_BASE_buildRotationTimelineHtmlSafe.apply(this, arguments) : '';
    }
  };

  const _THEMIS_V55_BASE_getVehicleRotationForecastSafe = getVehicleRotationForecastSafe;
  getVehicleRotationForecastSafe = function(vehicle, orderedRows) {
    const forecast = _THEMIS_V55_BASE_getVehicleRotationForecastSafe.apply(this, arguments) || {};
    if (!forecast.returnAfterLabel && Number.isFinite(forecast.predictedReturnMinutes)) {
      forecast.returnAfterLabel = `${Math.round(Number(forecast.predictedReturnMinutes || 0))}分後`;
    }
    if (!forecast.predictedReadyTime) forecast.predictedReadyTime = '-';
    return forecast;
  };
})();
/* ===== THEMIS v5.5.1 patch end ===== */


/* ===== THEMIS v5.5.3 patch start ===== */
(function(){
  function v553RoundTripMinutesForRow(row) {
    const stored = getStoredTravelMinutes(row?.casts?.travel_minutes ?? row?.travel_minutes);
    if (stored > 0) return Math.max(0, Math.round(stored * 2));
    const km = Number(row?.distance_km || row?.casts?.distance_km || 0);
    const area = normalizeAreaLabel(row?.destination_area || row?.cluster_area || row?.planned_area || row?.casts?.area || '');
    return Math.max(0, Math.round(estimateFallbackTravelMinutes(km, area) * 2));
  }

  function v553AreaNorm(rowOrArea) {
    if (typeof rowOrArea === 'string') return normalizeAreaLabel(rowOrArea || '');
    return normalizeAreaLabel(rowOrArea?.destination_area || rowOrArea?.cluster_area || rowOrArea?.planned_area || rowOrArea?.casts?.area || '');
  }

  function v553FriendlyArea(areaA, areaB) {
    const a = normalizeAreaLabel(areaA || '');
    const b = normalizeAreaLabel(areaB || '');
    if (!a || !b) return false;
    const ca = getCanonicalArea(a) || '';
    const cb = getCanonicalArea(b) || '';
    const ga = getAreaDisplayGroup(a) || '';
    const gb = getAreaDisplayGroup(b) || '';
    if (ca && cb && ca === cb) return true;
    if (ga && gb && ga === gb && getDirectionAffinityScore(a, b) > -20) return true;
    return getAreaAffinityScore(a, b) >= 60 || getDirectionAffinityScore(a, b) >= 22;
  }

  function v553HardReverseArea(areaA, areaB) {
    const a = normalizeAreaLabel(areaA || '');
    const b = normalizeAreaLabel(areaB || '');
    if (!a || !b) return false;
    return getAreaAffinityScore(a, b) <= 26 || getDirectionAffinityScore(a, b) <= -34;
  }

  function v553BuildVehicleStates(assignments, vehicles, hour) {
    const map = new Map();
    (Array.isArray(vehicles) ? vehicles : []).forEach(v => {
      map.set(Number(v.id), { vehicle: v, count: 0, rows: [], areas: [] });
    });
    (Array.isArray(assignments) ? assignments : [])
      .filter(r => Number(r?.actual_hour ?? 0) === Number(hour))
      .forEach(r => {
        const vid = Number(r?.vehicle_id || 0);
        if (!map.has(vid)) return;
        const state = map.get(vid);
        state.count += 1;
        state.rows.push(r);
        const area = v553AreaNorm(r);
        if (area) state.areas.push(area);
      });
    return map;
  }

  function v553FindLongFriendlyComponents(hourRows) {
    const rows = (Array.isArray(hourRows) ? hourRows : []).filter(r => v553RoundTripMinutesForRow(r) >= 55);
    const visited = new Set();
    const components = [];
    for (const row of rows) {
      const id = Number(row?.id || 0);
      if (!id || visited.has(id)) continue;
      const stack = [row];
      const component = [];
      visited.add(id);
      while (stack.length) {
        const cur = stack.pop();
        component.push(cur);
        const curArea = v553AreaNorm(cur);
        for (const other of rows) {
          const oid = Number(other?.id || 0);
          if (!oid || visited.has(oid)) continue;
          const otherArea = v553AreaNorm(other);
          if (v553FriendlyArea(curArea, otherArea) && !v553HardReverseArea(curArea, otherArea)) {
            visited.add(oid);
            stack.push(other);
          }
        }
      }
      if (component.length >= 2) components.push(component);
    }
    return components;
  }

  function v553VehicleCanAcceptComponent(state, vehicle, componentRows) {
    const seat = Math.max(1, Number(vehicle?.seat_capacity || 4));
    const componentIds = new Set((componentRows || []).map(r => Number(r?.id || 0)));
    const currentRows = Array.isArray(state?.rows) ? state.rows : [];
    const nonComponentRows = currentRows.filter(r => !componentIds.has(Number(r?.id || 0)));
    if (nonComponentRows.length + componentRows.length > seat) return false;
    const existingAreas = nonComponentRows.map(v553AreaNorm).filter(Boolean);
    for (const row of componentRows) {
      const area = v553AreaNorm(row);
      if (existingAreas.some(a => v553HardReverseArea(area, a))) return false;
    }
    return true;
  }

  function v553PickTargetVehicle(componentRows, states, monthlyMap) {
    const idsInComponent = new Set(componentRows.map(r => Number(r?.id || 0)));
    const componentAreas = componentRows.map(v553AreaNorm).filter(Boolean);
    let best = null;
    for (const [vehicleId, state] of states.entries()) {
      const vehicle = state.vehicle;
      if (!vehicle) continue;
      if (!v553VehicleCanAcceptComponent(state, vehicle, componentRows)) continue;

      const currentRows = Array.isArray(state.rows) ? state.rows : [];
      const componentAlreadyHere = currentRows.filter(r => idsInComponent.has(Number(r?.id || 0))).length;
      const nonComponentAreas = currentRows
        .filter(r => !idsInComponent.has(Number(r?.id || 0)))
        .map(v553AreaNorm)
        .filter(Boolean);

      let score = 0;
      score += componentAlreadyHere * 900;
      if (componentAlreadyHere > 0) score += 250;

      for (const compArea of componentAreas) {
        for (const area of nonComponentAreas) {
          if (v553FriendlyArea(compArea, area)) score += 120;
          if (v553HardReverseArea(compArea, area)) score -= 600;
        }
      }

      const monthly = monthlyMap?.get(Number(vehicleId)) || {};
      score -= Number(monthly.averageDistance || 0) * 0.04;
      score -= Number(state.count || 0) * 6;

      if (!best || score > best.score) best = { vehicleId: Number(vehicleId), score };
    }
    return best?.vehicleId || null;
  }

  function v553RebundleAssignments(assignments, vehicles, monthlyMap) {
    let rows = Array.isArray(assignments) ? assignments.map(r => ({ ...r })) : [];
    if (!rows.length) return rows;

    const hours = [...new Set(rows.map(r => Number(r?.actual_hour ?? 0)))];
    for (const hour of hours) {
      const hourRows = rows.filter(r => Number(r?.actual_hour ?? 0) === Number(hour));
      const components = v553FindLongFriendlyComponents(hourRows);
      if (!components.length) continue;

      for (const componentRows of components) {
        const currentVehicles = [...new Set(componentRows.map(r => Number(r?.vehicle_id || 0)).filter(Boolean))];
        if (currentVehicles.length <= 1) continue;
        const states = v553BuildVehicleStates(rows, vehicles, hour);
        const targetVehicleId = v553PickTargetVehicle(componentRows, states, monthlyMap);
        if (!targetVehicleId) continue;
        const componentIds = new Set(componentRows.map(r => Number(r?.id || 0)));
        rows = rows.map(r => componentIds.has(Number(r?.id || 0)) ? { ...r, vehicle_id: targetVehicleId } : r);
      }
    }

    return rows;
  }

  const _THEMIS_V553_BASE_optimizeAssignmentsByDistanceBalance = optimizeAssignmentsByDistanceBalance;
  optimizeAssignmentsByDistanceBalance = function(assignments, items, vehicles, monthlyMap) {
    let rows = _THEMIS_V553_BASE_optimizeAssignmentsByDistanceBalance.apply(this, arguments);
    if (!Array.isArray(rows)) rows = Array.isArray(assignments) ? assignments : [];
    rows = v553RebundleAssignments(rows, vehicles, monthlyMap || new Map());
    return rows;
  };

  const _THEMIS_V553_BASE_applyManualLastVehicleToAssignments = applyManualLastVehicleToAssignments;
  applyManualLastVehicleToAssignments = function(assignments, vehicles) {
    let rows = _THEMIS_V553_BASE_applyManualLastVehicleToAssignments.apply(this, arguments);
    if (!Array.isArray(rows)) rows = Array.isArray(assignments) ? assignments : [];
    const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
    rows = v553RebundleAssignments(rows, vehicles, monthlyMap || new Map());
    return rows;
  };

  function v553FormatClock(totalMinutes) {
    if (typeof formatClockTimeFromMinutesGlobal === 'function') return formatClockTimeFromMinutesGlobal(totalMinutes);
    if (typeof formatClockTimeFromMinutes === 'function') return formatClockTimeFromMinutes(totalMinutes);
    const safe = Math.max(0, Math.round(Number(totalMinutes || 0)));
    const h = Math.floor(safe / 60) % 24;
    const m = safe % 60;
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
  }

  const _THEMIS_V553_BASE_getVehicleRotationForecastSafe = getVehicleRotationForecastSafe;
  getVehicleRotationForecastSafe = function(vehicle, orderedRows) {
    const forecast = _THEMIS_V553_BASE_getVehicleRotationForecastSafe.apply(this, arguments) || {};
    const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
    if (!rows.length) {
      if (!forecast.returnAfterLabel) forecast.returnAfterLabel = '-';
      if (!forecast.predictedReadyTime) forecast.predictedReadyTime = '-';
      return forecast;
    }

    const timeSummary = getRowsTravelTimeSummary(rows);
    const isLastTrip = isDriverLastTripChecked(Number(vehicle?.id || 0));
    const fallbackMinutes = Math.max(0, Math.round(isLastTrip ? timeSummary.sendOnlyMinutes : timeSummary.totalMinutes));
    const runAt = Number(lastAutoDispatchRunAtMinutes || getCurrentClockMinutes() || 0);
    const fallbackReady = v553FormatClock(runAt + fallbackMinutes);

    const missingMinutes = !(Number(forecast.predictedReturnMinutes || 0) > 0);
    const missingReturnLabel = !forecast.returnAfterLabel || forecast.returnAfterLabel === '-' || forecast.returnAfterLabel === '0分後';
    const missingReady = !forecast.predictedReadyTime || forecast.predictedReadyTime === '-';
    const missingRotation = !(Number(forecast.rotationMinutes || 0) > 0);

    if (missingMinutes) forecast.predictedReturnMinutes = fallbackMinutes;
    if (missingReturnLabel) forecast.returnAfterLabel = `${fallbackMinutes}分後`;
    if (missingReady) forecast.predictedReadyTime = fallbackReady;
    if (missingRotation) forecast.rotationMinutes = fallbackMinutes;

    return forecast;
  };

  const _THEMIS_V553_BASE_buildRotationTimelineHtmlSafe = buildRotationTimelineHtmlSafe;
  buildRotationTimelineHtmlSafe = function(vehicles, activeItems) {
    try {
      const timeline = (Array.isArray(vehicles) ? vehicles : [])
        .map(vehicle => {
          const rows = (Array.isArray(activeItems) ? activeItems : []).filter(item => Number(item.vehicle_id) === Number(vehicle.id));
          if (!rows.length) return null;
          const orderedRows = (typeof moveManualLastItemsToEnd === 'function' && typeof sortItemsByNearestRoute === 'function')
            ? moveManualLastItemsToEnd(sortItemsByNearestRoute(rows))
            : rows;
          const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
          const summary = getVehicleDailySummary(vehicle, orderedRows);
          const rotationMinutes = Number(forecast?.rotationMinutes || forecast?.predictedReturnMinutes || 0);
          return {
            name: vehicle?.driver_name || vehicle?.plate_number || '-',
            returnAfterLabel: forecast?.returnAfterLabel || `${rotationMinutes}分後`,
            nextRunTime: forecast?.predictedReadyTime || '-',
            rotationMinutes,
            totalKm: Number(summary?.totalKm || 0),
            totalJobs: Number(summary?.jobCount || 0)
          };
        })
        .filter(Boolean);
      if (!timeline.length) return '';
      return `
        <div class="panel-card" style="margin-bottom:16px;">
          <h3 style="margin-bottom:10px;">車両稼働タイムライン</h3>
          <div style="display:flex; flex-wrap:wrap; gap:8px;">
            ${timeline.map(item => `
              <div class="chip" style="padding:8px 12px;">
                <strong>${escapeHtml(item.name)}</strong>
                / 戻り ${escapeHtml(item.returnAfterLabel)}
                / 次便可能 ${escapeHtml(item.nextRunTime)}
                / 回転時間 ${Math.round(Number(item.rotationMinutes || 0))}分
                / 累計 ${Number(item.totalKm || 0).toFixed(1)}km
                / ${Number(item.totalJobs || 0)}件
              </div>
            `).join('')}
          </div>
        </div>
      `;
    } catch (e) {
      console.error('buildRotationTimelineHtmlSafe v5.5.3 error:', e);
      return _THEMIS_V553_BASE_buildRotationTimelineHtmlSafe.apply(this, arguments);
    }
  };
})();

function renderDispatchSummaryFromDOM(){
  const resultWrap = document.getElementById("dailyDispatchResult");
  const bar = document.getElementById("dispatchSummaryBar");
  if(!resultWrap || !bar) return;

  const cards = resultWrap.querySelectorAll(".dispatch-card, .vehicle-card, .result-card");

  let vehicleCount = 0;
  let totalJobs = 0;
  let times = [];

  cards.forEach(card=>{
    vehicleCount++;

    const jobs = card.querySelectorAll(".job, .dispatch-job, li");
    totalJobs += jobs.length;

    // 時間っぽいテキスト探す（例: 45分）
    const text = card.innerText;
    const match = text.match(/(\d+)\s*分/);
    if(match){
      times.push(parseInt(match[1]));
    }
  });

  if(vehicleCount === 0){
    bar.innerHTML = "";
    return;
  }

  let min = "-", max = "-", avg = "-";

  if(times.length){
    min = Math.min(...times);
    max = Math.max(...times);
    avg = Math.round(times.reduce((a,b)=>a+b,0)/times.length);
  }

  bar.innerHTML = `
    🚗 ${vehicleCount}台｜
    📦 ${totalJobs}件｜
    ⏱ 最短 ${min}｜
    ⏳ 最長 ${max}｜
    📉 平均 ${avg}
  `;
}

// フック（既存処理に干渉しない）
const _origRender = window.renderDailyDispatchResult;
window.renderDailyDispatchResult = function(){
  if(_origRender) _origRender();
  setTimeout(renderDispatchSummaryFromDOM, 50);
};

/* ===== THEMIS v6.9.13 drop-order unify hotfix start ===== */
(function(){
  function __themisOrderDistance(row) {
    const v = Number(row?.distance_km ?? row?.casts?.distance_km ?? 0);
    if (Number.isFinite(v) && v > 0) return v;
    const pt = (typeof getThemisV54RowPoint === 'function') ? getThemisV54RowPoint(row) : null;
    if (pt && typeof estimateRoadKmBetweenPoints === 'function') {
      return Number(estimateRoadKmBetweenPoints(ORIGIN_LAT, ORIGIN_LNG, pt.lat, pt.lng) || 999999);
    }
    return 999999;
  }

  function __themisOrderTravel(row) {
    if (typeof getThemisV54StoredTravelMinutes === 'function') {
      const v = Number(getThemisV54StoredTravelMinutes(row) || 0);
      if (Number.isFinite(v) && v > 0) return v;
    }
    const v = Number(row?.travel_minutes ?? row?.casts?.travel_minutes ?? 0);
    return Number.isFinite(v) && v > 0 ? v : 999999;
  }

  function __themisOrderId(row) {
    const v = Number(row?.id || row?.cast_id || row?.casts?.id || 0);
    return Number.isFinite(v) ? v : 0;
  }

  function __themisCompareByOrigin(a, b) {
    const da = __themisOrderDistance(a);
    const db = __themisOrderDistance(b);
    if (da !== db) return da - db;

    const ta = __themisOrderTravel(a);
    const tb = __themisOrderTravel(b);
    if (ta !== tb) return ta - tb;

    return __themisOrderId(a) - __themisOrderId(b);
  }

  function __themisCompareFromPrevious(prev, a, b) {
    const prevD = __themisOrderDistance(prev);
    const da = Math.abs(__themisOrderDistance(a) - prevD);
    const db = Math.abs(__themisOrderDistance(b) - prevD);
    if (da !== db) return da - db;

    const oa = __themisOrderDistance(a);
    const ob = __themisOrderDistance(b);
    if (oa !== ob) return oa - ob;

    const ta = __themisOrderTravel(a);
    const tb = __themisOrderTravel(b);
    if (ta !== tb) return ta - tb;

    return __themisOrderId(a) - __themisOrderId(b);
  }

  sortItemsByNearestRoute = function(items) {
    const remaining = [...(items || [])].filter(Boolean);
    if (!remaining.length) return [];

    remaining.sort(__themisCompareByOrigin);
    const ordered = [remaining.shift()];

    while (remaining.length) {
      const prev = ordered[ordered.length - 1];
      remaining.sort((a, b) => __themisCompareFromPrevious(prev, a, b));
      ordered.push(remaining.shift());
    }

    return ordered;
  };
})();
/* ===== THEMIS v6.9.13 drop-order unify hotfix end ===== */
