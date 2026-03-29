
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
const ENABLE_DISPLAY_GROUP_FORCE_BRANCH = true;

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


function getOnTheWayCompatibility(areaA, areaB) {
  const a = normalizeAreaLabel(areaA);
  const b = normalizeAreaLabel(areaB);
  if (!a || !b || a === "無し" || b === "無し") return 0;

  const corridor = ["柏", "我孫子", "取手", "藤代", "牛久", "ひたち野うしく"];
  const idxA = corridor.findIndex(k => a.includes(k));
  const idxB = corridor.findIndex(k => b.includes(k));

  if (idxA === -1 || idxB === -1) return 0;

  const diff = Math.abs(idxA - idxB);
  if (diff === 0) return 80;
  if (diff === 1) return 100;
  if (diff === 2) return 60;
  return 0;
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

  const onTheWay = getOnTheWayCompatibility(areaA, areaB);
  if (onTheWay >= 80) penalty -= 120;
  else if (onTheWay >= 50) penalty -= 70;

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

function getVehicleMonthlyStatsMap(reportRows, targetMonth) {
  const map = new Map();

  function resolveVehicleIdFromReport(row) {
    const directId = Number(row?.vehicle_id || row?.vehicles?.id || 0);
    if (directId > 0) return directId;

    const reportDriver = String(row?.driver_name || row?.vehicles?.driver_name || "").trim();
    const reportPlate = String(row?.plate_number || row?.vehicles?.plate_number || "").trim();

    const matchedVehicle = (Array.isArray(allVehiclesCache) ? allVehiclesCache : []).find(vehicle => {
      const vehicleDriver = String(vehicle?.driver_name || "").trim();
      const vehiclePlate = String(vehicle?.plate_number || "").trim();
      return (reportDriver && vehicleDriver && reportDriver === vehicleDriver) ||
             (reportPlate && vehiclePlate && reportPlate === vehiclePlate);
    });

    return Number(matchedVehicle?.id || 0);
  }

  (Array.isArray(reportRows) ? reportRows : []).forEach(row => {
    if (getMonthKey(row?.report_date) !== targetMonth) return;

    const vehicleId = resolveVehicleIdFromReport(row);
    if (!vehicleId) return;

    const prev = map.get(vehicleId) || {
      totalDistance: 0,
      workedDays: 0,
      avgDistance: 0,
      byDate: new Map()
    };

    const dateKey = String(row?.report_date || "").trim();
    const distance = Number(row?.distance_km || 0);
    prev.totalDistance += distance;
    if (dateKey) {
      prev.byDate.set(dateKey, Number((prev.byDate.get(dateKey) || 0) + distance));
    }
    prev.workedDays = [...prev.byDate.values()].filter(value => Number(value || 0) > 0).length;
    prev.avgDistance = prev.workedDays > 0 ? prev.totalDistance / prev.workedDays : 0;
    map.set(vehicleId, prev);
  });

  map.forEach(stats => {
    stats.totalDistance = Number(Number(stats.totalDistance || 0).toFixed(1));
    stats.avgDistance = Number(Number(stats.avgDistance || 0).toFixed(1));
    delete stats.byDate;
  });

  return map;
}



function getUnifiedMonthlyUiStatsMap(reportRows = currentDailyReportsCache, baseDate) {
  const targetDate = baseDate || (els?.dispatchDate?.value || todayStr());
  const startDate = getMonthStartStr(targetDate);
  const endDate = getMonthEndStr(targetDate);
  const normalizedRows = normalizeMileageExportRows(
    (Array.isArray(reportRows) ? reportRows : []).filter(row => {
      const d = String(row?.report_date || "");
      return d && d >= startDate && d <= endDate;
    })
  );

  const calendar = buildMileageCalendarRows(normalizedRows, startDate, endDate);
  const map = new Map();

  const vehicles = Array.isArray(allVehiclesCache) ? allVehiclesCache : [];
  const normalizeText = value => String(value || "").trim();
  const findVehicleId = entry => {
    const directId = Number(entry?.vehicle_id || 0);
    if (directId > 0) return directId;

    const driver = normalizeText(entry?.driver_name);
    const plate = normalizeText(entry?.plate_number);

    const exact = vehicles.find(vehicle => {
      const vehicleDriver = normalizeText(vehicle?.driver_name);
      const vehiclePlate = normalizeText(vehicle?.plate_number);
      return (plate && vehiclePlate && plate === vehiclePlate) ||
             (driver && vehicleDriver && driver === vehicleDriver);
    });
    return Number(exact?.id || 0);
  };

  (calendar?.drivers || []).forEach(entry => {
    const vehicleId = findVehicleId(entry);
    if (!vehicleId) return;
    map.set(vehicleId, {
      totalDistance: Number(Number(entry.total_distance_km || 0).toFixed(1)),
      workedDays: Number(entry.worked_days || 0),
      avgDistance: Number(Number(entry.avg_distance_km || 0).toFixed(1))
    });
  });

  return map;
}

function buildMonthlyDistanceMapForCurrentMonth() {
  try {
    const targetDate = typeof getSelectedDispatchDate === "function"
      ? getSelectedDispatchDate()
      : (new Date()).toISOString().slice(0, 10);
    return getUnifiedMonthlyUiStatsMap(
      Array.isArray(currentDailyReportsCache) ? currentDailyReportsCache : [],
      targetDate
    );
  } catch (_) {
    return new Map();
  }
}

function optimizeAssignmentsByRouteFlow(assignments, items, vehicles) {
  return Array.isArray(assignments) ? assignments : [];
}

function applyManualLastVehicleToAssignments(assignments, vehicles) {
  return Array.isArray(assignments) ? assignments : [];
}

function joinManualLastVehicleToAssignments(assignments, vehicles) {
  return applyManualLastVehicleToAssignments(assignments, vehicles);
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

  if (tabId === "vehiclesTab") {
    const targetDate = els.dispatchDate?.value || todayStr();
    forceResetMileageReportInputs(targetDate);
    window.requestAnimationFrame(() => forceResetMileageReportInputs(els.dispatchDate?.value || todayStr()));
    window.setTimeout(() => forceResetMileageReportInputs(els.dispatchDate?.value || todayStr()), 0);
    window.setTimeout(() => forceResetMileageReportInputs(els.dispatchDate?.value || todayStr()), 120);
  }
}

function setupTabs() {
  document.querySelectorAll(".main-tab").forEach(btn => {
    btn.addEventListener("click", () => activateTab(btn.dataset.tab));
  });

  document.querySelectorAll(".go-tab-btn").forEach(btn => {
    btn.addEventListener("click", () => activateTab(btn.dataset.goTab));
  });
}

function getSortedVehiclesForDisplay() {
  const list = Array.isArray(allVehiclesCache) ? [...allVehiclesCache] : [];
  return list.sort((a, b) => {
    const plateA = String(a?.plate_number || "").trim();
    const plateB = String(b?.plate_number || "").trim();

    const matchA = plateA.match(/^([A-Za-z])(\d+)?$/);
    const matchB = plateB.match(/^([A-Za-z])(\d+)?$/);

    if (matchA && matchB) {
      const alpha = matchA[1].localeCompare(matchB[1], "en", { sensitivity: "base" });
      if (alpha !== 0) return alpha;
      return Number(matchA[2] || 0) - Number(matchB[2] || 0);
    }

    const plain = plateA.localeCompare(plateB, "ja", { numeric: true, sensitivity: "base" });
    if (plain !== 0) return plain;

    return Number(a?.id || 0) - Number(b?.id || 0);
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

function getMileageSelectedRange() {
  const baseDate = els.dispatchDate?.value || todayStr();
  return {
    startDate: els.mileageReportStartDate?.value || getMonthStartStr(baseDate),
    endDate: els.mileageReportEndDate?.value || baseDate
  };
}

function getDashboardMonthlyRange() {
  const baseDate = els.dispatchDate?.value || todayStr();
  return {
    startDate: getMonthStartStr(baseDate),
    endDate: baseDate,
    monthKey: getMonthKey(baseDate)
  };
}

function getVehicleStatsMapForDashboardMonth(reportRows = currentDailyReportsCache) {
  const { endDate } = getDashboardMonthlyRange();
  return getUnifiedMonthlyUiStatsMap(reportRows, endDate);
}

function renderHomeMonthlyVehicleList(reportRows = currentDailyReportsCache) {
  if (!els.homeMonthlyVehicleList) return;

  const statsMap = getVehicleStatsMapForDashboardMonth(reportRows);

  els.homeMonthlyVehicleList.innerHTML = "";

  if (!allVehiclesCache.length) {
    els.homeMonthlyVehicleList.innerHTML = `<div class="chip">車両なし</div>`;
    return;
  }

  getSortedVehiclesForDisplay().forEach(vehicle => {
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

async function refreshHomeMonthlyVehicleList() {
  if (!els.homeMonthlyVehicleList) return;
  const { startDate, endDate } = getDashboardMonthlyRange();

  if (typeof fetchDriverMileageRows !== 'function') {
    renderHomeMonthlyVehicleList(currentDailyReportsCache);
    return;
  }

  try {
    const freshRows = await fetchDriverMileageRows(startDate, endDate);
    renderHomeMonthlyVehicleList(Array.isArray(freshRows) ? freshRows : currentDailyReportsCache);
  } catch (error) {
    console.error('refreshHomeMonthlyVehicleList error:', error);
    renderHomeMonthlyVehicleList(currentDailyReportsCache);
  }
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

function applyCastLatLng() {
  const parsed = parseLatLngText(els.castLatLngText?.value || "");
  if (!parsed) {
    alert("座標形式が正しくありません");
    return;
  }

  if (els.castLat) els.castLat.value = parsed.lat;
  if (els.castLng) els.castLng.value = parsed.lng;

  if (els.castArea) {
    els.castArea.value = getCastManagementAreaLabel(
      parsed.lat,
      parsed.lng,
      els.castAddress?.value || ""
    );
  }

  if (els.castDistanceKm) {
    els.castDistanceKm.value = String(estimateRoadKmFromStation(parsed.lat, parsed.lng));
  }
  lastCastGeocodeKey = normalizeGeocodeAddressKey(els.castAddress?.value || "");
  setCastGeoStatus("manual", "✔ 座標反映済（手動入力）");
}

function getCastManagementAreaLabel(lat, lng, address = "") {
  return normalizeAreaLabel(guessArea(lat, lng, address));
}

function guessCastArea() {
  const lat = toNullableNumber(els.castLat?.value);
  const lng = toNullableNumber(els.castLng?.value);
  if (els.castArea) {
    els.castArea.value = getCastManagementAreaLabel(lat, lng, els.castAddress?.value || "");
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

function renderVehiclesTable() {
  if (!els.vehiclesTableBody) return;

  const statsMap = getVehicleStatsMapForDashboardMonth(currentDailyReportsCache);

  els.vehiclesTableBody.innerHTML = "";

  if (!allVehiclesCache.length) {
    els.vehiclesTableBody.innerHTML = `<tr><td colspan="9" class="muted">車両がありません</td></tr>`;
    return;
  }

  getSortedVehiclesForDisplay().forEach(vehicle => {
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


function syncMileageReportRange(dateStr = todayStr(), force = false) {
  const targetDate = dateStr || todayStr();
  const start = getMonthStartStr(targetDate);
  const end = targetDate;

  if (force) {
    forceResetMileageReportInputs(targetDate);
    return;
  }

  if (els.mileageReportStartDate && !els.mileageReportStartDate.value) {
    applyMileageDateValue(els.mileageReportStartDate, start);
  }
  if (els.mileageReportEndDate && !els.mileageReportEndDate.value) {
    applyMileageDateValue(els.mileageReportEndDate, end);
  }
}

function initializeMileageReportDefaultDates() {
  const today = todayStr();
  forceResetMileageReportInputs(today);
}

function applyMileageDateValue(input, value) {
  if (!input) return;
  input.value = value;
  input.defaultValue = value;
  try { input.setAttribute("value", value); } catch (e) {}
}

function replaceMileageDateInput(inputId) {
  const current = document.getElementById(inputId);
  if (!current || !current.parentNode) return current;
  const clone = current.cloneNode(true);
  clone.value = "";
  clone.defaultValue = "";
  try { clone.removeAttribute("value"); } catch (e) {}
  current.parentNode.replaceChild(clone, current);
  return clone;
}

function forceResetMileageReportInputs(dateStr = todayStr()) {
  const targetDate = dateStr || todayStr();
  const start = getMonthStartStr(targetDate);
  const end = targetDate;

  const startInput = replaceMileageDateInput("mileageReportStartDate");
  const endInput = replaceMileageDateInput("mileageReportEndDate");

  els.mileageReportStartDate = startInput || document.getElementById("mileageReportStartDate");
  els.mileageReportEndDate = endInput || document.getElementById("mileageReportEndDate");

  applyMileageDateValue(els.mileageReportStartDate, start);
  applyMileageDateValue(els.mileageReportEndDate, end);
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

  const targetDate = todayStr();
  const startDate = els.mileageReportStartDate?.value || getMonthStartStr(targetDate);
  const endDate = els.mileageReportEndDate?.value || targetDate;

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


function syncScheduleRendererDeps() {
  if (typeof window.setScheduleRendererDeps !== "function") return;

  window.setScheduleRendererDeps({
    els,
    plans: currentPlansCache,
    actuals: currentActualsCache,
    actions: {
      fillPlanForm,
      openGoogleMap,
      deletePlan
    },
    helpers: {
      normalizeStatus,
      buildMapLinkHtml,
      escapeHtml,
      normalizeAreaLabel,
      getStatusText,
      getHourLabel,
      getAreaDisplayGroup,
      getGroupedAreasByDisplay,
      getGroupedAreaHeaderHtml,
      todayStr
    }
  });
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
  syncScheduleRendererDeps();
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
    payload.created_by = getCurrentUserIdSafe();
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
        created_by: getCurrentUserIdSafe()
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
  syncScheduleRendererDeps();
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

  getSortedVehiclesForDisplay().forEach(vehicle => {
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
  return getSortedVehiclesForDisplay().filter(v => activeVehicleIdsForToday.has(Number(v.id)));
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
  const rawGroup =
    item?.display_group ||
    item?.area_group ||
    item?.group_area ||
    item?.actual_group ||
    item?.group ||
    item?.destination_area ||
    item?.cluster_area ||
    item?.planned_area ||
    item?.casts?.area ||
    "無し";
  const area = normalizeAreaLabel(rawGroup);
  if (typeof THEMIS_DISPLAY_GROUPS !== "undefined" && THEMIS_DISPLAY_GROUPS && THEMIS_DISPLAY_GROUPS.has(area)) return area;
  return getAreaDisplayGroup(area) || area || "東京方面";
}

function __rowDistanceForCapacitySplit(row) {
  return Number(
    row?.distance_km ??
    row?.casts?.distance_km ??
    row?.distanceKm ??
    0
  ) || 0;
}

function __rowTravelMinutesForCapacitySplit(row) {
  return Number(
    row?.travel_minutes ??
    row?.casts?.travel_minutes ??
    row?.travelMinutes ??
    0
  ) || 0;
}

function __sortRowsForGroupCapacitySplit(rows) {
  const safeRows = Array.isArray(rows) ? rows.filter(Boolean) : [];
  return [...safeRows].sort((a, b) => {
    const dist = __rowDistanceForCapacitySplit(b) - __rowDistanceForCapacitySplit(a);
    if (dist !== 0) return dist;
    const tm = __rowTravelMinutesForCapacitySplit(b) - __rowTravelMinutesForCapacitySplit(a);
    if (tm !== 0) return tm;
    return Number(a?.id || 0) - Number(b?.id || 0);
  });
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
      const aMax = Math.max(...a[1].map(row => Number(row?.distance_km || 0)), 0);
      const bMax = Math.max(...b[1].map(row => Number(row?.distance_km || 0)), 0);
      if (bMax !== aMax) return bMax - aMax;
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
      let remaining = __sortRowsForGroupCapacitySplit(rows);

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
        const chunk = __sortRowsForGroupCapacitySplit(remaining.splice(0, freeSeats));
        let stopOrder = picked.count + 1;

        for (const row of chunk) {
          assignments.push({
            item_id: Number(row?.id || 0),
            actual_hour: hour,
            vehicle_id: picked.id,
            driver_name: picked.vehicle?.driver_name || "",
            distance_km: __rowDistanceForCapacitySplit(row),
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
  if (!Array.isArray(assignments)) assignments = [];

  assignments = resolveCapacityOverflowLocally(
    assignments,
    unassignedItems,
    selectedVehicles,
    monthlyMap
  );

  if ((!Array.isArray(assignments) || !assignments.length) && typeof __buildEmergencyAssignments === "function") {
    assignments = __buildEmergencyAssignments(unassignedItems, selectedVehicles);
    assignments = resolveCapacityOverflowLocally(
      assignments,
      unassignedItems,
      selectedVehicles,
      monthlyMap
    );
  }

  if (!assignments.length) return;

  await applyAutoDispatchAssignments(assignments);
}

function __buildEmergencyAssignments(activeItems, selectedVehicles) {
  const items = Array.isArray(activeItems) ? activeItems.filter(Boolean) : [];
  const vehicles = Array.isArray(selectedVehicles) ? selectedVehicles.filter(Boolean) : [];
  if (!items.length || !vehicles.length) return [];

  const sortedItems = typeof sortItemsByNearestRoute === "function"
    ? sortItemsByNearestRoute(items)
    : [...items].sort((a, b) => Number(a?.actual_hour || 0) - Number(b?.actual_hour || 0));

  const seatMap = new Map(
    vehicles.map(v => [Number(v.id), Math.max(1, Number(v?.seat_capacity || 4))])
  );

  const hourVehicleLoads = new Map();
  const assignments = [];

  function getLoad(hour, vehicleId) {
    return hourVehicleLoads.get(`${hour}__${vehicleId}`) || 0;
  }
  function addLoad(hour, vehicleId) {
    const key = `${hour}__${vehicleId}`;
    hourVehicleLoads.set(key, getLoad(hour, vehicleId) + 1);
  }

  sortedItems.forEach((item, index) => {
    const hour = Number(item?.actual_hour ?? item?.plan_hour ?? 0);
    const area = normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.casts?.area || "無し");

    let bestVehicle = null;
    let bestScore = -Infinity;

    vehicles.forEach((vehicle, vehicleIndex) => {
      const vehicleId = Number(vehicle.id);
      const currentLoad = getLoad(hour, vehicleId);
      const seatCapacity = seatMap.get(vehicleId) || 4;
      if (currentLoad >= seatCapacity) return;

      const homeArea = normalizeAreaLabel(vehicle?.home_area || "");
      const vehicleArea = normalizeAreaLabel(vehicle?.vehicle_area || "");

      let score = 0;
      if (typeof getVehicleAreaMatchScore === "function") {
        score += Number(getVehicleAreaMatchScore(vehicle, area) || 0);
      }
      if (typeof getAreaAffinityScore === "function") {
        score += Number(getAreaAffinityScore(vehicleArea, area) || 0) * 0.6;
        score += Number(getAreaAffinityScore(homeArea, area) || 0) * 0.35;
      }
      if (typeof getDirectionAffinityScore === "function") {
        score += Number(getDirectionAffinityScore(vehicleArea, area) || 0) * 0.25;
      }

      const siblingAssignments = assignments.filter(a => Number(a.actual_hour ?? 0) === Number(hour) && Number(a.vehicle_id) === Number(vehicleId));
      const siblingAreas = siblingAssignments.map(a => {
        const existing = items.find(x => Number(x.id) === Number(a.item_id));
        return normalizeAreaLabel(existing?.destination_area || existing?.cluster_area || existing?.casts?.area || '');
      }).filter(Boolean);
      const siblingAnchor = siblingAssignments
        .map(a => items.find(x => Number(x.id) === Number(a.item_id)))
        .filter(Boolean)
        .sort((a, b) => Number(b?.distance_km || 0) - Number(a?.distance_km || 0))[0] || null;
      const siblingAnchorArea = normalizeAreaLabel(siblingAnchor?.destination_area || siblingAnchor?.cluster_area || siblingAnchor?.casts?.area || '');
      const siblingAnchorDistance = Number(siblingAnchor?.distance_km || 0);
      const itemDistance = Number(item?.distance_km || 0);
      let bundlePenalty = 0;
      siblingAreas.forEach(existingArea => {
        const direction = Number(getDirectionAffinityScore(existingArea, area) || 0);
        if (direction <= -38 || isHardReverseMixForRoute(existingArea, area)) bundlePenalty += 1400;
        else bundlePenalty += Math.max(0, 34 - direction) * 7;
      });
      if (siblingAnchorArea) {
        const anchorDirection = Number(getDirectionAffinityScore(siblingAnchorArea, area) || 0);
        if (anchorDirection <= -38 || isHardReverseMixForRoute(siblingAnchorArea, area)) bundlePenalty += 1800;
        else bundlePenalty += Math.max(0, 40 - anchorDirection) * 9;
        bundlePenalty += Math.abs(siblingAnchorDistance - itemDistance) * 0.7;
      }

      score -= bundlePenalty;
      score -= currentLoad * 18;
      score -= vehicleIndex * 0.01;
      score -= index * 0.001;

      if (bestVehicle === null || score > bestScore) {
        bestVehicle = vehicle;
        bestScore = score;
      }
    });

    if (!bestVehicle) return;

    addLoad(hour, Number(bestVehicle.id));
    assignments.push({
      item_id: Number(item.id),
      vehicle_id: Number(bestVehicle.id),
      driver_name: bestVehicle.driver_name || "",
      actual_hour: hour,
      manual_last_vehicle: typeof isManualLastVehicle === "function"
        ? Boolean(isManualLastVehicle(bestVehicle.id))
        : false
    });
  });

  return assignments;
}


/* ===== THEMIS overflow local resolver start ===== */

function __overflowGetAssignmentItem(ctx, assignment) {
  return ctx.itemMap.get(Number(assignment?.item_id || 0)) || null;
}

function __overflowGetVehicleHourKey(vehicleId, hour) {
  return `${Number(vehicleId || 0)}__${Number(hour || 0)}`;
}

function __overflowBuildContext(assignments, items, vehicles) {
  const itemMap = new Map((Array.isArray(items) ? items : []).map(item => [Number(item.id), item]));
  const vehicleMap = new Map((Array.isArray(vehicles) ? vehicles : []).map(vehicle => [Number(vehicle.id), vehicle]));
  const vehicleHourRowsMap = new Map();

  (Array.isArray(assignments) ? assignments : []).forEach(assignment => {
    const key = __overflowGetVehicleHourKey(assignment.vehicle_id, assignment.actual_hour);
    if (!vehicleHourRowsMap.has(key)) vehicleHourRowsMap.set(key, []);
    vehicleHourRowsMap.get(key).push(assignment);
  });

  return {
    itemMap,
    vehicleMap,
    vehicleHourRowsMap
  };
}

function __overflowFindBuckets(ctx) {
  const buckets = [];

  ctx.vehicleHourRowsMap.forEach((rows, key) => {
    const [vehicleId, hour] = String(key).split("__").map(Number);
    const vehicle = ctx.vehicleMap.get(Number(vehicleId));
    const seatCapacity = Math.max(1, Number(vehicle?.seat_capacity || 4));
    const count = Array.isArray(rows) ? rows.length : 0;

    if (count > seatCapacity) {
      buckets.push({
        key,
        vehicleId: Number(vehicleId),
        hour: Number(hour),
        seatCapacity,
        count,
        overCount: count - seatCapacity
      });
    }
  });

  buckets.sort((a, b) => {
    if (b.overCount !== a.overCount) return b.overCount - a.overCount;
    if (a.hour !== b.hour) return a.hour - b.hour;
    return a.vehicleId - b.vehicleId;
  });

  return buckets;
}

function __overflowGetRowDistanceFromOrigin(ctx, assignment) {
  const item = __overflowGetAssignmentItem(ctx, assignment);
  return Number(item?.distance_km || item?.casts?.distance_km || assignment?.distance_km || 0);
}

function __overflowSelectRows(sourceRows, overCount, ctx) {
  return [...(Array.isArray(sourceRows) ? sourceRows : [])]
    .sort((a, b) => {
      const ad = __overflowGetRowDistanceFromOrigin(ctx, a);
      const bd = __overflowGetRowDistanceFromOrigin(ctx, b);
      if (ad !== bd) return ad - bd;
      return Number(a?.item_id || 0) - Number(b?.item_id || 0);
    })
    .slice(0, Math.max(0, Number(overCount || 0)));
}

function __overflowGetAreaFromAssignment(ctx, assignment) {
  const item = __overflowGetAssignmentItem(ctx, assignment);
  return normalizeAreaLabel(
    item?.destination_area ||
    item?.cluster_area ||
    item?.planned_area ||
    item?.casts?.area ||
    ""
  );
}

function __overflowGetOrderedRowsFromAssignments(ctx, assignments) {
  const rows = (Array.isArray(assignments) ? assignments : [])
    .map(assignment => {
      const item = __overflowGetAssignmentItem(ctx, assignment);
      return item ? { ...item } : null;
    })
    .filter(Boolean);

  if (typeof sortItemsByNearestRoute === "function") {
    const sorted = sortItemsByNearestRoute(rows);
    if (typeof moveManualLastItemsToEnd === "function") {
      return moveManualLastItemsToEnd(sorted);
    }
    return sorted;
  }

  return rows.sort((a, b) => {
    const ad = Number(a?.distance_km || 0);
    const bd = Number(b?.distance_km || 0);
    if (ad !== bd) return ad - bd;
    return Number(a?.id || 0) - Number(b?.id || 0);
  });
}

function __overflowIsReverseDirectionAgainstVehicle(ctx, candidateAssignment, existingAssignments, vehicle) {
  const targetArea = __overflowGetAreaFromAssignment(ctx, candidateAssignment);
  const existingAreas = (Array.isArray(existingAssignments) ? existingAssignments : [])
    .map(row => __overflowGetAreaFromAssignment(ctx, row))
    .filter(Boolean);

  if (existingAreas.some(area => isHardReverseMixForRoute(area, targetArea))) {
    return true;
  }

  const vehicleHomeArea = normalizeAreaLabel(vehicle?.home_area || "");
  if (vehicleHomeArea && isHardReverseMixForRoute(vehicleHomeArea, targetArea)) {
    return true;
  }

  const vehicleArea = normalizeAreaLabel(vehicle?.vehicle_area || "");
  if (vehicleArea && isHardReverseMixForRoute(vehicleArea, targetArea)) {
    return true;
  }

  return false;
}

function __overflowGetTargetCandidates(ctx, movingAssignment, bucket, vehicles) {
  const candidates = [];
  const movingArea = __overflowGetAreaFromAssignment(ctx, movingAssignment);

  (Array.isArray(vehicles) ? vehicles : []).forEach(vehicle => {
    const vehicleId = Number(vehicle?.id || 0);
    if (!vehicleId) return;
    if (vehicleId === Number(bucket.vehicleId)) return;

    const key = __overflowGetVehicleHourKey(vehicleId, bucket.hour);
    const existingAssignments = ctx.vehicleHourRowsMap.get(key) || [];
    const seatCapacity = Math.max(1, Number(vehicle?.seat_capacity || 4));

    if (existingAssignments.length >= seatCapacity) return;
    if (__overflowIsReverseDirectionAgainstVehicle(ctx, movingAssignment, existingAssignments, vehicle)) return;

    const existingAreas = existingAssignments
      .map(row => __overflowGetAreaFromAssignment(ctx, row))
      .filter(Boolean);

    let normalSpreadOk = true;
    for (const area of existingAreas) {
      const direction = Number(getDirectionAffinityScore(area, movingArea) || 0);
      if (direction <= -38) {
        normalSpreadOk = false;
        break;
      }
    }
    if (!normalSpreadOk) return;

    candidates.push({
      vehicleId,
      vehicle,
      hour: bucket.hour,
      existingAssignments
    });
  });

  return candidates;
}

function __overflowEvaluateInsertion(ctx, movingAssignment, candidate, monthlyMap) {
  const simulatedAssignments = [...candidate.existingAssignments, movingAssignment];
  const orderedRows = __overflowGetOrderedRowsFromAssignments(ctx, simulatedAssignments);

  const routeKm = Number(calculateRouteDistanceGlobal(orderedRows) || 0);
  const timeSummary = getRowsTravelTimeSummary(orderedRows);
  const routeMinutes = Number(timeSummary?.totalMinutes || 0);

  const movingArea = __overflowGetAreaFromAssignment(ctx, movingAssignment);
  const existingAreas = candidate.existingAssignments
    .map(row => __overflowGetAreaFromAssignment(ctx, row))
    .filter(Boolean);

  let spreadScore = 0;
  existingAreas.forEach(area => {
    spreadScore += Number(getAreaAffinityScore(area, movingArea) || 0) * 2.0;
    spreadScore += Math.max(0, Number(getDirectionAffinityScore(area, movingArea) || 0)) * 1.8;
  });

  const vehicle = candidate.vehicle;
  spreadScore += Number(getVehicleAreaMatchScore(vehicle, movingArea) || 0) * 2.3;
  spreadScore += Number(getStrictHomeCompatibilityScore(movingArea, vehicle?.home_area || "") || 0) * 0.7;
  spreadScore += Math.max(0, Number(getDirectionAffinityScore(movingArea, vehicle?.home_area || "") || 0)) * 0.4;

  const month = monthlyMap?.get?.(Number(candidate.vehicleId)) || { totalDistance: 0, avgDistance: 0 };
  const loadPenalty = Number(month.totalDistance || 0) * 0.04 + Number(month.avgDistance || 0) * 0.35;

  let lastTripPenalty = 0;
  if (typeof isDriverLastTripChecked === "function" && isDriverLastTripChecked(candidate.vehicleId)) {
    lastTripPenalty = 18;
  }

  const score =
    spreadScore
    - routeMinutes * 1.0
    - routeKm * 0.7
    - loadPenalty
    - lastTripPenalty;

  return {
    vehicleId: candidate.vehicleId,
    vehicle,
    score,
    routeKm,
    routeMinutes
  };
}

function __overflowChooseBestTarget(ctx, movingAssignment, bucket, vehicles, monthlyMap) {
  const candidates = __overflowGetTargetCandidates(ctx, movingAssignment, bucket, vehicles);
  if (!candidates.length) return null;

  let best = null;

  candidates.forEach(candidate => {
    const evaluated = __overflowEvaluateInsertion(ctx, movingAssignment, candidate, monthlyMap);
    if (!best || evaluated.score > best.score) {
      best = evaluated;
    }
  });

  return best;
}

function __overflowRebuildStopOrderForVehicleHour(assignments, ctx, vehicleId, hour) {
  const targetAssignments = assignments.filter(a =>
    Number(a?.vehicle_id || 0) === Number(vehicleId) &&
    Number(a?.actual_hour || 0) === Number(hour)
  );

  if (!targetAssignments.length) return;

  const orderedRows = __overflowGetOrderedRowsFromAssignments(ctx, targetAssignments);
  const orderMap = new Map(orderedRows.map((row, index) => [Number(row.id), index + 1]));

  targetAssignments.forEach(assignment => {
    assignment.stop_order = Number(orderMap.get(Number(assignment.item_id)) || 1);
    const item = __overflowGetAssignmentItem(ctx, assignment);
    assignment.distance_km = Number(item?.distance_km || item?.casts?.distance_km || assignment?.distance_km || 0);
    if (!assignment.driver_name) {
      const vehicle = ctx.vehicleMap.get(Number(vehicleId));
      assignment.driver_name = vehicle?.driver_name || "";
    }
  });
}

function __overflowApplyMove(assignments, movingAssignment, bucket, targetVehicleId, items, vehicles) {
  const sourceIndex = assignments.findIndex(a =>
    Number(a?.item_id || 0) === Number(movingAssignment?.item_id || 0) &&
    Number(a?.vehicle_id || 0) === Number(bucket.vehicleId) &&
    Number(a?.actual_hour || 0) === Number(bucket.hour)
  );

  if (sourceIndex < 0) return assignments;

  const targetVehicle = (Array.isArray(vehicles) ? vehicles : []).find(v => Number(v?.id || 0) === Number(targetVehicleId));
  assignments[sourceIndex] = {
    ...assignments[sourceIndex],
    vehicle_id: Number(targetVehicleId),
    driver_name: targetVehicle?.driver_name || assignments[sourceIndex]?.driver_name || ""
  };

  let ctx = __overflowBuildContext(assignments, items, vehicles);
  __overflowRebuildStopOrderForVehicleHour(assignments, ctx, bucket.vehicleId, bucket.hour);

  ctx = __overflowBuildContext(assignments, items, vehicles);
  __overflowRebuildStopOrderForVehicleHour(assignments, ctx, targetVehicleId, bucket.hour);

  return assignments;
}

function __overflowFinalizeAssignments(assignments, items, vehicles) {
  const working = [...assignments];
  const touched = new Set(working.map(a => __overflowGetVehicleHourKey(a?.vehicle_id, a?.actual_hour)));

  touched.forEach(key => {
    const [vehicleId, hour] = String(key).split("__").map(Number);
    const ctx = __overflowBuildContext(working, items, vehicles);
    __overflowRebuildStopOrderForVehicleHour(working, ctx, vehicleId, hour);
  });

  return [...working].sort((a, b) => {
    const ah = Number(a?.actual_hour || 0);
    const bh = Number(b?.actual_hour || 0);
    if (ah !== bh) return ah - bh;

    const av = Number(a?.vehicle_id || 0);
    const bv = Number(b?.vehicle_id || 0);
    if (av !== bv) return av - bv;

    const ao = Number(a?.stop_order || 0);
    const bo = Number(b?.stop_order || 0);
    if (ao !== bo) return ao - bo;

    return Number(a?.item_id || 0) - Number(b?.item_id || 0);
  });
}

function resolveCapacityOverflowLocally(assignments, items, vehicles, monthlyMap) {
  if (!Array.isArray(assignments) || !assignments.length) return Array.isArray(assignments) ? assignments : [];

  let working = assignments.map(a => ({ ...a }));
  let ctx = __overflowBuildContext(working, items, vehicles);
  const buckets = __overflowFindBuckets(ctx);

  if (!buckets.length) return working;

  buckets.forEach(bucket => {
    ctx = __overflowBuildContext(working, items, vehicles);
    const sourceRows = ctx.vehicleHourRowsMap.get(bucket.key) || [];
    const overflowRows = __overflowSelectRows(sourceRows, bucket.overCount, ctx);

    overflowRows.forEach(movingAssignment => {
      ctx = __overflowBuildContext(working, items, vehicles);
      const best = __overflowChooseBestTarget(ctx, movingAssignment, bucket, vehicles, monthlyMap);
      if (!best) return;

      working = __overflowApplyMove(
        working,
        movingAssignment,
        bucket,
        best.vehicleId,
        items,
        vehicles
      );
    });
  });

  working = __overflowFinalizeAssignments(working, items, vehicles);
  return working;
}

/* ===== THEMIS overflow local resolver end ===== */

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
  let assignments = [];

  try {
    assignments = optimizeAssignments(activeItems, selectedVehicles, monthlyMap);
    if (!Array.isArray(assignments)) assignments = [];

    assignments = resolveCapacityOverflowLocally(
      assignments,
      activeItems,
      selectedVehicles,
      monthlyMap
    );
  } catch (error) {
    console.error("runAutoDispatch main pipeline error:", error);
    assignments = [];
  }

  if ((!Array.isArray(assignments) || !assignments.length) && typeof __buildEmergencyAssignments === "function") {
    try {
      assignments = __buildEmergencyAssignments(activeItems, selectedVehicles);
      assignments = resolveCapacityOverflowLocally(
        assignments,
        activeItems,
        selectedVehicles,
        monthlyMap
      );
    } catch (fallbackError) {
      console.error("runAutoDispatch emergency fallback error:", fallbackError);
      assignments = [];
    }
  }

  if (!Array.isArray(assignments)) assignments = [];

  const finalAssignedItemIds = new Set(assignments.map(a => Number(a?.item_id || 0)).filter(Boolean));
  const finalUnassignedCount = activeItems.filter(item => !finalAssignedItemIds.has(Number(item?.id || 0))).length;

  if (!assignments.length || finalUnassignedCount > 0) {
    console.error("auto dispatch incomplete", {
      totalItems: activeItems.length,
      assignedCount: finalAssignedItemIds.size,
      unassignedCount: finalUnassignedCount
    });
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
          created_by: getCurrentUserIdSafe()
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
          created_by: getCurrentUserIdSafe()
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
  const safeDispatchId = Number.isFinite(Number(dispatchId)) && Number(dispatchId) > 0 ? Number(dispatchId) : null;
  const safeItemId = Number.isFinite(Number(itemId)) && Number(itemId) > 0 ? Number(itemId) : null;

  const payload = {
    dispatch_id: safeDispatchId,
    item_id: safeItemId,
    action,
    message,
    acted_by: getCurrentUserIdSafe()
  };

  let { error } = await supabaseClient
    .from("dispatch_history")
    .insert(payload);

  if (error && String(error?.message || "").includes("dispatch_history_item_id_fkey")) {
    ({ error } = await supabaseClient
      .from("dispatch_history")
      .insert({
        ...payload,
        item_id: null
      }));
  }

  if (error) console.error(error);
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
      if (!row.created_by) row.created_by = getCurrentUserIdSafe();

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
      if (!row.created_by) row.created_by = getCurrentUserIdSafe();

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
      if (!row.created_by) row.created_by = getCurrentUserIdSafe();

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
      if (!row.created_by) row.created_by = getCurrentUserIdSafe();

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
      if (!row.created_by) row.created_by = getCurrentUserIdSafe();

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
      if (!row.acted_by) row.acted_by = getCurrentUserIdSafe();

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
          created_by: getCurrentUserIdSafe()
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
          created_by: getCurrentUserIdSafe()
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
  syncMileageReportRange(dateStr, true);

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
  syncMileageReportRange(dateStr, true);

  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
  renderManualLastVehicleInfo();
}

async function syncDateAndReloadFromActualDate() {
  const dateStr = els.actualDate?.value || todayStr();
  if (els.dispatchDate) els.dispatchDate.value = dateStr;
  if (els.planDate) els.planDate.value = dateStr;
  syncMileageReportRange(dateStr, true);

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
  if (els.optimizeBtn) els.optimizeBtn.addEventListener("click", () => (typeof window.runAutoDispatchV1 === "function" ? window.runAutoDispatchV1() : runAutoDispatch()));
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

let postDispatchEventsBound = false;
let dashboardInitialized = false;
let mileageSyncListenersBound = false;

function bindMileageReportSyncListeners() {
  if (mileageSyncListenersBound) return;
  mileageSyncListenersBound = true;

  const syncMileageReport = () => {
    syncMileageReportRange(els.dispatchDate?.value || todayStr(), true);
  };

  window.addEventListener("load", syncMileageReport, { once: true });
  window.addEventListener("pageshow", syncMileageReport, { once: true });
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      syncMileageReport();
    }
  });
}

function bindPostDispatchEvents() {
  if (postDispatchEventsBound) return;
  postDispatchEventsBound = true;

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
  els.previewMileageReportBtn?.addEventListener("click", async () => {
    await previewDriverMileageReport();
    await refreshHomeMonthlyVehicleList();
    renderVehiclesTable();
  });
  els.exportMileageReportBtn?.addEventListener("click", exportDriverMileageReportXlsx);
  els.mileageReportStartDate?.addEventListener("change", async () => {
    await refreshHomeMonthlyVehicleList();
    renderVehiclesTable();
  });
  els.mileageReportEndDate?.addEventListener("change", async () => {
    await refreshHomeMonthlyVehicleList();
    renderVehiclesTable();
  });

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

document.addEventListener("DOMContentLoaded", async () => {
  if (dashboardInitialized) {
    console.warn("[dashboard] initialization skipped: already initialized");
    return;
  }
  dashboardInitialized = true;

  try {
    console.log("SUPABASE_URL:", SUPABASE_URL);

    const ok = await ensureAuth();
    if (!ok) return;

    ensureCastTravelMinutesUi();
    initializeMileageReportDefaultDates();
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
    forceResetMileageReportInputs(today);

    syncScheduleRendererDeps();
    await loadHomeAndAll();
    syncMileageReportRange(today, true);
    window.requestAnimationFrame(() => syncMileageReportRange(els.dispatchDate?.value || todayStr(), true));
    window.setTimeout(() => syncMileageReportRange(els.dispatchDate?.value || todayStr(), true), 0);
    window.setTimeout(() => syncMileageReportRange(els.dispatchDate?.value || todayStr(), true), 120);
    bindMileageReportSyncListeners();
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

  const departDelayMinutes = (typeof getExpectedDepartureDelayMinutes === "function" ? getExpectedDepartureDelayMinutes(baseHour) : 0);
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
  const maxDetourKm = 10;
  try {
    const candidatePoint = typeof getThemisV54RowPoint === "function" ? getThemisV54RowPoint(item) : null;
    const siblingAssignmentsForDetour = working.filter(a =>
      Number(a.vehicle_id) === Number(vehicle?.id) &&
      Number(a.actual_hour ?? 0) === Number(currentHour ?? 0) &&
      Number(a.item_id) !== Number(item?.id)
    );
    const siblingItemsForDetour = siblingAssignmentsForDetour
      .map(a => itemMap.get(Number(a.item_id)))
      .filter(Boolean);
    const siblingPoints = siblingItemsForDetour
      .map(row => (typeof getThemisV54RowPoint === "function" ? getThemisV54RowPoint(row) : null))
      .filter(Boolean);

    if (candidatePoint && siblingPoints.length) {
      const avgLat = siblingPoints.reduce((sum, pt) => sum + Number(pt.lat || 0), 0) / siblingPoints.length;
      const avgLng = siblingPoints.reduce((sum, pt) => sum + Number(pt.lng || 0), 0) / siblingPoints.length;
      const detourKm = typeof estimateRoadKmBetweenPoints === "function"
        ? Number(estimateRoadKmBetweenPoints(avgLat, avgLng, candidatePoint.lat, candidatePoint.lng) || 0)
        : Number(Math.sqrt(((candidatePoint.lat - avgLat) ** 2) + ((candidatePoint.lng - avgLng) ** 2)) * 111 || 0);

      if (detourKm > maxDetourKm) {
        return -Infinity;
      }
    }
  } catch (bundleGuardError) {
    console.warn("bundle protection skipped:", bundleGuardError);
  }

    const area = getThemisV54RowArea(item);
    const compat =
      getStrictHomeCompatibilityScore(area, vehicle?.home_area || "") * 1.7 +
      Math.max(0, getDirectionAffinityScore(area, vehicle?.home_area || "")) * 0.95 +
      getAreaAffinityScore(area, vehicle?.home_area || "") * 0.85 +
      getVehicleAreaMatchScore(vehicle, area) * 0.55;

    const projectedAvg = getThemisV54VehicleProjectedAvg(vehicle?.id, monthlyMap, assignedDistance, assignedCount);
    const currentHourLoad = Number(hourLoads.get(`${Number(vehicle?.id)}__${Number(currentHour ?? 0)}`) || 0);

    const siblingAssignments = working.filter(a => Number(a.vehicle_id) === Number(vehicle?.id) && Number(a.actual_hour ?? 0) === Number(currentHour ?? 0) && Number(a.item_id) !== Number(item?.id));
    const siblingItems = siblingAssignments.map(a => itemMap.get(Number(a.item_id))).filter(Boolean);
    const siblingAreas = siblingItems.map(s => getThemisV54RowArea(s)).filter(Boolean);
    const siblingAnchor = [...siblingItems].sort((a, b) => Number(b?.distance_km || 0) - Number(a?.distance_km || 0))[0] || null;
    const siblingAnchorArea = siblingAnchor ? getThemisV54RowArea(siblingAnchor) : "";
    const siblingAnchorDistance = Number(siblingAnchor?.distance_km || 0);
    const itemDistance = Number(item?.distance_km || 0);

    let bundlePenalty = 0;
    siblingAreas.forEach(existingArea => {
      const direction = Number(getDirectionAffinityScore(area, existingArea) || 0);
      const affinity = Number(getAreaAffinityScore(area, existingArea) || 0);
      if (direction <= -38 || isHardReverseMixForRoute(area, existingArea)) {
        bundlePenalty += 1600;
      } else {
        bundlePenalty += Math.max(0, 36 - direction) * 8;
        bundlePenalty += Math.max(0, 55 - affinity) * 1.5;
      }
    });

    if (siblingAnchorArea) {
      const anchorDirection = Number(getDirectionAffinityScore(area, siblingAnchorArea) || 0);
      if (anchorDirection <= -38 || isHardReverseMixForRoute(area, siblingAnchorArea)) {
        bundlePenalty += 2200;
      } else {
        bundlePenalty += Math.max(0, 42 - anchorDirection) * 10;
        bundlePenalty += Math.abs(itemDistance - siblingAnchorDistance) * 0.8;
      }
    }

    return compat - projectedAvg * 1.35 - currentHourLoad * 18 - bundlePenalty;
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
    (vehicles || []).forEach(v => map.set(Number(v.id), { count:0, areas:[], rows:[], vehicle:v }));
    (assignments || []).filter(r => Number(r?.actual_hour ?? 0) === Number(hour)).forEach(r => {
      const id = Number(r.vehicle_id || 0);
      if (!map.has(id)) return;
      const state = map.get(id);
      state.count += 1;
      state.rows.push(r);
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
    const farthestBundleRow = [...(bundleRows || [])].sort((a, b) => Number(b?.distance_km || 0) - Number(a?.distance_km || 0))[0] || bundleRows[0];
    const rowArea = normalizeAreaLabel(farthestBundleRow?.destination_area || farthestBundleRow?.cluster_area || farthestBundleRow?.planned_area || farthestBundleRow?.casts?.area || '');
    const rowDistance = Number(farthestBundleRow?.distance_km || 0);
    let best = null;
    for (const [vehicleId, state] of vehicleStates.entries()) {
      const vehicle = state.vehicle || (vehicles || []).find(v => Number(v.id) === Number(vehicleId));
      if (!vehicle) continue;
      const capacity = Number(vehicle.seat_capacity || 4);
      if (state.count + bundleRows.length > capacity) continue;
      if (!bundleRows.every(r => v551CanMoveIntoVehicle(r, state, vehicle))) continue;

      let score = 0;
      const areas = state.areas || [];
      const existingRows = Array.isArray(state.rows) ? state.rows : [];
      const existingAnchorRow = [...existingRows].sort((a, b) => Number(b?.distance_km || 0) - Number(a?.distance_km || 0))[0] || null;
      const existingAnchorArea = normalizeAreaLabel(existingAnchorRow?.destination_area || existingAnchorRow?.cluster_area || existingAnchorRow?.planned_area || existingAnchorRow?.casts?.area || '');
      const existingAnchorDistance = Number(existingAnchorRow?.distance_km || 0);

      if (areas.length) {
        for (const area of areas) {
          const direction = Number(getDirectionAffinityScore(rowArea, area) || 0);
          const affinity = Number(getAreaAffinityScore(rowArea, area) || 0);
          if (v551IsStrongReverse(rowArea, area)) {
            score -= 1200;
          } else {
            score += direction * 1.8;
            score += affinity * 0.7;
            if (v551IsFriendlyDirection(rowArea, area)) score += 140;
          }
        }
      } else {
        score += 20;
      }

      if (existingAnchorArea) {
        const anchorDir = Number(getDirectionAffinityScore(rowArea, existingAnchorArea) || 0);
        const anchorAffinity = Number(getAreaAffinityScore(rowArea, existingAnchorArea) || 0);
        if (v551IsStrongReverse(rowArea, existingAnchorArea)) {
          score -= 1800;
        } else {
          score += anchorDir * 4.2;
          score += anchorAffinity * 1.2;
        }
        score -= Math.abs(existingAnchorDistance - rowDistance) * 0.9;
      }

      const monthly = monthlyMap.get(Number(vehicleId)) || {};
      const avg = Number(monthly.averageDistance || monthly.avgDistance || 0);
      score -= avg * 0.08;
      score -= Number(state.count || 0) * 5;

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

  const _THEMIS_V553_BASE_applyManualLastVehicleToAssignments = (typeof applyManualLastVehicleToAssignments === "function" ? applyManualLastVehicleToAssignments : function(assignments){ return Array.isArray(assignments) ? assignments : []; });
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


// ===== THEMIS field trial compatibility shim =====

function __themisNormalizeStrictAreaLabel(area) {
  return normalizeAreaLabel(area || "");
}

function __themisGetStrictAreaKeyFromItem(item) {
  const area = __themisNormalizeStrictAreaLabel(
    item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || ""
  );
  const displayGroup = typeof getAreaDisplayGroup === "function"
    ? (getAreaDisplayGroup(area) || area)
    : area;
  return __themisNormalizeStrictAreaLabel(displayGroup);
}

function __themisGetStrictAreaKeyFromAssignment(assignment, itemMap) {
  const item = itemMap.get(Number(assignment?.item_id || 0));
  return __themisGetStrictAreaKeyFromItem(item || assignment || {});
}

function __themisHasEnoughVehiclesForStrictAreaLayer(items, vehicles) {
  const safeItems = Array.isArray(items) ? items.filter(Boolean) : [];
  const safeVehicles = Array.isArray(vehicles) ? vehicles.filter(Boolean) : [];
  if (!safeItems.length || !safeVehicles.length) return false;

  const hourAreaMap = new Map();
  safeItems.forEach(item => {
    const hour = Number(item?.actual_hour ?? item?.plan_hour ?? 0);
    const areaKey = __themisGetStrictAreaKeyFromItem(item);
    if (!areaKey || areaKey === "無し") return;
    if (!hourAreaMap.has(hour)) hourAreaMap.set(hour, new Set());
    hourAreaMap.get(hour).add(areaKey);
  });

  if (!hourAreaMap.size) return false;
  for (const areaSet of hourAreaMap.values()) {
    if (areaSet.size > safeVehicles.length) return false;
  }
  return true;
}

function __themisGetStrictAreaVehicleScore(vehicle, areaKey, monthlyMap) {
  const month = monthlyMap?.get?.(Number(vehicle?.id || 0)) || { totalDistance: 0, avgDistance: 0, workedDays: 0 };
  let score = 0;
  if (typeof getVehicleAreaMatchScore === "function") {
    score += Number(getVehicleAreaMatchScore(vehicle, areaKey) || 0) * 3.2;
  }
  if (typeof getStrictHomeCompatibilityScore === "function") {
    score += Number(getStrictHomeCompatibilityScore(areaKey, vehicle?.home_area || "") || 0) * 2.6;
  }
  if (typeof getDirectionAffinityScore === "function") {
    score += Math.max(0, Number(getDirectionAffinityScore(areaKey, vehicle?.home_area || "") || 0)) * 1.2;
    score += Math.max(0, Number(getDirectionAffinityScore(areaKey, vehicle?.vehicle_area || "") || 0)) * 0.9;
  }
  if (typeof getAreaAffinityScore === "function") {
    score += Number(getAreaAffinityScore(areaKey, vehicle?.home_area || "") || 0) * 0.9;
    score += Number(getAreaAffinityScore(areaKey, vehicle?.vehicle_area || "") || 0) * 0.7;
  }
  score -= Number(month.totalDistance || 0) * 0.02;
  score -= Number(month.avgDistance || 0) * 0.12;
  if (typeof isDriverLastTripChecked === "function" && isDriverLastTripChecked(Number(vehicle?.id || 0))) {
    score -= 12;
  }
  return score;
}

function __themisBuildAssignmentsPreserveStrictAreas(items, vehicles, monthlyMap) {
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
    const byArea = new Map();
    hourItems.forEach(item => {
      const areaKey = __themisGetStrictAreaKeyFromItem(item);
      if (!byArea.has(areaKey)) byArea.set(areaKey, []);
      byArea.get(areaKey).push(item);
    });

    const areaEntries = [...byArea.entries()].map(([areaKey, rows]) => ({
      areaKey,
      rows: rows.slice(),
      anchor: rows.slice().sort((a, b) => Number(b?.distance_km || 0) - Number(a?.distance_km || 0) || Number(a?.id || 0) - Number(b?.id || 0))[0] || null,
      count: rows.length
    })).sort((a, b) => {
      const ad = Number(a.anchor?.distance_km || 0);
      const bd = Number(b.anchor?.distance_km || 0);
      if (bd !== ad) return bd - ad;
      if (b.count !== a.count) return b.count - a.count;
      return String(a.areaKey || "").localeCompare(String(b.areaKey || ""), "ja");
    });

    const usedVehicleIds = new Set();
    const dedicatedMap = new Map();

    areaEntries.forEach(entry => {
      let bestVehicle = null;
      let bestScore = -Infinity;
      safeVehicles.forEach(vehicle => {
        const vehicleId = Number(vehicle?.id || 0);
        if (!vehicleId || usedVehicleIds.has(vehicleId)) return;
        const score = __themisGetStrictAreaVehicleScore(vehicle, entry.areaKey, monthlyMap);
        if (score > bestScore) {
          bestVehicle = vehicle;
          bestScore = score;
        }
      });
      if (bestVehicle) {
        const vehicleId = Number(bestVehicle.id || 0);
        usedVehicleIds.add(vehicleId);
        dedicatedMap.set(entry.areaKey, bestVehicle);
      }
    });

    if (dedicatedMap.size !== areaEntries.length) {
      return [];
    }

    areaEntries.forEach(entry => {
      const vehicle = dedicatedMap.get(entry.areaKey);
      const vehicleId = Number(vehicle?.id || 0);
      const orderedRows = typeof sortItemsByNearestRoute === "function"
        ? sortItemsByNearestRoute(entry.rows)
        : entry.rows.slice().sort((a, b) => Number(a?.distance_km || 0) - Number(b?.distance_km || 0) || Number(a?.id || 0) - Number(b?.id || 0));
      orderedRows.forEach((row, index) => {
        assignments.push({
          item_id: Number(row?.id || 0),
          actual_hour: hour,
          vehicle_id: vehicleId,
          driver_name: vehicle?.driver_name || "",
          distance_km: Number(row?.distance_km || row?.casts?.distance_km || 0),
          stop_order: index + 1,
          strict_area_key: entry.areaKey
        });
      });
    });
  }

  return assignments;
}

function optimizeAssignments(items, vehicles, monthlyMap, options = {}) {
  if (!window.DispatchCore || typeof window.DispatchCore.optimizeAssignments !== 'function') {
    console.error('DispatchCore.optimizeAssignments is not available');
    return [];
  }
  const safeItems = Array.isArray(items) ? items : [];
  const safeVehicles = Array.isArray(vehicles)
    ? vehicles
    : (vehicles && typeof vehicles === 'object')
      ? Object.values(vehicles)
      : [];
  const safeMonthlyMap = monthlyMap instanceof Map ? monthlyMap : new Map();

  const useStrictAreaLayer = !options?.preAssignedAssignments && __themisHasEnoughVehiclesForStrictAreaLayer(safeItems, safeVehicles);
  if (useStrictAreaLayer) {
    const strictAssignments = __themisBuildAssignmentsPreserveStrictAreas(safeItems, safeVehicles, safeMonthlyMap);
    if (Array.isArray(strictAssignments) && strictAssignments.length === safeItems.length) {
      return strictAssignments;
    }
  }

  return window.DispatchCore.optimizeAssignments(safeItems, safeVehicles, safeMonthlyMap, options || {});
}

function callDispatchCoreSafe(items, vehicles, monthlyMap, options = {}) {
  return optimizeAssignments(items, vehicles, monthlyMap, options);
}

function runSimulationDispatchPreview() {
  if (typeof runAutoDispatch === 'function') {
    return runAutoDispatch();
  }
  return [];
}


/* ===== THEMIS monthly UI sync hotfix v4 start ===== */
(function(){
  let __themisMonthlyUiRows = [];

  function __themisMonthRange(dateStr) {
    const base = String(dateStr || (els?.dispatchDate?.value || todayStr?.() || new Date().toISOString().slice(0,10)));
    const monthKey = typeof getMonthKey === 'function' ? getMonthKey(base) : base.slice(0, 7);
    return { startDate: `${monthKey}-01`, endDate: base, baseDate: base };
  }

  function __themisResolveVehicleId(row) {
    const directId = Number(row?.vehicle_id || row?.vehicles?.id || 0);
    if (directId > 0) return directId;
    const reportDriver = String(row?.driver_name || row?.vehicles?.driver_name || '').trim();
    const reportPlate = String(row?.plate_number || row?.vehicles?.plate_number || '').trim();
    const matched = (Array.isArray(allVehiclesCache) ? allVehiclesCache : []).find(vehicle => {
      const vehicleDriver = String(vehicle?.driver_name || '').trim();
      const vehiclePlate = String(vehicle?.plate_number || '').trim();
      return (reportDriver && vehicleDriver && reportDriver === vehicleDriver) || (reportPlate && vehiclePlate && reportPlate === vehiclePlate);
    });
    return Number(matched?.id || 0);
  }

  function __themisNormalizeMonthlyRows(rows) {
    return (Array.isArray(rows) ? rows : []).map(row => ({
      ...row,
      vehicle_id: __themisResolveVehicleId(row),
      report_date: row?.report_date || '',
      driver_name: row?.driver_name || row?.vehicles?.driver_name || '',
      plate_number: row?.plate_number || row?.vehicles?.plate_number || '',
      distance_km: Number(row?.distance_km || 0)
    }));
  }

  async function __themisFetchMonthlyRows(baseDate) {
    const range = __themisMonthRange(baseDate);
    if (typeof fetchDriverMileageRows === 'function') {
      try {
        const rows = await fetchDriverMileageRows(range.startDate, range.endDate);
        return __themisNormalizeMonthlyRows(rows);
      } catch (error) {
        console.error('fetchDriverMileageRows monthly hotfix error:', error);
      }
    }

    if (supabaseClient?.from) {
      try {
        const { data, error } = await supabaseClient
          .from('vehicle_daily_reports')
          .select('vehicle_id, driver_name, distance_km, report_date, note, vehicles(id, plate_number, driver_name)')
          .gte('report_date', range.startDate)
          .lte('report_date', range.endDate)
          .order('report_date', { ascending: true });
        if (error) throw error;
        return __themisNormalizeMonthlyRows(data);
      } catch (error) {
        console.error('supabase monthly hotfix fetch error:', error);
      }
    }

    return __themisNormalizeMonthlyRows(currentDailyReportsCache);
  }

  function __themisStatsMap(rows, baseDate) {
    return getUnifiedMonthlyUiStatsMap(__themisNormalizeMonthlyRows(rows), baseDate || (els?.dispatchDate?.value || todayStr()));
  }

  async function __themisRefreshMonthlyUi(baseDate) {
    const rows = await __themisFetchMonthlyRows(baseDate);
    __themisMonthlyUiRows = rows;
    try { currentDailyReportsCache = rows; } catch(e) {}
    if (typeof renderHomeMonthlyVehicleList === 'function') renderHomeMonthlyVehicleList(rows);
    if (typeof renderVehiclesTable === 'function') renderVehiclesTable();
    if (typeof renderDailyVehicleChecklist === 'function') renderDailyVehicleChecklist();
  }

  renderHomeMonthlyVehicleList = function(reportRows = currentDailyReportsCache) {
    if (!els?.homeMonthlyVehicleList) return;
    const rows = (Array.isArray(__themisMonthlyUiRows) && __themisMonthlyUiRows.length) ? __themisMonthlyUiRows : __themisNormalizeMonthlyRows(reportRows);
    const statsMap = __themisStatsMap(rows, els?.dispatchDate?.value || todayStr());
    els.homeMonthlyVehicleList.innerHTML = '';
    if (!allVehiclesCache.length) {
      els.homeMonthlyVehicleList.innerHTML = `<div class="chip">車両なし</div>`;
      return;
    }
    getSortedVehiclesForDisplay().forEach(vehicle => {
      const stats = statsMap.get(Number(vehicle.id)) || { totalDistance: 0, workedDays: 0, avgDistance: 0 };
      const row = document.createElement('div');
      row.className = 'home-monthly-item';
      row.innerHTML = `
        <span class="chip">${escapeHtml(vehicle.driver_name || vehicle.plate_number || '-')}</span>
        <span class="chip">${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || '-'))}</span>
        <span class="chip">帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || '-'))}</span>
        <span class="chip">月間:${Number(stats.totalDistance || 0).toFixed(1)}km</span>
        <span class="chip">出勤:${Number(stats.workedDays || 0)}日</span>
        <span class="chip">平均:${Number(stats.avgDistance || 0).toFixed(1)}km</span>`;
      els.homeMonthlyVehicleList.appendChild(row);
    });
  };

  renderVehiclesTable = function() {
    if (!els?.vehiclesTableBody) return;
    const rows = (Array.isArray(__themisMonthlyUiRows) && __themisMonthlyUiRows.length) ? __themisMonthlyUiRows : __themisNormalizeMonthlyRows(currentDailyReportsCache);
    const statsMap = __themisStatsMap(rows, els?.dispatchDate?.value || todayStr());
    els.vehiclesTableBody.innerHTML = '';
    if (!allVehiclesCache.length) {
      els.vehiclesTableBody.innerHTML = `<tr><td colspan="9" class="muted">車両がありません</td></tr>`;
      return;
    }
    getSortedVehiclesForDisplay().forEach(vehicle => {
      const stats = statsMap.get(Number(vehicle.id)) || { totalDistance: 0, workedDays: 0, avgDistance: 0 };
      const tr = document.createElement('tr');
      tr.innerHTML = `
      <tr>
      <td>${escapeHtml(vehicle.driver_name || '-')}</td>
      <td>${escapeHtml(vehicle.plate_number || '-')}</td>
      <td>${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || '-'))}</td>
      <td>${escapeHtml(normalizeAreaLabel(vehicle.home_area || '-'))}</td>
      <td>${vehicle.seat_capacity ?? '-'}</td>
      <td>${Number(stats.totalDistance || 0).toFixed(1)}</td>
      <td>${Number(stats.workedDays || 0)}</td>
      <td>${Number(stats.avgDistance || 0).toFixed(1)}</td>
      <td class="actions-cell">
        <button class="btn ghost vehicle-edit-btn" data-id="${vehicle.id}">編集</button>
        <button class="btn danger vehicle-delete-btn" data-id="${vehicle.id}">削除</button>
      </td>`;
      els.vehiclesTableBody.appendChild(tr);
    });
    els.vehiclesTableBody.querySelectorAll('.vehicle-edit-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const vehicle = allVehiclesCache.find(x => Number(x.id) === Number(btn.dataset.id));
        if (vehicle) fillVehicleForm(vehicle);
      });
    });
    els.vehiclesTableBody.querySelectorAll('.vehicle-delete-btn').forEach(btn => {
      btn.addEventListener('click', async () => deleteVehicle(Number(btn.dataset.id)));
    });
  };

  renderDailyVehicleChecklist = function() {
    if (!els?.dailyVehicleChecklist) return;
    els.dailyVehicleChecklist.innerHTML = '';
    if (!allVehiclesCache.length) {
      els.dailyVehicleChecklist.innerHTML = `<div class="muted">車両がありません</div>`;
      return;
    }
    const rows = (Array.isArray(__themisMonthlyUiRows) && __themisMonthlyUiRows.length) ? __themisMonthlyUiRows : __themisNormalizeMonthlyRows(currentDailyReportsCache);
    const monthlyStatsMap = __themisStatsMap(rows, els?.dispatchDate?.value || todayStr());
    const header = document.createElement('div');
    header.className = 'vehicle-check-header';
    header.innerHTML = `<div class="vehicle-check-header-info"></div><div class="vehicle-check-header-col">可能車両</div><div class="vehicle-check-header-col">ラスト便</div>`;
    els.dailyVehicleChecklist.appendChild(header);
    getSortedVehiclesForDisplay().forEach(vehicle => {
      const stats = monthlyStatsMap.get(Number(vehicle.id)) || { totalDistance: 0, workedDays: 0, avgDistance: 0 };
      const avgDistanceText = `${Number(stats.avgDistance || 0).toFixed(1)}km`;
      const row = document.createElement('div');
      row.className = 'vehicle-check-item';
      row.innerHTML = `
        <div class="vehicle-check-info">
          <div class="vehicle-check-name">${escapeHtml(vehicle.driver_name || '-')}</div>
          <div class="vehicle-check-car">車両 ${escapeHtml(vehicle.plate_number || '-')}</div>
          <div class="vehicle-check-meta">担当 ${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || '-'))} / 帰宅 ${escapeHtml(normalizeAreaLabel(vehicle.home_area || '-'))} / 定員 ${vehicle.seat_capacity ?? '-'} / 1日平均距離 ${avgDistanceText}</div>
        </div>
        <label class="vehicle-check-toggle vehicle-check-toggle-work">
          <input class="vehicle-check-input" type="checkbox" data-id="${vehicle.id}" ${activeVehicleIdsForToday.has(Number(vehicle.id)) ? 'checked' : ''} />
          <span>可能車両</span>
        </label>
        <label class="vehicle-check-toggle vehicle-check-toggle-last">
          <input class="driver-last-trip-input" type="checkbox" data-id="${vehicle.id}" ${isDriverLastTripChecked(vehicle.id) ? 'checked' : ''} />
          <span>ラスト便</span>
        </label>`;
      els.dailyVehicleChecklist.appendChild(row);
    });
    renderOperationAndSimulationUI();
    els.dailyVehicleChecklist.querySelectorAll('.vehicle-check-input').forEach(input => {
      input.addEventListener('change', () => {
        const id = Number(input.dataset.id);
        if (input.checked) activeVehicleIdsForToday.add(id); else activeVehicleIdsForToday.delete(id);
        renderDailyMileageInputs();
        renderDailyDispatchResult();
        renderDailyVehicleChecklist();
      });
    });
    els.dailyVehicleChecklist.querySelectorAll('.driver-last-trip-input').forEach(input => {
      input.addEventListener('change', () => {
        const id = Number(input.dataset.id);
        setDriverLastTripChecked(id, input.checked);
        if (input.checked && !activeVehicleIdsForToday.has(id)) activeVehicleIdsForToday.add(id);
        renderDailyMileageInputs();
        renderDailyDispatchResult();
        renderDailyVehicleChecklist();
      });
    });
  };

  const __themisRefreshNames = ['saveDailyMileageReports','confirmDailyToMonthly','resetMonthlySummary','syncDateAndReloadFromDispatchDate','syncDateAndReloadFromPlanDate','syncDateAndReloadFromActualDate','loadHomeAndAll','previewDriverMileageReport'];
  __themisRefreshNames.forEach(name => {
    const orig = globalThis[name];
    if (typeof orig !== 'function') return;
    globalThis[name] = async function(...args) {
      const result = await orig.apply(this, args);
      try { await __themisRefreshMonthlyUi(els?.dispatchDate?.value || todayStr()); } catch (e) { console.error(`${name} monthly refresh error:`, e); }
      return result;
    };
  });

  window.addEventListener('load', () => {
    setTimeout(() => { __themisRefreshMonthlyUi(els?.dispatchDate?.value || todayStr()).catch(err => console.error('monthly refresh on load error:', err)); }, 300);
  });

  window.__themisRefreshMonthlyUi = __themisRefreshMonthlyUi;
})();
/* ===== THEMIS monthly UI sync hotfix v4 end ===== */


function getVehicleAreaMatchScore(vehicle, area) {
  const targetRaw = typeof normalizeAreaLabel === "function"
    ? normalizeAreaLabel(area || "")
    : String(area || "").trim();
  if (!targetRaw || targetRaw === "無し") return 0;

  const targetCanonical = typeof getCanonicalArea === "function"
    ? (getCanonicalArea(targetRaw) || targetRaw)
    : targetRaw;
  const targetGroup = (typeof THEMIS_DISPLAY_GROUPS !== "undefined" && THEMIS_DISPLAY_GROUPS && THEMIS_DISPLAY_GROUPS.has(targetRaw))
    ? targetRaw
    : (typeof getAreaDisplayGroup === "function" ? getAreaDisplayGroup(targetRaw) : targetCanonical);

  const vehicleAreaRaw = typeof normalizeAreaLabel === "function"
    ? normalizeAreaLabel(vehicle?.vehicle_area || "")
    : String(vehicle?.vehicle_area || "").trim();
  const homeAreaRaw = typeof normalizeAreaLabel === "function"
    ? normalizeAreaLabel(vehicle?.home_area || "")
    : String(vehicle?.home_area || "").trim();

  const vehicleCanonical = typeof getCanonicalArea === "function"
    ? (getCanonicalArea(vehicleAreaRaw) || vehicleAreaRaw)
    : vehicleAreaRaw;
  const homeCanonical = typeof getCanonicalArea === "function"
    ? (getCanonicalArea(homeAreaRaw) || homeAreaRaw)
    : homeAreaRaw;

  const vehicleGroup = vehicleAreaRaw
    ? ((typeof THEMIS_DISPLAY_GROUPS !== "undefined" && THEMIS_DISPLAY_GROUPS && THEMIS_DISPLAY_GROUPS.has(vehicleAreaRaw))
        ? vehicleAreaRaw
        : (typeof getAreaDisplayGroup === "function" ? getAreaDisplayGroup(vehicleAreaRaw) : vehicleCanonical))
    : vehicleCanonical;
  const homeGroup = homeAreaRaw
    ? ((typeof THEMIS_DISPLAY_GROUPS !== "undefined" && THEMIS_DISPLAY_GROUPS && THEMIS_DISPLAY_GROUPS.has(homeAreaRaw))
        ? homeAreaRaw
        : (typeof getAreaDisplayGroup === "function" ? getAreaDisplayGroup(homeAreaRaw) : homeCanonical))
    : homeCanonical;

  function calcScore(baseCanonical, baseGroup, weight, useHomeReverseGuard) {
    if (!baseCanonical && !baseGroup) return -999;

    let score = 0;
    if (baseCanonical && targetCanonical && baseCanonical === targetCanonical) {
      score = 100;
    } else if (baseGroup && targetGroup && baseGroup === targetGroup) {
      score = 88;
    } else {
      const affinity = typeof getAreaAffinityScore === "function"
        ? Number(getAreaAffinityScore(baseCanonical || baseGroup, targetCanonical) || 0)
        : 0;
      const direction = typeof getDirectionAffinityScore === "function"
        ? Number(getDirectionAffinityScore(baseCanonical || baseGroup, targetCanonical) || 0)
        : 0;

      score = affinity * 0.72 + direction * 0.34;
      if (direction <= -38) score -= 55;
      if (direction <= -95) score -= 95;
    }

    if (typeof isHardReverseMixForRoute === "function" && isHardReverseMixForRoute(baseCanonical || baseGroup, targetCanonical)) {
      score -= 130;
    }
    if (useHomeReverseGuard && typeof isHardReverseForHome === "function" && isHardReverseForHome(targetCanonical, baseCanonical || baseGroup)) {
      score -= 120;
    }

    return score * weight;
  }

  const vehicleScore = calcScore(vehicleCanonical, vehicleGroup, 1.0, false);
  const homeScore = calcScore(homeCanonical, homeGroup, 0.72, true);

  let best = Math.max(vehicleScore, homeScore, 0);

  if (vehicleGroup && homeGroup && vehicleGroup === homeGroup && vehicleGroup === targetGroup) {
    best += 8;
  }

  return Math.max(-160, Math.min(110, Math.round(best)));
}

