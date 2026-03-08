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

const els = {
  userEmail: document.getElementById("userEmail"),
  logoutBtn: document.getElementById("logoutBtn"),

  dispatchDate: document.getElementById("dispatchDate"),
  driverName: document.getElementById("driverName"),
  vehicleSelect: document.getElementById("vehicleSelect"),
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
  vehicleName: document.getElementById("vehicleName"),
  vehicleDriverName: document.getElementById("vehicleDriverName"),
  vehicleLineId: document.getElementById("vehicleLineId"),
  vehicleStatus: document.getElementById("vehicleStatus"),
  vehicleSeatCapacity: document.getElementById("vehicleSeatCapacity"),
  vehicleHomeArea: document.getElementById("vehicleHomeArea"),
  vehicleCapacityNote: document.getElementById("vehicleCapacityNote"),
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

function hasValue(v) {
  return v !== null && v !== undefined && v !== "";
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

    if (isValidLatLng(lat, lng)) {
      return { lat, lng };
    }
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
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function isFridayOrSaturday(dateStr) {
  const d = new Date(dateStr);
  const day = d.getDay();
  return day === 5 || day === 6;
}

function getPlanHourOptions(dateStr) {
  const base = [0, 1, 2, 3];
  return isFridayOrSaturday(dateStr) ? [...base, 5] : [...base, 4];
}

function isLastHour(dateStr, hour) {
  return isFridayOrSaturday(dateStr) ? Number(hour) === 5 : Number(hour) === 4;
}

/* =========================
   地域判定
========================= */

function classifyAreaByAddress(address) {
  const a = normalizeAddressText(address);
  if (!a) return "";

  if (a.includes("江戸川区")) {
    if (["西小岩", "南小岩", "北小岩", "上一色", "本一色"].some(k => a.includes(k))) return "江戸川・小岩";
    if (["篠崎", "瑞江", "一之江", "船堀", "葛西", "西葛西"].some(k => a.includes(k))) return "江戸川南部";
    return "江戸川";
  }

  if (a.includes("葛飾区")) {
    if (["金町", "柴又", "高砂", "新宿"].some(k => a.includes(k))) return "葛飾東部";
    if (["青戸", "立石", "亀有", "奥戸", "堀切"].some(k => a.includes(k))) return "葛飾西部";
    return "葛飾";
  }

  if (a.includes("足立区")) {
    if (["綾瀬", "六町", "北綾瀬"].some(k => a.includes(k))) return "足立東部";
    if (["北千住", "西新井", "竹ノ塚", "梅島"].some(k => a.includes(k))) return "足立西部";
    return "足立";
  }

  if (a.includes("江東区")) return "江東";
  if (a.includes("墨田区")) return "墨田";
  if (a.includes("台東区")) return "台東";
  if (a.includes("荒川区")) return "荒川";
  if (a.includes("中央区")) return "中央区";
  if (a.includes("千代田区")) return "千代田区";
  if (a.includes("文京区")) return "文京区";

  if (a.includes("船橋市")) {
    if (["西船", "西船橋", "本中山", "二子", "二子町", "海神", "山野町", "印内", "葛飾町", "東中山"].some(k => a.includes(k))) return "船橋西部";
    if (["前原", "前原西", "前原東", "津田沼", "東船橋"].some(k => a.includes(k))) return "船橋中部";
    if (["習志野台", "薬円台", "北習志野", "高根台", "芝山", "飯山満"].some(k => a.includes(k))) return "船橋東部";
    return "船橋";
  }

  if (a.includes("市川市")) {
    if (["八幡", "南八幡", "鬼越", "鬼高", "本八幡", "東菅野", "本北方", "北方", "若宮", "高石神", "中山"].some(k => a.includes(k))) return "市川・本八幡";
    if (["中国分", "国分", "国府台", "堀之内", "須和田", "曽谷", "宮久保", "菅野"].some(k => a.includes(k))) return "市川北部";
    if (["行徳", "妙典", "南行徳", "相之川", "新浜", "福栄", "末広"].some(k => a.includes(k))) return "行徳・南行徳";
    return "市川";
  }

  if (a.includes("松戸市")) {
    if (["新松戸", "新松戸北", "新松戸東", "幸谷", "八ケ崎", "八ヶ崎", "二ツ木", "中和倉"].some(k => a.includes(k))) return "新松戸";
    if (["北松戸", "馬橋", "西馬橋", "栄町", "栄町西", "中根", "中根長津町"].some(k => a.includes(k))) return "北松戸・馬橋";
    if (["常盤平", "常盤平陣屋前", "常盤平西窪町", "常盤平双葉町", "常盤平柳町", "五香", "五香西", "五香南", "金ケ作", "金ヶ作"].some(k => a.includes(k))) return "常盤平";
    if (["日暮", "河原塚", "千駄堀", "牧の原", "松飛台", "稔台", "みのり台", "八柱", "常盤平松葉町"].some(k => a.includes(k))) return "八柱・みのり台";
    if (["東松戸", "秋山", "高塚新田", "紙敷", "大橋", "和名ケ谷", "和名ヶ谷"].some(k => a.includes(k))) return "東松戸・秋山";
    if (["北国分", "下矢切", "上矢切", "栗山", "三矢小台", "矢切"].some(k => a.includes(k))) return "北国分・矢切";
    if (["根本", "小根本", "樋野口", "古ケ崎", "古ヶ崎", "上本郷", "岩瀬", "胡録台", "緑ケ丘", "緑ヶ丘"].some(k => a.includes(k))) return "松戸駅周辺";
    return "松戸";
  }

  if (a.includes("柏市")) {
    if (["旭町", "明原", "あけぼの", "末広町", "泉町", "中央町", "柏", "東上町"].some(k => a.includes(k))) return "柏駅周辺";
    if (["豊四季", "南柏", "今谷", "新富町"].some(k => a.includes(k))) return "南柏・豊四季";
    if (["北柏", "根戸", "松葉町", "花野井", "柏の葉"].some(k => a.includes(k))) return "北柏・柏の葉";
    return "柏";
  }

  if (a.includes("流山市")) {
    if (["南流山"].some(k => a.includes(k))) return "南流山";
    if (["おおたかの森", "初石", "駒木"].some(k => a.includes(k))) return "おおたかの森";
    if (["江戸川台", "西初石", "東初石"].some(k => a.includes(k))) return "江戸川台・初石";
    return "流山";
  }

  if (a.includes("我孫子市")) return "我孫子";
  if (a.includes("鎌ケ谷市") || a.includes("鎌ヶ谷市")) return "鎌ケ谷";
  if (a.includes("野田市")) return "野田";
  if (a.includes("習志野市")) return "習志野";
  if (a.includes("浦安市")) return "浦安";
  if (a.includes("八千代市")) return "八千代";
  if (a.includes("白井市")) return "白井";
  if (a.includes("印西市")) return "印西";
  if (a.includes("佐倉市")) return "佐倉";
  if (a.includes("成田市")) return "成田";

  if (a.includes("三郷市")) {
    if (["新三郷", "彦成", "彦糸", "采女", "早稲田"].some(k => a.includes(k))) return "三郷北部";
    if (["戸ケ崎", "戸ヶ崎", "高州", "鷹野", "彦沢", "栄"].some(k => a.includes(k))) return "三郷南部";
    return "三郷";
  }

  if (a.includes("八潮市")) {
    if (["八潮", "中央", "大瀬", "茜町"].some(k => a.includes(k))) return "八潮中心部";
    if (["大原", "木曽根", "浮塚", "南後谷", "伊勢野"].some(k => a.includes(k))) return "八潮南部";
    return "八潮";
  }

  if (a.includes("草加市")) {
    if (["草加", "氷川町", "中央", "高砂", "住吉"].some(k => a.includes(k))) return "草加中心部";
    if (["谷塚", "瀬崎", "新田", "松原", "獨協大学前"].some(k => a.includes(k))) return "草加周辺";
    return "草加";
  }

  if (a.includes("吉川市")) return "吉川";
  if (a.includes("越谷市")) return "越谷";
  if (a.includes("川口市")) return "川口";
  if (a.includes("さいたま市")) {
    if (a.includes("岩槻区")) return "さいたま・岩槻";
    if (a.includes("浦和区") || a.includes("南区") || a.includes("緑区")) return "さいたま南部";
    if (a.includes("大宮区") || a.includes("見沼区") || a.includes("北区") || a.includes("西区")) return "さいたま北部";
    return "さいたま";
  }

  if (a.includes("取手市")) {
    if (["取手", "新町", "井野", "白山", "台宿"].some(k => a.includes(k))) return "取手中心部";
    if (["藤代", "宮和田", "双葉", "谷中", "桜が丘"].some(k => a.includes(k))) return "藤代";
    return "取手";
  }

  if (a.includes("守谷市")) {
    if (["中央", "松ケ丘", "松ヶ丘", "けやき台", "ひがし野", "本町"].some(k => a.includes(k))) return "守谷中心部";
    if (["みずき野", "百合ケ丘", "百合ヶ丘", "御所ケ丘", "御所ヶ丘"].some(k => a.includes(k))) return "守谷西部";
    return "守谷";
  }

  if (a.includes("つくばみらい市")) {
    if (["みらい平", "陽光台", "紫峰ヶ丘", "富士見ヶ丘", "筒戸"].some(k => a.includes(k))) return "みらい平";
    return "つくばみらい";
  }

  if (a.includes("常総市")) return "常総";
  if (a.includes("龍ケ崎市") || a.includes("龍ヶ崎市")) return "龍ケ崎";
  if (a.includes("牛久市")) return "牛久";
  if (a.includes("つくば市")) return "つくば";
  if (a.includes("土浦市")) return "土浦";
  if (a.includes("坂東市")) return "坂東";
  if (a.includes("利根町")) return "利根町";

  return "";
}

function classifyAreaByLatLng(lat, lng) {
  if (!isValidLatLng(lat, lng)) return "周辺";

  if (lng >= 139.99) {
    if (lat >= 35.75) return "葛飾東部";
    if (lat >= 35.71) return "江戸川・小岩";
    return "東京東部";
  }

  if (lat >= 35.79 && lat < 35.86 && lng >= 139.78 && lng < 139.88) return "草加・八潮・三郷";
  if (lat >= 35.84 && lat < 35.92 && lng >= 139.84 && lng < 139.92) return "三郷・吉川";
  if (lat >= 35.86 && lat < 35.95 && lng >= 139.74 && lng < 139.84) return "越谷";

  if (lat >= 35.88 && lat < 35.98 && lng >= 140.00 && lng < 140.10) return "取手";
  if (lat >= 35.93 && lat < 36.03 && lng >= 139.95 && lng < 140.06) return "守谷";
  if (lat >= 35.97 && lat < 36.08 && lng >= 139.98 && lng < 140.08) return "つくばみらい";
  if (lat >= 35.94 && lat < 36.08 && lng >= 140.10 && lng < 140.22) return "龍ケ崎・牛久";

  if (lat >= 35.69 && lat < 35.74 && lng >= 139.93 && lng < 139.99) return "船橋西部";
  if (lat >= 35.71 && lat < 35.76 && lng >= 139.89 && lng < 139.96) return "市川・本八幡";
  if (lat >= 35.73 && lat < 35.78 && lng >= 139.87 && lng < 139.93) return "市川北部";
  if (lat >= 35.74 && lat < 35.80 && lng >= 139.98 && lng < 140.09) return "八千代・習志野";
  if (lat >= 35.81 && lat < 35.86 && lng >= 139.90 && lng < 139.96) return "新松戸";
  if (lat >= 35.80 && lat < 35.84 && lng >= 139.90 && lng < 139.95) return "北松戸・馬橋";
  if (lat >= 35.79 && lat < 35.83 && lng >= 139.85 && lng < 139.91) return "常盤平";
  if (lat >= 35.76 && lat < 35.81 && lng >= 139.89 && lng < 139.95) return "八柱・みのり台";
  if (lat >= 35.75 && lat < 35.79 && lng >= 139.93 && lng < 139.98) return "東松戸・秋山";
  if (lat >= 35.74 && lat < 35.78 && lng >= 139.86 && lng < 139.91) return "北国分・矢切";
  if (lat >= 35.84 && lng >= 139.94 && lng < 140.02) return "柏";
  if (lat >= 35.87 && lng >= 139.90 && lng < 139.98) return "流山";
  if (lat >= 35.86 && lng < 139.90) return "我孫子";
  if (lat >= 35.75 && lat < 35.79 && lng >= 139.87 && lng < 139.93) return "鎌ケ谷";

  return "広域周辺";
}

function guessArea(lat, lng, address = "") {
  const byAddress = classifyAreaByAddress(address);
  if (byAddress) return byAddress;
  return classifyAreaByLatLng(lat, lng);
}

function getDisplayArea(cast) {
  const manualArea = String(cast.area || "").trim();
  if (manualArea) return manualArea;

  const address = cast.address || "";
  const byAddress = classifyAreaByAddress(address);
  if (byAddress) return byAddress;

  const lat = toNullableNumber(cast.latitude);
  const lng = toNullableNumber(cast.longitude);

  if (lat !== null && lng !== null) {
    return classifyAreaByLatLng(lat, lng);
  }

  return "";
}

/* =========================
   AI配車 v2
========================= */

function areaDistanceScore(areaA, areaB) {
  if (!areaA || !areaB) return 999;
  if (areaA === areaB) return 0;

  const nearMap = {
    "松戸": ["新松戸", "北松戸・馬橋", "常盤平", "八柱・みのり台", "東松戸・秋山", "北国分・矢切", "松戸駅周辺"],
    "柏": ["柏", "柏駅周辺", "南柏・豊四季", "北柏・柏の葉", "流山", "南流山", "おおたかの森"],
    "八千代": ["八千代", "船橋東部", "船橋中部", "習志野", "船橋"],
    "三郷": ["三郷", "三郷北部", "三郷南部", "八潮", "草加", "吉川"],
    "守谷": ["守谷", "守谷中心部", "守谷西部", "取手", "藤代", "みらい平", "つくばみらい"],
    "江戸川": ["江戸川", "江戸川・小岩", "江戸川南部", "葛飾", "葛飾東部", "葛飾西部"],
    "市川": ["市川", "市川・本八幡", "市川北部", "行徳・南行徳", "船橋西部"],
    "船橋": ["船橋", "船橋西部", "船橋中部", "船橋東部", "八千代"]
  };

  const nearList = nearMap[areaB] || [];
  if (nearList.includes(areaA)) return 1;
  return 2;
}

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

function getGlobalWorstReturnDistance(assignments, vehicles) {
  let worst = 0;
  vehicles.forEach(vehicle => {
    const d = getVehicleCurrentMaxDistance(assignments, vehicle.id);
    if (d > worst) worst = d;
  });
  return worst;
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

function groupCastsByArea(casts) {
  const grouped = new Map();
  casts.forEach(cast => {
    const key = cast.area || "未分類";
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(cast);
  });

  return [...grouped.entries()].map(([area, members]) => ({
    area,
    members
  }));
}

function sortAreaGroupsByDistance(groups) {
  return groups.sort((a, b) => {
    const aMin = Math.min(...a.members.map(x => Number(x.distance_km || 9999)));
    const bMin = Math.min(...b.members.map(x => Number(x.distance_km || 9999)));
    return aMin - bMin;
  });
}

function chooseBestVehicleForCastBalanced(
  cast,
  vehicles,
  assignments,
  isLast,
  preferredArea = "",
  monthlyDistanceMap = new Map()
) {
  let bestVehicle = null;
  let bestScore = Infinity;
  const currentWorst = getGlobalWorstReturnDistance(assignments, vehicles);

  vehicles.forEach(vehicle => {
    if (!canAssignToVehicle(vehicle, assignments)) return;

    const vehicleAssignments = getVehicleAssignments(assignments, vehicle.id);
    const currentLoad = vehicleAssignments.length;
    const currentMaxDistance = getVehicleCurrentMaxDistance(assignments, vehicle.id);
    const castDistance = Number(cast.distance_km || 0);
    const newVehicleMaxDistance = Math.max(currentMaxDistance, castDistance);

    let newGlobalWorst = 0;
    vehicles.forEach(v => {
      if (Number(v.id) === Number(vehicle.id)) {
        if (newVehicleMaxDistance > newGlobalWorst) newGlobalWorst = newVehicleMaxDistance;
      } else {
        const d = getVehicleCurrentMaxDistance(assignments, v.id);
        if (d > newGlobalWorst) newGlobalWorst = d;
      }
    });

    let score = 0;

    score += newGlobalWorst * 30;
    score += Math.max(0, newGlobalWorst - currentWorst) * 50;
    score += currentLoad * 12;

    const sameAreaCount = vehicleAssignments.filter(x => x.area === cast.area).length;
    if (sameAreaCount > 0) score -= 8;

    if (isLast) {
      score += areaDistanceScore(cast.area, vehicle.home_area || preferredArea) * 25;
    }

    const avgDistance = getVehicleMonthlyAvgDistance(monthlyDistanceMap, vehicle.id);
    score += avgDistance * 0.15;

    if (score < bestScore) {
      bestScore = score;
      bestVehicle = vehicle;
    }
  });

  return bestVehicle;
}

function optimizeDispatchAssignmentsV2({ casts, vehicles, planDate, planHour, monthlyDistanceMap }) {
  const result = [];
  const activeVehicles = vehicles.filter(v =>
    v.status === "working" || v.status === "waiting" || v.status === "running"
  );
  const lastHour = isLastHour(planDate, planHour);
  const groups = sortAreaGroupsByDistance(groupCastsByArea(casts));

  groups.forEach(group => {
    const sortedMembers = [...group.members].sort((a, b) => {
      const aKm = Number(a.distance_km || 9999);
      const bKm = Number(b.distance_km || 9999);
      return aKm - bKm;
    });

    sortedMembers.forEach(cast => {
      const vehicle = chooseBestVehicleForCastBalanced(
        cast,
        activeVehicles,
        result,
        lastHour,
        group.area,
        monthlyDistanceMap
      );

      if (!vehicle) return;

      result.push({
        cast_id: cast.id,
        cast_name: cast.name,
        area: cast.area || "",
        distance_km: cast.distance_km || null,
        vehicle_id: vehicle.id,
        vehicle_label: `${vehicle.plate_number} ${vehicle.vehicle_name || ""}`.trim(),
        driver_name: vehicle.driver_name || "",
        home_area: vehicle.home_area || "",
        line_id: vehicle.line_id || "",
        plan_hour: planHour,
        plan_date: planDate
      });
    });
  });

  return result;
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
      const aVehicle = String(a.vehicle_id || "");
      const bVehicle = String(b.vehicle_id || "");
      if (aVehicle !== bVehicle) return aVehicle.localeCompare(bVehicle);
      return Number(a.stop_order || 0) - Number(b.stop_order || 0);
    })
    .forEach(item => {
      const key = Number(item.vehicle_id || 0);
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(item);
    });

  const lines = [];
  lines.push("【本日の送り配車】");
  lines.push("");

  for (const [vehicleId, rows] of grouped.entries()) {
    const vehicle = vehicles.find(v => Number(v.id) === Number(vehicleId));
    const driverName = vehicle?.driver_name || rows[0]?.driver_name || "ドライバー未設定";
    const lineId = vehicle?.line_id || "";
    const header = `${lineId ? lineId + " " : ""}${driverName}`;

    lines.push(header);

    rows.forEach((row, index) => {
      const castName = row.casts?.name || "不明";
      const area = row.destination_area || "-";
      lines.push(`${index + 1}件目 ${castName} / ${area}`);
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
    hasValue(cast.latitude) && hasValue(cast.longitude)
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

  const manualArea = els.castArea.value.trim();
  const autoArea =
    classifyAreaByAddress(els.castAddress.value) ||
    (lat !== null && lng !== null ? classifyAreaByLatLng(lat, lng) : "");

  const payload = {
    name,
    phone: els.castPhone.value.trim(),
    address: els.castAddress.value.trim(),
    area: manualArea || autoArea || null,
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

  await addHistory(
    null,
    null,
    editingCastId ? "update_cast" : "create_cast",
    editingCastId ? "キャストを更新" : "キャストを作成"
  );

  resetCastForm();
  await loadCasts();
  await refreshPlanCastSelect();
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
  await refreshPlanCastSelect();
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
  if (!els.castSelect) return;

  els.castSelect.innerHTML = `<option value="">選択してください</option>`;
  const usedCastIds = getUsedCastIdsInCurrentDispatch();

  casts
    .filter(cast => !usedCastIds.has(Number(cast.id)))
    .forEach(cast => {
      const displayArea = getDisplayArea(cast);
      const option = document.createElement("option");
      option.value = cast.id;
      option.textContent = `${cast.name} | ${displayArea || "地域未設定"}`;
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
      <p>地域: ${escapeHtml(displayArea || "-")}</p>
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
  els.vehicleName.value = "";
  els.vehicleDriverName.value = "";
  els.vehicleLineId.value = "";
  els.vehicleStatus.value = "waiting";
  els.vehicleSeatCapacity.value = "4";
  els.vehicleHomeArea.value = "";
  els.vehicleCapacityNote.value = "";
  els.vehicleMemo.value = "";
  els.saveVehicleBtn.textContent = "保存";
  els.cancelVehicleEditBtn.classList.add("hidden");
}

function fillVehicleForm(vehicle) {
  editingVehicleId = vehicle.id;
  els.vehiclePlateNumber.value = vehicle.plate_number || "";
  els.vehicleName.value = vehicle.vehicle_name || "";
  els.vehicleDriverName.value = vehicle.driver_name || "";
  els.vehicleLineId.value = vehicle.line_id || "";
  els.vehicleStatus.value = vehicle.status || "waiting";
  els.vehicleSeatCapacity.value = String(vehicle.seat_capacity || 4);
  els.vehicleHomeArea.value = vehicle.home_area || "";
  els.vehicleCapacityNote.value = vehicle.capacity_note || "";
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
    vehicle_name: els.vehicleName.value.trim(),
    driver_name: els.vehicleDriverName.value.trim(),
    line_id: els.vehicleLineId.value.trim(),
    status: els.vehicleStatus.value,
    seat_capacity: Number(els.vehicleSeatCapacity.value || 4),
    home_area: els.vehicleHomeArea.value.trim(),
    capacity_note: els.vehicleCapacityNote.value.trim(),
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

  await addHistory(
    null,
    null,
    editingVehicleId ? "update_vehicle" : "create_vehicle",
    editingVehicleId ? "車両を更新" : "車両を登録"
  );

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
  renderVehicleSelect(allVehiclesCache);
  renderVehicleList(allVehiclesCache);
  renderReportVehicleSelect(allVehiclesCache);
}

function renderVehicleSelect(vehicles) {
  if (!els.vehicleSelect) return;

  els.vehicleSelect.innerHTML = `<option value="">選択してください</option>`;
  vehicles.forEach(vehicle => {
    const option = document.createElement("option");
    option.value = vehicle.id;
    option.textContent = `${vehicle.plate_number} | ${vehicle.vehicle_name || "車種未設定"} | ${vehicleStatusLabel(vehicle.status)}`;
    option.dataset.label = `${vehicle.plate_number} ${vehicle.vehicle_name || ""}`.trim();
    els.vehicleSelect.appendChild(option);
  });
}

function renderReportVehicleSelect(vehicles) {
  if (!els.reportVehicleSelect) return;

  els.reportVehicleSelect.innerHTML = `<option value="">選択してください</option>`;
  vehicles.forEach(vehicle => {
    const option = document.createElement("option");
    option.value = vehicle.id;
    option.textContent = `${vehicle.plate_number} | ${vehicle.vehicle_name || "車種未設定"} | ${vehicle.driver_name || "-"}`;
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
      <p>車両名: ${escapeHtml(vehicle.vehicle_name || "-")}</p>
      <p>担当ドライバー: ${escapeHtml(vehicle.driver_name || "-")}</p>
      <p>LINE ID: ${escapeHtml(vehicle.line_id || "-")}</p>
      <p>状態: <span class="badge ${badgeClass}">${escapeHtml(vehicleStatusLabel(vehicle.status))}</span></p>
      <p>乗車可能人数: ${escapeHtml(vehicle.seat_capacity || 4)}人</p>
      <p>帰宅方面: ${escapeHtml(vehicle.home_area || "-")}</p>
      <p>積載メモ: ${escapeHtml(vehicle.capacity_note || "-")}</p>
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
   送り予定表
========================= */

function resetPlanForm() {
  editingPlanId = null;
  els.planDate.value = els.planDate.value || todayStr();
  rebuildPlanHourOptions(els.planDate.value);
  els.planCastSelect.value = "";
  els.planArea.value = "";
  els.planNote.value = "";
  els.savePlanBtn.textContent = "予定保存";
  els.cancelPlanEditBtn.classList.add("hidden");
}

function rebuildPlanHourOptions(dateStr) {
  const hours = getPlanHourOptions(dateStr);
  const current = Number(els.planHour.value || hours[0]);
  els.planHour.innerHTML = "";
  hours.forEach(hour => {
    const option = document.createElement("option");
    option.value = String(hour);
    option.textContent = `${hour}時`;
    els.planHour.appendChild(option);
  });

  if (hours.includes(current)) {
    els.planHour.value = String(current);
  } else {
    els.planHour.value = String(hours[0]);
  }
}

function refreshPlanCastSelect() {
  if (!els.planCastSelect) return;

  const blockedIds = getPlannedOrBlockedCastIds();
  const currentEditCastId = editingPlanId
    ? Number(currentPlansCache.find(x => Number(x.id) === Number(editingPlanId))?.cast_id || 0)
    : null;

  els.planCastSelect.innerHTML = `<option value="">選択してください</option>`;

  allCastsCache
    .filter(cast => !blockedIds.has(Number(cast.id)) || Number(cast.id) === currentEditCastId)
    .forEach(cast => {
      const displayArea = getDisplayArea(cast);
      const option = document.createElement("option");
      option.value = cast.id;
      option.textContent = `${cast.name} | ${displayArea || "地域未設定"}`;
      option.dataset.area = displayArea || "";
      els.planCastSelect.appendChild(option);
    });
}

function fillPlanForm(plan) {
  editingPlanId = plan.id;
  els.planDate.value = plan.plan_date;
  rebuildPlanHourOptions(plan.plan_date);
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

  const cast = allCastsCache.find(x => Number(x.id) === castId);
  const plannedArea = els.planArea.value.trim() || getDisplayArea(cast);

  const payload = {
    plan_date: planDate,
    plan_hour: planHour,
    cast_id: castId,
    planned_area: plannedArea || null,
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
  if (!els.planSelect) return;

  els.planSelect.innerHTML = `<option value="">予定を選択してください</option>`;

  plans
    .filter(plan => plan.status === "planned")
    .forEach(plan => {
      const option = document.createElement("option");
      option.value = plan.id;
      option.textContent = `${plan.plan_hour}時 | ${plan.casts?.name || "不明"} | ${plan.planned_area || "-"}`;
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

function renderPlansBoard(plans, dateStr) {
  els.plansBoard.innerHTML = "";

  if (!plans.length) {
    els.plansBoard.innerHTML = `<div class="item"><p>予定はまだありません。</p></div>`;
    return;
  }

  const hours = getPlanHourOptions(dateStr);

  hours.forEach(hour => {
    const hourPlans = plans.filter(x => Number(x.plan_hour) === Number(hour));
    const wrapper = document.createElement("div");
    wrapper.className = "item";

    if (!hourPlans.length) {
      wrapper.innerHTML = `<h3>${hour}時</h3><p>予定なし</p>`;
      els.plansBoard.appendChild(wrapper);
      return;
    }

    const grouped = new Map();
    hourPlans.forEach(plan => {
      const key = plan.planned_area || getDisplayArea(plan.casts || {}) || "未分類";
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(plan);
    });

    let inner = `<h3>${hour}時${isLastHour(dateStr, hour) ? "（ラスト便）" : ""}</h3>`;

    [...grouped.entries()].forEach(([area, list]) => {
      inner += `<p><strong>${escapeHtml(area)}</strong></p>`;
      list.forEach(plan => {
        const statusClass = planStatusBadgeClass(plan.status);
        inner += `
          <div class="item" style="margin-top:8px;">
            <p>${escapeHtml(plan.casts?.name || "不明")} / <span class="badge ${statusClass}">${escapeHtml(planStatusLabel(plan.status))}</span></p>
            <p>住所: ${escapeHtml(plan.casts?.address || "-")}</p>
            <p>メモ: ${escapeHtml(plan.note || "-")}</p>
            <div class="actions">
              <button class="btn edit-plan-btn" data-id="${plan.id}">編集</button>
              <button class="btn secondary plan-done-btn" data-id="${plan.id}">完了</button>
              <button class="btn danger plan-cancel-btn" data-id="${plan.id}">キャンセル</button>
              <button class="btn danger delete-plan-btn" data-id="${plan.id}">削除</button>
            </div>
          </div>
        `;
      });
    });

    wrapper.innerHTML = inner;

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

    wrapper.querySelectorAll(".delete-plan-btn").forEach(btn => {
      btn.addEventListener("click", async () => {
        await deletePlan(Number(btn.dataset.id));
      });
    });

    els.plansBoard.appendChild(wrapper);
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
    els.vehicleSelect.value = existing[0].vehicle_id ? String(existing[0].vehicle_id) : "";
    await loadDispatchItems(currentDispatchId);
    return;
  }

  const selectedVehicle = els.vehicleSelect.selectedOptions?.[0] || null;
  const vehicleId = els.vehicleSelect.value ? Number(els.vehicleSelect.value) : null;
  const vehicleLabel = selectedVehicle?.dataset.label || null;

  const { data, error } = await supabaseClient
    .from("dispatches")
    .insert({
      dispatch_date: dispatchDate,
      driver_name: driverName,
      vehicle_id: vehicleId,
      vehicle_label: vehicleLabel,
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
  await addHistory(currentDispatchId, null, "create_dispatch", `配車を作成: ${dispatchDate}`);
  await loadDispatchItems(currentDispatchId);
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
    els.vehicleSelect.value = dispatch.vehicle_id ? String(dispatch.vehicle_id) : "";
    await loadDispatchItems(dispatch.id);
  } else {
    currentDispatchItemsCache = [];
    renderCastSelect(allCastsCache);
    els.dispatchSummary.textContent = "";
    els.dispatchList.innerHTML = `<div class="item"><p>この日の配車はまだありません。</p></div>`;
    els.vehicleSelect.value = "";
  }
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

  const alreadyExists = currentDispatchItemsCache.some(item => Number(item.cast_id) === castId);
  if (alreadyExists) {
    alert("このキャストはすでに配車に追加されています");
    return;
  }

  const selected = els.castSelect.selectedOptions[0];
  const destinationAddress = selected?.dataset.address || "";
  const lat = toNullableNumber(selected?.dataset.lat);
  const lng = toNullableNumber(selected?.dataset.lng);
  const area = selected?.dataset.area || (lat !== null && lng !== null ? guessArea(lat, lng, destinationAddress) : "");
  const distanceKm = lat !== null && lng !== null ? estimateRoadKmFromStation(lat, lng) : null;
  const travelMinutes = distanceKm !== null ? estimateMinutes(distanceKm) : null;

  let stopOrder = Number(els.stopOrder.value || 1);
  if (!Number.isFinite(stopOrder) || stopOrder <= 0) {
    stopOrder = currentDispatchItemsCache.length + 1;
  }

  const selectedVehicle = els.vehicleSelect.selectedOptions?.[0] || null;
  const vehicleId = els.vehicleSelect.value ? Number(els.vehicleSelect.value) : null;
  const vehicleLabel = selectedVehicle?.dataset.label || null;
  const driverName = els.driverName.value.trim();

  const { data, error } = await supabaseClient
    .from("dispatch_items")
    .insert({
      dispatch_id: currentDispatchId,
      cast_id: castId,
      stop_order: stopOrder,
      pickup_label: ORIGIN_LABEL,
      destination_address: destinationAddress,
      destination_area: area,
      latitude: lat,
      longitude: lng,
      distance_km: distanceKm,
      travel_minutes: travelMinutes,
      vehicle_id: vehicleId,
      vehicle_label: vehicleLabel,
      driver_name: driverName,
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
  await addHistory(currentDispatchId, data.id, "add_cast", "キャストを配車に追加");
  await loadDispatchItems(currentDispatchId);
  await loadHistory();
}

async function addPlanToDispatch() {
  if (!currentDispatchId) {
    alert("先に配車を作成または読込してください");
    return;
  }

  const planId = Number(els.planSelect.value);
  if (!planId) {
    alert("予定を選択してください");
    return;
  }

  const plan = currentPlansCache.find(x => Number(x.id) === planId);
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
  const area = plan.planned_area || getDisplayArea(cast);
  const distanceKm = lat !== null && lng !== null ? estimateRoadKmFromStation(lat, lng) : null;
  const travelMinutes = distanceKm !== null ? estimateMinutes(distanceKm) : null;
  const stopOrder = currentDispatchItemsCache.length + 1;

  const selectedVehicle = els.vehicleSelect.selectedOptions?.[0] || null;
  const vehicleId = els.vehicleSelect.value ? Number(els.vehicleSelect.value) : null;
  const vehicleLabel = selectedVehicle?.dataset.label || null;
  const driverName = els.driverName.value.trim();

  const { data, error } = await supabaseClient
    .from("dispatch_items")
    .insert({
      dispatch_id: currentDispatchId,
      cast_id: cast.id,
      stop_order: stopOrder,
      pickup_label: ORIGIN_LABEL,
      destination_address: cast.address || "",
      destination_area: area || "",
      latitude: lat,
      longitude: lng,
      distance_km: distanceKm,
      travel_minutes: travelMinutes,
      vehicle_id: vehicleId,
      vehicle_label: vehicleLabel,
      driver_name: driverName,
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

  const nextOrder = currentDispatchItemsCache.length + 1;
  els.stopOrder.value = String(nextOrder);
}

function renderDispatchSummary(items) {
  if (!items.length) {
    els.dispatchSummary.textContent = "配車先はまだありません。";
    return;
  }

  const active = items.filter(x => normalizeStatus(x.status) === "pending");
  const done = items.filter(x => normalizeStatus(x.status) === "done");
  const cancel = items.filter(x => normalizeStatus(x.status) === "cancel");
  const totalKm = active.reduce((sum, x) => sum + Number(x.distance_km || 0), 0);
  const totalMin = active.reduce((sum, x) => sum + Number(x.travel_minutes || 0), 0);
  const areas = [...new Set(active.map(x => x.destination_area).filter(Boolean))];
  const vehicleText =
    els.vehicleSelect && els.vehicleSelect.selectedOptions.length
      ? els.vehicleSelect.selectedOptions[0]?.textContent || ""
      : "";

  els.dispatchSummary.textContent =
    `未完了 ${active.length} 件 / 完了 ${done.length} 件 / キャンセル ${cancel.length} 件 / 推定合計距離 ${totalKm.toFixed(1)} km / 推定合計時間 ${totalMin} 分${areas.length ? ` / 地域 ${areas.join("・")}` : ""}${vehicleText ? ` / 車両 ${vehicleText}` : ""}`;
}

function buildVehicleChangeOptions(selectedVehicleId) {
  let html = `<option value="">未設定</option>`;
  allVehiclesCache.forEach(vehicle => {
    const selected = Number(vehicle.id) === Number(selectedVehicleId) ? "selected" : "";
    html += `<option value="${vehicle.id}" ${selected}>${escapeHtml(vehicle.plate_number)} | ${escapeHtml(vehicle.vehicle_name || "車種未設定")} | ${escapeHtml(vehicle.driver_name || "-")}</option>`;
  });
  return html;
}

function renderDispatchItems(items) {
  els.dispatchList.innerHTML = "";

  if (!items.length) {
    els.dispatchList.innerHTML = `<div class="item"><p>配車先はまだありません。</p></div>`;
    return;
  }

  items.forEach((item, index) => {
    const status = normalizeStatus(item.status);
    const badgeClass = status === "done" ? "done" : status === "cancel" ? "cancel" : "";
    const cardClass = status === "done" ? "done-card" : status === "cancel" ? "cancel-card" : "";
    const displayArea = item.destination_area || (item.casts ? getDisplayArea(item.casts) : "") || "-";

    const div = document.createElement("div");
    div.className = `item ${cardClass}`;
    div.innerHTML = `
      <h3>${item.stop_order}件目 | ${escapeHtml(item.casts?.name || "不明")}</h3>
      <p>地域: ${escapeHtml(displayArea)}</p>
      <p>住所: ${escapeHtml(item.destination_address || "-")}</p>
      <p>車両: ${escapeHtml(item.vehicle_label || "-")}</p>
      <p>ドライバー: ${escapeHtml(item.driver_name || "-")}</p>
      <p>推定距離: ${item.distance_km ?? "-"} km</p>
      <p>推定時間: ${item.travel_minutes ?? "-"} 分</p>
      <p><span class="badge ${badgeClass}">${escapeHtml(status)}</span></p>

      <label>
        <span>車両変更</span>
        <select class="dispatch-vehicle-change" data-id="${item.id}">
          ${buildVehicleChangeOptions(item.vehicle_id)}
        </select>
      </label>

      <div class="actions">
        <button class="btn secondary route-btn" data-address="${escapeHtml(item.destination_address || "")}">ルート</button>
        <button class="btn secondary up-btn" data-id="${item.id}" ${index === 0 ? "disabled" : ""}>上へ</button>
        <button class="btn secondary down-btn" data-id="${item.id}" ${index === items.length - 1 ? "disabled" : ""}>下へ</button>
        <button class="btn done-btn" data-id="${item.id}">完了</button>
        <button class="btn danger cancel-btn" data-id="${item.id}">キャンセル</button>
        <button class="btn danger delete-item-btn" data-id="${item.id}">削除</button>
      </div>
    `;

    div.querySelector(".route-btn")?.addEventListener("click", e => {
      const address = e.currentTarget.dataset.address;
      if (address) openRouteFromMatsudo(address);
    });

    div.querySelector(".done-btn")?.addEventListener("click", async e => {
      await updateItemStatus(Number(e.currentTarget.dataset.id), "done");
    });

    div.querySelector(".cancel-btn")?.addEventListener("click", async e => {
      await updateItemStatus(Number(e.currentTarget.dataset.id), "cancel");
    });

    div.querySelector(".delete-item-btn")?.addEventListener("click", async e => {
      await deleteDispatchItem(Number(e.currentTarget.dataset.id));
    });

    div.querySelector(".up-btn")?.addEventListener("click", async e => {
      await moveDispatchItem(Number(e.currentTarget.dataset.id), -1);
    });

    div.querySelector(".down-btn")?.addEventListener("click", async e => {
      await moveDispatchItem(Number(e.currentTarget.dataset.id), 1);
    });

    div.querySelector(".dispatch-vehicle-change")?.addEventListener("change", async e => {
      await updateDispatchItemVehicle(Number(e.currentTarget.dataset.id), Number(e.currentTarget.value || 0));
    });

    els.dispatchList.appendChild(div);
  });
}

async function updateDispatchItemVehicle(itemId, vehicleId) {
  const vehicle = allVehiclesCache.find(v => Number(v.id) === Number(vehicleId));

  const payload = {
    vehicle_id: vehicleId || null,
    vehicle_label: vehicle ? `${vehicle.plate_number} ${vehicle.vehicle_name || ""}`.trim() : null,
    driver_name: vehicle?.driver_name || null
  };

  const { error } = await supabaseClient
    .from("dispatch_items")
    .update(payload)
    .eq("id", itemId);

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(currentDispatchId, itemId, "change_vehicle", "配車車両を変更");
  await loadDispatchItems(currentDispatchId);
  await loadHistory();
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

async function moveDispatchItem(itemId, direction) {
  const items = [...currentDispatchItemsCache].sort((a, b) => Number(a.stop_order) - Number(b.stop_order));
  const index = items.findIndex(item => Number(item.id) === Number(itemId));
  if (index < 0) return;

  const targetIndex = index + direction;
  if (targetIndex < 0 || targetIndex >= items.length) return;

  const temp = items[index];
  items[index] = items[targetIndex];
  items[targetIndex] = temp;

  for (let i = 0; i < items.length; i++) {
    await supabaseClient
      .from("dispatch_items")
      .update({ stop_order: i + 1 })
      .eq("id", items[i].id);
  }

  await addHistory(
    currentDispatchId,
    itemId,
    "move_dispatch_item",
    direction < 0 ? "配車項目を上へ移動" : "配車項目を下へ移動"
  );
  await loadDispatchItems(currentDispatchId);
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
  const hour = currentPlansCache.find(x => x.status === "assigned")?.plan_hour ?? 0;
  const monthlyDistanceMap = buildMonthlyDistanceMap(currentDailyReportsCache, getMonthKey(dateStr));

  const casts = activeItems.map(item => ({
    id: item.cast_id,
    name: item.casts?.name || "",
    area: item.destination_area || getDisplayArea(item.casts || {}),
    distance_km: Number(item.distance_km || 0)
  }));

  const aiResult = optimizeDispatchAssignmentsV2({
    casts,
    vehicles: allVehiclesCache,
    planDate: dateStr,
    planHour: hour,
    monthlyDistanceMap
  });

  for (let i = 0; i < aiResult.length; i++) {
    const match = activeItems.find(x => Number(x.cast_id) === Number(aiResult[i].cast_id));
    if (!match) continue;

    await supabaseClient
      .from("dispatch_items")
      .update({
        stop_order: i + 1,
        vehicle_id: aiResult[i].vehicle_id,
        vehicle_label: aiResult[i].vehicle_label,
        driver_name: aiResult[i].driver_name,
        plan_hour: aiResult[i].plan_hour,
        plan_date: aiResult[i].plan_date
      })
      .eq("id", match.id);
  }

  await addHistory(currentDispatchId, null, "optimize", "AI配車順を最適化");
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
        plate_number,
        vehicle_name
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
      <h3>${escapeHtml(row.vehicles?.plate_number || "-")} / ${escapeHtml(row.vehicles?.vehicle_name || "-")}</h3>
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
        plate_number,
        vehicle_name
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
      vehicle_label: `${row.vehicles?.plate_number || "-"} ${row.vehicles?.vehicle_name || ""}`.trim(),
      worked_days: 0,
      total_distance_km: 0
    };
    prev.worked_days += 1;
    prev.total_distance_km += Number(row.distance_km || 0);
    grouped.set(id, prev);
  });

  const rows = [...grouped.entries()].map(([vehicleId, value]) => ({
    vehicle_id: vehicleId,
    vehicle_label: value.vehicle_label,
    worked_days: value.worked_days,
    total_distance_km: value.total_distance_km,
    avg_distance_per_day: value.worked_days > 0 ? value.total_distance_km / value.worked_days : 0
  })).sort((a, b) => a.avg_distance_per_day - b.avg_distance_per_day);

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
      const manualArea = String(row.area || "").trim();
      const autoArea =
        classifyAreaByAddress(row.address) ||
        (lat !== null && lng !== null ? classifyAreaByLatLng(lat, lng) : "");

      return {
        name: String(row.name || "").trim(),
        phone: String(row.phone || "").trim(),
        address: String(row.address || "").trim(),
        area: manualArea || autoArea || null,
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
    rebuildPlanHourOptions(els.planDate.value || todayStr());
    await loadPlansByDate(els.planDate.value || todayStr());
  });

  els.planCastSelect?.addEventListener("change", () => {
    const selected = els.planCastSelect.selectedOptions?.[0];
    if (!els.planArea.value.trim()) {
      els.planArea.value = selected?.dataset.area || "";
    }
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

    const today = todayStr();
    if (els.dispatchDate) els.dispatchDate.value = today;
    if (els.planDate) els.planDate.value = today;
    if (els.reportDate) els.reportDate.value = today;

    rebuildPlanHourOptions(today);

    await loadCasts();
    await loadVehicles();
    await loadDispatchByDate(today);
    await loadPlansByDate(today);
    await loadDailyReports(today);
    await loadHistory();

    refreshPlanCastSelect();
    resetPlanForm();
  } catch (err) {
    console.error("dashboard init error:", err);
    alert("初期化中にエラーが発生しました。F12 の Console を確認してください。");
  }
});
