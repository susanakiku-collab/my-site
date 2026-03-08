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

let allCastsCache = [];
let allVehiclesCache = [];
let currentDispatchItemsCache = [];

const els = {
  userEmail: document.getElementById("userEmail"),
  logoutBtn: document.getElementById("logoutBtn"),

  dispatchDate: document.getElementById("dispatchDate"),
  driverName: document.getElementById("driverName"),
  vehicleSelect: document.getElementById("vehicleSelect"),
  createDispatchBtn: document.getElementById("createDispatchBtn"),
  loadDispatchBtn: document.getElementById("loadDispatchBtn"),
  optimizeBtn: document.getElementById("optimizeBtn"),
  castSelect: document.getElementById("castSelect"),
  stopOrder: document.getElementById("stopOrder"),
  addCastBtn: document.getElementById("addCastBtn"),
  dispatchSummary: document.getElementById("dispatchSummary"),
  dispatchList: document.getElementById("dispatchList"),

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
  castsList: document.getElementById("castsList"),

  vehiclePlateNumber: document.getElementById("vehiclePlateNumber"),
  vehicleName: document.getElementById("vehicleName"),
  vehicleDriverName: document.getElementById("vehicleDriverName"),
  vehicleStatus: document.getElementById("vehicleStatus"),
  vehicleCapacityNote: document.getElementById("vehicleCapacityNote"),
  vehicleMemo: document.getElementById("vehicleMemo"),
  saveVehicleBtn: document.getElementById("saveVehicleBtn"),
  cancelVehicleEditBtn: document.getElementById("cancelVehicleEditBtn"),
  vehiclesList: document.getElementById("vehiclesList"),

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
    if (["西船", "西船橋", "本中山", "二子", "二子町", "海神", "山野町", "印内", "葛飾町", "東中山"].some(k => a.includes(k))) {
      return "船橋西部";
    }
    if (["前原", "前原西", "前原東", "津田沼", "東船橋"].some(k => a.includes(k))) {
      return "船橋中部";
    }
    if (["習志野台", "薬円台", "北習志野", "高根台", "芝山"].some(k => a.includes(k))) {
      return "船橋東部";
    }
    return "船橋";
  }

  if (a.includes("市川市")) {
    if (["八幡", "南八幡", "鬼越", "鬼高", "本八幡", "東菅野", "本北方", "北方", "若宮", "高石神", "中山"].some(k => a.includes(k))) {
      return "市川・本八幡";
    }
    if (["中国分", "国分", "国府台", "堀之内", "須和田", "曽谷", "宮久保", "菅野"].some(k => a.includes(k))) {
      return "市川北部";
    }
    if (["行徳", "妙典", "南行徳", "相之川", "新浜", "福栄", "末広"].some(k => a.includes(k))) {
      return "行徳・南行徳";
    }
    return "市川";
  }

  if (a.includes("松戸市")) {
    if (["新松戸", "新松戸北", "新松戸東", "幸谷", "八ケ崎", "八ヶ崎", "二ツ木", "中和倉"].some(k => a.includes(k))) {
      return "新松戸";
    }
    if (["北松戸", "馬橋", "西馬橋", "栄町", "栄町西", "中根", "中根長津町"].some(k => a.includes(k))) {
      return "北松戸・馬橋";
    }
    if (["常盤平", "常盤平陣屋前", "常盤平西窪町", "常盤平双葉町", "常盤平柳町", "五香", "五香西", "五香南", "金ケ作", "金ヶ作"].some(k => a.includes(k))) {
      return "常盤平";
    }
    if (["日暮", "河原塚", "千駄堀", "牧の原", "松飛台", "稔台", "みのり台", "八柱", "常盤平松葉町"].some(k => a.includes(k))) {
      return "八柱・みのり台";
    }
    if (["東松戸", "秋山", "高塚新田", "紙敷", "大橋", "和名ケ谷", "和名ヶ谷"].some(k => a.includes(k))) {
      return "東松戸・秋山";
    }
    if (["北国分", "下矢切", "上矢切", "栗山", "三矢小台", "矢切"].some(k => a.includes(k))) {
      return "北国分・矢切";
    }
    if (["根本", "小根本", "樋野口", "古ケ崎", "古ヶ崎", "上本郷", "岩瀬", "胡録台", "緑ケ丘", "緑ヶ丘"].some(k => a.includes(k))) {
      return "松戸駅周辺";
    }
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

  const fallbackRules = [
    { name: "船橋西部", keywords: ["西船", "西船橋", "本中山", "二子町", "海神", "東中山"] },
    { name: "市川・本八幡", keywords: ["本八幡", "南八幡", "鬼越", "鬼高"] },
    { name: "新松戸", keywords: ["新松戸", "幸谷"] },
    { name: "常盤平", keywords: ["常盤平", "五香"] },
    { name: "東松戸・秋山", keywords: ["東松戸", "秋山"] },
    { name: "柏駅周辺", keywords: ["柏駅", "末広町", "旭町"] },
    { name: "南流山", keywords: ["南流山"] },
    { name: "葛飾東部", keywords: ["金町", "柴又", "高砂"] },
    { name: "江戸川・小岩", keywords: ["小岩", "西小岩", "南小岩", "北小岩"] }
  ];

  for (const rule of fallbackRules) {
    if (rule.keywords.some(keyword => a.includes(keyword))) {
      return rule.name;
    }
  }

  return "";
}

function classifyAreaByLatLng(lat, lng) {
  if (!isValidLatLng(lat, lng)) return "周辺";

  if (lng >= 139.99) {
    if (lat >= 35.75) return "葛飾東部";
    if (lat >= 35.71) return "江戸川・小岩";
    return "東京東部";
  }

  if (lat >= 35.69 && lat < 35.74 && lng >= 139.93 && lng < 139.99) {
    return "船橋西部";
  }

  if (lat >= 35.71 && lat < 35.76 && lng >= 139.89 && lng < 139.96) {
    return "市川・本八幡";
  }

  if (lat >= 35.73 && lat < 35.78 && lng >= 139.87 && lng < 139.93) {
    return "市川北部";
  }

  if (lat >= 35.81 && lat < 35.86 && lng >= 139.90 && lng < 139.96) {
    return "新松戸";
  }

  if (lat >= 35.80 && lat < 35.84 && lng >= 139.90 && lng < 139.95) {
    return "北松戸・馬橋";
  }

  if (lat >= 35.79 && lat < 35.83 && lng >= 139.85 && lng < 139.91) {
    return "常盤平";
  }

  if (lat >= 35.76 && lat < 35.81 && lng >= 139.89 && lng < 139.95) {
    return "八柱・みのり台";
  }

  if (lat >= 35.75 && lat < 35.79 && lng >= 139.93 && lng < 139.98) {
    return "東松戸・秋山";
  }

  if (lat >= 35.74 && lat < 35.78 && lng >= 139.86 && lng < 139.91) {
    return "北国分・矢切";
  }

  if (lat >= 35.84 && lng >= 139.94 && lng < 140.02) return "柏";
  if (lat >= 35.87 && lng >= 139.90 && lng < 139.98) return "流山";
  if (lat >= 35.86 && lng < 139.90) return "我孫子";

  if (lat >= 35.75 && lat < 35.79 && lng >= 139.87 && lng < 139.93) {
    return "鎌ケ谷";
  }

  return "松戸周辺";
}

function guessArea(lat, lng, address = "") {
  const byAddress = classifyAreaByAddress(address);
  if (byAddress) return byAddress;

  return classifyAreaByLatLng(lat, lng);
}

function getDisplayArea(cast) {
  const address = cast.address || "";
  const byAddress = classifyAreaByAddress(address);
  if (byAddress) return byAddress;

  const lat = toNullableNumber(cast.latitude);
  const lng = toNullableNumber(cast.longitude);

  if (lat !== null && lng !== null) {
    return classifyAreaByLatLng(lat, lng);
  }

  return cast.area || "";
}

function getUsedCastIdsInCurrentDispatch() {
  const ids = new Set();
  currentDispatchItemsCache.forEach(item => {
    if (item.cast_id) ids.add(Number(item.cast_id));
  });
  return ids;
}

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

  const autoArea =
    classifyAreaByAddress(els.castAddress.value) ||
    (lat !== null && lng !== null ? classifyAreaByLatLng(lat, lng) : "");

  const payload = {
    name,
    phone: els.castPhone.value.trim(),
    address: els.castAddress.value.trim(),
    area: autoArea || els.castArea.value.trim() || null,
    latitude: lat,
    longitude: lng,
    memo: els.castMemo.value.trim(),
    is_active: true
  };

  let error;

  if (editingCastId) {
    ({ error } = await supabaseClient
      .from("casts")
      .update(payload)
      .eq("id", editingCastId));
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
  if (!els.castsList) return;

  els.castsList.innerHTML = "";

  if (!casts.length) {
    els.castsList.innerHTML = `<div class="item"><p>キャストがまだ登録されていません。</p></div>`;
    return;
  }

  casts.forEach(cast => {
    const lat = toNullableNumber(cast.latitude);
    const lng = toNullableNumber(cast.longitude);
    const km =
      lat !== null && lng !== null
        ? estimateRoadKmFromStation(lat, lng)
        : null;

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

/* =========================
   車両管理
========================= */

function resetVehicleForm() {
  editingVehicleId = null;
  if (els.vehiclePlateNumber) els.vehiclePlateNumber.value = "";
  if (els.vehicleName) els.vehicleName.value = "";
  if (els.vehicleDriverName) els.vehicleDriverName.value = "";
  if (els.vehicleStatus) els.vehicleStatus.value = "waiting";
  if (els.vehicleCapacityNote) els.vehicleCapacityNote.value = "";
  if (els.vehicleMemo) els.vehicleMemo.value = "";
  if (els.saveVehicleBtn) els.saveVehicleBtn.textContent = "保存";
  if (els.cancelVehicleEditBtn) els.cancelVehicleEditBtn.classList.add("hidden");
}

function fillVehicleForm(vehicle) {
  editingVehicleId = vehicle.id;
  els.vehiclePlateNumber.value = vehicle.plate_number || "";
  els.vehicleName.value = vehicle.vehicle_name || "";
  els.vehicleDriverName.value = vehicle.driver_name || "";
  els.vehicleStatus.value = vehicle.status || "waiting";
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
    status: els.vehicleStatus.value,
    capacity_note: els.vehicleCapacityNote.value.trim(),
    memo: els.vehicleMemo.value.trim(),
    is_active: true
  };

  let error;

  if (editingVehicleId) {
    ({ error } = await supabaseClient
      .from("vehicles")
      .update(payload)
      .eq("id", editingVehicleId));
  } else {
    payload.created_by = currentUser.id;
    ({ error } = await supabaseClient
      .from("vehicles")
      .insert(payload));
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
}

function renderVehicleSelect(vehicles) {
  if (!els.vehicleSelect) return;

  els.vehicleSelect.innerHTML = `<option value="">選択してください</option>`;

  vehicles.forEach(vehicle => {
    const option = document.createElement("option");
    option.value = vehicle.id;
    option.textContent =
      `${vehicle.plate_number} | ${vehicle.vehicle_name || "車種未設定"} | ${vehicleStatusLabel(vehicle.status)}`;
    option.dataset.label =
      `${vehicle.plate_number} ${vehicle.vehicle_name || ""}`.trim();
    els.vehicleSelect.appendChild(option);
  });
}

function renderVehicleList(vehicles) {
  if (!els.vehiclesList) return;

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
      <p>積載メモ: ${escapeHtml(vehicle.capacity_note || "-")}</p>
      <p>状態: <span class="badge ${badgeClass}">${escapeHtml(vehicleStatusLabel(vehicle.status))}</span></p>
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
    if (els.vehicleSelect) {
      els.vehicleSelect.value = existing[0].vehicle_id ? String(existing[0].vehicle_id) : "";
    }
    await loadDispatchItems(currentDispatchId);
    return;
  }

  const selectedVehicle = els.vehicleSelect?.selectedOptions?.[0] || null;
  const vehicleId = els.vehicleSelect?.value ? Number(els.vehicleSelect.value) : null;
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
    if (els.vehicleSelect) {
      els.vehicleSelect.value = dispatch.vehicle_id ? String(dispatch.vehicle_id) : "";
    }
    await loadDispatchItems(dispatch.id);
  } else {
    currentDispatchItemsCache = [];
    renderCastSelect(allCastsCache);
    els.dispatchSummary.textContent = "";
    els.dispatchList.innerHTML = `<div class="item"><p>この日の配車はまだありません。</p></div>`;
    if (els.vehicleSelect) {
      els.vehicleSelect.value = "";
    }
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

  const alreadyExists = currentDispatchItemsCache.some(
    item => Number(item.cast_id) === castId
  );
  if (alreadyExists) {
    alert("このキャストはすでに配車に追加されています");
    return;
  }

  const selected = els.castSelect.selectedOptions[0];
  const destinationAddress = selected?.dataset.address || "";
  const lat = toNullableNumber(selected?.dataset.lat);
  const lng = toNullableNumber(selected?.dataset.lng);
  const area =
    selected?.dataset.area ||
    (lat !== null && lng !== null ? guessArea(lat, lng, destinationAddress) : "");
  const distanceKm =
    lat !== null && lng !== null ? estimateRoadKmFromStation(lat, lng) : null;
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
      destination_area: area,
      latitude: lat,
      longitude: lng,
      distance_km: distanceKm,
      travel_minutes: travelMinutes,
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

function renderDispatchItems(items) {
  if (!els.dispatchList) return;

  els.dispatchList.innerHTML = "";

  if (!items.length) {
    els.dispatchList.innerHTML = `<div class="item"><p>配車先はまだありません。</p></div>`;
    return;
  }

  items.forEach((item, index) => {
    const status = normalizeStatus(item.status);
    const badgeClass = status === "done" ? "done" : status === "cancel" ? "cancel" : "";
    const cardClass = status === "done" ? "done-card" : status === "cancel" ? "cancel-card" : "";

    const isDone = status === "done";
    const isCancel = status === "cancel";
    const disableDone = isDone ? "disabled" : "";
    const disableCancel = isCancel ? "disabled" : "";

    const displayArea =
      item.destination_area ||
      (item.casts ? getDisplayArea(item.casts) : "") ||
      "-";

    const div = document.createElement("div");
    div.className = `item ${cardClass}`;
    div.innerHTML = `
      <h3>${item.stop_order}件目 | ${escapeHtml(item.casts?.name || "不明")}</h3>
      <p>地域: ${escapeHtml(displayArea)}</p>
      <p>住所: ${escapeHtml(item.destination_address || "-")}</p>
      <p>推定距離: ${item.distance_km ?? "-"} km</p>
      <p>推定時間: ${item.travel_minutes ?? "-"} 分</p>
      <p><span class="badge ${badgeClass}">${escapeHtml(status)}</span></p>
      <div class="actions">
        <button class="btn secondary route-btn" data-address="${escapeHtml(item.destination_address || "")}">ルート</button>
        <button class="btn secondary up-btn" data-id="${item.id}" ${index === 0 ? "disabled" : ""}>上へ</button>
        <button class="btn secondary down-btn" data-id="${item.id}" ${index === items.length - 1 ? "disabled" : ""}>下へ</button>
        <button class="btn done-btn" data-id="${item.id}" ${disableDone}>完了</button>
        <button class="btn danger cancel-btn" data-id="${item.id}" ${disableCancel}>キャンセル</button>
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

    els.dispatchList.appendChild(div);
  });
}

async function updateItemStatus(itemId, status) {
  const { error } = await supabaseClient
    .from("dispatch_items")
    .update({ status })
    .eq("id", itemId);

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(currentDispatchId, itemId, "update_status", `状態を ${status} に変更`);
  await loadDispatchItems(currentDispatchId);
  await loadHistory();
}

async function deleteDispatchItem(itemId) {
  const ok = window.confirm("この配車項目を削除しますか？");
  if (!ok) return;

  const { error } = await supabaseClient
    .from("dispatch_items")
    .delete()
    .eq("id", itemId);

  if (error) {
    alert(error.message);
    return;
  }

  await normalizeStopOrders(currentDispatchId);
  await addHistory(currentDispatchId, itemId, "delete_dispatch_item", "配車項目を削除");
  await loadDispatchItems(currentDispatchId);
  await loadHistory();
}

async function moveDispatchItem(itemId, direction) {
  const items = [...currentDispatchItemsCache].sort(
    (a, b) => Number(a.stop_order) - Number(b.stop_order)
  );

  const index = items.findIndex(item => Number(item.id) === Number(itemId));
  if (index < 0) return;

  const targetIndex = index + direction;
  if (targetIndex < 0 || targetIndex >= items.length) return;

  const temp = items[index];
  items[index] = items[targetIndex];
  items[targetIndex] = temp;

  for (let i = 0; i < items.length; i++) {
    const stopOrder = i + 1;
    await supabaseClient
      .from("dispatch_items")
      .update({ stop_order: stopOrder })
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
  if (!els.historyList) return;

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
      <p>${new Date(row.created_at).toLocaleString("ja-JP")}</p>
    `;
    els.historyList.appendChild(div);
  });
}

/* =========================
   AI並び替え
========================= */

function optimizeDispatchOrder(items) {
  const pending = items.filter(x => normalizeStatus(x.status) === "pending");
  const doneOrCancel = items.filter(x => normalizeStatus(x.status) !== "pending");

  const grouped = new Map();

  pending.forEach(item => {
    const key = item.destination_area || "未分類";
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(item);
  });

  const orderedGroups = [...grouped.entries()].sort((a, b) => {
    const aMin = Math.min(...a[1].map(x => Number(x.distance_km || 9999)));
    const bMin = Math.min(...b[1].map(x => Number(x.distance_km || 9999)));
    return aMin - bMin;
  });

  const result = [];
  orderedGroups.forEach(([, list]) => {
    list.sort((a, b) => {
      const aKm = Number(a.distance_km || 9999);
      const bKm = Number(b.distance_km || 9999);

      if (aKm !== bKm) return aKm - bKm;

      const aLat = Number(a.latitude || 0);
      const bLat = Number(b.latitude || 0);
      return bLat - aLat;
    });

    result.push(...list);
  });

  return [...result, ...doneOrCancel];
}

async function runOptimize() {
  if (!currentDispatchId) {
    alert("先に配車を読み込んでください");
    return;
  }

  const { data, error } = await supabaseClient
    .from("dispatch_items")
    .select("*")
    .eq("dispatch_id", currentDispatchId);

  if (error) {
    alert(error.message);
    return;
  }

  const optimized = optimizeDispatchOrder(data || []);

  for (let i = 0; i < optimized.length; i++) {
    await supabaseClient
      .from("dispatch_items")
      .update({ stop_order: i + 1 })
      .eq("id", optimized[i].id);
  }

  await addHistory(currentDispatchId, null, "optimize", "AI配車順を最適化");
  await loadDispatchItems(currentDispatchId);
  await loadHistory();
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

  const existingNames = new Set(
    allCastsCache.map(c => String(c.name || "").trim().toLowerCase())
  );

  const payload = rows
    .map(row => {
      const lat = toNullableNumber(row.latitude);
      const lng = toNullableNumber(row.longitude);

      const autoArea =
        classifyAreaByAddress(row.address) ||
        (lat !== null && lng !== null ? classifyAreaByLatLng(lat, lng) : "");

      return {
        name: String(row.name || "").trim(),
        phone: String(row.phone || "").trim(),
        address: String(row.address || "").trim(),
        area: autoArea || String(row.area || "").trim() || null,
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

  if (!tabs.length || !panels.length) return;

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

  els.openGoogleMapBtn?.addEventListener("click", () => {
    openGoogleMapsForPin(els.castAddress.value);
  });

  els.applyLatLngBtn?.addEventListener("click", applyLatLngFromText);

  els.saveCastBtn?.addEventListener("click", saveCast);
  els.cancelEditBtn?.addEventListener("click", resetCastForm);
  els.importCsvBtn?.addEventListener("click", importCsv);

  els.saveVehicleBtn?.addEventListener("click", saveVehicle);
  els.cancelVehicleEditBtn?.addEventListener("click", resetVehicleForm);

  els.createDispatchBtn?.addEventListener("click", createDispatch);
  els.loadDispatchBtn?.addEventListener("click", async () => {
    await loadDispatchByDate(els.dispatchDate.value);
  });
  els.addCastBtn?.addEventListener("click", addCastToDispatch);
  els.optimizeBtn?.addEventListener("click", async () => {
    await runOptimize();
  });

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
}

document.addEventListener("DOMContentLoaded", async () => {
  try {
    const ok = await ensureAuth();
    if (!ok) return;

    setupTabs();
    setupEvents();

    resetCastForm();
    resetVehicleForm();

    if (els.dispatchDate) {
      els.dispatchDate.value = todayStr();
    }

    await loadCasts();
    await loadVehicles();
    await loadDispatchByDate(els.dispatchDate.value);
    await loadHistory();
  } catch (err) {
    console.error("dashboard init error:", err);
    alert("初期化中にエラーが発生しました。F12 の Console を確認してください。");
  }
});
