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

function guessArea(lat, lng) {
  if (!isValidLatLng(lat, lng)) return "周辺方面";
  if (lat >= 35.88) return "柏方面";
  if (lat >= 35.79) return "松戸方面";
  if (lng >= 139.95) return "市川・東京方面";
  return "周辺方面";
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
  els.castArea.value = cast.area || "";
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
    els.castArea.value = guessArea(parsed.lat, parsed.lng);
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
    area: els.castArea.value.trim() || (lat !== null && lng !== null ? guessArea(lat, lng) : null),
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
  els.castSelect.innerHTML = `<option value="">選択してください</option>`;

  const usedCastIds = getUsedCastIdsInCurrentDispatch();

  casts
    .filter(cast => !usedCastIds.has(Number(cast.id)))
    .forEach(cast => {
      const option = document.createElement("option");
      option.value = cast.id;
      option.textContent = `${cast.name} | ${cast.area || "方面未設定"}`;
      option.dataset.address = cast.address || "";
      option.dataset.lat = cast.latitude ?? "";
      option.dataset.lng = cast.longitude ?? "";
      option.dataset.area = cast.area || "";
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
    const km =
      lat !== null && lng !== null
        ? estimateRoadKmFromStation(lat, lng)
        : null;

    const div = document.createElement("div");
    div.className = "item";
    div.innerHTML = `
      <h3>${escapeHtml(cast.name)}</h3>
      <p>電話: ${escapeHtml(cast.phone || "-")}</p>
      <p>住所: ${escapeHtml(cast.address || "-")}</p>
      <p>方面: ${escapeHtml(cast.area || "-")}</p>
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
  els.vehiclePlateNumber.value = "";
  els.vehicleName.value = "";
  els.vehicleDriverName.value = "";
  els.vehicleStatus.value = "waiting";
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
    selected?.dataset.area || (lat !== null && lng !== null ? guessArea(lat, lng) : "");
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
        area
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

  const vehicleText =
    els.vehicleSelect && els.vehicleSelect.selectedOptions.length
      ? els.vehicleSelect.selectedOptions[0]?.textContent || ""
      : "";

  els.dispatchSummary.textContent =
    `未完了 ${active.length} 件 / 完了 ${done.length} 件 / キャンセル ${cancel.length} 件 / 推定合計距離 ${totalKm.toFixed(1)} km / 推定合計時間 ${totalMin} 分${vehicleText ? ` / 車両 ${vehicleText}` : ""}`;
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

    const isDone = status === "done";
    const isCancel = status === "cancel";
    const disableDone = isDone ? "disabled" : "";
    const disableCancel = isCancel ? "disabled" : "";

    const div = document.createElement("div");
    div.className = `item ${cardClass}`;
    div.innerHTML = `
      <h3>${item.stop_order}件目 | ${escapeHtml(item.casts?.name || "不明")}</h3>
      <p>方面: ${escapeHtml(item.destination_area || item.casts?.area || "-")}</p>
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

      return {
        name: String(row.name || "").trim(),
        phone: String(row.phone || "").trim(),
        address: String(row.address || "").trim(),
        area: String(row.area || "").trim() || (lat !== null && lng !== null ? guessArea(lat, lng) : null),
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
      document.getElementById(tab.dataset.tab)?.classList.add("active");
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
      els.castArea.value = guessArea(parsed.lat, parsed.lng);
    }
  });
}

document.addEventListener("DOMContentLoaded", async () => {
  const ok = await ensureAuth();
  if (!ok) return;

  setupTabs();
  setupEvents();

  resetCastForm();
  resetVehicleForm();

  els.dispatchDate.value = todayStr();

  await loadCasts();
  await loadVehicles();
  await loadDispatchByDate(els.dispatchDate.value);
  await loadHistory();
});
