const {
  SUPABASE_URL,
  SUPABASE_ANON_KEY,
  GOOGLE_MAPS_API_KEY,
  ORIGIN_LABEL,
  ORIGIN_LAT,
  ORIGIN_LNG
} = window.APP_CONFIG;

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

let currentUser = null;
let currentDispatchId = null;
let editingCastId = null;

const els = {
  userEmail: document.getElementById("userEmail"),
  logoutBtn: document.getElementById("logoutBtn"),

  dispatchDate: document.getElementById("dispatchDate"),
  driverName: document.getElementById("driverName"),
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
  castArea: document.getElementById("castArea"),
  castLat: document.getElementById("castLat"),
  castLng: document.getElementById("castLng"),
  castMemo: document.getElementById("castMemo"),
  geocodeBtn: document.getElementById("geocodeBtn"),
  saveCastBtn: document.getElementById("saveCastBtn"),
  cancelEditBtn: document.getElementById("cancelEditBtn"),
  csvFileInput: document.getElementById("csvFileInput"),
  importCsvBtn: document.getElementById("importCsvBtn"),
  castsList: document.getElementById("castsList"),

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

function openRouteFromMatsudo(address) {
  const origin = encodeURIComponent(ORIGIN_LABEL);
  const dest = encodeURIComponent(address);
  const url = `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${dest}&travelmode=driving`;
  window.open(url, "_blank");
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
  if (lat >= 35.88) return "柏方面";
  if (lat >= 35.79) return "松戸方面";
  if (lng >= 139.95) return "市川・東京方面";
  return "周辺方面";
}

function normalizeStatus(status) {
  if (status === "done") return "done";
  if (status === "cancel") return "cancel";
  return "pending";
}

async function geocodeAddress(address) {
  if (!GOOGLE_MAPS_API_KEY) {
    throw new Error("Google Maps APIキーが未設定です");
  }

  const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${GOOGLE_MAPS_API_KEY}`;
  const res = await fetch(url);
  const json = await res.json();

  if (!json.results || !json.results.length) {
    throw new Error("住所から座標を取得できませんでした");
  }

  const loc = json.results[0].geometry.location;
  return { lat: loc.lat, lng: loc.lng };
}

async function ensureAuth() {
  const { data, error } = await supabaseClient.auth.getUser();

  if (error) {
    console.error(error);
    window.location.href = "index.html";
    return false;
  }

  currentUser = data.user;

  if (!currentUser) {
    window.location.href = "index.html";
    return false;
  }

  // profiles に自動登録
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
  }

  els.userEmail.textContent = currentUser.email || "";
  return true;
}

async function logout() {
  await supabaseClient.auth.signOut();
  window.location.href = "index.html";
}

function resetCastForm() {
  editingCastId = null;
  els.castName.value = "";
  els.castPhone.value = "";
  els.castAddress.value = "";
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
  els.castArea.value = cast.area || "";
  els.castLat.value = cast.latitude ?? "";
  els.castLng.value = cast.longitude ?? "";
  els.castMemo.value = cast.memo || "";
  els.saveCastBtn.textContent = "更新";
  els.cancelEditBtn.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function saveCast() {
  const name = els.castName.value.trim();
  if (!name) {
    alert("名前を入力してください");
    return;
  }

  const lat = els.castLat.value ? Number(els.castLat.value) : null;
  const lng = els.castLng.value ? Number(els.castLng.value) : null;

  const payload = {
    name,
    phone: els.castPhone.value.trim(),
    address: els.castAddress.value.trim(),
    area: els.castArea.value.trim() || (lat && lng ? guessArea(lat, lng) : null),
    latitude: lat,
    longitude: lng,
    memo: els.castMemo.value.trim()
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

  await addHistory(null, null, editingCastId ? "update_cast" : "create_cast", editingCastId ? "キャストを更新" : "キャストを作成");
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

  renderCastSelect(data || []);
  renderCastList(data || []);
}

function renderCastSelect(casts) {
  els.castSelect.innerHTML = `<option value="">選択してください</option>`;
  casts.forEach(cast => {
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
    const km = cast.latitude && cast.longitude
      ? estimateRoadKmFromStation(Number(cast.latitude), Number(cast.longitude))
      : null;

    const div = document.createElement("div");
    div.className = "item";
    div.innerHTML = `
      <h3>${escapeHtml(cast.name)}</h3>
      <p>電話: ${escapeHtml(cast.phone || "-")}</p>
      <p>住所: ${escapeHtml(cast.address || "-")}</p>
      <p>方面: ${escapeHtml(cast.area || "-")}</p>
      <p>松戸駅から推定距離: ${km ? `${km} km` : "-"}</p>
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
    await loadDispatchItems(currentDispatchId);
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
    await loadDispatchItems(dispatch.id);
  } else {
    els.dispatchSummary.textContent = "";
    els.dispatchList.innerHTML = `<div class="item"><p>この日の配車はまだありません。</p></div>`;
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

  const selected = els.castSelect.selectedOptions[0];
  const destinationAddress = selected?.dataset.address || "";
  const lat = selected?.dataset.lat ? Number(selected.dataset.lat) : null;
  const lng = selected?.dataset.lng ? Number(selected.dataset.lng) : null;
  const area = selected?.dataset.area || "";
  const distanceKm = lat && lng ? estimateRoadKmFromStation(lat, lng) : null;
  const travelMinutes = distanceKm ? estimateMinutes(distanceKm) : null;
  const stopOrder = Number(els.stopOrder.value || 1);

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

  await addHistory(currentDispatchId, data.id, "add_cast", "キャストを配車に追加");
  await loadDispatchItems(currentDispatchId);
  await loadHistory();
}

async function loadDispatchItems(dispatchId) {
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
    .order("stop_order", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  renderDispatchSummary(data || []);
  renderDispatchItems(data || []);
}

function renderDispatchSummary(items) {
  if (!items.length) {
    els.dispatchSummary.textContent = "配車先はまだありません。";
    return;
  }

  const active = items.filter(x => normalizeStatus(x.status) === "pending");
  const totalKm = active.reduce((sum, x) => sum + Number(x.distance_km || 0), 0);
  const totalMin = active.reduce((sum, x) => sum + Number(x.travel_minutes || 0), 0);

  els.dispatchSummary.textContent =
    `未完了 ${active.length} 件 / 推定合計距離 ${totalKm.toFixed(1)} km / 推定合計時間 ${totalMin} 分`;
}

function renderDispatchItems(items) {
  els.dispatchList.innerHTML = "";

  if (!items.length) {
    els.dispatchList.innerHTML = `<div class="item"><p>配車先はまだありません。</p></div>`;
    return;
  }

  items.forEach(item => {
    const status = normalizeStatus(item.status);
    const badgeClass = status === "done" ? "done" : status === "cancel" ? "cancel" : "";
    const cardClass = status === "done" ? "done-card" : status === "cancel" ? "cancel-card" : "";

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
        <button class="btn done-btn" data-id="${item.id}">完了</button>
        <button class="btn danger cancel-btn" data-id="${item.id}">キャンセル</button>
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

function optimizeDispatchOrder(items) {
  const active = items.filter(x => normalizeStatus(x.status) === "pending");
  const grouped = new Map();

  active.forEach(item => {
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
    list.sort((a, b) => Number(a.distance_km || 9999) - Number(b.distance_km || 9999));
    result.push(...list);
  });

  return result;
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

function parseCsv(text) {
  const lines = text
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .filter(line => line.trim() !== "");

  if (!lines.length) return [];

  const headers = parseCsvLine(lines[0]).map(h => h.trim());
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

  const payload = rows.map(row => {
    const lat = row.latitude ? Number(row.latitude) : null;
    const lng = row.longitude ? Number(row.longitude) : null;

    return {
      name: row.name?.trim() || "",
      phone: row.phone?.trim() || "",
      address: row.address?.trim() || "",
      area: row.area?.trim() || (lat && lng ? guessArea(lat, lng) : null),
      latitude: lat,
      longitude: lng,
      memo: row.memo?.trim() || "",
      is_active: true,
      created_by: currentUser.id
    };
  }).filter(x => x.name);

  if (!payload.length) {
    alert("有効なname列がありません");
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
  els.logoutBtn.addEventListener("click", logout);

  els.geocodeBtn.addEventListener("click", async () => {
    const address = els.castAddress.value.trim();
    if (!address) {
      alert("住所を入力してください");
      return;
    }

    try {
      const { lat, lng } = await geocodeAddress(address);
      els.castLat.value = lat;
      els.castLng.value = lng;
      if (!els.castArea.value) {
        els.castArea.value = guessArea(lat, lng);
      }
    } catch (err) {
      alert(err.message);
    }
  });

  els.saveCastBtn.addEventListener("click", saveCast);
  els.cancelEditBtn.addEventListener("click", resetCastForm);
  els.importCsvBtn.addEventListener("click", importCsv);

  els.createDispatchBtn.addEventListener("click", createDispatch);
  els.loadDispatchBtn.addEventListener("click", async () => {
    await loadDispatchByDate(els.dispatchDate.value);
  });
  els.addCastBtn.addEventListener("click", addCastToDispatch);
  els.optimizeBtn.addEventListener("click", async () => {
    await runOptimize();
  });
}

document.addEventListener("DOMContentLoaded", async () => {
  const ok = await ensureAuth();
  if (!ok) return;

  setupTabs();
  setupEvents();
  resetCastForm();
  els.dispatchDate.value = todayStr();
  await loadCasts();
  await loadDispatchByDate(els.dispatchDate.value);
  await loadHistory();
});
