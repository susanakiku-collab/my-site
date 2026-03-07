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

const els = {
  loginView: document.getElementById("loginView"),
  appView: document.getElementById("appView"),
  userEmail: document.getElementById("userEmail"),
  logoutBtn: document.getElementById("logoutBtn"),

  loginEmail: document.getElementById("loginEmail"),
  loginPassword: document.getElementById("loginPassword"),
  loginBtn: document.getElementById("loginBtn"),
  signupBtn: document.getElementById("signupBtn"),
  authMessage: document.getElementById("authMessage"),

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

function showLoggedOut() {
  els.loginView.classList.remove("hidden");
  els.appView.classList.add("hidden");
  els.logoutBtn.classList.add("hidden");
  els.userEmail.textContent = "";
}

function showLoggedIn(user) {
  els.loginView.classList.add("hidden");
  els.appView.classList.remove("hidden");
  els.logoutBtn.classList.remove("hidden");
  els.userEmail.textContent = user?.email || "";
}

async function signUp() {
  const email = els.loginEmail.value.trim();
  const password = els.loginPassword.value.trim();

  if (!email || !password) {
    els.authMessage.textContent = "メールアドレスとパスワードを入力してください。";
    return;
  }

  const { data, error } = await supabaseClient.auth.signUp({ email, password });

  if (error) {
    els.authMessage.textContent = error.message;
    return;
  }

  if (data.user) {
    const { error: profileError } = await supabaseClient.from("profiles").upsert({
      id: data.user.id,
      email: data.user.email,
      display_name: data.user.email,
      role: "dispatcher"
    });

    if (profileError) {
      console.error(profileError);
    }
  }

  els.authMessage.textContent = "登録しました。確認メールが届く設定なら承認してください。";
}

async function signIn() {
  const email = els.loginEmail.value.trim();
  const password = els.loginPassword.value.trim();

  if (!email || !password) {
    els.authMessage.textContent = "メールアドレスとパスワードを入力してください。";
    return;
  }

  const { data, error } = await supabaseClient.auth.signInWithPassword({
    email,
    password
  });

  if (error) {
    els.authMessage.textContent = error.message;
    return;
  }

  currentUser = data.user;
  els.authMessage.textContent = "";
  await onAuthed();
}

async function signOut() {
  await supabaseClient.auth.signOut();
  currentUser = null;
  currentDispatchId = null;
  showLoggedOut();
}

async function bootstrapAuth() {
  const { data } = await supabaseClient.auth.getUser();
  currentUser = data.user;

  if (currentUser) {
    await onAuthed();
  } else {
    showLoggedOut();
  }

  supabaseClient.auth.onAuthStateChange(async (_event, session) => {
    currentUser = session?.user || null;
    if (currentUser) {
      await onAuthed();
    } else {
      showLoggedOut();
    }
  });
}

async function onAuthed() {
  showLoggedIn(currentUser);
  els.dispatchDate.value = todayStr();
  await loadCasts();
  await loadDispatchByDate(els.dispatchDate.value);
  await loadHistory();
}

function clearCastForm() {
  els.castName.value = "";
  els.castPhone.value = "";
  els.castAddress.value = "";
  els.castArea.value = "";
  els.castLat.value = "";
  els.castLng.value = "";
  els.castMemo.value = "";
}

async function saveCast() {
  if (!currentUser) return;

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
    memo: els.castMemo.value.trim(),
    created_by: currentUser.id
  };

  const { error } = await supabaseClient.from("casts").insert(payload);

  if (error) {
    alert(error.message);
    return;
  }

  clearCastForm();
  await loadCasts();
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
      </div>
    `;

    div.querySelector(".route-btn")?.addEventListener("click", e => {
      const address = e.currentTarget.dataset.address;
      if (address) openRouteFromMatsudo(address);
    });

    els.castsList.appendChild(div);
  });
}

async function createDispatch() {
  if (!currentUser) return;

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
  if (!currentUser) return;

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
  els.loginBtn.addEventListener("click", signIn);
  els.signupBtn.addEventListener("click", signUp);
  els.logoutBtn.addEventListener("click", signOut);

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
  setupTabs();
  setupEvents();
  els.dispatchDate.value = todayStr();
  await bootstrapAuth();
});