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
  mainView: document.getElementById("mainView"),
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
  loadTodayDispatchBtn: document.getElementById("loadTodayDispatchBtn"),
  castSelect: document.getElementById("castSelect"),
  stopOrder: document.getElementById("stopOrder"),
  addCastToDispatchBtn: document.getElementById("addCastToDispatchBtn"),
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

function formatDateInputToday() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

async function signUp() {
  const email = els.loginEmail.value.trim();
  const password = els.loginPassword.value.trim();

  if (!email || !password) {
    els.authMessage.textContent = "メールアドレスとパスワードを入力してください。";
    return;
  }

  const { data, error } = await supabaseClient.auth.signUp({
    email,
    password
  });

  if (error) {
    els.authMessage.textContent = error.message;
    return;
  }

  if (data.user) {
    await supabaseClient.from("profiles").upsert({
      id: data.user.id,
      email: data.user.email,
      display_name: data.user.email,
      role: "dispatcher"
    });
  }

  els.authMessage.textContent = "登録しました。メール確認が必要な設定の場合は確認してください。";
}

async function signIn() {
  const email = els.loginEmail.value.trim();
  const password = els.loginPassword.value.trim();

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

function showLoggedOut() {
  els.loginView.classList.remove("hidden");
  els.mainView.classList.add("hidden");
  els.logoutBtn.classList.add("hidden");
  els.userEmail.textContent = "";
}

function showLoggedIn(user) {
  els.loginView.classList.add("hidden");
  els.mainView.classList.remove("hidden");
  els.logoutBtn.classList.remove("hidden");
  els.userEmail.textContent = user?.email || "";
}

async function onAuthed() {
  showLoggedIn(currentUser);
  els.dispatchDate.value = formatDateInputToday();
  await loadCasts();
  await loadDispatchByDate(els.dispatchDate.value);
  await loadHistory();
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

async function geocodeAddress(address) {
  const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${GOOGLE_MAPS_API_KEY}`;
  const res = await fetch(url);
  const json = await res.json();

  if (!json.results || !json.results.length) {
    throw new Error("住所から座標を取得できませんでした。");
  }

  const loc = json.results[0].geometry.location;
  return { lat: loc.lat, lng: loc.lng };
}

function openRouteFromMatsudo(address) {
  const origin = encodeURIComponent(ORIGIN_LABEL);
  const dest = encodeURIComponent(address);
  const url = `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${dest}&travelmode=driving`;
  window.open(url, "_blank");
}

function guessAreaByLatLng(lat, lng) {
  if (lat > 35.86) return "柏";
  if (lat > 35.79) return "松戸";
  return "市川・東京方面";
}

async function saveCast() {
  if (!currentUser) return;

  const name = els.castName.value.trim();
  if (!name) {
    alert("名前を入力してください");
    return;
  }

  const payload = {
    name,
    phone: els.castPhone.value.trim(),
    address: els.castAddress.value.trim(),
    area: els.castArea.value.trim(),
    latitude: els.castLat.value ? Number(els.castLat.value) : null,
    longitude: els.castLng.value ? Number(els.castLng.value) : null,
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

function clearCastForm() {
  els.castName.value = "";
  els.castPhone.value = "";
  els.castAddress.value = "";
  els.castArea.value = "";
  els.castLat.value = "";
  els.castLng.value = "";
  els.castMemo.value = "";
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
  renderCastsList(data || []);
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
    els.castSelect.appendChild(option);
  });
}

function renderCastsList(casts) {
  els.castsList.innerHTML = "";
  casts.forEach(cast => {
    const div = document.createElement("div");
    div.className = "item";
    div.innerHTML = `
      <h3>${escapeHtml(cast.name)}</h3>
      <p>電話: ${escapeHtml(cast.phone || "-")}</p>
      <p>住所: ${escapeHtml(cast.address || "-")}</p>
      <p>方面: ${escapeHtml(cast.area || "-")}</p>
      <p>座標: ${cast.latitude ?? "-"}, ${cast.longitude ?? "-"}</p>
      <p>メモ: ${escapeHtml(cast.memo || "-")}</p>
      <div class="actions">
        <button class="btn secondary" data-route="${escapeAttr(cast.address || "")}">GoogleMap</button>
      </div>
    `;
    div.querySelector("[data-route]")?.addEventListener("click", e => {
      const address = e.currentTarget.dataset.route;
      if (address) openRouteFromMatsudo(address);
    });
    els.castsList.appendChild(div);
  });
}

async function createDispatch() {
  const dispatch_date = els.dispatchDate.value;
  const driver_name = els.driverName.value.trim();

  if (!dispatch_date) {
    alert("日付を選択してください");
    return;
  }

  const { data, error } = await supabaseClient
    .from("dispatches")
    .insert({
      dispatch_date,
      driver_name,
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
  await addHistory(data.id, null, "create_dispatch", `配車を作成しました: ${dispatch_date}`);
  await loadDispatchByDate(dispatch_date);
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
  const destination_address = selected?.dataset.address || "";
  const latitude = selected?.dataset.lat ? Number(selected.dataset.lat) : null;
  const longitude = selected?.dataset.lng ? Number(selected.dataset.lng) : null;

  const stop_order = Number(els.stopOrder.value || 1);

  const { data, error } = await supabaseClient
    .from("dispatch_items")
    .insert({
      dispatch_id: currentDispatchId,
      cast_id: castId,
      stop_order,
      pickup_label: ORIGIN_LABEL,
      destination_address,
      latitude,
      longitude,
      status: "pending"
    })
    .select()
    .single();

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(currentDispatchId, data.id, "add_cast", `配車にキャストを追加しました`);
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

  renderDispatchItems(data || []);
}

function renderDispatchItems(items) {
  els.dispatchList.innerHTML = "";

  if (!items.length) {
    els.dispatchList.innerHTML = `<div class="item"><p>配車先はまだありません。</p></div>`;
    return;
  }

  items.forEach(item => {
    const div = document.createElement("div");
    div.className = "item";
    div.innerHTML = `
      <h3>${item.stop_order}件目 | ${escapeHtml(item.casts?.name || "不明")}</h3>
      <p>方面: ${escapeHtml(item.casts?.area || "-")}</p>
      <p>住所: ${escapeHtml(item.destination_address || "-")}</p>
      <p>状態: ${escapeHtml(item.status || "pending")}</p>
      <div class="actions">
        <button class="btn secondary" data-route="${escapeAttr(item.destination_address || "")}">ルート</button>
        <button class="btn" data-complete="${item.id}">完了</button>
        <button class="btn danger" data-cancel="${item.id}">キャンセル</button>
      </div>
    `;

    div.querySelector("[data-route]")?.addEventListener("click", e => {
      const address = e.currentTarget.dataset.route;
      if (address) openRouteFromMatsudo(address);
    });

    div.querySelector("[data-complete]")?.addEventListener("click", async e => {
      await updateDispatchItemStatus(Number(e.currentTarget.dataset.complete), "done");
    });

    div.querySelector("[data-cancel]")?.addEventListener("click", async e => {
      await updateDispatchItemStatus(Number(e.currentTarget.dataset.cancel), "cancel");
    });

    els.dispatchList.appendChild(div);
  });
}

async function updateDispatchItemStatus(itemId, status) {
  const { data, error } = await supabaseClient
    .from("dispatch_items")
    .update({ status })
    .eq("id", itemId)
    .select()
    .single();

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(currentDispatchId, itemId, "update_status", `配車状態を ${status} に変更`);
  await loadDispatchItems(currentDispatchId);
  await loadHistory();
}

async function addHistory(dispatchId, itemId, action, message) {
  await supabaseClient.from("dispatch_history").insert({
    dispatch_id: dispatchId,
    item_id: itemId,
    action,
    message,
    acted_by: currentUser.id
  });
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
  (data || []).forEach(row => {
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

function setupTabs() {
  const tabs = document.querySelectorAll(".tab");
  const panels = document.querySelectorAll(".tab-panel");

  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      tabs.forEach(t => t.classList.remove("active"));
      panels.forEach(p => p.classList.remove("active"));

      tab.classList.add("active");
      document.getElementById(tab.dataset.tab)?.classList.add("active");
    });
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttr(value) {
  return escapeHtml(value);
}

function setupEvents() {
  els.loginBtn.addEventListener("click", signIn);
  els.signupBtn.addEventListener("click", signUp);
  els.logoutBtn.addEventListener("click", signOut);

  els.saveCastBtn.addEventListener("click", saveCast);

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
        els.castArea.value = guessAreaByLatLng(lat, lng);
      }
    } catch (err) {
      alert(err.message);
    }
  });

  els.createDispatchBtn.addEventListener("click", createDispatch);
  els.loadTodayDispatchBtn.addEventListener("click", async () => {
    await loadDispatchByDate(els.dispatchDate.value);
  });
  els.addCastToDispatchBtn.addEventListener("click", addCastToDispatch);
}

document.addEventListener("DOMContentLoaded", async () => {
  setupTabs();
  setupEvents();
  els.dispatchDate.value = formatDateInputToday();
  await bootstrapAuth();
});