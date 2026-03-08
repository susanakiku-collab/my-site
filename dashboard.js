/* =========================================================
   THEMIS AI Dispatch - dashboard.js 完全版
   前提:
   - dashboard.html に以下IDがある想定
     planDate, planHour,
     btnLoadPlans, btnAutoDispatch, btnImportDispatch,
     planList, vehicleBoard, statusMsg

   - window.SUPABASE_URL / window.SUPABASE_ANON_KEY
     または下の定数に直接設定
   - HTML側で supabase-js を読み込み済み
   ========================================================= */

// =========================
// Supabase
// =========================
const SUPABASE_URL =
  window.SUPABASE_URL ||
  "YOUR_SUPABASE_URL";

const SUPABASE_ANON_KEY =
  window.SUPABASE_ANON_KEY ||
  "YOUR_SUPABASE_ANON_KEY";

const supabase =
  window.supabaseClient ||
  window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// =========================
// State
// =========================
const state = {
  profile: null,
  vehicles: [],
  plans: [],
  autoAssignments: [],
  dispatches: [],
};

// =========================
// DOM
// =========================
const elPlanDate = document.getElementById("planDate");
const elPlanHour = document.getElementById("planHour");

const elBtnLoadPlans = document.getElementById("btnLoadPlans");
const elBtnAutoDispatch = document.getElementById("btnAutoDispatch");
const elBtnImportDispatch = document.getElementById("btnImportDispatch");

const elPlanList = document.getElementById("planList");
const elVehicleBoard = document.getElementById("vehicleBoard");
const elStatusMsg = document.getElementById("statusMsg");

// =========================
// Init
// =========================
document.addEventListener("DOMContentLoaded", async () => {
  try {
    setTodayIfEmpty();
    bindEvents();

    await loadProfile();
    await loadVehicles();
    await loadPlans();
  } catch (err) {
    console.error(err);
    showStatus(`初期化エラー: ${err.message}`, true);
  }
});

// =========================
// Event Bindings
// =========================
function bindEvents() {
  if (elBtnLoadPlans) {
    elBtnLoadPlans.addEventListener("click", async () => {
      await loadPlans();
    });
  }

  if (elBtnAutoDispatch) {
    elBtnAutoDispatch.addEventListener("click", async () => {
      await autoDispatch();
    });
  }

  if (elBtnImportDispatch) {
    elBtnImportDispatch.addEventListener("click", async () => {
      await importToActualDispatches();
    });
  }

  if (elPlanDate) {
    elPlanDate.addEventListener("change", async () => {
      await loadPlans();
    });
  }

  if (elPlanHour) {
    elPlanHour.addEventListener("change", async () => {
      await loadPlans();
    });
  }
}

// =========================
// Utility
// =========================
function setTodayIfEmpty() {
  if (!elPlanDate) return;
  if (elPlanDate.value) return;

  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  elPlanDate.value = `${y}-${m}-${d}`;
}

function showStatus(message, isError = false) {
  if (elStatusMsg) {
    elStatusMsg.textContent = message;
    elStatusMsg.style.color = isError ? "#ff7b7b" : "#9fe870";
  }
  console[isError ? "error" : "log"](message);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function toNumber(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function toText(value, fallback = "") {
  const s = String(value ?? "").trim();
  return s || fallback;
}

function normalizeHour(value) {
  if (value === null || value === undefined) return "";
  const v = String(value).trim();
  return v;
}

function normalizeArea(area) {
  const a = String(area ?? "").trim();
  if (!a) return "その他";

  if (a.includes("松戸")) return "松戸";
  if (a.includes("柏")) return "柏";
  if (a.includes("流山")) return "流山";
  if (a.includes("市川")) return "市川";
  if (a.includes("船橋")) return "船橋";
  if (a.includes("東京東")) return "東京東";
  if (a.includes("東京西")) return "東京西";
  if (a.includes("埼玉")) return "埼玉";
  if (a.includes("千葉北")) return "千葉北";
  if (a.includes("千葉南")) return "千葉南";

  return a || "その他";
}

function areaOrder(area) {
  const map = {
    松戸: 1,
    柏: 2,
    流山: 3,
    市川: 4,
    船橋: 5,
    東京東: 6,
    東京西: 7,
    埼玉: 8,
    千葉北: 9,
    千葉南: 10,
    その他: 99,
  };
  return map[area] ?? 99;
}

function statusLabel(status) {
  const s = String(status ?? "").toLowerCase();

  if (s === "pending") return "予定";
  if (s === "scheduled") return "配車済";
  if (s === "imported") return "取込済";
  if (s === "completed") return "完了";
  if (s === "cancelled") return "キャンセル";

  return status || "-";
}

function chunkBy(arr, getKey) {
  const map = {};
  for (const item of arr) {
    const key = getKey(item);
    if (!map[key]) map[key] = [];
    map[key].push(item);
  }
  return map;
}

function safeArray(value) {
  return Array.isArray(value) ? value : [];
}

function uid(prefix = "id") {
  return `${prefix}_${Math.random().toString(36).slice(2, 10)}`;
}

// =========================
// Profile
// =========================
async function loadProfile() {
  try {
    const { data: authData, error: authError } = await supabase.auth.getUser();
    if (authError) throw authError;

    const user = authData?.user;
    if (!user) {
      state.profile = null;
      showStatus("未ログイン状態です", true);
      return;
    }

    const { data, error } = await supabase
      .from("profiles")
      .select("*")
      .eq("id", user.id)
      .maybeSingle();

    if (error) throw error;

    state.profile = data || {
      id: user.id,
      full_name: "",
    };
  } catch (err) {
    console.error(err);
    state.profile = null;
    showStatus(`プロフィール取得エラー: ${err.message}`, true);
  }
}

// =========================
// Vehicles
// =========================
async function loadVehicles() {
  try {
    showStatus("車両を読み込み中...");

    let query = supabase
      .from("vehicles")
      .select("*");

    // display_order が存在しないケースに備えて try/catch しやすいよう分離
    const { data, error } = await query.order("display_order", { ascending: true });

    if (error) throw error;

    state.vehicles = safeArray(data).filter((v) => {
      if (typeof v.is_active === "boolean") return v.is_active;
      return true;
    });

    renderVehicleBoard([]);
    showStatus(`車両 ${state.vehicles.length} 台を読み込みました`);
  } catch (err) {
    console.error(err);

    // display_order がない場合の救済
    try {
      const { data, error } = await supabase
        .from("vehicles")
        .select("*");

      if (error) throw error;

      state.vehicles = safeArray(data).filter((v) => {
        if (typeof v.is_active === "boolean") return v.is_active;
        return true;
      });

      renderVehicleBoard([]);
      showStatus(`車両 ${state.vehicles.length} 台を読み込みました`);
    } catch (err2) {
      console.error(err2);
      state.vehicles = [];
      renderVehicleBoard([]);
      showStatus(`車両読込エラー: ${err2.message}`, true);
    }
  }
}

// =========================
// Plans
// =========================
async function loadPlans() {
  try {
    if (!elPlanDate) {
      throw new Error("planDate が見つかりません");
    }

    showStatus("予定表を読み込み中...");

    let query = supabase
      .from("dispatch_plans")
      .select(`
        id,
        plan_date,
        plan_hour,
        cast_id,
        destination_address,
        planned_area,
        distance_km,
        note,
        status,
        created_by,
        created_at,
        updated_at,
        vehicle_group,
        casts (
          id,
          name,
          area,
          phone
        )
      `)
      .eq("plan_date", elPlanDate.value)
      .order("plan_hour", { ascending: true })
      .order("created_at", { ascending: true });

    if (elPlanHour && elPlanHour.value) {
      query = query.eq("plan_hour", elPlanHour.value);
    }

    const { data, error } = await query;
    if (error) throw error;

    state.plans = safeArray(data).map((row) => ({
      id: row.id,
      plan_date: row.plan_date,
      plan_hour: normalizeHour(row.plan_hour),
      cast_id: row.cast_id,
      destination_address: toText(row.destination_address),
      planned_area: normalizeArea(row.planned_area),
      distance_km: toNumber(row.distance_km, 0),
      note: toText(row.note),
      status: toText(row.status, "pending"),
      created_by: row.created_by,
      created_at: row.created_at,
      updated_at: row.updated_at,
      vehicle_group: toText(row.vehicle_group),
      casts: row.casts || null,
    }));

    renderPlanList();
    renderVehicleBoard([]);
    state.autoAssignments = [];

    showStatus(`予定 ${state.plans.length} 件を読み込みました`);
  } catch (err) {
    console.error(err);
    state.plans = [];
    renderPlanList();
    renderVehicleBoard([]);
    showStatus(`予定読込エラー: ${err.message}`, true);
  }
}

function renderPlanList() {
  if (!elPlanList) return;

  if (!state.plans.length) {
    elPlanList.innerHTML = `
      <div class="empty">
        該当する予定はありません
      </div>
    `;
    return;
  }

  const html = state.plans.map((plan) => {
    const castName = plan.casts?.name || `cast:${plan.cast_id}`;
    const area = plan.planned_area || "その他";
    const status = plan.status || "pending";

    return `
      <div class="plan-card status-${escapeHtml(status)}" data-plan-id="${escapeHtml(plan.id)}">
        <div class="plan-top">
          <div class="plan-hour">${escapeHtml(plan.plan_hour || "-")}時</div>
          <div class="plan-cast">${escapeHtml(castName)}</div>
          <div class="plan-area">${escapeHtml(area)}</div>
        </div>

        <div class="plan-address">${escapeHtml(plan.destination_address || "")}</div>

        <div class="plan-meta">
          <span>距離: ${escapeHtml(String(plan.distance_km))}km</span>
          <span>状態: ${escapeHtml(statusLabel(status))}</span>
          <span>グループ: ${escapeHtml(plan.vehicle_group || "-")}</span>
        </div>

        ${
          plan.note
            ? `<div class="plan-note">備考: ${escapeHtml(plan.note)}</div>`
            : ""
        }
      </div>
    `;
  }).join("");

  elPlanList.innerHTML = html;
}

// =========================
// AI Auto Dispatch
// =========================
async function autoDispatch() {
  try {
    if (!state.plans.length) {
      showStatus("先に予定表を読み込んでください", true);
      return;
    }

    if (!state.vehicles.length) {
      showStatus("車両がありません。vehicles を確認してください", true);
      return;
    }

    showStatus("AI自動配車を実行中...");

    // 対象: pending / scheduled以外でも completed,cancelled,imported は除外
    const targetPlans = state.plans.filter((p) => {
      const s = String(p.status || "").toLowerCase();
      return !["completed", "cancelled", "imported"].includes(s);
    });

    if (!targetPlans.length) {
      showStatus("配車対象の予定がありません", true);
      return;
    }

    // 方面クラスタ + 時間帯 + 距離順
    const sortedPlans = [...targetPlans].sort((a, b) => {
      const h1 = toNumber(a.plan_hour, 0);
      const h2 = toNumber(b.plan_hour, 0);
      if (h1 !== h2) return h1 - h2;

      const ao1 = areaOrder(a.planned_area);
      const ao2 = areaOrder(b.planned_area);
      if (ao1 !== ao2) return ao1 - ao2;

      const d1 = toNumber(a.distance_km, 0);
      const d2 = toNumber(b.distance_km, 0);
      if (d1 !== d2) return d1 - d2;

      return String(a.id).localeCompare(String(b.id));
    });

    // hour + area 単位でまとめる
    const buckets = chunkBy(
      sortedPlans,
      (p) => `${p.plan_hour}__${p.planned_area}`
    );

    const bucketKeys = Object.keys(buckets).sort((a, b) => a.localeCompare(b));
    const assignments = [];

    // vehicle_group が一致すればそこへ優先割当
    let vehicleCursor = 0;

    for (const bucketKey of bucketKeys) {
      const bucketPlans = buckets[bucketKey];

      for (const plan of bucketPlans) {
        let selectedVehicle = null;

        // vehicle_group 優先
        if (plan.vehicle_group) {
          selectedVehicle = state.vehicles.find((v) => {
            const vg =
              String(v.group_name || v.vehicle_group || v.name || "")
                .trim()
                .toLowerCase();
            return vg && vg === String(plan.vehicle_group).trim().toLowerCase();
          });
        }

        // 見つからなければラウンドロビン
        if (!selectedVehicle) {
          selectedVehicle = state.vehicles[vehicleCursor % state.vehicles.length];
          vehicleCursor += 1;
        }

        assignments.push({
          assignment_id: uid("assign"),
          plan_id: plan.id,
          plan_date: plan.plan_date,
          plan_hour: plan.plan_hour,
          cast_id: plan.cast_id,
          cast_name: plan.casts?.name || "",
          destination_address: plan.destination_address,
          destination_area: plan.planned_area,
          distance_km: plan.distance_km,
          vehicle_id: selectedVehicle.id,
          vehicle_name:
            selectedVehicle.name ||
            selectedVehicle.plate_number ||
            "未設定車両",
          driver_name:
            selectedVehicle.driver_name ||
            selectedVehicle.default_driver_name ||
            "",
          stop_order: 0,
          source_status: plan.status,
        });
      }
    }

    // 車両ごとに stop_order ふりなおし
    const vehicleHourGroups = chunkBy(
      assignments,
      (a) => `${a.vehicle_id}__${a.plan_hour}`
    );

    for (const key of Object.keys(vehicleHourGroups)) {
      vehicleHourGroups[key]
        .sort((a, b) => {
          const ao1 = areaOrder(a.destination_area);
          const ao2 = areaOrder(b.destination_area);
          if (ao1 !== ao2) return ao1 - ao2;

          const d1 = toNumber(a.distance_km, 0);
          const d2 = toNumber(b.distance_km, 0);
          if (d1 !== d2) return d1 - d2;

          return String(a.plan_id).localeCompare(String(b.plan_id));
        })
        .forEach((item, idx) => {
          item.stop_order = idx + 1;
        });
    }

    state.autoAssignments = assignments;
    renderVehicleBoard(assignments);

    showStatus(`AI自動配車が完了しました。対象 ${assignments.length} 件`);
  } catch (err) {
    console.error(err);
    showStatus(`AI自動配車エラー: ${err.message}`, true);
  }
}

// =========================
// Import to dispatches / dispatch_items
// =========================
async function importToActualDispatches() {
  try {
    if (!state.autoAssignments.length) {
      showStatus("先にAI自動配車を実行してください", true);
      return;
    }

    if (!state.profile?.id) {
      showStatus("ログインユーザー情報が取得できていません", true);
      return;
    }

    showStatus("実際の送りへ取り込み中...");

    // dispatch 単位でまとめる
    const groups = chunkBy(
      state.autoAssignments,
      (a) => `${a.plan_date}__${a.plan_hour}__${a.vehicle_id}`
    );

    const createdDispatches = [];
    const importedPlanIds = [];

    for (const key of Object.keys(groups)) {
      const items = groups[key];
      if (!items.length) continue;

      const head = items[0];

      // すでに同じ便があるか確認
      let existingDispatch = null;
      {
        const { data, error } = await supabase
          .from("dispatches")
          .select("*")
          .eq("dispatch_date", head.plan_date)
          .eq("dispatch_hour", head.plan_hour)
          .eq("vehicle_id", head.vehicle_id)
          .limit(1)
          .maybeSingle();

        if (error) throw error;
        existingDispatch = data || null;
      }

      let dispatchId = null;

      if (existingDispatch) {
        dispatchId = existingDispatch.id;
      } else {
        const payload = {
          dispatch_date: head.plan_date,
          dispatch_hour: head.plan_hour,
          vehicle_id: head.vehicle_id,
          driver_name: head.driver_name || "",
          status: "scheduled",
          created_by: state.profile.id,
        };

        const { data, error } = await supabase
          .from("dispatches")
          .insert(payload)
          .select()
          .single();

        if (error) throw error;
        dispatchId = data.id;
        createdDispatches.push(data);
      }

      // 既存 item 取得
      const { data: existingItems, error: existingItemsError } = await supabase
        .from("dispatch_items")
        .select("dispatch_id, cast_id, plan_date, plan_hour, destination_address")
        .eq("dispatch_id", dispatchId);

      if (existingItemsError) throw existingItemsError;

      const existsSet = new Set(
        safeArray(existingItems).map(
          (x) =>
            `${x.dispatch_id}__${x.cast_id}__${x.plan_date}__${x.plan_hour}__${x.destination_address}`
        )
      );

      const itemsPayload = items
        .filter((item) => {
          const dupKey =
            `${dispatchId}__${item.cast_id}__${item.plan_date}__${item.plan_hour}__${item.destination_address}`;
          return !existsSet.has(dupKey);
        })
        .map((item) => ({
          dispatch_id: dispatchId,
          cast_id: item.cast_id,
          stop_order: item.stop_order,
          destination_address: item.destination_address,
          destination_area: item.destination_area,
          distance_km: item.distance_km,
          vehicle_id: item.vehicle_id,
          driver_name: item.driver_name || "",
          status: "scheduled",
          plan_hour: item.plan_hour,
          plan_date: item.plan_date,
        }));

      if (itemsPayload.length > 0) {
        const { error: itemsInsertError } = await supabase
          .from("dispatch_items")
          .insert(itemsPayload);

        if (itemsInsertError) throw itemsInsertError;
      }

      // plan status 更新対象
      for (const item of items) {
        importedPlanIds.push(item.plan_id);
      }
    }

    // dispatch_plans を imported に更新
    const uniquePlanIds = [...new Set(importedPlanIds)];
    if (uniquePlanIds.length > 0) {
      const { error: updateError } = await supabase
        .from("dispatch_plans")
        .update({
          status: "imported",
          updated_at: new Date().toISOString(),
        })
        .in("id", uniquePlanIds);

      if (updateError) throw updateError;
    }

    showStatus(
      `取り込み完了: 新規dispatch ${createdDispatches.length} 件 / 対象予定 ${uniquePlanIds.length} 件`
    );

    await loadPlans();
  } catch (err) {
    console.error(err);
    showStatus(`取り込みエラー: ${err.message}`, true);
  }
}

// =========================
// Vehicle Board Rendering
// =========================
function renderVehicleBoard(assignments = []) {
  if (!elVehicleBoard) return;

  if (!state.vehicles.length) {
    elVehicleBoard.innerHTML = `
      <div class="empty">車両データがありません</div>
    `;
    return;
  }

  const groupedByVehicle = {};
  for (const vehicle of state.vehicles) {
    groupedByVehicle[vehicle.id] = [];
  }

  for (const item of assignments) {
    if (!groupedByVehicle[item.vehicle_id]) {
      groupedByVehicle[item.vehicle_id] = [];
    }
    groupedByVehicle[item.vehicle_id].push(item);
  }

  const html = state.vehicles.map((vehicle) => {
    const items = safeArray(groupedByVehicle[vehicle.id]);
    const vehicleName =
      vehicle.name || vehicle.plate_number || "未設定車両";

    const bodyHtml = items.length
      ? items
          .sort((a, b) => {
            const h1 = toNumber(a.plan_hour, 0);
            const h2 = toNumber(b.plan_hour, 0);
            if (h1 !== h2) return h1 - h2;

            return toNumber(a.stop_order, 0) - toNumber(b.stop_order, 0);
          })
          .map((item) => {
            return `
              <div class="vehicle-stop">
                <div class="vehicle-stop-head">
                  <span class="badge-hour">${escapeHtml(String(item.plan_hour))}時</span>
                  <span class="badge-order">${escapeHtml(String(item.stop_order))}件目</span>
                </div>
                <div class="vehicle-stop-cast">${escapeHtml(item.cast_name || "")}</div>
                <div class="vehicle-stop-area">${escapeHtml(item.destination_area || "")}</div>
                <div class="vehicle-stop-address">${escapeHtml(item.destination_address || "")}</div>
                <div class="vehicle-stop-meta">距離 ${escapeHtml(String(item.distance_km || 0))}km</div>
              </div>
            `;
          })
          .join("")
      : `<div class="empty-mini">まだ配車なし</div>`;

    return `
      <div class="vehicle-card" data-vehicle-id="${escapeHtml(vehicle.id)}">
        <div class="vehicle-card-head">
          <div class="vehicle-name">${escapeHtml(vehicleName)}</div>
          <div class="vehicle-sub">${escapeHtml(vehicle.plate_number || "")}</div>
        </div>
        <div class="vehicle-card-body">
          ${bodyHtml}
        </div>
      </div>
    `;
  }).join("");

  elVehicleBoard.innerHTML = html;
}

// =========================
// Optional: 実際の便を読みたい時用
// =========================
async function loadDispatchesByDate() {
  try {
    if (!elPlanDate?.value) return [];

    const { data, error } = await supabase
      .from("dispatches")
      .select("*")
      .eq("dispatch_date", elPlanDate.value)
      .order("dispatch_hour", { ascending: true });

    if (error) throw error;

    state.dispatches = safeArray(data);
    return state.dispatches;
  } catch (err) {
    console.error(err);
    showStatus(`dispatches読込エラー: ${err.message}`, true);
    return [];
  }
}

// =========================
// Debug Helpers
// =========================
window.THEMIS = {
  state,
  loadProfile,
  loadVehicles,
  loadPlans,
  autoDispatch,
  importToActualDispatches,
  loadDispatchesByDate,
};
