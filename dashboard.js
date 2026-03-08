// =========================
// THEMIS AI Dispatch Core
// =========================

// 既存で読み込み済み想定
// import { createClient } from 'https://esm.sh/@supabase/supabase-js'
const SUPABASE_URL = window.SUPABASE_URL || 'YOUR_SUPABASE_URL';
const SUPABASE_ANON_KEY = window.SUPABASE_ANON_KEY || 'YOUR_SUPABASE_ANON_KEY';
const supabase = window.supabaseClient || window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// -------------------------
// state
// -------------------------
const state = {
  plans: [],
  vehicles: [],
  autoAssignments: [],
  profile: null
};

// -------------------------
// DOM
// -------------------------
const elPlanDate = document.getElementById('planDate');
const elPlanHour = document.getElementById('planHour');
const elBtnLoadPlans = document.getElementById('btnLoadPlans');
const elBtnAutoDispatch = document.getElementById('btnAutoDispatch');
const elBtnImportDispatch = document.getElementById('btnImportDispatch');
const elPlanList = document.getElementById('planList');
const elVehicleBoard = document.getElementById('vehicleBoard');
const elStatusMsg = document.getElementById('statusMsg');

// -------------------------
// init
// -------------------------
document.addEventListener('DOMContentLoaded', async () => {
  setToday();
  bindEvents();
  await loadProfile();
  await loadVehicles();
  await loadPlans();
});

// -------------------------
// util
// -------------------------
function setToday() {
  if (!elPlanDate.value) {
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    elPlanDate.value = `${yyyy}-${mm}-${dd}`;
  }
}

function bindEvents() {
  elBtnLoadPlans?.addEventListener('click', loadPlans);
  elBtnAutoDispatch?.addEventListener('click', autoDispatch);
  elBtnImportDispatch?.addEventListener('click', importToActualDispatches);
  elPlanDate?.addEventListener('change', loadPlans);
  elPlanHour?.addEventListener('change', loadPlans);
}

function showStatus(msg, isError = false) {
  if (!elStatusMsg) return;
  elStatusMsg.textContent = msg;
  elStatusMsg.style.color = isError ? '#ff6b6b' : '#9fe870';
  console.log(msg);
}

function escapeHtml(str = '') {
  return String(str)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function areaOrder(area) {
  const map = {
    '松戸': 1,
    '柏': 2,
    '流山': 3,
    '市川': 4,
    '船橋': 5,
    '東京東': 6,
    '東京西': 7,
    '埼玉': 8,
    '千葉北': 9,
    '千葉南': 10,
    'その他': 99
  };
  return map[area] ?? 99;
}

function normalizeArea(area) {
  if (!area) return 'その他';
  const a = String(area).trim();

  if (a.includes('松戸')) return '松戸';
  if (a.includes('柏')) return '柏';
  if (a.includes('流山')) return '流山';
  if (a.includes('市川')) return '市川';
  if (a.includes('船橋')) return '船橋';
  if (a.includes('東京東')) return '東京東';
  if (a.includes('東京西')) return '東京西';
  if (a.includes('埼玉')) return '埼玉';
  if (a.includes('千葉北')) return '千葉北';
  if (a.includes('千葉南')) return '千葉南';

  return 'その他';
}

function toNumber(v, fallback = 0) {
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

// -------------------------
// auth/profile
// -------------------------
async function loadProfile() {
  try {
    const { data: authData, error: authError } = await supabase.auth.getUser();
    if (authError) throw authError;

    const user = authData?.user;
    if (!user) {
      showStatus('未ログイン状態です', true);
      return;
    }

    const { data, error } = await supabase
      .from('profiles')
      .select('*')
      .eq('id', user.id)
      .single();

    if (error) {
      console.warn('profiles取得失敗:', error.message);
      state.profile = { id: user.id, full_name: '未設定' };
      return;
    }

    state.profile = data;
  } catch (err) {
    console.error(err);
    showStatus(`プロフィール読込エラー: ${err.message}`, true);
  }
}

// -------------------------
// vehicles
// -------------------------
async function loadVehicles() {
  try {
    let query = supabase
      .from('vehicles')
      .select('*')
      .order('display_order', { ascending: true });

    const { data, error } = await query;

    if (error) throw error;

    state.vehicles = (data || []).filter(v => {
      if (typeof v.is_active === 'boolean') return v.is_active;
      return true;
    });

    renderVehicleBoard([]);
  } catch (err) {
    console.error(err);
    showStatus(`車両読込エラー: ${err.message}`, true);
  }
}

// -------------------------
// plans
// -------------------------
async function loadPlans() {
  try {
    showStatus('予定表を読み込み中...');

    let query = supabase
      .from('dispatch_plans')
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
      .eq('plan_date', elPlanDate.value)
      .order('plan_hour', { ascending: true })
      .order('created_at', { ascending: true });

    if (elPlanHour.value) {
      query = query.eq('plan_hour', elPlanHour.value);
    }

    const { data, error } = await query;

    if (error) throw error;

    state.plans = (data || []).map(plan => ({
      ...plan,
      planned_area: normalizeArea(plan.planned_area),
      distance_km: toNumber(plan.distance_km, 0)
    }));

    renderPlanList();
    renderVehicleBoard([]);
    showStatus(`予定 ${state.plans.length} 件を読み込みました`);
  } catch (err) {
    console.error(err);
    showStatus(`予定読込エラー: ${err.message}`, true);
  }
}

function renderPlanList() {
  if (!elPlanList) return;

  if (!state.plans.length) {
    elPlanList.innerHTML = `<div class="empty">該当する予定はありません</div>`;
    return;
  }

  elPlanList.innerHTML = state.plans.map(plan => {
    const castName = plan.casts?.name || `cast:${plan.cast_id}`;
    const area = plan.planned_area || 'その他';
    const status = plan.status || 'pending';

    return `
      <div class="plan-card status-${escapeHtml(status)}">
        <div class="plan-top">
          <div class="plan-hour">${escapeHtml(plan.plan_hour || '-')}時</div>
          <div class="plan-cast">${escapeHtml(castName)}</div>
          <div class="plan-area">${escapeHtml(area)}</div>
        </div>
        <div class="plan-address">${escapeHtml(plan.destination_address || '')}</div>
        <div class="plan-meta">
          <span>距離: ${escapeHtml(String(plan.distance_km || 0))} km</span>
          <span>状態: ${escapeHtml(status)}</span>
          <span>車両グループ: ${escapeHtml(plan.vehicle_group || '-')}</span>
        </div>
        ${plan.note ? `<div class="plan-note">備考: ${escapeHtml(plan.note)}</div>` : ''}
      </div>
    `;
  }).join('');
}

// -------------------------
// AI auto dispatch
// -------------------------
async function autoDispatch() {
  try {
    if (!state.plans.length) {
      showStatus('先に予定表を読み込んでください', true);
      return;
    }

    if (!state.vehicles.length) {
      showStatus('有効な車両がありません', true);
      return;
    }

    showStatus('AI自動配車を実行中...');

    // 未処理だけ対象
    const targetPlans = state.plans.filter(p => {
      const s = (p.status || '').toLowerCase();
      return !['imported', 'completed', 'cancelled'].includes(s);
    });

    if (!targetPlans.length) {
      showStatus('配車対象の予定がありません', true);
      return;
    }

    // 時間帯 → 方面 → 距離順
    const sorted = [...targetPlans].sort((a, b) => {
      const h1 = toNumber(a.plan_hour, 0);
      const h2 = toNumber(b.plan_hour, 0);
      if (h1 !== h2) return h1 - h2;

      const a1 = areaOrder(a.planned_area);
      const a2 = areaOrder(b.planned_area);
      if (a1 !== a2) return a1 - a2;

      return toNumber(a.distance_km, 0) - toNumber(b.distance_km, 0);
    });

    // 車両へラウンドロビン + 方面まとまり優先
    const buckets = {};
    for (const plan of sorted) {
      const key = `${plan.plan_hour}_${plan.planned_area}`;
      if (!buckets[key]) buckets[key] = [];
      buckets[key].push(plan);
    }

    const assignments = [];
    const vehicles = [...state.vehicles];
    let vehicleCursor = 0;

    Object.keys(buckets).sort().forEach(bucketKey => {
      const groupPlans = buckets[bucketKey];

      for (const plan of groupPlans) {
        const vehicle = vehicles[vehicleCursor % vehicles.length];
        assignments.push({
          plan_id: plan.id,
          plan_date: plan.plan_date,
          plan_hour: plan.plan_hour,
          cast_id: plan.cast_id,
          cast_name: plan.casts?.name || '',
          destination_address: plan.destination_address,
          destination_area: plan.planned_area,
          distance_km: plan.distance_km,
          vehicle_id: vehicle.id,
          vehicle_name: vehicle.name || vehicle.plate_number || `車両${vehicleCursor + 1}`,
          driver_name: vehicle.driver_name || '',
          stop_order: 0
        });
        vehicleCursor++;
      }
    });

    // 車両ごとに stop_order 採番
    const grouped = {};
    for (const item of assignments) {
      const key = `${item.vehicle_id}_${item.plan_hour}`;
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(item);
    }

    Object.values(grouped).forEach(items => {
      items
        .sort((a, b) => toNumber(a.distance_km, 0) - toNumber(b.distance_km, 0))
        .forEach((item, idx) => {
          item.stop_order = idx + 1;
        });
    });

    state.autoAssignments = assignments;
    renderVehicleBoard(assignments);
    showStatus(`AI自動配車が完了しました。対象 ${assignments.length} 件`);
  } catch (err) {
    console.error(err);
    showStatus(`AI自動配車エラー: ${err.message}`, true);
  }
}

// -------------------------
// import to dispatches / dispatch_items
// -------------------------
async function importToActualDispatches() {
  try {
    if (!state.autoAssignments.length) {
      showStatus('先にAI自動配車を実行してください', true);
      return;
    }

    showStatus('実際の送りへ取り込み中...');

    // dispatch単位でまとめる
    const dispatchGroupMap = {};
    for (const item of state.autoAssignments) {
      const key = `${item.plan_date}_${item.plan_hour}_${item.vehicle_id}`;
      if (!dispatchGroupMap[key]) {
        dispatchGroupMap[key] = {
          dispatch_date: item.plan_date,
          dispatch_hour: item.plan_hour,
          vehicle_id: item.vehicle_id,
          driver_name: item.driver_name || '',
          items: []
        };
      }
      dispatchGroupMap[key].items.push(item);
    }

    const createdDispatchIds = [];

    for (const groupKey of Object.keys(dispatchGroupMap)) {
      const group = dispatchGroupMap[groupKey];

      // 1. dispatches 作成
      const { data: dispatchData, error: dispatchError } = await supabase
        .from('dispatches')
        .insert({
          dispatch_date: group.dispatch_date,
          dispatch_hour: group.dispatch_hour,
          vehicle_id: group.vehicle_id,
          driver_name: group.driver_name,
          status: 'scheduled',
          created_by: state.profile?.id || null
        })
        .select()
        .single();

      if (dispatchError) throw dispatchError;

      createdDispatchIds.push(dispatchData.id);

      // 2. dispatch_items 作成
      const itemsPayload = group.items.map(item => ({
        dispatch_id: dispatchData.id,
        cast_id: item.cast_id,
        stop_order: item.stop_order,
        destination_address: item.destination_address,
        destination_area: item.destination_area,
        distance_km: item.distance_km,
        vehicle_id: item.vehicle_id,
        driver_name: item.driver_name,
        status: 'scheduled',
        plan_hour: item.plan_hour,
        plan_date: item.plan_date
      }));

      const { error: itemsError } = await supabase
        .from('dispatch_items')
        .insert(itemsPayload);

      if (itemsError) throw itemsError;
    }

    // 3. 元の予定を imported に更新
    const planIds = state.autoAssignments.map(x => x.plan_id);

    const { error: updateError } = await supabase
      .from('dispatch_plans')
      .update({
        status: 'imported',
        updated_at: new Date().toISOString()
      })
      .in('id', planIds);

    if (updateError) throw updateError;

    showStatus(`取り込み完了: dispatch ${createdDispatchIds.length} 件作成`);
    await loadPlans();
  } catch (err) {
    console.error(err);
    showStatus(`取り込みエラー: ${err.message}`, true);
  }
}

// -------------------------
// render vehicle board
// -------------------------
function renderVehicleBoard(assignments = []) {
  if (!elVehicleBoard) return;

  if (!state.vehicles.length) {
    elVehicleBoard.innerHTML = `<div class="empty">車両データがありません</div>`;
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

  elVehicleBoard.innerHTML = state.vehicles.map(vehicle => {
    const items = groupedByVehicle[vehicle.id] || [];
    const name = vehicle.name || vehicle.plate_number || '未設定車両';

    const itemHtml = items.length
      ? items
          .sort((a, b) => {
            const h1 = toNumber(a.plan_hour, 0);
            const h2 = toNumber(b.plan_hour, 0);
            if (h1 !== h2) return h1 - h2;
            return a.stop_order - b.stop_order;
          })
          .map(item => `
            <div class="vehicle-stop">
              <div class="vehicle-stop-head">
                <span class="badge-hour">${escapeHtml(String(item.plan_hour))}時</span>
                <span class="badge-order">${item.stop_order}件目</span>
              </div>
              <div class="vehicle-stop-cast">${escapeHtml(item.cast_name || '')}</div>
              <div class="vehicle-stop-area">${escapeHtml(item.destination_area || '')}</div>
              <div class="vehicle-stop-address">${escapeHtml(item.destination_address || '')}</div>
              <div class="vehicle-stop-meta">距離 ${escapeHtml(String(item.distance_km || 0))}km</div>
            </div>
          `)
          .join('')
      : `<div class="empty-mini">まだ配車なし</div>`;

    return `
      <div class="vehicle-card">
        <div class="vehicle-card-head">
          <div class="vehicle-name">${escapeHtml(name)}</div>
          <div class="vehicle-sub">${escapeHtml(vehicle.plate_number || '')}</div>
        </div>
        <div class="vehicle-card-body">
          ${itemHtml}
        </div>
      </div>
    `;
  }).join('');
}
