/* =========================================================
   THEMIS AI Dispatch
   dashboard.js  安定版
========================================================= */

/* =========================
   Supabase 初期化
========================= */

const SUPABASE_URL = "https://cgtkbroxqdamnirjhzxy.supabase.co"
const SUPABASE_ANON_KEY = "sb_publishable_veQ6nay4yBjAHz95Hv55ng_QahlWYUz"

const supabase = window.supabase.createClient(
  SUPABASE_URL,
  SUPABASE_ANON_KEY
)

/* =========================
   State
========================= */

const state = {
  profile: null,
  vehicles: [],
  plans: [],
  autoAssignments: []
};

/* =========================
   DOM
========================= */

const elPlanDate = document.getElementById("planDate")
const elPlanHour = document.getElementById("planHour")

const elBtnLoadPlans = document.getElementById("btnLoadPlans")
const elBtnAutoDispatch = document.getElementById("btnAutoDispatch")
const elBtnImportDispatch = document.getElementById("btnImportDispatch")

const elPlanList = document.getElementById("planList")
const elVehicleBoard = document.getElementById("vehicleBoard")
const elStatusMsg = document.getElementById("statusMsg")

/* =========================
   初期化
========================= */

document.addEventListener("DOMContentLoaded", async () => {

  setToday()

  bindEvents()

  try {
    await loadVehicles()
  } catch(e){
    console.error(e)
  }

  try {
    await loadPlans()
  } catch(e){
    console.error(e)
  }

})

/* =========================
   Utility
========================= */

function setToday(){

  if(!elPlanDate) return

  const now = new Date()

  const y = now.getFullYear()
  const m = String(now.getMonth()+1).padStart(2,"0")
  const d = String(now.getDate()).padStart(2,"0")

  elPlanDate.value = `${y}-${m}-${d}`

}

function showStatus(msg,error=false){

  console.log(msg)

  if(!elStatusMsg) return

  elStatusMsg.textContent = msg
  elStatusMsg.style.color = error ? "#ff7a7a" : "#8fffc1"

}

function num(v){
  const n = Number(v)
  return Number.isFinite(n) ? n : 0
}

function text(v){
  return String(v ?? "")
}

function normalizeArea(a){

  if(!a) return "その他"

  const v = String(a)

  if(v.includes("松戸")) return "松戸"
  if(v.includes("柏")) return "柏"
  if(v.includes("流山")) return "流山"
  if(v.includes("市川")) return "市川"
  if(v.includes("船橋")) return "船橋"

  return "その他"

}

function areaOrder(a){

  const map = {
    松戸:1,
    柏:2,
    流山:3,
    市川:4,
    船橋:5,
    その他:99
  }

  return map[a] ?? 99

}

/* =========================
   Event
========================= */

function bindEvents(){

  if(elBtnLoadPlans)
  elBtnLoadPlans.onclick = loadPlans

  if(elBtnAutoDispatch)
  elBtnAutoDispatch.onclick = autoDispatch

  if(elBtnImportDispatch)
  elBtnImportDispatch.onclick = importDispatch

}

/* =========================
   車両取得
========================= */

async function loadVehicles(){

  showStatus("車両読み込み中...")

  const {data,error} = await supabase
  .from("vehicles")
  .select("*")
  .order("id")

  if(error){
    showStatus("vehicles取得エラー",true)
    throw error
  }

  state.vehicles = data || []

  showStatus(`車両 ${state.vehicles.length} 台`)

}

/* =========================
   予定取得
========================= */

async function loadPlans(){

  showStatus("予定表読み込み中...")

  let query = supabase
  .from("dispatch_plans")
  .select("*")
  .eq("plan_date", elPlanDate.value)

  if(elPlanHour.value)
  query = query.eq("plan_hour", elPlanHour.value)

  const {data,error} = await query

  if(error){
    showStatus("予定取得エラー",true)
    throw error
  }

  state.plans = (data || []).map(p=>({

    ...p,

    planned_area: normalizeArea(p.planned_area),
    distance_km: num(p.distance_km)

  }))

  renderPlans()

  showStatus(`予定 ${state.plans.length} 件`)

}

/* =========================
   予定表示
========================= */

function renderPlans(){

  if(!elPlanList) return

  if(state.plans.length===0){

    elPlanList.innerHTML = `<div class="empty">予定なし</div>`
    return

  }

  elPlanList.innerHTML = state.plans.map(p=>`

  <div class="plan-card">

    <div class="plan-top">

      <div class="plan-hour">${p.plan_hour}時</div>

      <div class="plan-cast">${p.cast_id}</div>

      <div class="plan-area">${p.planned_area}</div>

    </div>

    <div class="plan-address">
      ${p.destination_address}
    </div>

    <div class="plan-meta">
      距離 ${p.distance_km}km
    </div>

  </div>

  `).join("")

}

/* =========================
   AI配車
========================= */

function autoDispatch(){

  if(state.plans.length===0){
    showStatus("予定がありません",true)
    return
  }

  if(state.vehicles.length===0){
    showStatus("車両がありません",true)
    return
  }

  showStatus("AI配車実行")

  const sorted = [...state.plans].sort((a,b)=>{

    const h = num(a.plan_hour)-num(b.plan_hour)
    if(h!==0) return h

    const ar = areaOrder(a.planned_area)-areaOrder(b.planned_area)
    if(ar!==0) return ar

    return a.distance_km-b.distance_km

  })

  const assignments = []

  let i=0

  for(const p of sorted){

    const vehicle = state.vehicles[i % state.vehicles.length]

    assignments.push({

      ...p,

      vehicle_id: vehicle.id,
      vehicle_name: vehicle.name,
      stop_order:1

    })

    i++

  }

  state.autoAssignments = assignments

  renderVehicleBoard()

  showStatus("AI配車完了")

}

/* =========================
   車両表示
========================= */

function renderVehicleBoard(){

  if(!elVehicleBoard) return

  if(state.autoAssignments.length===0){

    elVehicleBoard.innerHTML=`<div class="empty">配車なし</div>`
    return

  }

  const group={}

  for(const a of state.autoAssignments){

    if(!group[a.vehicle_id])
      group[a.vehicle_id]=[]

    group[a.vehicle_id].push(a)

  }

  elVehicleBoard.innerHTML = Object.keys(group).map(id=>{

    const list = group[id]

    return `

    <div class="vehicle-card">

      <div class="vehicle-card-head">
        車両 ${id}
      </div>

      <div class="vehicle-card-body">

      ${
        list.map(x=>`

        <div class="vehicle-stop">

          <div class="vehicle-stop-head">
            <span>${x.plan_hour}時</span>
          </div>

          <div class="vehicle-stop-area">
            ${x.planned_area}
          </div>

          <div class="vehicle-stop-address">
            ${x.destination_address}
          </div>

        </div>

        `).join("")
      }

      </div>

    </div>

    `

  }).join("")

}

/* =========================
   dispatch登録
========================= */

async function importDispatch(){

  if(state.autoAssignments.length===0){

    showStatus("配車データなし",true)
    return

  }

  showStatus("dispatch登録中")

  for(const a of state.autoAssignments){

    await supabase
    .from("dispatch_items")
    .insert({

      cast_id:a.cast_id,
      vehicle_id:a.vehicle_id,
      destination_address:a.destination_address,
      destination_area:a.planned_area,
      distance_km:a.distance_km,
      plan_hour:a.plan_hour,
      plan_date:a.plan_date,
      stop_order:a.stop_order,
      status:"scheduled"

    })

  }

  showStatus("dispatch登録完了")

}


