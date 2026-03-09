/* =====================================================
THEMIS AI Dispatch
dashboard.js
PART 1 / 9
設定 / Supabase / 基本ユーティリティ
===================================================== */

/* ================================
Supabase
================================ */

const {
  SUPABASE_URL,
  SUPABASE_ANON_KEY,
  ORIGIN_LABEL,
  ORIGIN_LAT,
  ORIGIN_LNG
} = window.APP_CONFIG;

const supabaseClient = supabase.createClient(
  SUPABASE_URL,
  SUPABASE_ANON_KEY
);


/* ================================
キャッシュ
================================ */

let casts = []
let vehicles = []
let plans = []
let actuals = []

let currentActualsCache = []


/* ================================
状態
================================ */

let editingCastId = null
let editingVehicleId = null
let editingPlanId = null
let editingActualId = null


/* ================================
UTIL
================================ */

function normalizeStatus(s){

  if(!s) return "pending"

  return String(s).toLowerCase()

}


/* ================================
座標判定
================================ */

function isValidLatLng(lat,lng){

  return (
    typeof lat === "number" &&
    typeof lng === "number" &&
    !isNaN(lat) &&
    !isNaN(lng)
  )

}


/* ================================
距離計算
================================ */

function haversineKm(lat1,lng1,lat2,lng2){

  const R = 6371

  const dLat = (lat2-lat1) * Math.PI/180
  const dLng = (lng2-lng1) * Math.PI/180

  const a =
    Math.sin(dLat/2) * Math.sin(dLat/2) +
    Math.cos(lat1*Math.PI/180) *
    Math.cos(lat2*Math.PI/180) *
    Math.sin(dLng/2) * Math.sin(dLng/2)

  const c = 2 * Math.atan2(Math.sqrt(a),Math.sqrt(1-a))

  return R * c

}


/* ================================
キャスト座標
================================ */

function getStopLatLng(stop){

  const lat = Number(stop?.casts?.latitude)
  const lng = Number(stop?.casts?.longitude)

  if(isValidLatLng(lat,lng)){

    return {lat,lng}

  }

  return null

}


/* ================================
距離計算（ルート）
================================ */

function calcRouteDistanceKm(stops){

  let total = 0

  let prevLat = ORIGIN_LAT
  let prevLng = ORIGIN_LNG

  stops.forEach(s=>{

    const p = getStopLatLng(s)

    if(!p) return

    total += haversineKm(
      prevLat,
      prevLng,
      p.lat,
      p.lng
    )

    prevLat = p.lat
    prevLng = p.lng

  })

  return Math.round(total)

}


/* ================================
ルート最適化
================================ */

function sortStopsByNearestNeighbor(stops){

  const result = []
  const pool = [...stops]

  let currentLat = ORIGIN_LAT
  let currentLng = ORIGIN_LNG

  while(pool.length){

    let bestIndex = 0
    let bestDist = Infinity

    pool.forEach((s,i)=>{

      const p = getStopLatLng(s)
      if(!p) return

      const d = haversineKm(
        currentLat,
        currentLng,
        p.lat,
        p.lng
      )

      if(d < bestDist){

        bestDist = d
        bestIndex = i

      }

    })

    const next = pool.splice(bestIndex,1)[0]

    result.push(next)

    const p = getStopLatLng(next)

    if(p){

      currentLat = p.lat
      currentLng = p.lng

    }

  }

  return result

}


/* ================================
時間表示
================================ */

function getHourLabel(h){

  return `${h}時`

}
/* =====================================================
THEMIS AI Dispatch
dashboard.js
PART 2 / 9
キャスト管理
===================================================== */


/* ================================
キャスト取得
================================ */

async function loadCasts(){

  const {data,error} =
    await supabaseClient
      .from("casts")
      .select("*")
      .order("name",{ascending:true})

  if(error){

    console.error(error)
    return

  }

  casts = data || []

  renderCasts()

  populateCastSelect()

}


/* ================================
キャスト保存
================================ */

async function saveCast(){

  const name =
    document.getElementById("castName").value.trim()

  const address =
    document.getElementById("castAddress").value.trim()

  const area =
    document.getElementById("castArea").value.trim()

  const distance =
    Number(
      document.getElementById("castDistanceKm").value
    )

  const memo =
    document.getElementById("castMemo").value.trim()

  const phone =
    document.getElementById("castPhone").value.trim()

  const lat =
    Number(document.getElementById("castLat").value)

  const lng =
    Number(document.getElementById("castLng").value)

  if(!name){

    alert("氏名を入力してください")
    return

  }

  const payload = {

    name,
    address,
    area,
    distance_km:distance,
    memo,
    phone,
    latitude:lat,
    longitude:lng

  }

  if(editingCastId){

    const {error} =
      await supabaseClient
        .from("casts")
        .update(payload)
        .eq("id",editingCastId)

    if(error){

      console.error(error)
      alert("更新失敗")
      return

    }

  }else{

    const {error} =
      await supabaseClient
        .from("casts")
        .insert([payload])

    if(error){

      console.error(error)
      alert("登録失敗")
      return

    }

  }

  resetCastForm()

  loadCasts()

}


/* ================================
キャスト編集
================================ */

function editCast(id){

  const c =
    casts.find(x=>x.id===id)

  if(!c) return

  editingCastId = id

  document.getElementById("castName").value =
    c.name || ""

  document.getElementById("castAddress").value =
    c.address || ""

  document.getElementById("castArea").value =
    c.area || ""

  document.getElementById("castDistanceKm").value =
    c.distance_km || ""

  document.getElementById("castMemo").value =
    c.memo || ""

  document.getElementById("castPhone").value =
    c.phone || ""

  document.getElementById("castLat").value =
    c.latitude || ""

  document.getElementById("castLng").value =
    c.longitude || ""

}


/* ================================
キャスト削除
================================ */

async function deleteCast(id){

  if(!confirm("削除しますか？")) return

  const {error} =
    await supabaseClient
      .from("casts")
      .delete()
      .eq("id",id)

  if(error){

    console.error(error)
    alert("削除失敗")
    return

  }

  loadCasts()

}


/* ================================
キャスト表示
================================ */

function renderCasts(){

  const tbody =
    document.getElementById("castsTableBody")

  if(!tbody) return

  tbody.innerHTML = ""

  casts.forEach(c=>{

    const tr =
      document.createElement("tr")

    tr.innerHTML = `
      <td>${c.name || ""}</td>
      <td>${c.address || ""}</td>
      <td>${c.area || ""}</td>
      <td>${c.distance_km || ""}</td>
      <td>${c.memo || ""}</td>
      <td class="actions-cell">
        <button onclick="editCast(${c.id})">編集</button>
        <button onclick="deleteCast(${c.id})">削除</button>
      </td>
    `

    tbody.appendChild(tr)

  })

}


/* ================================
キャストセレクト
================================ */

function populateCastSelect(){

  const selects = [

    document.getElementById("planCastSelect"),
    document.getElementById("castSelect")

  ]

  selects.forEach(sel=>{

    if(!sel) return

    sel.innerHTML =
      `<option value="">選択してください</option>`

    casts.forEach(c=>{

      const opt =
        document.createElement("option")

      opt.value = c.id
      opt.textContent = c.name

      sel.appendChild(opt)

    })

  })

}


/* ================================
キャストフォームリセット
================================ */

function resetCastForm(){

  editingCastId = null

  document.getElementById("castName").value = ""
  document.getElementById("castAddress").value = ""
  document.getElementById("castArea").value = ""
  document.getElementById("castDistanceKm").value = ""
  document.getElementById("castMemo").value = ""
  document.getElementById("castPhone").value = ""
  document.getElementById("castLat").value = ""
  document.getElementById("castLng").value = ""

}
/* =====================================================
THEMIS AI Dispatch
dashboard.js
PART 3 / 9
車両管理
===================================================== */


/* ================================
車両取得
================================ */

async function loadVehicles(){

  const {data,error} =
    await supabaseClient
      .from("vehicles")
      .select("*")
      .order("plate_number",{ascending:true})

  if(error){

    console.error(error)
    return

  }

  vehicles = data || []

  renderVehicles()

  renderVehicleChecklist()

}


/* ================================
車両保存
================================ */

async function saveVehicle(){

  const plate =
    document.getElementById("vehiclePlateNumber").value.trim()

  const area =
    document.getElementById("vehicleArea").value.trim()

  const homeArea =
    document.getElementById("vehicleHomeArea").value.trim()

  const capacity =
    Number(
      document.getElementById("vehicleSeatCapacity").value
    )

  const driver =
    document.getElementById("vehicleDriverName").value.trim()

  const lineId =
    document.getElementById("vehicleLineId").value.trim()

  const status =
    document.getElementById("vehicleStatus").value

  const memo =
    document.getElementById("vehicleMemo").value.trim()

  if(!plate){

    alert("車両IDを入力してください")
    return

  }

  const payload = {

    plate_number:plate,
    vehicle_area:area,
    home_area:homeArea,
    seat_capacity:capacity,
    driver_name:driver,
    line_id:lineId,
    status:status,
    memo:memo

  }

  if(editingVehicleId){

    const {error} =
      await supabaseClient
        .from("vehicles")
        .update(payload)
        .eq("id",editingVehicleId)

    if(error){

      console.error(error)
      alert("更新失敗")
      return

    }

  }else{

    const {error} =
      await supabaseClient
        .from("vehicles")
        .insert([payload])

    if(error){

      console.error(error)
      alert("登録失敗")
      return

    }

  }

  resetVehicleForm()

  loadVehicles()

}


/* ================================
車両編集
================================ */

function editVehicle(id){

  const v =
    vehicles.find(x=>x.id===id)

  if(!v) return

  editingVehicleId = id

  document.getElementById("vehiclePlateNumber").value =
    v.plate_number || ""

  document.getElementById("vehicleArea").value =
    v.vehicle_area || ""

  document.getElementById("vehicleHomeArea").value =
    v.home_area || ""

  document.getElementById("vehicleSeatCapacity").value =
    v.seat_capacity || ""

  document.getElementById("vehicleDriverName").value =
    v.driver_name || ""

  document.getElementById("vehicleLineId").value =
    v.line_id || ""

  document.getElementById("vehicleStatus").value =
    v.status || "waiting"

  document.getElementById("vehicleMemo").value =
    v.memo || ""

}


/* ================================
車両削除
================================ */

async function deleteVehicle(id){

  if(!confirm("車両削除しますか？")) return

  const {error} =
    await supabaseClient
      .from("vehicles")
      .delete()
      .eq("id",id)

  if(error){

    console.error(error)
    alert("削除失敗")
    return

  }

  loadVehicles()

}


/* ================================
車両一覧表示
================================ */

function renderVehicles(){

  const tbody =
    document.getElementById("vehiclesTableBody")

  if(!tbody) return

  tbody.innerHTML = ""

  vehicles.forEach(v=>{

    const tr =
      document.createElement("tr")

    tr.innerHTML = `
      <td>${v.plate_number || ""}</td>
      <td>${v.vehicle_area || ""}</td>
      <td>${v.home_area || ""}</td>
      <td>${v.seat_capacity || ""}</td>
      <td>${v.driver_name || ""}</td>
      <td>${v.monthly_km || 0}</td>
      <td>${v.work_days || 0}</td>
      <td>${v.avg_km || 0}</td>
      <td class="actions-cell">
        <button onclick="editVehicle(${v.id})">編集</button>
        <button onclick="deleteVehicle(${v.id})">削除</button>
      </td>
    `

    tbody.appendChild(tr)

  })

}


/* ================================
車両フォームリセット
================================ */

function resetVehicleForm(){

  editingVehicleId = null

  document.getElementById("vehiclePlateNumber").value = ""
  document.getElementById("vehicleArea").value = ""
  document.getElementById("vehicleHomeArea").value = ""
  document.getElementById("vehicleSeatCapacity").value = ""
  document.getElementById("vehicleDriverName").value = ""
  document.getElementById("vehicleLineId").value = ""
  document.getElementById("vehicleMemo").value = ""

}


/* ================================
当日車両チェックリスト
================================ */

function renderVehicleChecklist(){

  const wrap =
    document.getElementById("dailyVehicleChecklist")

  if(!wrap) return

  wrap.innerHTML = ""

  vehicles.forEach(v=>{

    const div =
      document.createElement("div")

    div.className = "vehicle-check-item"

    const label =
      `${v.plate_number} / ${v.driver_name || ""}`

    div.innerHTML = `
      <label class="vehicle-check-label">${label}</label>
      <input
        type="checkbox"
        class="vehicle-check-input"
        data-vehicle='${JSON.stringify(v)}'
      >
    `

    wrap.appendChild(div)

  })

}
/* =====================================================
THEMIS AI Dispatch
dashboard.js
PART 4 / 9
予定表
===================================================== */


/* ================================
予定取得
================================ */

async function loadPlans(){

  const date =
    document.getElementById("planDate")?.value

  if(!date) return

  const {data,error} =
    await supabaseClient
      .from("dispatch_plans")
      .select(`
        *,
        casts(*)
      `)
      .eq("plan_date",date)
      .order("plan_hour",{ascending:true})

  if(error){

    console.error(error)
    return

  }

  plans = data || []

  renderPlans()

}


/* ================================
予定保存
================================ */

async function savePlan(){

  const castId =
    document.getElementById("planCastSelect").value

  const hour =
    Number(document.getElementById("planHour").value)

  const address =
    document.getElementById("planAddress").value.trim()

  const area =
    document.getElementById("planArea").value.trim()

  const distance =
    Number(document.getElementById("planDistanceKm").value)

  const note =
    document.getElementById("planNote").value.trim()

  const date =
    document.getElementById("planDate").value

  if(!castId){

    alert("キャストを選択してください")
    return

  }

  const payload = {

    cast_id:castId,
    plan_hour:hour,
    destination_address:address,
    planned_area:area,
    distance_km:distance,
    note:note,
    plan_date:date

  }

  if(editingPlanId){

    const {error} =
      await supabaseClient
        .from("dispatch_plans")
        .update(payload)
        .eq("id",editingPlanId)

    if(error){

      console.error(error)
      alert("更新失敗")
      return

    }

  }else{

    const {error} =
      await supabaseClient
        .from("dispatch_plans")
        .insert([payload])

    if(error){

      console.error(error)
      alert("登録失敗")
      return

    }

  }

  resetPlanForm()

  loadPlans()

}


/* ================================
予定編集
================================ */

function editPlan(id){

  const p =
    plans.find(x=>x.id===id)

  if(!p) return

  editingPlanId = id

  document.getElementById("planCastSelect").value =
    p.cast_id

  document.getElementById("planHour").value =
    p.plan_hour

  document.getElementById("planAddress").value =
    p.destination_address || ""

  document.getElementById("planArea").value =
    p.planned_area || ""

  document.getElementById("planDistanceKm").value =
    p.distance_km || ""

  document.getElementById("planNote").value =
    p.note || ""

}


/* ================================
予定削除
================================ */

async function deletePlan(id){

  if(!confirm("予定削除しますか？")) return

  const {error} =
    await supabaseClient
      .from("dispatch_plans")
      .delete()
      .eq("id",id)

  if(error){

    console.error(error)
    alert("削除失敗")
    return

  }

  loadPlans()

}


/* ================================
キャスト住所自動反映
================================ */

function applyCastToPlan(){

  const castId =
    document.getElementById("planCastSelect").value

  const cast =
    casts.find(c=>c.id==castId)

  if(!cast) return

  document.getElementById("planAddress").value =
    cast.address || ""

  document.getElementById("planArea").value =
    cast.area || ""

  document.getElementById("planDistanceKm").value =
    cast.distance_km || ""

}


/* ================================
予定表示
================================ */

function renderPlans(){

  const wrap =
    document.getElementById("plansGroupedTable")

  if(!wrap) return

  wrap.innerHTML = ""

  const hourMap = {}

  plans.forEach(p=>{

    const h = p.plan_hour

    if(!hourMap[h]) hourMap[h] = []

    hourMap[h].push(p)

  })


  Object.keys(hourMap)
    .sort((a,b)=>a-b)
    .forEach(hour=>{

      const section =
        document.createElement("div")

      section.className = "grouped-section"

      section.innerHTML =
        `<div class="grouped-hour-title">${hour}時</div>`

      const areaMap = {}

      hourMap[hour].forEach(p=>{

        const area =
          p.planned_area || "未分類"

        if(!areaMap[area]) areaMap[area] = []

        areaMap[area].push(p)

      })


      Object.keys(areaMap).forEach(area=>{

        const areaTitle =
          document.createElement("div")

        areaTitle.className="grouped-area-title"

        areaTitle.textContent = area

        section.appendChild(areaTitle)

        areaMap[area].forEach(p=>{

          const row =
            document.createElement("div")

          row.className="grouped-row"

          row.innerHTML = `
            <div>${p.casts?.name || ""}</div>
            <div>${p.destination_address || ""}</div>
            <div>${p.distance_km || ""}km</div>
            <div>${p.note || ""}</div>
            <div class="op-cell">
              <button onclick="editPlan(${p.id})">編集</button>
              <button onclick="deletePlan(${p.id})">削除</button>
            </div>
          `

          section.appendChild(row)

        })

      })

      wrap.appendChild(section)

    })

}


/* ================================
フォームリセット
================================ */

function resetPlanForm(){

  editingPlanId = null

  document.getElementById("planCastSelect").value=""
  document.getElementById("planAddress").value=""
  document.getElementById("planArea").value=""
  document.getElementById("planDistanceKm").value=""
  document.getElementById("planNote").value=""

}
/* =====================================================
THEMIS AI Dispatch
dashboard.js
PART 5 / 9
実際の送り
===================================================== */


/* ================================
実績取得
================================ */

async function loadActuals(){

  const date =
    document.getElementById("actualDate")?.value

  if(!date) return

  const {data,error} =
    await supabaseClient
      .from("dispatch_items")
      .select(`
        *,
        casts(*),
        vehicles(*)
      `)
      .eq("plan_date",date)
      .order("actual_hour",{ascending:true})

  if(error){

    console.error(error)
    return

  }

  actuals = data || []
  currentActualsCache = actuals

  renderActualTable()
  renderActualMatrix()

}


/* ================================
実績保存
================================ */

async function saveActual(){

  const castId =
    document.getElementById("castSelect").value

  const hour =
    Number(document.getElementById("actualHour").value)

  const address =
    document.getElementById("actualAddress").value.trim()

  const area =
    document.getElementById("actualArea").value.trim()

  const distance =
    Number(document.getElementById("actualDistanceKm").value)

  const note =
    document.getElementById("actualNote").value.trim()

  const status =
    document.getElementById("actualStatus").value

  const date =
    document.getElementById("actualDate").value

  if(!castId){

    alert("キャストを選択してください")
    return

  }

  const payload = {

    cast_id:castId,
    actual_hour:hour,
    destination_address:address,
    destination_area:area,
    distance_km:distance,
    note:note,
    status:status,
    plan_date:date

  }

  if(editingActualId){

    const {error} =
      await supabaseClient
        .from("dispatch_items")
        .update(payload)
        .eq("id",editingActualId)

    if(error){

      console.error(error)
      alert("更新失敗")
      return

    }

  }else{

    const {error} =
      await supabaseClient
        .from("dispatch_items")
        .insert([payload])

    if(error){

      console.error(error)
      alert("登録失敗")
      return

    }

  }

  resetActualForm()
  loadActuals()

}


/* ================================
予定 → 実績追加
================================ */

async function addPlanToActual(){

  const planId =
    document.getElementById("planSelect").value

  const plan =
    plans.find(p=>p.id==planId)

  if(!plan) return

  const payload = {

    cast_id:plan.cast_id,
    actual_hour:plan.plan_hour,
    destination_address:plan.destination_address,
    destination_area:plan.planned_area,
    distance_km:plan.distance_km,
    plan_date:plan.plan_date,
    status:"pending"

  }

  const {error} =
    await supabaseClient
      .from("dispatch_items")
      .insert([payload])

  if(error){

    console.error(error)
    alert("追加失敗")
    return

  }

  loadActuals()

}


/* ================================
実績削除
================================ */

async function deleteActual(id){

  if(!confirm("削除しますか？")) return

  const {error} =
    await supabaseClient
      .from("dispatch_items")
      .delete()
      .eq("id",id)

  if(error){

    console.error(error)
    alert("削除失敗")
    return

  }

  loadActuals()

}


/* ================================
実績テーブル表示
================================ */

function renderActualTable(){

  const wrap =
    document.getElementById("actualTableWrap")

  if(!wrap) return

  wrap.innerHTML = ""

  actuals.forEach(a=>{

    const row =
      document.createElement("div")

    row.className="grouped-row"

    row.innerHTML = `
      <div>${getHourLabel(a.actual_hour)}</div>
      <div>${a.casts?.name || ""}</div>
      <div>${a.destination_area || ""}</div>
      <div>${a.distance_km || ""}km</div>
      <div class="state-stack">

        <span class="badge-status ${a.status}">
          ${a.status}
        </span>

        <button onclick="deleteActual(${a.id})">
          削除
        </button>

      </div>
    `

    wrap.appendChild(row)

  })

}


/* ================================
時間×方面マトリクス
================================ */

function renderActualMatrix(){

  const wrap =
    document.getElementById("actualTimeAreaMatrix")

  if(!wrap) return

  wrap.innerHTML=""

  const map = {}

  actuals.forEach(a=>{

    const key =
      `${a.actual_hour}_${a.destination_area}`

    if(!map[key]) map[key] = []

    map[key].push(a)

  })


  Object.keys(map).forEach(key=>{

    const items = map[key]

    const card =
      document.createElement("div")

    card.className="matrix-card"

    card.innerHTML = `
      <div class="matrix-summary">
        ${items[0].actual_hour}時 / ${items[0].destination_area}
      </div>
    `

    items.forEach(i=>{

      const item =
        document.createElement("div")

      item.className="matrix-item"

      item.textContent =
        i.casts?.name || ""

      card.appendChild(item)

    })

    wrap.appendChild(card)

  })

}


/* ================================
フォームリセット
================================ */

function resetActualForm(){

  editingActualId = null

  document.getElementById("castSelect").value=""
  document.getElementById("actualAddress").value=""
  document.getElementById("actualArea").value=""
  document.getElementById("actualDistanceKm").value=""
  document.getElementById("actualNote").value=""

}
/* =====================================================
THEMIS AI Dispatch
dashboard.js
PART 6 / 9
PRO 配車アルゴリズム
===================================================== */


/* ================================
使用車両取得
================================ */

function getSelectedVehiclesForToday(){

  const checks =
    document.querySelectorAll(".vehicle-check-input")

  const list = []

  checks.forEach(c=>{

    if(c.checked){

      const v =
        JSON.parse(c.dataset.vehicle)

      list.push(v)

    }

  })

  return list

}


/* ================================
方面クラスタリング
================================ */

function clusterByArea(items){

  const map = {}

  items.forEach(i=>{

    const area =
      i.destination_area || "未分類"

    if(!map[area]) map[area] = []

    map[area].push(i)

  })

  return map

}


/* ================================
最短ルート
================================ */

function optimizeRouteNearest(stops){

  const result = []
  const pool = [...stops]

  let currentLat = ORIGIN_LAT
  let currentLng = ORIGIN_LNG

  while(pool.length){

    let bestIndex = 0
    let bestDist = Infinity

    pool.forEach((s,i)=>{

      const p = getStopLatLng(s)

      if(!p) return

      const d =
        haversineKm(
          currentLat,
          currentLng,
          p.lat,
          p.lng
        )

      if(d < bestDist){

        bestDist = d
        bestIndex = i

      }

    })

    const next =
      pool.splice(bestIndex,1)[0]

    result.push(next)

    const p = getStopLatLng(next)

    if(p){

      currentLat = p.lat
      currentLng = p.lng

    }

  }

  return result

}


/* ================================
距離バランス
================================ */

function getVehicleLoadScore(vehicle){

  const dist =
    Number(vehicle.loadDistance || 0)

  return dist * 0.5

}


/* ================================
帰宅方向
================================ */

function getHomeDirectionScore(vehicle,item){

  if(!vehicle.home_area) return 0

  if(vehicle.home_area === item.destination_area){

    return -8

  }

  return 0

}


/* ================================
PRO配車AI
================================ */

function optimizeAssignmentsPro(){

  const vehicles =
    getSelectedVehiclesForToday()

  if(!vehicles.length){

    alert("車両を選択してください")
    return

  }

  const items =
    currentActualsCache.filter(
      x=>normalizeStatus(x.status)!=="done" &&
         normalizeStatus(x.status)!=="cancel"
    )

  vehicles.forEach(v=>{

    v.assignedStops = []
    v.loadDistance = 0

  })

  const clusters =
    clusterByArea(items)

  Object.values(clusters).forEach(cluster=>{

    const ordered =
      optimizeRouteNearest(cluster)

    ordered.forEach(item=>{

      let bestVehicle = null
      let bestScore = Infinity

      vehicles.forEach(v=>{

        const capacity =
          Number(v.seat_capacity || 4)

        if(v.assignedStops.length >= capacity)
          return

        let score = 0

        /* 距離 */

        const p = getStopLatLng(item)

        if(p){

          score += haversineKm(
            ORIGIN_LAT,
            ORIGIN_LNG,
            p.lat,
            p.lng
          )

        }

        /* 距離バランス */

        score += getVehicleLoadScore(v)

        /* 帰宅 */

        score += getHomeDirectionScore(v,item)

        if(score < bestScore){

          bestScore = score
          bestVehicle = v

        }

      })

      if(bestVehicle){

        bestVehicle.assignedStops.push(item)

        item.vehicle_id = bestVehicle.id

        bestVehicle.loadDistance +=
          Number(item.distance_km || 0)

      }

    })

  })

  renderDispatchResult()

}
/* =====================================================
THEMIS AI Dispatch
dashboard.js
PART 7 / 9
配車結果UI
===================================================== */


/* ================================
配車結果描画
================================ */

function renderDispatchResult(){

  const wrap =
    document.getElementById("dailyDispatchResult")

  if(!wrap) return

  wrap.innerHTML = ""

  const vehicles =
    getSelectedVehiclesForToday()

  vehicles.forEach(vehicle=>{

    const rows =
      currentActualsCache.filter(
        i => Number(i.vehicle_id) === Number(vehicle.id)
      )

    const ordered =
      sortStopsByNearestNeighbor(rows)

    const distance =
      calcRouteDistanceKm(ordered)

    const card =
      document.createElement("div")

    card.className = "vehicle-result-card"

    card.innerHTML = `
      <div class="vehicle-result-head">

        <div class="vehicle-result-title">
          <h4>
            ${vehicle.driver_name || vehicle.plate_number}
          </h4>

          <div class="vehicle-result-meta">
            ${vehicle.vehicle_area || ""}
          </div>

        </div>

        <div class="vehicle-result-badges">

          <span class="metric-badge">
            ${ordered.length}件
          </span>

          <span class="metric-badge">
            ${distance}km
          </span>

        </div>

      </div>

      <div class="vehicle-result-body"></div>
    `

    const body =
      card.querySelector(".vehicle-result-body")

    if(!ordered.length){

      const empty =
        document.createElement("div")

      empty.className="empty-vehicle-text"

      empty.textContent="割り当てなし"

      body.appendChild(empty)

    }

    ordered.forEach((row,index)=>{

      const item =
        document.createElement("div")

      item.className="dispatch-row"

      const vehicleSelect =
        buildVehicleSelect(row)

      item.innerHTML = `
        <div class="dispatch-left">

          <span class="badge-order">
            ${index+1}
          </span>

          <span class="dispatch-name">
            ${row.casts?.name || ""}
          </span>

          <span class="dispatch-area">
            ${row.destination_area || ""}
          </span>

        </div>

        <div class="dispatch-right">

          <span class="dispatch-distance">
            ${row.distance_km || ""}km
          </span>

        </div>
      `

      const right =
        item.querySelector(".dispatch-right")

      right.appendChild(vehicleSelect)

      body.appendChild(item)

    })

    wrap.appendChild(card)

  })

}


/* ================================
車両変更
================================ */

function buildVehicleSelect(row){

  const select =
    document.createElement("select")

  select.className =
    "dispatch-vehicle-select"

  vehicles.forEach(v=>{

    const opt =
      document.createElement("option")

    opt.value = v.id

    opt.textContent =
      v.driver_name || v.plate_number

    if(Number(v.id) === Number(row.vehicle_id)){

      opt.selected = true

    }

    select.appendChild(opt)

  })

  select.addEventListener("change",async e=>{

    const newVehicleId =
      Number(e.target.value)

    const {error} =
      await supabaseClient
        .from("dispatch_items")
        .update({
          vehicle_id:newVehicleId
        })
        .eq("id",row.id)

    if(error){

      console.error(error)
      alert("車両変更失敗")
      return

    }

    loadActuals()

  })

  return select

}
/* =====================================================
THEMIS AI Dispatch
dashboard.js
PART 8 / 9
Googleルート / コピー
===================================================== */


/* ================================
Googleナビ生成
================================ */

function buildGoogleMapsStopLink(lat,lng,address=""){

  const nLat = Number(lat)
  const nLng = Number(lng)

  if(isValidLatLng(nLat,nLng)){

    return `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(ORIGIN_LABEL)}&destination=${nLat},${nLng}&travelmode=driving`

  }

  const dest = String(address||"")

  if(!dest) return ""

  return `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(ORIGIN_LABEL)}&destination=${encodeURIComponent(dest)}&travelmode=driving`

}


/* ================================
Google複数ルート
================================ */

function buildGoogleMapsMultiStopLink(rows){

  if(!rows.length) return ""

  const stops =
    rows.map(r=>{

      const lat =
        Number(r?.casts?.latitude)

      const lng =
        Number(r?.casts?.longitude)

      if(isValidLatLng(lat,lng)){

        return `${lat},${lng}`

      }

      return r.destination_address

    }).filter(Boolean)

  if(!stops.length) return ""

  if(stops.length === 1){

    return `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(ORIGIN_LABEL)}&destination=${encodeURIComponent(stops[0])}`

  }

  const dest =
    stops[stops.length-1]

  const way =
    stops.slice(0,-1)

  return `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(ORIGIN_LABEL)}&destination=${encodeURIComponent(dest)}&waypoints=${encodeURIComponent(way.join("|"))}`

}


/* ================================
コピー用テキスト
================================ */

function buildCopyResultText(){

  const vehicles =
    getSelectedVehiclesForToday()

  const lines = []

  vehicles.forEach(v=>{

    const rows =
      currentActualsCache
        .filter(
          i => Number(i.vehicle_id) === Number(v.id)
        )

    const ordered =
      sortStopsByNearestNeighbor(rows)

    const route =
      buildGoogleMapsMultiStopLink(ordered)

    const distance =
      calcRouteDistanceKm(ordered)

    lines.push("================================")

    lines.push(
      `${v.driver_name || v.plate_number}`
    )

    lines.push(
      `想定距離 ${distance}km`
    )

    lines.push("ルート")

    lines.push(route)

    lines.push("")

    ordered.forEach((r,i)=>{

      const link =
        buildGoogleMapsStopLink(
          r?.casts?.latitude,
          r?.casts?.longitude,
          r.destination_address
        )

      lines.push(
        `${i+1}件目 ${r.casts?.name}`
      )

      lines.push(link)

      lines.push("")

    })

  })

  return lines.join("\n")

}


/* ================================
コピー実行
================================ */

function copyDispatchResult(){

  const text =
    buildCopyResultText()

  navigator.clipboard.writeText(text)

  alert("配車結果コピーしました")

}
/* =====================================================
THEMIS AI Dispatch
dashboard.js
PART 8 / 9
Googleルート / コピー
===================================================== */


/* ================================
Googleナビ生成
================================ */

function buildGoogleMapsStopLink(lat,lng,address=""){

  const nLat = Number(lat)
  const nLng = Number(lng)

  if(isValidLatLng(nLat,nLng)){

    return `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(ORIGIN_LABEL)}&destination=${nLat},${nLng}&travelmode=driving`

  }

  const dest = String(address||"")

  if(!dest) return ""

  return `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(ORIGIN_LABEL)}&destination=${encodeURIComponent(dest)}&travelmode=driving`

}


/* ================================
Google複数ルート
================================ */

function buildGoogleMapsMultiStopLink(rows){

  if(!rows.length) return ""

  const stops =
    rows.map(r=>{

      const lat =
        Number(r?.casts?.latitude)

      const lng =
        Number(r?.casts?.longitude)

      if(isValidLatLng(lat,lng)){

        return `${lat},${lng}`

      }

      return r.destination_address

    }).filter(Boolean)

  if(!stops.length) return ""

  if(stops.length === 1){

    return `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(ORIGIN_LABEL)}&destination=${encodeURIComponent(stops[0])}`

  }

  const dest =
    stops[stops.length-1]

  const way =
    stops.slice(0,-1)

  return `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(ORIGIN_LABEL)}&destination=${encodeURIComponent(dest)}&waypoints=${encodeURIComponent(way.join("|"))}`

}


/* ================================
コピー用テキスト
================================ */

function buildCopyResultText(){

  const vehicles =
    getSelectedVehiclesForToday()

  const lines = []

  vehicles.forEach(v=>{

    const rows =
      currentActualsCache
        .filter(
          i => Number(i.vehicle_id) === Number(v.id)
        )

    const ordered =
      sortStopsByNearestNeighbor(rows)

    const route =
      buildGoogleMapsMultiStopLink(ordered)

    const distance =
      calcRouteDistanceKm(ordered)

    lines.push("================================")

    lines.push(
      `${v.driver_name || v.plate_number}`
    )

    lines.push(
      `想定距離 ${distance}km`
    )

    lines.push("ルート")

    lines.push(route)

    lines.push("")

    ordered.forEach((r,i)=>{

      const link =
        buildGoogleMapsStopLink(
          r?.casts?.latitude,
          r?.casts?.longitude,
          r.destination_address
        )

      lines.push(
        `${i+1}件目 ${r.casts?.name}`
      )

      lines.push(link)

      lines.push("")

    })

  })

  return lines.join("\n")

}


/* ================================
コピー実行
================================ */

function copyDispatchResult(){

  const text =
    buildCopyResultText()

  navigator.clipboard.writeText(text)

  alert("配車結果コピーしました")

}
