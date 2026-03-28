// supabaseService
// save/load/import service layer

async function ensureAuth() {
  return true;
}

function getCurrentUserIdSafe() {
  return currentUser?.id || window.currentUser?.id || null;
}

function formatLocalDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function todayStr() {
  return formatLocalDate(new Date());
}

function getMonthStartStr(dateStr) {
  const d = new Date(dateStr || todayStr());
  return formatLocalDate(new Date(d.getFullYear(), d.getMonth(), 1));
}

function getMonthEndStr(dateStr) {
  const d = new Date(dateStr || todayStr());
  return formatLocalDate(new Date(d.getFullYear(), d.getMonth() + 1, 0));
}

async function saveCast() {
  const name = els.castName?.value.trim();
  const address = els.castAddress?.value.trim();

  if (!name) {
    alert("氏名を入力してください");
    return;
  }

  const duplicate = typeof isDuplicateCast === "function" ? isDuplicateCast(name, address) : null;
  if (duplicate) {
    alert("このキャストは既に登録されています");
    return;
  }

  let lat = typeof toNullableNumber === "function" ? toNullableNumber(els.castLat?.value) : null;
  let lng = typeof toNullableNumber === "function" ? toNullableNumber(els.castLng?.value) : null;

  const addressKey = typeof normalizeGeocodeAddressKey === "function"
    ? normalizeGeocodeAddressKey(address)
    : String(address || "").trim().toLowerCase();

  if (address && typeof isValidLatLng === "function" && (!isValidLatLng(lat, lng) || addressKey !== lastCastGeocodeKey)) {
    const geocoded = typeof fillCastLatLngFromAddress === "function"
      ? await fillCastLatLngFromAddress({ silent: true, force: addressKey !== lastCastGeocodeKey })
      : null;
    lat = geocoded?.lat ?? (typeof toNullableNumber === "function" ? toNullableNumber(els.castLat?.value) : null);
    lng = geocoded?.lng ?? (typeof toNullableNumber === "function" ? toNullableNumber(els.castLng?.value) : null);
  }

  const manualArea = els.castArea?.value.trim() || "";
  const autoArea = typeof guessArea === "function" ? guessArea(lat, lng, address) : "";
  const autoDistance = typeof resolveDistanceKmFromOrigin === "function"
    ? await resolveDistanceKmFromOrigin(address, lat, lng)
    : null;

  const payload = {
    name,
    phone: els.castPhone?.value.trim() || "",
    address,
    area: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(manualArea || autoArea || "") : (manualArea || autoArea || ""),
    distance_km: (typeof toNullableNumber === "function" ? toNullableNumber(els.castDistanceKm?.value) : null) ?? autoDistance,
    travel_minutes: typeof getStoredTravelMinutes === "function" ? (getStoredTravelMinutes(els.castTravelMinutes?.value) || null) : null,
    latitude: lat,
    longitude: lng,
    memo: els.castMemo?.value.trim() || "",
    is_active: true
  };

  let error;
  if (editingCastId) {
    ({ error } = await supabaseClient.from("casts").update(payload).eq("id", editingCastId));
  } else {
    payload.created_by = getCurrentUserIdSafe();
    ({ error } = await supabaseClient.from("casts").insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, editingCastId ? "update_cast" : "create_cast", editingCastId ? "キャストを更新" : "キャストを作成");
  if (typeof resetCastForm === "function") resetCastForm();
  await loadCasts();
}

async function deleteCast(castId) {
  if (!window.confirm("このキャストを削除しますか？")) return;

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
}

async function importCastCsvFile() {
  const file = els.csvFileInput?.files?.[0];
  if (!file) {
    alert("CSVファイルを選択してください");
    return;
  }

  try {
    const text = await readCsvFileAsText(file);
    let rows = parseCsv(text);
    rows = normalizeCsvRows(rows);

    if (!rows.length) {
      alert("CSVデータが空です");
      return;
    }

    const uniqueMap = new Map();

    for (const row of rows) {
      const name = String(row.name || "").trim();
      const address = String(row.address || "").trim();
      if (!name || !address) continue;

      const key = `${name}__${address}`;
      if (!uniqueMap.has(key)) uniqueMap.set(key, row);
    }

    const mergedRows = [...uniqueMap.values()];
    const payloads = [];

    for (const row of mergedRows) {
      const name = String(row.name || "").trim();
      const address = String(row.address || "").trim();
      if (!name || !address) continue;

      const lat = typeof toNullableNumber === "function" ? toNullableNumber(row.latitude) : null;
      const lng = typeof toNullableNumber === "function" ? toNullableNumber(row.longitude) : null;
      const autoArea = typeof guessArea === "function" ? guessArea(lat, lng, address) : "";

      payloads.push({
        name,
        phone: String(row.phone || "").trim(),
        address,
        area: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(String(row.area || "").trim() || autoArea || "") : (String(row.area || "").trim() || autoArea || ""),
        distance_km:
          (typeof toNullableNumber === "function" ? toNullableNumber(row.distance_km) : null) ??
          ((typeof isValidLatLng === "function" && isValidLatLng(lat, lng) && typeof estimateRoadKmFromStation === "function") ? estimateRoadKmFromStation(lat, lng) : null),
        travel_minutes: typeof getStoredTravelMinutes === "function" ? (getStoredTravelMinutes(row.travel_minutes) || null) : null,
        latitude: lat,
        longitude: lng,
        memo: String(row.memo || "").trim(),
        is_active: true,
        created_by: getCurrentUserIdSafe()
      });
    }

    if (!payloads.length) {
      alert("取り込めるデータがありません");
      els.csvFileInput.value = "";
      return;
    }

    const { error } = await supabaseClient
      .from("casts")
      .upsert(payloads, { onConflict: "name,address" });

    if (error) {
      console.error("CSV import supabase error:", error);
      alert("CSV取込エラー: " + error.message);
      return;
    }

    els.csvFileInput.value = "";
    await addHistory(null, null, "import_csv", `${payloads.length}件のキャストをCSV取込/更新`);
    alert(`${payloads.length}件のキャストをCSV取込/更新しました`);
    await loadCasts();
  } catch (error) {
    console.error("importCastCsvFile error:", error);
    alert("CSV取込中にエラーが発生しました");
  }
}

async function saveVehicle() {
  const plateNumber = els.vehiclePlateNumber?.value.trim();
  if (!plateNumber) {
    alert("車両IDを入力してください");
    return;
  }

  const duplicate = typeof isDuplicateVehicle === "function" ? isDuplicateVehicle(plateNumber) : null;
  if (duplicate) {
    alert("この車両IDは既に登録されています");
    return;
  }

  const payload = {
    plate_number: plateNumber,
    vehicle_area: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(els.vehicleArea?.value.trim() || "") : (els.vehicleArea?.value.trim() || ""),
    home_area: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(els.vehicleHomeArea?.value.trim() || "") : (els.vehicleHomeArea?.value.trim() || ""),
    seat_capacity: Number(els.vehicleSeatCapacity?.value || 4),
    driver_name: els.vehicleDriverName?.value.trim() || "",
    line_id: els.vehicleLineId?.value.trim() || "",
    status: els.vehicleStatus?.value || "waiting",
    memo: els.vehicleMemo?.value.trim() || "",
    is_active: true
  };

  let error;
  if (editingVehicleId) {
    ({ error } = await supabaseClient.from("vehicles").update(payload).eq("id", editingVehicleId));
  } else {
    payload.created_by = getCurrentUserIdSafe();
    ({ error } = await supabaseClient.from("vehicles").insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, editingVehicleId ? "update_vehicle" : "create_vehicle", editingVehicleId ? "車両を更新" : "車両を登録");
  if (typeof resetVehicleForm === "function") resetVehicleForm();
  await loadVehicles();
}

async function deleteVehicle(vehicleId) {
  if (!window.confirm("この車両を削除しますか？")) return;

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
}

async function importVehicleCsvFile() {
  const file = els.vehicleCsvFileInput?.files?.[0];
  if (!file) {
    alert("CSVファイルを選択してください");
    return;
  }

  try {
    const text = await readCsvFileAsText(file);
    let rows = parseCsv(text);
    rows = normalizeCsvRows(rows);

    if (!rows.length) {
      alert("CSVデータが空です");
      return;
    }

    const inserts = [];

    for (const row of rows) {
      const plateNumber = String(row.plate_number || "").trim();
      if (!plateNumber) continue;

      const exists = Array.isArray(allVehiclesCache) ? allVehiclesCache.find(
        v => String(v.plate_number || "").trim() === plateNumber
      ) : null;
      if (exists) {
        console.log("車両重複スキップ:", plateNumber);
        continue;
      }

      inserts.push({
        plate_number: plateNumber,
        vehicle_area: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(String(row.vehicle_area || "").trim() || "") : String(row.vehicle_area || "").trim(),
        home_area: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(String(row.home_area || "").trim() || "") : String(row.home_area || "").trim(),
        seat_capacity: Number(row.seat_capacity || 4),
        driver_name: String(row.driver_name || "").trim(),
        line_id: String(row.line_id || "").trim(),
        status: String(row.status || "waiting").trim() || "waiting",
        memo: String(row.memo || "").trim(),
        is_active: true,
        created_by: getCurrentUserIdSafe()
      });
    }

    if (!inserts.length) {
      alert("新規車両はありません");
      els.vehicleCsvFileInput.value = "";
      return;
    }

    const { error } = await supabaseClient.from("vehicles").insert(inserts);
    if (error) {
      console.error("Vehicle CSV import error:", error);
      alert("車両CSV取込エラー: " + error.message);
      return;
    }

    els.vehicleCsvFileInput.value = "";
    await addHistory(null, null, "import_vehicle_csv", `${inserts.length}件の車両をCSV取込`);
    alert(`${inserts.length}件の車両を取り込みました`);
    await loadVehicles();
  } catch (error) {
    console.error("importVehicleCsvFile error:", error);
    alert("車両CSV取込中にエラーが発生しました");
  }
}

async function fetchDriverMileageRows(startDate, endDate) {
  const { data, error } = await supabaseClient
    .from("vehicle_daily_reports")
    .select(`
      *,
      vehicles (
        id,
        plate_number,
        driver_name
      )
    `)
    .gte("report_date", startDate)
    .lte("report_date", endDate)
    .order("report_date", { ascending: true })
    .order("vehicle_id", { ascending: true });

  if (error) {
    console.error(error);
    alert("走行実績の取得に失敗しました: " + error.message);
    return [];
  }

  return data || [];
}

async function loadCasts() {
  const { data, error } = await supabaseClient
    .from("casts")
    .select("*")
    .eq("is_active", true)
    .order("name", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  allCastsCache = data || [];
  if (typeof renderCastsTable === "function") renderCastsTable();
  if (typeof renderCastSearchResults === "function") renderCastSearchResults();
  if (typeof renderCastSelects === "function") renderCastSelects();
  if (typeof renderHomeSummary === "function") renderHomeSummary();
}

async function loadVehicles() {
  const { data, error } = await supabaseClient
    .from("vehicles")
    .select("*")
    .eq("is_active", true)
    .order("plate_number", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  allVehiclesCache = data || [];

  const validIds = new Set(allVehiclesCache.map(v => Number(v.id || 0)).filter(Boolean));
  activeVehicleIdsForToday = new Set(
    [...(activeVehicleIdsForToday || new Set())].filter(id => validIds.has(Number(id || 0)))
  );

  if (typeof renderVehiclesTable === "function") renderVehiclesTable();
  if (typeof renderDailyVehicleChecklist === "function") renderDailyVehicleChecklist();
  if (typeof renderDailyMileageInputs === "function") renderDailyMileageInputs();
  if (typeof renderDailyDispatchResult === "function") renderDailyDispatchResult();
  if (typeof renderHomeSummary === "function") renderHomeSummary();
  if (typeof refreshHomeMonthlyVehicleList === "function") await refreshHomeMonthlyVehicleList();
  else if (typeof renderHomeMonthlyVehicleList === "function") renderHomeMonthlyVehicleList();
}

async function loadDailyReports(dateStr) {
  const start = getMonthStartStr(dateStr || todayStr());
  const end = getMonthEndStr(dateStr || todayStr());

  const { data, error } = await supabaseClient
    .from("vehicle_daily_reports")
    .select(`
      *,
      vehicles (
        id,
        plate_number,
        driver_name
      )
    `)
    .gte("report_date", start)
    .lte("report_date", end)
    .order("report_date", { ascending: true })
    .order("vehicle_id", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  currentDailyReportsCache = data || [];
  if (typeof renderVehiclesTable === "function") renderVehiclesTable();
  if (typeof refreshHomeMonthlyVehicleList === "function") await refreshHomeMonthlyVehicleList();
  else if (typeof renderHomeMonthlyVehicleList === "function") renderHomeMonthlyVehicleList();
  if (typeof renderDailyMileageInputs === "function") renderDailyMileageInputs();
  if (typeof renderDailyDispatchResult === "function") renderDailyDispatchResult();
}

async function loadHistory() {
  const { data, error } = await supabaseClient
    .from("dispatch_history")
    .select("*")
    .order("created_at", { ascending: false })
    .limit(100);

  if (error) {
    console.error(error);
    return;
  }

  if (!els?.historyList) return;
  els.historyList.innerHTML = "";

  if (!data?.length) {
    els.historyList.innerHTML = `<div class="muted">履歴はありません</div>`;
    return;
  }

  data.forEach(row => {
    const div = document.createElement("div");
    div.className = "history-item";
    div.innerHTML = `
      <h4>${escapeHtml(row.action)}</h4>
      <p>${escapeHtml(row.message || "")}</p>
      <p class="muted">${escapeHtml(formatDateTimeJa(row.created_at))}</p>
    `;
    els.historyList.appendChild(div);
  });
}

async function loadHomeAndAll() {
  const dateStr = els?.dispatchDate?.value || todayStr();

  if (els?.dispatchDate) els.dispatchDate.value = dateStr;
  if (els?.planDate) els.planDate.value = dateStr;
  if (els?.actualDate) els.actualDate.value = dateStr;
  if (typeof syncMileageReportRange === "function") {
    syncMileageReportRange(dateStr, true);
  } else {
    if (els?.mileageReportStartDate) els.mileageReportStartDate.value = getMonthStartStr(dateStr);
    if (els?.mileageReportEndDate) els.mileageReportEndDate.value = dateStr;
  }

  await loadCasts();
  await loadVehicles();
  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
  await loadHistory();

  if (typeof renderDailyVehicleChecklist === "function") renderDailyVehicleChecklist();
  if (typeof renderDailyMileageInputs === "function") renderDailyMileageInputs();
  if (typeof renderDailyDispatchResult === "function") renderDailyDispatchResult();
  if (typeof renderHomeSummary === "function") renderHomeSummary();
  if (typeof refreshHomeMonthlyVehicleList === "function") await refreshHomeMonthlyVehicleList();
  else if (typeof renderHomeMonthlyVehicleList === "function") renderHomeMonthlyVehicleList();
}

async function loadAllData() {
  return loadHomeAndAll();
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
        distance_km,
        travel_minutes,
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
  renderPlanGroupedTable();
  renderPlansTimeAreaMatrix();
  renderPlanSelect();
  renderPlanCastSelect();
  renderHomeSummary();
  renderOperationAndSimulationUI();
}

async function loadActualsByDate(dateStr) {
  const { data: dispatches, error: dispatchError } = await supabaseClient
    .from("dispatches")
    .select("*")
    .eq("dispatch_date", dateStr)
    .order("id", { ascending: false })
    .limit(1);

  if (dispatchError) {
    console.error(dispatchError);
    return;
  }

  if (dispatches?.length) {
    currentDispatchId = dispatches[0].id;
  } else {
    const { data: inserted, error: createError } = await supabaseClient
      .from("dispatches")
      .insert({
        dispatch_date: dateStr,
        status: "draft",
        created_by: getCurrentUserIdSafe()
      })
      .select()
      .single();

    if (createError) {
      console.error(createError);
      return;
    }
    currentDispatchId = inserted.id;
  }

  const { data, error } = await supabaseClient
    .from("dispatch_items")
    .select(`
      *,
      casts (
        id,
        name,
        phone,
        address,
        area,
        distance_km,
        travel_minutes,
        latitude,
        longitude
      )
    `)
    .eq("dispatch_id", currentDispatchId)
    .order("actual_hour", { ascending: true })
    .order("stop_order", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  currentActualsCache = data || [];
  renderActualTable();
  renderActualTimeAreaMatrix();
  renderHomeSummary();
  renderCastSelects();
  renderManualLastVehicleInfo();
}

async function savePlan() {
  const cast = findCastByInputValue(els.planCastSelect?.value || "");
  const castId = Number(cast?.id || 0);
  if (!castId) {
    alert("キャストを選択または入力してください");
    return;
  }

  const planDate = els.planDate?.value || todayStr();
  const hour = Number(els.planHour?.value || 0);
  const address = els.planAddress?.value.trim() || "";
  let distanceKm = toNullableNumber(els.planDistanceKm?.value);
  if (distanceKm === null) {
    distanceKm = await resolveDistanceKmForCastRecord(cast, address);
    if (distanceKm !== null && els.planDistanceKm) els.planDistanceKm.value = String(distanceKm);
  }
  const area = els.planArea?.value.trim() || "";
  const note = els.planNote?.value.trim() || "";

  const payload = {
    plan_date: planDate,
    plan_hour: hour,
    cast_id: castId,
    destination_address: address,
    planned_area: normalizeAreaLabel(area || "無し"),
    distance_km: distanceKm,
    note,
    status: "planned"
  };

  let error;
  if (editingPlanId) {
    ({ error } = await supabaseClient.from("dispatch_plans").update(payload).eq("id", editingPlanId));
  } else {
    payload.created_by = getCurrentUserIdSafe();
    ({ error } = await supabaseClient.from("dispatch_plans").insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(
    null,
    null,
    editingPlanId ? "update_plan" : "create_plan",
    editingPlanId ? "予定を更新" : "予定を作成"
  );

  resetPlanForm();
  await loadPlansByDate(planDate);
}

async function deletePlan(planId) {
  if (!window.confirm("この予定を削除しますか？")) return;

  const { error } = await supabaseClient.from("dispatch_plans").delete().eq("id", planId);
  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "delete_plan", `予定ID ${planId} を削除`);
  await loadPlansByDate(els.planDate?.value || todayStr());
}

async function clearAllPlans() {
  if (!window.confirm("この日の予定を全消去しますか？")) return;

  const planDate = els.planDate?.value || todayStr();
  const { error } = await supabaseClient
    .from("dispatch_plans")
    .delete()
    .eq("plan_date", planDate);

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "clear_plans", `${planDate} の予定を全削除`);
  await loadPlansByDate(planDate);
}

async function saveActual() {
  const cast = findCastByInputValue(els.castSelect?.value || "");
  const castId = Number(cast?.id || 0);
  if (!castId) {
    alert("キャストを選択または入力してください");
    return;
  }

  const dateStr = els.actualDate?.value || todayStr();
  const hour = Number(els.actualHour?.value || 0);
  const address = els.actualAddress?.value.trim() || "";
  const area = normalizeAreaLabel(els.actualArea?.value.trim() || "無し");
  let distanceKm = toNullableNumber(els.actualDistanceKm?.value);
  if (distanceKm === null) {
    distanceKm = await resolveDistanceKmForCastRecord(cast, address);
    if (distanceKm !== null && els.actualDistanceKm) els.actualDistanceKm.value = String(distanceKm);
  }
  const status = els.actualStatus?.value || "pending";
  const note = els.actualNote?.value.trim() || "";

  const existingActual = editingActualId
    ? currentActualsCache.find(x => Number(x.id) === Number(editingActualId))
    : null;

  const stopOrder = existingActual
    ? Number(existingActual.stop_order || 1)
    : currentActualsCache.filter(
        x =>
          Number(x.actual_hour) === hour &&
          Number(x.id) !== Number(editingActualId || 0)
      ).length + 1;

  const payload = {
    dispatch_id: currentDispatchId,
    cast_id: castId,
    actual_hour: hour,
    stop_order: stopOrder,
    pickup_label: ORIGIN_LABEL,
    destination_address: address,
    destination_area: area,
    distance_km: distanceKm,
    status,
    note,
    plan_date: dateStr
  };

  let error;
  if (editingActualId) {
    ({ error } = await supabaseClient.from("dispatch_items").update(payload).eq("id", editingActualId));
  } else {
    ({ error } = await supabaseClient.from("dispatch_items").insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(
    currentDispatchId,
    editingActualId || null,
    editingActualId ? "update_actual" : "create_actual",
    editingActualId ? "実際の送りを更新" : "実際の送りを追加"
  );

  resetActualForm();
  await loadActualsByDate(dateStr);
  if (!editingActualId) {
    try {
      await assignUnassignedActualsForToday();
      await loadActualsByDate(dateStr);
    } catch (assignError) {
      console.error("assignUnassignedActualsForToday error:", assignError);
    }
  }
  await loadPlansByDate(els.planDate?.value || dateStr);
}

async function deleteActual(itemId) {
  if (!window.confirm("このActualを削除しますか？")) return;

  const { error } = await supabaseClient.from("dispatch_items").delete().eq("id", itemId);
  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(currentDispatchId, itemId, "delete_actual", `Actual ID ${itemId} を削除`);
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
}

async function updateActualStatus(itemId, status) {
  const item = currentActualsCache.find(x => Number(x.id) === Number(itemId));
  if (!item) {
    alert("対象のActualが見つかりません");
    return;
  }

  const { error } = await supabaseClient
    .from("dispatch_items")
    .update({ status })
    .eq("id", itemId);

  if (error) {
    alert(error.message);
    return;
  }

  const targetPlan = currentPlansCache.find(
    plan =>
      Number(plan.cast_id) === Number(item.cast_id) &&
      plan.plan_date === (els.actualDate?.value || todayStr()) &&
      Number(plan.plan_hour) === Number(item.actual_hour ?? -1)
  );

  if (targetPlan) {
    let nextPlanStatus = targetPlan.status;
    if (status === "done") nextPlanStatus = "done";
    else if (status === "cancel") nextPlanStatus = "cancel";
    else if (status === "pending") nextPlanStatus = "assigned";

    const { error: planError } = await supabaseClient
      .from("dispatch_plans")
      .update({ status: nextPlanStatus })
      .eq("id", targetPlan.id);

    if (planError) console.error(planError);
  }

  await addHistory(currentDispatchId, itemId, "update_actual_status", `Actual状態を ${status} に変更`);
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
}

async function addPlanToActual() {
  const planId = Number(els.planSelect?.value || 0);
  if (!planId) {
    alert("予定を選択してください");
    return;
  }

  const plan = currentPlansCache.find(x => Number(x.id) === Number(planId));
  if (!plan) {
    alert("予定が見つかりません");
    return;
  }

  if (isPlanAlreadyAddedToActual(plan)) {
    alert("その予定はすでにActualへ追加されています");
    renderPlanSelect();
    return;
  }

  if (currentActualsCache.some(x => Number(x.cast_id) === Number(plan.cast_id) && normalizeStatus(x.status) !== "cancel")) {
    alert("そのキャストはすでにActualにあります");
    renderPlanSelect();
    return;
  }

  const doneCastIds = getDoneCastIdsInActuals();
  if (doneCastIds.has(Number(plan.cast_id))) {
    alert("このキャストはすでに送り完了です");
    renderPlanSelect();
    return;
  }

  const payload = {
    dispatch_id: currentDispatchId,
    cast_id: plan.cast_id,
    actual_hour: Number(plan.plan_hour || 0),
    stop_order:
      currentActualsCache.filter(x => Number(x.actual_hour) === Number(plan.plan_hour || 0)).length + 1,
    pickup_label: ORIGIN_LABEL,
    destination_address: plan.destination_address || plan.casts?.address || "",
    destination_area: normalizeAreaLabel(plan.planned_area || "無し"),
    distance_km: plan.distance_km ?? plan.casts?.distance_km ?? null,
    status: "pending",
    note: plan.note || "",
    plan_date: plan.plan_date
  };

  const { error } = await supabaseClient.from("dispatch_items").insert(payload);
  if (error) {
    alert(error.message);
    return;
  }

  await supabaseClient
    .from("dispatch_plans")
    .update({ status: "assigned" })
    .eq("id", plan.id);

  await addHistory(currentDispatchId, null, "add_plan_to_actual", `予定ID ${plan.id} をActualへ追加`);
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
  if (els.planSelect) els.planSelect.value = "";
  renderPlanSelect();
}
