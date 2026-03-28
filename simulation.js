// simulation
// dashboard.js から安全に分離した シミュレーション診断 / 診断UI系

function getSimulationStateValue(key, fallback) {
  try {
    if (window.stateManager?.get) {
      const value = window.stateManager.get(key);
      return value === undefined ? fallback : value;
    }
  } catch (_) {}
  return fallback;
}

function setSimulationSlotHourSafe(value) {
  if (typeof window.setSimulationSlotHourState === "function") {
    return window.setSimulationSlotHourState(value);
  }
  if (window.stateManager?.setSimulationSlotHour) {
    return window.stateManager.setSimulationSlotHour(value);
  }
  window.simulationSlotHour = value;
  return value;
}

function getSimulationSlotHourSafe() {
  return getSimulationStateValue("simulationSlotHour", window.simulationSlotHour ?? null);
}

function setLastSimulationResultSafe(value) {
  if (typeof window.setLastSimulationResultState === "function") {
    return window.setLastSimulationResultState(value);
  }
  if (window.stateManager?.setLastSimulationResult) {
    return window.stateManager.setLastSimulationResult(value);
  }
  window.lastSimulationResult = value;
  return value;
}

function getLastSimulationResultSafe() {
  return getSimulationStateValue("lastSimulationResult", window.lastSimulationResult ?? null);
}

function setIsRefreshingHybridUISafe(value) {
  if (typeof window.setIsRefreshingHybridUIState === "function") {
    return window.setIsRefreshingHybridUIState(value);
  }
  if (window.stateManager?.setIsRefreshingHybridUI) {
    return window.stateManager.setIsRefreshingHybridUI(value);
  }
  window.isRefreshingHybridUI = Boolean(value);
  return Boolean(value);
}

function getIsRefreshingHybridUISafe() {
  return Boolean(getSimulationStateValue("isRefreshingHybridUI", window.isRefreshingHybridUI ?? false));
}

function setSuppressSimulationSlotChangeSafe(value) {
  if (typeof window.setSuppressSimulationSlotChangeState === "function") {
    return window.setSuppressSimulationSlotChangeState(value);
  }
  if (window.stateManager?.setSuppressSimulationSlotChange) {
    return window.stateManager.setSuppressSimulationSlotChange(value);
  }
  window.suppressSimulationSlotChange = Boolean(value);
  return Boolean(value);
}

function getSuppressSimulationSlotChangeSafe() {
  return Boolean(getSimulationStateValue("suppressSimulationSlotChange", window.suppressSimulationSlotChange ?? false));
}


function callDispatchCoreSafe(functionName, ...args) {
  if (typeof window.callDispatchCore === "function") {
    return window.callDispatchCore(functionName, ...args);
  }
  const core = window.DispatchCore || {};
  const fn = core?.[functionName];
  if (typeof fn !== "function") {
    if (functionName === "getVehicleDeadNamesForHour") return [];
    if (functionName === "getEffectiveVehiclesForHour") return [];
    if (functionName === "buildProjectedRowsForHour" || functionName === "buildSimulationRowsForHour") return functionName === "buildSimulationRowsForHour" ? { rows: [], summary: { slotPlanCount: 0, inflowPlanCount: 0 } } : [];
    return [];
  }
  return fn(...args);
}

function getBuildSimulationRowsForHourSafe(targetHour, options = {}) {
  return callDispatchCoreSafe("buildSimulationRowsForHour", targetHour, options);
}

function getBuildProjectedRowsForHourSafe(targetHour) {
  return callDispatchCoreSafe("buildProjectedRowsForHour", targetHour);
}

function getEffectiveVehiclesForHourSafe(targetHour) {
  return callDispatchCoreSafe("getEffectiveVehiclesForHour", targetHour);
}

function getVehicleDeadNamesForHourSafe(targetHour) {
  return callDispatchCoreSafe("getVehicleDeadNamesForHour", targetHour);
}

function diagnoseSimulationHourWindow(targetHour, options = {}) {
  const built = getBuildSimulationRowsForHourSafe(targetHour, options);
  const rows = built.rows;
  const vehicles = getEffectiveVehiclesForHourSafe(targetHour);
  const effectiveVehicleCount = vehicles.length;
  const totalCapacity = vehicles.reduce((sum, vehicle) => sum + Math.max(1, Number(vehicle?.seat_capacity || 4)), 0);
  const castCount = rows.length;
  const areaGroups = new Set(rows.map(row => getAreaDisplayGroup(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "無し"))));
  const shortageCount = Math.max(0, castCount - totalCapacity);
  let statusKey = "ok";
  let statusText = "OK";
  if (totalCapacity < castCount) {
    statusKey = "danger";
    statusText = "定員オーバー見込み";
  } else if (areaGroups.size > effectiveVehicleCount && castCount > 0) {
    statusKey = "warn";
    statusText = "車両不足見込み";
  }
  return {
    hour: Number(targetHour),
    rows,
    vehicles,
    castCount,
    effectiveVehicleCount,
    totalCapacity,
    areaGroupCount: areaGroups.size,
    shortageCount,
    statusKey,
    statusText,
    deadVehicleNames: getVehicleDeadNamesForHourSafe(targetHour),
    slotPlanCount: built.summary.slotPlanCount,
    inflowPlanCount: built.summary.inflowPlanCount
  };
}


function diagnoseHourWindow(targetHour) {
  const rows = getBuildProjectedRowsForHourSafe(targetHour);
  const vehicles = getEffectiveVehiclesForHourSafe(targetHour);
  const effectiveVehicleCount = vehicles.length;
  const totalCapacity = vehicles.reduce((sum, vehicle) => sum + Math.max(1, Number(vehicle?.seat_capacity || 4)), 0);
  const castCount = rows.length;
  const areaGroups = new Set(rows.map(row => getAreaDisplayGroup(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "無し"))));
  const shortageCount = Math.max(0, castCount - totalCapacity);
  let statusKey = "ok";
  let statusText = "OK";
  if (totalCapacity < castCount) {
    statusKey = "danger";
    statusText = "定員オーバー見込み";
  } else if (areaGroups.size > effectiveVehicleCount && castCount > 0) {
    statusKey = "warn";
    statusText = "車両不足見込み";
  }
  return {
    hour: Number(targetHour),
    rows,
    vehicles,
    castCount,
    effectiveVehicleCount,
    totalCapacity,
    areaGroupCount: areaGroups.size,
    shortageCount,
    statusKey,
    statusText,
    deadVehicleNames: getVehicleDeadNamesForHourSafe(targetHour)
  };
}


function getStatusPillClass(statusKey) {
  if (statusKey === "danger") return "danger";
  if (statusKey === "warn") return "warn";
  return "ok";
}


function buildStatusPill(label, value, statusKey = "ok") {
  return `<span class="status-pill ${getStatusPillClass(statusKey)}"><span class="status-pill-label">${escapeHtml(label)}</span><span class="status-pill-value">${escapeHtml(value)}</span></span>`;
}


function buildSummaryItemsHtml(items) {
  return items.map(item => `
    <div class="hybrid-summary-item ${escapeHtml(item.tone || '')}">
      <div class="hybrid-summary-icon">${escapeHtml(item.icon || '•')}</div>
      <div class="hybrid-summary-body">
        <div class="hybrid-summary-label">${escapeHtml(item.label)}</div>
        <div class="hybrid-summary-value">${escapeHtml(item.value)}</div>
      </div>
    </div>
  `).join("");
}


function buildOperationDiagnosisHtml(operationDiag, nextDiag, lastDiag) {
  const nextDead = nextDiag?.deadVehicleNames?.length ? nextDiag.deadVehicleNames.join(" / ") : "なし";
  return `
    <div class="hybrid-state-row">
      ${buildStatusPill("今", `${getHourLabel(operationDiag.hour)} / ${operationDiag.castCount}名`, operationDiag.statusKey)}
      ${buildStatusPill("次", nextDiag ? `${getHourLabel(nextDiag.hour)} / ${nextDiag.castCount}名` : '-', nextDiag?.statusKey || 'ok')}
      ${buildStatusPill("終", lastDiag ? `${getHourLabel(lastDiag.hour)} / ${lastDiag.castCount}名` : '-', lastDiag?.statusKey || 'ok')}
    </div>
    <div class="hybrid-legend-grid">
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便の状態</div>
        <div class="hybrid-legend-value">${escapeHtml(nextDiag ? nextDiag.statusText : '-')}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便NG車両</div>
        <div class="hybrid-legend-value">${escapeHtml(nextDead)}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便 実効/定員</div>
        <div class="hybrid-legend-value">${nextDiag ? `${nextDiag.effectiveVehicleCount}台 / ${nextDiag.totalCapacity}` : '-'}</div>
      </div>
    </div>
  `;
}


function buildSimulationDiagnosisHtml(diag) {
  return `
    <div class="hybrid-state-row">
      ${buildStatusPill("便", `${getHourLabel(diag.hour)}`, diag.statusKey)}
      ${buildStatusPill("人数", `${diag.castCount}名`, diag.statusKey)}
      ${buildStatusPill("車両", `${diag.effectiveVehicleCount}台`, diag.statusKey)}
      ${buildStatusPill("定員", `${diag.totalCapacity}`, diag.statusKey)}
    </div>
    <div class="hybrid-legend-grid">
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">判定</div>
        <div class="hybrid-legend-value">${escapeHtml(diag.statusText)}${diag.shortageCount > 0 ? ` / 未配車見込み ${diag.shortageCount}名` : ''}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">方面系統</div>
        <div class="hybrid-legend-value">${diag.areaGroupCount}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便NG車両</div>
        <div class="hybrid-legend-value">${escapeHtml(diag.deadVehicleNames.length ? diag.deadVehicleNames.join(' / ') : 'なし')}</div>
      </div>
    </div>
  `;
}


function renderOperationAndSimulationUI() {
  if (getIsRefreshingHybridUISafe()) return;
  setIsRefreshingHybridUISafe(true);
  try {
    const operationBaseHour = getOperationBaseHour();
    const realNextHour = getNextHourSlot(operationBaseHour);
    const lastHour = getLastHourSlot();
    const operationDiag = diagnoseHourWindow(operationBaseHour);
    const nextDiag = Number.isFinite(realNextHour) ? diagnoseHourWindow(realNextHour) : null;
    const lastDiag = Number.isFinite(lastHour) ? diagnoseHourWindow(lastHour) : null;

    if (els.operationContextSummary) {
      els.operationContextSummary.innerHTML = buildSummaryItemsHtml([
        { icon: "📍", label: "今の基準", value: getHourLabel(operationBaseHour), tone: "tone-blue" },
        { icon: "⏭", label: "次に見る便", value: Number.isFinite(realNextHour) ? getHourLabel(realNextHour) : "-", tone: "tone-indigo" },
        { icon: "🏁", label: "最終便", value: Number.isFinite(lastHour) ? getHourLabel(lastHour) : "-", tone: "tone-amber" },
        { icon: "🚐", label: "次便の使える車両", value: nextDiag ? `${nextDiag.effectiveVehicleCount}台` : "-", tone: "tone-green" }
      ]);
    }

    if (els.operationDiagnosis) {
      const cls = nextDiag?.statusKey === "danger" || lastDiag?.statusKey === "danger"
        ? "danger"
        : (nextDiag?.statusKey === "warn" || lastDiag?.statusKey === "warn" ? "warn" : "ok");      els.operationDiagnosis.className = `hybrid-diagnosis ${cls}`;
      els.operationDiagnosis.innerHTML = buildOperationDiagnosisHtml(operationDiag, nextDiag, lastDiag);
    }

    const hourOptions = getUnifiedHourSet();
    if (els.simulationSlotSelect) {
      const currentSimulationSlotHour = getSimulationSlotHourSafe();
      const previous = Number.isFinite(currentSimulationSlotHour) ? currentSimulationSlotHour : (Number.isFinite(realNextHour) ? realNextHour : operationBaseHour);
      setSuppressSimulationSlotChangeSafe(true);
      try {
        const optionHtml = hourOptions.map(hour => `<option value="${hour}">${getHourLabel(hour)}</option>`).join("");
        if (els.simulationSlotSelect.innerHTML !== optionHtml) {
          els.simulationSlotSelect.innerHTML = optionHtml;
        }
        const targetValue = hourOptions.includes(previous) ? previous : (hourOptions[0] ?? operationBaseHour);
        setSimulationSlotHourSafe(targetValue);
        if (String(els.simulationSlotSelect.value) !== String(targetValue)) {
          els.simulationSlotSelect.value = String(targetValue);
        }
      } finally {
        setTimeout(() => { setSuppressSimulationSlotChangeSafe(false); }, 0);
      }
    }

    if (!getLastSimulationResultSafe() && els.simulationDiagnosis) {
      els.simulationDiagnosis.className = 'hybrid-diagnosis muted';
      els.simulationDiagnosis.textContent = '試算対象便を選択してください';
    }
  } catch (error) {
    console.error(error);
    if (els.operationDiagnosis) {
      els.operationDiagnosis.className = 'hybrid-diagnosis danger';
      els.operationDiagnosis.textContent = `診断エラー: ${error?.message || error}`;
    }
  } finally {
    setIsRefreshingHybridUISafe(false);
  }
}


function runSlotDiagnosisPreview() {
  const hour = Number(els.simulationSlotSelect?.value ?? getSimulationSlotHourSafe() ?? getOperationBaseHour());
  setSimulationSlotHourSafe(hour);
  const diag = diagnoseSimulationHourWindow(hour, { includePlanInflow: Boolean(els.simulationIncludePlanInflow?.checked) });
  setLastSimulationResultSafe({ type: 'diagnosis', hour, diag });
  if (els.simulationDiagnosis) {    els.simulationDiagnosis.className = `hybrid-diagnosis ${diag.statusKey}`;
    els.simulationDiagnosis.innerHTML = buildSimulationDiagnosisHtml(diag);
  }
  if (els.simulationPreview) {
    if (!diag.rows.length) {
      els.simulationPreview.className = 'simulation-preview muted';
      els.simulationPreview.textContent = 'この便の予定対象はありません';
    } else {
      const list = diag.rows.map(row => `
        <div class="sim-mini-chip">
          <span class="sim-mini-name">${escapeHtml(row?.casts?.name || '-')}</span>
          <span class="sim-mini-area">${escapeHtml(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || '-'))}</span>
        </div>
      `).join('');
      els.simulationPreview.className = 'simulation-preview';
      els.simulationPreview.innerHTML = `
        <div class="sim-preview-head">
          <h4 class="sim-preview-title">試算対象一覧</h4>
          <span class="chip">${getHourLabel(diag.hour)}</span>
        </div>
        <div class="sim-preview-meta">対象便予定 ${diag.slotPlanCount}名 / 前便未処理予定 ${diag.inflowPlanCount}名</div>
        <div class="sim-mini-chip-grid">${list}</div>
      `;
    }
  }
}

