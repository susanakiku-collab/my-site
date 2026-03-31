// renderSchedule
// Fully restored wrapper/public export version for schedule rendering

(function () {
  function getGlobal(name, fallback) {
    try {
      if (typeof window !== 'undefined' && typeof window[name] !== 'undefined') return window[name];
    } catch (_) {}
    return fallback;
  }

  function setScheduleRendererDeps(deps) {
    try {
      window.__scheduleRendererDeps = deps || {};
    } catch (_) {}
  }

  function resolveScheduleRendererContext() {
    const injected = getGlobal('__scheduleRendererDeps', {}) || {};
    return {
      els: injected.els || getGlobal('els', {}),
      plans: Array.isArray(injected.plans) ? injected.plans : (getGlobal('currentPlansCache', []) || []),
      actuals: Array.isArray(injected.actuals) ? injected.actuals : (getGlobal('currentActualsCache', []) || []),
      helpers: injected.helpers || {},
      actions: injected.actions || {}
    };
  }

  function helper(name, fallback) {
    const ctx = resolveScheduleRendererContext();
    if (ctx.helpers && typeof ctx.helpers[name] === 'function') return ctx.helpers[name];
    const globalFn = getGlobal(name, null);
    if (typeof globalFn === 'function') return globalFn;
    return fallback;
  }

  function getMatrixLegendHtml() {
    return `
      <div class="matrix-legend" aria-label="色の説明">
        <span class="matrix-legend-item"><span class="matrix-dot pending"></span>黄色 = 未完了</span>
        <span class="matrix-legend-item"><span class="matrix-dot done"></span>緑 = 完了</span>
        <span class="matrix-legend-item"><span class="matrix-dot cancel"></span>赤 = キャンセル</span>
      </div>
    `;
  }

  function getPlanLinkedActualStatusCore(planRow, { actuals, els, helpers } = {}) {
    const normalizeStatus = (helpers && helpers.normalizeStatus) || helper('normalizeStatus', s => s || 'pending');
    const todayStr = (helpers && helpers.todayStr) || helper('todayStr', () => new Date().toISOString().slice(0, 10));
    const safeActuals = Array.isArray(actuals) ? actuals : [];
    const linked = safeActuals.find(
      item =>
        Number(item && item.cast_id) === Number(planRow && planRow.cast_id) &&
        String((item && item.actual_date) || (els && els.actualDate && els.actualDate.value) || todayStr()) ===
          String((planRow && planRow.plan_date) || (els && els.planDate && els.planDate.value) || todayStr()) &&
        Number((item && (item.actual_hour ?? item.plan_hour)) ?? -1) === Number((planRow && (planRow.plan_hour ?? planRow.actual_hour)) ?? -1)
    );
    if (linked) return normalizeStatus(linked.status);
    return normalizeStatus(planRow && planRow.status);
  }

  function buildMatrixMetaText(distanceKm, travelMinutes) {
    const parts = [];
    const distance = Number(distanceKm || 0);
    const minutes = Number(travelMinutes || 0);
    if (Number.isFinite(distance) && distance > 0) parts.push(`${distance.toFixed(1)}km`);
    if (Number.isFinite(minutes) && minutes > 0) parts.push(`片道${minutes}分`);
    return parts.length ? ` (${parts.join(' / ')})` : '';
  }

  function buildMatrixNameLine(row, status, addressKey = 'destination_address') {
    const normalizeStatus = helper('normalizeStatus', s => s || 'pending');
    const buildMapLinkHtml = helper('buildMapLinkHtml', ({ name }) => String(name || '-'));
    const escapeHtml = helper('escapeHtml', v => String(v ?? ''));
    const safeStatus = normalizeStatus(status);
    const nameHtml = buildMapLinkHtml({
      name: row && row.casts && row.casts.name,
      address: row && (row[addressKey] || (row.casts && row.casts.address)),
      lat: row && row.casts && row.casts.latitude,
      lng: row && row.casts && row.casts.longitude,
      className: `map-name-link matrix-name status-${safeStatus}`
    });
    const metaText = buildMatrixMetaText(row && row.distance_km, row && row.casts && row.casts.travel_minutes || row && row.travel_minutes);
    const overflowText = Number(row && row.vehicle_id || 0) <= 0 && String(row && row.driver_name || '').includes('あぶれ')
      ? `<span class="matrix-meta" style="margin-left:6px;"> あぶれ</span>`
      : '';
    return `<span class="matrix-line">${nameHtml}<span class="matrix-meta">${escapeHtml(metaText)}</span>${overflowText}</span>`;
  }

  function renderPlanGroupedTableCore({ els, plans, actuals, actions, helpers }) {
    if (!els || !els.plansGroupedTable) return;

    const safePlans = Array.isArray(plans) ? plans : [];
    const getHourLabel = (helpers && helpers.getHourLabel) || helper('getHourLabel', h => `${Number(h)}時`);
    const getGroupedAreasByDisplay = (helpers && helpers.getGroupedAreasByDisplay) || helper('getGroupedAreasByDisplay', (items, pick) => {
      const seen = new Set();
      return (items || []).map(x => {
        const detailArea = helper('normalizeAreaLabel', v => String(v || '無し'))(pick(x));
        if (seen.has(detailArea)) return null;
        seen.add(detailArea);
        return { detailArea };
      }).filter(Boolean);
    });
    const normalizeAreaLabel = (helpers && helpers.normalizeAreaLabel) || helper('normalizeAreaLabel', v => String(v || '無し'));
    const getGroupedAreaHeaderHtml = (helpers && helpers.getGroupedAreaHeaderHtml) || helper('getGroupedAreaHeaderHtml', area => String(area || ''));
    const buildMapLinkHtml = (helpers && helpers.buildMapLinkHtml) || helper('buildMapLinkHtml', ({ name }) => String(name || '-'));
    const escapeHtml = (helpers && helpers.escapeHtml) || helper('escapeHtml', v => String(v ?? ''));
    const getStatusText = (helpers && helpers.getStatusText) || helper('getStatusText', s => String(s || ''));

    if (!safePlans.length) {
      els.plansGroupedTable.innerHTML = `<div class="muted" style="padding:14px;">予定がありません</div>`;
      return;
    }

    const hours = [...new Set(safePlans.map(x => Number(x.plan_hour)))].sort((a, b) => a - b);
    let html = `<div class="grouped-plan-list">`;

    hours.forEach(hour => {
      const hourItems = safePlans.filter(x => Number(x.plan_hour) === hour);
      const groupedAreas = getGroupedAreasByDisplay(hourItems, x => x.planned_area || '無し');

      html += `<div class="grouped-section">`;
      html += `<div class="grouped-hour-title">${getHourLabel(hour)}</div>`;

      groupedAreas.forEach(({ detailArea }) => {
        const areaItems = hourItems.filter(
          x => normalizeAreaLabel(x.planned_area || '無し') === detailArea
        );

        html += `<div class="grouped-area-title">${getGroupedAreaHeaderHtml(detailArea)}</div>`;

        areaItems.forEach(plan => {
          const linkedStatus = getPlanLinkedActualStatusCore(plan, { actuals, els, helpers });
          html += `
            <div class="grouped-row">
              <div>${getHourLabel(hour)}</div>
              <div><strong>${buildMapLinkHtml({
               name: plan && plan.casts && plan.casts.name,
               address: (plan && plan.destination_address) || (plan && plan.casts && plan.casts.address),
               lat: plan && plan.casts && plan.casts.latitude,
               lng: plan && plan.casts && plan.casts.longitude
               })}</strong></div>
               <div>${escapeHtml(normalizeAreaLabel((plan && plan.planned_area) || '無し'))}</div>
               <div>${plan && plan.distance_km != null ? plan.distance_km : ''}</div>
               <div class="op-cell">
                <span class="badge-status ${linkedStatus}">${escapeHtml(getStatusText(linkedStatus))}</span>
                <button class="btn ghost plan-edit-btn" data-id="${plan && plan.id}">編集</button>
                <button class="btn ghost plan-route-btn" data-address="${escapeHtml((plan && plan.destination_address) || (plan && plan.casts && plan.casts.address) || '')}">ルート</button>
                <button class="btn danger plan-delete-btn" data-id="${plan && plan.id}">削除</button>
              </div>
            </div>
          `;
        });
      });

      html += `</div>`;
    });

    html += `</div>`;
    els.plansGroupedTable.innerHTML = html;

    const fillPlanForm = (actions && actions.fillPlanForm) || helper('fillPlanForm', null);
    const openGoogleMap = (actions && actions.openGoogleMap) || helper('openGoogleMap', null);
    const deletePlan = (actions && actions.deletePlan) || helper('deletePlan', null);

    els.plansGroupedTable.querySelectorAll('.plan-edit-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const plan = safePlans.find(x => Number(x.id) === Number(btn.dataset.id));
        if (plan && typeof fillPlanForm === 'function') fillPlanForm(plan);
      });
    });

    els.plansGroupedTable.querySelectorAll('.plan-route-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        if (typeof openGoogleMap === 'function') openGoogleMap(btn.dataset.address || '');
      });
    });

    els.plansGroupedTable.querySelectorAll('.plan-delete-btn').forEach(btn => {
      btn.addEventListener('click', async () => {
        if (typeof deletePlan === 'function') await deletePlan(Number(btn.dataset.id));
      });
    });
  }

  function renderPlansTimeAreaMatrixCore({ els, plans, actuals, helpers }) {
    if (!els || !els.plansTimeAreaMatrix) return;

    const safePlans = Array.isArray(plans) ? plans : [];
    const normalizeAreaLabel = (helpers && helpers.normalizeAreaLabel) || helper('normalizeAreaLabel', v => String(v || '無し'));
    const getAreaDisplayGroup = (helpers && helpers.getAreaDisplayGroup) || helper('getAreaDisplayGroup', v => String(v || '無し'));
    const getHourLabel = (helpers && helpers.getHourLabel) || helper('getHourLabel', h => `${Number(h)}時`);
    const escapeHtml = (helpers && helpers.escapeHtml) || helper('escapeHtml', v => String(v ?? ''));

    const hours = [0, 1, 2, 3, 4, 5];
    const areas = [
      ...new Set(
        safePlans.map(x => getAreaDisplayGroup(normalizeAreaLabel(x.planned_area || '無し')))
      )
    ];

    if (!areas.length) {
      els.plansTimeAreaMatrix.innerHTML = `<div class="muted" style="padding:14px;">一覧がありません</div>`;
      return;
    }

    let html = `
      ${getMatrixLegendHtml()}
      <table class="matrix-table">
        <thead>
          <tr>
            <th>時間</th>
            ${areas.map(area => `<th>${escapeHtml(area)}</th>`).join('')}
          </tr>
        </thead>
        <tbody>
    `;

    hours.forEach(hour => {
      html += `<tr><td>${getHourLabel(hour)}</td>`;

      areas.forEach(area => {
        const rows = safePlans.filter(
          plan =>
            Number((plan && plan.plan_hour) ?? 0) === hour &&
            getAreaDisplayGroup(normalizeAreaLabel((plan && plan.planned_area) || '無し')) === area
        );

        if (!rows.length) {
          html += `<td>-</td>`;
        } else {
          html += `
            <td>
              <div class="matrix-card">
                ${rows.map(row => `
                  <div class="matrix-item">
                    ${buildMatrixNameLine(row, getPlanLinkedActualStatusCore(row, { actuals, els, helpers }), 'destination_address')}
                    <div class="matrix-subarea">${escapeHtml(normalizeAreaLabel((row && row.planned_area) || '無し'))}</div>
                  </div>
                `).join('')}
              </div>
            </td>
          `;
        }
      });

      html += `</tr>`;
    });

    html += `</tbody></table>`;
    els.plansTimeAreaMatrix.innerHTML = html;
  }

  function renderPlanGroupedTable() {
    const ctx = resolveScheduleRendererContext();
    return renderPlanGroupedTableCore(ctx);
  }

  function renderPlansTimeAreaMatrix() {
    const ctx = resolveScheduleRendererContext();
    return renderPlansTimeAreaMatrixCore(ctx);
  }

  window.setScheduleRendererDeps = setScheduleRendererDeps;
  window.getMatrixLegendHtml = getMatrixLegendHtml;
  window.buildMatrixNameLine = buildMatrixNameLine;
  window.getPlanLinkedActualStatusCore = getPlanLinkedActualStatusCore;
  window.renderPlanGroupedTableCore = renderPlanGroupedTableCore;
  window.renderPlansTimeAreaMatrixCore = renderPlansTimeAreaMatrixCore;
  window.renderPlanGroupedTable = renderPlanGroupedTable;
  window.renderPlansTimeAreaMatrix = renderPlansTimeAreaMatrix;
})();
