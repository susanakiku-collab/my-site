// renderActual
// dashboard.js から安全に分離した実績表描画系
// core は依存注入で動作し、window へラッパー関数を公開する

(function(){
  const actualRendererDeps = {
    getEls: () => (typeof els !== "undefined" ? els : null),
    getActuals: () => (typeof currentActualsCache !== "undefined" ? currentActualsCache : []),
    getActions: () => ({
      fillActualForm: typeof fillActualForm === "function" ? fillActualForm : null,
      openGoogleMap: typeof openGoogleMap === "function" ? openGoogleMap : null,
      deleteActual: typeof deleteActual === "function" ? deleteActual : null,
      updateActualStatus: typeof updateActualStatus === "function" ? updateActualStatus : null
    }),
    getHelpers: () => ({
      normalizeStatus: typeof normalizeStatus === "function" ? normalizeStatus : (v => v),
      getHourLabel: typeof getHourLabel === "function" ? getHourLabel : (h => `${h}時`),
      getGroupedAreasByDisplay: typeof getGroupedAreasByDisplay === "function" ? getGroupedAreasByDisplay : (() => []),
      getGroupedAreaHeaderHtml: typeof getGroupedAreaHeaderHtml === "function" ? getGroupedAreaHeaderHtml : (v => String(v || "")),
      buildMapLinkHtml: typeof buildMapLinkHtml === "function" ? buildMapLinkHtml : ({ name }) => String(name || "-"),
      escapeHtml: typeof escapeHtml === "function" ? escapeHtml : (v => String(v ?? "")),
      normalizeAreaLabel: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel : (v => String(v || "")),
      getStatusText: typeof getStatusText === "function" ? getStatusText : (v => String(v || "")),
      getAreaDisplayGroup: typeof getAreaDisplayGroup === "function" ? getAreaDisplayGroup : (v => String(v || "")),
      getMatrixLegendHtml: typeof getMatrixLegendHtml === "function" ? getMatrixLegendHtml : (() => ""),
      buildMatrixNameLine: typeof buildMatrixNameLine === "function" ? buildMatrixNameLine : (() => "")
    })
  };

  function setActualRendererDeps(nextDeps = {}) {
    Object.assign(actualRendererDeps, nextDeps || {});
  }

  function resolveActualRendererContext() {
    const elsObj = typeof actualRendererDeps.getEls === "function" ? actualRendererDeps.getEls() : null;
    const actuals = typeof actualRendererDeps.getActuals === "function" ? actualRendererDeps.getActuals() : [];
    const actions = typeof actualRendererDeps.getActions === "function" ? actualRendererDeps.getActions() : {};
    const helpers = typeof actualRendererDeps.getHelpers === "function" ? actualRendererDeps.getHelpers() : {};
    return {
      els: elsObj || {},
      actuals: Array.isArray(actuals) ? actuals : [],
      actions: actions || {},
      helpers: helpers || {}
    };
  }

  function renderActualTableCore({ els, actuals, actions, helpers }) {
    if (!els.actualTableWrap) return;

    const {
      normalizeStatus,
      getHourLabel,
      getGroupedAreasByDisplay,
      getGroupedAreaHeaderHtml,
      buildMapLinkHtml,
      escapeHtml,
      normalizeAreaLabel,
      getStatusText
    } = helpers;

    const actionableItems = actuals.filter(item => {
      const status = normalizeStatus(item?.status);
      return status !== "done" && status !== "cancel";
    });

    if (!actionableItems.length) {
      els.actualTableWrap.innerHTML = `<div class="muted" style="padding:14px;">配車決定で対応待ちのActualはありません</div>`;
      return;
    }

    const hours = [...new Set(actionableItems.map(x => Number(x?.actual_hour ?? 0)))].sort((a, b) => a - b);
    let html = `<div class="grouped-actual-list">`;

    hours.forEach(hour => {
      const hourItems = actionableItems.filter(x => Number(x?.actual_hour ?? 0) === hour);
      const groupedAreas = getGroupedAreasByDisplay(hourItems, x => x?.destination_area || "無し");

      html += `<div class="grouped-section">`;
      html += `<div class="grouped-hour-title">${getHourLabel(hour)}</div>`;

      groupedAreas.forEach(({ detailArea }) => {
        const areaItems = hourItems.filter(
          item => normalizeAreaLabel(item?.destination_area || "無し") === detailArea
        );

        if (!areaItems.length) return;

        html += `<div class="grouped-area-title">${getGroupedAreaHeaderHtml(detailArea)}</div>`;

        areaItems.forEach(item => {
          html += `
            <div class="grouped-row">
              <div>${getHourLabel(hour)}</div>
              <div><strong>${buildMapLinkHtml({
                name: item?.casts?.name,
                address: item?.destination_address || item?.casts?.address,
                lat: item?.casts?.latitude,
                lng: item?.casts?.longitude
              })}</strong></div>
              <div>${escapeHtml(normalizeAreaLabel(item?.destination_area || "無し"))}</div>
              <div>${item?.distance_km ?? ""}</div>
              <div class="op-cell">
                <div class="state-stack">
                  <button class="btn primary actual-done-btn" data-id="${item?.id}">完了</button>
                  <button class="btn danger actual-cancel-btn" data-id="${item?.id}">キャンセル</button>
                  <span class="badge-status ${normalizeStatus(item?.status)}">${escapeHtml(getStatusText(item?.status))}</span>
                </div>
                <button class="btn ghost actual-edit-btn" data-id="${item?.id}">編集</button>
                <button class="btn ghost actual-route-btn" data-address="${escapeHtml(item?.destination_address || item?.casts?.address || "")}">ルート</button>
                <button class="btn danger actual-delete-btn" data-id="${item?.id}">削除</button>
              </div>
            </div>
          `;
        });
      });

      html += `</div>`;
    });

    html += `</div>`;
    els.actualTableWrap.innerHTML = html;

    els.actualTableWrap.querySelectorAll(".actual-edit-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const item = actuals.find(x => Number(x?.id) === Number(btn.dataset.id));
        if (item && typeof actions.fillActualForm === "function") actions.fillActualForm(item);
      });
    });

    els.actualTableWrap.querySelectorAll(".actual-route-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        if (typeof actions.openGoogleMap === "function") actions.openGoogleMap(btn.dataset.address || "");
      });
    });

    els.actualTableWrap.querySelectorAll(".actual-delete-btn").forEach(btn => {
      btn.addEventListener("click", async () => {
        if (typeof actions.deleteActual === "function") await actions.deleteActual(Number(btn.dataset.id));
      });
    });

    els.actualTableWrap.querySelectorAll(".actual-done-btn").forEach(btn => {
      btn.addEventListener("click", async () => {
        if (typeof actions.updateActualStatus === "function") await actions.updateActualStatus(Number(btn.dataset.id), "done");
      });
    });

    els.actualTableWrap.querySelectorAll(".actual-cancel-btn").forEach(btn => {
      btn.addEventListener("click", async () => {
        if (typeof actions.updateActualStatus === "function") await actions.updateActualStatus(Number(btn.dataset.id), "cancel");
      });
    });
  }

  function renderActualTimeAreaMatrixCore({ els, actuals, helpers }) {
    if (!els.actualTimeAreaMatrix) return;

    const {
      getAreaDisplayGroup,
      normalizeAreaLabel,
      escapeHtml,
      getHourLabel,
      getMatrixLegendHtml,
      buildMatrixNameLine,
      normalizeStatus
    } = helpers;

    const hours = [0, 1, 2, 3, 4, 5];
    const areas = [
      ...new Set(
        actuals.map(x => getAreaDisplayGroup(normalizeAreaLabel(x?.destination_area || "無し")))
      )
    ].filter(Boolean);

    if (!areas.length) {
      els.actualTimeAreaMatrix.innerHTML = `<div class="muted" style="padding:14px;">一覧がありません</div>`;
      return;
    }

    let html = `
      ${getMatrixLegendHtml()}
      <table class="matrix-table">
        <thead>
          <tr>
            <th>時間</th>
            ${areas.map(a => `<th>${escapeHtml(a)}</th>`).join("")}
          </tr>
        </thead>
        <tbody>
    `;

    hours.forEach(hour => {
      html += `<tr><td>${getHourLabel(hour)}</td>`;

      areas.forEach(area => {
        const rows = actuals.filter(
          x =>
            Number(x?.actual_hour ?? 0) === hour &&
            getAreaDisplayGroup(normalizeAreaLabel(x?.destination_area || "無し")) === area
        );

        if (!rows.length) {
          html += `<td>-</td>`;
        } else {
          html += `
            <td>
              <div class="matrix-card">
                ${rows.map(row => `
                  <div class="matrix-item">
                    ${buildMatrixNameLine(row, normalizeStatus(row?.status), "destination_address")}
                  </div>
                `).join("")}
              </div>
            </td>
          `;
        }
      });

      html += `</tr>`;
    });

    html += `</tbody></table>`;
    els.actualTimeAreaMatrix.innerHTML = html;
  }

  function renderActualTable() {
    return renderActualTableCore(resolveActualRendererContext());
  }

  function renderActualTimeAreaMatrix() {
    return renderActualTimeAreaMatrixCore(resolveActualRendererContext());
  }

  window.setActualRendererDeps = setActualRendererDeps;
  window.renderActualTableCore = renderActualTableCore;
  window.renderActualTimeAreaMatrixCore = renderActualTimeAreaMatrixCore;
  window.renderActualTable = renderActualTable;
  window.renderActualTimeAreaMatrix = renderActualTimeAreaMatrix;
})();
