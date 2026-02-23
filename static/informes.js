document.addEventListener('DOMContentLoaded', () => {
  const serviceAutoInput = document.getElementById('reportServiceAuto');
  const periodSelect = document.getElementById('reportPeriodSelect');
  const runSelect = document.getElementById('reportRunSelect');
  const receiverInput = document.getElementById('reportReceiver');
  const responsiblesList = document.getElementById('reportResponsiblesList');
  const observationInput = document.getElementById('reportObservation');
  const reportDateInput = document.getElementById('reportDate');
  const warehouseLeadInput = document.getElementById('reportWarehouseLead');
  const assetsManagerInput = document.getElementById('reportAssetsManager');
  const refreshBtn = document.getElementById('refreshReportFiltersBtn');
  const refreshA22HistoryBtn = document.getElementById('refreshA22HistoryBtn');
  const exportExcelBtn = document.getElementById('exportA22ExcelBtn');
  const exportPdfBtn = document.getElementById('exportA22PdfBtn');
  const exportAccountingMonthlyBtn = document.getElementById('exportAccountingMonthlyBtn');
  const statusEl = document.getElementById('reportStatus');
  const accountingStatusEl = document.getElementById('accountingStatus');
  const accountingMonthInput = document.getElementById('accountingMonth');
  const accountingYearInput = document.getElementById('accountingYear');
  const accountingReportTitleInput = document.getElementById('accountingReportTitle');
  const accountingGeneratedByInput = document.getElementById('accountingGeneratedBy');
  const refreshAccountingHistoryBtn = document.getElementById('refreshAccountingHistoryBtn');
  const accountingHistoryContainer = document.getElementById('accountingHistoryContainer');
  const a22HistoryContainer = document.getElementById('a22HistoryContainer');
  const assetObsSearch = document.getElementById('assetObsSearch');
  const assetObsContainer = document.getElementById('assetObsContainer');
  const tabButtons = Array.from(document.querySelectorAll('[data-report-tab]'));
  const panelA22 = document.getElementById('panelA22');
  const panelAccounting = document.getElementById('panelAccounting');

  let serviceAssets = [];
  const assetObservationMap = {};
  const runsById = new Map();
  const HISTORY_PAGE_SIZE = 10;
  const ASSET_OBS_PAGE_SIZE = 10;
  const historyPageByScope = {
    accounting: 1,
    a22: 1,
  };
  let assetObsPage = 1;

  function switchTab(tabName) {
    const isA22 = tabName === 'a22';
    panelA22?.classList.toggle('active', isA22);
    panelAccounting?.classList.toggle('active', !isA22);
    tabButtons.forEach((btn) => {
      const active = btn.dataset.reportTab === tabName;
      btn.classList.toggle('active', active);
      btn.setAttribute('aria-selected', active ? 'true' : 'false');
    });
  }

  function accountingPeriodParams() {
    const month = Number(accountingMonthInput?.value || 0);
    const year = Number(accountingYearInput?.value || 0);
    return {
      month: (month >= 1 && month <= 12) ? month : null,
      year: (year >= 2000 && year <= 2100) ? year : null,
    };
  }

  function accountingReportTitle() {
    return String(accountingReportTitleInput?.value || '').trim();
  }

  function accountingGeneratedBy() {
    return String(accountingGeneratedByInput?.value || '').trim();
  }

  function buildHistoryRows(rows, scope) {
    if (scope === 'accounting') {
      return rows.map((r) => `
        <tr>
          <td>${App.escapeHtml(String(r.id || ''))}</td>
          <td class="cell-clip" title="${App.escapeHtml(r.title || '')}">${App.escapeHtml(r.title || '')}</td>
          <td class="cell-clip" title="${App.escapeHtml(r.period_label || '-')}">${App.escapeHtml(r.period_label || '-')}</td>
          <td>${App.escapeHtml(String(r.generated_at || '').replace('T', ' ').slice(0, 16))}</td>
          <td class="cell-clip" title="${App.escapeHtml(r.file_name || '')}">${App.escapeHtml(r.file_name || '')}</td>
          <td><button type="button" class="mini-btn" data-report-id="${App.escapeHtml(String(r.id || ''))}">Descargar</button></td>
        </tr>
      `).join('');
    }
    return rows.map((r) => `
      <tr>
        <td>${App.escapeHtml(String(r.id || ''))}</td>
        <td>${App.escapeHtml(r.report_type === 'a22_pdf' ? 'A22 PDF' : 'A22 Excel')}</td>
        <td class="cell-clip" title="${App.escapeHtml(r.period_label || '-')}">${App.escapeHtml(r.period_label || '-')}</td>
        <td>${App.escapeHtml(String(r.generated_at || '').replace('T', ' ').slice(0, 16))}</td>
        <td class="cell-clip" title="${App.escapeHtml(r.file_name || '')}">${App.escapeHtml(r.file_name || '')}</td>
        <td><button type="button" class="mini-btn" data-a22-report-id="${App.escapeHtml(String(r.id || ''))}">Descargar</button></td>
      </tr>
    `).join('');
  }

  function renderHistoryTable(container, rows, scope) {
    if (!container) return;
    if (!rows.length) {
      container.innerHTML = scope === 'accounting'
        ? '<div class="empty-mini">No hay informes generados todavia.</div>'
        : '<div class="empty-mini">No hay informes A22 generados todavia.</div>';
      return;
    }

    const totalPages = Math.max(1, Math.ceil(rows.length / HISTORY_PAGE_SIZE));
    const requestedPage = Number(historyPageByScope[scope] || 1);
    const currentPage = Math.min(Math.max(requestedPage, 1), totalPages);
    historyPageByScope[scope] = currentPage;
    const startIdx = (currentPage - 1) * HISTORY_PAGE_SIZE;
    const pageRows = rows.slice(startIdx, startIdx + HISTORY_PAGE_SIZE);
    const from = startIdx + 1;
    const to = Math.min(startIdx + HISTORY_PAGE_SIZE, rows.length);

    const headCols = scope === 'accounting'
      ? '<th>ID</th><th>TITULO</th><th>PERIODO</th><th>GENERADO</th><th>ARCHIVO</th><th>ACCION</th>'
      : '<th>ID</th><th>TIPO</th><th>PERIODO</th><th>GENERADO</th><th>ARCHIVO</th><th>ACCION</th>';

    container.innerHTML = `
      <div class="history-table-wrap">
      <table class="report-history-table">
        <thead>
          <tr>${headCols}</tr>
        </thead>
        <tbody>
          ${buildHistoryRows(pageRows, scope)}
        </tbody>
      </table>
      </div>
      <div class="history-pagination">
        <div class="history-page-meta">Mostrando ${from}-${to} de ${rows.length}</div>
        <div class="history-page-controls">
          <button
            type="button"
            class="mini-btn history-page-btn"
            data-history-scope="${scope}"
            data-history-action="prev"
            ${currentPage <= 1 ? 'disabled' : ''}
          >Anterior</button>
          <span class="history-page-indicator">Pagina ${currentPage} de ${totalPages}</span>
          <button
            type="button"
            class="mini-btn history-page-btn"
            data-history-scope="${scope}"
            data-history-action="next"
            ${currentPage >= totalPages ? 'disabled' : ''}
          >Siguiente</button>
        </div>
      </div>
    `;
  }

  function moveHistoryPage(scope, action) {
    const current = Number(historyPageByScope[scope] || 1);
    historyPageByScope[scope] = action === 'prev' ? current - 1 : current + 1;
    if (scope === 'accounting') {
      loadAccountingHistory().catch((err) => App.setStatus(accountingStatusEl || statusEl, err.message, true));
    } else {
      loadA22History().catch((err) => App.setStatus(statusEl, err.message, true));
    }
  }

  async function loadAccountingHistory() {
    if (!accountingHistoryContainer) return;
    const data = await App.get('/reports/accounting_monthly_history');
    const rows = data.items || [];
    renderHistoryTable(accountingHistoryContainer, rows, 'accounting');
  }

  async function loadA22History() {
    if (!a22HistoryContainer) return;
    const periodId = periodSelect?.value ? Number(periodSelect.value) : null;
    if (!periodId) {
      renderHistoryTable(a22HistoryContainer, [], 'a22');
      return;
    }
    const data = await App.get(`/reports/a22_history?period_id=${encodeURIComponent(String(periodId))}`);
    const rows = data.items || [];
    renderHistoryTable(a22HistoryContainer, rows, 'a22');
  }

  async function loadResponsibles() {
    const data = await App.get('/responsibles');
    responsiblesList.innerHTML = (data.responsibles || [])
      .map((name) => `<option value="${App.escapeHtml(name)}"></option>`)
      .join('');
  }

  async function loadRunsForPeriod() {
    if (!runSelect) return;
    const pid = periodSelect?.value ? Number(periodSelect.value) : null;
    if (!pid) {
      runsById.clear();
      runSelect.innerHTML = '<option value="">-- Selecciona jornada --</option>';
      updateServiceFromRun();
      serviceAssets = [];
      renderAssetObsTable();
      return;
    }
    const params = new URLSearchParams();
    if (pid) params.set('period_id', String(pid));
    const data = await App.get('/runs' + (params.toString() ? `?${params.toString()}` : ''));
    const runs = (data.runs || []);
    runsById.clear();
    runs.forEach((r) => runsById.set(Number(r.id), r));
    runSelect.innerHTML = '<option value="">-- Selecciona jornada --</option>' + runs
      .map((r) => {
        const svc = r.service ? ` [${App.escapeHtml(r.service)}]` : '';
        const st = r.status === 'active' ? 'Activa' : (r.status === 'cancelled' ? 'Anulada' : 'Cerrada');
        return `<option value="${r.id}">${r.id} - ${App.escapeHtml(r.name)}${svc} - ${st}</option>`;
      }).join('');
    if (runs.length === 1) runSelect.value = String(runs[0].id);
    updateServiceFromRun();
  }

  function selectedRun() {
    const id = runSelect?.value ? Number(runSelect.value) : null;
    if (!id) return null;
    return runsById.get(id) || null;
  }

  function updateServiceFromRun() {
    const run = selectedRun();
    const service = run?.service || '';
    if (serviceAutoInput) serviceAutoInput.value = service;
    return service;
  }

  function renderAssetObsTable() {
    if (!updateServiceFromRun()) {
      assetObsContainer.innerHTML = '<div class="empty-mini">Selecciona una jornada para cargar codigos.</div>';
      return;
    }
    const q = String(assetObsSearch?.value || '').trim().toUpperCase();
    const rows = serviceAssets.filter((a) => {
      if (!q) return true;
      return String(a.C_ACT || '').toUpperCase().includes(q) || String(a.NOM || '').toUpperCase().includes(q);
    });
    if (!rows.length) {
      assetObsContainer.innerHTML = '<div class="empty-mini">No hay activos para el filtro indicado.</div>';
      return;
    }
    const totalPages = Math.max(1, Math.ceil(rows.length / ASSET_OBS_PAGE_SIZE));
    const currentPage = Math.min(Math.max(assetObsPage, 1), totalPages);
    assetObsPage = currentPage;
    const startIdx = (currentPage - 1) * ASSET_OBS_PAGE_SIZE;
    const pageRows = rows.slice(startIdx, startIdx + ASSET_OBS_PAGE_SIZE);
    const from = startIdx + 1;
    const to = Math.min(startIdx + ASSET_OBS_PAGE_SIZE, rows.length);
    assetObsContainer.innerHTML = `
      <table class="obs-table">
        <thead>
          <tr>
            <th>COD ACTIVO</th>
            <th>DESCRIPCION ACTIVO</th>
            <th>OBSERVACION ESPECIFICA</th>
          </tr>
        </thead>
        <tbody>
          ${pageRows.map((a) => {
            const code = String(a.C_ACT || '');
            const current = assetObservationMap[code] ?? '';
            return `<tr>
              <td>${App.escapeHtml(code)}</td>
              <td>${App.escapeHtml(a.NOM)}</td>
              <td>
                <input
                  class="obs-input"
                  data-code="${App.escapeHtml(code)}"
                  value="${App.escapeHtml(current)}"
                  placeholder="Observacion para este activo (opcional)"
                />
              </td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
      <div class="history-pagination">
        <div class="history-page-meta">Mostrando ${from}-${to} de ${rows.length}</div>
        <div class="history-page-controls">
          <button
            type="button"
            class="mini-btn history-page-btn"
            data-asset-obs-action="prev"
            ${currentPage <= 1 ? 'disabled' : ''}
          >Anterior</button>
          <span class="history-page-indicator">Pagina ${currentPage} de ${totalPages}</span>
          <button
            type="button"
            class="mini-btn history-page-btn"
            data-asset-obs-action="next"
            ${currentPage >= totalPages ? 'disabled' : ''}
          >Siguiente</button>
        </div>
      </div>
    `;
  }

  async function loadAssetsForService() {
    const service = updateServiceFromRun();
    serviceAssets = [];
    assetObsPage = 1;
    Object.keys(assetObservationMap).forEach((k) => delete assetObservationMap[k]);
    if (!service) {
      assetObsContainer.innerHTML = '<div class="empty-mini">Selecciona primero una jornada para cargar codigos.</div>';
      return;
    }
    const data = await App.get(`/assets?service=${encodeURIComponent(service)}`);
    serviceAssets = data.assets || [];
    renderAssetObsTable();
  }

  function buildPayload() {
    const perAssetObservations = {};
    Object.keys(assetObservationMap).forEach((code) => {
      const v = String(assetObservationMap[code] || '').trim();
      if (v) perAssetObservations[code] = v;
    });
    const run = selectedRun();
    const service = (run?.service || '').trim();
    const receiver = (receiverInput?.value || '').trim();
    const observation = (observationInput?.value || '').trim();
    const reportDate = (reportDateInput?.value || '').trim();
    const warehouseLead = (warehouseLeadInput?.value || '').trim();
    const assetsManager = (assetsManagerInput?.value || '').trim();
    const periodId = (periodSelect?.value || '').trim();
    const runId = (runSelect?.value || '').trim();
    return {
      service,
      period_id: periodId ? Number(periodId) : null,
      run_id: runId ? Number(runId) : null,
      receiver,
      observation,
      report_date: reportDate,
      warehouse_lead: warehouseLead,
      assets_manager: assetsManager,
      per_asset_observations: perAssetObservations,
    };
  }

  function validateRequiredFields() {
    const periodId = (periodSelect?.value || '').trim();
    const runId = (runSelect?.value || '').trim();
    const service = (updateServiceFromRun() || '').trim();
    const warehouseLead = (warehouseLeadInput?.value || '').trim();
    const assetsManager = (assetsManagerInput?.value || '').trim();
    if (!periodId) {
      App.setStatus(statusEl, 'Debes seleccionar el periodo de inventario.', true);
      periodSelect?.focus();
      return false;
    }
    if (!runId) {
      App.setStatus(statusEl, 'Debes seleccionar la jornada del periodo.', true);
      runSelect?.focus();
      return false;
    }
    if (!service) {
      App.setStatus(statusEl, 'La jornada seleccionada no tiene servicio asociado.', true);
      runSelect?.focus();
      return false;
    }
    if (!warehouseLead) {
      App.setStatus(statusEl, 'Debes escribir el lider de almacen.', true);
      warehouseLeadInput?.focus();
      return false;
    }
    if (!assetsManager) {
      App.setStatus(statusEl, 'Debes escribir el responsable de activos fijos.', true);
      assetsManagerInput?.focus();
      return false;
    }
    return true;
  }

  async function exportFile(url, fallbackName) {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(buildPayload()),
    });
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || `Error HTTP ${res.status}`);
    }
    const blob = await res.blob();
    const cd = res.headers.get('Content-Disposition') || '';
    const match = cd.match(/filename="?([^"]+)"?/i);
    const filename = (match && match[1]) || fallbackName;
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(link.href);
  }

  async function reloadFilters() {
    await Promise.all([
      App.loadPeriods(periodSelect),
      loadResponsibles(),
    ]);
    await loadRunsForPeriod();
    await loadAssetsForService();
  }

  refreshBtn?.addEventListener('click', async () => {
    try {
      await reloadFilters();
      App.setStatus(statusEl, 'Filtros actualizados');
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  exportExcelBtn?.addEventListener('click', async () => {
    try {
      if (!validateRequiredFields()) return;
      await exportFile('/export', 'A22.xlsx');
      App.setStatus(statusEl, 'A22 Excel generado');
      await loadA22History();
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  exportPdfBtn?.addEventListener('click', async () => {
    try {
      if (!validateRequiredFields()) return;
      await exportFile('/export_a22_pdf', 'A22.pdf');
      App.setStatus(statusEl, 'A22 PDF generado');
      await loadA22History();
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  exportAccountingMonthlyBtn?.addEventListener('click', async () => {
    try {
      const reportTitle = accountingReportTitle();
      if (!reportTitle) {
        App.setStatus(accountingStatusEl || statusEl, 'Debes escribir el titulo del informe contable.', true);
        accountingReportTitleInput?.focus();
        return;
      }
      const generatedBy = accountingGeneratedBy();
      if (!generatedBy) {
        App.setStatus(accountingStatusEl || statusEl, 'Debes escribir el usuario que genera el informe.', true);
        accountingGeneratedByInput?.focus();
        return;
      }
      App.setStatus(accountingStatusEl || statusEl, 'Generando informe contable mensual...');
      const qp = accountingPeriodParams();
      const params = new URLSearchParams();
      if (qp.month) params.set('month', String(qp.month));
      if (qp.year) params.set('year', String(qp.year));
      params.set('report_title', reportTitle);
      params.set('generated_by', generatedBy);
      params.set('refresh', '1');
      const urlReq = '/reports/accounting_monthly_excel' + (params.toString() ? `?${params.toString()}` : '');
      const res = await fetch(urlReq);
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || 'No fue posible generar el informe contable');
      }
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      const disposition = res.headers.get('Content-Disposition') || '';
      const match = disposition.match(/filename=\"?([^\";]+)\"?/i);
      a.href = url;
      a.download = match ? match[1] : `informe_contabilidad_mensual_${Date.now()}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
      App.setStatus(accountingStatusEl || statusEl, 'Informe contable mensual generado correctamente.');
      await loadAccountingHistory();
    } catch (err) {
      App.setStatus(accountingStatusEl || statusEl, err.message, true);
    }
  });

  refreshAccountingHistoryBtn?.addEventListener('click', () => {
    loadAccountingHistory().catch((err) => App.setStatus(accountingStatusEl || statusEl, err.message, true));
  });
  refreshA22HistoryBtn?.addEventListener('click', () => {
    loadA22History().catch((err) => App.setStatus(statusEl, err.message, true));
  });

  tabButtons.forEach((btn) => {
    btn.addEventListener('click', () => switchTab(btn.dataset.reportTab || 'a22'));
  });

  periodSelect?.addEventListener('change', () => {
    loadRunsForPeriod()
      .then(loadAssetsForService)
      .then(() => Promise.all([
        loadAccountingHistory().catch(() => {}),
        loadA22History().catch(() => {}),
      ]))
      .then(() => {
        if (periodSelect?.value) {
          App.setStatus(statusEl, 'Periodo seleccionado. Ahora selecciona una jornada para continuar.');
        } else {
          App.setStatus(statusEl, 'Selecciona un periodo para consultar jornadas y activos.', true);
        }
      })
      .catch((err) => App.setStatus(statusEl, err.message, true));
  });
  runSelect?.addEventListener('change', () => {
    updateServiceFromRun();
    loadAssetsForService().catch((err) => App.setStatus(statusEl, err.message, true));
  });
  assetObsSearch?.addEventListener('input', () => {
    assetObsPage = 1;
    renderAssetObsTable();
  });

  assetObsContainer?.addEventListener('input', (e) => {
    const input = e.target.closest('input.obs-input[data-code]');
    if (!input) return;
    assetObservationMap[input.dataset.code] = input.value || '';
  });
  assetObsContainer?.addEventListener('click', (e) => {
    const navBtn = e.target.closest('button[data-asset-obs-action]');
    if (!navBtn) return;
    const action = navBtn.getAttribute('data-asset-obs-action');
    assetObsPage = action === 'prev' ? assetObsPage - 1 : assetObsPage + 1;
    renderAssetObsTable();
  });

  accountingHistoryContainer?.addEventListener('click', (e) => {
    const navBtn = e.target.closest('button[data-history-action][data-history-scope]');
    if (navBtn) {
      moveHistoryPage(navBtn.getAttribute('data-history-scope'), navBtn.getAttribute('data-history-action'));
      return;
    }
    const btn = e.target.closest('button[data-report-id]');
    if (!btn) return;
    const reportId = btn.getAttribute('data-report-id');
    if (!reportId) return;
    window.open(`/reports/accounting_monthly_history/${encodeURIComponent(reportId)}/download`, '_blank');
  });
  a22HistoryContainer?.addEventListener('click', (e) => {
    const navBtn = e.target.closest('button[data-history-action][data-history-scope]');
    if (navBtn) {
      moveHistoryPage(navBtn.getAttribute('data-history-scope'), navBtn.getAttribute('data-history-action'));
      return;
    }
    const btn = e.target.closest('button[data-a22-report-id]');
    if (!btn) return;
    const reportId = btn.getAttribute('data-a22-report-id');
    if (!reportId) return;
    window.open(`/reports/a22_history/${encodeURIComponent(reportId)}/download`, '_blank');
  });

  reloadFilters()
    .then(() => {
      switchTab('a22');
      const today = new Date();
      if (accountingMonthInput && !accountingMonthInput.value) {
        accountingMonthInput.value = String(today.getMonth() + 1);
      }
      if (accountingYearInput && !accountingYearInput.value) {
        accountingYearInput.value = String(today.getFullYear());
      }
      if (accountingReportTitleInput && !accountingReportTitleInput.value) {
        const monthNames = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
        accountingReportTitleInput.value = `INFORME CONCILIACION ACTIVOS FIJOS-CONTABILIDAD ${monthNames[today.getMonth()]} ${today.getFullYear()}`;
      }
      if (reportDateInput && !reportDateInput.value) {
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const dd = String(today.getDate()).padStart(2, '0');
        reportDateInput.value = `${yyyy}-${mm}-${dd}`;
      }
      if (warehouseLeadInput && !warehouseLeadInput.value) warehouseLeadInput.value = 'YUDI ELENA CAMBINDO MINA';
      loadAccountingHistory().catch(() => {});
      loadA22History().catch(() => {});
      if (!periodSelect?.value) {
        App.setStatus(statusEl, 'Selecciona un periodo para consultar jornadas y activos.', true);
      } else {
        App.setStatus(statusEl, 'Listo para generar A22');
      }
    })
    .catch((err) => App.setStatus(statusEl, err.message, true));
});
