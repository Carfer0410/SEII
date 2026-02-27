document.addEventListener('DOMContentLoaded', () => {
  const tabButtons = Array.from(document.querySelectorAll('[data-format-tab]'));
  const panelA22 = document.getElementById('panelA22');
  const panelClearance = document.getElementById('panelClearance');

  // A22
  const serviceAutoInput = document.getElementById('reportServiceAuto');
  const periodSelect = document.getElementById('reportPeriodSelect');
  const runSelect = document.getElementById('reportRunSelect');
  const receiverInput = document.getElementById('reportReceiver');
  const a22ResponsiblesList = document.getElementById('reportResponsiblesList');
  const observationInput = document.getElementById('reportObservation');
  const reportDateInput = document.getElementById('reportDate');
  const warehouseLeadInput = document.getElementById('reportWarehouseLead');
  const assetsManagerInput = document.getElementById('reportAssetsManager');
  const refreshBtn = document.getElementById('refreshReportFiltersBtn');
  const refreshA22HistoryBtn = document.getElementById('refreshA22HistoryBtn');
  const exportExcelBtn = document.getElementById('exportA22ExcelBtn');
  const exportPdfBtn = document.getElementById('exportA22PdfBtn');
  const a22StatusEl = document.getElementById('reportStatus');
  const a22HistoryContainer = document.getElementById('a22HistoryContainer');
  const assetObsSearch = document.getElementById('assetObsSearch');
  const assetObsContainer = document.getElementById('assetObsContainer');

  // Paz y salvo
  const clearancePeriodSelect = document.getElementById('formatPeriodSelect');
  const clearanceRunSelect = document.getElementById('formatRunSelect');
  const clearanceDateInput = document.getElementById('formatDate');
  const clearanceOutgoingInput = document.getElementById('formatOutgoing');
  const clearanceIncomingInput = document.getElementById('formatIncoming');
  const clearanceIssuedByInput = document.getElementById('formatIssuedBy');
  const clearanceNotesInput = document.getElementById('formatNotes');
  const clearanceValidateBtn = document.getElementById('formatValidateBtn');
  const clearanceGenerateBtn = document.getElementById('formatGenerateBtn');
  const clearanceStatusEl = document.getElementById('formatStatus');
  const clearanceResponsiblesList = document.getElementById('formatResponsiblesList');

  const runsById = new Map();
  let serviceAssets = [];
  const assetObservationMap = {};
  const HISTORY_PAGE_SIZE = 10;
  const ASSET_OBS_PAGE_SIZE = 10;
  let a22HistoryPage = 1;
  let assetObsPage = 1;

  function switchTab(tabName) {
    const isA22 = tabName === 'a22';
    panelA22?.classList.toggle('active', isA22);
    panelClearance?.classList.toggle('active', !isA22);
    tabButtons.forEach((btn) => {
      const active = btn.dataset.formatTab === tabName;
      btn.classList.toggle('active', active);
      btn.setAttribute('aria-selected', active ? 'true' : 'false');
    });
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

  async function loadResponsiblesAll() {
    const data = await App.get('/responsibles');
    const items = data.responsibles || [];
    if (a22ResponsiblesList) {
      a22ResponsiblesList.innerHTML = items.map((name) => `<option value="${App.escapeHtml(name)}"></option>`).join('');
    }
    if (clearanceResponsiblesList) {
      clearanceResponsiblesList.innerHTML = items.map((name) => `<option value="${App.escapeHtml(name)}"></option>`).join('');
    }
  }

  async function loadA22RunsForPeriod() {
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
    const data = await App.get(`/runs?period_id=${encodeURIComponent(String(pid))}`);
    const runs = data.runs || [];
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

  async function loadClearanceRunsForPeriod() {
    if (!clearanceRunSelect) return;
    const pid = clearancePeriodSelect?.value ? Number(clearancePeriodSelect.value) : null;
    if (!pid) {
      clearanceRunSelect.innerHTML = '<option value="">-- Selecciona jornada cerrada --</option>';
      return;
    }
    const data = await App.get(`/runs?period_id=${encodeURIComponent(String(pid))}`);
    const closed = (data.runs || []).filter((r) => String(r.status || '').toLowerCase() === 'closed');
    clearanceRunSelect.innerHTML = '<option value="">-- Selecciona jornada cerrada --</option>' + closed
      .map((r) => {
        const svc = r.service_scope_label || r.service ? ` [${App.escapeHtml(r.service_scope_label || r.service)}]` : '';
        const metrics = ` E:${Number(r.found || 0)} / NE:${Number(r.not_found || 0)}`;
        return `<option value="${r.id}">${r.id} - ${App.escapeHtml(r.name)}${svc} |${metrics}</option>`;
      }).join('');
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

  function renderAssetObsTable() {
    if (!assetObsContainer) return;
    if (!updateServiceFromRun()) {
      assetObsContainer.innerHTML = '<div class="empty-mini">Selecciona una jornada para cargar codigos.</div>';
      return;
    }
    const q = String(assetObsSearch?.value || '').trim().toUpperCase();
    const rows = serviceAssets.filter((a) => !q || String(a.C_ACT || '').toUpperCase().includes(q) || String(a.NOM || '').toUpperCase().includes(q));
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
        <thead><tr><th>COD ACTIVO</th><th>DESCRIPCION ACTIVO</th><th>OBSERVACION ESPECIFICA</th></tr></thead>
        <tbody>
          ${pageRows.map((a) => {
            const code = String(a.C_ACT || '');
            const current = assetObservationMap[code] ?? '';
            return `<tr>
              <td>${App.escapeHtml(code)}</td>
              <td>${App.escapeHtml(a.NOM)}</td>
              <td><input class="obs-input" data-code="${App.escapeHtml(code)}" value="${App.escapeHtml(current)}" placeholder="Observacion para este activo (opcional)" /></td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
      <div class="history-pagination">
        <div class="history-page-meta">Mostrando ${from}-${to} de ${rows.length}</div>
        <div class="history-page-controls">
          <button type="button" class="mini-btn history-page-btn" data-asset-obs-action="prev" ${currentPage <= 1 ? 'disabled' : ''}>Anterior</button>
          <span class="history-page-indicator">Pagina ${currentPage} de ${totalPages}</span>
          <button type="button" class="mini-btn history-page-btn" data-asset-obs-action="next" ${currentPage >= totalPages ? 'disabled' : ''}>Siguiente</button>
        </div>
      </div>`;
  }

  async function loadA22History() {
    if (!a22HistoryContainer) return;
    const pid = periodSelect?.value ? Number(periodSelect.value) : null;
    if (!pid) {
      a22HistoryContainer.innerHTML = '<div class="empty-mini">Selecciona un periodo para ver historial A22.</div>';
      return;
    }
    const data = await App.get(`/reports/a22_history?period_id=${encodeURIComponent(String(pid))}`);
    const rows = data.items || [];
    if (!rows.length) {
      a22HistoryContainer.innerHTML = '<div class="empty-mini">No hay informes A22 generados todavia.</div>';
      return;
    }
    const totalPages = Math.max(1, Math.ceil(rows.length / HISTORY_PAGE_SIZE));
    a22HistoryPage = Math.min(Math.max(a22HistoryPage, 1), totalPages);
    const startIdx = (a22HistoryPage - 1) * HISTORY_PAGE_SIZE;
    const pageRows = rows.slice(startIdx, startIdx + HISTORY_PAGE_SIZE);
    const from = startIdx + 1;
    const to = Math.min(startIdx + HISTORY_PAGE_SIZE, rows.length);
    a22HistoryContainer.innerHTML = `
      <div class="history-table-wrap">
        <table class="report-history-table">
          <thead><tr><th>ID</th><th>TIPO</th><th>PERIODO</th><th>GENERADO</th><th>ARCHIVO</th><th>ACCION</th></tr></thead>
          <tbody>
            ${pageRows.map((r) => `
              <tr>
                <td>${App.escapeHtml(String(r.id || ''))}</td>
                <td>${App.escapeHtml(r.report_type === 'a22_pdf' ? 'A22 PDF' : 'A22 Excel')}</td>
                <td class="cell-clip" title="${App.escapeHtml(r.period_label || '-')}">${App.escapeHtml(r.period_label || '-')}</td>
                <td>${App.escapeHtml(App.formatDateTime(r.generated_at_local || r.generated_at || ''))}</td>
                <td class="cell-clip" title="${App.escapeHtml(r.file_name || '')}">${App.escapeHtml(r.file_name || '')}</td>
                <td><button type="button" class="mini-btn" data-a22-report-id="${App.escapeHtml(String(r.id || ''))}">Descargar</button></td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>
      <div class="history-pagination">
        <div class="history-page-meta">Mostrando ${from}-${to} de ${rows.length}</div>
        <div class="history-page-controls">
          <button type="button" class="mini-btn history-page-btn" data-a22-history-action="prev" ${a22HistoryPage <= 1 ? 'disabled' : ''}>Anterior</button>
          <span class="history-page-indicator">Pagina ${a22HistoryPage} de ${totalPages}</span>
          <button type="button" class="mini-btn history-page-btn" data-a22-history-action="next" ${a22HistoryPage >= totalPages ? 'disabled' : ''}>Siguiente</button>
        </div>
      </div>`;
  }

  function buildA22Payload() {
    const perAssetObservations = {};
    Object.keys(assetObservationMap).forEach((code) => {
      const v = String(assetObservationMap[code] || '').trim();
      if (v) perAssetObservations[code] = v;
    });
    const run = selectedRun();
    const service = (run?.service || '').trim();
    return {
      service,
      period_id: periodSelect?.value ? Number(periodSelect.value) : null,
      run_id: runSelect?.value ? Number(runSelect.value) : null,
      receiver: (receiverInput?.value || '').trim(),
      observation: (observationInput?.value || '').trim(),
      report_date: (reportDateInput?.value || '').trim(),
      warehouse_lead: (warehouseLeadInput?.value || '').trim(),
      assets_manager: (assetsManagerInput?.value || '').trim(),
      per_asset_observations: perAssetObservations,
    };
  }

  function validateA22() {
    const payload = buildA22Payload();
    if (!payload.period_id) return App.setStatus(a22StatusEl, 'Debes seleccionar el periodo de inventario.', true), false;
    if (!payload.run_id) return App.setStatus(a22StatusEl, 'Debes seleccionar la jornada del periodo.', true), false;
    if (!payload.service) return App.setStatus(a22StatusEl, 'La jornada seleccionada no tiene servicio asociado.', true), false;
    if (!payload.warehouse_lead) return App.setStatus(a22StatusEl, 'Debes escribir el lider de almacen.', true), false;
    if (!payload.assets_manager) return App.setStatus(a22StatusEl, 'Debes escribir el responsable de activos fijos.', true), false;
    return true;
  }

  async function exportA22(url, fallbackName) {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(buildA22Payload()),
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

  async function validateClearance() {
    const pid = clearancePeriodSelect?.value ? Number(clearancePeriodSelect.value) : null;
    const rid = clearanceRunSelect?.value ? Number(clearanceRunSelect.value) : null;
    if (!pid) return App.setStatus(clearanceStatusEl, 'Selecciona un periodo', true);
    if (!rid) return App.setStatus(clearanceStatusEl, 'Selecciona una jornada cerrada', true);
    const data = await App.get(`/paz_y_salvo/validate?period_id=${encodeURIComponent(String(pid))}&run_id=${encodeURIComponent(String(rid))}`);
    App.setStatus(clearanceStatusEl, data.message || '', !data.allowed);
  }

  async function generateClearance() {
    const pid = clearancePeriodSelect?.value ? Number(clearancePeriodSelect.value) : null;
    const rid = clearanceRunSelect?.value ? Number(clearanceRunSelect.value) : null;
    if (!pid) return App.setStatus(clearanceStatusEl, 'Selecciona un periodo', true);
    if (!rid) return App.setStatus(clearanceStatusEl, 'Selecciona una jornada cerrada', true);
    const outgoing = String(clearanceOutgoingInput?.value || '').trim();
    const incoming = String(clearanceIncomingInput?.value || '').trim();
    if (!outgoing) return App.setStatus(clearanceStatusEl, 'Selecciona responsable saliente', true);
    if (!incoming) return App.setStatus(clearanceStatusEl, 'Selecciona responsable entrante', true);
    const res = await fetch('/paz_y_salvo/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        period_id: pid,
        run_id: rid,
        outgoing_responsible: outgoing,
        incoming_responsible: incoming,
        issued_by: String(clearanceIssuedByInput?.value || '').trim(),
        report_date: String(clearanceDateInput?.value || '').trim(),
        observations: String(clearanceNotesInput?.value || '').trim(),
      }),
    });
    if (!res.ok) {
      const payload = await res.json().catch(() => ({}));
      throw new Error(payload.error || `Error HTTP ${res.status}`);
    }
    const blob = await res.blob();
    const header = String(res.headers.get('Content-Disposition') || '');
    const match = header.match(/filename=\"?([^\";]+)\"?/i);
    const filename = match?.[1] || 'paz_y_salvo.pdf';
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    App.setStatus(clearanceStatusEl, 'Paz y salvo generado. El PDF fue descargado y guardado en Documentos.');
  }

  async function reloadA22Filters() {
    await App.loadPeriods(periodSelect);
    await loadA22RunsForPeriod();
    await loadAssetsForService();
    await loadA22History();
  }

  tabButtons.forEach((btn) => btn.addEventListener('click', () => switchTab(btn.dataset.formatTab || 'a22')));

  refreshBtn?.addEventListener('click', () => {
    reloadA22Filters()
      .then(() => App.setStatus(a22StatusEl, 'Filtros A22 actualizados'))
      .catch((err) => App.setStatus(a22StatusEl, err.message, true));
  });

  periodSelect?.addEventListener('change', () => {
    loadA22RunsForPeriod()
      .then(loadAssetsForService)
      .then(loadA22History)
      .catch((err) => App.setStatus(a22StatusEl, err.message, true));
  });

  runSelect?.addEventListener('change', () => {
    updateServiceFromRun();
    loadAssetsForService().catch((err) => App.setStatus(a22StatusEl, err.message, true));
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

  exportExcelBtn?.addEventListener('click', () => {
    if (!validateA22()) return;
    exportA22('/export', 'A22.xlsx')
      .then(() => {
        App.setStatus(a22StatusEl, 'A22 Excel generado');
        return loadA22History();
      })
      .catch((err) => App.setStatus(a22StatusEl, err.message, true));
  });

  exportPdfBtn?.addEventListener('click', () => {
    if (!validateA22()) return;
    exportA22('/export_a22_pdf', 'A22.pdf')
      .then(() => {
        App.setStatus(a22StatusEl, 'A22 PDF generado');
        return loadA22History();
      })
      .catch((err) => App.setStatus(a22StatusEl, err.message, true));
  });

  refreshA22HistoryBtn?.addEventListener('click', () => {
    loadA22History().catch((err) => App.setStatus(a22StatusEl, err.message, true));
  });

  a22HistoryContainer?.addEventListener('click', (e) => {
    const navBtn = e.target.closest('button[data-a22-history-action]');
    if (navBtn) {
      a22HistoryPage += navBtn.getAttribute('data-a22-history-action') === 'prev' ? -1 : 1;
      loadA22History().catch((err) => App.setStatus(a22StatusEl, err.message, true));
      return;
    }
    const btn = e.target.closest('button[data-a22-report-id]');
    if (!btn) return;
    const reportId = btn.getAttribute('data-a22-report-id');
    if (!reportId) return;
    window.open(`/reports/a22_history/${encodeURIComponent(reportId)}/download`, '_blank');
  });

  clearancePeriodSelect?.addEventListener('change', () => {
    loadClearanceRunsForPeriod().catch((err) => App.setStatus(clearanceStatusEl, err.message, true));
  });

  clearanceValidateBtn?.addEventListener('click', () => {
    validateClearance().catch((err) => App.setStatus(clearanceStatusEl, err.message, true));
  });

  clearanceGenerateBtn?.addEventListener('click', () => {
    generateClearance().catch((err) => App.setStatus(clearanceStatusEl, err.message, true));
  });

  if (reportDateInput && !reportDateInput.value) reportDateInput.value = new Date().toISOString().slice(0, 10);
  if (clearanceDateInput && !clearanceDateInput.value) clearanceDateInput.value = new Date().toISOString().slice(0, 10);
  if (clearanceIssuedByInput && !clearanceIssuedByInput.value) clearanceIssuedByInput.value = 'Responsable activos fijos';

  Promise.all([
    App.loadPeriods(periodSelect),
    App.loadPeriods(clearancePeriodSelect),
    loadResponsiblesAll(),
  ])
    .then(() => loadA22RunsForPeriod())
    .then(() => loadAssetsForService())
    .then(() => loadA22History())
    .then(() => loadClearanceRunsForPeriod())
    .then(() => switchTab('a22'))
    .catch((err) => {
      App.setStatus(a22StatusEl, err.message, true);
      App.setStatus(clearanceStatusEl, err.message, true);
    });
});
