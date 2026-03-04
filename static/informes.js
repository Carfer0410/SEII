document.addEventListener('DOMContentLoaded', () => {
  const exportAccountingMonthlyBtn = document.getElementById('exportAccountingMonthlyBtn');
  const accountingStatusEl = document.getElementById('accountingStatus');
  const accountingMonthInput = document.getElementById('accountingMonth');
  const accountingYearInput = document.getElementById('accountingYear');
  const accountingReportTitleInput = document.getElementById('accountingReportTitle');
  const accountingGeneratedByInput = document.getElementById('accountingGeneratedBy');
  const refreshAccountingHistoryBtn = document.getElementById('refreshAccountingHistoryBtn');
  const accountingHistoryContainer = document.getElementById('accountingHistoryContainer');

  const accountingBaseMonthInput = document.getElementById('accountingBaseMonth');
  const accountingBaseYearInput = document.getElementById('accountingBaseYear');
  const accountingBaseUploadedByInput = document.getElementById('accountingBaseUploadedBy');
  const accountingBaseFileInput = document.getElementById('accountingBaseFile');
  const uploadAccountingBaseBtn = document.getElementById('uploadAccountingBaseBtn');
  const refreshAccountingBasesBtn = document.getElementById('refreshAccountingBasesBtn');
  const refreshOverridesSummaryBtn = document.getElementById('refreshOverridesSummaryBtn');
  const accountingBaseStatusEl = document.getElementById('accountingBaseStatus');
  const accountingBasesContainer = document.getElementById('accountingBasesContainer');
  const accountingOverridesSummaryContainer = document.getElementById('accountingOverridesSummaryContainer');

  const HISTORY_PAGE_SIZE = 10;
  const BASES_PAGE_SIZE = 8;
  let historyPage = 1;
  let basesPage = 1;
  let lastHistoryRows = [];
  let lastBaseRows = [];

  function accountingPeriodParams() {
    const month = Number(accountingMonthInput?.value || 0);
    const year = Number(accountingYearInput?.value || 0);
    return {
      month: (month >= 1 && month <= 12) ? month : null,
      year: (year >= 2000 && year <= 2100) ? year : null,
    };
  }

  function accountingBasePeriodParams() {
    const month = Number(accountingBaseMonthInput?.value || 0);
    const year = Number(accountingBaseYearInput?.value || 0);
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

  function selectedUploadedBy() {
    return String(accountingBaseUploadedByInput?.value || '').trim();
  }

  function syncBasePeriodWithReport() {
    if (accountingBaseMonthInput && accountingMonthInput) accountingBaseMonthInput.value = accountingMonthInput.value;
    if (accountingBaseYearInput && accountingYearInput) accountingBaseYearInput.value = accountingYearInput.value;
  }

  function pagedRows(rows, page, size) {
    const totalPages = Math.max(1, Math.ceil(rows.length / size));
    const safePage = Math.min(Math.max(page, 1), totalPages);
    const startIdx = (safePage - 1) * size;
    return {
      pageRows: rows.slice(startIdx, startIdx + size),
      from: rows.length ? (startIdx + 1) : 0,
      to: Math.min(startIdx + size, rows.length),
      totalPages,
      safePage,
    };
  }

  function renderHistoryTable(rows) {
    if (!accountingHistoryContainer) return;
    if (!rows.length) {
      accountingHistoryContainer.innerHTML = '<div class="empty-mini">No hay informes generados todavia.</div>';
      return;
    }
    const pageData = pagedRows(rows, historyPage, HISTORY_PAGE_SIZE);
    historyPage = pageData.safePage;

    accountingHistoryContainer.innerHTML = `
      <div class="history-table-wrap">
      <table class="report-history-table compact-table">
        <thead>
          <tr><th>ID</th><th>TITULO</th><th>PERIODO</th><th>BASE MES</th><th>GENERADO</th><th>ARCHIVO</th><th>ACCION</th></tr>
        </thead>
        <tbody>
          ${pageData.pageRows.map((r) => `
            <tr>
              <td>${App.escapeHtml(String(r.id || ''))}</td>
              <td class="cell-clip" title="${App.escapeHtml(r.title || '')}">${App.escapeHtml(r.title || '')}</td>
              <td>${App.escapeHtml(r.period_label || '-')}</td>
              <td>${App.escapeHtml((r.accounting_base?.source_file_name) || `#${r.accounting_base_id || '-'}`)}</td>
              <td>${App.escapeHtml(App.formatDateTime(r.generated_at_local || r.generated_at || ''))}</td>
              <td class="cell-clip" title="${App.escapeHtml(r.file_name || '')}">${App.escapeHtml(r.file_name || '')}</td>
              <td><button type="button" class="mini-btn" data-report-id="${App.escapeHtml(String(r.id || ''))}">Descargar</button></td>
            </tr>`).join('')}
        </tbody>
      </table>
      </div>
      <div class="history-pagination">
        <div class="history-page-meta">Mostrando ${pageData.from}-${pageData.to} de ${rows.length}</div>
        <div class="history-page-controls">
          <button type="button" class="mini-btn history-page-btn" data-history-action="prev" ${historyPage <= 1 ? 'disabled' : ''}>Anterior</button>
          <span class="history-page-indicator">Pagina ${historyPage} de ${pageData.totalPages}</span>
          <button type="button" class="mini-btn history-page-btn" data-history-action="next" ${historyPage >= pageData.totalPages ? 'disabled' : ''}>Siguiente</button>
        </div>
      </div>
    `;
  }

  function renderBasesTable(rows) {
    if (!accountingBasesContainer) return;
    if (!rows.length) {
      accountingBasesContainer.innerHTML = '<div class="empty-mini">No hay bases mensuales cargadas para el periodo seleccionado.</div>';
      return;
    }
    const pageData = pagedRows(rows, basesPage, BASES_PAGE_SIZE);
    basesPage = pageData.safePage;

    accountingBasesContainer.innerHTML = `
      <div class="history-table-wrap">
      <table class="report-history-table compact-table">
        <thead>
          <tr><th>ID</th><th>PERIODO</th><th>ARCHIVO BASE</th><th>ACTIVOS</th><th>USUARIO</th><th>CARGADA</th><th>ACCION</th></tr>
        </thead>
        <tbody>
          ${pageData.pageRows.map((r) => `
            <tr>
              <td>${App.escapeHtml(String(r.id || ''))}</td>
              <td>${App.escapeHtml(r.period_label || '-')}</td>
              <td class="cell-clip" title="${App.escapeHtml(r.source_file_name || '')}">${App.escapeHtml(r.source_file_name || '')}</td>
              <td>${App.escapeHtml(String(r.asset_count || 0))}</td>
              <td>${App.escapeHtml(r.uploaded_by || '-')}</td>
              <td>${App.escapeHtml(App.formatDateTime(r.uploaded_at_local || r.uploaded_at || ''))}</td>
              <td><button type="button" class="mini-btn" data-base-id="${App.escapeHtml(String(r.id || ''))}">Descargar</button></td>
            </tr>`).join('')}
        </tbody>
      </table>
      </div>
      <div class="history-pagination">
        <div class="history-page-meta">Mostrando ${pageData.from}-${pageData.to} de ${rows.length}</div>
        <div class="history-page-controls">
          <button type="button" class="mini-btn history-base-btn" data-base-action="prev" ${basesPage <= 1 ? 'disabled' : ''}>Anterior</button>
          <span class="history-page-indicator">Pagina ${basesPage} de ${pageData.totalPages}</span>
          <button type="button" class="mini-btn history-base-btn" data-base-action="next" ${basesPage >= pageData.totalPages ? 'disabled' : ''}>Siguiente</button>
        </div>
      </div>
    `;
  }

  function money(v) {
    if (v === null || v === undefined || Number.isNaN(Number(v))) return '-';
    return Number(v).toLocaleString('es-CO', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function renderOverridesSummary(payload) {
    if (!accountingOverridesSummaryContainer) return;
    const summary = payload?.summary || {};
    const totals = summary.totals || {};
    const rows = summary.rows || [];
    const periodLabel = payload?.period?.label || '-';
    const baseName = payload?.base?.source_file_name || '-';
    if (!rows.length) {
      accountingOverridesSummaryContainer.innerHTML = '<div class="empty-mini">No hay ajustes contables configurados para auditar.</div>';
      return;
    }
    accountingOverridesSummaryContainer.innerHTML = `
      <div class="compact-kpis">
        <div class="compact-kpi"><strong>Periodo:</strong> ${App.escapeHtml(periodLabel)}</div>
        <div class="compact-kpi"><strong>Base:</strong> ${App.escapeHtml(baseName)}</div>
        <div class="compact-kpi"><strong>Ajustes contables:</strong> ${App.escapeHtml(String(totals.configured || 0))}</div>
        <div class="compact-kpi"><strong>Presentes:</strong> ${App.escapeHtml(String(totals.present || 0))}</div>
        <div class="compact-kpi"><strong>Faltantes:</strong> ${App.escapeHtml(String(totals.missing || 0))}</div>
        <div class="compact-kpi"><strong>Duplicados:</strong> ${App.escapeHtml(String(totals.duplicates || 0))}</div>
        <div class="compact-kpi"><strong>Delta total:</strong> ${App.escapeHtml(money(totals.delta_total || 0))}</div>
      </div>
      <div class="history-table-wrap">
      <table class="report-history-table compact-table">
        <thead>
          <tr><th>C_ACT</th><th>ESTADO</th><th>HITS</th><th>FAMILIA</th><th>COSTO BASE</th><th>COSTO AJUSTADO</th><th>DELTA</th></tr>
        </thead>
        <tbody>
          ${rows.map((r) => `
            <tr>
              <td>${App.escapeHtml(r.c_act || '')}</td>
              <td>${r.present ? '<span class="tag-ok">Aplicado</span>' : '<span class="tag-miss">No encontrado</span>'}</td>
              <td>${App.escapeHtml(String(r.hits || 0))}</td>
              <td class="cell-clip" title="${App.escapeHtml(`${r.c_fam || ''} ${r.nom_fam || ''}`.trim())}">${App.escapeHtml(`${r.c_fam || ''} ${r.nom_fam || ''}`.trim() || '-')}</td>
              <td>${App.escapeHtml(money(r.base_cost))}</td>
              <td>${App.escapeHtml(money(r.override_cost))}</td>
              <td>${App.escapeHtml(money(r.delta))}</td>
            </tr>`).join('')}
        </tbody>
      </table>
      </div>
    `;
  }

  async function loadAccountingHistory() {
    if (!accountingHistoryContainer) return;
    const data = await App.get('/reports/accounting_monthly_history');
    lastHistoryRows = data.items || [];
    renderHistoryTable(lastHistoryRows);
  }

  async function loadAccountingBases() {
    if (!accountingBasesContainer) return;
    const qp = accountingBasePeriodParams();
    const params = new URLSearchParams();
    if (qp.month) params.set('month', String(qp.month));
    if (qp.year) params.set('year', String(qp.year));
    const data = await App.get('/reports/accounting_monthly_bases' + (params.toString() ? `?${params.toString()}` : ''));
    lastBaseRows = data.items || [];
    renderBasesTable(lastBaseRows);
  }

  async function loadOverridesSummary() {
    if (!accountingOverridesSummaryContainer) return;
    const qp = accountingPeriodParams();
    if (!qp.month || !qp.year) {
      accountingOverridesSummaryContainer.innerHTML = '<div class="empty-mini">Selecciona mes y año para consultar ajustes contables.</div>';
      return;
    }
    const params = new URLSearchParams({ month: String(qp.month), year: String(qp.year) });
    try {
      const data = await App.get('/reports/accounting_monthly_overrides_summary?' + params.toString());
      renderOverridesSummary(data);
    } catch (err) {
      accountingOverridesSummaryContainer.innerHTML = `<div class="empty-mini">${App.escapeHtml(err.message || 'No fue posible consultar ajustes contables.')}</div>`;
    }
  }

  uploadAccountingBaseBtn?.addEventListener('click', async () => {
    try {
      const file = accountingBaseFileInput?.files?.[0];
      if (!file) {
        App.setStatus(accountingBaseStatusEl, 'Debes seleccionar el archivo base mensual.', true);
        return;
      }
      const qp = accountingBasePeriodParams();
      if (!qp.month || !qp.year) {
        App.setStatus(accountingBaseStatusEl, 'Debes seleccionar mes y año validos.', true);
        return;
      }
      const fd = new FormData();
      fd.append('file', file);
      fd.append('month', String(qp.month));
      fd.append('year', String(qp.year));
      fd.append('uploaded_by', selectedUploadedBy());
      App.setStatus(accountingBaseStatusEl, 'Cargando base mensual...');
      const res = await fetch('/reports/accounting_monthly_bases/upload', { method: 'POST', body: fd });
      const payload = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(payload.error || 'No fue posible cargar la base mensual');
      App.setStatus(accountingBaseStatusEl, `Base mensual cargada. Activos: ${payload.base?.asset_count || 0}`);
      accountingBaseFileInput.value = '';
      await loadAccountingBases();
      await loadOverridesSummary();
    } catch (err) {
      App.setStatus(accountingBaseStatusEl, err.message, true);
    }
  });

  exportAccountingMonthlyBtn?.addEventListener('click', async () => {
    try {
      const reportTitle = accountingReportTitle();
      if (!reportTitle) {
        App.setStatus(accountingStatusEl, 'Debes escribir el titulo del informe contable.', true);
        accountingReportTitleInput?.focus();
        return;
      }
      const generatedBy = accountingGeneratedBy();
      if (!generatedBy) {
        App.setStatus(accountingStatusEl, 'Debes escribir el usuario que genera el informe.', true);
        accountingGeneratedByInput?.focus();
        return;
      }
      App.setStatus(accountingStatusEl, 'Generando informe contable mensual...');
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
      App.setStatus(accountingStatusEl, 'Informe contable mensual generado correctamente.');
      await loadAccountingHistory();
      await loadOverridesSummary();
    } catch (err) {
      App.setStatus(accountingStatusEl, err.message, true);
    }
  });

  refreshAccountingHistoryBtn?.addEventListener('click', () => {
    loadAccountingHistory().catch((err) => App.setStatus(accountingStatusEl, err.message, true));
  });
  refreshAccountingBasesBtn?.addEventListener('click', () => {
    loadAccountingBases().catch((err) => App.setStatus(accountingBaseStatusEl, err.message, true));
  });
  refreshOverridesSummaryBtn?.addEventListener('click', () => {
    loadOverridesSummary().catch((err) => App.setStatus(accountingBaseStatusEl, err.message, true));
  });

  accountingHistoryContainer?.addEventListener('click', (e) => {
    const navBtn = e.target.closest('button[data-history-action]');
    if (navBtn) {
      historyPage += navBtn.getAttribute('data-history-action') === 'prev' ? -1 : 1;
      renderHistoryTable(lastHistoryRows);
      return;
    }
    const btn = e.target.closest('button[data-report-id]');
    if (!btn) return;
    const reportId = btn.getAttribute('data-report-id');
    if (!reportId) return;
    window.open(`/reports/accounting_monthly_history/${encodeURIComponent(reportId)}/download`, '_blank');
  });

  accountingBasesContainer?.addEventListener('click', (e) => {
    const navBtn = e.target.closest('button[data-base-action]');
    if (navBtn) {
      basesPage += navBtn.getAttribute('data-base-action') === 'prev' ? -1 : 1;
      renderBasesTable(lastBaseRows);
      return;
    }
    const btn = e.target.closest('button[data-base-id]');
    if (!btn) return;
    const baseId = btn.getAttribute('data-base-id');
    if (!baseId) return;
    window.open(`/reports/accounting_monthly_bases/${encodeURIComponent(baseId)}/download`, '_blank');
  });

  const today = new Date();
  if (accountingMonthInput && !accountingMonthInput.value) accountingMonthInput.value = String(today.getMonth() + 1);
  if (accountingYearInput && !accountingYearInput.value) accountingYearInput.value = String(today.getFullYear());
  syncBasePeriodWithReport();

  accountingMonthInput?.addEventListener('change', () => {
    syncBasePeriodWithReport();
    loadOverridesSummary().catch((err) => App.setStatus(accountingBaseStatusEl, err.message, true));
  });
  accountingYearInput?.addEventListener('change', () => {
    syncBasePeriodWithReport();
    loadOverridesSummary().catch((err) => App.setStatus(accountingBaseStatusEl, err.message, true));
  });
  accountingBaseMonthInput?.addEventListener('change', () => {
    loadAccountingBases().catch((err) => App.setStatus(accountingBaseStatusEl, err.message, true));
  });
  accountingBaseYearInput?.addEventListener('change', () => {
    loadAccountingBases().catch((err) => App.setStatus(accountingBaseStatusEl, err.message, true));
  });

  loadAccountingHistory().catch((err) => App.setStatus(accountingStatusEl, err.message, true));
  loadAccountingBases().catch((err) => App.setStatus(accountingBaseStatusEl, err.message, true));
  loadOverridesSummary().catch((err) => App.setStatus(accountingBaseStatusEl, err.message, true));
});
