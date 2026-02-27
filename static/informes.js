document.addEventListener('DOMContentLoaded', () => {
  const exportAccountingMonthlyBtn = document.getElementById('exportAccountingMonthlyBtn');
  const accountingStatusEl = document.getElementById('accountingStatus');
  const accountingMonthInput = document.getElementById('accountingMonth');
  const accountingYearInput = document.getElementById('accountingYear');
  const accountingReportTitleInput = document.getElementById('accountingReportTitle');
  const accountingGeneratedByInput = document.getElementById('accountingGeneratedBy');
  const refreshAccountingHistoryBtn = document.getElementById('refreshAccountingHistoryBtn');
  const accountingHistoryContainer = document.getElementById('accountingHistoryContainer');

  const HISTORY_PAGE_SIZE = 10;
  let historyPage = 1;

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

  function renderHistoryTable(rows) {
    if (!accountingHistoryContainer) return;
    if (!rows.length) {
      accountingHistoryContainer.innerHTML = '<div class="empty-mini">No hay informes generados todavia.</div>';
      return;
    }
    const totalPages = Math.max(1, Math.ceil(rows.length / HISTORY_PAGE_SIZE));
    historyPage = Math.min(Math.max(historyPage, 1), totalPages);
    const startIdx = (historyPage - 1) * HISTORY_PAGE_SIZE;
    const pageRows = rows.slice(startIdx, startIdx + HISTORY_PAGE_SIZE);
    const from = startIdx + 1;
    const to = Math.min(startIdx + HISTORY_PAGE_SIZE, rows.length);

    accountingHistoryContainer.innerHTML = `
      <div class="history-table-wrap">
      <table class="report-history-table">
        <thead>
          <tr><th>ID</th><th>TITULO</th><th>PERIODO</th><th>GENERADO</th><th>ARCHIVO</th><th>ACCION</th></tr>
        </thead>
        <tbody>
          ${pageRows.map((r) => `
            <tr>
              <td>${App.escapeHtml(String(r.id || ''))}</td>
              <td class="cell-clip" title="${App.escapeHtml(r.title || '')}">${App.escapeHtml(r.title || '')}</td>
              <td class="cell-clip" title="${App.escapeHtml(r.period_label || '-')}">${App.escapeHtml(r.period_label || '-')}</td>
              <td>${App.escapeHtml(App.formatDateTime(r.generated_at_local || r.generated_at || ''))}</td>
              <td class="cell-clip" title="${App.escapeHtml(r.file_name || '')}">${App.escapeHtml(r.file_name || '')}</td>
              <td><button type="button" class="mini-btn" data-report-id="${App.escapeHtml(String(r.id || ''))}">Descargar</button></td>
            </tr>`).join('')}
        </tbody>
      </table>
      </div>
      <div class="history-pagination">
        <div class="history-page-meta">Mostrando ${from}-${to} de ${rows.length}</div>
        <div class="history-page-controls">
          <button type="button" class="mini-btn history-page-btn" data-history-action="prev" ${historyPage <= 1 ? 'disabled' : ''}>Anterior</button>
          <span class="history-page-indicator">Pagina ${historyPage} de ${totalPages}</span>
          <button type="button" class="mini-btn history-page-btn" data-history-action="next" ${historyPage >= totalPages ? 'disabled' : ''}>Siguiente</button>
        </div>
      </div>
    `;
  }

  async function loadAccountingHistory() {
    if (!accountingHistoryContainer) return;
    const data = await App.get('/reports/accounting_monthly_history');
    renderHistoryTable(data.items || []);
  }

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
    } catch (err) {
      App.setStatus(accountingStatusEl, err.message, true);
    }
  });

  refreshAccountingHistoryBtn?.addEventListener('click', () => {
    loadAccountingHistory().catch((err) => App.setStatus(accountingStatusEl, err.message, true));
  });

  accountingHistoryContainer?.addEventListener('click', (e) => {
    const navBtn = e.target.closest('button[data-history-action]');
    if (navBtn) {
      historyPage += navBtn.getAttribute('data-history-action') === 'prev' ? -1 : 1;
      loadAccountingHistory().catch((err) => App.setStatus(accountingStatusEl, err.message, true));
      return;
    }
    const btn = e.target.closest('button[data-report-id]');
    if (!btn) return;
    const reportId = btn.getAttribute('data-report-id');
    if (!reportId) return;
    window.open(`/reports/accounting_monthly_history/${encodeURIComponent(reportId)}/download`, '_blank');
  });

  const today = new Date();
  if (accountingMonthInput && !accountingMonthInput.value) {
    accountingMonthInput.value = String(today.getMonth() + 1);
  }
  if (accountingYearInput && !accountingYearInput.value) {
    accountingYearInput.value = String(today.getFullYear());
  }
  loadAccountingHistory().catch((err) => App.setStatus(accountingStatusEl, err.message, true));
});
