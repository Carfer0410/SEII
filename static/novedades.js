document.addEventListener('DOMContentLoaded', () => {
  const periodSelect = document.getElementById('issuesPeriodSelect');
  const scanBtn = document.getElementById('issuesScanBtn');
  const refreshBtn = document.getElementById('issuesRefreshBtn');
  const pdfBtn = document.getElementById('issuesPdfBtn');
  const statusFilter = document.getElementById('issuesStatusFilter');
  const severityFilter = document.getElementById('issuesSeverityFilter');
  const analyzeBaseCheck = document.getElementById('issuesAnalyzeBase');
  const helpBtn = document.getElementById('issuesHelpBtn');
  const helpModal = document.getElementById('issuesHelpModal');
  const helpCloseBtn = document.getElementById('issuesHelpCloseBtn');
  const statusEl = document.getElementById('issuesStatus');
  const kpisEl = document.getElementById('issuesKpis');
  const tableEl = document.getElementById('issuesTable');

  let currentItems = [];
  const ISSUES_PAGE_SIZE = 10;
  let issuesPage = 1;

  function money(v) {
    const n = Number(v || 0);
    return `$${n.toLocaleString('es-CO', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
  }

  function selectedPeriodId() {
    return periodSelect?.value ? Number(periodSelect.value) : null;
  }

  function renderKpis(summary) {
    const s = summary || {};
    kpisEl.innerHTML = `
      <div class="jour-kpi-box">
        <span>Total novedades</span>
        <strong>${App.escapeHtml(String(s.total || 0))}</strong>
      </div>
      <div class="jour-kpi-box jour-kpi-pending">
        <span>Abiertas</span>
        <strong>${App.escapeHtml(String(s.open || 0))}</strong>
      </div>
      <div class="jour-kpi-box jour-kpi-ok">
        <span>Valor en riesgo</span>
        <strong>${App.escapeHtml(money(s.value_risk || 0))}</strong>
      </div>
    `;
  }

  function renderTable(items) {
    if (!items.length) {
      tableEl.innerHTML = '<div class="empty-mini">No hay novedades para este filtro.</div>';
      return;
    }
    const totalPages = Math.max(1, Math.ceil(items.length / ISSUES_PAGE_SIZE));
    const currentPage = Math.min(Math.max(issuesPage, 1), totalPages);
    issuesPage = currentPage;
    const startIdx = (currentPage - 1) * ISSUES_PAGE_SIZE;
    const pageItems = items.slice(startIdx, startIdx + ISSUES_PAGE_SIZE);
    const from = startIdx + 1;
    const to = Math.min(startIdx + ISSUES_PAGE_SIZE, items.length);

    tableEl.innerHTML = `
      <div class="history-table-wrap">
        <table class="issues-table">
          <thead>
            <tr>
              <th>ID</th><th>TIPO</th><th>ACTIVO</th><th>SERVICIO</th><th>SEV</th><th>ESTADO</th><th>IMPACTO</th><th>EVIDENCIA</th><th>ASIGNADO</th><th>VENCE</th><th>NOTA</th><th>ACCION</th>
            </tr>
          </thead>
          <tbody>
            ${pageItems.map((r) => `
              <tr>
                <td>${App.escapeHtml(String(r.id || ''))}</td>
                <td class="cell-clip" title="${App.escapeHtml(r.description || '')}">${App.escapeHtml(r.issue_type_label || '')}</td>
                <td class="cell-clip" title="${App.escapeHtml(`${r.asset_code || ''} ${r.asset_name || ''}`.trim())}">${App.escapeHtml(`${r.asset_code || ''} ${r.asset_name || ''}`.trim())}</td>
                <td class="cell-clip" title="${App.escapeHtml(r.service || '')}">${App.escapeHtml(r.service || '')}</td>
                <td>
                  <select class="issues-input" data-field="severity" data-id="${r.id}">
                    ${['Alta','Media','Baja'].map(v => `<option value="${v}" ${v===r.severity?'selected':''}>${v}</option>`).join('')}
                  </select>
                </td>
                <td>
                  <select class="issues-input" data-field="status" data-id="${r.id}">
                    ${['Nuevo','En analisis','Escalado','Cerrado'].map(v => `<option value="${v}" ${v===r.status?'selected':''}>${v}</option>`).join('')}
                  </select>
                </td>
                <td class="cell-money">${App.escapeHtml(money(r.detected_value || 0))}</td>
                <td class="cell-evidence" title="${App.escapeHtml(r.description || '')}">${App.escapeHtml(r.description || '')}</td>
                <td><input class="issues-input" data-field="assigned_to" data-id="${r.id}" value="${App.escapeHtml(r.assigned_to || '')}" placeholder="Asignado a"/></td>
                <td><input class="issues-input" type="date" data-field="due_date" data-id="${r.id}" value="${App.escapeHtml(r.due_date || '')}"/></td>
                <td><input class="issues-input" data-field="resolution_notes" data-id="${r.id}" value="${App.escapeHtml(r.resolution_notes || '')}" placeholder="Nota de gestion"/></td>
                <td><button class="mini-btn" data-save-id="${r.id}" type="button">Guardar</button></td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
      <div class="history-pagination">
        <div class="history-page-meta">Mostrando ${from}-${to} de ${items.length}</div>
        <div class="history-page-controls">
          <button
            type="button"
            class="mini-btn history-page-btn"
            data-issues-action="prev"
            ${currentPage <= 1 ? 'disabled' : ''}
          >Anterior</button>
          <span class="history-page-indicator">Pagina ${currentPage} de ${totalPages}</span>
          <button
            type="button"
            class="mini-btn history-page-btn"
            data-issues-action="next"
            ${currentPage >= totalPages ? 'disabled' : ''}
          >Siguiente</button>
        </div>
      </div>
    `;
  }

  async function loadPeriods() {
    const data = await App.get('/periods');
    const periods = data.periods || [];
    periodSelect.innerHTML = '<option value="">-- Selecciona periodo --</option>' +
      periods.map((p) => `<option value="${p.id}">${App.escapeHtml(p.name)}${p.status === 'open' ? ' (Abierto)' : (p.status === 'cancelled' ? ' (Anulado)' : ' (Cerrado)')}</option>`).join('');
  }

  async function loadIssues() {
    const pid = selectedPeriodId();
    if (!pid) {
      renderKpis({ total: 0, open: 0, value_risk: 0 });
      renderTable([]);
      return;
    }
    const p = new URLSearchParams({ period_id: String(pid) });
    if (statusFilter?.value) p.set('status', statusFilter.value);
    if (severityFilter?.value) p.set('severity', severityFilter.value);
    const data = await App.get(`/issues?${p.toString()}`);
    currentItems = data.items || [];
    renderKpis(data.summary || {});
    renderTable(currentItems);
  }

  scanBtn?.addEventListener('click', async () => {
    const pid = selectedPeriodId();
    if (!pid) return App.setStatus(statusEl, 'Selecciona un periodo.', true);
    try {
      App.setStatus(statusEl, 'Analizando novedades...');
      const data = await App.post('/issues/scan', { period_id: pid, analyze_base: !!analyzeBaseCheck?.checked });
      issuesPage = 1;
      await loadIssues();
      App.setStatus(statusEl, `Analisis completado. Novedades detectadas: ${data.created || 0}`);
    } catch (e) {
      App.setStatus(statusEl, e.message, true);
    }
  });

  refreshBtn?.addEventListener('click', () => {
    issuesPage = 1;
    loadIssues().then(() => App.setStatus(statusEl, 'Tablero actualizado')).catch((e) => App.setStatus(statusEl, e.message, true));
  });
  periodSelect?.addEventListener('change', () => {
    issuesPage = 1;
    loadIssues().catch((e) => App.setStatus(statusEl, e.message, true));
  });
  statusFilter?.addEventListener('change', () => {
    issuesPage = 1;
    loadIssues().catch((e) => App.setStatus(statusEl, e.message, true));
  });
  severityFilter?.addEventListener('change', () => {
    issuesPage = 1;
    loadIssues().catch((e) => App.setStatus(statusEl, e.message, true));
  });

  pdfBtn?.addEventListener('click', () => {
    const pid = selectedPeriodId();
    if (!pid) return App.setStatus(statusEl, 'Selecciona un periodo.', true);
    window.location = `/issues/report_pdf?period_id=${encodeURIComponent(pid)}`;
  });

  helpBtn?.addEventListener('click', () => {
    if (helpModal) helpModal.hidden = false;
  });

  helpCloseBtn?.addEventListener('click', () => {
    if (helpModal) helpModal.hidden = true;
  });

  helpModal?.addEventListener('click', (e) => {
    if (e.target === helpModal) helpModal.hidden = true;
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && helpModal && !helpModal.hidden) {
      helpModal.hidden = true;
    }
  });

  tableEl?.addEventListener('click', async (e) => {
    const navBtn = e.target.closest('button[data-issues-action]');
    if (navBtn) {
      const action = navBtn.getAttribute('data-issues-action');
      issuesPage = action === 'prev' ? issuesPage - 1 : issuesPage + 1;
      renderTable(currentItems);
      return;
    }
    const btn = e.target.closest('button[data-save-id]');
    if (!btn) return;
    const id = Number(btn.getAttribute('data-save-id'));
    if (!id) return;
    const rowInputs = tableEl.querySelectorAll(`[data-id="${id}"]`);
    const payload = {};
    rowInputs.forEach((input) => {
      const field = input.getAttribute('data-field');
      if (field) payload[field] = input.value || '';
    });
    try {
      await App.patch(`/issues/${id}`, payload);
      App.setStatus(statusEl, `Novedad ${id} actualizada.`);
      await loadIssues();
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  loadPeriods()
    .then(() => {
      if (analyzeBaseCheck) analyzeBaseCheck.checked = false;
      return loadIssues();
    })
    .catch((e) => App.setStatus(statusEl, e.message, true));
});
