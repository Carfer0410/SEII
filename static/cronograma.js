document.addEventListener('DOMContentLoaded', () => {
  const periodSelect = document.getElementById('cronPeriodSelect');
  const refreshBtn = document.getElementById('cronRefreshBtn');
  const statusEl = document.getElementById('cronStatus');

  const kpisEl = document.getElementById('cronServicesKpis');
  const barEl = document.getElementById('cronServicesBar');
  const metaEl = document.getElementById('cronServicesMeta');
  const tableEl = document.getElementById('cronServicesTable');
  const pendingEl = document.getElementById('cronPendingServices');
  const runsEl = document.getElementById('cronRunsTable');
  const recommendationsEl = document.getElementById('cronRecommendations');
  const operationalPlanEl = document.getElementById('cronOperationalPlan');

  function selectedPeriodId() {
    return periodSelect?.value ? Number(periodSelect.value) : null;
  }

  function renderCoverage(data) {
    const summary = data?.summary || {};
    const rows = data?.services || [];
    const total = Number(summary.total_services || 0);
    const done = Number(summary.done_services || 0);
    const pending = Number(summary.pending_services || 0);
    const donePct = Number(summary.done_pct || 0);
    const pendingPct = Number(summary.pending_pct || 0);

    kpisEl.innerHTML = `
      <div class="jour-kpi-box">
        <span>Total servicios</span>
        <strong>${total}</strong>
      </div>
      <div class="jour-kpi-box jour-kpi-ok">
        <span>Inventariados</span>
        <strong>${done} (${donePct.toFixed(2)}%)</strong>
      </div>
      <div class="jour-kpi-box jour-kpi-pending">
        <span>Pendientes</span>
        <strong>${pending} (${pendingPct.toFixed(2)}%)</strong>
      </div>
    `;
    barEl.style.width = `${Math.max(0, Math.min(100, donePct))}%`;
    metaEl.textContent = `Cobertura actual del periodo: ${donePct.toFixed(2)}%.`;

    tableEl.innerHTML = rows.length ? `
      <div class="closed-runs-wrap">
        <table class="closed-runs-table">
          <thead><tr><th>SERVICIO</th><th>ACTIVOS</th><th>ESTADO</th><th>% AVANCE</th></tr></thead>
          <tbody>
            ${rows.map((r) => `
              <tr>
                <td>${App.escapeHtml(r.service || '')}</td>
                <td>${App.escapeHtml(String(r.asset_count || 0))}</td>
                <td>${App.escapeHtml(r.status || '')}</td>
                <td>${App.escapeHtml(String(r.status_pct || 0))}%</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    ` : '<div class="empty-mini">No hay servicios para mostrar.</div>';

    const pendingRows = rows
      .filter((r) => r.status !== 'Inventariado')
      .sort((a, b) => Number(b.asset_count || 0) - Number(a.asset_count || 0));
    pendingEl.innerHTML = pendingRows.length ? `
      <div class="closed-runs-wrap">
        <table class="closed-runs-table">
          <thead><tr><th>SERVICIO PENDIENTE</th><th>ACTIVOS</th></tr></thead>
          <tbody>${pendingRows.map((r) => `<tr><td>${App.escapeHtml(r.service || '')}</td><td>${App.escapeHtml(String(r.asset_count || 0))}</td></tr>`).join('')}</tbody>
        </table>
      </div>
    ` : '<div class="empty-mini">No hay servicios pendientes. Cobertura completa.</div>';

    const recs = data?.recommendations || [];
    recommendationsEl.innerHTML = recs.length ? `
      <ol class="cron-list">
        ${recs.map((r) => `<li><b>${App.escapeHtml(r.service || '')}</b> - ${App.escapeHtml(r.reason || '')}</li>`).join('')}
      </ol>
    ` : '<div class="empty-mini">No hay recomendaciones pendientes para este periodo.</div>';

    const highLoad = pendingRows.filter((r) => Number(r.asset_count || 0) >= 120);
    const midLoad = pendingRows.filter((r) => Number(r.asset_count || 0) >= 40 && Number(r.asset_count || 0) < 120);
    const lowLoad = pendingRows.filter((r) => Number(r.asset_count || 0) < 40);
    const totalPendingAssets = pendingRows.reduce((acc, r) => acc + Number(r.asset_count || 0), 0);
    const estimatedRuns = pendingRows.length;

    operationalPlanEl.innerHTML = `
      <ul class="cron-list">
        <li><b>Jornadas pendientes estimadas:</b> ${estimatedRuns}</li>
        <li><b>Activos pendientes por inventariar:</b> ${totalPendingAssets}</li>
        <li><b>Servicios de carga alta:</b> ${highLoad.length} ${highLoad.length ? `(>=120 activos)` : ''}</li>
        <li><b>Servicios de carga media:</b> ${midLoad.length} ${midLoad.length ? `(40 a 119 activos)` : ''}</li>
        <li><b>Servicios de cierre rapido:</b> ${lowLoad.length} ${lowLoad.length ? `(<40 activos)` : ''}</li>
        <li><b>Sugerencia:</b> programa primero 1 servicio de carga alta por jornada, y cierra el dia con 1-2 servicios de carga baja.</li>
      </ul>
    `;
  }

  function renderRuns(data) {
    const runs = data?.runs || [];
    runsEl.innerHTML = runs.length ? `
      <div class="closed-runs-wrap">
        <table class="closed-runs-table">
          <thead><tr><th>ID</th><th>JORNADA</th><th>SERVICIO</th><th>ESTADO</th></tr></thead>
          <tbody>
            ${runs.map((r) => `
              <tr>
                <td>${App.escapeHtml(r.id || '')}</td>
                <td>${App.escapeHtml(r.name || '')}</td>
                <td>${App.escapeHtml(r.service || 'TODOS')}</td>
                <td>${App.escapeHtml(r.status === 'active' ? 'Activa' : (r.status === 'cancelled' ? 'Anulada' : 'Cerrada'))}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    ` : '<div class="empty-mini">No hay jornadas registradas en este periodo.</div>';
  }

  async function loadData() {
    const pid = selectedPeriodId();
    if (!pid) {
      kpisEl.innerHTML = '<div class="empty-mini">Selecciona un periodo para ver el cronograma.</div>';
      barEl.style.width = '0%';
      metaEl.textContent = '';
      tableEl.innerHTML = '<div class="empty-mini">Sin datos para mostrar.</div>';
      pendingEl.innerHTML = '<div class="empty-mini">Sin datos para mostrar.</div>';
      runsEl.innerHTML = '<div class="empty-mini">Sin datos para mostrar.</div>';
      recommendationsEl.innerHTML = '<div class="empty-mini">Sin datos para mostrar.</div>';
      operationalPlanEl.innerHTML = '<div class="empty-mini">Sin datos para mostrar.</div>';
      App.setStatus(statusEl, 'Selecciona un periodo para consultar el cronograma.', true);
      return;
    }
    const [coverage, runs] = await Promise.all([
      App.get(`/periods/${pid}/service_coverage`),
      App.get(`/runs?period_id=${encodeURIComponent(pid)}`),
    ]);
    renderCoverage(coverage);
    renderRuns(runs);
    App.setStatus(statusEl, 'Cronograma actualizado.');
  }

  async function loadPeriods() {
    const data = await App.get('/periods');
    const periods = data.periods || [];
    periodSelect.innerHTML = '<option value="">-- Selecciona periodo --</option>' +
      periods.map((p) => `<option value="${p.id}">${App.escapeHtml(p.name)}${p.status === 'open' ? ' (Abierto)' : (p.status === 'cancelled' ? ' (Anulado)' : ' (Cerrado)')}</option>`).join('');
  }

  refreshBtn?.addEventListener('click', () => loadData().catch((e) => App.setStatus(statusEl, e.message, true)));
  periodSelect?.addEventListener('change', () => loadData().catch((e) => App.setStatus(statusEl, e.message, true)));

  loadPeriods()
    .then(loadData)
    .catch((e) => App.setStatus(statusEl, e.message, true));
});
