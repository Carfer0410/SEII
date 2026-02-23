document.addEventListener('DOMContentLoaded', () => {
  const serviceAuto = document.getElementById('serviceAuto');
  const periodSelect = document.getElementById('periodSelect');
  const runSelect = document.getElementById('runSelect');
  const refreshBtn = document.getElementById('refreshDashboardBtn');
  const exportPdfBtn = document.getElementById('exportDashboardPdfBtn');
  const statusEl = document.getElementById('dashboardStatus');
  const kpisEl = document.getElementById('dashboardKpis');
  const serviceChartEl = document.getElementById('serviceChart');
  const typeChartEl = document.getElementById('typeChart');
  const areaChartEl = document.getElementById('areaChart');
  const statusPieChartEl = document.getElementById('statusPieChart');
  const topServiceEl = document.getElementById('topService');
  const topTypeEl = document.getElementById('topType');
  const topAreaEl = document.getElementById('topArea');
  const comparePeriodA = document.getElementById('comparePeriodA');
  const comparePeriodB = document.getElementById('comparePeriodB');
  const comparePeriodsBtn = document.getElementById('comparePeriodsBtn');
  const compareStatus = document.getElementById('compareStatus');
  const compareContainer = document.getElementById('compareContainer');
  const notFoundSpecificStatus = document.getElementById('notFoundSpecificStatus');
  const notFoundSpecificContainer = document.getElementById('notFoundSpecificContainer');

  let serviceChart = null;
  let typeChart = null;
  let areaChart = null;
  let statusPieChart = null;
  let lastPayload = null;
  let runsCache = [];

  const DASHBOARD_CHART_THEME = {
    axisText: '#2f556f',
    grid: 'rgba(31, 95, 143, 0.14)',
    tooltipBg: 'rgba(15, 44, 66, 0.94)',
    tooltipTitle: '#e8f7ff',
    tooltipBody: '#cce9f7',
    stroke: 'rgba(255,255,255,0.95)',
    barPalette: [
      ['#24a9de', '#1689c2'],
      ['#1f9b4a', '#2aa65a'],
      ['#1f5f8f', '#2f76a8'],
      ['#f2c335', '#dfab1c'],
      ['#2e8ec2', '#1f5f8f'],
      ['#56bb63', '#2f9a4d'],
    ],
    piePalette: ['#1f9b4a', '#f26a5b', '#f2c335'],
  };

  function selectedRunId() {
    return runSelect.value ? Number(runSelect.value) : null;
  }

  function selectedRun() {
    const id = selectedRunId();
    if (!id) return null;
    return runsCache.find((r) => Number(r.id) === id) || null;
  }

  function updateServiceFromRun() {
    const run = selectedRun();
    if (serviceAuto) serviceAuto.value = run?.service || '';
  }

  function clearDashboardView(message = 'Selecciona un periodo para consultar dashboard.') {
    lastPayload = null;
    if (kpisEl) kpisEl.innerHTML = '<div class="empty-mini">Sin datos para mostrar.</div>';
    if (notFoundSpecificContainer) notFoundSpecificContainer.innerHTML = '<div class="empty-mini">Sin datos para mostrar.</div>';
    if (notFoundSpecificStatus) notFoundSpecificStatus.textContent = '';
    if (serviceChart) { serviceChart.destroy(); serviceChart = null; }
    if (typeChart) { typeChart.destroy(); typeChart = null; }
    if (areaChart) { areaChart.destroy(); areaChart = null; }
    if (statusPieChart) { statusPieChart.destroy(); statusPieChart = null; }
    App.setStatus(statusEl, message, true);
  }

  function shortenLabel(text, maxLen = 34) {
    const s = String(text || '');
    if (s.length <= maxLen) return s;
    return `${s.slice(0, maxLen - 1)}...`;
  }

  function getBarGradient(ctx, chartArea, index) {
    const pair = DASHBOARD_CHART_THEME.barPalette[index % DASHBOARD_CHART_THEME.barPalette.length];
    const gradient = ctx.createLinearGradient(chartArea.left, 0, chartArea.right, 0);
    gradient.addColorStop(0, pair[0]);
    gradient.addColorStop(1, pair[1]);
    return gradient;
  }

  function renderBarChart(instance, canvasEl, rows, label, topN) {
    if (instance) instance.destroy();
    const limitedRows = (rows || []).slice(0, topN);
    const dynamicHeight = Math.max(320, limitedRows.length * 30);
    const wrapper = canvasEl.closest('.chart-wrap');
    if (wrapper) wrapper.style.height = `${Math.min(dynamicHeight, 560)}px`;
    canvasEl.height = dynamicHeight;
    canvasEl.style.width = '100%';
    canvasEl.style.height = `${dynamicHeight}px`;

    const valueLabelPlugin = {
      id: `valueLabel-${label}`,
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const dataset = chart.data.datasets[0];
        const meta = chart.getDatasetMeta(0);
        ctx.save();
        ctx.fillStyle = '#1f5f8f';
        ctx.font = '700 11px "Segoe UI", sans-serif';
        ctx.textAlign = 'left';
        ctx.textBaseline = 'middle';
        meta.data.forEach((bar, i) => {
          const value = Number(dataset.data[i] || 0);
          const pos = bar.getProps(['x', 'y'], true);
          ctx.fillText(String(value), pos.x + 8, pos.y);
        });
        ctx.restore();
      },
    };

    return new Chart(canvasEl.getContext('2d'), {
      type: 'bar',
      plugins: [{
        id: 'clearCanvas',
        beforeDraw: (chart) => {
          const { ctx, width, height } = chart;
          ctx.clearRect(0, 0, width, height);
        },
      }, valueLabelPlugin],
      data: {
        labels: limitedRows.map((r) => String(r.name || '')),
        datasets: [{
          label,
          data: limitedRows.map((r) => Number(r.total || 0)),
          borderWidth: 1,
          borderColor: DASHBOARD_CHART_THEME.stroke,
          borderRadius: 8,
          borderSkipped: false,
          maxBarThickness: 24,
          backgroundColor: (context) => {
            const { chart, dataIndex } = context;
            const { ctx, chartArea } = chart;
            if (!chartArea) return '#1f5f8f';
            return getBarGradient(ctx, chartArea, dataIndex);
          },
        }],
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        animation: { duration: 700, easing: 'easeOutQuart' },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: DASHBOARD_CHART_THEME.tooltipBg,
            titleColor: DASHBOARD_CHART_THEME.tooltipTitle,
            bodyColor: DASHBOARD_CHART_THEME.tooltipBody,
            borderColor: 'rgba(54, 132, 179, 0.5)',
            borderWidth: 1,
            padding: 10,
            displayColors: false,
            callbacks: {
              title: (items) => items?.[0]?.label || '',
              label: (item) => `${label}: ${Number(item.parsed.x || 0).toLocaleString('es-CO')}`,
            },
          },
        },
        scales: {
          x: {
            beginAtZero: true,
            grid: { color: DASHBOARD_CHART_THEME.grid, drawBorder: false },
            ticks: {
              color: DASHBOARD_CHART_THEME.axisText,
              font: { size: 11, weight: '600' },
            },
          },
          y: {
            grid: { display: false, drawBorder: false },
            ticks: {
              autoSkip: false,
              callback: (_, idx, ticks) => shortenLabel(ticks[idx].label),
              color: DASHBOARD_CHART_THEME.axisText,
              font: { size: 11, weight: '600' },
            },
          },
        },
      },
    });
  }

  function renderStatusPie(instance, canvasEl, kpis) {
    if (!canvasEl || typeof Chart === 'undefined') return instance;
    if (instance) instance.destroy();
    const found = Number(kpis?.found || 0);
    const notFound = Number(kpis?.not_found || 0);
    const pending = Number(kpis?.pending || 0);
    return new Chart(canvasEl.getContext('2d'), {
      type: 'doughnut',
      data: {
        labels: ['Encontrados', 'No encontrados', 'Pendientes'],
        datasets: [{
          data: [found, notFound, pending],
          backgroundColor: DASHBOARD_CHART_THEME.piePalette,
          borderColor: '#ffffff',
          borderWidth: 3,
          hoverOffset: 8,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '58%',
        animation: { duration: 800, easing: 'easeOutQuart' },
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              usePointStyle: true,
              pointStyle: 'circle',
              padding: 16,
              color: DASHBOARD_CHART_THEME.axisText,
              font: { size: 12, weight: '700' },
            },
          },
          tooltip: {
            backgroundColor: DASHBOARD_CHART_THEME.tooltipBg,
            titleColor: DASHBOARD_CHART_THEME.tooltipTitle,
            bodyColor: DASHBOARD_CHART_THEME.tooltipBody,
            borderColor: 'rgba(54, 132, 179, 0.5)',
            borderWidth: 1,
            padding: 10,
            callbacks: {
              label: (ctx) => {
                const value = Number(ctx.parsed || 0);
                const total = [found, notFound, pending].reduce((a, b) => a + b, 0);
                const pct = total ? ((value / total) * 100).toFixed(1) : '0.0';
                return `${ctx.label}: ${value.toLocaleString('es-CO')} (${pct}%)`;
              },
            },
          },
        },
      },
    });
  }

  async function loadRunsForPeriod() {
    const periodId = periodSelect?.value ? Number(periodSelect.value) : null;
    if (!periodId) {
      runsCache = [];
      runSelect.innerHTML = '<option value="">-- Todas las jornadas del periodo --</option>';
      updateServiceFromRun();
      return;
    }
    const params = new URLSearchParams();
    if (periodId) params.set('period_id', String(periodId));
    const data = await App.get('/runs' + (params.toString() ? `?${params.toString()}` : ''));
    runsCache = data.runs || [];
    runSelect.innerHTML = '<option value="">-- Todas las jornadas del periodo --</option>' + runsCache.map((r) => {
      const svc = r.service ? ` [${App.escapeHtml(r.service)}]` : '';
      const st = r.status === 'active' ? 'Activa' : (r.status === 'cancelled' ? 'Anulada' : 'Cerrada');
      return `<option value="${r.id}">${r.id} - ${App.escapeHtml(r.name)}${svc} - ${st}</option>`;
    }).join('');
    if (runsCache.length === 1) runSelect.value = String(runsCache[0].id);
    updateServiceFromRun();
  }

  async function loadComparePeriods() {
    const data = await App.get('/periods');
    const options = '<option value="">Selecciona periodo</option>' + (data.periods || [])
      .map((p) => `<option value="${p.id}">${App.escapeHtml(p.name)}${p.status === 'open' ? ' (Abierto)' : (p.status === 'cancelled' ? ' (Anulado)' : ' (Cerrado)')}</option>`)
      .join('');
    if (comparePeriodA) comparePeriodA.innerHTML = options;
    if (comparePeriodB) comparePeriodB.innerHTML = options;
  }

  function renderCompare(data) {
    if (!compareContainer) return;
    const a = data.period_a || {};
    const b = data.period_b || {};
    const d = data.delta || {};
    compareContainer.innerHTML = `
      <table class="mini-status-table">
        <thead><tr><th>Indicador</th><th>${App.escapeHtml(a.name || 'Periodo A')}</th><th>${App.escapeHtml(b.name || 'Periodo B')}</th><th>Diferencia</th></tr></thead>
        <tbody>
          <tr><td>Total activos</td><td>${a.total || 0}</td><td>${b.total || 0}</td><td>${d.total || 0}</td></tr>
          <tr><td>Encontrados</td><td>${a.found || 0}</td><td>${b.found || 0}</td><td>${d.found || 0}</td></tr>
          <tr><td>No encontrados</td><td>${a.not_found || 0}</td><td>${b.not_found || 0}</td><td>${d.not_found || 0}</td></tr>
          <tr><td>% Encontrados</td><td>${a.found_pct || 0}%</td><td>${b.found_pct || 0}%</td><td>${d.found_pct || 0}%</td></tr>
          <tr><td>% No encontrados</td><td>${a.not_found_pct || 0}%</td><td>${b.not_found_pct || 0}%</td><td>${d.not_found_pct || 0}%</td></tr>
        </tbody>
      </table>
    `;
  }

  function renderNotFoundSpecific(data) {
    if (!notFoundSpecificContainer) return;
    const rows = data.not_found_assets || [];
    const total = Number(data.not_found_assets_total || rows.length || 0);
    if (notFoundSpecificStatus) {
      const capped = data.not_found_assets_capped ? ` (mostrando ${rows.length})` : '';
      notFoundSpecificStatus.textContent = `Total no encontrados: ${total}${capped}`;
    }
    if (!rows.length) {
      notFoundSpecificContainer.innerHTML = '<div class="empty-mini">No hay activos no encontrados en este corte.</div>';
      return;
    }
    notFoundSpecificContainer.innerHTML = `
      <div class="closed-runs-wrap compact-notfound-wrap">
        <table class="compact-notfound-table">
          <thead>
            <tr>
              <th>CODIGO</th>
              <th>ACTIVO</th>
              <th>SERVICIO</th>
              <th>RESPONSABLE</th>
              <th>UBICACION</th>
              <th>VALOR LIBRO</th>
            </tr>
          </thead>
          <tbody>
            ${rows.slice(0, 120).map((r) => `<tr>
              <td>${App.escapeHtml(r.code || '')}</td>
              <td>${App.escapeHtml(r.name || '')}</td>
              <td>${App.escapeHtml(r.service || '')}</td>
              <td>${App.escapeHtml(r.responsible || '')}</td>
              <td>${App.escapeHtml(r.location || '')}</td>
              <td>${new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', maximumFractionDigits: 0 }).format(Number(r.value || 0))}</td>
            </tr>`).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  async function comparePeriods() {
    const a = comparePeriodA?.value || '';
    const b = comparePeriodB?.value || '';
    if (!a || !b) {
      App.setStatus(compareStatus, 'Selecciona ambos periodos para comparar', true);
      return;
    }
    const params = new URLSearchParams();
    params.set('period_a', a);
    params.set('period_b', b);
    const data = await App.get('/dashboard/compare_periods?' + params.toString());
    renderCompare(data);
    App.setStatus(compareStatus, `Comparativo generado: ${data.period_a?.name || ''} vs ${data.period_b?.name || ''}`);
  }

  async function loadDashboard() {
    if (!periodSelect.value) {
      clearDashboardView('Selecciona un periodo para consultar dashboard');
      return;
    }
    const params = [];
    if (periodSelect.value) params.push(`period_id=${encodeURIComponent(periodSelect.value)}`);
    if (selectedRunId()) params.push(`run_id=${encodeURIComponent(selectedRunId())}`);
    const url = '/dashboard/summary' + (params.length ? `?${params.join('&')}` : '');
    const data = await App.get(url);
    lastPayload = data;
    renderDashboardFromPayload();
  }

  function renderDashboardFromPayload() {
    if (!lastPayload) return;
    const data = lastPayload;
    const k = data.kpis || {};
    const c = data.coverage || {};
    kpisEl.innerHTML = `
      <div>Total: ${k.total || 0} | Encontrados: ${k.found || 0} (${k.found_pct || 0}%) | No encontrados: ${k.not_found || 0} (${k.not_found_pct || 0}%) | Pendientes: ${k.pending || 0}</div>
      <div style="margin-top:6px;color:#345266;font-size:13px;">
        Cobertura sobre base: ${c.scope_assets || 0}/${c.base_total_assets || 0} activos (${c.scope_assets_pct || 0}%) |
        Valor cubierto: ${new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', maximumFractionDigits: 0 }).format(Number(c.scope_value || 0))}
        / ${new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', maximumFractionDigits: 0 }).format(Number(c.base_total_value || 0))}
        (${c.scope_value_pct || 0}%)
      </div>
    `;
    const topService = Number(topServiceEl?.value || 15);
    const topType = Number(topTypeEl?.value || 15);
    const topArea = Number(topAreaEl?.value || 10);
    serviceChart = renderBarChart(serviceChart, serviceChartEl, data.by_service || [], 'Servicios', topService);
    typeChart = renderBarChart(typeChart, typeChartEl, data.by_type || [], 'Tipos', topType);
    areaChart = renderBarChart(areaChart, areaChartEl, data.by_area || [], 'Areas', topArea);
    statusPieChart = renderStatusPie(statusPieChart, statusPieChartEl, k);
    renderNotFoundSpecific(data);
    App.setStatus(statusEl, `Actualizado: ${data.meta?.generated_at || ''}`);
  }

  refreshBtn?.addEventListener('click', () => loadDashboard().catch((e) => App.setStatus(statusEl, e.message, true)));
  periodSelect?.addEventListener('change', () => {
    loadRunsForPeriod().then(loadDashboard).catch((e) => App.setStatus(statusEl, e.message, true));
  });
  runSelect?.addEventListener('change', () => {
    updateServiceFromRun();
    loadDashboard().catch((e) => App.setStatus(statusEl, e.message, true));
  });
  topServiceEl?.addEventListener('change', renderDashboardFromPayload);
  topTypeEl?.addEventListener('change', renderDashboardFromPayload);
  topAreaEl?.addEventListener('change', renderDashboardFromPayload);
  comparePeriodsBtn?.addEventListener('click', () => comparePeriods().catch((e) => App.setStatus(compareStatus, e.message, true)));

  exportPdfBtn?.addEventListener('click', () => {
    if (!periodSelect.value) {
      App.setStatus(statusEl, 'Selecciona periodo para exportar PDF', true);
      return;
    }
    const params = [];
    if (periodSelect.value) params.push(`period_id=${encodeURIComponent(periodSelect.value)}`);
    if (selectedRunId()) params.push(`run_id=${encodeURIComponent(selectedRunId())}`);
    const url = '/dashboard/report_pdf' + (params.length ? `?${params.join('&')}` : '');
    window.location = url;
  });

  Promise.all([
    App.loadPeriods(periodSelect),
    loadComparePeriods(),
  ])
    .then(loadRunsForPeriod)
    .then(loadDashboard)
    .catch((err) => App.setStatus(statusEl, err.message, true));
});
