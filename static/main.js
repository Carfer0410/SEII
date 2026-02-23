document.addEventListener('DOMContentLoaded', () => {
  const uploadForm = document.getElementById('uploadForm');
  const fileInput = document.getElementById('fileInput');
  const importResult = document.getElementById('importResult');
  const serviceSelect = document.getElementById('serviceSelect');
  const refreshBtn = document.getElementById('refreshBtn');
  const assetsContainer = document.getElementById('assetsContainer');
  const scanInput = document.getElementById('scanInput');
  const startCamera = document.getElementById('startCamera');
  const stopCamera = document.getElementById('stopCamera');
  const exportBtn = document.getElementById('exportBtn');
  const receiver = document.getElementById('receiver');
  const importStatus = document.getElementById('importStatus');
  const scanStatus = document.getElementById('scanStatus');
  const runName = document.getElementById('runName');
  const runSelect = document.getElementById('runSelect');
  const createRunBtn = document.getElementById('createRunBtn');
  const refreshRunsBtn = document.getElementById('refreshRunsBtn');
  const closeRunBtn = document.getElementById('closeRunBtn');
  const runSummary = document.getElementById('runSummary');
  const disposalCode = document.getElementById('disposalCode');
  const disposalReason = document.getElementById('disposalReason');
  const createDisposalBtn = document.getElementById('createDisposalBtn');
  const disposalStatusFilter = document.getElementById('disposalStatusFilter');
  const refreshDisposalsBtn = document.getElementById('refreshDisposalsBtn');
  const disposalStatus = document.getElementById('disposalStatus');
  const disposalsContainer = document.getElementById('disposalsContainer');
  const refreshDashboardBtn = document.getElementById('refreshDashboardBtn');
  const exportDashboardPdfBtn = document.getElementById('exportDashboardPdfBtn');
  const dashboardStatus = document.getElementById('dashboardStatus');
  const dashboardKpis = document.getElementById('dashboardKpis');
  const serviceChartEl = document.getElementById('serviceChart');
  const typeChartEl = document.getElementById('typeChart');
  const areaChartEl = document.getElementById('areaChart');

  let html5QrcodeScanner = null;
  let lastScanCode = null;
  let lastScanAt = 0;
  let serviceChart = null;
  let typeChart = null;
  let areaChart = null;

  function on(el, eventName, handler) {
    if (!el) return;
    el.addEventListener(eventName, handler);
  }

  function setStatus(el, message, isError = false) {
    if (!el) return;
    el.textContent = message || '';
    el.style.color = isError ? '#b00020' : '#004d40';
  }

  function getSelectedRunId() {
    const v = runSelect?.value || '';
    return v ? Number(v) : null;
  }

  async function parseJsonOrThrow(res) {
    const payload = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(payload.error || `Error HTTP ${res.status}`);
    return payload;
  }

  on(uploadForm, 'submit', async (e) => {
    e.preventDefault();
    if (!fileInput.files.length) {
      alert('Seleccione un archivo');
      return;
    }

    const fd = new FormData();
    fd.append('file', fileInput.files[0]);

    try {
      setStatus(importStatus, 'Importando archivo...');
      const res = await fetch('/import', { method: 'POST', body: fd });
      const j = await parseJsonOrThrow(res);
      importResult.textContent = JSON.stringify(j);
      setStatus(importStatus, 'Importacion completada.');
      await loadServices();
      await loadAssets();
      await loadDisposals();
      await loadDashboard();
    } catch (err) {
      setStatus(importStatus, err.message || 'Error al importar archivo', true);
    }
  });

  on(refreshBtn, 'click', async () => {
    try {
      await loadServices();
      await loadAssets();
      await loadDisposals();
      await refreshSummary();
      await loadDashboard();
      setStatus(scanStatus, '');
    } catch (err) {
      setStatus(scanStatus, err.message || 'Error al refrescar', true);
    }
  });

  on(serviceSelect, 'change', async () => {
    try {
      await loadAssets();
      await refreshSummary();
      await loadDisposals();
      await loadDashboard();
    } catch (err) {
      setStatus(scanStatus, err.message || 'Error al filtrar activos', true);
    }
  });

  on(runSelect, 'change', async () => {
    try {
      await loadAssets();
      await refreshSummary();
      await loadDashboard();
    } catch (err) {
      setStatus(scanStatus, err.message || 'Error al cambiar jornada', true);
    }
  });

  on(refreshRunsBtn, 'click', async () => {
    try {
      await loadRuns();
      await refreshSummary();
      await loadDisposals();
      await loadDashboard();
    } catch (err) {
      setStatus(scanStatus, err.message || 'Error al cargar jornadas', true);
    }
  });

  on(createRunBtn, 'click', async () => {
    const name = (runName.value || '').trim();
    if (!name) {
      setStatus(scanStatus, 'Debe escribir nombre de jornada', true);
      return;
    }
    try {
      const res = await fetch('/runs', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name,
          service: serviceSelect.value || null,
          created_by: 'usuario_movil',
        }),
      });
      const j = await parseJsonOrThrow(res);
      runName.value = '';
      await loadRuns(j.run.id);
      await loadAssets();
      await refreshSummary();
      await loadDisposals();
      await loadDashboard();
      setStatus(scanStatus, `Jornada creada: ${j.run.name}`);
    } catch (err) {
      setStatus(scanStatus, err.message || 'Error creando jornada', true);
    }
  });

  on(closeRunBtn, 'click', async () => {
    const runId = getSelectedRunId();
    if (!runId) {
      setStatus(scanStatus, 'Seleccione una jornada para cerrar', true);
      return;
    }
    if (!confirm('Va a cerrar la jornada y marcar "No encontrado" los pendientes. Continuar?')) return;
    try {
      const res = await fetch(`/runs/${runId}/close`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ user: 'usuario_movil' }),
      });
      const j = await parseJsonOrThrow(res);
      await loadRuns(runId);
      await loadAssets();
      await refreshSummary();
      await loadDisposals();
      await loadDashboard();
      setStatus(scanStatus, `Jornada cerrada. No encontrados auto: ${j.auto_marked_not_found}`);
    } catch (err) {
      setStatus(scanStatus, err.message || 'Error cerrando jornada', true);
    }
  });

  on(createDisposalBtn, 'click', async () => {
    const code = (disposalCode.value || '').trim();
    const reason = (disposalReason.value || '').trim();
    if (!code) {
      setStatus(disposalStatus, 'Debe escribir codigo de activo', true);
      return;
    }
    try {
      const res = await fetch('/disposals', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          code,
          reason,
          requested_by: 'usuario_movil',
        }),
      });
      const j = await parseJsonOrThrow(res);
      disposalCode.value = '';
      disposalReason.value = '';
      await loadDisposals();
      await loadAssets();
      await loadDashboard();
      setStatus(disposalStatus, `Marcado para baja: ${j.disposal.asset.C_ACT}`);
    } catch (err) {
      setStatus(disposalStatus, err.message || 'Error marcando baja', true);
    }
  });

  on(refreshDisposalsBtn, 'click', async () => {
    try {
      await loadDisposals();
      setStatus(disposalStatus, '');
    } catch (err) {
      setStatus(disposalStatus, err.message || 'Error cargando bajas', true);
    }
  });

  on(disposalStatusFilter, 'change', loadDisposals);

  on(refreshDashboardBtn, 'click', async () => {
    try {
      await loadDashboard();
    } catch (err) {
      setStatus(dashboardStatus, err.message || 'Error cargando dashboard', true);
    }
  });

  on(exportDashboardPdfBtn, 'click', () => {
    const service = serviceSelect ? serviceSelect.value : '';
    const runId = getSelectedRunId();
    let url = '/dashboard/report_pdf';
    const params = [];
    if (service) params.push(`service=${encodeURIComponent(service)}`);
    if (runId) params.push(`run_id=${encodeURIComponent(runId)}`);
    if (params.length) url += `?${params.join('&')}`;
    window.location = url;
  });

  on(disposalsContainer, 'click', async (e) => {
    const button = e.target.closest('button[data-disposal-id]');
    if (!button) return;
    const disposalId = Number(button.getAttribute('data-disposal-id'));
    const nextStatus = button.getAttribute('data-next-status');
    if (!disposalId || !nextStatus) return;
    try {
      const res = await fetch(`/disposals/${disposalId}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          status: nextStatus,
          reviewed_by: 'usuario_movil',
        }),
      });
      await parseJsonOrThrow(res);
      await loadDisposals();
      await loadAssets();
      await loadDashboard();
      setStatus(disposalStatus, `Estado actualizado: ${nextStatus}`);
    } catch (err) {
      setStatus(disposalStatus, err.message || 'Error actualizando baja', true);
    }
  });

  async function loadServices() {
    if (!serviceSelect) return;
    const res = await fetch('/services');
    const j = await parseJsonOrThrow(res);
    serviceSelect.innerHTML =
      '<option value="">-- Todos --</option>' +
      j.services.map((s) => `<option>${escapeHtml(s)}</option>`).join('');
  }

  async function loadRuns(preferredRunId = null) {
    if (!runSelect) return;
    const previous = preferredRunId || getSelectedRunId();
    const res = await fetch('/runs');
    const j = await parseJsonOrThrow(res);
    runSelect.innerHTML = '<option value="">-- Sin jornada --</option>' + j.runs
      .map((r) => {
        const status = r.status === 'active' ? 'Activa' : 'Cerrada';
        const svc = r.service ? ` [${r.service}]` : '';
        return `<option value="${r.id}">${r.id} - ${escapeHtml(r.name)}${svc} - ${status}</option>`;
      })
      .join('');
    if (previous) {
      runSelect.value = String(previous);
    }
  }

  async function refreshSummary() {
    if (!runSummary) return;
    const runId = getSelectedRunId();
    if (!runId) {
      runSummary.textContent = 'Sin jornada seleccionada.';
      return;
    }
    const res = await fetch(`/runs/${runId}/summary`);
    const j = await parseJsonOrThrow(res);
    const s = j.summary;
    runSummary.textContent = `Total: ${s.total} | Encontrados: ${s.found} | No encontrados: ${s.not_found} | Pendientes: ${s.pending}`;
  }

  async function loadAssets() {
    if (!assetsContainer) return;
    const service = serviceSelect ? serviceSelect.value : '';
    const runId = getSelectedRunId();
    let url = '/assets';
    const params = [];
    if (service) params.push(`service=${encodeURIComponent(service)}`);
    if (runId) params.push(`run_id=${encodeURIComponent(runId)}`);
    if (params.length) url += `?${params.join('&')}`;
    const res = await fetch(url);
    const j = await parseJsonOrThrow(res);
    assetsContainer.innerHTML = renderTable(j.assets || []);
  }

  async function loadDisposals() {
    if (!disposalsContainer || !disposalStatusFilter) return;
    const service = serviceSelect ? serviceSelect.value : '';
    const status = disposalStatusFilter.value;
    let url = '/disposals';
    const params = [];
    if (service) params.push(`service=${encodeURIComponent(service)}`);
    if (status) params.push(`status=${encodeURIComponent(status)}`);
    if (params.length) url += `?${params.join('&')}`;
    const res = await fetch(url);
    const j = await parseJsonOrThrow(res);
    disposalsContainer.innerHTML = renderDisposals(j.disposals || []);
  }

  async function loadDashboard() {
    if (!dashboardKpis) return;
    const service = serviceSelect ? serviceSelect.value : '';
    const runId = getSelectedRunId();
    let url = '/dashboard/summary';
    const params = [];
    if (service) params.push(`service=${encodeURIComponent(service)}`);
    if (runId) params.push(`run_id=${encodeURIComponent(runId)}`);
    if (params.length) url += `?${params.join('&')}`;
    const res = await fetch(url);
    const j = await parseJsonOrThrow(res);
    renderDashboard(j);
  }

  function renderDashboard(data) {
    if (!dashboardKpis) return;
    const k = data.kpis || {};
    dashboardKpis.innerHTML =
      `Total: ${k.total || 0} | Encontrados: ${k.found || 0} (${k.found_pct || 0}%) | ` +
      `No encontrados: ${k.not_found || 0} (${k.not_found_pct || 0}%) | Pendientes: ${k.pending || 0}`;

    if (serviceChartEl) {
      serviceChart = renderBarChart(serviceChart, serviceChartEl, data.by_service || [], 'Servicios');
    }
    if (typeChartEl) {
      typeChart = renderBarChart(typeChart, typeChartEl, (data.by_type || []).slice(0, 12), 'Tipos');
    }
    if (areaChartEl) {
      areaChart = renderBarChart(areaChart, areaChartEl, data.by_area || [], 'Areas');
    }

    const generatedAt = data.meta?.generated_at_local || data.meta?.generated_at || '';
    const formattedGeneratedAt = (window.App?.formatDateTime ? window.App.formatDateTime(generatedAt) : generatedAt);
    setStatus(dashboardStatus, formattedGeneratedAt ? `Actualizado: ${formattedGeneratedAt}` : '');
  }

  function renderBarChart(instance, canvasEl, rows, label) {
    if (!canvasEl || typeof Chart === 'undefined') return instance;
    if (instance) instance.destroy();
    const labels = rows.map((r) => r.name);
    const values = rows.map((r) => r.total);
    return new Chart(canvasEl.getContext('2d'), {
      type: 'bar',
      data: {
        labels,
        datasets: [{
          label,
          data: values,
          backgroundColor: '#1976d2',
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
        },
        scales: {
          y: { beginAtZero: true },
        },
      },
    });
  }

  function renderTable(items) {
    if (!items.length) return '<div>No hay activos.</div>';
    const headers = Object.keys(items[0]);
    return `<table><thead><tr>${headers
      .map((h) => `<th>${escapeHtml(App.formatHeader ? App.formatHeader(h) : h)}</th>`)
      .join('')}</tr></thead><tbody>${items
      .map(
        (r) =>
          `<tr>${headers
            .map((h) => {
              const raw = String(r[h] || '');
              const formatted = (window.App?.formatDateTime ? window.App.formatDateTime(raw) : raw);
              return `<td>${escapeHtml(formatted)}</td>`;
            })
            .join('')}</tr>`
      )
      .join('')}</tbody></table>`;
  }

  function escapeHtml(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function renderDisposals(items) {
    if (!items.length) return '<div>No hay activos en bajas pendientes.</div>';
    const headers = ['Cod activo', 'Descripcion', 'Servicio', 'Tipo activo', 'Saldo por depreciar', 'Estado baja', 'Motivo', 'Solicitado por', 'Fecha solicitud', 'Acciones'];
    return `<table><thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>${items.map((row) => {
      const a = row.asset || {};
      const approveBtn = `<button data-disposal-id="${row.id}" data-next-status="Aprobada para baja">Aprobar</button>`;
      const rejectBtn = `<button data-disposal-id="${row.id}" data-next-status="Rechazada">Rechazar</button>`;
      const pendingBtn = `<button data-disposal-id="${row.id}" data-next-status="Pendiente baja">Pendiente</button>`;
      const actions = `${approveBtn} ${rejectBtn} ${pendingBtn}`;
      return `<tr>
        <td>${escapeHtml(String(a.C_ACT || ''))}</td>
        <td>${escapeHtml(String(a.NOM || ''))}</td>
        <td>${escapeHtml(String(a.NOM_CCOS || ''))}</td>
        <td>${escapeHtml(String(a.DESC_TIAC || ''))}</td>
        <td>${escapeHtml(String(a.SALDO || ''))}</td>
        <td>${escapeHtml(String(row.status || ''))}</td>
        <td>${escapeHtml(String(row.reason || ''))}</td>
        <td>${escapeHtml(String(row.requested_by || ''))}</td>
        <td>${escapeHtml(window.App?.formatDateTime ? window.App.formatDateTime(String(row.requested_at || '')) : String(row.requested_at || ''))}</td>
        <td>${actions}</td>
      </tr>`;
    }).join('')}</tbody></table>`;
  }

  let scanInputTimer = null;

  async function flushScanInput() {
    const code = scanInput?.value?.trim() || '';
    if (scanInput) scanInput.value = '';
    if (!code) return;

    try {
      await sendScan(code);
      await loadAssets();
      await refreshSummary();
      await loadDashboard();
    } catch (err) {
      setStatus(scanStatus, err.message || 'Error en escaneo', true);
    }
  }

  on(scanInput, 'keydown', async (e) => {
    if (e.key === 'Enter' || e.key === 'NumpadEnter' || e.key === 'Tab') {
      e.preventDefault();
      if (scanInputTimer) {
        clearTimeout(scanInputTimer);
        scanInputTimer = null;
      }
      await flushScanInput();
    }
  });

  on(scanInput, 'input', () => {
    if (scanInputTimer) clearTimeout(scanInputTimer);
    scanInputTimer = setTimeout(() => {
      flushScanInput();
      scanInputTimer = null;
    }, 120);
  });

  async function sendScan(code) {
    const now = Date.now();
    if (code === lastScanCode && now - lastScanAt < 1500) return;
    lastScanCode = code;
    lastScanAt = now;

    const res = await fetch('/scan', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        code,
        user: 'usuario_movil',
        run_id: getSelectedRunId(),
      }),
    });
    const j = await parseJsonOrThrow(res);
    if (j.found) setStatus(scanStatus, `Encontrado: ${j.asset.C_ACT}`);
    else setStatus(scanStatus, `Codigo no encontrado en la base: ${code}`, true);
  }

  on(startCamera, 'click', () => {
    if (html5QrcodeScanner) return;
    html5QrcodeScanner = new Html5Qrcode('reader');
    html5QrcodeScanner
      .start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: 250 },
        async (decoded) => {
          try {
            await sendScan(decoded);
            await loadAssets();
            await refreshSummary();
            await loadDashboard();
          } catch (err) {
            setStatus(scanStatus, err.message || 'Error procesando lectura', true);
          }
        },
        () => {}
      )
      .catch((err) => setStatus(scanStatus, `Error camara: ${err}`, true));
  });

  on(stopCamera, 'click', () => {
    if (!html5QrcodeScanner) return;
    html5QrcodeScanner.stop().then(() => {
      html5QrcodeScanner.clear();
      html5QrcodeScanner = null;
    });
  });

  on(exportBtn, 'click', () => {
    const service = serviceSelect ? serviceSelect.value : '';
    const rec = receiver ? receiver.value : '';
    const runId = getSelectedRunId();
    const url =
      '/export' +
      (service ? `?service=${encodeURIComponent(service)}` : '') +
      (runId ? ((service ? '&' : '?') + `run_id=${encodeURIComponent(runId)}`) : '') +
      (rec ? (((service || runId) ? '&' : '?') + `receiver=${encodeURIComponent(rec)}`) : '');
    window.location = url;
  });

  loadServices()
    .then(() => loadRuns())
    .then(() => loadAssets())
    .then(() => loadDisposals())
    .then(() => refreshSummary())
    .then(() => loadDashboard())
    .catch((err) => setStatus(scanStatus, err.message || 'Error al iniciar', true));
});
