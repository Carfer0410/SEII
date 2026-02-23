document.addEventListener('DOMContentLoaded', () => {
  const serviceSelect = document.getElementById('serviceSelect');
  const periodSelect = document.getElementById('periodSelect');
  const refreshBtn = document.getElementById('refreshBtn');
  const exportFoundBtn = document.getElementById('exportFoundBtn');
  const exportNotFoundBtn = document.getElementById('exportNotFoundBtn');
  const exportConsolidatedBtn = document.getElementById('exportConsolidatedBtn');
  const assetsContainer = document.getElementById('assetsContainer');
  const scanInput = document.getElementById('scanInput');
  const startCamera = document.getElementById('startCamera');
  const stopCamera = document.getElementById('stopCamera');
  const scanStatus = document.getElementById('scanStatus');
  const activeRunBannerText = document.getElementById('activeRunBannerText');

  const kpiTotalAssets = document.getElementById('kpiTotalAssets');
  const kpiFoundAssets = document.getElementById('kpiFoundAssets');
  const kpiNotFoundAssets = document.getElementById('kpiNotFoundAssets');
  const kpiPendingAssets = document.getElementById('kpiPendingAssets');
  const foundAssetsContainer = document.getElementById('foundAssetsContainer');
  const notFoundAssetsContainer = document.getElementById('notFoundAssetsContainer');
  const pendingAssetsContainer = document.getElementById('pendingAssetsContainer');
  const opPeriod = document.getElementById('opPeriod');
  const opRun = document.getElementById('opRun');
  const opService = document.getElementById('opService');
  const opLastScan = document.getElementById('opLastScan');
  const exportStateHelp = document.getElementById('exportStateHelp');
  const workflowHints = document.getElementById('workflowHints');
  const startGuideBtn = document.getElementById('startGuideBtn');

  let scanner = null;
  let lastScanCode = null;
  let lastScanAt = 0;
  let currentAssets = [];
  let assetsPage = 1;
  let assetsPageSize = 10;
  let assetsSearchTerm = '';
  const miniPages = { found: 1, not_found: 1, pending: 1 };
  const miniPageSizes = { found: 10, not_found: 10, pending: 10 };
  let realtimeTimer = null;
  let realtimeBusy = false;
  let lastRealtimeHash = '';
  let activeRunId = null;
  let lastScanDisplay = '-';
  let guideIndex = 0;
  let guideSteps = [];
  let guidePopover = null;
  let exportBusy = false;
  let loadingOverlay = null;
  let disposalReasonOverlay = null;
  let actionResultOverlay = null;

  const classificationOptions = [
    'Pendiente verificacion',
    'En mantenimiento',
    'Prestado',
    'Activo de control',
    'Para baja',
    'Baja aprobada',
  ];

  function normalizeInventoryStatus(value) {
    const s = String(value || '').trim().toUpperCase();
    if (s === 'ENCONTRADO') return 'Encontrado';
    if (s === 'NO ENCONTRADO') return 'No encontrado';
    return 'Pendiente';
  }

  function renderAssetsTable(items) {
    const term = String(assetsSearchTerm || '').trim().toUpperCase();
    const filteredItems = (items || []).filter((row) => {
      if (!term) return true;
      const haystack = [
        row.C_ACT,
        row.NOM,
        row.MODELO,
        row.REF,
        row.SERIE,
        row.NOM_CCOS,
        row.DES_UBI,
        row.NOM_RESP,
        row.estado_inventario,
        row.estado_jornada,
      ].map((v) => String(v || '').toUpperCase()).join(' | ');
      return haystack.includes(term);
    });
    if (!filteredItems.length) {
      return `
        <div class="table-toolbar">
          <div class="field-help">Mostrando 0 de ${items.length} activos</div>
          <div class="pagination-row">
            <label class="field-help">Buscar:</label>
            <input type="text" class="assets-search-input" data-assets-search value="${App.escapeHtml(assetsSearchTerm)}" placeholder="Buscar por codigo, descripcion, servicio, ubicacion..." />
          </div>
        </div>
        <div class="empty-mini">No hay activos para ese filtro de busqueda.</div>
      `;
    }

    const total = filteredItems.length;
    const totalPages = Math.max(1, Math.ceil(total / assetsPageSize));
    if (assetsPage > totalPages) assetsPage = totalPages;
    if (assetsPage < 1) assetsPage = 1;
    const start = (assetsPage - 1) * assetsPageSize;
    const end = Math.min(start + assetsPageSize, total);
    const pageItems = filteredItems.slice(start, end);

    const visibleHeaders = ['C_ACT', 'NOM', 'MODELO', 'REF', 'SERIE', 'NOM_CCOS', 'DES_UBI', 'NOM_RESP', 'estado_inventario', 'estado_jornada', 'estado_baja'];
    const head = ['Clasificacion'].concat(visibleHeaders.map((h) => App.formatHeader(h)));
    return `
      <div class="table-toolbar">
        <div class="field-help">Mostrando ${start + 1}-${end} de ${total} activos</div>
        <div class="pagination-row">
          <label class="field-help">Buscar:</label>
          <input type="text" class="assets-search-input" data-assets-search value="${App.escapeHtml(assetsSearchTerm)}" placeholder="Buscar por codigo, descripcion, servicio, ubicacion..." />
          <label class="field-help">Filas:</label>
          <select class="page-size-select" data-page-size-select>
            ${[10, 25, 50, 100, 200].map((n) => `<option value="${n}" ${n === assetsPageSize ? 'selected' : ''}>${n}</option>`).join('')}
          </select>
          <button type="button" class="pager-btn" data-page-action="prev" ${assetsPage <= 1 ? 'disabled' : ''}>Anterior</button>
          <span class="field-help">Pagina ${assetsPage} de ${totalPages}</span>
          <button type="button" class="pager-btn" data-page-action="next" ${assetsPage >= totalPages ? 'disabled' : ''}>Siguiente</button>
        </div>
      </div>
      <table class="assets-table"><thead><tr>${head.map((h) => `<th>${App.escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>${pageItems.map((row) => {
      const selected = row.estado_inventario || 'Pendiente verificacion';
      const normalizedRunStatus = normalizeInventoryStatus(row.estado_jornada || '');
      const managedByRun = Boolean(row.gestionado_jornada);
      const lockedStatus = managedByRun && (normalizedRunStatus === 'Encontrado' || normalizedRunStatus === 'No encontrado')
        ? normalizedRunStatus
        : null;
      const select = lockedStatus
        ? `<span class="state-chip ${lockedStatus === 'Encontrado' ? 'chip-ok' : 'chip-bad'}">${App.escapeHtml(lockedStatus)}</span>`
        : `<select data-asset-id="${row.id}" class="asset-classification">
            ${classificationOptions.map((opt) => `<option value="${App.escapeHtml(opt)}" ${opt === selected ? 'selected' : ''}>${App.escapeHtml(opt)}</option>`).join('')}
          </select>`;
      const cells = visibleHeaders.map((h) => {
        if (h === 'estado_jornada') {
          const rawRun = String(row[h] || '').trim();
          if (!rawRun) {
            return '<td><span class="state-chip chip-info">-</span></td>';
          }
          const normalized = normalizeInventoryStatus(rawRun);
          const cls = normalized === 'Encontrado'
            ? 'state-chip chip-ok'
            : normalized === 'No encontrado'
              ? 'state-chip chip-bad'
              : 'state-chip chip-warn';
          return `<td><span class="${cls}">${App.escapeHtml(normalized)}</span></td>`;
        }
        if (h === 'estado_inventario') {
          const raw = String(row[h] || '');
          const normalized = normalizeInventoryStatus(raw);
          const cls = normalized === 'Encontrado'
            ? 'state-chip chip-ok'
            : normalized === 'No encontrado'
              ? 'state-chip chip-bad'
              : 'state-chip chip-warn';
          return `<td><span class="${cls}">${App.escapeHtml(normalized)}</span></td>`;
        }
        return `<td>${App.escapeHtml(row[h])}</td>`;
      }).join('');
      return `<tr><td>${select}</td>${cells}</tr>`;
    }).join('')}</tbody></table>`;
  }

  function renderMiniTable(key, items, emptyText) {
    if (!items.length) return `<div class="empty-mini">${App.escapeHtml(emptyText)}</div>`;
    const size = miniPageSizes[key] || 25;
    const total = items.length;
    const totalPages = Math.max(1, Math.ceil(total / size));
    if (miniPages[key] > totalPages) miniPages[key] = totalPages;
    if (miniPages[key] < 1) miniPages[key] = 1;
    const start = (miniPages[key] - 1) * size;
    const end = Math.min(start + size, total);
    const cut = items.slice(start, end);
    const rows = cut.map((a) => `<tr>
      <td>${App.escapeHtml(a.C_ACT)}</td>
      <td>${App.escapeHtml(a.NOM)}</td>
      <td>${App.escapeHtml(a.DES_UBI)}</td>
      <td>${App.escapeHtml(a.NOM_RESP)}</td>
    </tr>`).join('');
    return `
      <div class="table-toolbar mini-toolbar">
        <div class="field-help">Mostrando ${start + 1}-${end} de ${total} activos</div>
        <div class="pagination-row">
          <label class="field-help">Filas:</label>
          <select class="page-size-select" data-mini-page-size="${key}">
            ${[10, 25, 50, 100].map((n) => `<option value="${n}" ${n === size ? 'selected' : ''}>${n}</option>`).join('')}
          </select>
          <button type="button" class="pager-btn" data-mini-action="prev" data-mini-key="${key}" ${miniPages[key] <= 1 ? 'disabled' : ''}>Anterior</button>
          <span class="field-help">Pagina ${miniPages[key]} de ${totalPages}</span>
          <button type="button" class="pager-btn" data-mini-action="next" data-mini-key="${key}" ${miniPages[key] >= totalPages ? 'disabled' : ''}>Siguiente</button>
        </div>
      </div>
      <table class="mini-status-table"><thead><tr><th>CODIGO</th><th>ACTIVO</th><th>UBICACION</th><th>RESPONSABLE</th></tr></thead><tbody>${rows}</tbody></table>
    `;
  }

  function renderReconciliation(items) {
    const found = [];
    const notFound = [];
    const pending = [];
    (items || []).forEach((a) => {
      const st = normalizeInventoryStatus(a.estado_inventario);
      if (st === 'Encontrado') found.push(a);
      else if (st === 'No encontrado') notFound.push(a);
      else pending.push(a);
    });

    if (kpiTotalAssets) kpiTotalAssets.textContent = String(items.length);
    if (kpiFoundAssets) kpiFoundAssets.textContent = String(found.length);
    if (kpiNotFoundAssets) kpiNotFoundAssets.textContent = String(notFound.length);
    if (kpiPendingAssets) kpiPendingAssets.textContent = String(pending.length);

    if (foundAssetsContainer) foundAssetsContainer.innerHTML = renderMiniTable('found', found, 'No hay encontrados.');
    if (notFoundAssetsContainer) notFoundAssetsContainer.innerHTML = renderMiniTable('not_found', notFound, 'No hay no encontrados.');
    if (pendingAssetsContainer) pendingAssetsContainer.innerHTML = renderMiniTable('pending', pending, 'No hay pendientes.');
  }

  function getAssetsUrl() {
    const service = serviceSelect.value;
    const params = [];
    if (service) params.push(`service=${encodeURIComponent(service)}`);
    if (activeRunId) params.push(`run_id=${encodeURIComponent(activeRunId)}`);
    return '/assets' + (params.length ? `?${params.join('&')}` : '');
  }

  function renderRunBanner(run) {
    if (!activeRunBannerText) return;
    activeRunId = run ? run.id : null;
    if (!run) {
      activeRunBannerText.textContent = 'Sin jornada activa. Para escanear debes iniciar una jornada.';
      activeRunBannerText.className = 'run-banner-text run-banner-off';
      if (scanInput) scanInput.disabled = true;
      if (startCamera) startCamera.disabled = true;
      if (stopCamera) stopCamera.disabled = true;
      if (scanner) {
        scanner.stop().then(() => scanner.clear()).finally(() => { scanner = null; });
      }
      updateOperationalStrip();
      updateActionLocks();
      return;
    }
    const svcLabel = String(run.service_scope_label || run.service || '').trim();
    const svc = svcLabel ? ` | Servicio(s): ${svcLabel}` : '';
    const per = run.period_name ? ` | Periodo: ${run.period_name}` : '';
    activeRunBannerText.textContent = `Jornada activa: ${run.name}${per}${svc} | Inicio: ${String(run.started_at || '').slice(0, 16).replace('T', ' ')}`;
    activeRunBannerText.className = 'run-banner-text run-banner-on';
    if (scanInput) scanInput.disabled = false;
    if (startCamera) startCamera.disabled = false;
    if (stopCamera) stopCamera.disabled = !scanner;
    updateOperationalStrip();
    updateActionLocks();
  }

  function selectedLabel(selectEl, fallback = '-') {
    if (!selectEl) return fallback;
    const opt = selectEl.options[selectEl.selectedIndex];
    if (!opt) return fallback;
    const text = String(opt.textContent || '').trim();
    if (!text || text.startsWith('--')) return fallback;
    return text;
  }

  function updateOperationalStrip() {
    if (opPeriod) opPeriod.textContent = `Periodo: ${selectedLabel(periodSelect, '-')}`;
    if (opRun) opRun.textContent = `Jornada: ${activeRunId ? `Activa #${activeRunId}` : 'Sin jornada activa'}`;
    if (opService) {
      const serviceText = selectedLabel(serviceSelect, 'TODOS');
      opService.textContent = `Servicio: ${serviceText === '-' ? 'TODOS' : serviceText}`;
    }
    if (opLastScan) opLastScan.textContent = `Ultimo escaneo: ${lastScanDisplay}`;
  }

  function updateActionLocks() {
    const hasPeriod = Boolean(periodSelect?.value);
    const hasRun = Boolean(activeRunId);
    const canExport = hasPeriod;
    if (exportFoundBtn) exportFoundBtn.disabled = !canExport;
    if (exportNotFoundBtn) exportNotFoundBtn.disabled = !canExport;
    if (exportConsolidatedBtn) exportConsolidatedBtn.disabled = !canExport;
    if (exportStateHelp) {
      exportStateHelp.textContent = canExport
        ? 'Exportacion habilitada para el contexto activo.'
        : 'Selecciona un periodo para habilitar exportaciones.';
    }
    renderWorkflowHints();
  }

  function clearGuideHighlight() {
    document.querySelectorAll('.guide-highlight').forEach((el) => el.classList.remove('guide-highlight'));
  }

  function ensureGuidePopover() {
    if (guidePopover) return guidePopover;
    const pop = document.createElement('div');
    pop.className = 'guide-popover hidden';
    pop.innerHTML = `
      <div class="guide-kicker">Guia de uso</div>
      <div id="guideTitle" class="guide-title"></div>
      <div id="guideDesc" class="guide-desc"></div>
      <div class="guide-nav">
        <div id="guideStep" class="guide-step"></div>
        <div class="guide-actions">
          <button id="guidePrev" type="button" class="guide-btn">Anterior</button>
          <button id="guideNext" type="button" class="guide-btn guide-btn-primary">Siguiente</button>
          <button id="guideClose" type="button" class="guide-btn">Cerrar</button>
        </div>
      </div>
    `;
    document.body.appendChild(pop);
    const prev = pop.querySelector('#guidePrev');
    const next = pop.querySelector('#guideNext');
    const close = pop.querySelector('#guideClose');
    prev?.addEventListener('click', () => moveGuide(-1));
    next?.addEventListener('click', () => moveGuide(1));
    close?.addEventListener('click', stopGuide);
    guidePopover = pop;
    return pop;
  }

  function buildGuideSteps() {
    guideSteps = [
      {
        title: 'Contexto y jornada',
        desc: 'Selecciona periodo/servicio y verifica jornada activa. Si falta base, cargala primero desde Jornadas.',
        selector: '.inv-screen .card:nth-of-type(2)',
      },
      {
        title: 'Escaneo operativo',
        desc: 'Escanea con pistola o camara. Cada lectura valida el activo y actualiza estado en tiempo real.',
        selector: '.inv-screen .card:nth-of-type(3)',
      },
      {
        title: 'Validacion de resultados',
        desc: 'Revisa KPI y listas de encontrados/no encontrados/pendientes para controlar avance y brechas.',
        selector: '.inv-screen .card:nth-of-type(4)',
      },
      {
        title: 'Detalle de activos',
        desc: 'Usa paginacion y clasificacion operativa. Encontrado/no encontrado lo controla la jornada.',
        selector: '.inv-screen .card:nth-of-type(5)',
      },
      {
        title: 'Reportes operativos',
        desc: 'Exporta reportes operativos del contexto activo. Para A22 y reportes formales usa Informes.',
        selector: '.inv-screen .card:nth-of-type(6)',
      },
    ];
  }

  function renderGuideStep() {
    const pop = ensureGuidePopover();
    const step = guideSteps[guideIndex];
    if (!step) return;
    clearGuideHighlight();
    const target = document.querySelector(step.selector);
    if (target) {
      target.classList.add('guide-highlight');
      target.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
    const titleEl = pop.querySelector('#guideTitle');
    const descEl = pop.querySelector('#guideDesc');
    const idxEl = pop.querySelector('#guideStep');
    const prevBtn = pop.querySelector('#guidePrev');
    const nextBtn = pop.querySelector('#guideNext');
    if (titleEl) titleEl.textContent = step.title;
    if (descEl) descEl.textContent = step.desc;
    if (idxEl) idxEl.textContent = `Paso ${guideIndex + 1} de ${guideSteps.length}`;
    if (prevBtn) prevBtn.disabled = guideIndex === 0;
    if (nextBtn) nextBtn.textContent = guideIndex >= guideSteps.length - 1 ? 'Finalizar' : 'Siguiente';
  }

  function moveGuide(delta) {
    if (!guideSteps.length) return;
    const next = guideIndex + delta;
    if (next < 0) return;
    if (next >= guideSteps.length) {
      stopGuide();
      return;
    }
    guideIndex = next;
    renderGuideStep();
  }

  function startGuide() {
    buildGuideSteps();
    guideIndex = 0;
    const pop = ensureGuidePopover();
    pop.classList.remove('hidden');
    renderGuideStep();
  }

  function stopGuide() {
    clearGuideHighlight();
    if (guidePopover) guidePopover.classList.add('hidden');
  }

  function renderWorkflowHints() {
    if (!workflowHints) return;
    const hasPeriod = Boolean(periodSelect?.value);
    const hasRun = Boolean(activeRunId);
    const hasScans = lastScanDisplay !== '-';
    const totalAssets = Array.isArray(currentAssets) ? currentAssets.length : 0;
    const foundCount = Number(kpiFoundAssets?.textContent || 0);
    const notFoundCount = Number(kpiNotFoundAssets?.textContent || 0);
    const pendingCount = Number(kpiPendingAssets?.textContent || 0);

    const hints = [];
    hints.push({
      ok: hasPeriod,
      text: hasPeriod
        ? 'Periodo seleccionado. El inventario quedara ordenado por corte.'
        : 'Selecciona un periodo para ordenar la trazabilidad del inventario.',
    });
    hints.push({
      ok: hasRun,
      text: hasRun
        ? 'Jornada activa detectada. Puedes escanear activos.'
        : 'Activa una jornada para habilitar escaneo y registrar encontrados.',
    });
    hints.push({
      ok: hasScans,
      text: hasScans
        ? `Escaneo en marcha. Ultimo registro: ${lastScanDisplay}.`
        : 'Empieza a escanear con pistola o camara para actualizar estados.',
    });
    hints.push({
      ok: totalAssets > 0,
      text: totalAssets > 0
        ? `Base visible cargada (${totalAssets} activos en contexto).`
        : 'No hay activos cargados. Importa la base en Jornadas o ajusta filtros.',
    });
    hints.push({
      ok: pendingCount === 0 && totalAssets > 0,
      text: totalAssets > 0
        ? `Resultado actual: ${foundCount} encontrados, ${notFoundCount} no encontrados, ${pendingCount} pendientes.`
        : 'Sin resultados de conciliacion aun.',
    });

    workflowHints.innerHTML = hints.map((h) => {
      const cls = h.ok ? 'inv-hint inv-hint-ok' : 'inv-hint inv-hint-next';
      const icon = h.ok ? 'OK' : '!';
      return `<div class="${cls}"><span class="inv-hint-badge">${icon}</span><span>${App.escapeHtml(h.text)}</span></div>`;
    }).join('');
  }

  async function refreshActiveRunBanner() {
    if (!periodSelect?.value) {
      renderRunBanner(null);
      return;
    }
    try {
      const params = new URLSearchParams();
      if (periodSelect?.value) params.set('period_id', periodSelect.value);
      params.set('status', 'active');
      const data = await App.get('/runs' + (params.toString() ? `?${params.toString()}` : ''));
      const active = (data.runs || []).find((r) => r.status === 'active') || null;
      renderRunBanner(active);
    } catch (_) {
      if (activeRunBannerText) {
        activeRunBannerText.textContent = 'No se pudo consultar jornada activa.';
        activeRunBannerText.className = 'run-banner-text run-banner-off';
      }
    }
  }

  function computeAssetsHash(items) {
    return (items || []).map((a) => [
      a.id,
      a.C_ACT,
      a.estado_inventario,
      a.estado_jornada || '',
      a.estado_baja || '',
      a.fecha_verificacion || '',
    ].join('|')).join('||');
  }

  async function loadAssets({ preservePage = false } = {}) {
    await refreshActiveRunBanner();
    if (!periodSelect?.value) {
      currentAssets = [];
      if (!preservePage) assetsPage = 1;
      if (assetsContainer) {
        assetsContainer.innerHTML = '<div class="empty-mini">Selecciona un periodo para consultar activos.</div>';
      }
      renderReconciliation([]);
      lastRealtimeHash = '';
      renderWorkflowHints();
      return;
    }
    const url = getAssetsUrl();
    const data = await App.get(url);
    currentAssets = data.assets || [];
    if (!preservePage) assetsPage = 1;
    assetsContainer.innerHTML = renderAssetsTable(currentAssets);
    renderReconciliation(currentAssets);
    lastRealtimeHash = computeAssetsHash(currentAssets);
    renderWorkflowHints();
  }

  async function refreshAssetsSilently() {
    if (realtimeBusy) return;
    if (!periodSelect?.value) return;
    realtimeBusy = true;
    try {
      await refreshActiveRunBanner();
      const data = await App.get(getAssetsUrl());
      const nextItems = data.assets || [];
      const nextHash = computeAssetsHash(nextItems);
      if (nextHash !== lastRealtimeHash) {
        currentAssets = nextItems;
        assetsContainer.innerHTML = renderAssetsTable(currentAssets);
        renderReconciliation(currentAssets);
        lastRealtimeHash = nextHash;
        renderWorkflowHints();
      }
    } catch (_) {
      // Evita ruido visual; el usuario ya ve errores en acciones directas.
    } finally {
      realtimeBusy = false;
    }
  }

  function startRealtimeSync() {
    if (realtimeTimer) clearInterval(realtimeTimer);
    realtimeTimer = setInterval(() => {
      if (document.hidden) return;
      refreshAssetsSilently();
    }, 5000);
  }

  function reconciliationUrl(path) {
    const service = (serviceSelect?.value || '').trim();
    const params = new URLSearchParams();
    if (service) params.set('service', service);
    // Prioriza jornada activa para evitar inconsistencias temporales periodo-jornada.
    if (activeRunId) {
      params.set('run_id', String(activeRunId));
    } else if (periodSelect?.value) {
      params.set('period_id', periodSelect.value);
    }
    if (!params.get('period_id') && !params.get('run_id')) {
      App.setStatus(scanStatus, 'Selecciona un periodo para exportar reportes', true);
      return null;
    }
    return path + (params.toString() ? `?${params.toString()}` : '');
  }

  function ensureConfirmModal() {
    let overlay = document.getElementById('appConfirmOverlay');
    if (overlay) return overlay;
    overlay = document.createElement('div');
    overlay.id = 'appConfirmOverlay';
    overlay.className = 'app-confirm-overlay';
    overlay.hidden = true;
    overlay.innerHTML = `
      <div class="app-confirm-card" role="dialog" aria-modal="true" aria-labelledby="appConfirmTitle">
        <div class="app-confirm-head">
          <h4 id="appConfirmTitle">Confirmacion</h4>
        </div>
        <div id="appConfirmBody" class="app-confirm-body"></div>
        <div class="app-confirm-actions">
          <button id="appConfirmCancel" type="button" class="mini-btn app-confirm-cancel">Cancelar</button>
          <button id="appConfirmOk" type="button" class="mini-btn app-confirm-ok">Confirmar</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    return overlay;
  }

  function ensureLoadingOverlay() {
    if (loadingOverlay) return loadingOverlay;
    loadingOverlay = document.getElementById('appLoadingOverlay');
    if (loadingOverlay) return loadingOverlay;
    loadingOverlay = document.createElement('div');
    loadingOverlay.id = 'appLoadingOverlay';
    loadingOverlay.className = 'app-loading-overlay';
    loadingOverlay.hidden = true;
    loadingOverlay.innerHTML = `
      <div class="app-loading-card" role="status" aria-live="polite" aria-busy="true">
        <div class="app-loading-spinner" aria-hidden="true"></div>
        <div id="appLoadingText" class="app-loading-text">Generando informe...</div>
      </div>
    `;
    document.body.appendChild(loadingOverlay);
    return loadingOverlay;
  }

  function ensureDisposalReasonModal() {
    if (disposalReasonOverlay) return disposalReasonOverlay;
    disposalReasonOverlay = document.getElementById('disposalReasonOverlay');
    if (disposalReasonOverlay) return disposalReasonOverlay;
    disposalReasonOverlay = document.createElement('div');
    disposalReasonOverlay.id = 'disposalReasonOverlay';
    disposalReasonOverlay.className = 'app-confirm-overlay';
    disposalReasonOverlay.hidden = true;
    disposalReasonOverlay.innerHTML = `
      <div class="app-reason-card" role="dialog" aria-modal="true" aria-labelledby="disposalReasonTitle">
        <div class="app-confirm-head">
          <h4 id="disposalReasonTitle">Motivo de baja</h4>
        </div>
        <div id="disposalReasonBody" class="app-confirm-body"></div>
        <textarea id="disposalReasonInput" class="app-reason-input" rows="4" placeholder="Escribe el motivo real de baja..."></textarea>
        <div class="app-reason-note">Este motivo se registrara en la tabla de bajas y en los reportes.</div>
        <div class="app-confirm-actions">
          <button id="disposalReasonCancel" type="button" class="mini-btn app-confirm-cancel">Cancelar</button>
          <button id="disposalReasonSave" type="button" class="mini-btn app-confirm-ok">Guardar motivo</button>
        </div>
      </div>
    `;
    document.body.appendChild(disposalReasonOverlay);
    return disposalReasonOverlay;
  }

  function ensureActionResultModal() {
    if (actionResultOverlay) return actionResultOverlay;
    actionResultOverlay = document.getElementById('actionResultOverlay');
    if (actionResultOverlay) return actionResultOverlay;
    actionResultOverlay = document.createElement('div');
    actionResultOverlay.id = 'actionResultOverlay';
    actionResultOverlay.className = 'app-confirm-overlay';
    actionResultOverlay.hidden = true;
    actionResultOverlay.innerHTML = `
      <div class="app-confirm-card" role="dialog" aria-modal="true" aria-labelledby="actionResultTitle">
        <div class="app-confirm-head">
          <h4 id="actionResultTitle">Actualizacion completada</h4>
        </div>
        <div id="actionResultBody" class="app-confirm-body"></div>
        <div class="app-confirm-actions">
          <button id="actionResultOk" type="button" class="mini-btn app-confirm-ok">Entendido</button>
        </div>
      </div>
    `;
    document.body.appendChild(actionResultOverlay);
    return actionResultOverlay;
  }

  function showActionResultModal({ title = 'Actualizacion completada', message = '' } = {}) {
    return new Promise((resolve) => {
      const overlay = ensureActionResultModal();
      const titleEl = overlay.querySelector('#actionResultTitle');
      const bodyEl = overlay.querySelector('#actionResultBody');
      const okBtn = overlay.querySelector('#actionResultOk');

      const close = () => {
        overlay.hidden = true;
        document.removeEventListener('keydown', onKeydown);
        overlay.removeEventListener('click', onOverlayClick);
        okBtn?.removeEventListener('click', onOk);
        resolve();
      };

      const onOk = () => close();
      const onOverlayClick = (e) => {
        if (e.target === overlay) close();
      };
      const onKeydown = (e) => {
        if (e.key === 'Escape' || e.key === 'Enter') close();
      };

      if (titleEl) titleEl.textContent = title;
      if (bodyEl) bodyEl.textContent = message || '';

      overlay.hidden = false;
      okBtn?.addEventListener('click', onOk);
      overlay.addEventListener('click', onOverlayClick);
      document.addEventListener('keydown', onKeydown);
      setTimeout(() => okBtn?.focus({ preventScroll: true }), 0);
    });
  }

  function classificationDestinationLabel(classification) {
    if (classification === 'Para baja' || classification === 'Baja aprobada') return 'Bajas';
    if (classification === 'En mantenimiento') return 'Mantenimiento';
    if (classification === 'Prestado') return 'Prestamo';
    if (classification === 'Activo de control') return 'Activos de control';
    if (classification === 'Pendiente verificacion') return 'Pendiente de verificacion';
    return classification || 'clasificacion';
  }

  function showDisposalReasonModal({ title = 'Motivo de baja', message = '' } = {}) {
    return new Promise((resolve) => {
      const overlay = ensureDisposalReasonModal();
      const titleEl = overlay.querySelector('#disposalReasonTitle');
      const bodyEl = overlay.querySelector('#disposalReasonBody');
      const inputEl = overlay.querySelector('#disposalReasonInput');
      const saveBtn = overlay.querySelector('#disposalReasonSave');
      const cancelBtn = overlay.querySelector('#disposalReasonCancel');

      const close = (result) => {
        overlay.hidden = true;
        document.removeEventListener('keydown', onKeydown);
        overlay.removeEventListener('click', onOverlayClick);
        saveBtn?.removeEventListener('click', onSave);
        cancelBtn?.removeEventListener('click', onCancel);
        resolve(result);
      };

      const onSave = () => {
        const value = String(inputEl?.value || '').trim();
        if (!value) {
          if (inputEl) inputEl.focus({ preventScroll: true });
          return;
        }
        close(value);
      };
      const onCancel = () => close(null);
      const onOverlayClick = (e) => {
        if (e.target === overlay) close(null);
      };
      const onKeydown = (e) => {
        if (e.key === 'Escape') close(null);
      };

      if (titleEl) titleEl.textContent = title;
      if (bodyEl) bodyEl.textContent = message || '';
      if (inputEl) inputEl.value = '';

      overlay.hidden = false;
      saveBtn?.addEventListener('click', onSave);
      cancelBtn?.addEventListener('click', onCancel);
      overlay.addEventListener('click', onOverlayClick);
      document.addEventListener('keydown', onKeydown);
      setTimeout(() => inputEl?.focus({ preventScroll: true }), 0);
    });
  }

  function showLoadingOverlay(message) {
    const overlay = ensureLoadingOverlay();
    const textEl = overlay.querySelector('#appLoadingText');
    if (textEl) textEl.textContent = message || 'Generando informe...';
    overlay.hidden = false;
  }

  function hideLoadingOverlay() {
    const overlay = ensureLoadingOverlay();
    overlay.hidden = true;
  }

  function showConfirmModal({ title = 'Confirmacion', message = '', okText = 'Confirmar', cancelText = 'Cancelar' } = {}) {
    return new Promise((resolve) => {
      const overlay = ensureConfirmModal();
      const titleEl = overlay.querySelector('#appConfirmTitle');
      const bodyEl = overlay.querySelector('#appConfirmBody');
      const okBtn = overlay.querySelector('#appConfirmOk');
      const cancelBtn = overlay.querySelector('#appConfirmCancel');

      const close = (result) => {
        overlay.hidden = true;
        document.removeEventListener('keydown', onKeydown);
        overlay.removeEventListener('click', onOverlayClick);
        okBtn?.removeEventListener('click', onOk);
        cancelBtn?.removeEventListener('click', onCancel);
        resolve(result);
      };

      const onOk = () => close(true);
      const onCancel = () => close(false);
      const onOverlayClick = (e) => {
        if (e.target === overlay) close(false);
      };
      const onKeydown = (e) => {
        if (e.key === 'Escape') close(false);
      };

      if (titleEl) titleEl.textContent = title;
      if (bodyEl) {
        bodyEl.innerHTML = String(message || '')
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/\n/g, '<br>');
      }
      if (okBtn) okBtn.textContent = okText;
      if (cancelBtn) {
        const hasCancel = String(cancelText || '').trim().length > 0;
        cancelBtn.textContent = hasCancel ? cancelText : 'Cancelar';
        cancelBtn.hidden = !hasCancel;
      }

      overlay.hidden = false;
      okBtn?.addEventListener('click', onOk);
      cancelBtn?.addEventListener('click', onCancel);
      overlay.addEventListener('click', onOverlayClick);
      document.addEventListener('keydown', onKeydown);
    });
  }

  async function postScanWithResponse(payload) {
    const res = await fetch('/scan', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload || {}),
    });
    const data = await res.json().catch(() => ({}));
    return { ok: res.ok, status: res.status, data };
  }

  async function sendScan(code) {
    if (!activeRunId) {
      App.setStatus(scanStatus, 'Debes iniciar una jornada activa para escanear', true);
      return;
    }
    const now = Date.now();
    if (code === lastScanCode && now - lastScanAt < 1500) return;
    lastScanCode = code;
    lastScanAt = now;
    const payload = { code, user: 'usuario_movil' };
    payload.run_id = activeRunId;
    const scanResult = await postScanWithResponse(payload);
    let data = scanResult.data || {};
    let movedFromMismatch = false;

    if (!scanResult.ok && data.code === 'SERVICE_MISMATCH') {
      const expected = String(data.expected_service || '').trim() || 'SIN SERVICIO';
      const expectedLabel = String(data.expected_service_label || expected).trim() || expected;
      const current = String(data.current_service || '').trim() || 'SIN SERVICIO';
      const assetCode = String(data?.asset?.C_ACT || code || '');
      const assetName = String(data?.asset?.NOM || '').trim();
      const assetLabel = assetName ? `${assetName} (${assetCode})` : assetCode;
      const shouldMove = await showConfirmModal({
        title: 'Activo en servicio diferente',
        message:
        `Activo ${assetLabel} escaneado en otro servicio.\n` +
        `Base: ${current}\n` +
        `Alcance jornada: ${expectedLabel}\n\n` +
        'Quieres cambiar el servicio de este activo al de la jornada y marcarlo como encontrado?',
        okText: 'Mover y marcar encontrado',
        cancelText: 'Mantener servicio actual',
      });
      if (!shouldMove) {
        App.setStatus(scanStatus, `Activo ${assetLabel} no movido. Servicio base actual: ${current}.`, true);
        lastScanDisplay = `${assetLabel} (Servicio distinto)`;
        updateOperationalStrip();
        renderWorkflowHints();
        return;
      }

      const assetId = Number(data?.asset?.id || 0);
      if (!assetId) {
        throw new Error('No se pudo identificar el activo para actualizar servicio');
      }
      await App.patch(`/assets/${assetId}/service`, {
        service: expected,
        run_id: activeRunId,
        user: 'usuario_movil',
      });

      const retry = await postScanWithResponse(payload);
      if (!retry.ok) {
        throw new Error(retry.data?.error || `Error HTTP ${retry.status}`);
      }
      data = retry.data || {};
      movedFromMismatch = true;
      App.setStatus(scanStatus, `Activo ${assetLabel} movido a ${expected} y marcado como encontrado.`);
    } else if (!scanResult.ok) {
      throw new Error(data.error || `Error HTTP ${scanResult.status}`);
    }

    if (data.found) {
      if (!movedFromMismatch) {
        App.setStatus(scanStatus, `Encontrado: ${data.asset.C_ACT}`);
      }
      lastScanDisplay = `${data.asset.C_ACT} (Encontrado)`;
    } else {
      App.setStatus(scanStatus, `Codigo no encontrado: ${code}`, true);
      lastScanDisplay = `${code} (No existe en base)`;
    }
    updateOperationalStrip();
    renderWorkflowHints();
  }

  refreshBtn?.addEventListener('click', async () => {
    try {
      await App.loadServices(serviceSelect);
      await loadAssets();
    } catch (err) {
      App.setStatus(scanStatus, err.message, true);
    }
  });

  exportFoundBtn?.addEventListener('click', () => {
    downloadReconciliationReport('/reconciliation/export_found', 'Base depurada (encontrados)');
  });
  exportNotFoundBtn?.addEventListener('click', () => {
    downloadReconciliationReport('/reconciliation/export_not_found', 'Listado no encontrados');
  });
  exportConsolidatedBtn?.addEventListener('click', () => {
    downloadReconciliationReport('/reconciliation/export_consolidated', 'Consolidado final');
  });

  serviceSelect?.addEventListener('change', () => {
    updateOperationalStrip();
    renderWorkflowHints();
    loadAssets().catch((e) => App.setStatus(scanStatus, e.message, true));
  });
  periodSelect?.addEventListener('change', () => {
    updateOperationalStrip();
    updateActionLocks();
    renderWorkflowHints();
    loadAssets().catch((e) => App.setStatus(scanStatus, e.message, true));
  });

  let scanInputTimer = null;
  let scanInputFirstAt = 0;
  let scanInputLastAt = 0;

  async function flushScanInput() {
    if (!scanInput) return;
    const code = scanInput.value.trim();
    scanInput.value = '';
    if (!code) return;
    try {
      await sendScan(code);
      await loadAssets();
    } catch (err) {
      App.setStatus(scanStatus, err.message, true);
    }
  }

  scanInput?.addEventListener('keydown', async (e) => {
    if (e.key === 'Enter' || e.key === 'NumpadEnter' || e.key === 'Tab') {
      e.preventDefault();
      if (scanInputTimer) {
        clearTimeout(scanInputTimer);
        scanInputTimer = null;
      }
      await flushScanInput();
    }
  });

  scanInput?.addEventListener('input', (e) => {
    const now = Date.now();
    if (!scanInputFirstAt) scanInputFirstAt = now;
    scanInputLastAt = now;
    if (scanInputTimer) clearTimeout(scanInputTimer);
    // Si no llega Enter, auto-procesa solo cuando parece lectura en rafaga (pistola)
    // o cuando el usuario pega un codigo.
    scanInputTimer = setTimeout(() => {
      const code = (scanInput?.value || '').trim();
      const burstMs = scanInputFirstAt ? (scanInputLastAt - scanInputFirstAt) : 9999;
      const looksLikeScannerBurst = code.length >= 4 && burstMs >= 0 && burstMs <= 450;
      const fromPaste = String(e?.inputType || '') === 'insertFromPaste';
      if (looksLikeScannerBurst || fromPaste) {
        flushScanInput();
      }
      scanInputFirstAt = 0;
      scanInputLastAt = 0;
      scanInputTimer = null;
    }, 120);
  });

  assetsContainer?.addEventListener('change', async (e) => {
    const pageSizeSelect = e.target.closest('select[data-page-size-select]');
    if (pageSizeSelect) {
      assetsPageSize = Number(pageSizeSelect.value) || 10;
      assetsPage = 1;
      assetsContainer.innerHTML = renderAssetsTable(currentAssets);
      return;
    }
    const select = e.target.closest('select[data-asset-id]');
    if (!select) return;
    const assetId = Number(select.dataset.assetId);
    const classification = select.value;
    if (!assetId || !classification) return;
    const previousValue = String(select.dataset.prevValue || '').trim();
    let disposalReason = '';
    if (classification === 'Para baja' || classification === 'Baja aprobada') {
      const promptTitle = classification === 'Para baja' ? 'para baja' : 'baja aprobada';
      const entered = await showDisposalReasonModal({
        title: `Motivo real de baja (${promptTitle})`,
        message: 'Registra el motivo real que justifica este cambio de estado.',
      });
      if (entered === null) {
        if (previousValue) select.value = previousValue;
        return;
      }
      disposalReason = String(entered || '').trim();
      if (!disposalReason) {
        App.setStatus(scanStatus, 'Debes escribir el motivo real de baja para continuar.', true);
        if (previousValue) select.value = previousValue;
        return;
      }
    }
    try {
      const payload = {
        classification,
        disposal_reason: disposalReason,
        user: 'usuario_movil',
      };
      if (activeRunId) payload.run_id = activeRunId;
      if (periodSelect?.value) payload.period_id = Number(periodSelect.value);
      await App.patch(`/assets/${assetId}/classification`, payload);
      const row = currentAssets.find((item) => Number(item.id) === assetId) || {};
      const assetCode = String(row.C_ACT || assetId);
      const destination = classificationDestinationLabel(classification);
      App.setStatus(scanStatus, `Activo ${assetCode} actualizado: ${classification}`);
      await loadAssets({ preservePage: true });
      await showActionResultModal({
        title: 'Activo actualizado correctamente',
        message: `Activo ${assetCode} enviado correctamente a ${destination}.`,
      });
    } catch (err) {
      if (previousValue) select.value = previousValue;
      App.setStatus(scanStatus, err.message, true);
    }
  });

  assetsContainer?.addEventListener('focusin', (e) => {
    const select = e.target.closest('select[data-asset-id]');
    if (!select) return;
    select.dataset.prevValue = select.value || '';
  });

  assetsContainer?.addEventListener('input', (e) => {
    const input = e.target.closest('input[data-assets-search]');
    if (!input) return;
    assetsSearchTerm = input.value || '';
    const caret = Number.isInteger(input.selectionStart) ? input.selectionStart : assetsSearchTerm.length;
    assetsPage = 1;
    assetsContainer.innerHTML = renderAssetsTable(currentAssets);
    const nextInput = assetsContainer.querySelector('input[data-assets-search]');
    if (nextInput) {
      nextInput.focus({ preventScroll: true });
      const pos = Math.min(caret, String(nextInput.value || '').length);
      nextInput.setSelectionRange(pos, pos);
    }
  });

  assetsContainer?.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-page-action]');
    if (!btn) return;
    const action = btn.dataset.pageAction;
    const totalPages = Math.max(1, Math.ceil(currentAssets.length / assetsPageSize));
    if (action === 'prev' && assetsPage > 1) assetsPage -= 1;
    if (action === 'next' && assetsPage < totalPages) assetsPage += 1;
    assetsContainer.innerHTML = renderAssetsTable(currentAssets);
  });

  document.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-mini-action][data-mini-key]');
    if (!btn) return;
    const action = btn.dataset.miniAction;
    const key = btn.dataset.miniKey;
    if (!key) return;
    if (action === 'prev' && miniPages[key] > 1) miniPages[key] -= 1;
    if (action === 'next') miniPages[key] += 1;
    renderReconciliation(currentAssets);
  });

  document.addEventListener('change', (e) => {
    const select = e.target.closest('select[data-mini-page-size]');
    if (!select) return;
    const key = select.dataset.miniPageSize;
    if (!key) return;
    miniPageSizes[key] = Number(select.value) || 10;
    miniPages[key] = 1;
    renderReconciliation(currentAssets);
  });

  startCamera?.addEventListener('click', async () => {
    await refreshActiveRunBanner();
    if (!activeRunId) {
      App.setStatus(scanStatus, 'Debes iniciar una jornada activa para usar la camara', true);
      return;
    }
    if (typeof window.Html5Qrcode === 'undefined') {
      App.setStatus(scanStatus, 'No se pudo cargar el lector de camara. Recarga la pagina (Ctrl+F5).', true);
      return;
    }
    const host = window.location.hostname;
    const secureAllowed = window.isSecureContext || host === 'localhost' || host === '127.0.0.1';
    if (!secureAllowed) {
      App.setStatus(scanStatus, 'La camara requiere HTTPS (o localhost). Abre la app con URL segura.', true);
      return;
    }
    if (scanner) return;
    scanner = new Html5Qrcode('reader');
    if (stopCamera) stopCamera.disabled = false;
    const hasFormatsApi = typeof window.Html5QrcodeSupportedFormats !== 'undefined';
    const formatsToSupport = hasFormatsApi ? [
      Html5QrcodeSupportedFormats.CODE_128,
      Html5QrcodeSupportedFormats.CODE_39,
      Html5QrcodeSupportedFormats.CODE_93,
      Html5QrcodeSupportedFormats.EAN_13,
      Html5QrcodeSupportedFormats.EAN_8,
      Html5QrcodeSupportedFormats.UPC_A,
      Html5QrcodeSupportedFormats.UPC_E,
      Html5QrcodeSupportedFormats.ITF,
      Html5QrcodeSupportedFormats.CODABAR,
      Html5QrcodeSupportedFormats.QR_CODE,
    ] : undefined;
    const scanConfig = {
      fps: 12,
      qrbox: { width: 320, height: 140 },
      ...(formatsToSupport ? { formatsToSupport } : {}),
    };
    scanner.start({ facingMode: 'environment' }, scanConfig, async (decoded) => {
      try {
        await sendScan(decoded);
        await loadAssets();
      } catch (err) {
        App.setStatus(scanStatus, err.message, true);
      }
    }, () => {}).catch((err) => App.setStatus(scanStatus, `Error camara: ${err}`, true));
  });

  stopCamera?.addEventListener('click', () => {
    if (!scanner) return;
    scanner.stop().then(() => scanner.clear()).finally(() => {
      scanner = null;
      if (stopCamera) stopCamera.disabled = true;
    });
  });

  Promise.all([
    App.loadServices(serviceSelect),
    App.loadPeriods(periodSelect, { status: 'open' }),
  ])
    .then(() => loadAssets())
    .then(() => {
      updateOperationalStrip();
      updateActionLocks();
      renderWorkflowHints();
    })
    .then(() => startRealtimeSync())
    .catch((err) => App.setStatus(scanStatus, err.message, true));

  startGuideBtn?.addEventListener('click', startGuide);

  function setExportLoadingState(loading, label = 'Generando informe') {
    exportBusy = loading;
    [exportFoundBtn, exportNotFoundBtn, exportConsolidatedBtn].forEach((btn) => {
      if (btn) btn.disabled = loading;
    });
    if (scanStatus) scanStatus.classList.toggle('status-loading', loading);
    if (loading) {
      if (scanStatus) App.setStatus(scanStatus, `${label}...`);
      showLoadingOverlay(`${label}...`);
    } else {
      hideLoadingOverlay();
    }
  }

  async function downloadReconciliationReport(path, label) {
    if (exportBusy) return;
    const url = reconciliationUrl(path);
    if (!url) return;
    try {
      const startedAt = Date.now();
      setExportLoadingState(true, `Generando informe: ${label}`);
      const res = await fetch(url);
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || `Error HTTP ${res.status}`);
      }
      const blob = await res.blob();
      const disposition = res.headers.get('Content-Disposition') || '';
      const match = disposition.match(/filename=\"?([^\";]+)\"?/i);
      const downloadName = (match && match[1]) || `${label.replace(/\s+/g, '_')}.xlsx`;
      const objectUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = objectUrl;
      link.download = downloadName;
      document.body.appendChild(link);
      link.click();
      link.remove();
      // Evita cortar la descarga en algunos navegadores.
      setTimeout(() => URL.revokeObjectURL(objectUrl), 60000);

      const elapsed = Date.now() - startedAt;
      if (elapsed < 700) {
        await new Promise((resolve) => setTimeout(resolve, 700 - elapsed));
      }
      scanStatus?.classList.remove('status-loading');
      if (scanStatus) App.setStatus(scanStatus, `${label} generado correctamente.`);
    } catch (err) {
      // Fallback directo si falla la descarga por fetch/blob en el navegador.
      try {
        window.location = url;
      } catch (_) {}
      scanStatus?.classList.remove('status-loading');
      if (scanStatus) App.setStatus(scanStatus, err.message, true);
    } finally {
      setExportLoadingState(false);
    }
  }
});

