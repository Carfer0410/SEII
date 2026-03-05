document.addEventListener('DOMContentLoaded', () => {
  const codeInput = document.getElementById('lifeCodeInput');
  const searchBtn = document.getElementById('lifeSearchBtn');
  const pdfBtn = document.getElementById('lifePdfBtn');
  const statusEl = document.getElementById('lifeStatus');
  const previewEl = document.getElementById('lifeSheetPreview');
  const startCameraBtn = document.getElementById('lifeStartCameraBtn');
  const stopCameraBtn = document.getElementById('lifeStopCameraBtn');
  const tabSheetBtn = document.getElementById('lifeTabSheetBtn');
  const tabAssistBtn = document.getElementById('lifeTabAssistBtn');
  const panelSheet = document.getElementById('lifePanelSheet');
  const panelAssist = document.getElementById('lifePanelAssist');

  const quickCategoryInput = document.getElementById('lifeQuickCategory');
  const quickServiceInput = document.getElementById('lifeQuickServiceHint');
  const quickLocationInput = document.getElementById('lifeQuickLocationHint');
  const quickQueryInput = document.getElementById('lifeQuickQueryText');
  const quickTechnicalInput = document.getElementById('lifeQuickTechnicalText');
  const quickLimitInput = document.getElementById('lifeQuickLimit');
  const quickSearchBtn = document.getElementById('lifeQuickSearchBtn');
  const quickStatusEl = document.getElementById('lifeQuickStatus');
  const quickCandidatesEl = document.getElementById('lifeQuickCandidates');

  let currentCode = '';
  let scanner = null;
  let scanBusy = false;
  let lastDecoded = '';
  let lastDecodedAt = 0;
  let currentTab = 'sheet';

  function readCode() {
    return String(codeInput?.value || '').trim();
  }

  function looksLikeBarcode(code) {
    const txt = String(code || '').trim();
    if (!txt) return false;
    const compact = txt.replace(/\s+/g, '');
    if (compact.length >= 7) return true;
    if (/[^0-9]/.test(compact)) return true;
    return false;
  }

  async function setLifeTab(nextTab) {
    const tab = nextTab === 'assist' ? 'assist' : 'sheet';
    currentTab = tab;
    const isSheet = tab === 'sheet';
    tabSheetBtn?.classList.toggle('active', isSheet);
    tabAssistBtn?.classList.toggle('active', !isSheet);
    panelSheet?.classList.toggle('active', isSheet);
    panelAssist?.classList.toggle('active', !isSheet);
    if (tabSheetBtn) tabSheetBtn.setAttribute('aria-selected', isSheet ? 'true' : 'false');
    if (tabAssistBtn) tabAssistBtn.setAttribute('aria-selected', isSheet ? 'false' : 'true');
    if (!isSheet) await stopCamera();
  }

  function row(label, value) {
    return `
      <div class="life-cell life-label">${App.escapeHtml(label)}</div>
      <div class="life-cell life-value">${App.escapeHtml(String(value || '-'))}</div>
    `;
  }

  function money(value) {
    return Number(value || 0).toLocaleString('es-CO', { maximumFractionDigits: 2 });
  }

  function renderPreview(item) {
    if (!previewEl) return;
    previewEl.className = 'life-sheet';
    previewEl.innerHTML = `
      <div class="life-sheet-head">
        <h4>HOJA DE VIDA DE ACTIVOS</h4>
        <div class="life-sheet-meta">
          Generado: ${App.escapeHtml(item.fecha_generacion || '')} | Coincidencia: ${App.escapeHtml(item.matched_by || 'C_ACT')}
        </div>
      </div>
      <div class="life-grid">
        ${row('Codigo', item.codigo)}
        ${row('Codigo inteligente', item.codigo_inteligente)}
        ${row('Descripcion activo fijo', item.descripcion_activo)}
        ${row('Familia', `${item.familia_codigo || ''} - ${item.familia_nombre || ''}`)}
        ${row('Tipo de activo', `${item.tipo_codigo || ''} - ${item.tipo_nombre || ''}`)}
        ${row('Subtipo de activo', `${item.subtipo_codigo || ''} - ${item.subtipo_nombre || ''}`)}
        ${row('Marca', item.marca)}
        ${row('Modelo', item.modelo)}
        ${row('No. serial o referencia', item.serial_referencia)}
        ${row('Color', item.color)}
        ${row('NIT proveedor', item.nit_proveedor)}
        ${row('Descripcion proveedor', item.proveedor)}
        ${row('Fecha incorporacion', item.fecha_incorporacion)}
        ${row('Forma de adquisicion', item.forma_adquisicion)}
        ${row('En garantia', item.en_garantia)}
        ${row('Entidad', item.entidad)}
        ${row('Desde / Hasta garantia', `${item.garantia_desde || ''} / ${item.garantia_hasta || ''}`)}
        ${row('Estado', item.estado)}
        ${row('Condicion', item.condicion)}
        ${row('Metodo depreciacion', item.metodo_deprec)}
        ${row('Costo del activo', money(item.costo_activo))}
        ${row('Saldo', money(item.saldo))}
        ${row('Total activo', money(item.total_activo))}
        ${row('Responsable', item.responsable)}
        ${row('Ubicacion', item.ubicacion)}
        ${row('Centro de costo', item.centro_costo)}
        ${row('Servicio', item.servicio)}
        ${row('Agencia', item.agencia)}
        ${row('Area', item.area)}
        ${row('Observaciones', item.observaciones)}
      </div>
    `;
  }

  function clearPreview() {
    if (!previewEl) return;
    previewEl.className = 'life-sheet-empty';
    previewEl.textContent = 'Ingresa un codigo para consultar la hoja de vida del activo.';
  }

  async function searchAsset(codeOverride = '') {
    const code = String(codeOverride || readCode()).trim();
    if (!code) {
      App.setStatus(statusEl, 'Escribe o escanea un codigo de activo.', true);
      clearPreview();
      currentCode = '';
      if (pdfBtn) pdfBtn.disabled = true;
      return;
    }
    try {
      const allowBarcode = looksLikeBarcode(code) ? '1' : '0';
      const data = await App.get(`/asset_life_sheet?code=${encodeURIComponent(code)}&allow_barcode=${allowBarcode}`);
      renderPreview(data.item || {});
      currentCode = code;
      if (pdfBtn) pdfBtn.disabled = false;
      App.setStatus(statusEl, `Activo consultado: ${data.item?.codigo || code}`);
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
      clearPreview();
      currentCode = '';
      if (pdfBtn) pdfBtn.disabled = true;
    }
  }

  function canUseCamera() {
    if (typeof window.Html5Qrcode === 'undefined') {
      App.setStatus(statusEl, 'No se pudo cargar el lector de camara. Recarga la pagina (Ctrl+F5).', true);
      return false;
    }
    const host = window.location.hostname;
    const secureAllowed = window.isSecureContext || host === 'localhost' || host === '127.0.0.1';
    if (!secureAllowed) {
      App.setStatus(statusEl, 'La camara requiere HTTPS (o localhost). Abre la app con URL segura.', true);
      return false;
    }
    return true;
  }

  async function stopCamera() {
    if (!scanner) return;
    try { await scanner.stop(); } catch (_) {}
    try { await scanner.clear(); } catch (_) {}
    scanner = null;
    if (startCameraBtn) startCameraBtn.disabled = false;
    if (stopCameraBtn) stopCameraBtn.disabled = true;
  }

  async function startCamera() {
    if (!canUseCamera() || scanner) return;
    scanner = new Html5Qrcode('lifeReader');
    if (startCameraBtn) startCameraBtn.disabled = true;
    if (stopCameraBtn) stopCameraBtn.disabled = false;

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

    try {
      await scanner.start(
        { facingMode: 'environment' },
        scanConfig,
        async (decoded) => {
          const code = String(decoded || '').trim();
          if (!code) return;
          const now = Date.now();
          if (scanBusy) return;
          if (code === lastDecoded && now - lastDecodedAt < 1500) return;
          scanBusy = true;
          lastDecoded = code;
          lastDecodedAt = now;
          if (codeInput) codeInput.value = code;
          App.setStatus(statusEl, `Codigo escaneado: ${code}. Consultando activo...`);
          try {
            await searchAsset(code);
          } finally {
            setTimeout(() => { scanBusy = false; }, 350);
          }
        },
        () => {}
      );
      App.setStatus(statusEl, 'Camara activa. Apunta al codigo para consultar la hoja de vida.');
    } catch (err) {
      App.setStatus(statusEl, `Error camara: ${err}`, true);
      await stopCamera();
    }
  }

  function renderQuickCandidates(payload) {
    if (!quickCandidatesEl) return;
    const analysis = payload?.analysis || {};
    const candidates = payload?.candidates || [];
    if (!candidates.length) {
      quickCandidatesEl.innerHTML = '<div class="life-assist-empty">No se encontraron candidatos con estos criterios.</div>';
      return;
    }
    quickCandidatesEl.innerHTML = `
      <div class="life-assist-summary">
        <div><strong>Tipo:</strong> ${App.escapeHtml(analysis.category_label || 'Sin tipo especifico')}</div>
        <div><strong>Pool no encontrados:</strong> ${App.escapeHtml(String(analysis.not_found_pool_size || 0))}</div>
        <div><strong>Candidatos devueltos:</strong> ${App.escapeHtml(String(analysis.returned_candidates || 0))}</div>
      </div>
      <div class="history-table-wrap">
        <table class="report-history-table compact-table life-assist-table">
          <thead>
            <tr>
              <th>#</th><th>CODIGO</th><th>DESCRIPCION</th><th>SERVICIO</th><th>UBICACION</th><th>ESTADO</th><th>SCORE</th><th>ACCION</th>
            </tr>
          </thead>
          <tbody>
            ${candidates.map((row, idx) => {
              const a = row.asset || {};
              return `
                <tr>
                  <td>${idx + 1}</td>
                  <td>${App.escapeHtml(a.codigo || '')}</td>
                  <td class="cell-clip" title="${App.escapeHtml(a.descripcion || '')}">${App.escapeHtml(a.descripcion || '')}</td>
                  <td>${App.escapeHtml(a.servicio || '-')}</td>
                  <td class="cell-clip" title="${App.escapeHtml(a.ubicacion || '')}">${App.escapeHtml(a.ubicacion || '-')}</td>
                  <td>${App.escapeHtml(a.estado_inventario || '-')}</td>
                  <td>${App.escapeHtml(String(row.score || 0))}</td>
                  <td><button type="button" class="mini-btn" data-life-candidate-code="${App.escapeHtml(a.codigo || '')}">Usar</button></td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  async function quickLookup() {
    const body = {
      category_key: String(quickCategoryInput?.value || '').trim(),
      service_hint: String(quickServiceInput?.value || '').trim(),
      location_hint: String(quickLocationInput?.value || '').trim(),
      query_text: String(quickQueryInput?.value || '').trim(),
      technical_text: String(quickTechnicalInput?.value || '').trim(),
      limit: Number(quickLimitInput?.value || 30),
    };
    if (!body.category_key && !body.service_hint && !body.location_hint && !body.query_text && !body.technical_text) {
      App.setStatus(quickStatusEl, 'Indica al menos un criterio para buscar.', true);
      return;
    }
    try {
      App.setStatus(quickStatusEl, 'Buscando candidatos similares en no encontrados...');
      const payload = await App.post('/asset_life_sheet/quick_lookup', body);
      renderQuickCandidates(payload);
      const total = payload?.analysis?.returned_candidates || 0;
      App.setStatus(quickStatusEl, `Busqueda completada. Candidatos: ${total}.`);
    } catch (err) {
      App.setStatus(quickStatusEl, err.message || 'No fue posible ejecutar la busqueda.', true);
      if (quickCandidatesEl) {
        quickCandidatesEl.innerHTML = '<div class="life-assist-empty">No fue posible generar candidatos.</div>';
      }
    }
  }

  searchBtn?.addEventListener('click', () => searchAsset());
  tabSheetBtn?.addEventListener('click', () => { setLifeTab('sheet'); });
  tabAssistBtn?.addEventListener('click', () => { setLifeTab('assist'); });
  codeInput?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === 'NumpadEnter') {
      e.preventDefault();
      searchAsset();
    }
  });
  pdfBtn?.addEventListener('click', () => {
    const code = currentCode || readCode();
    if (!code) {
      App.setStatus(statusEl, 'Primero consulta un activo para generar el PDF.', true);
      return;
    }
    const allowBarcode = looksLikeBarcode(code) ? '1' : '0';
    window.location = `/asset_life_sheet/pdf?code=${encodeURIComponent(code)}&allow_barcode=${allowBarcode}`;
  });
  startCameraBtn?.addEventListener('click', () => startCamera());
  stopCameraBtn?.addEventListener('click', () => stopCamera());
  quickSearchBtn?.addEventListener('click', () => quickLookup());

  quickCandidatesEl?.addEventListener('click', async (e) => {
    const btn = e.target.closest('button[data-life-candidate-code]');
    if (!btn) return;
    const code = String(btn.getAttribute('data-life-candidate-code') || '').trim();
    if (!code) return;
    if (currentTab !== 'sheet') await setLifeTab('sheet');
    if (codeInput) codeInput.value = code;
    await searchAsset(code);
    App.setStatus(quickStatusEl, `Activo seleccionado: ${code}. Hoja de vida cargada.`);
  });

  window.addEventListener('beforeunload', () => { stopCamera(); });
  setLifeTab('sheet');
});
