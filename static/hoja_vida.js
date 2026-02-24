document.addEventListener('DOMContentLoaded', () => {
  const codeInput = document.getElementById('lifeCodeInput');
  const searchBtn = document.getElementById('lifeSearchBtn');
  const pdfBtn = document.getElementById('lifePdfBtn');
  const statusEl = document.getElementById('lifeStatus');
  const previewEl = document.getElementById('lifeSheetPreview');
  const startCameraBtn = document.getElementById('lifeStartCameraBtn');
  const stopCameraBtn = document.getElementById('lifeStopCameraBtn');

  let currentCode = '';
  let scanner = null;
  let scanBusy = false;
  let lastDecoded = '';
  let lastDecodedAt = 0;

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

  function row(label, value) {
    return `
      <div class="life-cell life-label">${App.escapeHtml(label)}</div>
      <div class="life-cell life-value">${App.escapeHtml(String(value || '-'))}</div>
    `;
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
        ${row('Costo del activo', Number(item.costo_activo || 0).toLocaleString('es-CO'))}
        ${row('Saldo', Number(item.saldo || 0).toLocaleString('es-CO'))}
        ${row('Total activo', Number(item.total_activo || 0).toLocaleString('es-CO'))}
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
    try {
      await scanner.stop();
    } catch (_) {
      // Ignorar si ya estaba detenida.
    }
    try {
      await scanner.clear();
    } catch (_) {
      // Ignorar limpieza fallida del componente.
    }
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

  searchBtn?.addEventListener('click', () => searchAsset());
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
  window.addEventListener('beforeunload', () => { stopCamera(); });
});
