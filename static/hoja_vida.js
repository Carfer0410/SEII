document.addEventListener('DOMContentLoaded', () => {
  const codeInput = document.getElementById('lifeCodeInput');
  const searchBtn = document.getElementById('lifeSearchBtn');
  const pdfBtn = document.getElementById('lifePdfBtn');
  const statusEl = document.getElementById('lifeStatus');
  const previewEl = document.getElementById('lifeSheetPreview');

  let currentCode = '';

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

  async function searchAsset() {
    const code = readCode();
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
});
