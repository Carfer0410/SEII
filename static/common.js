window.App = (() => {
  const HEADER_LABELS = {
    id: 'ID',
    C_ACT: 'Cod activo',
    NOM: 'Descripcion',
    MODELO: 'Modelo',
    REF: 'Referencia',
    SERIE: 'Serial',
    NOM_MARCA: 'Marca',
    C_TIAC: 'Cod tipo activo',
    DESC_TIAC: 'Tipo activo',
    DES_SUBTIAC: 'Subtipo activo',
    DEPRECIA: 'Deprecia',
    VIDA_UTIL: 'Vida util',
    TIPO_ACTIVO: 'Tipo activo clasificado',
    DES_UBI: 'Ubicacion',
    NOM_CCOS: 'Servicio',
    NOM_RESP: 'Responsable',
    EST: 'Estado sistema',
    COSTO: 'Costo inicial',
    SALDO: 'Saldo por depreciar',
    FECHA_COMPRA: 'Fecha adquisicion',
    estado_inventario: 'Estado inventario',
    estado_jornada: 'Estado jornada',
    gestionado_jornada: 'Gestionado jornada',
    estado_baja: 'Estado baja',
    fecha_verificacion: 'Fecha verificacion',
    usuario_verificador: 'Verificado por',
    observacion_inventario: 'Observacion inventario',
  };

  function formatHeader(key) {
    return HEADER_LABELS[key] || String(key || '');
  }

  async function parseJsonOrThrow(res) {
    const payload = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(payload.error || `Error HTTP ${res.status}`);
    return payload;
  }

  async function get(url) {
    const res = await fetch(url);
    return parseJsonOrThrow(res);
  }

  async function post(url, body) {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body || {}),
    });
    return parseJsonOrThrow(res);
  }

  async function patch(url, body) {
    const res = await fetch(url, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body || {}),
    });
    return parseJsonOrThrow(res);
  }

  function setStatus(el, msg, isError = false) {
    if (!el) return;
    el.textContent = msg || '';
    el.style.color = isError ? '#b42318' : '#0d7a52';
  }

  function escapeHtml(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function formatDateTime(value) {
    const raw = String(value || '').trim();
    if (!raw) return '';
    if (!raw.includes('T')) return raw;
    const normalized = raw.endsWith('Z') ? raw.replace('Z', '+00:00') : raw;
    const dt = new Date(normalized);
    if (Number.isNaN(dt.getTime())) return raw.replace('T', ' ').slice(0, 16);
    const yyyy = dt.getFullYear();
    const mm = String(dt.getMonth() + 1).padStart(2, '0');
    const dd = String(dt.getDate()).padStart(2, '0');
    const hh = String(dt.getHours()).padStart(2, '0');
    const mi = String(dt.getMinutes()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd} ${hh}:${mi}`;
  }

  function renderTable(items) {
    if (!items || !items.length) return '<div>No hay datos.</div>';
    const headers = Object.keys(items[0]);
    return `<table><thead><tr>${headers.map((h) => `<th>${escapeHtml(formatHeader(h))}</th>`).join('')}</tr></thead><tbody>${items.map((r) => `<tr>${headers.map((h) => `<td>${escapeHtml(formatDateTime(r[h]))}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  }

  async function loadServices(selectEl) {
    if (!selectEl) return;
    const data = await get('/services');
    selectEl.innerHTML = '<option value="">-- Todos los servicios --</option>' + data.services.map((s) => `<option>${escapeHtml(s)}</option>`).join('');
  }

  async function loadRuns(selectEl, filters = {}) {
    if (!selectEl) return;
    const params = new URLSearchParams();
    if (filters.period_id) params.set('period_id', filters.period_id);
    if (filters.status) params.set('status', filters.status);
    const url = '/runs' + (params.toString() ? `?${params.toString()}` : '');
    const data = await get(url);
    selectEl.innerHTML = '<option value="">-- Sin jornada --</option>' + data.runs.map((r) => {
      const status = r.status === 'active' ? 'Activa' : 'Cerrada';
      const svcLabel = r.service_scope_label || r.service || '';
      const svc = svcLabel ? ` [${svcLabel}]` : '';
      const period = r.period_name ? ` (${r.period_name})` : '';
      return `<option value="${r.id}">${r.id} - ${escapeHtml(r.name)}${svc}${period} - ${status}</option>`;
    }).join('');
  }

  async function loadPeriods(selectEl, filters = {}) {
    if (!selectEl) return;
    const params = new URLSearchParams();
    if (filters.status) params.set('status', filters.status);
    const url = '/periods' + (params.toString() ? `?${params.toString()}` : '');
    const data = await get(url);
    selectEl.innerHTML = '<option value="">-- Selecciona periodo --</option>' + (data.periods || []).map((p) => {
      const suffix = p.status === 'open' ? ' (Abierto)' : (p.status === 'cancelled' ? ' (Anulado)' : ' (Cerrado)');
      return `<option value="${p.id}">${escapeHtml(p.name)}${suffix}</option>`;
    }).join('');
  }

  return { get, post, patch, setStatus, escapeHtml, formatDateTime, renderTable, loadServices, loadRuns, loadPeriods, formatHeader };
})();
