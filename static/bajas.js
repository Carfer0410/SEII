document.addEventListener('DOMContentLoaded', () => {
  const serviceSelect = document.getElementById('serviceSelect');
  const periodSelect = document.getElementById('disposalPeriodSelect');
  const periodFilter = document.getElementById('disposalPeriodFilter');
  const codeInput = document.getElementById('disposalCode');
  const reasonInput = document.getElementById('disposalReason');
  const createBtn = document.getElementById('createDisposalBtn');
  const statusFilter = document.getElementById('disposalStatusFilter');
  const reclassifyBtn = document.getElementById('reclassifyBtn');
  const exportGeneralExcelBtn = document.getElementById('exportGeneralExcelBtn');
  const exportGeneralPdfBtn = document.getElementById('exportGeneralPdfBtn');
  const exportControlExcelBtn = document.getElementById('exportControlExcelBtn');
  const exportControlPdfBtn = document.getElementById('exportControlPdfBtn');
  const refreshBtn = document.getElementById('refreshDisposalsBtn');
  const statusEl = document.getElementById('disposalStatus');
  const container = document.getElementById('disposalsContainer');
  let reasonOverlay = null;
  let manageOverlay = null;
  let deleteConfirmOverlay = null;
  const disposalsById = new Map();

  const BASE_SECTIONS = [
    { key: 'BIOMEDICO', title: 'Bajas Biomedicos' },
    { key: 'MUEBLE Y ENSER', title: 'Bajas Mueble y Enser' },
    { key: 'INDUSTRIAL', title: 'Bajas Industriales' },
    { key: 'TECNOLOGICO', title: 'Bajas Tecnologicos' },
  ];

  function formatMoney(value) {
    const n = Number(value || 0);
    return new Intl.NumberFormat('es-CO', {
      style: 'currency',
      currency: 'COP',
      maximumFractionDigits: 0,
    }).format(Number.isFinite(n) ? n : 0);
  }

  function normalizeDisposalType(row) {
    const a = row.asset || {};
    const invStatus = String(a.estado_inventario || '').toUpperCase().trim();
    if (invStatus === 'ACTIVO DE CONTROL') return 'CONTROL - TECNOLOGICO';

    const raw = String(a.TIPO_ACTIVO || '').toUpperCase().trim();
    if (raw.includes('CONTROL')) return raw;
    if (raw.includes('BIOMED')) return 'BIOMEDICO';
    if (raw.includes('MUEBLE')) return 'MUEBLE Y ENSER';
    if (raw.includes('INDUSTR')) return 'INDUSTRIAL';
    return 'TECNOLOGICO';
  }

  function controlSubtypeLabel(typeKey) {
    const raw = String(typeKey || '').toUpperCase();
    if (raw.includes('BIOMED')) return 'BIOMEDICO';
    if (raw.includes('MUEBLE')) return 'MUEBLE Y ENSER';
    if (raw.includes('INDUSTR')) return 'INDUSTRIAL';
    if (raw.includes('TECNOLOG')) return 'TECNOLOGICO';
    return 'OTROS';
  }

  function renderDisposalsTable(items) {
    const headers = [
      'COD ACTIVO FIJO',
      'DESCRIPCION',
      'COSTO INICIAL',
      'SALDO POR DEPRECIAR',
      'FECHA ADQUISICION',
      'MOTIVO DE BAJA',
      'ACCION',
    ];
    return `<table class="disposal-table"><thead><tr>${headers.map((h) => `<th>${h}</th>`).join('')}</tr></thead><tbody>${items.map((row) => {
      const a = row.asset || {};
      const fechaRaw = String(a.FECHA_COMPRA || '').trim();
      const fechaOnly = fechaRaw ? fechaRaw.split('T')[0].split(' ')[0] : '';
      return `<tr>
        <td>${App.escapeHtml(a.C_ACT)}</td>
        <td>${App.escapeHtml(a.NOM)}</td>
        <td>${App.escapeHtml(formatMoney(a.COSTO))}</td>
        <td>${App.escapeHtml(formatMoney(a.SALDO))}</td>
        <td>${App.escapeHtml(fechaOnly)}</td>
        <td>${App.escapeHtml(row.reason)}</td>
        <td><button type="button" class="mini-btn disposal-edit-btn" data-manage-disposal-id="${App.escapeHtml(String(row.id || ''))}">Editar</button></td>
      </tr>`;
    }).join('')}</tbody></table>`;
  }

  function ensureReasonModal() {
    if (reasonOverlay) return reasonOverlay;
    reasonOverlay = document.getElementById('disposalReasonEditOverlay');
    if (reasonOverlay) return reasonOverlay;
    reasonOverlay = document.createElement('div');
    reasonOverlay.id = 'disposalReasonEditOverlay';
    reasonOverlay.className = 'app-confirm-overlay';
    reasonOverlay.hidden = true;
    reasonOverlay.innerHTML = `
      <div class="app-reason-card" role="dialog" aria-modal="true" aria-labelledby="disposalReasonEditTitle">
        <div class="app-confirm-head">
          <h4 id="disposalReasonEditTitle">Editar motivo de baja</h4>
        </div>
        <div id="disposalReasonEditBody" class="app-confirm-body"></div>
        <textarea id="disposalReasonEditInput" class="app-reason-input" rows="4" placeholder="Escribe el motivo real de baja..."></textarea>
        <div class="app-reason-note">Este motivo se reflejara en la tabla y en los reportes exportados.</div>
        <div class="app-confirm-actions">
          <button id="disposalReasonEditCancel" type="button" class="mini-btn app-confirm-cancel">Cancelar</button>
          <button id="disposalReasonEditSave" type="button" class="mini-btn app-confirm-ok">Guardar motivo</button>
        </div>
      </div>
    `;
    document.body.appendChild(reasonOverlay);
    return reasonOverlay;
  }

  function showReasonModal({ title = 'Editar motivo de baja', message = '', initialValue = '' } = {}) {
    return new Promise((resolve) => {
      const overlay = ensureReasonModal();
      const titleEl = overlay.querySelector('#disposalReasonEditTitle');
      const bodyEl = overlay.querySelector('#disposalReasonEditBody');
      const inputEl = overlay.querySelector('#disposalReasonEditInput');
      const saveBtn = overlay.querySelector('#disposalReasonEditSave');
      const cancelBtn = overlay.querySelector('#disposalReasonEditCancel');

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
          inputEl?.focus({ preventScroll: true });
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
      if (bodyEl) bodyEl.textContent = message;
      if (inputEl) inputEl.value = initialValue || '';

      overlay.hidden = false;
      saveBtn?.addEventListener('click', onSave);
      cancelBtn?.addEventListener('click', onCancel);
      overlay.addEventListener('click', onOverlayClick);
      document.addEventListener('keydown', onKeydown);
      setTimeout(() => inputEl?.focus({ preventScroll: true }), 0);
    });
  }

  function ensureManageModal() {
    if (manageOverlay) return manageOverlay;
    manageOverlay = document.getElementById('disposalManageOverlay');
    if (manageOverlay) return manageOverlay;
    manageOverlay = document.createElement('div');
    manageOverlay.id = 'disposalManageOverlay';
    manageOverlay.className = 'app-confirm-overlay';
    manageOverlay.hidden = true;
    manageOverlay.innerHTML = `
      <div class="app-confirm-card" role="dialog" aria-modal="true" aria-labelledby="disposalManageTitle">
        <div class="app-confirm-head">
          <h4 id="disposalManageTitle">Editar baja</h4>
        </div>
        <div id="disposalManageBody" class="app-confirm-body"></div>
        <div class="app-confirm-actions">
          <button id="disposalManageEdit" type="button" class="mini-btn app-confirm-ok">Editar motivo</button>
          <button id="disposalManageRemove" type="button" class="mini-btn app-confirm-cancel">Eliminar de bajas</button>
          <button id="disposalManageClose" type="button" class="mini-btn app-confirm-cancel">Cerrar</button>
        </div>
      </div>
    `;
    document.body.appendChild(manageOverlay);
    return manageOverlay;
  }

  function showManageModal({ title = 'Editar baja', message = '' } = {}) {
    return new Promise((resolve) => {
      const overlay = ensureManageModal();
      const titleEl = overlay.querySelector('#disposalManageTitle');
      const bodyEl = overlay.querySelector('#disposalManageBody');
      const editBtn = overlay.querySelector('#disposalManageEdit');
      const removeBtn = overlay.querySelector('#disposalManageRemove');
      const closeBtn = overlay.querySelector('#disposalManageClose');

      const close = (result) => {
        overlay.hidden = true;
        document.removeEventListener('keydown', onKeydown);
        overlay.removeEventListener('click', onOverlayClick);
        editBtn?.removeEventListener('click', onEdit);
        removeBtn?.removeEventListener('click', onRemove);
        closeBtn?.removeEventListener('click', onClose);
        resolve(result);
      };

      const onEdit = () => close('edit_reason');
      const onRemove = () => close('remove_disposal');
      const onClose = () => close(null);
      const onOverlayClick = (e) => {
        if (e.target === overlay) close(null);
      };
      const onKeydown = (e) => {
        if (e.key === 'Escape') close(null);
      };

      if (titleEl) titleEl.textContent = title;
      if (bodyEl) bodyEl.textContent = message;

      overlay.hidden = false;
      editBtn?.addEventListener('click', onEdit);
      removeBtn?.addEventListener('click', onRemove);
      closeBtn?.addEventListener('click', onClose);
      overlay.addEventListener('click', onOverlayClick);
      document.addEventListener('keydown', onKeydown);
      setTimeout(() => editBtn?.focus({ preventScroll: true }), 0);
    });
  }

  function ensureDeleteConfirmModal() {
    if (deleteConfirmOverlay) return deleteConfirmOverlay;
    deleteConfirmOverlay = document.getElementById('disposalDeleteConfirmOverlay');
    if (deleteConfirmOverlay) return deleteConfirmOverlay;
    deleteConfirmOverlay = document.createElement('div');
    deleteConfirmOverlay.id = 'disposalDeleteConfirmOverlay';
    deleteConfirmOverlay.className = 'app-confirm-overlay';
    deleteConfirmOverlay.hidden = true;
    deleteConfirmOverlay.innerHTML = `
      <div class="app-confirm-card" role="dialog" aria-modal="true" aria-labelledby="disposalDeleteConfirmTitle">
        <div class="app-confirm-head">
          <h4 id="disposalDeleteConfirmTitle">Confirmar eliminacion</h4>
        </div>
        <div id="disposalDeleteConfirmBody" class="app-confirm-body"></div>
        <div class="app-confirm-actions">
          <button id="disposalDeleteConfirmCancel" type="button" class="mini-btn app-confirm-cancel">Cancelar</button>
          <button id="disposalDeleteConfirmOk" type="button" class="mini-btn app-confirm-ok">Eliminar</button>
        </div>
      </div>
    `;
    document.body.appendChild(deleteConfirmOverlay);
    return deleteConfirmOverlay;
  }

  function showDeleteConfirmModal(message) {
    return new Promise((resolve) => {
      const overlay = ensureDeleteConfirmModal();
      const bodyEl = overlay.querySelector('#disposalDeleteConfirmBody');
      const okBtn = overlay.querySelector('#disposalDeleteConfirmOk');
      const cancelBtn = overlay.querySelector('#disposalDeleteConfirmCancel');

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

      if (bodyEl) bodyEl.textContent = message || 'Estas seguro de eliminar este activo de bajas?';

      overlay.hidden = false;
      okBtn?.addEventListener('click', onOk);
      cancelBtn?.addEventListener('click', onCancel);
      overlay.addEventListener('click', onOverlayClick);
      document.addEventListener('keydown', onKeydown);
      setTimeout(() => okBtn?.focus({ preventScroll: true }), 0);
    });
  }

  async function deleteDisposal(disposalId) {
    const res = await fetch(`/disposals/${disposalId}`, { method: 'DELETE' });
    const payload = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(payload.error || `Error HTTP ${res.status}`);
    return payload;
  }

  function renderDisposals(items) {
    const groups = BASE_SECTIONS.reduce((acc, section) => {
      acc[section.key] = [];
      return acc;
    }, {});
    const controlGroups = {};

    (items || []).forEach((row) => {
      const bucket = normalizeDisposalType(row);
      if (String(bucket).includes('CONTROL')) {
        if (!controlGroups[bucket]) controlGroups[bucket] = [];
        controlGroups[bucket].push(row);
      } else {
        if (!groups[bucket]) groups[bucket] = [];
        groups[bucket].push(row);
      }
    });

    const controlSections = Object.keys(controlGroups).sort().map((key) => ({
      key,
      title: `Bajas Control - ${controlSubtypeLabel(key)}`,
      isControl: true,
      exactType: key,
    }));
    const sections = BASE_SECTIONS.concat(controlSections);

    return sections.map((section) => {
      const rows = section.isControl ? (controlGroups[section.key] || []) : (groups[section.key] || []);
      const totalCost = rows.reduce((acc, r) => acc + Number((r.asset || {}).COSTO || 0), 0);
      const totalSaldo = rows.reduce((acc, r) => acc + Number((r.asset || {}).SALDO || 0), 0);
      const body = rows.length
        ? renderDisposalsTable(rows)
        : '<div class="empty-mini">Sin activos en este tipo.</div>';
      return `<section class="disposal-section ${section.isControl ? 'control-section' : ''}">
        <div class="section-head">
          <h4>${section.title}</h4>
          <div class="section-tools">
            <span>${rows.length} activos</span>
            <span>Costo: ${App.escapeHtml(formatMoney(totalCost))}</span>
            <span>Saldo: ${App.escapeHtml(formatMoney(totalSaldo))}</span>
            <button type="button" data-export-type="${section.key}" ${section.exactType ? `data-export-type-exact="${section.exactType}"` : ''} data-export-format="excel" class="mini-btn">Excel</button>
            <button type="button" data-export-type="${section.key}" ${section.exactType ? `data-export-type-exact="${section.exactType}"` : ''} data-export-format="pdf" class="mini-btn">PDF</button>
          </div>
        </div>
        ${body}
      </section>`;
    }).join('');
  }

  async function loadDisposals() {
    if (!periodFilter?.value) {
      container.innerHTML = '<div class="empty-mini">Selecciona un periodo para consultar bajas.</div>';
      return;
    }
    const params = [];
    if (periodFilter?.value) params.push(`period_id=${encodeURIComponent(periodFilter.value)}`);
    if (serviceSelect.value) params.push(`service=${encodeURIComponent(serviceSelect.value)}`);
    if (statusFilter.value) params.push(`status=${encodeURIComponent(statusFilter.value)}`);
    const url = '/disposals' + (params.length ? `?${params.join('&')}` : '');
    const data = await App.get(url);
    const rows = data.disposals || [];
    disposalsById.clear();
    rows.forEach((row) => {
      if (row && row.id != null) disposalsById.set(Number(row.id), row);
    });
    container.innerHTML = renderDisposals(rows);
  }

  createBtn?.addEventListener('click', async () => {
    const code = (codeInput.value || '').trim();
    if (!code) return App.setStatus(statusEl, 'Escribe el codigo del activo', true);
    const selectedPeriodId = Number(periodSelect?.value || 0);
    if (!selectedPeriodId) return App.setStatus(statusEl, 'Selecciona el periodo para registrar la baja', true);
    try {
      const data = await App.post('/disposals', {
        code,
        reason: (reasonInput.value || '').trim(),
        requested_by: 'usuario_movil',
        period_id: selectedPeriodId,
      });
      codeInput.value = '';
      reasonInput.value = '';
      await loadDisposals();
      App.setStatus(statusEl, `Activo marcado: ${data.disposal.asset.C_ACT}`);
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  refreshBtn?.addEventListener('click', () => loadDisposals().catch((e) => App.setStatus(statusEl, e.message, true)));
  exportGeneralExcelBtn?.addEventListener('click', () => {
    if (!periodFilter?.value) return App.setStatus(statusEl, 'Selecciona un periodo para exportar.', true);
    const params = new URLSearchParams();
    if (periodFilter?.value) params.set('period_id', periodFilter.value);
    if (serviceSelect.value) params.set('service', serviceSelect.value);
    if (statusFilter.value) params.set('status', statusFilter.value);
    window.location = '/disposals/export_general_excel' + (params.toString() ? `?${params.toString()}` : '');
  });
  exportGeneralPdfBtn?.addEventListener('click', () => {
    if (!periodFilter?.value) return App.setStatus(statusEl, 'Selecciona un periodo para exportar.', true);
    const params = new URLSearchParams();
    if (periodFilter?.value) params.set('period_id', periodFilter.value);
    if (serviceSelect.value) params.set('service', serviceSelect.value);
    if (statusFilter.value) params.set('status', statusFilter.value);
    window.location = '/disposals/export_general_pdf' + (params.toString() ? `?${params.toString()}` : '');
  });
  exportControlExcelBtn?.addEventListener('click', () => {
    if (!periodFilter?.value) return App.setStatus(statusEl, 'Selecciona un periodo para exportar.', true);
    const params = new URLSearchParams();
    if (periodFilter?.value) params.set('period_id', periodFilter.value);
    if (serviceSelect.value) params.set('service', serviceSelect.value);
    if (statusFilter.value) params.set('status', statusFilter.value);
    window.location = '/disposals/export_general_control_excel' + (params.toString() ? `?${params.toString()}` : '');
  });
  exportControlPdfBtn?.addEventListener('click', () => {
    if (!periodFilter?.value) return App.setStatus(statusEl, 'Selecciona un periodo para exportar.', true);
    const params = new URLSearchParams();
    if (periodFilter?.value) params.set('period_id', periodFilter.value);
    if (serviceSelect.value) params.set('service', serviceSelect.value);
    if (statusFilter.value) params.set('status', statusFilter.value);
    window.location = '/disposals/export_general_control_pdf' + (params.toString() ? `?${params.toString()}` : '');
  });
  reclassifyBtn?.addEventListener('click', async () => {
    try {
      App.setStatus(statusEl, 'Reclasificando activos...');
      const payload = {
        service: serviceSelect.value || '',
        only_disposals: true,
      };
      const data = await App.post('/maintenance/reclassify', payload);
      await loadDisposals();
      App.setStatus(
        statusEl,
        `Reclasificacion completada. Procesados: ${data.total_processed}, actualizados: ${data.updated}`,
      );
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  serviceSelect?.addEventListener('change', () => loadDisposals().catch((e) => App.setStatus(statusEl, e.message, true)));
  statusFilter?.addEventListener('change', () => loadDisposals().catch((e) => App.setStatus(statusEl, e.message, true)));
  periodFilter?.addEventListener('change', () => loadDisposals().catch((e) => App.setStatus(statusEl, e.message, true)));
  container?.addEventListener('click', (e) => {
    const editBtn = e.target.closest('button[data-manage-disposal-id]');
    if (editBtn) {
      const disposalId = Number(editBtn.getAttribute('data-manage-disposal-id'));
      if (!disposalId) return;
      const row = disposalsById.get(disposalId);
      const code = String(row?.asset?.C_ACT || '').trim();
      const desc = String(row?.asset?.NOM || '').trim();
      const currentReason = String(row?.reason || '').trim();
      showManageModal({
        title: 'Editar baja',
        message: `Activo: ${code}${desc ? ` - ${desc}` : ''}`,
      }).then(async (action) => {
        if (!action) return;
        if (action === 'edit_reason') {
          const reason = await showReasonModal({
            title: 'Editar motivo de baja',
            message: `Activo: ${code}${desc ? ` - ${desc}` : ''}`,
            initialValue: currentReason,
          });
          if (reason === null) return;
          try {
            await App.patch(`/disposals/${disposalId}`, { reason });
            await loadDisposals();
            App.setStatus(statusEl, `Motivo actualizado para activo ${code || disposalId}.`);
          } catch (err) {
            App.setStatus(statusEl, err.message, true);
          }
          return;
        }
        if (action === 'remove_disposal') {
          try {
            const confirmed = await showDeleteConfirmModal(`Estas seguro de eliminar ${code || 'este activo'} de bajas?`);
            if (!confirmed) return;
            await deleteDisposal(disposalId);
            await loadDisposals();
            App.setStatus(statusEl, `Activo ${code || disposalId} eliminado de bajas y clasificado en Pendiente verificacion.`);
          } catch (err) {
            App.setStatus(statusEl, err.message, true);
          }
        }
      });
      return;
    }

    const btn = e.target.closest('button[data-export-type][data-export-format]');
    if (!btn) return;
    const type = btn.getAttribute('data-export-type') || '';
    const typeExact = btn.getAttribute('data-export-type-exact') || '';
    const format = btn.getAttribute('data-export-format') || 'excel';
    const params = new URLSearchParams();
    if (typeExact) params.set('type_exact', typeExact);
    else if (type) params.set('type', type);
    if (!periodFilter?.value) {
      App.setStatus(statusEl, 'Selecciona un periodo para exportar.', true);
      return;
    }
    if (periodFilter?.value) params.set('period_id', periodFilter.value);
    if (serviceSelect.value) params.set('service', serviceSelect.value);
    if (statusFilter.value) params.set('status', statusFilter.value);
    const endpoint = format === 'pdf' ? '/disposals/export_pdf' : '/disposals/export_excel';
    window.location = endpoint + (params.toString() ? `?${params.toString()}` : '');
  });

  Promise.all([
    App.loadServices(serviceSelect),
    App.loadPeriods(periodSelect),
    App.loadPeriods(periodFilter),
  ])
    .then(() => {
      if (periodSelect && periodFilter && periodSelect.value && !periodFilter.value) {
        periodFilter.value = periodSelect.value;
      }
      return loadDisposals();
    })
    .catch((err) => App.setStatus(statusEl, err.message, true));

  periodSelect?.addEventListener('change', () => {
    if (!periodFilter) return;
    periodFilter.value = periodSelect.value || '';
    loadDisposals().catch((e) => App.setStatus(statusEl, e.message, true));
  });
});
