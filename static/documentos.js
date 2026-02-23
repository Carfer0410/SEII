document.addEventListener('DOMContentLoaded', () => {
  const linkTypeEl = document.getElementById('docLinkType');
  const assetCodeEl = document.getElementById('docAssetCode');
  const assetNameEl = document.getElementById('docAssetName');
  const typeEl = document.getElementById('docType');
  const titleEl = document.getElementById('docTitle');
  const dateEl = document.getElementById('docDate');
  const areaEl = document.getElementById('docAreaService');
  const radicadoEl = document.getElementById('docRadicado');
  const descriptionEl = document.getElementById('docDescription');
  const fileEl = document.getElementById('docFile');
  const saveBtn = document.getElementById('docSaveBtn');
  const clearBtn = document.getElementById('docClearBtn');
  const statusEl = document.getElementById('docsStatus');

  const searchEl = document.getElementById('docsSearch');
  const filterStatusEl = document.getElementById('docsFilterStatus');
  const filterLinkTypeEl = document.getElementById('docsFilterLinkType');
  const filterTypeEl = document.getElementById('docsFilterType');
  const filterAreaEl = document.getElementById('docsFilterArea');
  const refreshBtn = document.getElementById('docsRefreshBtn');
  const tableWrapEl = document.getElementById('docsTableWrap');

  let documentTypes = [];
  let currentItems = [];
  let docsPage = 1;
  let docsPageSize = 10;
  let editOverlay = null;
  let noticeOverlay = null;
  let searchDebounce = null;

  function formatSize(bytes) {
    const n = Number(bytes || 0);
    if (n >= 1024 * 1024) return `${(n / (1024 * 1024)).toFixed(1)} MB`;
    if (n >= 1024) return `${(n / 1024).toFixed(1)} KB`;
    return `${n} B`;
  }

  function setAssetFieldsDisabled(disabled) {
    if (assetCodeEl) assetCodeEl.disabled = disabled;
    if (assetNameEl) assetNameEl.disabled = disabled;
    if (disabled && assetNameEl) assetNameEl.value = '';
  }

  async function loadTypes() {
    const data = await App.get('/documents/types');
    documentTypes = data.types || [];
    const options = documentTypes.map((t) => `<option value="${App.escapeHtml(t)}">${App.escapeHtml(t)}</option>`).join('');
    if (typeEl) typeEl.innerHTML = `<option value="">Selecciona tipo...</option>${options}`;
    if (filterTypeEl) filterTypeEl.innerHTML = `<option value="">Todos los tipos</option>${options}`;
  }

  async function lookupAssetByCode() {
    const code = String(assetCodeEl?.value || '').trim();
    if (!code || String(linkTypeEl?.value || '') !== 'asset') {
      if (assetNameEl) assetNameEl.value = '';
      return;
    }
    try {
      const data = await App.get(`/assets/find_by_code?code=${encodeURIComponent(code)}`);
      if (assetNameEl) assetNameEl.value = data.asset?.name || '';
    } catch (_) {
      if (assetNameEl) assetNameEl.value = '';
    }
  }

  function ensureEditModal() {
    if (editOverlay) return editOverlay;
    editOverlay = document.createElement('div');
    editOverlay.className = 'app-confirm-overlay docs-edit-overlay';
    editOverlay.hidden = true;
    editOverlay.innerHTML = `
      <div class="app-confirm-card docs-edit-card" role="dialog" aria-modal="true" aria-labelledby="docsEditTitle">
        <div class="app-confirm-head">
          <h4 id="docsEditTitle">Editar documento</h4>
        </div>
        <div class="app-confirm-body">
          <div class="docs-edit-grid">
            <label><span>Tipo documento</span><select id="docsEditType"></select></label>
            <label><span>Titulo</span><input id="docsEditTitleInput" /></label>
            <label><span>Fecha doc</span><input id="docsEditDate" type="date" /></label>
            <label><span>Area/Servicio</span><input id="docsEditArea" /></label>
            <label><span>Radicado</span><input id="docsEditRadicado" /></label>
            <label class="docs-edit-col2"><span>Descripcion</span><textarea id="docsEditDescription" rows="3"></textarea></label>
          </div>
        </div>
        <div class="app-confirm-actions">
          <button id="docsEditCancel" type="button" class="mini-btn app-confirm-cancel">Cancelar</button>
          <button id="docsEditSave" type="button" class="mini-btn app-confirm-ok">Guardar cambios</button>
        </div>
      </div>
    `;
    document.body.appendChild(editOverlay);
    return editOverlay;
  }

  function openEditModal(item) {
    return new Promise((resolve) => {
      const overlay = ensureEditModal();
      const typeInput = overlay.querySelector('#docsEditType');
      const titleInput = overlay.querySelector('#docsEditTitleInput');
      const dateInput = overlay.querySelector('#docsEditDate');
      const areaInput = overlay.querySelector('#docsEditArea');
      const radicadoInput = overlay.querySelector('#docsEditRadicado');
      const descriptionInput = overlay.querySelector('#docsEditDescription');
      const cancelBtn = overlay.querySelector('#docsEditCancel');
      const saveBtn = overlay.querySelector('#docsEditSave');

      typeInput.innerHTML = documentTypes.map((t) => `<option value="${App.escapeHtml(t)}">${App.escapeHtml(t)}</option>`).join('');
      typeInput.value = item.document_type || '';
      titleInput.value = item.title || '';
      dateInput.value = item.doc_date || '';
      areaInput.value = item.area_service || '';
      radicadoInput.value = item.radicado || '';
      descriptionInput.value = item.description || '';

      const close = (payload = null) => {
        overlay.hidden = true;
        overlay.removeEventListener('click', onOverlayClick);
        document.removeEventListener('keydown', onKeyDown);
        cancelBtn.removeEventListener('click', onCancel);
        saveBtn.removeEventListener('click', onSave);
        resolve(payload);
      };
      const onCancel = () => close(null);
      const onSave = () => close({
        document_type: String(typeInput.value || '').trim(),
        title: String(titleInput.value || '').trim(),
        doc_date: String(dateInput.value || '').trim(),
        area_service: String(areaInput.value || '').trim(),
        radicado: String(radicadoInput.value || '').trim(),
        description: String(descriptionInput.value || '').trim(),
      });
      const onOverlayClick = (e) => {
        if (e.target === overlay) close(null);
      };
      const onKeyDown = (e) => {
        if (e.key === 'Escape') close(null);
      };

      overlay.hidden = false;
      cancelBtn.addEventListener('click', onCancel);
      saveBtn.addEventListener('click', onSave);
      overlay.addEventListener('click', onOverlayClick);
      document.addEventListener('keydown', onKeyDown);
      setTimeout(() => titleInput.focus({ preventScroll: true }), 0);
    });
  }

  function ensureNoticeModal() {
    if (noticeOverlay) return noticeOverlay;
    noticeOverlay = document.createElement('div');
    noticeOverlay.className = 'app-confirm-overlay docs-notice-overlay';
    noticeOverlay.hidden = true;
    noticeOverlay.innerHTML = `
      <div class="app-confirm-card docs-notice-card" role="dialog" aria-modal="true" aria-labelledby="docsNoticeTitle">
        <div class="app-confirm-head">
          <h4 id="docsNoticeTitle"></h4>
        </div>
        <div class="app-confirm-body">
          <p id="docsNoticeMessage" class="docs-notice-message"></p>
        </div>
        <div class="app-confirm-actions">
          <button id="docsNoticeCancel" type="button" class="mini-btn app-confirm-cancel">Cancelar</button>
          <button id="docsNoticeOk" type="button" class="mini-btn app-confirm-ok">Aceptar</button>
        </div>
      </div>
    `;
    document.body.appendChild(noticeOverlay);
    return noticeOverlay;
  }

  function showNoticeModal({ title, message, okText = 'Aceptar', cancelText = 'Cancelar', showCancel = false }) {
    return new Promise((resolve) => {
      const overlay = ensureNoticeModal();
      const titleEl = overlay.querySelector('#docsNoticeTitle');
      const messageEl = overlay.querySelector('#docsNoticeMessage');
      const cancelBtn = overlay.querySelector('#docsNoticeCancel');
      const okBtn = overlay.querySelector('#docsNoticeOk');

      titleEl.textContent = String(title || '').trim() || 'Confirmacion';
      messageEl.textContent = String(message || '').trim();
      okBtn.textContent = okText;
      cancelBtn.textContent = cancelText;
      cancelBtn.hidden = !showCancel;

      const close = (accepted) => {
        overlay.hidden = true;
        overlay.removeEventListener('click', onOverlayClick);
        document.removeEventListener('keydown', onKeyDown);
        cancelBtn.removeEventListener('click', onCancel);
        okBtn.removeEventListener('click', onOk);
        resolve(accepted);
      };
      const onCancel = () => close(false);
      const onOk = () => close(true);
      const onOverlayClick = (e) => {
        if (e.target === overlay) close(false);
      };
      const onKeyDown = (e) => {
        if (e.key === 'Escape') close(false);
      };

      overlay.hidden = false;
      cancelBtn.addEventListener('click', onCancel);
      okBtn.addEventListener('click', onOk);
      overlay.addEventListener('click', onOverlayClick);
      document.addEventListener('keydown', onKeyDown);
      setTimeout(() => okBtn.focus({ preventScroll: true }), 0);
    });
  }

  function renderTable(items) {
    if (!tableWrapEl) return;
    if (!items.length) {
      tableWrapEl.innerHTML = '<div class="empty-mini">No hay documentos registrados para este filtro.</div>';
      return;
    }
    const total = items.length;
    const totalPages = Math.max(1, Math.ceil(total / docsPageSize));
    if (docsPage > totalPages) docsPage = totalPages;
    if (docsPage < 1) docsPage = 1;
    const start = (docsPage - 1) * docsPageSize;
    const end = Math.min(start + docsPageSize, total);
    const pageItems = items.slice(start, end);
    tableWrapEl.innerHTML = `
      <div class="history-table-wrap">
        <table class="docs-table">
          <thead>
            <tr>
              <th>ID</th>
              <th>Tipo</th>
              <th>Titulo</th>
              <th>Vinculacion</th>
              <th>Activo</th>
              <th>Area</th>
              <th>Fecha doc</th>
              <th>Subido</th>
              <th>Archivo</th>
              <th>Acciones</th>
            </tr>
          </thead>
          <tbody>
            ${pageItems.map((r) => `
              <tr>
                <td>${App.escapeHtml(String(r.id || ''))}</td>
                <td><span class="doc-type-chip">${App.escapeHtml(r.document_type || '')}</span></td>
                <td class="cell-clip" title="${App.escapeHtml(r.title || '')}">${App.escapeHtml(r.title || '')}</td>
                <td>${r.link_type === 'asset' ? 'Con activo' : 'General'}</td>
                <td class="cell-clip" title="${App.escapeHtml(`${r.asset_code || ''} ${r.asset_name || ''}`.trim())}">
                  ${App.escapeHtml(`${r.asset_code || ''} ${r.asset_name || ''}`.trim() || '-')}
                </td>
                <td>${App.escapeHtml(r.area_service || '-')}</td>
                <td>${App.escapeHtml(r.doc_date || '-')}</td>
                <td>${App.escapeHtml(App.formatDateTime(r.uploaded_at_local || r.uploaded_at || ''))}</td>
                <td class="cell-clip" title="${App.escapeHtml(r.file_name || '')}">
                  ${App.escapeHtml(r.file_name || '')}<br/>
                  <span class="file-meta">${App.escapeHtml(String((r.file_ext || '').toUpperCase()))} Â· ${formatSize(r.file_size)}</span>
                </td>
                <td>
                  <button class="mini-btn" type="button" data-doc-download="${r.id}">Descargar</button>
                  <button class="mini-btn" type="button" data-doc-edit="${r.id}">Editar</button>
                  <button class="mini-btn docs-archive-btn" type="button" data-doc-archive="${r.id}">Archivar</button>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
      <div class="history-pagination">
        <span class="field-help">Mostrando ${start + 1}-${end} de ${total}</span>
        <label class="field-help">Filas:</label>
        <select id="docsPageSize" class="page-size-select">
          ${[10, 25, 50, 100].map((n) => `<option value="${n}" ${n === docsPageSize ? 'selected' : ''}>${n}</option>`).join('')}
        </select>
        <button type="button" class="mini-btn" data-doc-page="prev" ${docsPage <= 1 ? 'disabled' : ''}>Anterior</button>
        <span class="field-help">Pagina ${docsPage} de ${totalPages}</span>
        <button type="button" class="mini-btn" data-doc-page="next" ${docsPage >= totalPages ? 'disabled' : ''}>Siguiente</button>
      </div>
    `;
  }

  async function loadDocuments() {
    const params = new URLSearchParams();
    if (searchEl?.value) params.set('search', searchEl.value.trim());
    if (filterStatusEl?.value) params.set('status', filterStatusEl.value);
    if (filterLinkTypeEl?.value) params.set('link_type', filterLinkTypeEl.value);
    if (filterTypeEl?.value) params.set('document_type', filterTypeEl.value);
    if (filterAreaEl?.value) params.set('area_service', filterAreaEl.value.trim());
    const data = await App.get('/documents' + (params.toString() ? `?${params.toString()}` : ''));
    currentItems = data.items || [];
    renderTable(currentItems);
  }

  function clearForm() {
    if (assetCodeEl) assetCodeEl.value = '';
    if (assetNameEl) assetNameEl.value = '';
    if (typeEl) typeEl.value = '';
    if (titleEl) titleEl.value = '';
    if (dateEl) dateEl.value = '';
    if (areaEl) areaEl.value = '';
    if (radicadoEl) radicadoEl.value = '';
    if (descriptionEl) descriptionEl.value = '';
    if (fileEl) fileEl.value = '';
    App.setStatus(statusEl, '');
  }

  async function saveDocument() {
    const fd = new FormData();
    fd.append('link_type', String(linkTypeEl?.value || 'general'));
    fd.append('asset_code', String(assetCodeEl?.value || '').trim());
    fd.append('document_type', String(typeEl?.value || '').trim());
    fd.append('title', String(titleEl?.value || '').trim());
    fd.append('doc_date', String(dateEl?.value || '').trim());
    fd.append('area_service', String(areaEl?.value || '').trim());
    fd.append('radicado', String(radicadoEl?.value || '').trim());
    fd.append('description', String(descriptionEl?.value || '').trim());
    fd.append('uploaded_by', 'usuario_movil');
    const file = fileEl?.files?.[0];
    if (file) fd.append('file', file);

    const res = await fetch('/documents', { method: 'POST', body: fd });
    const payload = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(payload.error || `Error HTTP ${res.status}`);
    return payload;
  }

  linkTypeEl?.addEventListener('change', () => {
    const isAsset = String(linkTypeEl.value || '') === 'asset';
    setAssetFieldsDisabled(!isAsset);
    if (isAsset) {
      assetCodeEl?.focus({ preventScroll: true });
    }
  });

  assetCodeEl?.addEventListener('blur', () => lookupAssetByCode());
  assetCodeEl?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === 'NumpadEnter') {
      e.preventDefault();
      lookupAssetByCode();
    }
  });

  saveBtn?.addEventListener('click', async () => {
    try {
      App.setStatus(statusEl, 'Subiendo documento...');
      await saveDocument();
      clearForm();
      await loadDocuments();
      App.setStatus(statusEl, 'Documento guardado correctamente.');
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  clearBtn?.addEventListener('click', () => clearForm());

  [filterStatusEl, filterLinkTypeEl, filterTypeEl, filterAreaEl].forEach((el) => {
    el?.addEventListener('change', () => {
      docsPage = 1;
      loadDocuments().catch((err) => App.setStatus(statusEl, err.message, true));
    });
  });
  searchEl?.addEventListener('input', () => {
    if (searchDebounce) clearTimeout(searchDebounce);
    searchDebounce = setTimeout(() => {
      docsPage = 1;
      loadDocuments().catch((err) => App.setStatus(statusEl, err.message, true));
    }, 220);
  });
  searchEl?.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    if (searchDebounce) clearTimeout(searchDebounce);
    docsPage = 1;
    loadDocuments().catch((err) => App.setStatus(statusEl, err.message, true));
  });

  refreshBtn?.addEventListener('click', () => {
    loadDocuments()
      .then(() => App.setStatus(statusEl, 'Repositorio actualizado'))
      .catch((err) => App.setStatus(statusEl, err.message, true));
  });

  tableWrapEl?.addEventListener('click', async (e) => {
    const btn = e.target.closest('button[data-doc-download]');
    if (btn) {
      const id = Number(btn.getAttribute('data-doc-download') || 0);
      if (!id) return;
      window.location = `/documents/${id}/download`;
      return;
    }

    const editBtn = e.target.closest('button[data-doc-edit]');
    if (editBtn) {
      const id = Number(editBtn.getAttribute('data-doc-edit') || 0);
      if (!id) return;
      const item = currentItems.find((x) => Number(x.id) === id);
      if (!item) return;
      openEditModal(item)
        .then(async (payload) => {
          if (!payload) return;
          if (!payload.title) {
            App.setStatus(statusEl, 'El titulo no puede quedar vacio.', true);
            return;
          }
          if (!payload.document_type) {
            App.setStatus(statusEl, 'Debes seleccionar tipo de documento.', true);
            return;
          }
          const confirmEdit = await showNoticeModal({
            title: 'Confirmar edicion',
            message: `Vas a guardar cambios del documento "${item.title || `#${id}`}".`,
            okText: 'Guardar',
            cancelText: 'Cancelar',
            showCancel: true,
          });
          if (!confirmEdit) return;
          const res = await fetch(`/documents/${id}`, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
          });
          const body = await res.json().catch(() => ({}));
          if (!res.ok) throw new Error(body.error || `Error HTTP ${res.status}`);
          App.setStatus(statusEl, `Documento ${id} actualizado.`);
          await loadDocuments();
          await showNoticeModal({
            title: 'Edicion completada',
            message: 'El documento se actualizo correctamente.',
            okText: 'Entendido',
          });
        })
        .catch((err) => App.setStatus(statusEl, err.message, true));
      return;
    }

    const pageBtn = e.target.closest('button[data-doc-page]');
    if (pageBtn) {
      const action = pageBtn.getAttribute('data-doc-page');
      if (action === 'prev') docsPage = Math.max(1, docsPage - 1);
      if (action === 'next') docsPage += 1;
      renderTable(currentItems);
      return;
    }

    const sizeSelect = e.target.closest('#docsPageSize');
    if (sizeSelect) {
      docsPageSize = Math.max(10, Number(sizeSelect.value || 10));
      docsPage = 1;
      renderTable(currentItems);
      return;
    }

    const archiveBtn = e.target.closest('button[data-doc-archive]');
    if (archiveBtn) {
      const id = Number(archiveBtn.getAttribute('data-doc-archive') || 0);
      if (!id) return;
      const item = currentItems.find((x) => Number(x.id) === id);
      const label = item?.title || `Documento ${id}`;
      const ok = await showNoticeModal({
        title: 'Archivar documento',
        message: `Seguro que deseas archivar "${label}"?`,
        okText: 'Archivar',
        cancelText: 'Cancelar',
        showCancel: true,
      });
      if (!ok) return;
      fetch(`/documents/${id}/archive`, { method: 'POST' })
        .then(async (res) => {
          const payload = await res.json().catch(() => ({}));
          if (!res.ok) throw new Error(payload.error || `Error HTTP ${res.status}`);
          App.setStatus(statusEl, `Documento ${id} archivado.`);
          docsPage = 1;
          await loadDocuments();
          await showNoticeModal({
            title: 'Documento archivado',
            message: 'El documento fue archivado correctamente.',
            okText: 'Aceptar',
          });
        })
        .catch((err) => App.setStatus(statusEl, err.message, true));
    }
  });

  tableWrapEl?.addEventListener('change', (e) => {
    const sizeSelect = e.target.closest('#docsPageSize');
    if (!sizeSelect) return;
    docsPageSize = Math.max(10, Number(sizeSelect.value || 10));
    docsPage = 1;
    renderTable(currentItems);
  });

  setAssetFieldsDisabled(true);
  Promise.all([loadTypes(), loadDocuments()])
    .catch((err) => App.setStatus(statusEl, err.message, true));
});
