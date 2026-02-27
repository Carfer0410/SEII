document.addEventListener("DOMContentLoaded", () => {
  const uploadForm = document.getElementById("uploadForm");
  const fileInput = document.getElementById("fileInput");
  const importResult = document.getElementById("importResult");
  const importStatus = document.getElementById("importStatus");
  const importCurrentBase = document.getElementById("importCurrentBase");
  const runName = document.getElementById("runName");
  const periodSelect = document.getElementById("periodSelect");
  const periodName = document.getElementById("periodName");
  const periodType = document.getElementById("periodType");
  const createPeriodBtn = document.getElementById("createPeriodBtn");
  const refreshPeriodsBtn = document.getElementById("refreshPeriodsBtn");
  const periodFilterSelect = document.getElementById("periodFilterSelect");
  const closePeriodBtn = document.getElementById("closePeriodBtn");
  const cancelPeriodBtn = document.getElementById("cancelPeriodBtn");
  const serviceSearch = document.getElementById("serviceSearch");
  const servicePickerList = document.getElementById("servicePickerList");
  const selectedServicesList = document.getElementById("selectedServicesList");
  const selectedServicesCount = document.getElementById(
    "selectedServicesCount",
  );
  const servicePolicyHint = document.getElementById("servicePolicyHint");
  const selectAllServicesBtn = document.getElementById("selectAllServicesBtn");
  const clearServicesBtn = document.getElementById("clearServicesBtn");
  const createRunBtn = document.getElementById("createRunBtn");
  const runSelect = document.getElementById("runSelect");
  const refreshRunsBtn = document.getElementById("refreshRunsBtn");
  const closeRunBtn = document.getElementById("closeRunBtn");
  const cancelRunBtn = document.getElementById("cancelRunBtn");
  const quickA22Btn = document.getElementById("quickA22Btn");
  const quickClearanceBtn = document.getElementById("quickClearanceBtn");
  const runSummary = document.getElementById("runSummary");
  const closedRunsContainer = document.getElementById("closedRunsContainer");
  const statusEl = document.getElementById("scanStatus");
  const closedRunsMeta = document.getElementById("closedRunsMeta");
  const clearanceSection = document.getElementById("clearanceSection");
  const clearanceRunSelect = document.getElementById("clearanceRunSelect");
  const clearanceDate = document.getElementById("clearanceDate");
  const clearanceOutgoing = document.getElementById("clearanceOutgoing");
  const clearanceIncoming = document.getElementById("clearanceIncoming");
  const clearanceIssuedBy = document.getElementById("clearanceIssuedBy");
  const clearanceNotes = document.getElementById("clearanceNotes");
  const validateClearanceBtn = document.getElementById("validateClearanceBtn");
  const generateClearanceBtn = document.getElementById("generateClearanceBtn");
  const clearanceStatus = document.getElementById("clearanceStatus");

  let cachedRuns = [];
  let cachedPeriods = [];
  let closedPage = 1;
  let closedPageSize = 10;
  let allServices = [];
  let selectedServiceSet = new Set();
  let blockedServiceMap = new Map();
  let reasonOverlay = null;
  let successOverlay = null;

  function ensureConfirmModal() {
    let overlay = document.getElementById("appConfirmOverlay");
    if (overlay) return overlay;
    overlay = document.createElement("div");
    overlay.id = "appConfirmOverlay";
    overlay.className = "app-confirm-overlay";
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

  function showConfirmModal({
    title = "Confirmacion",
    message = "",
    okText = "Confirmar",
    cancelText = "Cancelar",
  } = {}) {
    return new Promise((resolve) => {
      const overlay = ensureConfirmModal();
      const titleEl = overlay.querySelector("#appConfirmTitle");
      const bodyEl = overlay.querySelector("#appConfirmBody");
      const okBtn = overlay.querySelector("#appConfirmOk");
      const cancelBtn = overlay.querySelector("#appConfirmCancel");

      const close = (result) => {
        overlay.hidden = true;
        document.removeEventListener("keydown", onKeydown);
        overlay.removeEventListener("click", onOverlayClick);
        okBtn?.removeEventListener("click", onOk);
        cancelBtn?.removeEventListener("click", onCancel);
        resolve(result);
      };

      const onOk = () => close(true);
      const onCancel = () => close(false);
      const onOverlayClick = (e) => {
        if (e.target === overlay) close(false);
      };
      const onKeydown = (e) => {
        if (e.key === "Escape") close(false);
      };

      if (titleEl) titleEl.textContent = title;
      if (bodyEl) {
        bodyEl.innerHTML = String(message || "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/\n/g, "<br>");
      }
      if (okBtn) okBtn.textContent = okText;
      if (cancelBtn) {
        const hasCancel = String(cancelText || "").trim().length > 0;
        cancelBtn.textContent = hasCancel ? cancelText : "Cancelar";
        cancelBtn.hidden = !hasCancel;
      }

      overlay.hidden = false;
      okBtn?.addEventListener("click", onOk);
      cancelBtn?.addEventListener("click", onCancel);
      overlay.addEventListener("click", onOverlayClick);
      document.addEventListener("keydown", onKeydown);
    });
  }

  function ensureReasonModal() {
    if (reasonOverlay) return reasonOverlay;
    reasonOverlay = document.getElementById("jourReasonOverlay");
    if (reasonOverlay) return reasonOverlay;
    reasonOverlay = document.createElement("div");
    reasonOverlay.id = "jourReasonOverlay";
    reasonOverlay.className = "app-confirm-overlay";
    reasonOverlay.hidden = true;
    reasonOverlay.innerHTML = `
      <div class="app-reason-card" role="dialog" aria-modal="true" aria-labelledby="jourReasonTitle">
        <div class="app-confirm-head">
          <h4 id="jourReasonTitle">Motivo</h4>
        </div>
        <div id="jourReasonBody" class="app-confirm-body"></div>
        <textarea id="jourReasonInput" class="app-reason-input" rows="4" placeholder="Describe el motivo..."></textarea>
        <div class="app-confirm-actions">
          <button id="jourReasonCancel" type="button" class="mini-btn app-confirm-cancel">Cancelar</button>
          <button id="jourReasonSave" type="button" class="mini-btn app-confirm-ok">Confirmar</button>
        </div>
      </div>
    `;
    document.body.appendChild(reasonOverlay);
    return reasonOverlay;
  }

  function showReasonModal({ title = "Motivo", message = "" } = {}) {
    return new Promise((resolve) => {
      const overlay = ensureReasonModal();
      const titleEl = overlay.querySelector("#jourReasonTitle");
      const bodyEl = overlay.querySelector("#jourReasonBody");
      const inputEl = overlay.querySelector("#jourReasonInput");
      const saveBtn = overlay.querySelector("#jourReasonSave");
      const cancelBtn = overlay.querySelector("#jourReasonCancel");

      const close = (result) => {
        overlay.hidden = true;
        document.removeEventListener("keydown", onKeydown);
        overlay.removeEventListener("click", onOverlayClick);
        saveBtn?.removeEventListener("click", onSave);
        cancelBtn?.removeEventListener("click", onCancel);
        resolve(result);
      };

      const onSave = () => {
        const value = String(inputEl?.value || "").trim();
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
        if (e.key === "Escape") close(null);
      };

      if (titleEl) titleEl.textContent = title;
      if (bodyEl) bodyEl.textContent = message || "";
      if (inputEl) inputEl.value = "";

      overlay.hidden = false;
      saveBtn?.addEventListener("click", onSave);
      cancelBtn?.addEventListener("click", onCancel);
      overlay.addEventListener("click", onOverlayClick);
      document.addEventListener("keydown", onKeydown);
      setTimeout(() => inputEl?.focus({ preventScroll: true }), 0);
    });
  }

  function ensureSuccessModal() {
    if (successOverlay) return successOverlay;
    successOverlay = document.getElementById("appSuccessOverlay");
    if (successOverlay) return successOverlay;
    successOverlay = document.createElement("div");
    successOverlay.id = "appSuccessOverlay";
    successOverlay.className = "app-success-overlay";
    successOverlay.hidden = true;
    successOverlay.innerHTML = `
      <div class="app-success-card" role="dialog" aria-modal="true" aria-labelledby="appSuccessTitle">
        <div class="app-success-icon" aria-hidden="true">✓</div>
        <div class="app-success-head">
          <h4 id="appSuccessTitle">Base de datos importada correctamente</h4>
        </div>
        <div id="appSuccessBody" class="app-success-body"></div>
        <div class="app-success-actions">
          <button id="appSuccessOk" type="button" class="mini-btn app-success-ok">Aceptar</button>
        </div>
      </div>
    `;
    document.body.appendChild(successOverlay);
    return successOverlay;
  }

  function showSuccessModal({
    title = "Base de datos importada correctamente",
    message = "",
  } = {}) {
    return new Promise((resolve) => {
      const overlay = ensureSuccessModal();
      const titleEl = overlay.querySelector("#appSuccessTitle");
      const bodyEl = overlay.querySelector("#appSuccessBody");
      const okBtn = overlay.querySelector("#appSuccessOk");

      const close = () => {
        overlay.hidden = true;
        document.removeEventListener("keydown", onKeydown);
        overlay.removeEventListener("click", onOverlayClick);
        okBtn?.removeEventListener("click", onOk);
        resolve(true);
      };

      const onOk = () => close();
      const onOverlayClick = (e) => {
        if (e.target === overlay) close();
      };
      const onKeydown = (e) => {
        if (e.key === "Escape" || e.key === "Enter") close();
      };

      if (titleEl) titleEl.textContent = title;
      if (bodyEl) bodyEl.innerHTML = String(message || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\n/g, "<br>");

      overlay.hidden = false;
      okBtn?.addEventListener("click", onOk);
      overlay.addEventListener("click", onOverlayClick);
      document.addEventListener("keydown", onKeydown);
      setTimeout(() => okBtn?.focus({ preventScroll: true }), 0);
    });
  }

  function isBlockingWorkflowMessage(message) {
    const txt = String(message || "").toLowerCase();
    if (!txt) return false;
    return (
      txt.includes("no puedes") ||
      txt.includes("trazabilidad de escaneo") ||
      txt.includes("jornadas activas") ||
      txt.includes("novedades registradas") ||
      txt.includes("bajas asociadas") ||
      txt.includes("sin jornadas registradas") ||
      txt.includes("periodo con jornadas activas")
    );
  }

  async function showBlockingModal(message, title = "Accion no permitida") {
    await showConfirmModal({
      title,
      message: String(message || "No fue posible completar la accion solicitada."),
      okText: "Aceptar",
      cancelText: "",
    });
  }

  function selectedServices() {
    return allServices.filter((svc) => selectedServiceSet.has(svc));
  }

  function normalizeKey(value) {
    return String(value || "").trim().toLowerCase();
  }

  function blockedInfoForService(serviceName) {
    return blockedServiceMap.get(normalizeKey(serviceName)) || null;
  }

  function renderServicePolicyHint() {
    if (!servicePolicyHint) return;
    const pid = periodSelect?.value ? Number(periodSelect.value) : null;
    if (!pid) {
      servicePolicyHint.textContent =
        "Selecciona un periodo para bloquear automaticamente servicios ya cerrados.";
      return;
    }
    const totalBlocked = blockedServiceMap.size;
    if (!totalBlocked) {
      servicePolicyHint.textContent =
        "En este periodo aun no hay servicios cerrados. Puedes seleccionar cualquier servicio.";
      return;
    }
    servicePolicyHint.textContent = `Servicios bloqueados por cierre en este periodo: ${totalBlocked}. Se muestran con la etiqueta "Cerrado".`;
  }

  async function loadImportStatus() {
    if (!importCurrentBase) return;
    try {
      const data = await App.get("/import/status");
      if (!data.has_import) {
        importCurrentBase.innerHTML =
          '<div class="import-current-base-empty">No hay una base general importada actualmente.</div>';
        return;
      }
      const fileName = String(data.file_name || "").trim() || "Base previamente cargada";
      const importedAt = String(data.imported_at_local || "").trim() || "Sin fecha registrada";
      const imported = Number(data.imported || 0);
      const updated = Number(data.updated || 0);
      const total = imported + updated;
      importCurrentBase.innerHTML = `
        <div class="import-current-base-card">
          <div class="import-current-base-title">Base actual importada</div>
          <div class="import-current-base-meta">
            Archivo: <strong>${App.escapeHtml(fileName)}</strong><br/>
            Fecha de importacion: <strong>${App.escapeHtml(importedAt)}</strong><br/>
            Ultimo procesamiento: <strong>${App.escapeHtml(String(total))}</strong> registros
          </div>
        </div>
      `;
    } catch (_) {
      importCurrentBase.innerHTML =
        '<div class="import-current-base-empty">No se pudo consultar el estado actual de la base.</div>';
    }
  }

  function updateSelectedServicesUi() {
    const selected = selectedServices();
    if (selectedServicesCount)
      selectedServicesCount.textContent = String(selected.length);
    if (!selectedServicesList) return;
    if (!selected.length) {
      selectedServicesList.innerHTML =
        '<div class="empty-mini">No has seleccionado servicios.</div>';
      return;
    }
    selectedServicesList.innerHTML = selected
      .map(
        (svc) =>
          `<button type="button" class="jour-service-chip" data-service-remove="${App.escapeHtml(svc)}">
        <span>${App.escapeHtml(svc)}</span>
        <span class="jour-service-chip-x">x</span>
      </button>`,
      )
      .join("");
  }

  function renderServicePicker() {
    if (!servicePickerList) return;
    const term = String(serviceSearch?.value || "")
      .trim()
      .toLowerCase();
    const filtered = allServices.filter(
      (svc) => !term || svc.toLowerCase().includes(term),
    );
    if (!filtered.length) {
      servicePickerList.innerHTML =
        '<div class="empty-mini">No hay servicios para ese filtro.</div>';
      updateSelectedServicesUi();
      return;
    }
    servicePickerList.innerHTML = filtered
      .map((svc) => {
        const blockedInfo = blockedInfoForService(svc);
        const isBlocked = !!blockedInfo;
        const checked = selectedServiceSet.has(svc) ? "checked" : "";
        const disabled = isBlocked ? "disabled" : "";
        const blockedClass = isBlocked ? "is-blocked" : "";
        const blockedMeta = isBlocked
          ? ` title="Servicio ya cerrado en la jornada ${App.escapeHtml(blockedInfo.last_run_name || "")}"`
          : "";
        return `
        <label class="jour-service-item ${blockedClass}"${blockedMeta}>
          <input type="checkbox" data-service-check="${App.escapeHtml(svc)}" ${checked} ${disabled} />
          <span>${App.escapeHtml(svc)}</span>
          ${isBlocked ? '<span class="jour-service-badge">Cerrado</span>' : ""}
        </label>
      `;
      })
      .join("");
    updateSelectedServicesUi();
  }

  async function loadServicePicker() {
    const data = await App.get("/services");
    allServices = (data.services || []).slice();
    const keep = new Set();
    for (const svc of allServices) {
      if (selectedServiceSet.has(svc) && !blockedInfoForService(svc)) keep.add(svc);
    }
    selectedServiceSet = keep;
    renderServicePicker();
  }

  async function loadClosedServicesForCreatePeriod() {
    const pid = periodSelect?.value ? Number(periodSelect.value) : null;
    blockedServiceMap = new Map();
    if (!pid) {
      renderServicePolicyHint();
      renderServicePicker();
      return;
    }
    const data = await App.get(`/periods/${pid}/closed_services`);
    for (const row of data.items || []) {
      const svc = String(row.service || "").trim();
      if (!svc) continue;
      blockedServiceMap.set(normalizeKey(svc), row);
    }

    for (const svc of Array.from(selectedServiceSet)) {
      if (blockedInfoForService(svc)) selectedServiceSet.delete(svc);
    }
    renderServicePolicyHint();
    renderServicePicker();
  }

  function selectedRunId() {
    return runSelect.value ? Number(runSelect.value) : null;
  }

  function selectedClearanceRunId() {
    return clearanceRunSelect?.value ? Number(clearanceRunSelect.value) : null;
  }

  function renderClearanceRuns(preferredRunId = null) {
    if (!clearanceRunSelect) return;
    const closed = cachedRuns.filter(
      (r) => String(r.status || "").toLowerCase() === "closed",
    );
    clearanceRunSelect.innerHTML =
      '<option value="">-- Selecciona jornada cerrada --</option>' +
      closed
        .map((r) => {
          const svc =
            r.service_scope_label || r.service
              ? ` [${App.escapeHtml(r.service_scope_label || r.service)}]`
              : "";
          const metrics = ` E:${Number(r.found || 0)} / NE:${Number(r.not_found || 0)}`;
          return `<option value="${r.id}">${r.id} - ${App.escapeHtml(r.name)}${svc} |${metrics}</option>`;
        })
        .join("");
    if (
      preferredRunId &&
      closed.some((r) => Number(r.id) === Number(preferredRunId))
    ) {
      clearanceRunSelect.value = String(preferredRunId);
    }
  }

  function renderClosedRunsTable(rows) {
    if (!rows.length)
      return '<div class="empty-mini">No hay jornadas en historial.</div>';
    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / closedPageSize));
    if (closedPage > totalPages) closedPage = totalPages;
    if (closedPage < 1) closedPage = 1;
    const start = (closedPage - 1) * closedPageSize;
    const end = Math.min(start + closedPageSize, total);
    const cut = rows.slice(start, end);
    return `
      <div class="table-toolbar mini-toolbar">
        <div class="field-help">Mostrando ${start + 1}-${end} de ${total} jornadas cerradas</div>
        <div class="pagination-row">
          <label class="field-help">Filas:</label>
          <select class="page-size-select" data-closed-page-size>
            ${[5, 10, 15, 20, 30].map((n) => `<option value="${n}" ${n === closedPageSize ? "selected" : ""}>${n}</option>`).join("")}
          </select>
          <button type="button" class="pager-btn" data-closed-action="prev" ${closedPage <= 1 ? "disabled" : ""}>Anterior</button>
          <span class="field-help">Pagina ${closedPage} de ${totalPages}</span>
          <button type="button" class="pager-btn" data-closed-action="next" ${closedPage >= totalPages ? "disabled" : ""}>Siguiente</button>
        </div>
      </div>
      <div class="closed-runs-wrap"><table class="closed-runs-table closed-runs-table-compact"><thead><tr>
      <th>ID</th><th>JORNADA</th><th>SERVICIO</th><th>ESTADO</th><th>MOTIVO</th><th>HALLAZGOS <span title="E=Encontrados, NE=No encontrados">(E/NE)</span></th><th>INICIO</th><th>CIERRE</th>
    </tr></thead><tbody>${cut
      .map(
        (r) => {
          const rawStatus = String(r.status || "").trim().toLowerCase();
          let statusLabel = "Cerrada";
          if (rawStatus === "cancelled" || rawStatus === "anulada")
            statusLabel = "Anulada";
          else if (rawStatus === "deleted" || rawStatus === "eliminada")
            statusLabel = "Eliminada";
          else if (rawStatus === "closed" || rawStatus === "cerrada")
            statusLabel = "Cerrada";
          const reason =
            String(r.cancel_reason || "").trim() ||
            (statusLabel === "Cerrada" ? "Cierre normal" : "Sin motivo registrado");
          return `<tr>
      <td>${App.escapeHtml(r.id)}</td>
      <td>
        <div class="closed-run-title">${App.escapeHtml(r.name)}</div>
        <div class="closed-run-sub">${App.escapeHtml(r.period_name || "-")}</div>
      </td>
      <td>${App.escapeHtml(r.service_scope_label || r.service || "TODOS")}</td>
      <td>${App.escapeHtml(statusLabel)}</td>
      <td class="cell-clip" title="${App.escapeHtml(reason)}">${App.escapeHtml(reason)}</td>
      <td>
        <div class="closed-metrics">
          <span class="closed-metric ok">E: ${App.escapeHtml(r.found || 0)}</span>
          <span class="closed-metric warn">NE: ${App.escapeHtml(r.not_found || 0)}</span>
        </div>
      </td>
      <td>${App.escapeHtml(
        String(r.started_at || "")
          .replace("T", " ")
          .slice(0, 16),
      )}</td>
      <td>${App.escapeHtml(
        String(r.closed_at || "")
          .replace("T", " ")
          .slice(0, 16),
      )}</td>
    </tr>`;
        },
      )
      .join("")}</tbody></table></div>
    `;
  }

  function selectedPeriodId() {
    return periodFilterSelect?.value ? Number(periodFilterSelect.value) : null;
  }

  async function loadPeriodsView() {
    const data = await App.get("/periods");
    cachedPeriods = data.periods || [];
    const options =
      '<option value="">-- Selecciona periodo --</option>' +
      cachedPeriods
        .map(
          (p) => {
            const suffix = p.status === "open" ? " (Abierto)" : (p.status === "cancelled" ? " (Anulado)" : " (Cerrado)");
            return `<option value="${p.id}">${App.escapeHtml(p.name)}${suffix}</option>`;
          },
        )
        .join("");
    if (periodSelect) periodSelect.innerHTML = options;
    if (periodFilterSelect) periodFilterSelect.innerHTML = options;
  }

  async function loadRunsView(preferredRunId = null) {
    const pid = selectedPeriodId();
    if (!pid) {
      cachedRuns = [];
      runSelect.innerHTML = '<option value="">-- Sin jornada activa --</option>';
      if (clearanceRunSelect) {
        clearanceRunSelect.innerHTML = '<option value="">-- Selecciona jornada cerrada --</option>';
      }
      closedRunsContainer.innerHTML = '<div class="empty-mini">Selecciona un periodo para consultar jornadas.</div>';
      if (closedRunsMeta) {
        closedRunsMeta.textContent = 'Sin periodo seleccionado.';
      }
      return;
    }
    const params = new URLSearchParams();
    if (pid) params.set("period_id", String(pid));
    const data = await App.get(
      "/runs" + (params.toString() ? `?${params.toString()}` : ""),
    );
    cachedRuns = data.runs || [];
    const active = cachedRuns.filter((r) => r.status === "active");
    const closed = cachedRuns.filter((r) => r.status !== "active");

    runSelect.innerHTML =
      '<option value="">-- Sin jornada activa --</option>' +
      active
        .map((r) => {
          const svc =
            r.service_scope_label || r.service
              ? ` [${App.escapeHtml(r.service_scope_label || r.service)}]`
              : "";
          const period = r.period_name
            ? ` (${App.escapeHtml(r.period_name)})`
            : "";
          return `<option value="${r.id}">${r.id} - ${App.escapeHtml(r.name)}${svc}${period}</option>`;
        })
        .join("");

    if (preferredRunId && active.some((r) => r.id === Number(preferredRunId))) {
      runSelect.value = String(preferredRunId);
    } else if (active.length === 1) {
      runSelect.value = String(active[0].id);
    }

    renderClearanceRuns(preferredRunId);
    closedRunsContainer.innerHTML = renderClosedRunsTable(closed);
    if (closedRunsMeta) {
      const totalFound = closed.reduce(
        (acc, r) => acc + Number(r.found || 0),
        0,
      );
      const totalNotFound = closed.reduce(
        (acc, r) => acc + Number(r.not_found || 0),
        0,
      );
      closedRunsMeta.textContent = `Cerradas: ${closed.length} | Encontrados acumulados: ${totalFound} | No encontrados acumulados: ${totalNotFound}`;
    }
  }

  async function refreshSummary() {
    const id = selectedRunId();
    if (!id) {
      runSummary.textContent = "Sin jornada activa seleccionada.";
      return;
    }
    const data = await App.get(`/runs/${id}/summary`);
    const s = data.summary;
    runSummary.textContent = `Total: ${s.total} | Encontrados: ${s.found} | No encontrados: ${s.not_found} | Pendientes: ${s.pending}`;
  }

  uploadForm?.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (!fileInput?.files?.length)
      return App.setStatus(importStatus, "Selecciona un archivo", true);
    const fd = new FormData();
    fd.append("file", fileInput.files[0]);
    try {
      App.setStatus(importStatus, "Importando base...");
      const res = await fetch("/import", { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Error importando base");
      const imported = Number(data.imported || 0);
      const updated = Number(data.updated || 0);
      const total = imported + updated;
      if (importResult) {
        importResult.innerHTML = `
          <div class="import-summary-card">
            <div class="import-summary-title">Base general importada exitosamente</div>
            <div class="import-summary-meta">
              Registros procesados: <strong>${total}</strong><br/>
              Activos nuevos: <strong>${imported}</strong><br/>
              Activos actualizados: <strong>${updated}</strong>
            </div>
          </div>
        `;
      }
      await loadServicePicker();
      await loadPeriodsView();
      await loadClosedServicesForCreatePeriod();
      await loadRunsView(selectedRunId());
      await refreshSummary();
      await loadImportStatus();
      App.setStatus(importStatus, "Base importada correctamente");
      App.setStatus(
        statusEl,
        "Base actualizada. Puedes crear/continuar jornadas.",
      );
      await showSuccessModal({
        title: "Base de datos importada correctamente",
        message: `La base general fue importada con exito.\nRegistros procesados: ${total}\nActivos nuevos: ${imported}\nActivos actualizados: ${updated}`,
      });
    } catch (err) {
      App.setStatus(importStatus, err.message, true);
    }
  });

  createRunBtn?.addEventListener("click", async () => {
    const name = (runName.value || "").trim();
    if (!name)
      return App.setStatus(statusEl, "Escribe el nombre de la jornada", true);
    const pid = periodSelect?.value ? Number(periodSelect.value) : null;
    if (!pid)
      return App.setStatus(
        statusEl,
        "Selecciona un periodo para crear la jornada",
        true,
      );
    const services = selectedServices();
    if (!services.length)
      return App.setStatus(
        statusEl,
        "Selecciona al menos un servicio para la jornada",
        true,
      );
    const blockedSelected = services.filter((svc) => blockedInfoForService(svc));
    if (blockedSelected.length)
      return App.setStatus(
        statusEl,
        `No puedes incluir servicios ya cerrados en el periodo: ${blockedSelected.join(", ")}`,
        true,
      );
    try {
      const data = await App.post("/runs", {
        name,
        period_id: pid,
        services,
        created_by: "usuario_movil",
      });
      runName.value = "";
      await loadRunsView(data.run.id);
      await refreshSummary();
      // Modal de éxito reutilizando el estilo de confirmación
      await showConfirmModal({
        title: "Jornada creada",
        message: `La jornada "${data.run.name}" fue creada correctamente.`,
        okText: "Aceptar",
        cancelText: "",
      });
      App.setStatus(statusEl, `Jornada creada: ${data.run.name}`);
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  refreshRunsBtn?.addEventListener("click", async () => {
    try {
      await loadRunsView(selectedRunId());
      await refreshSummary();
    } catch (e) {
      App.setStatus(statusEl, e.message, true);
    }
  });

  runSelect?.addEventListener("change", () =>
    refreshSummary().catch((e) => App.setStatus(statusEl, e.message, true)),
  );
  periodFilterSelect?.addEventListener("change", () => {
    loadRunsView(null)
      .then(refreshSummary)
      .catch((e) => App.setStatus(statusEl, e.message, true));
  });
  periodSelect?.addEventListener("change", () => {
    loadClosedServicesForCreatePeriod().catch((e) =>
      App.setStatus(statusEl, e.message, true),
    );
  });

  createPeriodBtn?.addEventListener("click", async () => {
    const name = (periodName?.value || "").trim();
    const type = (periodType?.value || "semestral").trim();
    if (!name)
      return App.setStatus(statusEl, "Escribe el nombre del periodo", true);
    try {
      await App.post("/periods", { name, period_type: type });
      if (periodName) periodName.value = "";
      await loadPeriodsView();
      await loadClosedServicesForCreatePeriod();
      await loadRunsView(null);
      await refreshSummary();
      await showConfirmModal({
        title: "Periodo creado",
        message: `El periodo "${name}" fue creado correctamente.`,
        okText: "Aceptar",
        cancelText: "",
      });
      App.setStatus(statusEl, "Periodo creado correctamente");
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  refreshPeriodsBtn?.addEventListener("click", async () => {
    try {
      await loadPeriodsView();
      await loadClosedServicesForCreatePeriod();
      await loadRunsView(selectedRunId());
      await refreshSummary();
      App.setStatus(statusEl, "Periodos actualizados");
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  closePeriodBtn?.addEventListener("click", async () => {
    const pid = selectedPeriodId();
    if (!pid) return App.setStatus(statusEl, "Selecciona un periodo", true);
    const selectedPeriod = cachedPeriods.find((p) => Number(p.id) === Number(pid));
    const periodLabel = selectedPeriod?.name || `ID ${pid}`;
    const accepted = await showConfirmModal({
      title: "Cerrar periodo",
      message: `Estas seguro de cerrar el periodo "${periodLabel}"?\n\nEsta accion no se puede deshacer.`,
      okText: "Si, cerrar periodo",
      cancelText: "Cancelar",
    });
    if (!accepted) return;
    try {
      await App.post(`/periods/${pid}/close`, {});
      await loadPeriodsView();
      await loadClosedServicesForCreatePeriod();
      await loadRunsView(null);
      await refreshSummary();
      App.setStatus(statusEl, "Periodo cerrado");
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
      if (isBlockingWorkflowMessage(err.message)) {
        await showBlockingModal(err.message, "No se pudo cerrar el periodo");
      }
    }
  });

  cancelPeriodBtn?.addEventListener("click", async () => {
    const pid = selectedPeriodId();
    if (!pid) return App.setStatus(statusEl, "Selecciona un periodo", true);
    const selectedPeriod = cachedPeriods.find((p) => Number(p.id) === Number(pid));
    if (!selectedPeriod) return App.setStatus(statusEl, "Periodo no encontrado en la lista", true);
    if (selectedPeriod.status === "cancelled") return App.setStatus(statusEl, "El periodo ya esta anulado", true);
    const periodLabel = selectedPeriod?.name || `ID ${pid}`;
    const reason = await showReasonModal({
      title: "Anular periodo",
      message: `Estas seguro de anular el periodo "${periodLabel}"?\nEsta accion es solo para periodos creados por error y no se debe usar si ya hay trazabilidad.`,
    });
    if (reason === null) return;
    try {
      const data = await App.post(`/periods/${pid}/cancel`, {
        reason,
        user: "usuario_movil",
      });
      await loadPeriodsView();
      await loadClosedServicesForCreatePeriod();
      await loadRunsView(null);
      await refreshSummary();
      App.setStatus(
        statusEl,
        `Periodo anulado. Jornadas anuladas por arrastre: ${Number(data.cancelled_runs || 0)}`,
      );
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
      if (isBlockingWorkflowMessage(err.message)) {
        await showBlockingModal(err.message, "No se pudo anular el periodo");
      }
    }
  });

  closeRunBtn?.addEventListener("click", async () => {
    const id = selectedRunId();
    if (!id)
      return App.setStatus(statusEl, "Selecciona una jornada activa", true);
    const accepted = await showConfirmModal({
      title: "Cerrar jornada activa",
      message:
        'Se marcaran como "No encontrado" todos los activos pendientes dentro del alcance de la jornada.\n\nDeseas continuar?',
      okText: "Si, cerrar jornada",
      cancelText: "Cancelar",
    });
    if (!accepted) return;
    try {
      const data = await App.post(`/runs/${id}/close`, {
        user: "usuario_movil",
      });
      await loadClosedServicesForCreatePeriod();
      await loadRunsView(data.run.id);
      await refreshSummary();
      // Modal de éxito reutilizando el estilo de confirmación
      await showConfirmModal({
        title: "Jornada cerrada",
        message: `La jornada "${data.run.name}" fue cerrada correctamente. No encontrados auto: ${data.auto_marked_not_found}`,
        okText: "Aceptar",
        cancelText: "",
      });
      App.setStatus(
        statusEl,
        `Jornada cerrada. No encontrados auto: ${data.auto_marked_not_found}`,
      );
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
      if (isBlockingWorkflowMessage(err.message)) {
        await showBlockingModal(err.message, "No se pudo cerrar la jornada");
      }
    }
  });

  cancelRunBtn?.addEventListener("click", async () => {
    const id = selectedRunId();
    if (!id) return App.setStatus(statusEl, "Selecciona una jornada", true);
    const selectedRun = cachedRuns.find((r) => Number(r.id) === Number(id));
    if (!selectedRun) return App.setStatus(statusEl, "Jornada no encontrada en la lista", true);
    if (selectedRun.status === "cancelled") return App.setStatus(statusEl, "La jornada ya esta anulada", true);
    const runLabel = selectedRun?.name || `ID ${id}`;
    const reason = await showReasonModal({
      title: "Anular jornada",
      message: `Estas seguro de anular la jornada "${runLabel}"?\nSolo debe anularse si fue creada por error y aun no tiene trazabilidad.`,
    });
    if (reason === null) return;
    try {
      await App.post(`/runs/${id}/cancel`, {
        reason,
        user: "usuario_movil",
      });
      await loadRunsView(null);
      await refreshSummary();
      App.setStatus(statusEl, `Jornada anulada correctamente: ${runLabel}`);
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
      if (isBlockingWorkflowMessage(err.message)) {
        await showBlockingModal(err.message, "No se pudo anular la jornada");
      }
    }
  });

  serviceSearch?.addEventListener("input", renderServicePicker);

  selectAllServicesBtn?.addEventListener("click", () => {
    const term = String(serviceSearch?.value || "")
      .trim()
      .toLowerCase();
    allServices.forEach((svc) => {
      if ((!term || svc.toLowerCase().includes(term)) && !blockedInfoForService(svc))
        selectedServiceSet.add(svc);
    });
    renderServicePicker();
  });

  clearServicesBtn?.addEventListener("click", () => {
    selectedServiceSet.clear();
    renderServicePicker();
  });

  servicePickerList?.addEventListener("change", (e) => {
    const chk = e.target.closest("input[data-service-check]");
    if (!chk) return;
    const svc = chk.getAttribute("data-service-check") || "";
    if (!svc) return;
    if (blockedInfoForService(svc)) {
      chk.checked = false;
      return;
    }
    if (chk.checked) selectedServiceSet.add(svc);
    else selectedServiceSet.delete(svc);
    updateSelectedServicesUi();
  });

  selectedServicesList?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-service-remove]");
    if (!btn) return;
    const svc = btn.getAttribute("data-service-remove") || "";
    if (!svc) return;
    selectedServiceSet.delete(svc);
    renderServicePicker();
  });

  closedRunsContainer?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-closed-action]");
    if (!btn) return;
    const action = btn.getAttribute("data-closed-action");
    const closed = cachedRuns.filter((r) => r.status !== "active");
    const totalPages = Math.max(1, Math.ceil(closed.length / closedPageSize));
    if (action === "prev" && closedPage > 1) closedPage -= 1;
    if (action === "next" && closedPage < totalPages) closedPage += 1;
    closedRunsContainer.innerHTML = renderClosedRunsTable(closed);
  });

  closedRunsContainer?.addEventListener("change", (e) => {
    const select = e.target.closest("select[data-closed-page-size]");
    if (!select) return;
    closedPageSize = Number(select.value) || 10;
    closedPage = 1;
    const closed = cachedRuns.filter((r) => r.status !== "active");
    closedRunsContainer.innerHTML = renderClosedRunsTable(closed);
  });

  quickClearanceBtn?.addEventListener("click", () => {
    clearanceSection?.scrollIntoView({ behavior: "smooth", block: "start" });
    App.setStatus(
      statusEl,
      "Seccion Paz y salvo visible. Completa los datos y genera el PDF.",
    );
  });

  quickA22Btn?.addEventListener("click", () => {
    window.location.href = "/inventario";
  });

  validateClearanceBtn?.addEventListener("click", async () => {
    const pid = selectedPeriodId();
    const runId = selectedClearanceRunId();
    if (!pid) return App.setStatus(clearanceStatus, "Selecciona un periodo", true);
    if (!runId)
      return App.setStatus(
        clearanceStatus,
        "Selecciona una jornada cerrada para validar",
        true,
      );
    try {
      const data = await App.get(
        `/paz_y_salvo/validate?period_id=${encodeURIComponent(String(pid))}&run_id=${encodeURIComponent(String(runId))}`,
      );
      App.setStatus(clearanceStatus, data.message || "", !data.allowed);
      if (!data.allowed) {
        await showBlockingModal(
          data.message || "No cumple condiciones",
          "Paz y salvo bloqueado",
        );
      }
    } catch (err) {
      App.setStatus(clearanceStatus, err.message, true);
    }
  });

  generateClearanceBtn?.addEventListener("click", async () => {
    const pid = selectedPeriodId();
    const runId = selectedClearanceRunId();
    if (!pid) return App.setStatus(clearanceStatus, "Selecciona un periodo", true);
    if (!runId)
      return App.setStatus(
        clearanceStatus,
        "Selecciona una jornada cerrada para generar paz y salvo",
        true,
      );

    const outgoing = String(clearanceOutgoing?.value || "").trim();
    const incoming = String(clearanceIncoming?.value || "").trim();
    const issuedBy = String(clearanceIssuedBy?.value || "").trim();
    const reportDate = String(clearanceDate?.value || "").trim();
    const observations = String(clearanceNotes?.value || "").trim();
    if (!outgoing)
      return App.setStatus(clearanceStatus, "Debes indicar responsable saliente", true);
    if (!incoming)
      return App.setStatus(clearanceStatus, "Debes indicar responsable entrante", true);

    try {
      App.setStatus(clearanceStatus, "Generando paz y salvo...");
      const res = await fetch("/paz_y_salvo/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          period_id: pid,
          run_id: runId,
          outgoing_responsible: outgoing,
          incoming_responsible: incoming,
          issued_by: issuedBy,
          report_date: reportDate,
          observations,
        }),
      });
      if (!res.ok) {
        const payload = await res.json().catch(() => ({}));
        throw new Error(payload.error || `Error HTTP ${res.status}`);
      }
      const blob = await res.blob();
      const header = String(res.headers.get("Content-Disposition") || "");
      const match = header.match(/filename=\"?([^\";]+)\"?/i);
      const filename = match?.[1] || "paz_y_salvo.pdf";
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      App.setStatus(
        clearanceStatus,
        "Paz y salvo generado. El PDF fue descargado y guardado en Documentos.",
      );
    } catch (err) {
      App.setStatus(clearanceStatus, err.message, true);
      await showBlockingModal(err.message, "No se pudo generar paz y salvo");
    }
  });

  if (clearanceDate && !clearanceDate.value) {
    clearanceDate.value = new Date().toISOString().slice(0, 10);
  }
  if (clearanceIssuedBy && !clearanceIssuedBy.value) {
    clearanceIssuedBy.value = "Responsable activos fijos";
  }

  Promise.all([loadServicePicker(), loadPeriodsView()])
    .then(() => loadClosedServicesForCreatePeriod())
    .then(() => loadRunsView(null))
    .then(refreshSummary)
    .then(loadImportStatus)
    .catch((err) => App.setStatus(statusEl, err.message, true));
});
