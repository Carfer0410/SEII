document.addEventListener("DOMContentLoaded", () => {
  const uploadForm = document.getElementById("uploadForm");
  const fileInput = document.getElementById("fileInput");
  const importResult = document.getElementById("importResult");
  const importStatus = document.getElementById("importStatus");
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
  const selectAllServicesBtn = document.getElementById("selectAllServicesBtn");
  const clearServicesBtn = document.getElementById("clearServicesBtn");
  const createRunBtn = document.getElementById("createRunBtn");
  const runSelect = document.getElementById("runSelect");
  const refreshRunsBtn = document.getElementById("refreshRunsBtn");
  const closeRunBtn = document.getElementById("closeRunBtn");
  const cancelRunBtn = document.getElementById("cancelRunBtn");
  const runSummary = document.getElementById("runSummary");
  const closedRunsContainer = document.getElementById("closedRunsContainer");
  const statusEl = document.getElementById("scanStatus");
  const closedRunsMeta = document.getElementById("closedRunsMeta");

  let cachedRuns = [];
  let cachedPeriods = [];
  let closedPage = 1;
  let closedPageSize = 10;
  let allServices = [];
  let selectedServiceSet = new Set();
  let reasonOverlay = null;

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
      if (cancelBtn) cancelBtn.textContent = cancelText;

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

  function selectedServices() {
    return allServices.filter((svc) => selectedServiceSet.has(svc));
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
        const checked = selectedServiceSet.has(svc) ? "checked" : "";
        return `
        <label class="jour-service-item">
          <input type="checkbox" data-service-check="${App.escapeHtml(svc)}" ${checked} />
          <span>${App.escapeHtml(svc)}</span>
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
      if (selectedServiceSet.has(svc)) keep.add(svc);
    }
    selectedServiceSet = keep;
    renderServicePicker();
  }

  function selectedRunId() {
    return runSelect.value ? Number(runSelect.value) : null;
  }

  function renderClosedRunsTable(rows) {
    if (!rows.length)
      return '<div class="empty-mini">No hay jornadas cerradas.</div>';
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
      <div class="closed-runs-wrap"><table class="closed-runs-table"><thead><tr>
      <th>ID</th><th>PERIODO</th><th>NOMBRE</th><th>ESTADO</th><th>SERVICIO</th><th>ENCONTRADOS</th><th>NO ENCONTRADOS</th><th>INICIO</th><th>CIERRE</th>
    </tr></thead><tbody>${cut
      .map(
        (r) => `<tr>
      <td>${App.escapeHtml(r.id)}</td>
      <td>${App.escapeHtml(r.period_name || "-")}</td>
      <td>${App.escapeHtml(r.name)}</td>
      <td>${App.escapeHtml(r.status === "cancelled" ? "Anulada" : "Cerrada")}</td>
      <td>${App.escapeHtml(r.service_scope_label || r.service || "TODOS")}</td>
      <td>${App.escapeHtml(r.found || 0)}</td>
      <td>${App.escapeHtml(r.not_found || 0)}</td>
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
    </tr>`,
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
      if (importResult) {
        const imported = Number(data.imported || 0);
        const updated = Number(data.updated || 0);
        const total = imported + updated;
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
      await loadRunsView(selectedRunId());
      await refreshSummary();
      App.setStatus(importStatus, "Base importada correctamente");
      App.setStatus(
        statusEl,
        "Base actualizada. Puedes crear/continuar jornadas.",
      );
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

  createPeriodBtn?.addEventListener("click", async () => {
    const name = (periodName?.value || "").trim();
    const type = (periodType?.value || "semestral").trim();
    if (!name)
      return App.setStatus(statusEl, "Escribe el nombre del periodo", true);
    try {
      await App.post("/periods", { name, period_type: type });
      if (periodName) periodName.value = "";
      await loadPeriodsView();
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
      await loadRunsView(null);
      await refreshSummary();
      App.setStatus(statusEl, "Periodo cerrado");
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
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
      await loadRunsView(null);
      await refreshSummary();
      App.setStatus(
        statusEl,
        `Periodo anulado. Jornadas anuladas por arrastre: ${Number(data.cancelled_runs || 0)}`,
      );
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
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
      await loadRunsView(null);
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
    }
  });

  serviceSearch?.addEventListener("input", renderServicePicker);

  selectAllServicesBtn?.addEventListener("click", () => {
    const term = String(serviceSearch?.value || "")
      .trim()
      .toLowerCase();
    allServices.forEach((svc) => {
      if (!term || svc.toLowerCase().includes(term))
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

  Promise.all([loadServicePicker(), loadPeriodsView()])
    .then(() => loadRunsView(null))
    .then(refreshSummary)
    .catch((err) => App.setStatus(statusEl, err.message, true));
});
