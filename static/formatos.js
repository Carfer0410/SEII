document.addEventListener("DOMContentLoaded", () => {
  const openA22ModuleBtn = document.getElementById("openA22ModuleBtn");
  const a22FormatStatus = document.getElementById("a22FormatStatus");
  const periodSelect = document.getElementById("formatPeriodSelect");
  const runSelect = document.getElementById("formatRunSelect");
  const dateInput = document.getElementById("formatDate");
  const outgoingInput = document.getElementById("formatOutgoing");
  const incomingInput = document.getElementById("formatIncoming");
  const issuedByInput = document.getElementById("formatIssuedBy");
  const notesInput = document.getElementById("formatNotes");
  const validateBtn = document.getElementById("formatValidateBtn");
  const generateBtn = document.getElementById("formatGenerateBtn");
  const statusEl = document.getElementById("formatStatus");
  const responsiblesList = document.getElementById("formatResponsiblesList");

  function selectedPeriodId() {
    return periodSelect?.value ? Number(periodSelect.value) : null;
  }

  function selectedRunId() {
    return runSelect?.value ? Number(runSelect.value) : null;
  }

  async function loadClosedRuns() {
    if (!runSelect) return;
    const pid = selectedPeriodId();
    if (!pid) {
      runSelect.innerHTML = '<option value="">-- Selecciona jornada cerrada --</option>';
      return;
    }
    const data = await App.get(`/runs?period_id=${encodeURIComponent(String(pid))}`);
    const closed = (data.runs || []).filter(
      (r) => String(r.status || "").toLowerCase() === "closed",
    );
    runSelect.innerHTML =
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
  }

  async function loadResponsibles() {
    if (!responsiblesList) return;
    const data = await App.get("/responsibles");
    const items = data.responsibles || [];
    responsiblesList.innerHTML = "";
    items.forEach((name) => {
      const opt = document.createElement("option");
      opt.value = String(name || "");
      responsiblesList.appendChild(opt);
    });
    App.setStatus(statusEl, `Responsables disponibles: ${items.length}`);
  }

  async function validateClearance() {
    const pid = selectedPeriodId();
    const rid = selectedRunId();
    if (!pid) return App.setStatus(statusEl, "Selecciona un periodo", true);
    if (!rid) return App.setStatus(statusEl, "Selecciona una jornada cerrada", true);

    const data = await App.get(
      `/paz_y_salvo/validate?period_id=${encodeURIComponent(String(pid))}&run_id=${encodeURIComponent(String(rid))}`,
    );
    App.setStatus(statusEl, data.message || "", !data.allowed);
    return data;
  }

  async function generateClearance() {
    const pid = selectedPeriodId();
    const rid = selectedRunId();
    if (!pid) return App.setStatus(statusEl, "Selecciona un periodo", true);
    if (!rid) return App.setStatus(statusEl, "Selecciona una jornada cerrada", true);

    const outgoing = String(outgoingInput?.value || "").trim();
    const incoming = String(incomingInput?.value || "").trim();
    const issuedBy = String(issuedByInput?.value || "").trim();
    const reportDate = String(dateInput?.value || "").trim();
    const observations = String(notesInput?.value || "").trim();
    if (!outgoing) return App.setStatus(statusEl, "Selecciona responsable saliente", true);
    if (!incoming) return App.setStatus(statusEl, "Selecciona responsable entrante", true);

    App.setStatus(statusEl, "Generando paz y salvo...");
    const res = await fetch("/paz_y_salvo/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        period_id: pid,
        run_id: rid,
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
      statusEl,
      "Paz y salvo generado. El PDF fue descargado y guardado en Documentos.",
    );
  }

  openA22ModuleBtn?.addEventListener("click", () => {
    window.location.href = "/inventario";
    App.setStatus(a22FormatStatus, "Abriendo modulo Inventario y escaneo...");
  });

  periodSelect?.addEventListener("change", async () => {
    try {
      await loadClosedRuns();
      await loadResponsibles();
      App.setStatus(statusEl, "");
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  runSelect?.addEventListener("change", async () => {
    try {
      await loadResponsibles();
      App.setStatus(statusEl, "");
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  validateBtn?.addEventListener("click", async () => {
    try {
      await validateClearance();
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  generateBtn?.addEventListener("click", async () => {
    try {
      await generateClearance();
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  if (dateInput && !dateInput.value) {
    dateInput.value = new Date().toISOString().slice(0, 10);
  }
  if (issuedByInput && !issuedByInput.value) {
    issuedByInput.value = "Responsable activos fijos";
  }

  App.loadPeriods(periodSelect)
    .then(() => loadClosedRuns())
    .then(() => loadResponsibles())
    .catch((err) => App.setStatus(statusEl, err.message, true));
});
