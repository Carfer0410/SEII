(() => {
  const tableBody = document.querySelector("#usersTable tbody");
  const statusEl = document.getElementById("adminStatus");
  const createBtn = document.getElementById("createUserBtn");
  const newUsername = document.getElementById("newUsername");
  const newFullName = document.getElementById("newFullName");
  const newEmail = document.getElementById("newEmail");
  const newPassword = document.getElementById("newPassword");
  const newIsAdmin = document.getElementById("newIsAdmin");

  function esc(v) {
    return App.escapeHtml(String(v || ""));
  }

  function yesNo(v) {
    return v ? "Si" : "No";
  }

  function badge(text, tone) {
    return `<span class="admin-badge ${tone || ""}">${esc(text)}</span>`;
  }

  async function loadUsers() {
    const data = await App.get("/admin/users/api");
    const rows = data.items || [];
    tableBody.innerHTML = rows
      .map((u) => {
        const lastLogin = u.last_login_at_local || "-";
        return `
          <tr data-id="${u.id}">
            <td>${u.id}</td>
            <td><span class="admin-user-name">${esc(u.username)}</span></td>
            <td><input class="cell-full-name admin-cell-input" type="text" value="${esc(u.full_name)}" /></td>
            <td><input class="cell-email admin-cell-input" type="email" value="${esc(u.email)}" /></td>
            <td><label class="admin-inline-check"><input class="cell-is-admin" type="checkbox" ${u.is_admin ? "checked" : ""} /></label></td>
            <td>${u.is_active ? badge("Activo", "ok") : badge("Inactivo", "muted")}</td>
            <td>${esc(lastLogin)}</td>
            <td class="row-actions">
              <button class="btn-save admin-btn-save">Guardar</button>
              <button class="btn-toggle admin-btn-toggle">${u.is_active ? "Desactivar" : "Activar"}</button>
              <button class="btn-reset admin-btn-reset">Clave</button>
            </td>
          </tr>
        `;
      })
      .join("");
  }

  async function createUser() {
    const payload = {
      username: String(newUsername.value || "").trim(),
      full_name: String(newFullName.value || "").trim(),
      email: String(newEmail.value || "").trim(),
      password: String(newPassword.value || ""),
      is_admin: !!newIsAdmin.checked,
    };
    await App.post("/admin/users/api", payload);
    newUsername.value = "";
    newFullName.value = "";
    newEmail.value = "";
    newPassword.value = "";
    newIsAdmin.checked = false;
    await loadUsers();
    App.setStatus(statusEl, "Usuario creado correctamente.");
  }

  async function saveUser(rowEl) {
    const id = Number(rowEl.dataset.id);
    const payload = {
      full_name: rowEl.querySelector(".cell-full-name")?.value || "",
      email: rowEl.querySelector(".cell-email")?.value || "",
      is_admin: !!rowEl.querySelector(".cell-is-admin")?.checked,
    };
    await App.patch(`/admin/users/api/${id}`, payload);
    await loadUsers();
    App.setStatus(statusEl, "Usuario actualizado.");
  }

  async function toggleUser(rowEl) {
    const id = Number(rowEl.dataset.id);
    await App.post(`/admin/users/api/${id}/toggle_active`, {});
    await loadUsers();
    App.setStatus(statusEl, "Estado de usuario actualizado.");
  }

  async function resetPassword(rowEl) {
    const id = Number(rowEl.dataset.id);
    const pwd = window.prompt("Nueva contraseña (minimo 6 caracteres):");
    if (!pwd) return;
    await App.post(`/admin/users/api/${id}/reset_password`, { new_password: pwd });
    App.setStatus(statusEl, "Contraseña reiniciada correctamente.");
  }

  createBtn?.addEventListener("click", async () => {
    try {
      await createUser();
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  tableBody?.addEventListener("click", async (e) => {
    const btn = e.target.closest("button");
    if (!btn) return;
    const rowEl = e.target.closest("tr");
    if (!rowEl) return;
    try {
      if (btn.classList.contains("btn-save")) {
        await saveUser(rowEl);
      } else if (btn.classList.contains("btn-toggle")) {
        await toggleUser(rowEl);
      } else if (btn.classList.contains("btn-reset")) {
        await resetPassword(rowEl);
      }
    } catch (err) {
      App.setStatus(statusEl, err.message, true);
    }
  });

  loadUsers().catch((err) => App.setStatus(statusEl, err.message, true));
})();
