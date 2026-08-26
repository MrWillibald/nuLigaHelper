function showToast(message, ok) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.toggle("error", !ok);
  toast.classList.add("show");
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => toast.classList.remove("show"), 2600);
}

async function postJSON(url, body) {
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    return await response.json();
  } catch (err) {
    return { ok: false, error: "Server nicht erreichbar" };
  }
}

function flashCard(gameId) {
  const card = document.getElementById("game-" + gameId);
  if (!card) return;
  card.classList.remove("saved");
  void card.offsetWidth;
  card.classList.add("saved");
}

document.querySelectorAll("select[data-role]").forEach((select) => {
  select.addEventListener("change", async () => {
    const result = await postJSON("/api/assignment", {
      game_id: Number(select.dataset.game),
      role: select.dataset.role,
      slot: select.dataset.slot === "" ? null : Number(select.dataset.slot),
      person_id: select.value ? Number(select.value) : null,
    });
    if (result.ok) {
      flashCard(select.dataset.game);
      const option = select.selectedOptions[0];
      if (option && option.classList.contains("option-playing")) {
        showToast("Achtung: Person spielt selbst in diesem Spiel", false);
        select.classList.add("select-warn");
      } else if (option && option.classList.contains("foreign-option")) {
        showToast("Hinweis: Person gehört nicht zum zugewiesenen Team", false);
        select.classList.add("select-warn");
      } else {
        showToast("Dienst gespeichert", true);
        select.classList.remove("select-warn");
      }
    } else {
      showToast(result.error || "Fehler beim Speichern", false);
    }
  });
});

document.querySelectorAll(".team-select").forEach((select) => {
  select.addEventListener("change", async () => {
    const gameId = select.dataset.game;
    const result = await postJSON(`/api/games/${gameId}/team`, {
      team_id: select.value ? Number(select.value) : null,
    });
    if (result.ok) {
      flashCard(gameId);
      showToast("Team gespeichert", true);
      setTimeout(() => {
        window.location.hash = "game-" + gameId;
        window.location.reload();
      }, 500);
    } else {
      showToast(result.error || "Fehler beim Speichern", false);
      const previous = select.getAttribute("data-prev") || "";
      select.value = previous;
    }
  });
  select.addEventListener("focus", () => select.setAttribute("data-prev", select.value));
});

document.querySelectorAll("[data-delete-person]").forEach((button) => {
  button.addEventListener("click", async () => {
    const name = button.dataset.name;
    if (!confirm(`'${name}' wirklich löschen? Alle Diensteinträge werden mit entfernt.`)) {
      return;
    }
    try {
      await fetch(`/personen/${button.dataset.deletePerson}/delete`, { method: "POST" });
      window.location.reload();
    } catch (err) {
      showToast("Server nicht erreichbar", false);
    }
  });
});
