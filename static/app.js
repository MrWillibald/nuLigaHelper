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
      showToast("Dienst gespeichert", true);
    } else {
      showToast(result.error || "Fehler beim Speichern", false);
    }
  });
});

document.querySelectorAll(".jteam-input").forEach((input) => {
  let timer = null;
  input.addEventListener("input", () => {
    clearTimeout(timer);
    timer = setTimeout(async () => {
      const result = await postJSON(`/api/games/${input.dataset.game}/jteam`, {
        value: input.value,
      });
      if (result.ok) {
        flashCard(input.dataset.game);
        showToast("Kampfgericht-Team gespeichert", true);
      } else {
        showToast(result.error || "Fehler beim Speichern", false);
      }
    }, 500);
  });
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
