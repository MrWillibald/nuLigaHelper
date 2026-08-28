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

function otherRoleSelects(card, currentSelect) {
  return Array.from(card.querySelectorAll("select[data-role]")).filter(
    (s) => s !== currentSelect
  );
}

function removePersonOption(card, currentSelect, personId) {
  const value = String(personId);
  otherRoleSelects(card, currentSelect).forEach((s) => {
    const option = s.querySelector('option[value="' + value + '"]');
    if (option) option.remove();
  });
}

function insertOptionSorted(select, option) {
  const label = option.textContent.trim();
  const options = Array.from(select.options);
  let pos = 1; // keep the "offen" placeholder first
  while (pos < options.length && options[pos].textContent.trim().localeCompare(label) < 0) {
    pos++;
  }
  select.insertBefore(option, options[pos] || null);
}

function addPersonOption(card, currentSelect, option) {
  const clone = option.cloneNode(true);
  clone.removeAttribute("selected");
  otherRoleSelects(card, currentSelect).forEach((s) => {
    if (s.querySelector('option[value="' + clone.value + '"]')) return;
    insertOptionSorted(s, clone.cloneNode(true));
  });
}

const prevOptions = new WeakMap();
document.querySelectorAll("select[data-role]").forEach((select) => {
  select.addEventListener("focus", () => {
    prevOptions.set(select, select.selectedOptions[0]);
  });
  select.addEventListener("change", async () => {
    const previous = prevOptions.get(select);
    const previousId = previous && previous.value ? Number(previous.value) : null;
    const newId = select.value ? Number(select.value) : null;
    const result = await postJSON("/api/assignment", {
      game_id: Number(select.dataset.game),
      role: select.dataset.role,
      slot: select.dataset.slot === "" ? null : Number(select.dataset.slot),
      person_id: newId,
    });
    if (!result.ok) {
      select.value = previousId === null ? "" : String(previousId);
      showToast(result.error || "Fehler beim Speichern", false);
      return;
    }
    flashCard(select.dataset.game);
    const card = document.getElementById("game-" + select.dataset.game);
    // the newly assigned person must not be offered for other tasks
    if (newId) removePersonOption(card, select, newId);
    // a freed person may be offered again for other tasks
    if (previousId !== null && previousId !== newId) {
      addPersonOption(card, select, previous);
    }
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

document.querySelectorAll(".mv-select").forEach((select) => {
  select.addEventListener("focus", () => select.setAttribute("data-prev", select.value));
  select.addEventListener("change", async () => {
    const result = await postJSON(`/api/teams/${select.dataset.team}/mv`, {
      person_id: select.value ? Number(select.value) : null,
    });
    if (result.ok) {
      showToast("MV gespeichert", true);
      window.location.reload();
    } else {
      showToast(result.error || "Fehler beim Speichern", false);
      select.value = select.getAttribute("data-prev") || "";
    }
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
