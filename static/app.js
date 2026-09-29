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
      headers: {
        "Content-Type": "application/json",
        "X-CSRF-Token": document.querySelector('meta[name="csrf-token"]').content,
      },
      body: JSON.stringify(body),
    });
    if (response.status === 401) {
      return {
        ok: false, loginRequired: true,
        error: "Deine Sitzung ist abgelaufen. Bitte melde dich erneut an.",
      };
    }
    if (response.status === 409) {
      const conflict = await response.json();
      return { ...conflict, conflict: true };
    }
    let data;
    try {
      data = await response.json();
    } catch (err) {
      data = { ok: false, error: `Serverfehler (${response.status})` };
    }
    return data;
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

function flashBlock(blockId) {
  const card = document.getElementById("block-" + blockId);
  if (!card) return;
  card.classList.remove("saved");
  void card.offsetWidth;
  card.classList.add("saved");
}

function otherRoleSelects(card, currentSelect) {
  return Array.from(card.querySelectorAll("select[data-role], select[data-block-assignment]")).filter(
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

function comparePersonOptions(left, right) {
  const groupDifference = Number(left.dataset.sortGroup) - Number(right.dataset.sortGroup);
  if (groupDifference !== 0) return groupDifference;
  const nameDifference = left.dataset.sortName.localeCompare(right.dataset.sortName);
  if (nameDifference !== 0) return nameDifference;
  return Number(left.dataset.sortId) - Number(right.dataset.sortId);
}

function insertOptionSorted(select, option) {
  const options = Array.from(select.options);
  let pos = 1; // keep the "offen" placeholder first
  while (pos < options.length && comparePersonOptions(options[pos], option) < 0) {
    pos++;
  }
  select.insertBefore(option, options[pos] || null);
}

function addPersonOption(card, currentSelect, option) {
  if (!option.dataset.sortId) return;
  const clone = option.cloneNode(true);
  clone.removeAttribute("selected");
  otherRoleSelects(card, currentSelect).forEach((s) => {
    const allowed = card._candidateSlots?.get(candidateSlotKey(s));
    if (allowed && !allowed.has(Number(clone.value))) return;
    if (s.querySelector('option[value="' + clone.value + '"]')) return;
    insertOptionSorted(s, clone.cloneNode(true));
  });
}

function candidateSlotKey(select) {
  return select.hasAttribute("data-block-assignment")
    ? select.dataset.slot : `${select.dataset.role}:${select.dataset.slot}`;
}

function makeCandidateOption(person) {
  const option = document.createElement("option");
  option.value = String(person.id);
  option.dataset.sortGroup = String(person.sort_group);
  option.dataset.sortName = person.sort_name;
  option.dataset.sortId = String(person.id);
  option.textContent = `${person.name} · ${person.team_label}`;
  if (person.hint === "playing") {
    option.className = "option-playing";
    option.title = `${person.name} spielt in diesem Spiel selbst`;
    option.textContent += " • spielt selbst";
  } else if (person.hint === "outside") {
    option.className = "foreign-option";
    option.title = `${person.name} gehört zu ${person.team_label}`;
    option.textContent += " • außerhalb";
  }
  return option;
}

function populateCandidateCard(card, payload) {
  const selects = Array.from(card.querySelectorAll("select[data-role], select[data-block-assignment]"));
  const people = new Map(payload.people.map((person) => [person.id, person]));
  const taken = new Set(Array.from(card.querySelectorAll("[data-occupant-id]"))
    .map((label) => Number(label.dataset.occupantId)).filter(Boolean));
  const plans = selects.map((select) => {
    const key = candidateSlotKey(select);
    const slot = payload.slots[key];
    const selectedId = select.value ? Number(select.value) : null;
    if (!slot || slot.occupant_id !== selectedId) {
      throw new Error("Die Einteilung hat sich geändert. Bitte lade die Seite neu.");
    }
    const allowed = new Set(slot.candidate_ids);
    const options = slot.candidate_ids
      .filter((id) => people.has(id) && (!taken.has(id) || id === selectedId))
      .map((id) => makeCandidateOption(people.get(id)))
      .sort(comparePersonOptions);
    return { select, key, selectedId, allowed, options };
  });
  const allowedBySlot = new Map();
  plans.forEach(({ select, key, selectedId, allowed, options }) => {
    allowedBySlot.set(key, allowed);
    if (selectedId && people.has(selectedId)) select.selectedOptions[0].remove();
    const fragment = document.createDocumentFragment();
    options.forEach((option) => {
      if (Number(option.value) === selectedId) option.selected = true;
      fragment.appendChild(option);
    });
    select.appendChild(fragment);
    if (selectedId) select.value = String(selectedId);
    select.disabled = false;
  });
  card._candidateSlots = allowedBySlot;
}

function candidateStatus(card, message, retry = false, login = false) {
  const status = card.querySelector("[data-candidate-status]");
  if (!status) return;
  status.hidden = !message;
  status.querySelector("[data-candidate-message]").textContent = message;
  status.querySelector("[data-candidate-retry]").hidden = !retry;
  status.querySelector("[data-candidate-login]").hidden = !login;
}

async function loadCandidateCard(card) {
  if (card._candidateLoaded || card._candidateLoading) return;
  card._candidateLoading = true;
  candidateStatus(card, "Helferliste wird geladen …");
  try {
    const response = await fetch(card.dataset.candidateUrl, { cache: "no-store" });
    if (response.status === 401) {
      throw Object.assign(new Error("Deine Sitzung ist abgelaufen. Bitte melde dich erneut an."), { login: true });
    }
    if (!response.ok) throw new Error("Helferliste konnte nicht geladen werden.");
    populateCandidateCard(card, await response.json());
    card._candidateLoaded = true;
    candidateStatus(card, "");
  } catch (error) {
    candidateStatus(card, error.message || "Helferliste konnte nicht geladen werden.", true, error.login);
  } finally {
    card._candidateLoading = false;
  }
}

globalThis.nuLigaCandidateTools = { populateCandidateCard, loadCandidateCard };

document.querySelectorAll("details[data-candidate-url]").forEach((card) => {
  if (!card.querySelector("select[data-role], select[data-block-assignment]")) return;
  card.addEventListener("toggle", () => {
    if (card.open) loadCandidateCard(card);
  });
  card.querySelector("[data-candidate-retry]").addEventListener("click", () => {
    if (card.querySelector("[data-candidate-message]").textContent.includes("Bitte lade die Seite neu")) {
      window.location.reload();
    } else {
      loadCandidateCard(card);
    }
  });
  if (card.open) loadCandidateCard(card);
});

function updateCoverage(card, previousId, newId, role = null) {
  const coverage = card && card.querySelector(".coverage");
  if (!coverage || (coverage.dataset.progressKind === "game" && role === "Unterstützung")) return;
  const change = Number(newId !== null) - Number(previousId !== null);
  if (change === 0) return;
  const total = Number(coverage.dataset.progressTotal);
  const filled = Math.max(0, Math.min(total, Number(coverage.dataset.progressFilled) + change));
  const percent = Math.round(filled * 100 / total);
  coverage.dataset.progressFilled = String(filled);
  const unit = coverage.dataset.progressKind === "game" ? "Pflichtdiensten" : "Plätzen";
  coverage.querySelector("[data-progress-count]").textContent = `${filled} von ${total} ${unit} besetzt`;
  coverage.querySelector("[data-progress-percent]").textContent = `${percent} %`;
  coverage.querySelector(".coverage-fill").style.width = `${percent}%`;
  coverage.querySelector('[role="progressbar"]').setAttribute("aria-valuenow", String(filled));
}

// Expose the small, side-effect-free option helpers for the DOM regression test.
globalThis.nuLigaOptionTools = {
  comparePersonOptions, insertOptionSorted, addPersonOption, removePersonOption,
};
globalThis.nuLigaProgressTools = { updateCoverage };

const prevOptions = new WeakMap();
document.querySelectorAll("select[data-role]").forEach((select) => {
  select.addEventListener("focus", () => {
    prevOptions.set(select, select.selectedOptions[0]);
  });
  select.addEventListener("change", async () => {
    const previous = prevOptions.get(select);
    const previousId = previous && previous.value ? Number(previous.value) : null;
    const newId = select.value ? Number(select.value) : null;
    const requestBody = {
      game_id: Number(select.dataset.game), role: select.dataset.role,
      slot: Number(select.dataset.slot), expected_person_id: previousId,
    };
    let result = { ok: true };
    if (previousId !== null) {
      result = await postJSON("/api/assignment/release", requestBody);
    }
    if (result.ok && newId !== null) {
      result = await postJSON("/api/assignment/claim", {
        ...requestBody, expected_person_id: null, person_id: newId,
      });
    }
    if (!result.ok) {
      select.value = previousId === null ? "" : String(previousId);
      showToast(result.error || "Fehler beim Speichern", false);
      if (result.loginRequired) setTimeout(() => { window.location.href = "/login"; }, 1200);
      if (result.conflict || previousId !== null) {
        setTimeout(() => { window.location.reload(); }, 500);
      }
      return;
    }
    const card = document.getElementById("game-" + select.dataset.game);
    const label = select.closest?.("[data-occupant-id]");
    if (label) label.dataset.occupantId = newId || "";
    updateCoverage(card, previousId, newId, select.dataset.role);
    flashCard(select.dataset.game);
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

document.querySelectorAll("select[data-block-assignment]").forEach((select) => {
  select.addEventListener("focus", () => {
    prevOptions.set(select, select.selectedOptions[0]);
  });
  select.addEventListener("change", async () => {
    const previous = prevOptions.get(select);
    const previousId = previous && previous.value ? Number(previous.value) : null;
    const newId = select.value ? Number(select.value) : null;
    const requestBody = {
      block_id: Number(select.dataset.block),
      slot: Number(select.dataset.slot),
      expected_person_id: previousId,
    };
    let result = { ok: true };
    if (previousId !== null) {
      result = await postJSON("/api/block-assignment/release", requestBody);
    }
    if (result.ok && newId !== null) {
      result = await postJSON("/api/block-assignment/claim", {
        ...requestBody, expected_person_id: null, person_id: newId,
      });
    }
    if (!result.ok) {
      select.value = previousId === null ? "" : String(previousId);
      showToast(result.error || "Fehler beim Speichern", false);
      if (result.loginRequired) setTimeout(() => { window.location.href = "/login"; }, 1200);
      if (result.conflict || previousId !== null) {
        setTimeout(() => { window.location.reload(); }, 500);
      }
      return;
    }
    const card = document.getElementById("block-" + select.dataset.block);
    const label = select.closest?.("[data-occupant-id]");
    if (label) label.dataset.occupantId = newId || "";
    if (newId) removePersonOption(card, select, newId);
    if (previousId !== null && previousId !== newId) addPersonOption(card, select, previous);
    updateCoverage(card, previousId, newId);
    flashBlock(select.dataset.block);
    showToast("Tagesdienst gespeichert", true);
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
    if (!confirm(`'${name}' wirklich löschen? Alle Diensteinträge werden entfernt. Zum Ausscheiden bitte stattdessen deaktivieren.`)) {
      return;
    }
    const form = document.createElement("form");
    form.method = "POST";
    form.action = `/personen/${button.dataset.deletePerson}/delete`;
    const token = document.createElement("input");
    token.type = "hidden"; token.name = "csrf_token";
    token.value = document.querySelector('meta[name="csrf-token"]').content;
    form.appendChild(token); document.body.appendChild(form); form.submit();
  });
});

document.querySelectorAll("[data-team-dialog-open]").forEach((button) => {
  button.addEventListener("click", () => {
    const dialog = document.getElementById(button.dataset.teamDialogOpen);
    if (!dialog) return;
    if (dialog.matches("[data-team-picker-dialog]")) {
      dialog._checkedSnapshot = Array.from(
        dialog.querySelectorAll("[data-team-picker-input]")
      ).map((input) => input.checked);
      dialog._pickerApplied = false;
    }
    dialog.showModal();
  });
});

document.querySelectorAll(".team-membership-dialog").forEach((dialog) => {
  const form = dialog.querySelector("form");
  dialog.querySelectorAll("[data-team-dialog-close]").forEach((button) => {
    button.addEventListener("click", () => dialog.close());
  });
  dialog.addEventListener("close", () => form?.reset());
  dialog.addEventListener("click", (event) => {
    if (event.target !== dialog) return;
    const bounds = dialog.getBoundingClientRect();
    const inside = event.clientX >= bounds.left && event.clientX <= bounds.right
      && event.clientY >= bounds.top && event.clientY <= bounds.bottom;
    if (!inside) dialog.close();
  });
});

document.querySelectorAll("[data-team-picker-dialog]").forEach((dialog) => {
  const inputs = Array.from(dialog.querySelectorAll("[data-team-picker-input]"));
  const badgeContainer = document.querySelector("[data-new-team-badges]");

  function renderTeamBadges() {
    if (!badgeContainer) return;
    badgeContainer.replaceChildren();
    const selected = inputs.filter((input) => input.checked);
    if (!selected.length) {
      const empty = document.createElement("span");
      empty.className = "new-user-team-empty";
      empty.textContent = "Noch keine Mannschaft ausgewählt";
      badgeContainer.appendChild(empty);
      return;
    }
    selected.forEach((input) => {
      const badge = document.createElement("span");
      badge.className = "new-user-team-badge";
      badge.textContent = input.dataset.teamName;
      badgeContainer.appendChild(badge);
    });
  }

  dialog.querySelector("[data-team-picker-apply]")?.addEventListener("click", () => {
    dialog._pickerApplied = true;
    renderTeamBadges();
    dialog.close();
  });
  dialog.addEventListener("close", () => {
    if (!dialog._pickerApplied && dialog._checkedSnapshot) {
      inputs.forEach((input, index) => {
        input.checked = dialog._checkedSnapshot[index];
      });
    }
  });
});

document.querySelectorAll("[data-auth-form]").forEach((form) => {
  const channelInputs = Array.from(form.querySelectorAll('input[name="channel"]'));
  const emailInput = form.querySelector('input[name="email"]');
  const phoneInput = form.querySelector('input[name="phone"]');
  const countrySelect = form.querySelector('[name="country_code"]');
  const customCountry = form.querySelector("[data-custom-country]");
  const customCountryInput = form.querySelector('[name="custom_country_code"]');
  const requestButton = form.querySelector('button[value="request_code"]');
  const availability = form.querySelector(".auth-route-availability");
  const locked = form.dataset.authLocked === "true";

  function emailIsValid() {
    return Boolean(emailInput?.value.trim() && emailInput.checkValidity());
  }

  function phoneIsValid() {
    const candidate = phoneInput?.value.trim() || "";
    if (!candidate || !/^[+\d][\d\s()./-]*$/.test(candidate)) return false;
    const callingCode = countrySelect?.value === "custom"
      ? customCountryInput?.value.trim() || ""
      : countrySelect?.value || "";
    if (!/^\+?[1-9]\d{0,2}$/.test(callingCode)) return false;
    const digits = candidate.replace(/\D/g, "");
    const prefixDigits = callingCode.replace(/\D/g, "");
    const internationalDigits = digits.replace(/^00/, "");
    const explicitInternational = candidate.startsWith("+") || candidate.startsWith("00");
    if (explicitInternational && !internationalDigits.startsWith(prefixDigits)) {
      return false;
    }
    const totalDigits = explicitInternational
      ? internationalDigits.length
      : prefixDigits.length + digits.length;
    return digits.length >= 6 && totalDigits <= 15;
  }

  function updateRouteAvailability() {
    if (!channelInputs.length) return;
    form.classList.add("auth-enhanced");
    const validRoutes = { email: emailIsValid(), sms: phoneIsValid() };
    channelInputs.forEach((input) => {
      input.disabled = locked || !validRoutes[input.value];
      if (input.disabled) input.checked = false;
    });
    const availableRoutes = channelInputs.filter((input) => !input.disabled);
    if (!channelInputs.some((input) => input.checked) && availableRoutes.length === 1) {
      availableRoutes[0].checked = true;
    }

    const invalidSupplied = Boolean(
      (emailInput?.value.trim() && !validRoutes.email)
      || (phoneInput?.value.trim() && !validRoutes.sms)
    );
    const selected = channelInputs.some((input) => input.checked && !input.disabled);
    if (requestButton) {
      requestButton.disabled = locked || invalidSupplied || !selected;
    }
    if (availability) {
      availability.textContent = invalidSupplied
        ? "Bitte korrigiere oder leere ungültige Kontaktangaben."
        : availableRoutes.length
          ? "Wähle einen der verfügbaren Kontaktwege."
          : "Gib zuerst eine gültige E-Mail-Adresse oder Mobilnummer ein.";
    }
    updateCustomCountry();
  }

  function updateCustomCountry() {
    if (!countrySelect || !customCountry) return;
    const active = countrySelect.value === "custom";
    customCountry.hidden = !active;
    customCountry.querySelectorAll("input").forEach((input) => {
      input.disabled = locked || !active;
    });
  }

  channelInputs.forEach((input) => input.addEventListener("change", updateRouteAvailability));
  [emailInput, phoneInput, customCountryInput].forEach((input) => {
    if (input) input.addEventListener("input", updateRouteAvailability);
  });
  if (countrySelect) countrySelect.addEventListener("change", updateRouteAvailability);
  updateRouteAvailability();

  const invalid = form.querySelector('[aria-invalid="true"]:not(:disabled)');
  const code = form.querySelector('input[name="code"]:not(:disabled)');
  if (invalid) invalid.focus();
  else if (code) code.focus();
});
