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
