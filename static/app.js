function showToast(message, ok) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.toggle("error", !ok);
  toast.classList.add("show");
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => toast.classList.remove("show"), 2600);
}

const activeTaskHelp = new Set();

function taskHelpPlacement(anchor, size, viewport) {
  const margin = 12;
  const gap = 6;
  const viewportLeft = viewport.left || 0;
  const viewportTop = viewport.top || 0;
  const below = Math.max(0, viewportTop + viewport.height - margin - anchor.bottom - gap);
  const above = Math.max(0, anchor.top - gap - viewportTop - margin);
  const side = size.height <= below || below >= above ? "below" : "above";
  const maxHeight = Math.max(0, Math.min(viewport.height - margin * 2, side === "below" ? below : above));
  const height = Math.min(size.height, maxHeight);
  const left = Math.max(viewportLeft + margin, Math.min(
    anchor.left + ((anchor.width ?? anchor.right - anchor.left) - size.width) / 2,
    viewportLeft + viewport.width - margin - size.width,
  ));
  const top = Math.max(viewportTop + margin, Math.min(
    side === "below" ? anchor.bottom + gap : anchor.top - gap - height,
    viewportTop + viewport.height - margin - height,
  ));
  return { left, top, maxHeight, side };
}

function positionTaskHelp(button, description) {
  if (!button.getBoundingClientRect || !description.getBoundingClientRect) return;
  const visual = globalThis.window?.visualViewport;
  const viewport = {
    width: visual?.width || globalThis.window?.innerWidth || document.documentElement?.clientWidth,
    height: visual?.height || globalThis.window?.innerHeight || document.documentElement?.clientHeight,
    left: visual?.offsetLeft || 0,
    top: visual?.offsetTop || 0,
  };
  if (!viewport.width || !viewport.height) return;
  const scrollTop = description.scrollTop;
  description.style.width = `${Math.min(400, viewport.width - 24)}px`;
  description.style.maxHeight = `${viewport.height - 24}px`;
  const placement = taskHelpPlacement(
    button.getBoundingClientRect(), description.getBoundingClientRect(), viewport,
  );
  description.style.left = `${placement.left}px`;
  description.style.top = `${placement.top}px`;
  description.style.maxHeight = `${placement.maxHeight}px`;
  description.scrollTop = scrollTop;
}

function refreshTaskHelp() {
  Array.from(activeTaskHelp).forEach((help) => {
    if (help.available()) help.position();
    else help.dismiss();
  });
}

function initializeTaskHelp(field) {
  const button = field.querySelector("[data-task-help]");
  const description = field.querySelector("[data-task-description]");
  if (!button || !description || button._taskHelpBound) return;
  button._taskHelpBound = true;
  let hovered = false;
  let focused = false;
  let pinned = false;
  let hoverDismissed = false;
  let focusDismissed = false;
  let leaveTimer;
  const nativePopover = typeof description.showPopover === "function";
  const available = () => {
    if (button.isConnected === false) return false;
    for (let card = field.closest?.("details"); card; card = card.parentElement?.closest?.("details")) {
      if (card.open === false) return false;
    }
    return true;
  };
  const help = {
    available, position: () => positionTaskHelp(button, description),
    dismiss: () => dismiss(),
  };
  const render = () => {
    const visible = available() && (pinned || (hovered && !hoverDismissed) || (focused && !focusDismissed));
    if (visible) {
      const opening = description.hidden;
      description.hidden = false;
      if (opening && nativePopover) description.showPopover();
      activeTaskHelp.add(help);
      help.position();
    } else {
      if (!description.hidden && nativePopover) {
        try {
          description.hidePopover();
        } catch (error) {
          // Removed popovers have already left the top layer before the
          // mutation observer runs; their disconnected state forbids this API.
          if (error.name !== "InvalidStateError") throw error;
        }
      }
      description.hidden = true;
      activeTaskHelp.delete(help);
    }
    button.setAttribute("aria-expanded", String(visible));
  };
  const dismiss = () => {
    clearTimeout(leaveTimer);
    if (description.contains?.(document.activeElement) && available()) {
      // A hidden bubble cannot retain focus. Return keyboard readers to the
      // information control before closing, without scrolling the page.
      button.focus({ preventScroll: true });
    }
    pinned = false;
    hoverDismissed = hovered;
    focusDismissed = focused;
    render();
  };
  const enter = () => {
    clearTimeout(leaveTimer);
    hovered = true;
    hoverDismissed = false;
    render();
  };
  const leave = (event) => {
    if (button.contains?.(event.relatedTarget) || description.contains?.(event.relatedTarget)) return;
    clearTimeout(leaveTimer);
    // Allow the pointer to cross the small gap into the floating description.
    leaveTimer = setTimeout(() => {
      hovered = false;
      hoverDismissed = false;
      render();
    }, 120);
  };
  button.addEventListener("pointerenter", enter);
  button.addEventListener("pointerleave", leave);
  description.addEventListener("pointerenter", enter);
  description.addEventListener("pointerleave", leave);
  field.addEventListener("pointerleave", leave);
  const focus = () => {
    focused = true;
    focusDismissed = false;
    render();
  };
  const blur = (event) => {
    if (button.contains?.(event?.relatedTarget) || description.contains?.(event?.relatedTarget)) return;
    focused = false;
    focusDismissed = false;
    render();
  };
  button.addEventListener("focus", focus);
  button.addEventListener("blur", blur);
  description.addEventListener("focus", focus);
  description.addEventListener("blur", blur);
  button.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (pinned) {
      dismiss();
    } else {
      pinned = true;
      render();
    }
  });
  field.addEventListener("keydown", (event) => {
    if (event.key !== "Escape" || description.hidden) return;
    event.preventDefault();
    event.stopPropagation();
    // Keep focus on the information control or the existing assignment field.
    dismiss();
  });
}

let dynamicTaskHelpIndex = 0;
function createTaskHelp(label, description, editable = false) {
  const index = `dynamic-${++dynamicTaskHelpIndex}`;
  const field = document.createElement("div");
  field.className = "assign-field";
  const heading = document.createElement("div");
  heading.className = "assign-task-heading";
  const name = document.createElement(editable ? "label" : "span");
  name.textContent = label;
  if (editable) name.setAttribute("for", `task-assignment-${index}`);
  const button = document.createElement("button");
  button.className = "task-help";
  button.type = "button";
  button.setAttribute("data-task-help", "");
  button.setAttribute("aria-label", `Informationen zu ${label}`);
  button.setAttribute("aria-expanded", "false");
  button.setAttribute("aria-controls", `task-description-${index}`);
  button.setAttribute("aria-describedby", `task-description-${index}`);
  const icon = document.createElement("span");
  icon.setAttribute("aria-hidden", "true");
  icon.textContent = "i";
  button.appendChild(icon);
  heading.appendChild(name);
  heading.appendChild(button);
  const text = document.createElement("p");
  text.className = "task-description";
  text.id = `task-description-${index}`;
  text.setAttribute("data-task-description", "");
  text.setAttribute("popover", "manual");
  text.setAttribute("tabindex", "0");
  text.hidden = true;
  text.textContent = description;
  field.appendChild(heading);
  field.appendChild(text);
  initializeTaskHelp(field);
  return field;
}

document.querySelectorAll(".assign-field").forEach(initializeTaskHelp);
globalThis.window?.addEventListener?.("resize", refreshTaskHelp);
globalThis.window?.visualViewport?.addEventListener("resize", refreshTaskHelp);
globalThis.window?.visualViewport?.addEventListener("scroll", refreshTaskHelp);
document.addEventListener?.("scroll", (event) => {
  // Reading the bounded bubble must not remeasure it and reset its scroll.
  if (event.target?.getAttribute?.("data-task-description") != null) return;
  refreshTaskHelp();
}, true);
document.addEventListener?.("keydown", (event) => {
  if (event.key !== "Escape" || !activeTaskHelp.size) return;
  event.preventDefault();
  Array.from(activeTaskHelp).forEach((help) => help.dismiss());
});
if (globalThis.MutationObserver && document.body) {
  new MutationObserver(refreshTaskHelp).observe(document.body, {
    childList: true, subtree: true, attributes: true, attributeFilter: ["open"],
  });
}
globalThis.nuLigaTaskHelpTools = {
  initializeTaskHelp, createTaskHelp, taskHelpPlacement, positionTaskHelp, refreshTaskHelp,
};

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
  if (payload.block) renderCakeBlock(card, payload.block);
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
    // Rebuild candidates after sibling changes; retain a private/inactive
    // current occupant which the response intentionally does not list.
    Array.from(select.options).forEach((option) => {
      if (option.value && (Number(option.value) !== selectedId || people.has(selectedId))) {
        option.remove();
      }
    });
    const fragment = document.createDocumentFragment();
    options.forEach((option) => {
      if (Number(option.value) === selectedId) option.selected = true;
      fragment.appendChild(option);
    });
    select.appendChild(fragment);
    if (selectedId) select.value = String(selectedId);
    select.disabled = Boolean(card._mutationPending);
  });
  card._candidateSlots = allowedBySlot;
  if (payload.staffing) updateGameStaffing(card, payload.staffing);
}

function candidateStatus(card, message, retry = false, login = false) {
  const status = card.querySelector("[data-candidate-status]");
  if (!status) return;
  status.hidden = !message;
  status.querySelector("[data-candidate-message]").textContent = message;
  status.querySelector("[data-candidate-retry]").hidden = !retry;
  status.querySelector("[data-candidate-login]").hidden = !login;
}

async function loadCandidateCard(card, refresh = false, afterMutation = false) {
  if (!card?.dataset?.candidateUrl) return;
  if (card._mutationPending && !afterMutation) return;
  if (refresh) card._candidateLoaded = false;
  if (card._candidateLoaded || card._candidateLoading) return;
  card._candidateLoading = true;
  const revision = card._candidateRevision || 0;
  card.querySelectorAll?.("select[data-role], select[data-block-assignment]").forEach((select) => {
    select.disabled = true;
  });
  candidateStatus(card, "Helferliste wird geladen …");
  try {
    const response = await fetch(card.dataset.candidateUrl, { cache: "no-store" });
    if (response.status === 401) {
      throw Object.assign(new Error("Deine Sitzung ist abgelaufen. Bitte melde dich erneut an."), { login: true });
    }
    if (!response.ok) throw new Error("Helferliste konnte nicht geladen werden.");
    const payload = await response.json();
    if (revision !== (card._candidateRevision || 0)) return;
    populateCandidateCard(card, payload);
    card._candidateLoaded = true;
    candidateStatus(card, "");
  } catch (error) {
    if (revision === (card._candidateRevision || 0)) {
      candidateStatus(card, error.message || "Helferliste konnte nicht geladen werden.", true, error.login);
    }
  } finally {
    if (revision === (card._candidateRevision || 0)) card._candidateLoading = false;
  }
}

globalThis.nuLigaCandidateTools = { populateCandidateCard, loadCandidateCard };

document.querySelectorAll("details[data-candidate-url]").forEach((card) => {
  card.addEventListener("toggle", () => {
    if (card.open && card.querySelector("select[data-role], select[data-block-assignment]")) {
      loadCandidateCard(card, true);
    }
  });
  card.querySelector("[data-candidate-retry]").addEventListener("click", () => {
    if (card.querySelector("[data-candidate-message]").textContent.includes("Bitte lade die Seite neu")) {
      window.location.reload();
    } else {
      loadCandidateCard(card);
    }
  });
  if (card.open && card.querySelector("select[data-role], select[data-block-assignment]")) loadCandidateCard(card);
});

function setCoverage(coverage, filled, total) {
  const percent = total ? Math.round(filled * 100 / total) : 0;
  coverage.dataset.progressFilled = String(filled);
  coverage.dataset.progressTotal = String(total);
  const kind = coverage.dataset.progressKind;
  const unit = kind === "game" ? "Pflichtdiensten" : kind === "cake" ? "Kuchen zugesagt" : "Plätzen";
  coverage.querySelector("[data-progress-count]").textContent = `${filled} von ${total} ${unit}${kind === "cake" ? "" : " besetzt"}`;
  coverage.querySelector("[data-progress-percent]").textContent = `${percent} %`;
  coverage.querySelector(".coverage-fill").style.width = `${percent}%`;
  const track = coverage.querySelector('[role="progressbar"]');
  track.setAttribute("aria-valuenow", String(filled));
  track.setAttribute("aria-valuemax", String(total));
}

function updateCoverage(card, previousId, newId) {
  const coverage = card && card.querySelector(".coverage");
  if (!coverage || coverage.dataset.progressKind === "game") return;
  const change = Number(newId !== null) - Number(previousId !== null);
  if (change === 0) return;
  const total = Number(coverage.dataset.progressTotal);
  const filled = Math.max(0, Math.min(total, Number(coverage.dataset.progressFilled) + change));
  setCoverage(coverage, filled, total);
}

function updateGameStaffing(card, staffing) {
  const coverage = card?.querySelector(".coverage");
  if (coverage && coverage.dataset.progressKind === "game") {
    setCoverage(coverage, staffing.required_filled, staffing.required_total);
  }
  updateEligibility(card, staffing.deficiencies);
}

function updateEligibility(card, deficiencies = []) {
  const status = card?.querySelector("[data-eligibility-status]");
  if (!status || status.dataset?.eligibilityPast === "True") return;
  status.replaceChildren();
  deficiencies.forEach((deficiency) => {
    const message = document.createElement("p");
    message.textContent = deficiency.message;
    status.appendChild(message);
  });
  status.hidden = deficiencies.length === 0;
}

// Expose the small, side-effect-free option helpers for the DOM regression test.
globalThis.nuLigaOptionTools = {
  comparePersonOptions, insertOptionSorted, addPersonOption, removePersonOption,
};
globalThis.nuLigaProgressTools = { updateCoverage, setCoverage, updateGameStaffing };

const prevOptions = new WeakMap();
document.querySelectorAll("select[data-role]").forEach((select) => {
  select.addEventListener("focus", () => {
    prevOptions.set(select, select.selectedOptions[0]);
  });
  select.addEventListener("change", async () => {
    const previous = prevOptions.get(select);
    const previousId = previous && previous.value ? Number(previous.value) : null;
    const newId = select.value ? Number(select.value) : null;
    const card = document.getElementById("game-" + select.dataset.game);
    const requestBody = {
      game_id: Number(select.dataset.game), role: select.dataset.role,
      slot: Number(select.dataset.slot), expected_person_id: previousId,
    };
    let savedId = previousId;
    const label = select.closest?.("[data-occupant-id]");
    if (card._mutationPending) {
      select.value = label?.dataset.occupantId || (previousId === null ? "" : String(previousId));
      return;
    }
    setCardMutationPending(card, true);
    try {
      let result = { ok: true };
      if (previousId !== null) {
        result = await postJSON("/api/assignment/release", requestBody);
        if (result.ok) {
          savedId = null;
          select.value = "";
          if (label) label.dataset.occupantId = "";
          if (result.staffing) updateGameStaffing(card, result.staffing);
        }
      }
      if (result.ok && newId !== null) {
        result = await postJSON("/api/assignment/claim", {
          ...requestBody, expected_person_id: null, person_id: newId,
        });
      }
      if (!result.ok) {
        select.value = savedId === null ? "" : String(savedId);
        prevOptions.set(select, select.selectedOptions[0]);
        showToast(result.error || "Fehler beim Speichern", false);
        if (result.loginRequired) setTimeout(() => { window.location.href = "/login"; }, 1200);
        if (result.conflict) {
          setTimeout(() => { window.location.reload(); }, 500);
        } else {
          await loadCandidateCard(card, true, true);
        }
        return;
      }
      select.value = newId === null ? "" : String(newId);
      if (label) label.dataset.occupantId = newId || "";
      flashCard(select.dataset.game);
      // the newly assigned person must not be offered for other tasks
      if (newId) removePersonOption(card, select, newId);
      // a freed person may be offered again for other tasks
      if (previousId !== null && previousId !== newId) {
        addPersonOption(card, select, previous);
      }
      if (result.staffing) updateGameStaffing(card, result.staffing);
      if (select.dataset.releaseOnly === "true" && newId === null) label?.remove();
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
      await loadCandidateCard(card, true, true);
      prevOptions.set(select, select.selectedOptions[0]);
    } finally {
      setCardMutationPending(card, false);
    }
  });
});

function bindBlockSelect(select) {
  prevOptions.set(select, select.selectedOptions[0]);
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
    const card = document.getElementById("block-" + select.dataset.block);
    const label = select.closest?.("[data-occupant-id]");
    if (card._mutationPending) {
      select.value = label?.dataset.occupantId || (previousId === null ? "" : String(previousId));
      return;
    }
    setCardMutationPending(card, true);
    try {
      let savedId = previousId;
      let result = { ok: true };
      if (previousId !== null) {
        result = await postJSON("/api/block-assignment/release", requestBody);
        if (result.ok) {
          savedId = null;
          select.value = "";
          if (label) label.dataset.occupantId = "";
          updateCoverage(card, previousId, null);
        }
      }
      if (result.ok && newId !== null) {
        result = await postJSON("/api/block-assignment/claim", {
          ...requestBody, expected_person_id: null, person_id: newId,
        });
      }
      if (!result.ok) {
        select.value = savedId === null ? "" : String(savedId);
        prevOptions.set(select, select.selectedOptions[0]);
        showToast(result.error || "Fehler beim Speichern", false);
        if (result.loginRequired) setTimeout(() => { window.location.href = "/login"; }, 1200);
        if (result.conflict) {
          card._candidateLoaded = false;
          setTimeout(() => { window.location.reload(); }, 500);
        } else {
          await loadCandidateCard(card, true, true);
        }
        return;
      }
      select.value = newId === null ? "" : String(newId);
      if (label) label.dataset.occupantId = newId || "";
      if (newId) removePersonOption(card, select, newId);
      if (previousId !== null && previousId !== newId) addPersonOption(card, select, previous);
      updateCoverage(card, savedId, newId);
      if (result.block) {
        invalidateCandidateCard(card);
        renderCakeBlock(card, result.block);
      }
      flashBlock(select.dataset.block);
      showToast("Tagesdienst gespeichert", true);
      await loadCandidateCard(card, true, true);
      prevOptions.set(select, select.selectedOptions[0]);
    } finally {
      setCardMutationPending(card, false);
    }
  });
}

document.querySelectorAll("select[data-block-assignment]").forEach(bindBlockSelect);

function invalidateCandidateCard(card) {
  card._candidateRevision = (card._candidateRevision || 0) + 1;
  card._candidateLoaded = false;
  card._candidateLoading = false;
}

function setCardMutationPending(card, pending) {
  if (pending) invalidateCandidateCard(card);
  card._mutationPending = pending;
  card.querySelector("[data-cake-config]")?.querySelectorAll("input, button").forEach((control) => {
    control.disabled = pending;
  });
  card.querySelectorAll("select[data-role], select[data-block-assignment]").forEach((control) => {
    control.disabled = pending || !card._candidateLoaded || card._candidateLoading;
  });
}

function renderCakeBlock(card, block, resetForm = false) {
  if (block.phase !== "cake_delivery") return;
  card.querySelector("[data-cake-time]").textContent = block.delivery_time || "Lieferzeit offen";
  card.classList.toggle("task-block-no-time", !block.delivery_time);
  card.querySelector("[data-cake-quantity]").textContent = block.cake_quantity === null
    ? "Anzahl offen" : `${block.cake_quantity} Kuchen angefragt`;
  const status = card.querySelector("[data-cake-status]");
  status.hidden = block.configured && block.cake_quantity > 0;
  status.textContent = !block.configured ? "Admin-Einrichtung erforderlich"
    : block.cake_quantity === 0 ? "Keine Kuchen angefragt" : "";
  const coverage = card.querySelector(".coverage");
  coverage.hidden = !block.configured || block.cake_quantity === 0;
  setCoverage(coverage, block.progress.filled, block.progress.total);
  const form = card.querySelector("[data-cake-config]");
  const dirty = form && (
    form.querySelector('[name="delivery_time"]').value !== form.dataset.expectedTime
    || form.querySelector('[name="cake_quantity"]').value !== form.dataset.expectedQuantity
  );
  // A roster read may complete while an administrator edits. Preserve both the
  // draft and the settings it was based on so the later save still uses CAS.
  if (form && (resetForm || !dirty)) {
    form.dataset.expectedTime = block.delivery_time || "";
    form.dataset.expectedQuantity = block.cake_quantity === null ? "" : String(block.cake_quantity);
    form.querySelector('[name="delivery_time"]').value = form.dataset.expectedTime;
    form.querySelector('[name="cake_quantity"]').value = form.dataset.expectedQuantity;
  }
  const positions = card.querySelector("[data-cake-slots]");
  const previousFields = Array.from(positions.children);
  const fieldsBySlot = new Map(previousFields.map((field) => [field.dataset.blockSlot, field]));
  const nextFields = block.slots.map((slot) => {
    const description = slot.description || block.description;
    const signature = JSON.stringify([
      slot.label, description, Boolean(slot.editable), slot.person_id || null,
      slot.person_name || "", slot.person_team_label || "",
    ]);
    let field = fieldsBySlot.get(String(slot.slot));
    if (field && !field._cakeSlotSignature) {
      // Recognize the initial server markup without replacing a help control
      // that the viewer may already be reading while candidates are loading.
      const name = field.querySelector("label") || field.querySelector("span");
      const select = field.querySelector("select");
      const occupant = select?.selectedOptions[0] || field.querySelector("strong");
      const occupantText = slot.person_id ? `${slot.person_name} · ${slot.person_team_label}` : "";
      if (name?.textContent === slot.label
          && field.querySelector("[data-task-description]")?.textContent === description
          && Boolean(select) === Boolean(slot.editable)
          && field.dataset.occupantId === (slot.person_id ? String(slot.person_id) : "")
          && (!slot.person_id || occupant?.textContent === occupantText)) {
        field._cakeSlotSignature = signature;
      }
    }
    if (field?._cakeSlotSignature === signature) return field;
    field = createTaskHelp(slot.label, description, slot.editable);
    field.dataset.blockSlot = String(slot.slot);
    field.dataset.occupantId = slot.person_id ? String(slot.person_id) : "";
    field._cakeSlotSignature = signature;
    if (slot.editable) {
      const select = document.createElement("select");
      select.id = field.querySelector("label").getAttribute("for");
      select.dataset.block = String(block.id);
      select.dataset.slot = String(slot.slot);
      select.setAttribute("data-block-assignment", "");
      select.disabled = true;
      const empty = document.createElement("option");
      empty.value = "";
      empty.textContent = "– offen –";
      select.appendChild(empty);
      if (slot.person_id) {
        const occupant = document.createElement("option");
        occupant.value = String(slot.person_id);
        occupant.textContent = `${slot.person_name} · ${slot.person_team_label}`;
        occupant.selected = true;
        select.appendChild(occupant);
      }
      bindBlockSelect(select);
      field.appendChild(select);
    } else {
      const occupant = document.createElement("strong");
      occupant.textContent = slot.person_name
        ? `${slot.person_name} · ${slot.person_team_label}` : "– offen –";
      field.appendChild(occupant);
    }
    return field;
  });
  // Leave unchanged nodes in place: detach/reappend would discard keyboard
  // focus even when a saved roster refresh has identical slot metadata.
  previousFields.forEach((field) => {
    if (!nextFields.includes(field)) field.remove();
  });
  nextFields.forEach((field, index) => {
    if (positions.children[index] !== field) {
      positions.insertBefore(field, positions.children[index] || null);
    }
  });
}

document.querySelectorAll("[data-cake-config]").forEach((form) => {
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const card = form.closest("details");
    if (card._mutationPending) return;
    const rawQuantity = form.querySelector('[name="cake_quantity"]').value;
    const quantity = Number(rawQuantity);
    const message = form.querySelector("[data-cake-settings-message]");
    if (!/^\d+$/.test(rawQuantity) || !Number.isSafeInteger(quantity)) {
      message.textContent = "Die Kuchenanzahl muss eine nichtnegative ganze Zahl sein.";
      message.hidden = false;
      return;
    }
    setCardMutationPending(card, true);
    try {
      const result = await postJSON(form.dataset.settingsUrl, {
        delivery_time: form.querySelector('[name="delivery_time"]').value,
        cake_quantity: quantity,
        expected_delivery_time: form.dataset.expectedTime || null,
        expected_cake_quantity: form.dataset.expectedQuantity === "" ? null : Number(form.dataset.expectedQuantity),
      });
      if (result.block) {
        invalidateCandidateCard(card);
        renderCakeBlock(card, result.block, true);
        if (card.open && result.block.slots.length) await loadCandidateCard(card, true, true);
        else candidateStatus(card, "");
      } else if (card.open && card.querySelector("select[data-block-assignment]")) {
        await loadCandidateCard(card, true, true);
      }
      message.textContent = result.ok ? "Kucheneinstellungen gespeichert." : result.error || "Fehler beim Speichern";
      message.hidden = false;
      if (result.ok) flashBlock(card.dataset.blockCard);
      if (result.loginRequired) setTimeout(() => { window.location.href = "/login"; }, 1200);
    } finally {
      setCardMutationPending(card, false);
    }
  });
});

globalThis.nuLigaCakeTools = { renderCakeBlock, invalidateCandidateCard };

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
