/* Shared plain-text action results. No form or API payload belongs here. */
(() => {
  const region = document.getElementById("feedback");
  if (!region) return;
  const labels = { success: "Erfolg", info: "Information", warning: "Warnung", error: "Fehler" };
  const key = "nuliga-action-feedback";
  const entries = new Set();
  let sequence = 0;
  // Persistent empty live regions are registered before their text changes.
  // Visible entries themselves have no live role, so each outcome is announced once.
  const announcers = {};
  ["polite", "assertive"].forEach((channel) => {
    const node = document.createElement("div");
    node.className = "feedback-announcer";
    node.setAttribute("role", channel === "assertive" ? "alert" : "status");
    node.setAttribute("aria-live", channel);
    node.setAttribute("aria-atomic", "true");
    document.body.appendChild(node);
    announcers[channel] = { node, pending: [], scheduled: false };
  });
  function announce(message, severity) {
    const channel = announcers[severity === "error" ? "assertive" : "polite"];
    channel.pending.push(message);
    if (channel.scheduled) return;
    channel.scheduled = true;
    channel.node.textContent = "";
    setTimeout(() => {
      channel.node.textContent = channel.pending.join(" ");
      channel.pending = [];
      channel.scheduled = false;
    }, 0);
  }
  const now = () => Date.now();
  const destination = () => window.location.pathname + window.location.search;
  const normalize = (value) => {
    if (!value || typeof value.message !== "string" || !value.message.trim()
        || value.message.length > 2000 || !Object.hasOwn(labels, value.severity)
        || (value.action !== undefined && !["refresh", "signin"].includes(value.action))) return null;
    return { message: value.message, severity: value.severity,
      ...(value.action ? { action: value.action } : {}) };
  };
  function clearTransfer() {
    try { sessionStorage.removeItem(key); } catch (_) { /* Storage is optional. */ }
  }
  function placement() {
    const visual = window.visualViewport;
    const width = visual?.width || window.innerWidth;
    const height = visual?.height || window.innerHeight;
    region.style.left = `${(visual?.offsetLeft || 0) + width / 2}px`;
    region.style.top = `${(visual?.offsetTop || 0) + height - 12}px`;
    region.style.width = `${Math.max(0, Math.min(560, width - 24))}px`;
    region.style.maxHeight = `${Math.max(0, height * 0.4)}px`;
    globalThis.nuLigaTaskHelpTools?.refreshTaskHelp();
  }
  function visible(entry) {
    if (document.hidden) return false;
    const rect = entry.getBoundingClientRect();
    const bounds = region.getBoundingClientRect();
    return rect.top >= bounds.top - 1 && rect.bottom <= bounds.bottom + 1;
  }
  function show(value) {
    const descriptor = normalize(value);
    if (!descriptor) return;
    const entry = document.createElement("div");
    entry.className = `feedback-entry feedback-${descriptor.severity}`;
    entry.dataset.severity = descriptor.severity;
    entry.setAttribute("role", "group");
    const text = document.createElement("span");
    text.textContent = `${labels[descriptor.severity]}: ${descriptor.message}`;
    entry.appendChild(text);
    announce(text.textContent, descriptor.severity);
    if (descriptor.action) {
      const action = document.createElement("a");
      action.textContent = descriptor.action === "signin" ? "Erneut anmelden" : "Seite neu laden";
      action.href = descriptor.action === "signin" ? "/login" : destination();
      action.addEventListener("click", () => clearTransfer());
      entry.appendChild(action);
    }
    const state = { entry, remaining: 5000, last: now(), hovered: false, focused: false };
    entries.add(state);
    const remove = () => {
      entries.delete(state);
      entry.remove();
      if (!entries.size && typeof region.hidePopover === "function") region.hidePopover();
      placement();
    };
    state.remove = remove;
    const pause = (field, value) => { state[field] = value; state.last = now(); };
    entry.addEventListener("pointerenter", () => pause("hovered", true));
    entry.addEventListener("pointerleave", () => pause("hovered", false));
    entry.addEventListener("focusin", () => pause("focused", true));
    entry.addEventListener("focusout", (event) => {
      if (!entry.contains(event.relatedTarget)) pause("focused", false);
    });
    region.appendChild(entry);
    if (typeof region.showPopover === "function") {
      // Re-enter the top layer after task help, without closing or focusing it.
      if (region.matches(":popover-open")) region.hidePopover();
      region.showPopover();
    }
    placement();
    region.scrollTop = region.scrollHeight;
    return entry;
  }
  function tick() {
    const timestamp = now();
    entries.forEach((state) => {
      if (!state.hovered && !state.focused && visible(state.entry)) {
        state.remaining -= Math.max(0, timestamp - state.last);
        if (state.remaining <= 0) state.remove();
      }
      state.last = timestamp;
    });
  }
  function navigate(values, path = destination()) {
    const descriptors = values.map(normalize);
    const url = new URL(path, window.location.href);
    if (url.origin !== window.location.origin || descriptors.some((value) => !value)
        || !descriptors.length || descriptors.length > 5) return false;
    const target = url.pathname + url.search;
    try {
      sessionStorage.setItem(key, JSON.stringify({ destination: target, expires: now() + 60000,
        entries: descriptors.map((value) => ({ id: `${now()}-${++sequence}`, ...value })) }));
    } catch (_) {
      descriptors.forEach((value) => show({ ...value, action: target === "/login" ? "signin" : "refresh" }));
      return false;
    }
    if (target === destination()) window.location.reload();
    else window.location.href = target;
    return true;
  }
  function consume() {
    let raw;
    try { raw = sessionStorage.getItem(key); sessionStorage.removeItem(key); } catch (_) { return []; }
    if (!raw || raw.length > 12000) return [];
    try {
      const envelope = JSON.parse(raw);
      if (envelope.destination !== destination() || !Number.isFinite(envelope.expires)
          || envelope.expires <= now() || envelope.expires > now() + 60000
          || !Array.isArray(envelope.entries) || !envelope.entries.length || envelope.entries.length > 5) return [];
      const ids = envelope.entries.map((value) => value.id);
      if (ids.some((id) => typeof id !== "string" || id.length > 100) || new Set(ids).size !== ids.length) return [];
      const values = envelope.entries.map(normalize);
      return values.every(Boolean) ? values : [];
    } catch (_) { return []; }
  }
  const server = Array.from(region.querySelectorAll("[data-feedback-message]")).map((entry) => ({
    message: entry.querySelector("[data-feedback-text]").textContent,
    severity: entry.dataset.severity,
  }));
  // Login/logout/verification must never replay a previous account's transfer.
  const transferred = region.dataset.authBoundary === "true" ? (clearTransfer(), []) : consume();
  region.replaceChildren();
  region.classList.add("feedback-enhanced");
  region.setAttribute("tabindex", "-1");
  if (typeof region.showPopover === "function") region.setAttribute("popover", "manual");
  [...server, ...transferred].forEach(show);
  document.addEventListener("visibilitychange", () => entries.forEach((state) => { state.last = now(); }));
  document.querySelectorAll('form[action="/logout"]').forEach((form) => form.addEventListener("submit", clearTransfer));
  window.addEventListener("pagehide", () => {
    // A restored history document must not replay consumed results. Keep only a
    // newly prepared navigation envelope for its immediate destination.
    entries.clear();
    region.replaceChildren();
    Object.values(announcers).forEach((channel) => { channel.node.textContent = ""; channel.pending = []; });
    if (typeof region.hidePopover === "function" && region.matches(":popover-open")) region.hidePopover();
  });
  window.addEventListener("resize", placement);
  window.visualViewport?.addEventListener("resize", placement);
  window.visualViewport?.addEventListener("scroll", placement);
  globalThis.nuLigaFeedback = { show, navigate, clearTransfer, placement, tick };
  setInterval(tick, 100);
})();
