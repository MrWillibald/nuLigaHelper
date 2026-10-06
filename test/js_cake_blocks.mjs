const feedback = [];
globalThis.nuLigaFeedback = {
  show: (value) => feedback.push(value),
  navigate: (values) => { feedback.push(...values); return true; },
};
// Exercise the cake card's saved configuration and dynamically created controls.
function check(condition, message) {
  if (!condition) throw new Error(message);
}

class Element {
  constructor(tag = "div") {
    this.tag = tag;
    this.children = [];
    this.dataset = {};
    this.style = {};
    this.listeners = {};
    this.attrs = {};
    this.textContent = "";
    this.hidden = false;
    this.disabled = false;
    this._value = "";
    const classes = new Set();
    this.classList = {
      add: (value) => classes.add(value), remove: (value) => classes.delete(value),
      contains: (value) => classes.has(value),
      toggle: (value, enabled) => enabled ? classes.add(value) : classes.delete(value),
    };
  }
  get options() { return this.children; }
  get selectedOptions() { return this.children.filter((option) => option.value === this.value); }
  get value() {
    return this.tag === "select" ? this.children.find((option) => option.selected)?.value || "" : this._value;
  }
  set value(value) {
    this._value = value;
    if (this.tag === "select") this.children.forEach((option) => { option.selected = option.value === value; });
  }
  addEventListener(name, callback) { this.listeners[name] = callback; }
  get parentElement() { return this.parent; }
  get tagName() { return this.tag.toUpperCase(); }
  get isConnected() { return this === document.body || Boolean(this.parent?.isConnected); }
  getBoundingClientRect() {
    return { top: 100, bottom: 120, left: 100, right: 120,
      width: this.tag === "p" ? 300 : 20, height: this.tag === "p" ? 120 : 20 };
  }
  contains(element) { return this === element || this.children.some((child) => child.contains(element)); }
  showPopover() { this.popoverOpen = true; }
  hidePopover() { this.popoverOpen = false; }
  focus() {
    document.activeElement?.listeners.blur?.();
    document.activeElement = this;
    this.listeners.focus?.();
  }
  blur() {
    if (document.activeElement !== this) return;
    document.activeElement = null;
    this.listeners.blur?.();
  }
  setAttribute(name, value) { this.attrs[name] = value; }
  getAttribute(name) { return this.attrs[name] ?? null; }
  hasAttribute(name) { return name in this.attrs; }
  removeAttribute(name) { delete this.attrs[name]; if (name === "selected") this.selected = false; }
  appendChild(child) {
    if (child.tag === "fragment") [...child.children].forEach((item) => this.appendChild(item));
    else {
      if (child.parent) child.remove();
      child.parent = this;
      this.children.push(child);
    }
  }
  replaceChildren(...children) {
    this.children.forEach((child) => { child.parent = null; });
    this.children = [];
    children.forEach((child) => this.appendChild(child));
  }
  remove() {
    if (this.parent) this.parent.children = this.parent.children.filter((child) => child !== this);
    this.parent = null;
  }
  insertBefore(child, before) {
    if (child === before) return;
    if (child.parent) child.remove();
    child.parent = this;
    this.children.splice(before ? this.children.indexOf(before) : this.children.length, 0, child);
  }
  closest(selector) {
    let current = this;
    while (current) {
      if (current.tag === selector) return current;
      current = current.parent;
    }
    return null;
  }
  cloneNode() {
    const child = new Element(this.tag);
    child.value = this.value;
    child.dataset = { ...this.dataset };
    child.textContent = this.textContent;
    return child;
  }
  querySelector(selector) {
    const expected = selector.match(/option\[value="([^"]+)"\]/)?.[1];
    if (expected) return this.children.find((child) => child.value === expected) || null;
    if (this.elements?.[selector]) return this.elements[selector];
    const attribute = selector.match(/^\[([^\]]+)\]$/)?.[1];
    for (const child of this.children) {
      if (attribute ? child.hasAttribute(attribute) : child.tag === selector) return child;
      const descendant = child.querySelector(selector);
      if (descendant) return descendant;
    }
    return null;
  }
  querySelectorAll() { return []; }
}

const coverage = new Element();
coverage.dataset = { progressKind: "cake", progressFilled: "0", progressTotal: "0" };
coverage.elements = Object.fromEntries([
  "[data-progress-count]", "[data-progress-percent]", ".coverage-fill", '[role="progressbar"]',
].map((selector) => [selector, new Element()]));
const positions = new Element();
const time = new Element("input");
const quantity = new Element("input");
const message = new Element();
const button = new Element("button");
const form = new Element("form");
form.dataset = { expectedTime: "", expectedQuantity: "", settingsUrl: "/api/blocks/1/cake-settings" };
form.elements = { '[name="delivery_time"]': time, '[name="cake_quantity"]': quantity,
  "[data-cake-settings-message]": message };
form.querySelectorAll = () => [time, quantity, button];
const candidateMessage = new Element();
const retry = new Element("button");
const login = new Element();
const candidateStatus = new Element();
candidateStatus.elements = { "[data-candidate-message]": candidateMessage,
  "[data-candidate-retry]": retry, "[data-candidate-login]": login };
const card = new Element("details");
card.open = false;
card.appendChild(form);
card.appendChild(positions);
card.dataset = { blockCard: "1", candidateUrl: "/api/blocks/1/candidates" };
card.elements = {
  "[data-cake-time]": new Element(), "[data-cake-quantity]": new Element(),
  "[data-cake-status]": new Element(), ".coverage": coverage,
  "[data-cake-config]": form, "[data-cake-slots]": positions,
  "[data-candidate-status]": candidateStatus, "[data-candidate-retry]": retry,
  "[data-candidate-message]": candidateMessage,
};
const originalQuery = card.querySelector.bind(card);
card.querySelectorAll = (selector) => selector === "[data-occupant-id]" ? positions.children
  : positions.children.flatMap((label) => label.children.filter((child) => child.tag === "select"));
card.querySelector = (selector) => ["select[data-role], select[data-block-assignment]", "select[data-block-assignment]"].includes(selector)
  ? card.querySelectorAll(selector)[0] || null : originalQuery(selector);
const toast = new Element();
globalThis.document = {
  activeElement: null,
  listeners: {},
  documentElement: { clientWidth: 360, clientHeight: 640 },
  addEventListener(name, callback) { this.listeners[name] = callback; },
  querySelectorAll: (selector) => selector === "details[data-candidate-url]" ? [card]
    : selector === "[data-cake-config]" ? [form] : [],
  querySelector: () => ({ content: "synthetic-csrf" }),
  getElementById: (id) => id === "toast" ? toast : card,
  createElement: (tag) => new Element(tag),
  createDocumentFragment: () => new Element("fragment"),
};
document.body = new Element("body");
document.body.appendChild(card);
globalThis.window = { location: { reload() {} }, innerWidth: 360, innerHeight: 640,
  addEventListener() {} };
globalThis.setTimeout = () => 0;
globalThis.clearTimeout = () => {};
await import("../static/app.js");
const person = { id: 7, name: "Cake Helper", team_label: "Supporter", sort_group: 0,
  sort_name: "cake helper", hint: "" };
const cakeDescription = 'Zur angezeigten Zeit einen Kuchen bringen. <script>plain text</script> & Kuchen';
function savedBlock(count, occupant = null) {
  return {
    id: 1, phase: "cake_delivery", configured: count !== null,
    description: cakeDescription,
    cake_quantity: count, delivery_time: count === null ? null : "10:00",
    progress: { filled: occupant ? 1 : 0, total: count || 0 },
    slots: Array.from({ length: count || 0 }, (_, slot) => ({
      slot, label: `Kuchenlieferung ${slot + 1}`, editable: true,
      person_id: slot === 0 ? occupant : null,
      person_name: slot === 0 && occupant ? person.name : "",
      person_team_label: slot === 0 && occupant ? person.team_label : "",
    })),
  };
}
function checkTaskHelpForSlots(count) {
  const descriptionIds = new Set();
  check(positions.children.length === count, "saved cake help lost assignment positions");
  for (const [slot, field] of positions.children.entries()) {
    const label = field.querySelector("label");
    const select = field.children.find((child) => child.tag === "select");
    const help = field.querySelector("[data-task-help]");
    const description = field.querySelector("[data-task-description]");
    check(label.textContent === `Kuchenlieferung ${slot + 1}` && label.getAttribute("for") === select.id,
      "dynamic cake task label lost its select association");
    check(help.type === "button" && help.getAttribute("aria-label") === `Informationen zu ${label.textContent}`,
      "dynamic cake help is not a task-named non-submit control");
    check(help.getAttribute("aria-controls") === description.id
      && help.getAttribute("aria-describedby") === description.id && description.hidden,
      "dynamic cake help lost its description association or starts expanded");
    check(description.textContent === cakeDescription && !description.children.length,
      "dynamic cake descriptions diverged or interpreted description text as HTML");
    check(!descriptionIds.has(description.id), "dynamic cake controls reuse description identifiers");
    descriptionIds.add(description.id);
  }
}
let current = savedBlock(null);
const requests = [];
let configurationReply;
let candidateFailure = false;
let lastConfigurationBody;
globalThis.fetch = async (url, options) => {
  requests.push(url);
  if (url.endsWith("/cake-settings")) {
    const body = JSON.parse(options.body);
    lastConfigurationBody = body;
    check(options.headers["X-CSRF-Token"] === "synthetic-csrf", "settings omitted CSRF");
    if (configurationReply) return configurationReply;
    current = savedBlock(body.cake_quantity);
    return { status: 200, json: async () => ({ ok: true, block: current }) };
  }
  if (url.endsWith("/claim")) {
    current = savedBlock(current.cake_quantity, 7);
    return { status: 200, json: async () => ({ ok: true, block: current }) };
  }
  if (candidateFailure) throw new Error("offline");
  return { ok: true, status: 200, json: async () => ({
    block: current, people: [person], slots: Object.fromEntries(current.slots.map((slot) => [
      String(slot.slot), { occupant_id: slot.person_id, candidate_ids: [7] },
    ])),
  }) };
};

check(!requests.length, "unconfigured collapsed cake card fetched a roster");
time.value = "10:00";
quantity.value = "4";
await form.listeners.submit({ preventDefault() {} });
check(!card.open && positions.children.length === 4, "saved setup expanded card or lost positions");
checkTaskHelpForSlots(4);
check(requests.join() === form.dataset.settingsUrl, "collapsed setup loaded candidates");
check(coverage.dataset.progressTotal === "4" && coverage.dataset.progressFilled === "0", "saved quantity did not refresh progress");
check(card.querySelectorAll("").every((select) => select.disabled), "new positions enabled without candidates");
card.open = true;
await globalThis.nuLigaCandidateTools.loadCandidateCard(card, true);
let select = card.querySelectorAll("")[0];
check(!select.disabled && select.options.some((option) => option.value === "7"), "expanded cake did not load candidates");

// A candidate response may arrive while a helper reads the information control.
// Keep the same field/button when saved slot metadata has not changed, so focus
// and persistent help are not destroyed by the asynchronous roster refresh.
const focusedField = positions.children[0];
// Server-rendered fields have no private reconciliation signature yet.
delete focusedField._cakeSlotSignature;
const focusedHelp = focusedField.querySelector("[data-task-help]");
const focusedDescription = focusedField.querySelector("[data-task-description]");
focusedHelp.focus();
focusedHelp.listeners.click({ preventDefault() {}, stopPropagation() {} });
const ordinaryCandidateFetch = globalThis.fetch;
let finishHelpCandidates;
globalThis.fetch = async (url) => {
  requests.push(url);
  return { ok: true, status: 200, json: () => new Promise((resolve) => { finishHelpCandidates = resolve; }) };
};
const pendingHelpCandidates = globalThis.nuLigaCandidateTools.loadCandidateCard(card, true);
await Promise.resolve();
check(document.activeElement === focusedHelp && !focusedDescription.hidden,
  "starting a candidate read removed focused cake guidance");
finishHelpCandidates({ block: current, people: [person], slots: Object.fromEntries(current.slots.map((slot) => [
  String(slot.slot), { occupant_id: slot.person_id, candidate_ids: [7] },
])) });
await pendingHelpCandidates;
check(positions.children[0] === focusedField
  && positions.children[0].querySelector("[data-task-help]") === focusedHelp,
  "unchanged candidate response replaced the focused cake information control");
check(document.activeElement === focusedHelp && !focusedDescription.hidden
  && focusedHelp.getAttribute("aria-expanded") === "true",
  "candidate response lost keyboard position or persistent task guidance");
focusedField.listeners.keydown({ key: "Escape", preventDefault() {}, stopPropagation() {} });
focusedHelp.blur();
globalThis.fetch = ordinaryCandidateFetch;
select = card.querySelectorAll("")[0];
select.value = "7";
candidateFailure = true;
await select.listeners.change();
check(feedback.at(-1).severity === "success" && feedback.at(-1).message.includes("gespeichert"),
  "candidate refresh failure reclassified the confirmed save");
check(retry.hidden === false && card.querySelectorAll("").every((control) => control.disabled),
  "confirmed save with failed candidates enabled incomplete pickers or lost recovery");
candidateFailure = false;
await globalThis.nuLigaCandidateTools.loadCandidateCard(card, true);
check(coverage.dataset.progressFilled === "1" && positions.children[0].dataset.occupantId === "7", "cake claim did not show saved progress and occupant");
check(card.querySelectorAll("")[1].options.every((option) => option.value !== "7"), "cake volunteer still offered for another cake");

configurationReply = { status: 400, json: async () => ({ ok: false, error: "Belegte Plätze freigeben." }) };
quantity.value = "0";
await form.listeners.submit({ preventDefault() {} });
check(coverage.dataset.progressFilled === "1" && coverage.dataset.progressTotal === "4"
  && positions.children.length === 4, "refused reduction fabricated unsaved progress or removed positions");
check(!card.querySelectorAll("")[0].disabled, "refused configuration left valid assignment controls disabled");
current = savedBlock(5, 7);
configurationReply = { status: 409, json: async () => ({ ok: false, error: "Veraltet", block: current }) };
await form.listeners.submit({ preventDefault() {} });
check(quantity.value === "5" && positions.children.length === 5 && coverage.dataset.progressFilled === "1", "stale config did not show current saved values");
checkTaskHelpForSlots(5);
check(card.open, "saved configuration changed expansion state");

// Candidate reads and assignment responses must preserve an unsaved draft and
// the original settings it expects, including when a delayed GET finishes.
quantity.value = "7";
time.value = "11:00";
const originalExpected = { time: form.dataset.expectedTime, quantity: form.dataset.expectedQuantity };
current = savedBlock(6, 7);
await globalThis.nuLigaCandidateTools.loadCandidateCard(card, true);
check(quantity.value === "7" && time.value === "11:00", "candidate refresh overwrote an unsaved settings draft");
check(form.dataset.expectedQuantity === originalExpected.quantity
  && form.dataset.expectedTime === originalExpected.time, "candidate refresh moved a draft's CAS expectations");
globalThis.nuLigaCakeTools.renderCakeBlock(card, savedBlock(8, 7));
check(quantity.value === "7" && form.dataset.expectedQuantity === "5", "assignment refresh moved unsaved settings expectations");
configurationReply = { status: 409, json: async () => ({ ok: false, error: "Veraltet", block: current }) };
await form.listeners.submit({ preventDefault() {} });
check(lastConfigurationBody.expected_cake_quantity === 5
  && lastConfigurationBody.expected_delivery_time === "10:00", "unsaved draft bypassed stale config rejection");
check(quantity.value === "6" && time.value === "10:00"
  && form.dataset.expectedQuantity === "6", "explicit stale refusal did not reset fields to saved state");

let completeDelayedRead;
const ordinaryFetch = globalThis.fetch;
globalThis.fetch = async (url) => ({ ok: true, status: 200, json: () => new Promise((resolve) => {
  completeDelayedRead = resolve;
}) });
const delayedRead = globalThis.nuLigaCandidateTools.loadCandidateCard(card, true);
await Promise.resolve();
quantity.value = "9";
time.value = "12:00";
const delayedBlock = savedBlock(10, 7);
completeDelayedRead({ block: delayedBlock, people: [person], slots: Object.fromEntries(delayedBlock.slots.map((slot) => [
  String(slot.slot), { occupant_id: slot.person_id, candidate_ids: [7] },
])) });
await delayedRead;
check(quantity.value === "9" && time.value === "12:00" && form.dataset.expectedQuantity === "6",
  "a delayed candidate response overwrote a draft or advanced its stale-check baseline");
globalThis.fetch = ordinaryFetch;
current = delayedBlock;

// Configuration and slot writes in one card share a guard. Pending settings
// disable claims, and a pending release prevents another config POST.
let finishSettings;
configurationReply = { status: 200, json: () => new Promise((resolve) => { finishSettings = resolve; }) };
const beforeGuard = requests.length;
const pendingSettings = form.listeners.submit({ preventDefault() {} });
await Promise.resolve();
await Promise.resolve();
check(card._mutationPending && time.disabled && quantity.disabled
  && card.querySelectorAll("").every((control) => control.disabled), "pending configuration did not lock both kinds of controls");
select = card.querySelectorAll("")[0];
select.value = "";
await select.listeners.change();
await form.listeners.submit({ preventDefault() {} });
check(requests.length === beforeGuard + 1 && select.value === "7", "overlapping claim/configuration wrote during the same card mutation");
current = savedBlock(9, 7);
finishSettings({ ok: true, block: current });
await pendingSettings;
check(!card._mutationPending && !time.disabled && !card.querySelectorAll("")[0].disabled, "saved configuration failed to unlock refreshed controls");

let finishRelease;
globalThis.fetch = async (url, options) => {
  if (!url.endsWith("/release")) return ordinaryFetch(url, options);
  requests.push(url);
  return { status: 200, json: () => new Promise((resolve) => { finishRelease = resolve; }) };
};
select = card.querySelectorAll("")[0];
select.value = "";
const pendingRelease = select.listeners.change();
await Promise.resolve();
await Promise.resolve();
const pendingCount = requests.length;
check(card._mutationPending && button.disabled && requests.at(-1).endsWith("/release"),
  "changing a saved occupant without focus failed to send its release or lock configuration");
await form.listeners.submit({ preventDefault() {} });
check(requests.length === pendingCount, "configuration ran while the slot release was pending");
current = savedBlock(9);
finishRelease({ ok: true, block: current });
await pendingRelease;
check(!card._mutationPending && coverage.dataset.progressFilled === "0", "saved release failed to refresh/unlock the card");
globalThis.fetch = ordinaryFetch;

// A candidate GET started before a mutation must not replace its newer saved
// response when it eventually resolves.
configurationReply = undefined;
let finishOldCandidates;
let firstCandidateRead = true;
globalThis.fetch = async (url, options) => {
  if (url.endsWith("/candidates") && firstCandidateRead) {
    firstCandidateRead = false;
    return { ok: true, status: 200, json: () => new Promise((resolve) => { finishOldCandidates = resolve; }) };
  }
  return ordinaryFetch(url, options);
};
const oldCandidateRead = globalThis.nuLigaCandidateTools.loadCandidateCard(card, true);
await Promise.resolve();
quantity.value = "4";
await form.listeners.submit({ preventDefault() {} });
const oldBlock = savedBlock(12, 7);
finishOldCandidates({ block: oldBlock, people: [person], slots: Object.fromEntries(oldBlock.slots.map((slot) => [
  String(slot.slot), { occupant_id: slot.person_id, candidate_ids: [7] },
])) });
await oldCandidateRead;
check(coverage.dataset.progressTotal === "4" && coverage.dataset.progressFilled === "0"
  && positions.children.length === 4, "delayed pre-mutation candidates replaced newer saved cake progress");
globalThis.fetch = ordinaryFetch;

candidateFailure = true;
await globalThis.nuLigaCandidateTools.loadCandidateCard(card, true);
check(card.querySelectorAll("").every((control) => control.disabled), "candidate failure enabled incomplete cake controls");
check(!retry.hidden, "cake candidate failure omitted retry");
const removedHelp = positions.children[3].querySelector("[data-task-help]");
const removedDescription = positions.children[3].querySelector("[data-task-description]");
removedHelp.listeners.click({ preventDefault() {}, stopPropagation() {}, pointerType: "touch" });
check(!removedDescription.hidden && removedDescription.popoverOpen,
  "dynamic cake help did not open as a floating popover");
card.open = false;
globalThis.nuLigaTaskHelpTools.refreshTaskHelp();
check(removedDescription.hidden && !removedDescription.popoverOpen
  && removedHelp.getAttribute("aria-expanded") === "false",
  "closing a cake card left task guidance visible over other page content");
card.open = true;
removedHelp.listeners.click({ preventDefault() {}, stopPropagation() {}, pointerType: "touch" });
globalThis.nuLigaCakeTools.renderCakeBlock(card, savedBlock(0));
globalThis.nuLigaTaskHelpTools.refreshTaskHelp();
check(removedDescription.hidden && !removedDescription.popoverOpen
  && removedHelp.getAttribute("aria-expanded") === "false",
  "removing a configured cake position left its floating task guidance visible");
check(positions.children.length === 0 && coverage.hidden, "zero cakes left positions or a visible percentage");
check(card.elements["[data-cake-status]"].textContent === "Keine Kuchen angefragt", "zero cake status confused with setup");
check(!coverage.elements["[data-progress-percent]"].textContent.includes("NaN"), "zero cake progress divided by zero");
globalThis.nuLigaCakeTools.renderCakeBlock(card, savedBlock(null));
check(coverage.hidden && card.elements["[data-cake-status]"].textContent.includes("Einrichtung"), "unset cake status appears fully staffed");

check(feedback.some((value) => value.severity === "success" && value.message === "Kucheneinstellungen gespeichert."),
  "cake configuration has no shared success result");
check(feedback.some((value) => value.severity === "error"), "cake refusals have no shared error result");
check(!message.textContent.includes("gespeichert"), "cake result duplicated its transient confirmation inline");
