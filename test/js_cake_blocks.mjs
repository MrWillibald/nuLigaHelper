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
  setAttribute(name, value) { this.attrs[name] = value; }
  hasAttribute(name) { return name in this.attrs; }
  removeAttribute(name) { delete this.attrs[name]; if (name === "selected") this.selected = false; }
  appendChild(child) {
    if (child.tag === "fragment") child.children.forEach((item) => this.appendChild(item));
    else { child.parent = this; this.children.push(child); }
  }
  replaceChildren() { this.children = []; }
  remove() { this.parent.children = this.parent.children.filter((child) => child !== this); }
  insertBefore(child, before) {
    child.parent = this;
    this.children.splice(before ? this.children.indexOf(before) : this.children.length, 0, child);
  }
  closest() { return this.tag === "form" ? card : this.parent; }
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
    return this.elements?.[selector] || null;
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
  querySelectorAll: (selector) => selector === "details[data-candidate-url]" ? [card]
    : selector === "[data-cake-config]" ? [form] : [],
  querySelector: () => ({ content: "synthetic-csrf" }),
  getElementById: (id) => id === "toast" ? toast : card,
  createElement: (tag) => new Element(tag),
  createDocumentFragment: () => new Element("fragment"),
};
globalThis.window = { location: { reload() {} } };
globalThis.setTimeout = () => 0;
globalThis.clearTimeout = () => {};
await import("../static/app.js");
const person = { id: 7, name: "Cake Helper", team_label: "Supporter", sort_group: 0,
  sort_name: "cake helper", hint: "" };
function savedBlock(count, occupant = null) {
  return {
    id: 1, phase: "cake_delivery", configured: count !== null,
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
check(requests.join() === form.dataset.settingsUrl, "collapsed setup loaded candidates");
check(coverage.dataset.progressTotal === "4" && coverage.dataset.progressFilled === "0", "saved quantity did not refresh progress");
check(card.querySelectorAll("").every((select) => select.disabled), "new positions enabled without candidates");
card.open = true;
await globalThis.nuLigaCandidateTools.loadCandidateCard(card, true);
let select = card.querySelectorAll("")[0];
check(!select.disabled && select.options.some((option) => option.value === "7"), "expanded cake did not load candidates");
select.value = "7";
await select.listeners.change();
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
globalThis.nuLigaCakeTools.renderCakeBlock(card, savedBlock(0));
check(positions.children.length === 0 && coverage.hidden, "zero cakes left positions or a visible percentage");
check(card.elements["[data-cake-status]"].textContent === "Keine Kuchen angefragt", "zero cake status confused with setup");
check(!coverage.elements["[data-progress-percent]"].textContent.includes("NaN"), "zero cake progress divided by zero");
globalThis.nuLigaCakeTools.renderCakeBlock(card, savedBlock(null));
check(coverage.hidden && card.elements["[data-cake-status]"].textContent.includes("Einrichtung"), "unset cake status appears fully staffed");
