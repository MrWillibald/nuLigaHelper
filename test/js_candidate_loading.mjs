globalThis.document = {
  querySelectorAll: () => [],
  createElement: () => new FakeOption(),
  createDocumentFragment: () => ({ children: [], appendChild(child) { this.children.push(child); } }),
};
await import("../static/app.js");

function check(condition, message) {
  if (!condition) throw new Error(message);
}

class FakeOption {
  constructor(value = "", label = "– offen –", selected = false) {
    this.value = String(value);
    this.textContent = label;
    this.selected = selected;
    this.dataset = {};
    this.className = "";
    this.title = "";
  }
  cloneNode() {
    const copy = new FakeOption(this.value, this.textContent, this.selected);
    copy.dataset = { ...this.dataset };
    copy.className = this.className;
    copy.title = this.title;
    return copy;
  }
  remove() {
    const index = this.owner.options.indexOf(this);
    if (index >= 0) this.owner.options.splice(index, 1);
  }
  removeAttribute(name) {
    if (name === "selected") this.selected = false;
  }
}

class FakeSelect {
  constructor(role, slot, occupant = null, block = false) {
    this.dataset = block ? { slot: String(slot) } : { role, slot: String(slot) };
    this.block = block;
    this.disabled = true;
    this.options = [new FakeOption("", "– offen –", occupant === null)];
    if (occupant !== null) this.options.push(new FakeOption(occupant, `Current ${occupant}`, true));
    this.options.forEach((option) => { option.owner = this; });
  }
  hasAttribute(name) { return name === "data-block-assignment" && this.block; }
  get selectedOptions() { return this.options.filter((option) => option.selected); }
  get value() { return this.selectedOptions[0]?.value || ""; }
  set value(value) { this.options.forEach((option) => { option.selected = option.value === value; }); }
  appendChild(fragment) {
    fragment.children.forEach((option) => { option.owner = this; this.options.push(option); });
  }
  insertBefore(option, before) {
    const index = before ? this.options.indexOf(before) : this.options.length;
    option.owner = this;
    this.options.splice(index, 0, option);
  }
  querySelector(selector) {
    const value = selector.match(/value="([^"]+)"/)?.[1];
    return this.options.find((option) => option.value === value) || null;
  }
}

class FakeCard {
  constructor(url, selects, occupied) {
    this.dataset = { candidateUrl: url };
    this.selects = selects;
    this.labels = occupied.map((id) => ({ dataset: { occupantId: String(id) } }));
    this.message = { textContent: "" };
    this.retry = { hidden: true, addEventListener(name, callback) { this[name] = callback; } };
    this.login = { hidden: true };
    this.status = {
      hidden: true,
      querySelector: (selector) => ({
        "[data-candidate-message]": this.message,
        "[data-candidate-retry]": this.retry,
        "[data-candidate-login]": this.login,
      })[selector],
    };
    this.listeners = {};
  }
  addEventListener(name, callback) { this.listeners[name] = callback; }
  querySelectorAll(selector) {
    return selector === "[data-occupant-id]" ? this.labels : this.selects;
  }
  querySelector(selector) {
    if (selector === "[data-candidate-status]") return this.status;
    if (selector === "[data-candidate-retry]") return this.retry;
    if (selector === "[data-candidate-message]") return this.message;
    if (selector === "select[data-role], select[data-block-assignment]") return this.selects[0];
    return null;
  }
}

const people = [
  { id: 1, name: "Alex", team_label: "Responsible", sort_group: 1, sort_name: "alex", hint: "" },
  { id: 2, name: "Bea", team_label: "Supporter", sort_group: 2, sort_name: "bea", hint: "" },
  { id: 3, name: "Alex", team_label: "Other", sort_group: 3, sort_name: "alex", hint: "outside" },
  { id: 4, name: "Carl", team_label: "Playing", sort_group: 4, sort_name: "carl", hint: "playing" },
];
const allIds = people.map((person) => person.id);
const payload = {
  people,
  slots: {
    "Zeitnehmer:0": { candidate_ids: allIds, occupant_id: 2 },
    "Sekretär:0": { candidate_ids: allIds, occupant_id: null },
    "Verkauf:0": { candidate_ids: allIds, occupant_id: 99 },
  },
};
const time = new FakeSelect("Zeitnehmer", 0, 2);
const secretary = new FakeSelect("Sekretär", 0);
const sale = new FakeSelect("Verkauf", 0, 99);
const game = new FakeCard("/api/games/1/candidates", [time, secretary, sale], [2, 99]);
const untouched = new FakeCard("/api/games/2/candidates", [new FakeSelect("Zeitnehmer", 0)], []);
let requested = [];
globalThis.fetch = async (url) => {
  requested.push(url);
  return { ok: true, status: 200, json: async () => payload };
};
await globalThis.nuLigaCandidateTools.loadCandidateCard(game);
check(requested.join() === game.dataset.candidateUrl, "opening a card fetched another card");
check(game._candidateLoaded && !time.disabled && !secretary.disabled && !sale.disabled, "successful load did not enable controls");
check(untouched.selects[0].disabled, "unopened card was populated");
check(time.options.map((option) => option.value).join() === ",1,2,3,4", "game category order changed");
check(secretary.options.map((option) => option.value).join() === ",1,3,4", "assigned helper remained in sibling slot");
check(sale.value === "99" && sale.querySelector('option[value="99"]').selected, "inactive occupant was lost");
check(secretary.querySelector('option[value="3"]').className === "foreign-option", "outside hint missing");
check(secretary.querySelector('option[value="4"]').textContent.includes("spielt selbst"), "playing hint missing");
check(secretary.querySelector('option[value="1"]').textContent === "Alex · Responsible", "duplicate names lost their team label");
globalThis.nuLigaOptionTools.addPersonOption(game, time, time.querySelector('option[value="2"]'));
check(secretary.options.map((option) => option.value).join() === ",1,2,3,4", "released person did not return in order");
check(sale.querySelector('option[value="2"]'), "released person did not return to another eligible slot");
globalThis.nuLigaOptionTools.removePersonOption(game, secretary, 3);
check(!time.querySelector('option[value="3"]'), "newly claimed person remained in sibling slot");

const blockOne = new FakeSelect("", 0, 2, true);
const blockTwo = new FakeSelect("", 1, null, true);
const block = new FakeCard("/api/blocks/1/candidates", [blockOne, blockTwo], [2]);
globalThis.nuLigaCandidateTools.populateCandidateCard(block, {
  people: people.slice(1, 3).map((person) => ({ ...person, sort_group: 0 })),
  slots: {
    "0": { candidate_ids: [2, 3], occupant_id: 2 },
    "1": { candidate_ids: [2, 3], occupant_id: null },
  },
});
check(blockTwo.options.map((option) => option.value).join() === ",3", "block occupant remained in sibling slot");
check(time.querySelector('option[value="2"]'), "block loading changed a separate game's options");

const failed = new FakeCard("/api/games/3/candidates", [new FakeSelect("Zeitnehmer", 0)], []);
globalThis.fetch = async () => { throw new Error("offline"); };
await globalThis.nuLigaCandidateTools.loadCandidateCard(failed);
check(failed.selects[0].disabled && failed.retry.hidden === false, "failed load left an enabled control or no retry");
globalThis.fetch = async () => ({ ok: true, status: 200, json: async () => ({
  people: people.slice(0, 1),
  slots: { "Zeitnehmer:0": { candidate_ids: [1], occupant_id: null } },
}) });
await globalThis.nuLigaCandidateTools.loadCandidateCard(failed);
check(!failed.selects[0].disabled && failed.status.hidden, "retry did not recover the card");

const expired = new FakeCard("/api/games/4/candidates", [new FakeSelect("Zeitnehmer", 0)], []);
globalThis.fetch = async () => ({ status: 401 });
await globalThis.nuLigaCandidateTools.loadCandidateCard(expired);
check(expired.login.hidden === false && expired.selects[0].disabled, "expired session did not show sign-in action");
