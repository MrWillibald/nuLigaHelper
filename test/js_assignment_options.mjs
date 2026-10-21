globalThis.document = { querySelectorAll: () => [] };
await import("../static/app.js");

class FakeOption {
  constructor(value, group, name, id, label, className = "", title = "") {
    this.value = String(value);
    this.dataset = { sortGroup: String(group), sortName: name, sortId: String(id) };
    this.label = label;
    this.className = className;
    this.title = title;
    this.selected = false;
  }
  cloneNode() {
    const clone = new FakeOption(
      this.value, this.dataset.sortGroup, this.dataset.sortName,
      this.dataset.sortId, this.label, this.className, this.title,
    );
    clone.selected = this.selected;
    return clone;
  }
  removeAttribute(name) {
    if (name === "selected") this.selected = false;
  }
  remove() {
    if (!this.owner) return;
    const index = this.owner.options.indexOf(this);
    if (index >= 0) this.owner.options.splice(index, 1);
  }
}

class FakeSelect {
  constructor(options) {
    this.options = options;
    this.options.forEach((option) => { option.owner = this; });
  }
  querySelector(selector) {
    const match = selector.match(/value="([^"]+)"/);
    return match ? this.options.find((option) => option.value === match[1]) : null;
  }
  insertBefore(option, before) {
    const index = before ? this.options.indexOf(before) : this.options.length;
    option.owner = this;
    this.options.splice(index, 0, option);
  }
}

const placeholder = () => new FakeOption("", 0, "", 0, "– offen –");
const current = new FakeSelect([placeholder()]);
const first = new FakeSelect([
  placeholder(),
  new FakeOption(2, 1, "anna", 2, "Anna · Alpha"),
  new FakeOption(9, 3, "zoe", 9, "Zoe · ohne Team"),
]);
const second = new FakeSelect([placeholder()]);
const card = { querySelectorAll: () => [current, first, second] };
const released = new FakeOption(
  7, 2, "mira", 7, "Mira · Alpha, Beta • spielt selbst",
  "option-playing", "Mira spielt in diesem Spiel selbst",
);
released.selected = true;

globalThis.nuLigaOptionTools.addPersonOption(card, current, released);
for (const select of [first, second]) {
  const clone = select.querySelector('option[value="7"]');
  if (!clone) throw new Error("released person was not reinserted");
  if (clone.dataset.sortGroup !== "2" || clone.dataset.sortName !== "mira") {
    throw new Error("server-provided sort metadata was not preserved");
  }
  if (clone.label !== released.label || clone.className !== "option-playing" || clone.title !== released.title) {
    throw new Error("membership label or warning metadata was not preserved");
  }
  if (clone.selected) throw new Error("reinserted option remained selected");
}
if (first.options.map((option) => option.value).join(",") !== ",2,7,9") {
  throw new Error("released person was inserted at the wrong position");
}

// A block is one independent assignment container: claiming in one card removes
// the person only from sibling slots, not from another block or a game card.
const claimed = new FakeSelect([placeholder(), released.cloneNode()]);
const sibling = new FakeSelect([placeholder(), released.cloneNode()]);
const otherContainer = new FakeSelect([placeholder(), released.cloneNode()]);
const blockCard = { querySelectorAll: () => [claimed, sibling] };
globalThis.nuLigaOptionTools.removePersonOption(blockCard, claimed, 7);
if (sibling.querySelector('option[value="7"]')) {
  throw new Error("claimed block person remained available in a sibling slot");
}
if (!otherContainer.querySelector('option[value="7"]')) {
  throw new Error("block claim leaked into another assignment container");
}

const { progressFixture } = await import("./js_progress_fixture.mjs");
const { setCoverage, updateGameStaffing, updateBlockSummary } = globalThis.nuLigaProgressTools;
globalThis.document.createElement = () => ({ textContent: "" });
const game = progressFixture();
updateGameStaffing(game.card, { required_filled: 4, required_total: 5, vacancies: { Verkauf: 1 }, deficiencies: [] });
if (game.count.textContent !== "4/5" || game.arc.attrs["stroke-dashoffset"] !== "20"
    || game.gauge.attrs["aria-valuetext"] !== "4 von 5 Pflichtdiensten besetzt"
    || game.vacancies.children[0].textContent !== "Verkauf") {
  throw new Error("saved game summary lost its fraction, arc, semantics or vacancy");
}
// A remote helper also claimed a position: use the saved block, not a local +1.
for (const kind of ["preparation", "cleanup"]) {
  const block = progressFixture("block", 0, 3);
  updateBlockSummary(block.card, { phase: kind, progress: { filled: 2, total: 3 }, vacancies: { Tagesdienst: 1 } });
  if (block.count.textContent !== "2/3" || block.vacancies.children[0].textContent !== "Tagesdienst"
      || block.gauge.attrs["aria-valuetext"] !== "2 von 3 Plätzen besetzt") {
    throw new Error("saved block summary ignored concurrent occupancy");
  }
}
const adult = progressFixture();
setCoverage(adult.coverage, 7, 7);
if (adult.count.textContent !== "7/7" || adult.arc.attrs["stroke-dashoffset"] !== "0") {
  throw new Error("adult ring lost its seven-position denominator");
}
