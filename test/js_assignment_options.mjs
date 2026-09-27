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

function coverageCard(kind, filled, total) {
  const count = { textContent: "" };
  const percent = { textContent: "" };
  const fill = { style: { width: "" } };
  const bar = { attributes: {}, setAttribute(name, value) { this.attributes[name] = value; } };
  const elements = {
    "[data-progress-count]": count,
    "[data-progress-percent]": percent,
    ".coverage-fill": fill,
    '[role="progressbar"]': bar,
  };
  const coverage = {
    dataset: { progressKind: kind, progressFilled: String(filled), progressTotal: String(total) },
    querySelector: (selector) => elements[selector],
  };
  return { card: { querySelector: () => coverage }, coverage, count, percent, fill, bar };
}

const gameProgress = coverageCard("game", 3, 5);
const update = globalThis.nuLigaProgressTools.updateCoverage;
update(gameProgress.card, null, 7, "Zeitnehmer");
if (gameProgress.count.textContent !== "4 von 5 Pflichtdiensten besetzt"
    || gameProgress.percent.textContent !== "80 %"
    || gameProgress.fill.style.width !== "80%"
    || gameProgress.bar.attributes["aria-valuenow"] !== "4") {
  throw new Error("successful game claim did not update all progress views");
}
update(gameProgress.card, 7, 9, "Zeitnehmer");
update(gameProgress.card, null, 10, "Unterstützung");
if (gameProgress.coverage.dataset.progressFilled !== "4") {
  throw new Error("replacement or optional support changed required progress");
}
update(gameProgress.card, 9, null, "Zeitnehmer");
if (gameProgress.coverage.dataset.progressFilled !== "3") {
  throw new Error("game release did not reduce required progress");
}
const blockProgress = coverageCard("block", 1, 3);
update(blockProgress.card, null, 7);
if (blockProgress.count.textContent !== "2 von 3 Plätzen besetzt"
    || blockProgress.percent.textContent !== "67 %") {
  throw new Error("block claim did not update three-slot progress");
}
update(blockProgress.card, 7, null);
if (blockProgress.coverage.dataset.progressFilled !== "1") {
  throw new Error("block release did not reduce progress");
}
