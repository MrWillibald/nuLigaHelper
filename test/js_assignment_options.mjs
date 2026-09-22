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
}

class FakeSelect {
  constructor(options) { this.options = options; }
  querySelector(selector) {
    const match = selector.match(/value="([^"]+)"/);
    return match ? this.options.find((option) => option.value === match[1]) : null;
  }
  insertBefore(option, before) {
    const index = before ? this.options.indexOf(before) : this.options.length;
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
