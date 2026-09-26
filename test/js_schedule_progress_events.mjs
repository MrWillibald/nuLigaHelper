// A rejected claim must leave the card and its progress at the saved value.
const elements = {
  "[data-progress-count]": { textContent: "" },
  "[data-progress-percent]": { textContent: "" },
  ".coverage-fill": { style: { width: "60%" } },
  '[role="progressbar"]': { setAttribute() {} },
};
const coverage = {
  dataset: { progressKind: "game", progressFilled: "3", progressTotal: "5" },
  querySelector: (selector) => elements[selector],
};
const empty = { value: "", classList: { contains: () => false } };
const candidate = { value: "7", classList: { contains: () => false } };
const listeners = {};
const assignmentSelect = {
  dataset: { game: "1", role: "Zeitnehmer", slot: "0" },
  value: "",
  addEventListener(name, callback) { listeners[name] = callback; },
  get selectedOptions() { return [this.value ? candidate : empty]; },
  classList: { add() {}, remove() {} },
};
const card = {
  classList: { add() {}, remove() {} },
  offsetWidth: 1,
  querySelector: () => coverage,
  querySelectorAll: () => [assignmentSelect],
};
const toast = { classList: { toggle() {}, add() {}, remove() {} } };
globalThis.document = {
  querySelectorAll: (selector) => selector === "select[data-role]" ? [assignmentSelect] : [],
  querySelector: () => ({ content: "test-csrf" }),
  getElementById: (id) => id === "toast" ? toast : card,
};
globalThis.fetch = async () => ({ status: 409, json: async () => ({ ok: false, error: "Stale" }) });
globalThis.setTimeout = () => 0;
globalThis.clearTimeout = () => {};
await import("../static/app.js");
listeners.focus();
assignmentSelect.value = "7";
await listeners.change();
if (assignmentSelect.value !== "" || coverage.dataset.progressFilled !== "3") {
  throw new Error("rejected claim changed the selected person or displayed progress");
}
