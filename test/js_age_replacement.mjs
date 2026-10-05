// A replacement is two saved mutations: a rejected claim cannot undo its release.
const displays = {
  "[data-progress-count]": { textContent: "2 von 5 Pflichtdiensten besetzt" },
  "[data-progress-percent]": { textContent: "40 %" },
  ".coverage-fill": { style: { width: "40%" } },
  '[role="progressbar"]': { setAttribute() {} },
};
const coverage = {
  dataset: { progressKind: "game", progressFilled: "2", progressTotal: "5" },
  querySelector: (selector) => displays[selector],
};
const options = [{ value: "" }, { value: "7" }, { value: "9" }];
const listeners = {};
const label = { dataset: { occupantId: "7" } };
const select = {
  dataset: { game: "1", role: "Verkauf", slot: "0" }, value: "7",
  addEventListener(name, callback) { listeners[name] = callback; },
  get selectedOptions() { return options.filter((option) => option.value === this.value); },
  closest() { return label; }, classList: { add() {}, remove() {} },
};
const card = {
  classList: { add() {}, remove() {} }, offsetWidth: 1,
  querySelector: (selector) => selector === ".coverage" ? coverage : null,
  querySelectorAll: () => [select],
};
const toast = { classList: { toggle() {}, add() {}, remove() {} } };
globalThis.document = {
  querySelectorAll: (selector) => selector === "select[data-role]" ? [select] : [],
  querySelector: () => ({ content: "test-csrf" }),
  getElementById: (id) => id === "toast" ? toast : card,
};
const requests = [];
globalThis.fetch = async (url) => {
  requests.push(url);
  return { status: 200, json: async () => url.endsWith("/release")
    ? { ok: true, staffing: { required_filled: 1, required_total: 5, deficiencies: [] } }
    : { ok: false, error: "Verkauf benötigt eine erwachsene Person." } };
};
globalThis.setTimeout = () => 0;
globalThis.clearTimeout = () => {};
await import("../static/app.js");
listeners.focus();
select.value = "9";
await listeners.change();
if (requests.join() !== "/api/assignment/release,/api/assignment/claim") {
  throw new Error("replacement did not preserve per-slot release/claim operations");
}
if (select.value !== "" || label.dataset.occupantId !== "" || coverage.dataset.progressFilled !== "1") {
  throw new Error("rejected replacement restored a released occupant or fabricated occupancy");
}
