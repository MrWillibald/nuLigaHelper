import { installConfirmation } from "./js_confirmation_fixture.mjs";
const feedback = [];
globalThis.nuLigaFeedback = {
  show: (value) => feedback.push(value),
  navigate: (values) => { feedback.push(...values); return true; },
};
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
const empty = { value: "", dataset: {}, classList: { contains: () => false } };
const candidate = { value: "7", dataset: {}, classList: { contains: () => false } };
const listeners = {};
const assignmentSelect = {
  dataset: { game: "1", role: "Zeitnehmer", slot: "0" },
  value: "",
  addEventListener(name, callback) { listeners[name] = callback; },
  get selectedOptions() { return [this.value ? candidate : empty]; },
  classList: { add() {}, remove() {}, toggle() {} },
};
const card = {
  classList: { add() {}, remove() {}, toggle() {} },
  offsetWidth: 1,
  querySelector: (selector) => selector === ".coverage" ? coverage : null,
  querySelectorAll: (selector) => selector === "[data-request-replacement]" ? [] : [assignmentSelect],
};
const toast = { classList: { toggle() {}, add() {}, remove() {} } };
globalThis.document = {
  querySelectorAll: (selector) => selector === "select[data-role]" ? [assignmentSelect] : [],
  querySelector: () => ({ content: "test-csrf" }),
  getElementById: (id) => id === "toast" ? toast : card,
};
installConfirmation(document, { autoConfirm: true });
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

// Adult Reinigung positions and Kasse receive saved staffing, without deriving
// requiredness from their role text or from the locally selected value.
for (const [role, slot, total, filled] of [
  ["Kasse", "0", 8, 4],
  ["Reinigung", "0", 8, 5],
  ["Reinigung", "1", 8, 6],
]) {
  assignmentSelect.value = "";
  assignmentSelect.dataset.role = role;
  assignmentSelect.dataset.slot = slot;
  listeners.focus();
  assignmentSelect.value = "7";
  globalThis.fetch = async () => ({ status: 200, json: async () => ({
    ok: true,
    staffing: { required_filled: filled, required_total: total, deficiencies: [] },
  }) });
  await listeners.change();
  if (assignmentSelect.value !== "7"
      || coverage.dataset.progressFilled !== String(filled)
      || coverage.dataset.progressTotal !== String(total)) {
    throw new Error(`${role}:${slot} inferred progress instead of using saved staffing`);
  }
}

listeners.focus();
assignmentSelect.value = "";
globalThis.fetch = async () => ({ status: 409, json: async () => ({ ok: false, error: "Stale" }) });
await listeners.change();
if (assignmentSelect.value !== "7" || coverage.dataset.progressFilled !== "6") {
  throw new Error("stale adult cleaning release changed saved progress");
}

listeners.focus();
assignmentSelect.value = "";
globalThis.fetch = async () => { throw new Error("an offered occupied duty must never be released"); };
await listeners.change();
if (assignmentSelect.value !== "7" || coverage.dataset.progressFilled !== "6") {
  throw new Error("selecting the empty option withdrew a binding duty");
}

// A retained youth duty can only be released and disappears after saving.
let removed = false;
const retainedLabel = { dataset: { occupantId: "7" }, remove() { removed = true; } };
assignmentSelect.closest = () => retainedLabel;
assignmentSelect.dataset.releaseOnly = "true";
assignmentSelect.dataset.role = "Kasse";
assignmentSelect.value = "7";
listeners.focus();
assignmentSelect.value = "";
globalThis.fetch = async () => ({ status: 200, json: async () => ({
  ok: true, staffing: { required_filled: 5, required_total: 5, deficiencies: [] },
}) });
await listeners.change();
if (!removed || coverage.dataset.progressFilled !== "5" || coverage.dataset.progressTotal !== "5") {
  throw new Error("retained youth duty release failed to remove its field or changed baseline coverage");
}

if (feedback.at(-1).severity !== "success" || !feedback.at(-1).message.includes("freigegeben")) {
  throw new Error("saved release was mislabeled as an assignment");
}
assignmentSelect.dataset.releaseOnly = "false";
listeners.focus();
assignmentSelect.value = "7";
globalThis.fetch = async () => ({ status: 200, json: async () => ({
  ok: true, warning: "Person spielt selbst.",
  staffing: { required_filled: 5, required_total: 5, deficiencies: [] },
}) });
await listeners.change();
if (feedback.at(-1).severity !== "warning" || !feedback.at(-1).message.includes("gespeichert")
    || assignmentSelect.value !== "7") throw new Error("authoritative advisory did not confirm the saved assignment");
