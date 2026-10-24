import assert from "node:assert/strict";
import { progressFixture } from "./js_progress_fixture.mjs";
import { installConfirmation } from "./js_confirmation_fixture.mjs";
const feedback = [];
globalThis.nuLigaFeedback = {
  show: (value) => feedback.push(value),
  navigate: (values) => { feedback.push(...values); return true; },
};
// A rejected claim must leave the card and its progress at the saved value.
const summary = progressFixture();
const { coverage } = summary;
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
  createElement: () => ({ textContent: "" }),
  getElementById: (id) => id === "toast" ? toast : card,
};
installConfirmation(document, { autoConfirm: true });
globalThis.fetch = async () => ({ status: 409, json: async () => ({ ok: false, error: "Stale" }) });
globalThis.setTimeout = () => 0;
globalThis.clearTimeout = () => {};
await import("../static/app.js");
globalThis.nuLigaProgressTools.updateGameStaffing(card, { required_filled: 3, required_total: 5,
  vacancies: { Verkauf: 1, Ordnungsdienst: 1 }, deficiencies: [] });
const beforeRefusal = summary.snapshot();
listeners.focus();
assignmentSelect.value = "7";
await listeners.change();
if (assignmentSelect.value !== "" || coverage.dataset.progressFilled !== "3") {
  throw new Error("rejected claim changed the selected person or displayed progress");
}

assert.equal(summary.snapshot(), beforeRefusal, "refused selection changed bar, percentage or semantics");

// Adult Wischer and Kasse receive saved staffing, without deriving
// requiredness from their role text or from the locally selected value.
for (const [role, slot, total, filled] of [
  ["Kasse", "0", 7, 4],
  ["Wischer", "0", 7, 5],
]) {
  assignmentSelect.value = "";
  assignmentSelect.dataset.role = role;
  assignmentSelect.dataset.slot = slot;
  listeners.focus();
  assignmentSelect.value = "7";
  globalThis.fetch = async () => ({ status: 200, json: async () => ({
    ok: true,
    staffing: { required_filled: filled, required_total: total, vacancies: {}, deficiencies: [] },
  }) });
  await listeners.change();
  assert.equal(summary.count.textContent, `${filled} von ${total} Pflichtdiensten besetzt`);
  assert.equal(summary.fill.style.width, `${Math.round(filled * 100 / total)}%`);
  assert.equal(summary.percent.textContent, `${Math.round(filled * 100 / total)} %`);
  assert.equal(summary.track.attrs["aria-valuetext"], `${filled} von ${total} Pflichtdiensten besetzt`);
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
if (assignmentSelect.value !== "7" || coverage.dataset.progressFilled !== "5") {
  throw new Error("stale adult cleaning release changed saved progress");
}

listeners.focus();
assignmentSelect.value = "";
globalThis.fetch = async () => { throw new Error("an offered occupied duty must never be released"); };
await listeners.change();
if (assignmentSelect.value !== "7" || coverage.dataset.progressFilled !== "5") {
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
  ok: true, staffing: { required_filled: 5, required_total: 5, vacancies: {}, deficiencies: [] },
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
  staffing: { required_filled: 5, required_total: 5, vacancies: {}, deficiencies: [] },
}) });
await listeners.change();
if (feedback.at(-1).severity !== "warning" || !feedback.at(-1).message.includes("gespeichert")
    || assignmentSelect.value !== "7") throw new Error("authoritative advisory did not confirm the saved assignment");

// Candidate responses and writes share exactly one public summary renderer.
const tools = globalThis.nuLigaProgressTools;
tools.updateGameStaffing(card, { required_filled: 3, required_total: 7,
  vacancies: { Wischer: 1, Verkauf: 2, Kasse: 1 }, deficiencies: [] });
assert.equal(summary.count.textContent, "3 von 7 Pflichtdiensten besetzt");
assert.equal(summary.percent.textContent, "43 %");
card.querySelectorAll = () => [];
globalThis.nuLigaCandidateTools.populateCandidateCard(card, { people: [], slots: {}, staffing: {
  required_filled: 6, required_total: 7, vacancies: { Wischer: 1 }, deficiencies: [],
} });
assert.equal(summary.count.textContent, "6 von 7 Pflichtdiensten besetzt");
assert.equal(summary.fill.style.width, "86%");
assert.equal(summary.percent.textContent, "86 %");
for (const filled of [0, 7]) {
  tools.updateGameStaffing(card, { required_filled: filled, required_total: 7,
    vacancies: filled ? {} : { Verkauf: 2 }, deficiencies: [] });
  assert.equal(summary.fill.style.width, filled ? "100%" : "0%");
  assert.equal(summary.percent.textContent, filled ? "100 %" : "0 %");
}
// A pre-write candidate read cannot roll back the saved count and bar.
card.dataset = { candidateUrl: "/api/games/1/candidates" };
let finishOld;
globalThis.fetch = async () => ({ ok: true, status: 200,
  json: () => new Promise(resolve => { finishOld = resolve; }) });
const pending = globalThis.nuLigaCandidateTools.loadCandidateCard(card, true);
await Promise.resolve();
card._candidateRevision = (card._candidateRevision || 0) + 1;
const saved = summary.snapshot();
finishOld({ people: [], slots: {}, staffing: { required_filled: 0, required_total: 7,
  vacancies: { Verkauf: 2 }, deficiencies: [] } });
await pending;
assert.equal(summary.snapshot(), saved, "old candidates overwrote a newer saved game summary");

// Preparation/cleanup writes and candidate reads must use the authoritative
// block occupancy, including a second helper saved concurrently.
delete card.dataset.candidateUrl;
coverage.dataset.progressKind = "block";
assignmentSelect.dataset.block = "1";
retainedLabel.dataset.occupantId = "";
assignmentSelect.value = "";
listeners.focus();
assignmentSelect.value = "7";
const savedBlock = { phase: "preparation", progress: { filled: 2, total: 3 }, vacancies: { Vorbereitung: 1 } };
globalThis.fetch = async () => ({ status: 200, json: async () => ({ ok: true, block: savedBlock }) });
await listeners.change();
assert.equal(summary.count.textContent, "2 von 3 Plätzen besetzt", "local +1 lost a concurrent saved block claim");
assert.equal(summary.fill.style.width, "67%");
globalThis.nuLigaCandidateTools.populateCandidateCard(card, { people: [], slots: {}, block: {
  phase: "cleanup", progress: { filled: 3, total: 3 }, vacancies: {},
} });
assert.equal(summary.count.textContent, "3 von 3 Plätzen besetzt");
assert.equal(summary.percent.textContent, "100 %");
