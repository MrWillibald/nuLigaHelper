import assert from "node:assert/strict";
import { progressFixture } from "./js_progress_fixture.mjs";
import { assignmentFixture } from "./js_assignment_fixture.mjs";

const page = assignmentFixture();
await import("../../nuligahelper/web/static/app.js");
const tools = globalThis.nuLigaAssignmentTools;
const cases = ["game", "Vorbereitung", "Kuchenlieferung", "Aufräumen"];

for (const kind of cases) {
  const duty = page.createDuty({ kind, occupant: 7, token: "current-search" });
  const summary = progressFixture(kind === "game" ? "game" : "block", 1, 3);
  const query = duty.card.querySelector.bind(duty.card);
  duty.card.querySelector = selector => selector === ".coverage" ? summary.coverage : query(selector);
  const savedSummary = kind === "game" ? {
    staffing: { required_filled: 1, required_total: 3, vacancies: { Verkauf: 2 }, deficiencies: [] },
  } : { block: { phase: kind, progress: { filled: 1, total: 3 }, vacancies: { [kind]: 2 } } };
  if (savedSummary.staffing) globalThis.nuLigaProgressTools.updateGameStaffing(duty.card, savedSummary.staffing);
  else globalThis.nuLigaProgressTools.updateBlockSummary(duty.card, savedSummary.block);
  const beforeSummary = summary.snapshot();
  tools.bindReplacementRequest(duty.button);
  tools.applyDutyState(duty.field, { person_id: 7, seeking_replacement: true });
  assert.equal(duty.button.hidden, false);
  assert.equal(duty.button.textContent, "Ersatzsuche beenden");
  duty.button.focus();
  const requests = [];
  let finish;
  globalThis.fetch = (url, options) => {
    requests.push({ url, options, body: JSON.parse(options.body) });
    return new Promise(resolve => { finish = resolve; });
  };
  const pending = duty.button.dispatch("click");
  await duty.button.dispatch("click");
  assert.equal(requests.length, 1, "pending cancellation must submit only once");
  assert.equal(requests[0].url, kind === "game"
    ? "/api/assignment/cancel-replacement" : "/api/block-assignment/cancel-replacement");
  assert.equal(requests[0].body.expected_person_id, 7);
  assert.equal(requests[0].body.expected_request_token, "current-search");
  assert.equal(requests[0].options.headers["X-CSRF-Token"], "synthetic-csrf");
  assert.equal(duty.field.dataset.occupantId, "7", "pending closure cannot create a vacancy");
  finish({ status: 200, json: async () => ({ ok: true, ...savedSummary, assignment: {
    person_id: 7, person_name: "Original Helper", replacement_token: null,
    seeking_replacement: false, locked: false, can_request_replacement: true,
  } }) });
  await pending;
  assert.equal(duty.field.dataset.occupantId, "7");
  assert.equal(duty.field.dataset.replacementToken, "");
  assert.equal(summary.snapshot(), beforeSummary, "replacement closure changed physical occupancy or bar");
  assert.equal(duty.select.hidden, true);
  assert.equal(duty.button.textContent, "Ersatz suchen");
  assert.equal(document.activeElement, duty.button);
  assert.match(page.feedback.at(-1).message, /Ersatzsuche beendet/);

  const opening = duty.button.dispatch("click");
  finish({ status: 200, json: async () => ({ ok: true, ...savedSummary, assignment: {
    person_id: 7, replacement_token: "reopened", seeking_replacement: true,
    locked: false, can_request_replacement: true,
  } }) });
  await opening;
  assert.equal(summary.snapshot(), beforeSummary, "replacement search created a vacancy");
  tools.applyDutyState(duty.field, { person_id: 7, replacement_token: "reopened", seeking_replacement: true });
  globalThis.fetch = async (url, options) => {
    requests.push({ url, body: JSON.parse(options.body) });
    duty.field.dataset.occupantId = "11";
    duty.field.dataset.replacementToken = "winner-search";
    return { status: 409, json: async () => ({ ok: false, error: "Ersatzsuche wurde geändert." }) };
  };
  await duty.button.dispatch("click");
  assert.equal(requests.at(-1).body.expected_request_token, "reopened");
  assert.equal(duty.field.dataset.occupantId, "11", "stale cancellation must preserve the current helper");
  assert.equal(duty.field.dataset.replacementToken, "winner-search");
  assert.equal(page.feedback.at(-1).severity, "error");

  const before = requests.length;
  for (const state of [
    { locked: true, can_request_replacement: true, replacement_token: "fixed" },
    { locked: false, can_request_replacement: false, replacement_token: "foreign" },
    { locked: false, can_request_replacement: true, replacement_token: null },
  ]) {
    tools.applyDutyState(duty.field, { person_id: 11, seeking_replacement: true, ...state });
    await duty.button.dispatch("click");
  }
  assert.equal(requests.length, before, "locked, foreign or missing-generation cancellation cannot submit");
}
