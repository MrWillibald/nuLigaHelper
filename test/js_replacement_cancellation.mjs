import assert from "node:assert/strict";
import { assignmentFixture } from "./js_assignment_fixture.mjs";

const page = assignmentFixture();
await import("../static/app.js");
const tools = globalThis.nuLigaAssignmentTools;
const cases = ["game", "Vorbereitung", "Kuchenlieferung", "Aufräumen"];

for (const kind of cases) {
  const duty = page.createDuty({ kind, occupant: 7, token: "current-search" });
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
  finish({ status: 200, json: async () => ({ ok: true, assignment: {
    person_id: 7, person_name: "Original Helper", replacement_token: null,
    seeking_replacement: false, locked: false, can_request_replacement: true,
  } }) });
  await pending;
  assert.equal(duty.field.dataset.occupantId, "7");
  assert.equal(duty.field.dataset.replacementToken, "");
  assert.equal(duty.select.hidden, true);
  assert.equal(duty.button.textContent, "Ersatz suchen");
  assert.equal(document.activeElement, duty.button);
  assert.match(page.feedback.at(-1).message, /Ersatzsuche beendet/);

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
