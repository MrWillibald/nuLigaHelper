import assert from "node:assert/strict";
import { assignmentFixture } from "./js_assignment_fixture.mjs";

const page = assignmentFixture();
await import("../static/app.js");
const tools = globalThis.nuLigaAssignmentTools;
const requests = [];
globalThis.fetch = async (url, options) => {
  requests.push({ url, options, body: JSON.parse(options.body) });
  return { status: 200, json: async () => ({ ok: true, assignment: {
    person_id: 9, person_name: "Replacement <b>plain</b>", person_team_label: "Team A",
    locked: true, replacement_token: null, seeking_replacement: false,
  } }) };
};
const cases = [
  ...["Zeitnehmer", "Sekretär", "Verkauf", "Ordnungsdienst", "Kasse", "Reinigung"].map((role) => ({ kind: "game", role })),
  ...["Vorbereitung", "Kuchenlieferung", "Aufräumen"].map((kind) => ({ kind })),
];

for (const settings of cases) {
  for (const dismissal of ["cancel", "escape", "outside", "close"]) {
    const { select, field } = page.createDuty(settings);
    tools.bindAssignmentSelect(select); select.focus(); select.value = "9";
    const before = requests.length;
    const pending = select.dispatch("change");
    assert.equal(page.dialog.open, true, "selection must open confirmation before a write");
    assert.equal(select.value, "", "an unconfirmed selection must not replace saved occupancy");
    assert.equal(requests.length, before, "selection and dialog display must not send a write");
    assert.equal(document.activeElement, page.cancel, "initial focus must allow keyboard cancellation");
    assert.ok(page.context.textContent.includes("Replacement <b>plain</b>"), "confirmation must name the intended helper as text");
    assert.ok(page.context.textContent.includes(settings.role || settings.kind), "confirmation must identify the duty");
    assert.ok(page.context.textContent.includes("28.11.2026"), "confirmation must identify the date");
    if (dismissal === "cancel") page.cancel.dispatch("click");
    else if (dismissal === "escape") page.dialog.dispatch("cancel");
    else if (dismissal === "outside") page.dialog.dispatch("click", { clientX: 10, clientY: 10 });
    else page.dialog.close();
    await pending;
    assert.equal(requests.length, before, `${dismissal} must not submit a write or compensating release`);
    assert.equal(field.dataset.occupantId, ""); assert.equal(select.value, "");
    assert.equal(document.activeElement, select, "cancellation must return focus to the initiating control");
    assert.equal(page.dialog.open, false);
  }
  const { select, field, status } = page.createDuty({ ...settings, late: true });
  tools.bindAssignmentSelect(select); select.focus(); select.value = "9";
  const before = requests.length;
  const pending = select.dispatch("change");
  assert.equal(page.late.hidden, false, "late free-duty claims must explicitly explain the immediate lock");
  assert.equal(requests.length, before);
  page.confirm.dispatch("click"); page.confirm.dispatch("click"); await pending;
  assert.equal(requests.length, before + 1, "repeated confirmation must submit exactly once");
  const request = requests.at(-1);
  assert.equal(request.url, settings.kind === "game" ? "/api/assignment/claim" : "/api/block-assignment/claim");
  assert.equal(request.body.expected_person_id, null, "confirmation must keep the original empty-slot expectation");
  assert.equal(request.body.person_id, 9);
  assert.equal(request.options.headers["X-CSRF-Token"], "synthetic-csrf");
  assert.equal(field.dataset.occupantId, "9");
  assert.equal(select.hidden, true, "confirmed late assignments must expose no editing control");
  assert.match(status.textContent, /Fest eingeteilt/);
}

for (const settings of cases) {
  const { select, field, strong } = page.createDuty(settings);
  tools.bindAssignmentSelect(select); select.focus(); select.value = "9";
  const before = requests.length;
  const pending = select.dispatch("change");
  field.dataset.occupantId = "11"; strong.textContent = "Concurrent Winner"; select.value = "11";
  page.cancel.dispatch("click"); await pending;
  assert.equal(field.dataset.occupantId, "11");
  assert.equal(select.value, "11", "Cancel must preserve a concurrent winner");
  assert.equal(requests.length, before);
}

// Confirming after a remote winner uses the original expectation and conflicts.
for (const settings of cases) {
  const { select, field } = page.createDuty(settings);
  tools.bindAssignmentSelect(select); select.focus(); select.value = "9";
  const pending = select.dispatch("change");
  field.dataset.occupantId = "11"; select.value = "11";
  globalThis.fetch = async (url, options) => {
    requests.push({ url, options, body: JSON.parse(options.body) });
    return { status: 409, json: async () => ({ ok: false, error: "Zwischenzeitlich belegt.", current_person_id: 11 }) };
  };
  page.confirm.dispatch("click"); await pending;
  assert.equal(requests.at(-1).body.expected_person_id, null);
  assert.equal(field.dataset.occupantId, "11"); assert.equal(select.value, "11");
  assert.equal(page.feedback.at(-1).severity, "error");
}

const fixed = page.createDuty({ occupant: 7, token: "closed", locked: true });
tools.bindAssignmentSelect(fixed.select); tools.bindReplacementRequest(fixed.button);
tools.applyDutyState(fixed.field, { person_id: 7, locked: true, seeking_replacement: false });
const beforeFixed = requests.length;
fixed.select.focus(); fixed.select.value = "9";
await fixed.select.dispatch("change"); await fixed.button.dispatch("click");
assert.equal(requests.length, beforeFixed);
assert.equal(fixed.select.value, "7");
assert.equal(fixed.select.hidden, true); assert.equal(fixed.button.hidden, true);

// A newly filled empty field gains its replacement request action without reload.
for (const settings of cases) {
  const duty = page.createDuty(settings);
  tools.bindReplacementRequest(duty.button);
  tools.applyDutyState(duty.field, { person_id: null, locked: false, can_request_replacement: false });
  assert.equal(duty.button.hidden, true);
  tools.applyDutyState(duty.field, { person_id: 7, person_name: "Original Helper", locked: false, can_request_replacement: true });
  assert.equal(duty.button.hidden, false, "a saved claim must expose its authorized replacement action");
  const before = requests.length;
  globalThis.fetch = async (url, options) => {
    requests.push({ url, body: JSON.parse(options.body) });
    return { status: 200, json: async () => ({ ok: true, assignment: {
      person_id: 7, person_name: "Original Helper", replacement_token: "new-search", seeking_replacement: true,
      locked: false, can_request_replacement: true,
    } }) };
  };
  await duty.button.dispatch("click");
  assert.equal(requests.length, before + 1);
  assert.equal(requests.at(-1).url, settings.kind === "game"
    ? "/api/assignment/request-replacement" : "/api/block-assignment/request-replacement");
  assert.equal(requests.at(-1).body.expected_person_id, 7);
  assert.equal(duty.field.dataset.occupantId, "7", "opening a search must retain its responsible helper");
  assert.match(duty.status.textContent, /Ersatz gesucht/);
  assert.match(duty.status.textContent, /Original Helper bleibt/);
  assert.equal(duty.select.hidden, false);
  assert.equal(duty.button.hidden, false, "the assigned helper can end the newly opened search");
  assert.equal(duty.button.textContent, "Ersatzsuche beenden");
}
