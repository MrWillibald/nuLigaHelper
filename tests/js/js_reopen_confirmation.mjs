import assert from "node:assert/strict";
import { assignmentFixture } from "./js_assignment_fixture.mjs";
import { installConfirmation } from "./js_confirmation_fixture.mjs";

const page = assignmentFixture();
const dialog = installConfirmation(document, { id: "reopen-confirmation" });
await import("../../nuligahelper/web/static/app.js");
const tools = globalThis.nuLigaAssignmentTools;

function reopenable(kind) {
  const duty = page.createDuty({ kind, occupant: 7, locked: true });
  duty.field.dataset.canReopen = "true";
  duty.button.setAttribute("data-reopen", "");
  duty.button.removeAttribute("data-request-replacement");
  tools.bindReopen(duty.button);
  return duty;
}

const requests = [];
globalThis.fetch = async (url, options) => {
  requests.push({ url, body: JSON.parse(options.body) });
  return { status: 200, json: async () => ({ ok: true, notified: requests.length === 1 }) };
};

for (const [kind, url] of [["game", "/api/assignment/reopen"], ["Aufräumen", "/api/block-assignment/reopen"]]) {
  const duty = reopenable(kind);
  const before = requests.length;

  for (const dismiss of [() => dialog.cancel.dispatch("click"), () => dialog.dialog.dispatch("cancel")]) {
    const pending = duty.button.dispatch("click");
    await Promise.resolve();
    assert.equal(dialog.dialog.open, true, "reopening opens its confirmation, also on a locked duty");
    assert.match(dialog.context.textContent, /Original Helper verliert/);
    dismiss();
    await pending;
    assert.equal(requests.length, before, "cancel and Escape must not send a request");
    assert.equal(document.activeElement, duty.button, "focus returns to the control");
    assert.equal(duty.field.dataset.occupantId, "7");
  }

  const confirmed = duty.button.dispatch("click");
  await Promise.resolve();
  dialog.confirm.dispatch("click");
  await confirmed;
  assert.equal(requests.length, before + 1);
  assert.equal(requests.at(-1).url, url);
  assert.equal(requests.at(-1).body.expected_person_id, 7);
}

assert.equal(page.feedback.at(-2).severity, "success");
assert.equal(page.feedback.at(-1).severity, "warning", "an unreached helper leaves a warning for the admin");
assert.match(page.feedback.at(-1).message, /persönlich/);

const hidden = reopenable("game");
tools.applyDutyState(hidden.field, { person_id: null, can_reopen: false });
const count = requests.length;
await hidden.button.dispatch("click");
assert.equal(requests.length, count, "an empty slot cannot be reopened");
