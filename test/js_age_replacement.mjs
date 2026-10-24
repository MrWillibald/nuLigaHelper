import assert from "node:assert/strict";
import { progressFixture } from "./js_progress_fixture.mjs";
import { assignmentFixture } from "./js_assignment_fixture.mjs";

const page = assignmentFixture();
await import("../static/app.js");
const tools = globalThis.nuLigaAssignmentTools;
const requests = [];
const settings = [{ kind: "game", role: "Verkauf" }, { kind: "Vorbereitung" }, { kind: "Kuchenlieferung" }, { kind: "Aufräumen" }];

for (const fixture of settings) {
  const { select, field, status } = page.createDuty({ ...fixture, occupant: 7, token: "original-request" });
  tools.bindAssignmentSelect(select);
  const prefix = fixture.kind === "game" ? "/api/assignment/" : "/api/block-assignment/";
  const originalQuery = field.parent.querySelector.bind(field.parent);
  const summary = progressFixture(fixture.kind === "game" ? "game" : "block", 2, 3);
  const { coverage } = summary;
  field.parent.querySelector = selector => selector === ".coverage" ? coverage : originalQuery(selector);
  const savedSummary = fixture.kind === "game" ? {
    staffing: { required_filled: 2, required_total: 3, vacancies: { Verkauf: 1 }, deficiencies: [] },
  } : { block: { phase: fixture.kind, progress: { filled: 2, total: 3 }, vacancies: { [fixture.kind]: 1 } } };
  if (savedSummary.staffing) globalThis.nuLigaProgressTools.updateGameStaffing(field.parent, savedSummary.staffing);
  else globalThis.nuLigaProgressTools.updateBlockSummary(field.parent, savedSummary.block);
  const initialSummary = summary.snapshot();
  globalThis.fetch = async (url, options) => {
    requests.push({ url, body: JSON.parse(options.body) });
    return { status: 400, json: async () => ({ ok: false, error: "Verkauf benötigt eine erwachsene Person." }) };
  };
  select.focus(); select.value = "9";
  const cancelled = select.dispatch("change");
  const before = requests.length;
  page.cancel.dispatch("click"); await cancelled;
  assert.equal(requests.length, before, "cancelled takeover must not release or claim anything");
  assert.equal(field.dataset.occupantId, "7");
  assert.equal(field.dataset.replacementToken, "original-request");
  assert.equal(summary.snapshot(), initialSummary);
  select.value = "9";
  const refused = select.dispatch("change");
  page.confirm.dispatch("click"); await refused;
  assert.equal(requests.length, before + 1, "takeover is one mutation, never release followed by claim");
  assert.equal(requests.at(-1).url, prefix + "takeover");
  assert.equal(requests.at(-1).body.expected_person_id, 7);
  assert.equal(requests.at(-1).body.expected_request_token, "original-request");
  assert.equal(field.dataset.occupantId, "7", "refused takeover must retain original responsibility");
  assert.equal(select.value, "7");
  assert.equal(coverage.dataset.progressFilled, "2", "a refused transfer must not create a visible vacancy");
  assert.equal(field.dataset.replacementToken, "original-request");
  assert.equal(page.feedback.at(-1).severity, "error");
  assert.ok(!page.feedback.at(-1).message.includes("freigegeben"));

  globalThis.fetch = async (url, options) => { requests.push({ url, body: JSON.parse(options.body) }); throw new Error("lost response"); };
  select.value = "9";
  const lost = select.dispatch("change"); page.confirm.dispatch("click"); await lost;
  assert.equal(select.value, "7");
  assert.match(page.feedback.at(-1).message, /nicht bestätigt/);
  assert.ok(!page.feedback.at(-1).message.includes("freigegeben"));
  assert.ok(!page.feedback.at(-1).message.includes("abgelehnt"));

  // Close/reopen during confirmation must not substitute a newer request token.
  select.value = "9";
  const stale = select.dispatch("change");
  field.dataset.replacementToken = "reopened-request";
  globalThis.fetch = async (url, options) => {
    requests.push({ url, body: JSON.parse(options.body) });
    return { status: 409, json: async () => ({ ok: false, error: "Ersatzsuche wurde geändert." }) };
  };
  page.confirm.dispatch("click"); await stale;
  assert.equal(requests.at(-1).body.expected_request_token, "original-request");
  assert.equal(field.dataset.occupantId, "7");
  assert.equal(field.dataset.replacementToken, "reopened-request");

  globalThis.fetch = async (url, options) => {
    requests.push({ url, body: JSON.parse(options.body) });
    return { status: 200, json: async () => ({ ok: true, ...savedSummary, assignment: {
      person_id: 9, person_name: "Replacement Helper", replacement_token: null, seeking_replacement: false, locked: false,
    } }) };
  };
  select.value = "9";
  const accepted = select.dispatch("change"); page.confirm.dispatch("click"); await accepted;
  assert.equal(requests.at(-1).body.expected_request_token, "reopened-request");
  assert.equal(field.dataset.occupantId, "9");
  assert.equal(field.dataset.replacementToken, "");
  assert.equal(coverage.dataset.progressFilled, "2");
  assert.equal(summary.snapshot(), initialSummary, "takeover must preserve the occupied count and bar");
  assert.equal(select.hidden, true);
  assert.match(status.textContent, /Verbindlich/);
}

assert.ok(requests.every((request) => request.url.endsWith("/takeover")), "ordinary replacements must never use release/claim endpoints");
