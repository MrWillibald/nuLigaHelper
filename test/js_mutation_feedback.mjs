// Settings must preserve the saved/refused selection and carry exactly its result.
import assert from 'node:assert/strict';
const calls = [];
const feedback = [];
const transfers = [];
let navigates = true;
function setting(dataset) {
  return { dataset, value: '7', listeners: {}, disabled: false,
    addEventListener(event, callback) { this.listeners[event] = callback; }, closest() { return null; } };
}
const team = setting({ game: '1' });
const mv = setting({ team: '2' });
globalThis.document = { querySelector: () => ({ content: 'synthetic-csrf' }),
  querySelectorAll: (selector) => selector === '.team-select' ? [team] : selector === '.mv-select' ? [mv] : [] };
globalThis.nuLigaFeedback = { show: (value) => feedback.push(value),
  navigate: (values, path) => { transfers.push({ values, path });
    if (!navigates) feedback.push(...values.map((value) => ({ ...value, action: path === '/login' ? 'signin' : 'refresh' })));
    return navigates;
  } };
let response = { status: 200, json: async () => ({ ok: true }) };
globalThis.fetch = async (url, options) => { calls.push({ url, body: JSON.parse(options.body) }); return response; };
await import('../static/app.js');
for (const [select, endpoint, field, text] of [[team, '/api/games/1/team', 'team_id', 'Verantwortliche Mannschaft'],
  [mv, '/api/teams/2/mv', 'person_id', 'Mannschaftsverantwortlicher']]) {
  select.listeners.focus(); select.value = '9'; await select.listeners.change();
  assert.equal(calls.at(-1).url, endpoint); assert.equal(calls.at(-1).body[field], 9);
  assert.equal(transfers.at(-1).values[0].severity, 'success');
  assert(transfers.at(-1).values[0].message.includes(text)); assert.equal(feedback.length, 0, 'result displayed before navigation');
  response = { status: 403, json: async () => ({ ok: false, error: 'Keine Berechtigung.' }) };
  select.listeners.focus(); select.value = '11'; await select.listeners.change();
  assert.equal(select.value, '9'); assert.equal(select.disabled, false);
  assert.equal(feedback.at(-1).severity, 'error');
  response = { status: 401 }; select.value = '11'; await select.listeners.change();
  assert.equal(transfers.at(-1).path, '/login'); assert(transfers.at(-1).values[0].message.includes('Sitzung ist abgelaufen'));
  assert.equal(select.value, '9');
  response = { status: 200, json: async () => ({ ok: true }) };
  navigates = false; select.value = '12'; await select.listeners.change();
  assert.equal(select.value, '12'); assert.equal(feedback.at(-1).severity, 'success'); assert.equal(feedback.at(-1).action, 'refresh');
  navigates = true;
  response = { status: 200, json: async () => { throw Error('unreadable'); } };
  select.value = '13'; await select.listeners.change();
  assert.equal(feedback.at(-1).action, 'refresh'); assert(feedback.at(-1).message.includes('nicht bestätigt'));
  response = { status: 200, json: async () => ({ arbitrary: 'malformed' }) };
  select.value = '13'; await select.listeners.change();
  assert(feedback.at(-1).message.includes('nicht bestätigt'));
  response = { status: 200, json: async () => ({ ok: true }) };
  feedback.length = 0;
}
