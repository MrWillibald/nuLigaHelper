// Dependency-free lifecycle/transport tests with a controllable clock and viewport.
import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
const source = fs.readFileSync('static/feedback.js', 'utf8');
class Element {
  constructor() { this.children = []; this.dataset = {}; this.style = {}; this.listeners = {}; this.attrs = {}; this.classList = { add() {} }; }
  appendChild(child) { child.parent = this; this.children.push(child); return child; }
  replaceChildren() { this.children = []; }
  querySelectorAll() { return this.children.filter((child) => child.server); }
  querySelector(selector) { return selector === 'button' ? this.children.find((child) => child.attrs['aria-label']) : this.children[0]; }
  setAttribute(name, value) { this.attrs[name] = value; }
  addEventListener(name, callback) { this.listeners[name] = callback; }
  remove() { this.parent.children.splice(this.parent.children.indexOf(this), 1); }
  contains(child) { return this === child || this.children.includes(child); }
  getBoundingClientRect() { return this.rect || { top: 200, bottom: 300, height: 100 }; }
  focus() { this.focused = true; }
  matches() { return this.open; }
  showPopover() { this.open = true; }
  hidePopover() { this.open = false; }
  get scrollHeight() { return this.children.length * 100; }
}
function boot({ stored, blocked = false, boundary = false, server = [], path = '/' } = {}) {
  let clock = 100000;
  let data = stored;
  const region = new Element();
  region.dataset.authBoundary = String(boundary);
  server.forEach((value) => {
    const entry = new Element(); entry.server = true; entry.dataset.severity = value.severity;
    const text = new Element(); text.textContent = value.message; entry.appendChild(text); region.appendChild(entry);
  });
  const document = { body: new Element(), hidden: false, getElementById: () => region, createElement: () => new Element(),
    querySelectorAll: () => [], addEventListener(name, callback) { this[name] = callback; } };
  const location = { pathname: path, search: '', href: `https://club.test${path}`, origin: 'https://club.test', reload() { this.reloaded = true; } };
  const window = { location, innerWidth: 320, innerHeight: 600, listeners: {}, addEventListener(name, callback) { this.listeners[name] = callback; }, visualViewport: {
    width: 300, height: 400, offsetLeft: 10, offsetTop: 50, addEventListener() {},
  } };
  const storage = { getItem: () => data, removeItem() { if (blocked) throw Error('blocked'); data = null; },
    setItem(key, value) { if (blocked) throw Error('blocked'); data = value; } };
  const context = { document, window, sessionStorage: storage, URL, Date: { now: () => clock }, setInterval() {}, setTimeout(callback) { callback(); } };
  vm.createContext(context); vm.runInContext(source, context);
  return { api: context.nuLigaFeedback, region, document, window, location, storage: () => data,
    advance(ms) { clock += ms; context.nuLigaFeedback.tick(); }, clock: () => clock };
}
const page = boot({ server: [{ message: '<script>plain name</script>', severity: 'success' }] });
assert.equal(page.region.children.length, 1, 'hydration duplicates server feedback');
assert.equal(page.region.children[0].children[0].textContent, 'Erfolg: <script>plain name</script>');
assert.equal(page.region.style.left, '160px');
assert.equal(page.region.style.top, '438px');
assert.equal(page.region.style.width, '276px');
assert.equal(page.region.style.maxHeight, '160px');
page.advance(4999); assert.equal(page.region.children.length, 1);
page.advance(1); assert.equal(page.region.children.length, 0);
for (const severity of ['success', 'info', 'warning', 'error']) {
  const entry = page.api.show({ message: 'Timed result', severity });
  assert.equal(entry.querySelector('button'), undefined, `${severity} still has a close button`);
  page.advance(4999); assert.equal(page.region.children.length, 1, `${severity} expired too soon`);
  page.advance(1); assert.equal(page.region.children.length, 0, `${severity} did not expire after five seconds`);
}
for (const severity of ['success', 'info', 'warning', 'error']) for (const pause of ['pointerenter', 'focusin']) {
  const entry = page.api.show({ message: 'Reading', severity });
  page.advance(2000); entry.listeners[pause](); page.advance(9000);
  assert.equal(page.region.children.length, 1, `${pause} failed to pause`);
  entry.listeners[pause === 'pointerenter' ? 'pointerleave' : 'focusout']({ relatedTarget: null });
  page.advance(2999); assert.equal(page.region.children.length, 1);
  page.advance(1); assert.equal(page.region.children.length, 0);
}
let entry = page.api.show({ message: 'Visible', severity: 'success' });
page.advance(1000); entry.rect = { top: 0, bottom: 100 }; page.advance(10000);
entry.rect = null; page.advance(3999); assert.equal(page.region.children.length, 1);
page.document.hidden = true; page.document.visibilitychange(); page.advance(10000);
page.document.hidden = false; page.document.visibilitychange(); page.advance(1);
assert.equal(page.region.children.length, 0, 'visible timer did not resume');
const stack = boot();
for (const severity of ['warning', 'error', 'error']) {
  const unread = stack.api.show({ severity, message: 'Unread' });
  unread.rect = { top: 0, bottom: 100 };
}
stack.advance(90000); assert.equal(stack.region.children.length, 3, 'unread results expired outside the visible stack');
entry = stack.api.show({ severity: 'success', message: 'Newest' });
assert.equal(stack.region.scrollTop, stack.region.scrollHeight, 'new result hidden behind unread errors');
assert.equal(stack.region.children.length, 4);
assert.equal(stack.region.children.at(-1).attrs.role, 'group');
stack.advance(5000); assert.equal(stack.region.children.length, 3);
stack.region.children.forEach((unread) => { unread.rect = null; });
stack.advance(4999); assert.equal(stack.region.children.length, 3);
stack.advance(1); assert.equal(stack.region.children.length, 0, 'warnings/errors did not expire after visible reading time');
assert.equal(page.document.body.children[1].attrs.role, 'alert');
assert.equal(stack.document.body.children[1].textContent, 'Fehler: Unread');
assert.equal(page.document.body.children[0].attrs.role, 'status');
assert.equal(page.api.show({ message: 'Bad', severity: 'invalid' }), undefined);
const transfer = boot();
assert.equal(transfer.api.navigate([{ message: 'Saved', severity: 'success', csrf_token: 'SECRET', email: 'secret@test' }]), true);
const raw = transfer.storage();
assert(!raw.includes('SECRET') && !raw.includes('email') && !raw.includes('csrf_token'), 'transport copied API/form data');
const arrival = boot({ stored: raw });
assert.equal(arrival.region.children.length, 1); assert.equal(arrival.storage(), null);
arrival.advance(4999); assert.equal(arrival.region.children.length, 1);
arrival.advance(1); assert.equal(arrival.region.children.length, 0);
assert.equal(boot({ stored: arrival.storage() }).region.children.length, 0, 'consumed feedback replayed');
for (const stored of ['broken', '{}', JSON.stringify({ ...JSON.parse(raw), expires: 99999 }),
  JSON.stringify({ ...JSON.parse(raw), destination: '/other' }),
  JSON.stringify({ ...JSON.parse(raw), entries: [{ id: 'x', message: 'bad', severity: 'invalid' }] })]) {
  const rejected = boot({ stored }); assert.equal(rejected.region.children.length, 0); assert.equal(rejected.storage(), null);
}
assert.equal(boot({ stored: raw, boundary: true }).region.children.length, 0, 'auth boundary replayed old result');
const signin = boot(); signin.api.navigate([{ message: 'Expired', severity: 'error' }], '/login');
const expiredSession = boot({ stored: signin.storage(), path: '/login' });
assert.equal(expiredSession.region.children.length, 1, 'expiry notice lost on login destination');
expiredSession.advance(4999); assert.equal(expiredSession.region.children.length, 1, 'transferred error expired too soon');
expiredSession.advance(1); assert.equal(expiredSession.region.children.length, 0, 'transferred error did not expire after five seconds');
const recovery = boot({ blocked: true });
assert.equal(recovery.api.navigate([{ severity: 'success', message: 'Saved' }]), false);
assert.equal(recovery.location.reloaded, undefined);
assert.equal(recovery.region.children[0].children[1].textContent, 'Seite neu laden');
const focus = new Element(); page.document.activeElement = focus;
page.api.show({ message: 'Asynchronous result', severity: 'success' });
assert.equal(page.document.activeElement, focus, 'feedback stole task-help focus');

page.window.listeners.pagehide();
assert.equal(page.region.children.length, 0, 'history restoration replayed consumed messages');
transfer.window.listeners.pagehide();
assert.equal(transfer.storage(), raw, 'pagehide removed the intentional immediate transfer');
