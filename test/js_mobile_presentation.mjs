import assert from "node:assert/strict";

class Element {
  constructor(tag, dataset = {}) {
    this.tag = tag;
    this.dataset = dataset;
    this.children = [];
    this.listeners = new Map();
    this.attrs = {};
    this.style = {};
    this.hidden = false;
    this.open = true;
    this.classList = { add() {} };
  }
  addEventListener(name, callback) {
    const list = this.listeners.get(name) || [];
    list.push(callback);
    this.listeners.set(name, list);
  }
  dispatch(name, properties = {}) {
    const event = { target: this, preventDefault() { this.defaultPrevented = true; },
      stopPropagation() {}, ...properties };
    for (const callback of this.listeners.get(name) || []) callback(event);
    return event;
  }
  appendChild(child) { child.parentElement = this; this.children.push(child); }
  contains(node) { return node === this || this.children.some((child) => child.contains(node)); }
  matches(selector) { return selector.includes(this.tag); }
  querySelector(selector) {
    if (selector === '[aria-invalid="true"]') return this.invalidField || null;
    for (const child of this.children) {
      if (child.tag === selector || child.attrs[selector.slice(1, -1)] != null) return child;
      const descendant = child.querySelector(selector);
      if (descendant) return descendant;
    }
    return null;
  }
  setAttribute(key, value) { this.attrs[key] = String(value); }
  getAttribute(key) { return this.attrs[key] ?? null; }
  getBoundingClientRect() { return { left: 10, top: 10, bottom: 54, right: 54, width: 44, height: 44 }; }
  focus(options) { document.activeElement = this; this.focusOptions = options; }
  closest() { return null; }
}

const viewport = new Element("viewport");
viewport.matches = process.argv[2] !== "desktop";
const reduced = { matches: false };
const frames = [];
const observers = [];
globalThis.MutationObserver = class {
  constructor(callback) { observers.push(callback); }
  observe() {}
};
const details = Array.from({ length: 7 }, (_, index) => {
  const node = new Element("details", { responsiveDisclosure: "", desktopOpen: index < 5 ? "true" : "false" });
  node.open = index < 5;
  node.appendChild(new Element("summary"));
  node.appendChild(new Element("input"));
  return node;
});
details[4].dataset.disclosureError = "true";
const link = new Element("a");
const top = new Element("nav");
const feedback = new Element("section");
let dialog;
const documentEvents = new Element("document");
globalThis.document = {
  body: new Element("body"),
  activeElement: null,
  querySelectorAll: (selector) => selector === "[data-responsive-disclosure]" ? details : [],
  querySelector: (selector) => selector === "dialog[open]" ? dialog : null,
  getElementById: (id) => ({ "return-to-top": link, "page-navigation": top, feedback })[id],
  addEventListener: (...args) => documentEvents.addEventListener(...args),
  createElement: (tag) => new Element(tag),
};
globalThis.window = new Element("window");
Object.assign(window, {
  innerWidth: 390, innerHeight: 800, scrollY: 0,
  matchMedia(query) {
    assert.ok(["(max-width: 700px)", "(prefers-reduced-motion: reduce)"].includes(query));
    return query.includes("700") ? viewport : reduced;
  },
  requestAnimationFrame: (callback) => frames.push(callback),
  scrollTo(value) { this.lastScroll = value; this.scrollY = value.top; },
});
let requests = 0;
globalThis.fetch = async () => { requests++; throw new Error("presentation must not request candidates"); };
await import("../static/app.js");
const flush = () => { while (frames.length) frames.splice(0).forEach((callback) => callback()); };

assert.deepEqual(details.map((node) => node.open), viewport.matches
  ? [false, false, false, false, true, false, false]
  : [true, true, true, true, true, false, false], "initial responsive defaults or error reopening changed");
assert.equal(link.hidden, true, "floating action must be hidden at the top");

// Native activation also covers keyboard-generated clicks. Programmatic toggle
// events are deliberately ignored and must not lock an untouched disclosure.
details[0].children[0].dispatch("click");
details[0].open = !details[0].open;
const chosen = details[0].open;
details[1].open = true;
details[1].children[1].value = "unsaved value";
details[1].dispatch("input");
details[2].open = true;
document.activeElement = details[2].children[1];
details[3].dispatch("toggle");
viewport.matches = !viewport.matches;
viewport.dispatch("change");
assert.equal(details[0].open, chosen, "resize overrode deliberate activation");
assert.equal(details[1].open, true, "resize closed an unsaved form");
assert.equal(details[1].children[1].value, "unsaved value");
assert.equal(details[2].open, true, "resize removed the focused field");
assert.equal(details[3].open, !viewport.matches, "programmatic toggle was mistaken for a user choice");
assert.equal(details[4].open, true, "resize hid a validation error");
details[4].children[0].dispatch("click");
details[4].open = false;
viewport.matches = !viewport.matches;
viewport.dispatch("change");
assert.equal(details[4].open, false, "resize overrode deliberately closing an error disclosure");
details[5].children[0].dispatch("click");
details[5].open = true;
details[6].children[0].dispatch("click");
details[6].open = true;
assert.ok(details[5].open && details[6].open, "statistics sections must open independently");
assert.equal(requests, 0, "new disclosures started candidate-loading requests");

document.activeElement = null;
window.scrollY = 799;
window.dispatch("scroll"); flush();
assert.equal(link.hidden, true);
window.scrollY = 800;
window.dispatch("scroll"); flush();
assert.equal(link.hidden, false, "arrow did not appear at one viewport");
feedback.appendChild(new Element("message"));
observers.forEach((callback) => callback()); flush();
assert.equal(link.hidden, true, "arrow obstructed feedback");
feedback.children = [];
dialog = new Element("dialog");
observers.forEach((callback) => callback()); flush();
assert.equal(link.hidden, true, "arrow obstructed a dialog");
dialog = null;
document.activeElement = new Element("input");
documentEvents.dispatch("focusin"); flush();
assert.equal(link.hidden, true, "arrow obstructed a focused form control");
document.activeElement = null;
observers.forEach((callback) => callback()); flush();
assert.equal(link.hidden, false);

const field = globalThis.nuLigaTaskHelpTools.createTaskHelp("Verkauf", "Synthetic guidance");
const help = field.querySelector("[data-task-help]");
help.dispatch("click");
observers.forEach((callback) => callback()); flush();
assert.equal(link.hidden, true, "arrow obstructed task-help popover");
help.dispatch("click");
observers.forEach((callback) => callback()); flush();
assert.equal(link.hidden, false);

link.focus();
window.scrollY = 0;
window.dispatch("scroll"); flush();
assert.equal(link.hidden, false, "visibility update discarded keyboard focus");
reduced.matches = true;
const activation = link.dispatch("click"); flush();
assert.equal(activation.defaultPrevented, true);
assert.deepEqual(window.lastScroll, { top: 0, behavior: "instant" }, "reduced motion must disable animation");
assert.equal(document.activeElement, top, "return action lost keyboard position");
assert.deepEqual(top.focusOptions, { preventScroll: true });
assert.equal(link.hidden, true);
window.scrollY = 900;
window.dispatch("scroll"); flush();
reduced.matches = false;
link.dispatch("click"); flush();
assert.equal(window.lastScroll.behavior, "smooth");
assert.equal(requests, 0);
