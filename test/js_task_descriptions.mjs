// Exercise public task help using DOM events without a browser dependency.
import assert from "node:assert/strict";

class Element {
  constructor(tag) {
    this.tag = tag;
    this.children = [];
    this.attrs = {};
    this.listeners = new Map();
    this._scrollTop = 0;
    this.style = new Proxy({}, {
      set: (styles, name, value) => {
        styles[name] = value;
        // Browsers clamp an element's reading position when its scrollable
        // area temporarily grows while floating help is being measured.
        if (name === "maxHeight") this.scrollTop = this.scrollTop;
        return true;
      },
    });
    this.dataset = {};
    this.hidden = false;
    this.textContent = "";
    this.rect = { left: 100, top: 100, width: tag === "p" ? 300 : 20, height: tag === "p" ? 120 : 20 };
  }
  setAttribute(name, value) { this.attrs[name] = String(value); }
  getAttribute(name) { return this.attrs[name] ?? null; }
  get parentElement() { return this.parent instanceof Element ? this.parent : null; }
  get tagName() { return this.tag.toUpperCase(); }
  appendChild(child) { child.parent = this; this.children.push(child); }
  remove() {
    if (this.parent) this.parent.children = this.parent.children.filter((child) => child !== this);
    this.parent = null;
  }
  get isConnected() { return this === document.body || Boolean(this.parent?.isConnected); }
  get scrollHeight() { return this.contentHeight ?? this.rect.height; }
  get clientHeight() {
    const maximum = Number.parseFloat(this.style.maxHeight);
    return Math.min(this.scrollHeight, Number.isFinite(maximum) ? maximum : this.rect.height);
  }
  get scrollTop() { return this._scrollTop; }
  set scrollTop(value) { this._scrollTop = Math.max(0, Math.min(value, this.scrollHeight - this.clientHeight)); }
  getBoundingClientRect() {
    this.measurements = (this.measurements || 0) + 1;
    return { ...this.rect, right: this.rect.left + this.rect.width, bottom: this.rect.top + this.rect.height };
  }
  closest(selector) {
    let current = this;
    while (current) {
      if (current.tag === selector) return current;
      current = current.parent;
    }
    return null;
  }
  contains(element) { return this === element || this.children.some((child) => child.contains(element)); }
  showPopover() { this.popoverOpen = true; this.popoverShows = (this.popoverShows || 0) + 1; }
  hidePopover() { this.popoverOpen = false; }
  addEventListener(name, callback) {
    const callbacks = this.listeners.get(name) || [];
    callbacks.push(callback);
    this.listeners.set(name, callbacks);
  }
  querySelector(selector) {
    const attribute = selector.match(/^\[([^\]]+)\]$/)?.[1];
    for (const child of this.children) {
      if (attribute ? attribute in child.attrs : child.tag === selector) return child;
      const descendant = child.querySelector(selector);
      if (descendant) return descendant;
    }
    return null;
  }
  dispatch(name, properties = {}) {
    const event = {
      target: this, bubbles: ["click", "keydown", "change", "submit"].includes(name),
      defaultPrevented: false, stopped: false,
      preventDefault() { this.defaultPrevented = true; },
      stopPropagation() { this.stopped = true; },
      ...properties,
    };
    let current = this;
    do {
      event.currentTarget = current;
      for (const callback of current.listeners.get(name) || []) callback(event);
      current = event.bubbles && !event.stopped ? current.parent : null;
    } while (current);
    return event;
  }
  focus() {
    if (document.activeElement === this) return;
    document.activeElement?.dispatch("blur", { relatedTarget: this });
    document.activeElement = this;
    this.dispatch("focus");
  }
  blur() {
    if (document.activeElement !== this) return;
    document.activeElement = null;
    this.dispatch("blur");
  }
}

globalThis.document = {
  activeElement: null,
  listeners: new Map(),
  documentElement: { clientWidth: 360, clientHeight: 640 },
  querySelectorAll: () => [],
  createElement: (tag) => new Element(tag),
  addEventListener: Element.prototype.addEventListener,
  dispatch: Element.prototype.dispatch,
};
document.body = new Element("body");
document.body.parent = document;
const windowListeners = new Map();
globalThis.window = {
  innerWidth: 360, innerHeight: 640,
  addEventListener(name, callback) {
    const callbacks = windowListeners.get(name) || [];
    callbacks.push(callback);
    windowListeners.set(name, callbacks);
  },
  dispatch(name) { for (const callback of windowListeners.get(name) || []) callback({ type: name }); },
};
const pendingTimers = new Map();
let nextTimer = 0;
globalThis.setTimeout = (callback) => { const id = ++nextTimer; pendingTimers.set(id, callback); return id; };
globalThis.clearTimeout = (id) => pendingTimers.delete(id);
function flushTimers() {
  const callbacks = [...pendingTimers.values()];
  pendingTimers.clear();
  callbacks.forEach((callback) => callback());
}
let requests = 0;
globalThis.fetch = async () => { requests++; throw new Error("task help must not contact assignment APIs"); };
await import("../static/app.js");
const { initializeTaskHelp, createTaskHelp, taskHelpPlacement, refreshTaskHelp } = globalThis.nuLigaTaskHelpTools;
const card = new Element("details");
card.open = true;
document.body.appendChild(card);
const form = new Element("form");
card.appendChild(form);
const guidance = '<script>Text remains plain</script> & "quoted"';
const field = createTaskHelp("Verkauf 1", guidance, true);
form.appendChild(field);
const help = field.querySelector("[data-task-help]");
const description = field.querySelector("[data-task-description]");
const label = field.querySelector("label");
const select = new Element("select");
select.id = label.getAttribute("for");
select.value = "saved-helper";
field.appendChild(select);
let submissions = 0;
let changes = 0;
let cardClicks = 0;
form.addEventListener("submit", () => { submissions++; });
select.addEventListener("change", () => { changes++; });
card.addEventListener("click", () => { cardClicks++; card.open = !card.open; });

assert.equal(help.type, "button", "task help must not trigger a form's default submit action");
assert.equal(help.getAttribute("aria-label"), "Informationen zu Verkauf 1");
assert.equal(help.getAttribute("aria-controls"), description.id);
assert.equal(help.getAttribute("aria-describedby"), description.id);
assert.equal(description.textContent, guidance, "guidance must be written as plain text");
assert.equal(description.children.length, 0, "HTML-like guidance must not create elements");
assert.equal(label.getAttribute("for"), select.id, "assignment label activation must still target the select");
assert.equal(description.hidden, true);

function visible(expected, reason) {
  assert.equal(!description.hidden, expected, reason);
  assert.equal(help.getAttribute("aria-expanded"), String(expected), "expanded state must match displayed guidance");
}

help.dispatch("pointerenter", { pointerType: "mouse" });
visible(true, "pointer hover must expose task guidance");
// Cross the small gap between the icon and the floating paragraph without
// dismissing it before its own pointerenter event can run.
help.dispatch("pointerleave", { pointerType: "mouse", relatedTarget: null });
field.dispatch("pointerleave", { pointerType: "mouse", relatedTarget: null });
visible(true, "moving from the information icon to the bubble must not immediately hide it");
assert.ok(pendingTimers.size > 0, "crossing empty space should allow a brief transition into the bubble");
description.dispatch("pointerenter", { pointerType: "mouse" });
flushTimers();
visible(true, "guidance must remain readable when the pointer enters its text");
description.dispatch("pointerleave", { pointerType: "mouse" });
field.dispatch("pointerleave", { pointerType: "mouse" });
flushTimers();
visible(false, "leaving unpinned help must dismiss its description");

help.focus();
visible(true, "keyboard focus must expose task guidance");
const ordinaryKey = help.dispatch("keydown", { key: "ArrowDown" });
assert.equal(ordinaryKey.defaultPrevented, false, "unrelated keyboard behavior must stay available");
visible(true, "unrelated keys must not dismiss task guidance");
const escape = help.dispatch("keydown", { key: "Escape" });
assert.equal(escape.defaultPrevented, true);
visible(false, "Escape must dismiss even while the information button stays focused");
assert.equal(document.activeElement, help, "Escape must preserve the keyboard position");
help.blur();
help.focus();
visible(true, "guidance must become available again on a new focus interaction");
assert.equal(description.getAttribute("tabindex"), "0", "long task guidance must be reachable for keyboard scrolling");
help.dispatch("blur", { relatedTarget: description });
visible(true, "transferring keyboard focus into the bubble must not briefly hide it");
description.focus();
visible(true, "keyboard users must be able to read and scroll the focused bubble");
description.dispatch("keydown", { key: "Escape" });
visible(false, "Escape inside the floating description must dismiss it");
assert.equal(document.activeElement, help, "Escape inside task guidance must resume keyboard navigation at its information button");
description.blur();
help.focus();
description.focus();
select.focus();
visible(false, "leaving unpinned guidance for the assignment select must dismiss it");
help.focus();

const pinClick = help.dispatch("click");
assert.equal(pinClick.defaultPrevented, true);
help.blur();
field.dispatch("pointerleave");
flushTimers();
visible(true, "click must keep guidance open after hover and focus leave");
help.dispatch("click");
visible(false, "a second click must close pinned guidance");

// Touch synthesizes click, and need not generate mouse hover or focus first.
help.dispatch("click", { pointerType: "touch" });
visible(true, "touch tap must expose guidance without hover");
help.dispatch("click", { pointerType: "touch" });
visible(false, "a second touch tap must dismiss guidance");

help.focus();
help.dispatch("pointerenter");
help.dispatch("click");
help.dispatch("keydown", { key: "Escape" });
visible(false, "Escape must suppress both focus and hover after closing pinned guidance");
assert.equal(document.activeElement, help);
field.dispatch("pointerleave");
flushTimers();
help.blur();
help.dispatch("pointerenter");
visible(true, "dismissal must not permanently disable pointer guidance");
field.dispatch("pointerleave");
flushTimers();

// A repeated initializer must not register duplicate toggles on refreshed cards.
initializeTaskHelp(field);
help.dispatch("click");
visible(true, "initialization must not bind duplicate click handlers");
select.focus();
select.dispatch("keydown", { key: "Escape" });
visible(false, "Escape from the adjacent assignment control must dismiss pinned guidance");
assert.equal(document.activeElement, select, "dismissing help must not steal focus from an assignment");

const readOnly = createTaskHelp("Verkauf 2", guidance);
assert.equal(readOnly.children[0].children[0].tag, "span", "guest task text must not activate a select");
assert.notEqual(readOnly.querySelector("[data-task-description]").id, description.id,
  "repeated semantic tasks need distinct accessible description identifiers");
assert.equal(readOnly.querySelector("[data-task-description]").textContent, guidance,
  "numbered tasks must share the same semantic description");
assert.equal(submissions, 0, "help interactions must not submit assignments");
assert.equal(changes, 0, "help interactions must not change an assignment");
assert.equal(requests, 0, "help interactions must not fetch or mutate assignments");
assert.equal(select.value, "saved-helper");
assert.equal(cardClicks, 0, "information clicks must not propagate to unrelated card toggles");
assert.equal(card.open, true);

// Test viewport outcomes independently of exact offsets or CSS declarations.
for (const [anchor, side] of [
  [{ left: 0, right: 20, top: 0, bottom: 20 }, "below"],
  [{ left: 340, right: 360, top: 0, bottom: 20 }, "below"],
  [{ left: 0, right: 20, top: 620, bottom: 640 }, "above"],
  [{ left: 340, right: 360, top: 620, bottom: 640 }, "above"],
]) {
  const placement = taskHelpPlacement(anchor, { width: 220, height: 120 }, { width: 360, height: 640 });
  assert.equal(placement.side, side, "edge placement must choose the side with readable space");
  assert.ok(placement.left >= 0 && placement.left + 220 <= 360, "the bubble must stay within horizontal viewport edges");
  assert.ok(placement.top >= 0 && placement.top + Math.min(120, placement.maxHeight) <= 640,
    "the bubble must stay within vertical viewport edges");
}
const tall = taskHelpPlacement({ left: 100, right: 120, top: 50, bottom: 70 },
  { width: 220, height: 900 }, { width: 360, height: 640 });
assert.ok(tall.maxHeight > 0 && tall.maxHeight < 900 && tall.top + tall.maxHeight <= 640,
  "long guidance must get a readable bounded height within the viewport");
const zoomed = taskHelpPlacement({ left: 320, right: 340, top: 360, bottom: 380 },
  { width: 220, height: 120 }, { width: 300, height: 300, left: 40, top: 80 });
assert.ok(zoomed.left >= 40 && zoomed.left + 220 <= 340
  && zoomed.top >= 80 && zoomed.top + Math.min(120, zoomed.maxHeight) <= 380,
  "a shifted visual viewport must keep help within its visible area");

// A pinned bubble must follow its information control after a viewport change.
help.dispatch("click");
visible(true, "opening help should expose its floating bubble");
assert.equal(description.popoverOpen, true, "supported browsers should expose help in the popover top layer");
help.rect = { left: 340, top: 620, width: 20, height: 20 };
window.dispatch("resize");
assert.ok(Number.parseFloat(description.style.left) >= 0
  && Number.parseFloat(description.style.left) + description.rect.width <= 360,
  "a visible bubble must reposition within the viewport after resizing");
assert.ok(Number.parseFloat(description.style.top) < help.rect.top,
  "a bubble near the bottom edge must reposition above its information control");
help.rect = { left: 5, top: 0, width: 20, height: 20 };
document.dispatch("scroll");
assert.ok(Number.parseFloat(description.style.top) >= help.getBoundingClientRect().bottom,
  "scrolling must reposition an open bubble at its current anchor");

// Long help on a narrow screen must scroll internally. Its captured scroll
// event must not restart floating placement or erase the reader's position.
window.innerWidth = 320;
window.innerHeight = 568;
help.rect = { left: 140, top: 270, width: 20, height: 20 };
description.rect.height = 352;
description.contentHeight = 352;
window.dispatch("resize");
assert.ok(description.clientHeight < description.scrollHeight, "the narrow-screen fixture must require text scrolling");
description.focus();
description.scrollTop = 60;
assert.equal(description.scrollTop, 60, "the fixture must have enough content to read below the first screenful");
const beforeInternalScroll = description.measurements;
document.dispatch("scroll", { target: description });
assert.equal(description.measurements, beforeInternalScroll,
  "scrolling within task guidance must not trigger overlay measurement or repositioning");
assert.equal(description.scrollTop, 60, "scrolling the bubble must retain the reader's position");
assert.equal(document.activeElement, description, "reading lower task guidance must retain keyboard focus");
help.rect.top = 260;
document.dispatch("scroll", { target: document });
assert.ok(description.measurements > beforeInternalScroll, "page scrolling must still reposition open guidance");
assert.equal(description.scrollTop, 60,
  "page repositioning must preserve reading position across temporary measurement expansion");
window.innerHeight = 548;
window.dispatch("resize");
assert.equal(description.scrollTop, 60, "viewport resizing must preserve a still-valid task-guidance reading position");
window.innerWidth = 360;
window.innerHeight = 640;

const outside = new Element("button");
document.body.appendChild(outside);
outside.focus();
document.dispatch("keydown", { key: "Escape" });
visible(false, "Escape must dismiss a pinned bubble even after focus leaves its assignment field");
assert.equal(document.activeElement, outside, "global Escape dismissal must preserve the user's keyboard position");
assert.equal(description.popoverOpen, false);

help.dispatch("click", { pointerType: "touch" });
card.open = false;
refreshTaskHelp();
visible(false, "closing a card must dismiss its floating task guidance");
assert.equal(description.popoverOpen, false, "closed cards must not leave detached top-layer help visible");
card.open = true;
refreshTaskHelp();
visible(false, "reopening a card must not restore a stale pinned bubble");
help.dispatch("click", { pointerType: "touch" });
field.remove();
refreshTaskHelp();
visible(false, "removing a dynamic assignment field must dismiss its floating task guidance");
assert.equal(description.popoverOpen, false, "removed fields must not leave help in the browser's top layer");
assert.equal(requests, 0, "placement and dismissal must not contact assignment APIs");
assert.equal(submissions, 0);
assert.equal(changes, 0);

// Older browsers still expose the same interactions through the fixed overlay.
const ordinaryCreateElement = document.createElement;
document.createElement = (tag) => {
  const element = ordinaryCreateElement(tag);
  element.showPopover = undefined;
  element.hidePopover = undefined;
  return element;
};
const fallback = createTaskHelp("Vorbereitung 1", "Zur angezeigten Zeit vorbereiten.");
document.createElement = ordinaryCreateElement;
form.appendChild(fallback);
const fallbackControl = fallback.querySelector("[data-task-help]");
const fallbackText = fallback.querySelector("[data-task-description]");
fallbackControl.dispatch("click", { pointerType: "touch" });
assert.equal(fallbackText.hidden, false, "browsers without native popovers must still expose task help");
assert.ok(Number.isFinite(Number.parseFloat(fallbackText.style.left)), "fallback help must receive overlay placement");
document.dispatch("keydown", { key: "Escape" });
assert.equal(fallbackText.hidden, true, "Escape must also dismiss the older-browser fallback");
