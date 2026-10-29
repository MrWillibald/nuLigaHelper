import { installConfirmation } from "./js_confirmation_fixture.mjs";

export function assignmentFixture({ autoConfirm = false } = {}) {
  class Element {
    constructor(tag = "div") {
      this.tag = tag; this.children = []; this.dataset = {}; this.attrs = {};
      this.listeners = new Map(); this.textContent = ""; this.hidden = false; this.disabled = false;
      this.style = {}; this.selected = false; this._value = "";
      this.classList = { add() {}, remove() {}, toggle() {}, contains() { return false; } };
    }
    get options() { return this.children; }
    get value() { return this.tag === "select" ? this.children.find((child) => child.selected)?.value || "" : this._value; }
    set value(value) {
      this._value = String(value);
      if (this.tag === "select") this.children.forEach((child) => { child.selected = child.value === String(value); });
    }
    get selectedOptions() { return this.children.filter((child) => child.selected); }
    get isConnected() { return this === document.body || Boolean(this.parent?.isConnected); }
    setAttribute(name, value) { this.attrs[name] = String(value); }
    getAttribute(name) { return this.attrs[name] ?? null; }
    hasAttribute(name) { return name in this.attrs; }
    removeAttribute(name) { delete this.attrs[name]; if (name === "selected") this.selected = false; }
    appendChild(child) {
      if (child.tag === "fragment") [...child.children].forEach((item) => this.appendChild(item));
      else { child.remove(); child.parent = this; this.children.push(child); }
    }
    remove() { if (this.parent) this.parent.children = this.parent.children.filter((child) => child !== this); this.parent = null; }
    insertBefore(child, before) {
      child.remove(); child.parent = this;
      this.children.splice(before ? this.children.indexOf(before) : this.children.length, 0, child);
    }
    replaceChildren(...children) { [...this.children].forEach((child) => child.remove()); children.forEach((child) => this.appendChild(child)); }
    cloneNode() {
      const clone = new Element(this.tag); clone.value = this.value; clone.textContent = this.textContent;
      clone.dataset = { ...this.dataset }; clone.selected = this.selected; return clone;
    }
    matches(selector) {
      if (selector === ".day-block") return this.dayBlock === true;
      if (selector === ".day-title") return this.dayTitle === true;
      if (selector.startsWith("option[value=")) return this.tag === "option" && this.value === selector.match(/value="([^"]+)"/)[1];
      if (selector === "select[data-role]") return this.tag === "select" && this.dataset.role !== undefined;
      if (selector === "select[data-block-assignment]") return this.tag === "select" && this.hasAttribute("data-block-assignment");
      const attribute = selector.match(/^\[([^\]]+)\]$/)?.[1];
      return attribute ? this.hasAttribute(attribute) : this.tag === selector;
    }
    closest(selector) { for (let item = this; item; item = item.parent) if (item.matches(selector)) return item; return null; }
    querySelectorAll(selector) {
      const selectors = selector.split(", ");
      return this.children.flatMap((child) => [
        ...(selectors.some((part) => child.matches(part)) ? [child] : []), ...child.querySelectorAll(selector),
      ]);
    }
    querySelector(selector) { return this.querySelectorAll(selector)[0] || null; }
    addEventListener(name, callback) {
      const callbacks = this.listeners.get(name) || new Set(); callbacks.add(callback); this.listeners.set(name, callbacks);
    }
    async dispatch(name, properties = {}) {
      const event = { target: this, preventDefault() {}, ...properties };
      for (const callback of this.listeners.get(name) || []) await callback(event);
    }
    focus() { document.activeElement = this; for (const callback of this.listeners.get("focus") || []) callback(); }
  }
  const cards = new Map();
  const feedback = [];
  globalThis.document = {
    querySelectorAll: () => [], querySelector: () => ({ content: "synthetic-csrf" }),
    getElementById: (id) => cards.get(id) || null,
    createElement: (tag) => new Element(tag), createDocumentFragment: () => new Element("fragment"),
    activeElement: null, body: new Element("body"),
  };
  globalThis.nuLigaFeedback = {
    show: (value) => feedback.push(value),
    navigate: (values) => { feedback.push(...values); return true; },
  };
  const confirmation = installConfirmation(document, { autoConfirm });
  let nextId = 0;
  function createDuty({ kind = "game", role = "Verkauf", occupant = null, token = null, locked = false, late = false } = {}) {
    const id = String(++nextId);
    const day = new Element(); day.dayBlock = true; document.body.appendChild(day);
    const title = new Element(); title.dayTitle = true; title.textContent = "Samstag 28.11.2026"; day.appendChild(title);
    const card = new Element("details"); day.appendChild(card); card._candidateLoaded = true;
    cards.set(`${kind === "game" ? "game" : "block"}-${id}`, card);
    const field = new Element(); field.setAttribute("data-occupant-id", "");
    field.dataset = { occupantId: occupant ? String(occupant) : "", replacementToken: token || "", locked: String(locked),
      seekingReplacement: String(Boolean(token)), personName: occupant ? "Original Helper" : "", canRequestReplacement: "true" };
    card.appendChild(field);
    const strong = new Element("strong"); strong.setAttribute("data-duty-occupant", ""); strong.textContent = occupant ? "Original Helper · Supporter" : ""; field.appendChild(strong);
    const status = new Element("p"); status.setAttribute("data-duty-state", ""); field.appendChild(status);
    const select = new Element("select");
    select.dataset = { ...(kind === "game" ? { game: id, role } : { block: id }), slot: "0", dutyLabel: kind === "game" ? role : kind,
      replacementToken: token || "", locked: String(locked), lateClaimLocked: String(late) };
    if (kind !== "game") select.setAttribute("data-block-assignment", "");
    for (const [value, name] of [["", "– offen –"], ["7", "Original Helper · Supporter"], ["9", "Replacement <b>plain</b> · Team A"], ["11", "Concurrent Winner · Team B"]]) {
      const option = new Element("option"); option.value = value; option.textContent = name; select.appendChild(option);
    }
    select.value = occupant ? String(occupant) : ""; field.appendChild(select);
    const button = new Element("button"); button.setAttribute("data-request-replacement", ""); button.dataset = { ...select.dataset }; field.appendChild(button);
    return { card, field, select, button, status, strong };
  }
  return { createDuty, feedback, ...confirmation };
}
