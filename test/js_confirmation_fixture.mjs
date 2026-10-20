// A native-dialog boundary for DOM regressions. Confirmation is opt-in so tests
// can inspect pending promises before any assignment request is allowed.
export function installConfirmation(document, { autoConfirm = false } = {}) {
  class Control {
    constructor() { this.listeners = new Map(); this.hidden = false; this.textContent = ""; }
    addEventListener(name, callback) {
      const callbacks = this.listeners.get(name) || new Set();
      callbacks.add(callback); this.listeners.set(name, callbacks);
    }
    removeEventListener(name, callback) { this.listeners.get(name)?.delete(callback); }
    dispatch(name, properties = {}) {
      const event = { target: this, preventDefault() {}, ...properties };
      for (const callback of [...(this.listeners.get(name) || [])]) callback(event);
    }
    focus() { document.activeElement = this; }
  }
  const confirm = new Control();
  const cancel = new Control();
  const context = new Control();
  const late = new Control();
  const dialog = new Control();
  dialog.open = false;
  dialog.querySelector = (selector) => ({
    "[data-confirmation-confirm]": confirm, "[data-confirmation-cancel]": cancel,
    "[data-confirmation-context]": context, "[data-confirmation-late]": late,
  })[selector] || null;
  dialog.showModal = () => {
    dialog.open = true;
    if (autoConfirm) queueMicrotask(() => confirm.dispatch("click"));
  };
  dialog.close = () => { dialog.open = false; dialog.dispatch("close"); };
  dialog.getBoundingClientRect = () => ({ top: 50, right: 400, bottom: 500, left: 50 });
  const originalGet = document.getElementById?.bind(document);
  document.getElementById = (id) => id === "assignment-confirmation" ? dialog : originalGet?.(id);
  return { dialog, confirm, cancel, context, late };
}
