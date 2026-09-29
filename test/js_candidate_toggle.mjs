const requested = [];
const select = { disabled: true };
const message = { textContent: "" };
const retry = { hidden: true, addEventListener(name, callback) { this[name] = callback; } };
const login = { hidden: true };
const status = {
  hidden: true,
  querySelector(selector) {
    return ({ "[data-candidate-message]": message,
      "[data-candidate-retry]": retry,
      "[data-candidate-login]": login })[selector];
  },
};
const card = {
  open: false,
  dataset: { candidateUrl: "/api/games/1/candidates" },
  listeners: {},
  addEventListener(name, callback) { this.listeners[name] = callback; },
  querySelector(selector) {
    if (selector === "select[data-role], select[data-block-assignment]") return select;
    if (selector === "[data-candidate-status]") return status;
    if (selector === "[data-candidate-retry]") return retry;
    if (selector === "[data-candidate-message]") return message;
    return null;
  },
};
globalThis.document = {
  querySelectorAll(selector) {
    return selector === "details[data-candidate-url]" ? [card] : [];
  },
};
globalThis.fetch = async (url) => {
  requested.push(url);
  throw new Error("offline");
};
await import("../static/app.js");
if (requested.length || !card.listeners.toggle) {
  throw new Error("collapsed card loaded candidates or did not register expansion");
}
card.open = true;
card.listeners.toggle();
await new Promise((resolve) => setTimeout(resolve, 0));
if (requested.join() !== card.dataset.candidateUrl || !select.disabled || retry.hidden) {
  throw new Error("card expansion did not load only its candidates and expose retry on failure");
}
