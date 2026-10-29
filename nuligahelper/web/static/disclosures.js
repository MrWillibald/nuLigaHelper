// This classic script runs after each summary, before its content is parsed.
// Repeated executions share state without resetting an earlier native choice.
(() => {
  const details = document.currentScript?.parentElement;
  if (!details?.matches("[data-responsive-disclosure]")) return;

  if (!globalThis.nuLigaDisclosures) {
    const viewport = globalThis.window?.matchMedia?.("(max-width: 700px)");
    // Keep the server's native fallback if responsive enhancement is unavailable.
    if (!viewport || !(viewport.addEventListener || viewport.addListener)) return;
    const states = new WeakMap();
    globalThis.nuLigaDisclosures = {
      initialize(panel) {
        if (states.has(panel)) return;
        const state = { deliberate: false, dirty: false };
        const applyDefault = () => {
          if (state.deliberate) return;
          // The server marker works even before any form fields exist.
          if (panel.dataset.disclosureError === "true"
              || panel.querySelector('[aria-invalid="true"]')) {
            panel.open = true;
            return;
          }
          if (state.dirty || panel.contains(document.activeElement)) return;
          panel.open = !viewport.matches && panel.dataset.desktopOpen === "true";
        };
        // Native keyboard and touch activation both produce summary clicks.
        // Toggle events also follow programmatic defaults, so do not track them.
        panel.querySelector("summary")?.addEventListener("click", () => { state.deliberate = true; });
        panel.addEventListener("input", () => { state.dirty = true; });
        panel.addEventListener("change", () => { state.dirty = true; });
        if (viewport.addEventListener) viewport.addEventListener("change", applyDefault);
        else viewport.addListener(applyDefault);
        states.set(panel, state);
        applyDefault();
      },
    };
  }
  globalThis.nuLigaDisclosures.initialize(details);
})();
