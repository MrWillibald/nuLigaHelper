// Public summary nodes shared by the synthetic DOM regressions.
export function progressFixture(kind = "game", filled = 3, total = 5) {
  const node = () => ({
    textContent: "", style: {}, attrs: {}, children: [], hidden: false,
    setAttribute(name, value) { this.attrs[name] = String(value); },
    removeAttribute(name) { delete this.attrs[name]; },
    replaceChildren(...children) { this.children = children; },
  });
  const count = node(), fill = node(), track = node(), percent = node();
  const elements = { "[data-progress-count]": count, ".coverage-fill": fill,
    ".coverage-track": track, "[data-progress-percent]": percent };
  const coverage = { dataset: { progressKind: kind, progressFilled: String(filled), progressTotal: String(total) },
    querySelector: selector => elements[selector] };
  const card = { querySelector: selector => selector === ".coverage" ? coverage : null };
  const snapshot = () => JSON.stringify({ dataset: coverage.dataset, hidden: coverage.hidden,
    count: count.textContent, width: fill.style.width, track: track.attrs, percent: percent.textContent });
  return { coverage, card, count, fill, track, percent, snapshot };
}
