// Public summary nodes shared by the synthetic DOM regressions.
export function progressFixture(kind = "game", filled = 3, total = 5) {
  const node = () => ({
    textContent: "", attrs: {}, children: [], hidden: false,
    setAttribute(name, value) { this.attrs[name] = String(value); },
    removeAttribute(name) { delete this.attrs[name]; },
    replaceChildren(...children) { this.children = children; },
  });
  const count = node(), arc = node(), gauge = node(), vacancies = node();
  const elements = { "[data-progress-count]": count, ".coverage-fill": arc,
    ".coverage-gauge": gauge, "[data-vacancies]": vacancies };
  const coverage = { dataset: { progressKind: kind, progressFilled: String(filled), progressTotal: String(total) },
    querySelector: selector => elements[selector] };
  const card = { querySelector: selector => selector === ".coverage" ? coverage : null };
  const snapshot = () => JSON.stringify({ dataset: coverage.dataset, hidden: coverage.hidden,
    count: count.textContent, arc: arc.attrs, gauge: gauge.attrs,
    badges: vacancies.children.map(badge => badge.textContent), badgesHidden: vacancies.hidden });
  return { coverage, card, count, arc, gauge, vacancies, snapshot };
}
