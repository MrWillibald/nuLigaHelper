## Why

The home game overview is difficult to scan when visitors need to find games for a specific team or helper. Past games already sit in a collapsed section, but their separation and muted appearance need a clearer, consistent treatment. The footer's default blue legal links also clash with the existing navy footer.

## What Changes

- Add combinable filters for playing team, responsible team, and assigned person to the home game overview, with a clear way to reset them and a useful empty state.
- Keep past games collapsed by default, show them when expanded, and refine their muted styling and section control to fit the existing visual language.
- Style the shared footer's Impressum and Datenschutzerklärung links to fit its dark background, including hover and keyboard focus states.
- Preserve the public schedule's privacy boundary: guests can filter by names already present on assigned game cards, without a roster or person IDs in the response.

## Capabilities

### New Capabilities

- `schedule-overview`: Filtering and presentation of upcoming and past home games.
- `public-page-presentation`: Visual treatment and accessibility of the shared legal footer links.

### Modified Capabilities

None. Existing schedule access and assignment authorization requirements continue to apply.

## Impact

The plan affects the schedule rendering in `webapp.py`, `templates/schedule.html`, shared footer in `templates/base.html`, and `static/style.css` (with `static/app.js` if filtering is client side). It requires focused web UI tests and no schema change, migration, external dependency, or new public endpoint.
