# Proposal

## Why

Long introductory text, stacked filter controls and fully expanded person-edit forms make the mobile interface difficult to browse as the roster and season grow. Selective disclosure and a shared return-to-top action will make the existing information and actions easier to reach.

## What Changes

- Add a floating arrow with the German accessible label "Nach oben", appearing after approximately one viewport of scrolling on shared pages.
- Collapse general introductory descriptions on the schedule, helper-management and statistics pages on mobile, behind "Hinweise"; keep them expanded on desktop. Operational warnings and validation errors stay visible.
- Make statistics sections independently collapsible with meaningful result counts. Preserve the order: "Spiele pro Mannschaft", "Dienste pro Person", then "Offene Dienste, Altersanforderungen und Einrichtung". All start collapsed on mobile; only "Spiele pro Mannschaft" starts expanded on desktop.
- Make helper cards show name, every team and applicable status first, with authorized maintenance actions initially collapsed behind "Bearbeitung" on mobile, desktop and without JavaScript. Reopen the affected card on validation failure.
- Collapse schedule and roster filter forms on mobile behind "Filter", retaining a readable summary of applied filters and a clear-filter action outside the disclosure. Desktop filters remain expanded.
- Keep disclosure captions concise and stable in both states: "Hinweise", "Filter", "Bearbeitung", "Details" for game/day cards, "Vergangene Spieltage" and the statistics section titles, without action verbs.
- Add a triangle beside "Details" on schedule game/day cards, pointing right when closed and down when open, including without JavaScript.
- Preserve keyboard, touch, reduced-motion and no-JavaScript usability, using the existing mobile breakpoint and native disclosure controls.

## Capabilities

### New Capabilities

- `statistics-overview`: Section order, responsive initial expansion, independent disclosure, counts and empty states for the statistics page.

### Modified Capabilities

- `public-page-presentation`: Shared return-to-top navigation and responsive introductory-text disclosure.
- `schedule-overview`: Compact mobile filters with applied-filter visibility and existing filtering semantics.
- `access-control`: Compact roster maintenance on all viewports and mobile filter presentation within existing person-data and action permissions.

## Impact

- Templates: `base.html`, `schedule.html`, `persons.html` and `statistik.html`.
- Shared presentation and interaction: `static/style.css` and `static/app.js`; check coexistence with `static/feedback.js` and task-help popovers.
- `webapp.py` only if needed for safe filter-summary presentation or validation-return context; no new API, dependency or database migration is expected.
- Extend offline presentation/interaction verification and update `README.MD` for the new UI behavior.
- This scope excludes shortening game matchup/context summaries, collapsible match days, collapsible administration panels, pagination and navigation redesign. Existing staffing calculations, task descriptions, authorization and privacy remain authoritative.
