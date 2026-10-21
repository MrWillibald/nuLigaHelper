# Design

## Context

See `proposal.md` for motivation and `specs/schedule-overview/spec.md` for the behavior contract. This design spans the Flask view data, Jinja summaries, responsive CSS and saved-state JavaScript updates, so explicit presentation and update decisions are needed before implementation.

`templates/schedule.html` shares a `coverage` macro across game and day-block cards. It currently renders a long count label, a percentage and a horizontal bar. Below 900 CSS pixels the summary switches to one column; `.game-info` and `.game-summary-side` also stack their children, increasing card height. The statistics page already renders `.chip.chip-gap` badges with role counts from `db.staffing_status(game)["vacancies"]`.

Game mutation and candidate responses already include authoritative staffing with occupied/required counts, vacancies and eligibility deficiencies. Block mutations already return `block_view(block)` with saved progress, settings and slots. Preparation/cleanup candidate responses currently omit the block view, whereas cake candidate responses include it. `setCoverage()` updates bar width; `updateCoverage()` derives block progress from a local occupancy delta; `renderCakeBlock()` handles only cake-specific rendering. These paths need one consistent summary update without changing domain writers.

The shared presentation spec requires a visible `Details` caption and disclosure triangle. Preserve both within the compact layout. Assignment confirmation, replacement tokens, locks, lazy candidates, guest privacy and eligibility evaluation remain existing constraints.

## Goals / Non-Goals

**Goals:**

- Keep a single occupancy meaning across ring, count, badge group and accessible description.
- Render a complete initial summary on the server, then update it from saved response data without optimistic claim display or per-role database lookups.
- Reduce normal collapsed-card height while allowing long names, warnings and enlarged text to reflow naturally.

**Non-Goals:**

- Change assignments, capacity rules, notifications, statistics layouts or database schema.
- Add polling, real-time synchronization, a chart library or additional roster payloads to the initial schedule.
- Introduce fixed card heights, abbreviated duty names, warning truncation or reserved blank badge rows.

## Decisions

### 1. Render a small SVG ring with a central fraction

Replace the inner bar markup in the shared coverage component with a roughly 44–48 CSS pixel ring, a neutral track and an arc derived from the occupied/required ratio. Use inline SVG circles with presentation attributes for the initial arc and a single update helper for subsequent values. Keep the existing navy/neutral palette for occupancy; red remains the missing-duty badge color. This avoids introducing another success/failure meaning into a full ring when eligibility is unresolved.

Retain a `progressbar` semantic container with minimum, maximum, current value and a contextual German accessible value text such as `5 von 8 Pflichtdiensten besetzt`. Decorative circles are hidden from assistive technology. The visible fraction is central; do not repeat the long count or a percentage on another row. Gauge geometry is noninteractive. Any transition must respect reduced-motion preferences.

SVG provides a predictable initial arc and explicit geometry for zero/full occupancy without a library or client initialization. A CSS-only conic gradient was considered, but SVG makes the same ring and its updates straightforward across supported browsers. A half-circle gauge uses more horizontal space and does not match the agreed ring design.

For unconfigured cakes or a configured quantity of zero, hide the complete ring-and-vacancy component and use the existing setup or no-cakes status. A hidden zero-total component must not expose an invalid zero-maximum progressbar. Keep unusually large valid cake counts readable by letting the fraction's presentation adapt rather than clipping it.

### 2. Reuse domain vacancy data and saved block views

Include the existing staffing `vacancies` mapping in the game's initial view data. Render game badges in `db.ROLE_SLOT_COUNT` order, visiting each role once and displaying only positive vacancies in the offered set. Reuse the statistics `.chip.chip-gap` treatment and label convention: the role name alone for one missing position, `role (N×)` for several. Do not iterate the repeated UI `SLOT_LABELS` list or infer missing roles from gauge totals.

Extend `block_view()` with public summary vacancy data calculated from the same saved occupied slots and `db.block_capacity()` as progress. Each block produces at most one grouped vacancy entry labelled by its existing block label. A positive, configured cake block uses the requested capacity; zero/unconfigured states produce no invented vacancies. These fields are count/label data only. Preserve guest rendering boundaries: rendering these fields must not serialize the private block view, person IDs or slot identifiers into a public summary.

There is no need for a new vacancy endpoint or new counting policy. Game status and block views already supply the authoritative facts, while client-only role counting would duplicate offered-role rules and risk treating replacement requests as vacancies.

### 3. Update all summary elements from one saved snapshot

Adapt `setCoverage()` to update fraction, arc, values and accessible text. Make `updateGameStaffing()` update both coverage and vacancy badges before applying its separate eligibility indication. Use this path for saved claims, takeovers, request opening/closure and existing candidate refreshes. A takeover or replacement-state change preserves physical occupancy and therefore preserves vacancies.

Introduce a common block-summary updater using `result.block.progress` and its vacancy data for preparation, cleanup and cake delivery. Invoke it before cake-specific settings/slot rendering; `renderCakeBlock()` can keep its phase guard. Replace local block occupancy-delta updates when a saved block view is supplied, so the ring and badge cannot disagree after concurrent changes. Include the block summary/view in preparation and cleanup candidate responses as well, within the existing authorized candidate request, to refresh both summary elements together.

Build dynamic badges using text nodes and stable role/block ordering. Empty badge containers must not reserve layout space. Preserve existing mutation-pending and candidate-request freshness protections: an older candidate response must not overwrite a newer saved mutation. Do not update on selection, dialog opening, cancellation or dismissal. A refused write may retain the previous saved snapshot or refresh authoritative current state through existing recovery; it must never display the attempted assignment as saved. Do not add an automatic background refresh or a new write on card opening.

Keeping the existing block delta calculation was considered, but it can miss another helper's saved claim and has no single source for the corresponding badge count. Existing block responses make a common saved-state renderer feasible without changes to assignment semantics.

### 4. Use a content-sized wrapping badge group and a compact disclosure summary

Use a summary layout with information, a shared staffing area and the `Details` caption/triangle. On wide displays place information and staffing beside each other. On phones allow the information above staffing; keep ring and badge group side by side as available width permits. Give shrinking text/grid tracks `min-width: 0`, and allow full names to wrap. Time and matchup/block label share their first line when readable; metadata and responsible-team identity remain in compact secondary lines.

The staffing area has a fixed small ring track and a flexible badge group. Give the ring a 6 px outer margin on all sides so it has breathing room next to the badges and surrounding card information. Use a flex row with content-sized badges and a small gap. Wrap onto the next line when the next badge does not fit; never stretch badges or reserve equal-width columns. Long individual labels may wrap within their badge at narrow effective widths or with enlarged text. Use no fixed row count or minimum height. Keep the full German labels, including `Ordnungsdienst`; do not rely on hover to reveal a truncated role.

Reduce summary padding and inner line gaps from the current 18/20 px padding and 7–12 px stacking gaps, using roughly 10–12 px card padding and 4–6 px inner gaps as initial tuning values. Keep the caption and triangle in existing summary space, with the whole native summary as the touch target and a visible keyboard focus outline. Keep the caption in the lower right corner of every summary, including phone and desktop layouts. Reserve horizontal space beside the staffing area on desktop and bottom padding below it on phones so the caption cannot overlap badges or warnings. Preserve the established right/down triangle behavior without JavaScript. Age/classification and setup warnings remain independently readable in the collapsed summary and can increase its height when necessary.

```text
Phone (schematic; parentheses stand for the ring)
+------------------------------------------------+
| 14:30  TuS Raubling - Gastverein                |
| BL M · Halle 1 · Nr. 123                        |
| Verantwortlich: Supporter                      |
| (3/8)  [Zeitnehmer] [Verkauf (2x)]              |
|        [Kasse]      [Reinigung]      > Details  |
+------------------------------------------------+
```

Content-sized badges flow in reading order beside the ring. Natural line wrapping uses the available width without the gaps created by equal-width grid columns. A fixed maximum card height would hide names or warnings, so rows remain content-driven.

### 5. Verify observable state and actual layout

Update existing schedule assertions and JavaScript fixtures around progress to cover the ring and badges together. Keep tests focused on behavior: adult/youth/SPF offered duties, retained roles, grouped vacancies, replacement requests and takeovers, confirmation cancellation, rejected/stale writes, saved refreshes and cake capacity changes. Reuse synthetic offline fixtures; protect guest HTML and no-JavaScript initial content.

Perform browser layout verification with synthetic games at 320, 360, 390, 700, 900 and a desktop width around 1280 CSS pixels, including 200-percent text enlargement, long matchup/responsible-team names, every offered role vacant, one vacancy, full occupancy, an age deficiency, each block phase and zero/unconfigured cakes. Verify no horizontal overflow, clipping or overlap; independent keyboard/touch expansion; visible focus and `Details`; and unchanged initial filter/disclosure behavior. Check reduced motion and delayed/disabled JavaScript. Compare representative collapsed-card heights with the pre-change layout at the same widths/data and record the result; common mobile cases should be shorter, while exceptional long content may grow for readability.

## Risks / Trade-offs

- [A six-role vacancy group can exceed the ring's height] → Allow content-driven rows and tune spacing; never hide required staffing information to achieve a fixed height.
- [Nested page/day/card padding leaves less width than expected at 320 px] → Verify actual component widths, reduce redundant summary spacing and let badges flow onto additional lines and long labels wrap.
- [A full ring is mistaken for valid staffing] → Describe the ring as physical occupancy and retain the separate visible eligibility indication; do not add an unconditional complete/success label.
- [Partial or old response handling updates only the ring or only badges] → Route saved game and block snapshots through shared summary updaters and cover response ordering and failure recovery in existing JavaScript regressions.
- [Global badge/disclosure CSS changes affect statistics or other pages] → Reuse shared chip colors but scope badge layout and summary density to schedule cards, then check statistics as a regression.
- [Gauge markup or dynamic badges lose accessible meaning] → Preserve progress semantics, contextual German text, text-node rendering and a labeled native disclosure; verify initial and updated states.

## Migration Plan

Deploy the Flask view/response changes, schedule template, CSS and JavaScript together through the existing deployment process. No schema migration, backfill or provider configuration is required. Run the offline suite and responsive verification before deployment. Rollback restores the previous application assets and view builders together; assignments and audits require no conversion.
