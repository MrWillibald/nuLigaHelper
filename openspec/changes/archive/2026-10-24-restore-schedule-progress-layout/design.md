# Design

## Context

See `proposal.md` for motivation and `specs/schedule-overview/spec.md` for the behavior contract. The shared `coverage` macro in `templates/schedule.html` currently renders SVG rings and vacancy badges. `static/app.js` updates those through `setCoverage()`, `updateGameStaffing()` and `updateBlockSummary()`. Game staffing and block views already carry authoritative occupancy; block candidate responses also return the saved block view.

Commit `98694d3` introduced the rings, badges and compressed spacing. Commit `db8abaa` later changed information stacking and disclosure placement. The presentation reference is the parent of `98694d3` (`1859deb`), corroborated by the user's `1000014311.png`: a prominent time above the title, metadata and optional responsible team, followed on mobile by descriptive occupancy, percentage, a horizontal bar and right-aligned `Details`. This is a selective presentation restoration on the current application.

The design is needed because templates, CSS and live updates must change together; reverting markup alone would leave JavaScript targeting SVG elements. Current seven-position adult staffing, one Wischer, revised age eligibility, sticky duties, replacement generation checks and guest privacy all remain constraints.

## Goals / Non-Goals

**Goals:**

- Reuse the earlier bar structure and spacing while retaining current saved-state update paths.
- Keep server-rendered and dynamically updated labels, percentages, widths and accessible values consistent.
- Preserve independent warnings and natural reflow across all card types and past/upcoming dates.

**Non-Goals:**

- Revert complete commits, restore old staffing counts, modify domain writers, change response contracts or remove statistics badges.
- Change navigation, page-wide disclosure defaults, filtering, candidate loading, confirmation flows or cake configuration behavior.
- Introduce new libraries, schema changes or fixed card heights.

## Decisions

### 1. Restore the shared horizontal component

Replace the macro's SVG and badge markup with `.coverage-text` containing the descriptive count and percentage, followed by `.coverage-track` and `.coverage-fill`. Keep existing progress data attributes and a shared renderer for games, preparation, cleanup and cakes. Use the wording defined in the delta spec; the count describes occupancy, never eligibility compliance.

Put progress semantics on the track, with minimum, current, maximum and contextual German accessible text. Render values and width on the server first. Hide the entire component when cake setup is incomplete or capacity is zero, and do not expose a zero-maximum progressbar. Retain reduced-motion support for width transitions.

Alternative: adding a CSS bar alongside the ring would retain obsolete markup, redundant accessible content and unnecessary update paths. Restore one component instead.

### 2. Use historical layout values as the starting point

Restore the old summary grid: `minmax(240px,340px) minmax(0,1fr)` with a 22 px gap and 18/20 px vertical/horizontal padding. At at most 900 CSS pixels use a single column with a 12 px gap. Restore 7 px information-row gaps, 8 px spacing in the progress/status column, a 10 px rounded track with 8 px top margin, and the earlier 1.5 rem time emphasis. Keep information stacked on desktop as in the old reference; do not reintroduce the compact time/title row or increase time size further.

Move `Details` into the progress/status column after coverage and eligibility/setup messages, right-aligned in normal flow. Remove the absolute positioning, extra desktop right reservation and artificial mobile bottom reservation introduced for rings/badges. Keep full-summary keyboard/touch activation and the right/down triangle.

Preserve useful overflow protections such as shrinking tracks, wrapping long names/date qualifiers, bounded cake inputs and the direct-child `.past-details > summary` selectors. The latter prevent past-section styling from affecting nested game summaries. Keep page, month and day framing and the shared icon/color treatment. The historical values are implementation starting points; natural wrapping at narrow effective widths takes priority over fixed dimensions.

Alternative: restoring the complete historical stylesheet would discard later fixes unrelated to the requested appearance. Selective CSS/template edits restore the reference while retaining them.

### 3. Retain authoritative updates and remove only badge presentation

Adapt `setCoverage()` to update the long count, whole-number percentage, fill width and accessible state together. Match `progress_data()` rounding: `(filled * 100 + floor(total / 2)) / total`, taking the integer floor for positive totals. Use the same displayed percentage for initial and live bar widths, with an explicit zero-total guard.

Keep `updateGameStaffing()` and `updateBlockSummary()` driven by saved staffing/block snapshots, including snapshots returned by candidate refreshes. Preserve mutation-pending and response-freshness guards. Do not restore the older block occupancy-delta calculation: it can lose concurrent saved changes. Request opening/closure and takeovers keep occupancy unchanged; saved eligibility changes still refresh the separate warning.

Remove schedule-only `updateVacancies()` calls and rendering, including client role-ordering work used solely for badges. Keep domain vacancy calculations and existing response fields compatible; there is no need to change `webapp.py` response contracts or shared statistics chips. Selection or confirmation dismissal still performs no write or speculative progress change.

Alternative: reverting the original gauge commit's entire JavaScript/backend diff would unnecessarily remove authoritative block refresh improvements and risk newer duty behavior.

### 4. Verify behavior and the reference presentation

Adapt the existing progress fixture and Python/JavaScript assertions to the bar contract. Reuse current cancellation, replacement, candidate-refresh, age-reporting and cake cases; avoid tests that merely mirror markup. Cover adult seven-position and youth/SPF five-position totals, retained assignments, partial/empty/full progress, positive/zero/unconfigured cakes and read-only guest output. Verify that claims, takeovers, requests, rejected writes and stale responses retain their established saved-state meaning.

Adapt the existing `test/browser_schedule_gauges.mjs` and synthetic fixture for bar verification; rename them if useful and update `test/README.md` references together. Compare representative mobile cards with the supplied screenshot and pre-`98694d3` CSS, and desktop arrangement with the historical two-column layout. Use synthetic data at 320, 360, 390, 700, 900, 901 and approximately 1280 CSS pixels, with 200-percent text enlargement, long names, warnings, all block types and past dates. Check no clipping/overlap/overflow, readable count/percentage, visible focus, native disclosure and disabled/delayed JavaScript. There is no card-height reduction target.

## Risks / Trade-offs

- [Restored padding makes cards taller] -> Accept the reference's spacing; allow additional height for long content and warnings.
- [Count and percentage compete for space] -> Allow the text row to wrap while the track retains available inner width.
- [Hidden cake components reappear during an update] -> Apply configuration/positive-capacity visibility through the shared saved block renderer and verify transitions between setup, zero and positive quantities.
- [A full bar is read as complete eligibility] -> Keep the independent German eligibility warning visible and avoid unconditional success labels.
- [Shared CSS or broad rollback affects other pages] -> Scope changes to schedule progress/layout and preserve statistics styling, privacy boundaries and later disclosure fixes.

## Migration Plan

Ship the template, stylesheet and JavaScript together using the existing release process after offline tests and browser verification. Update `README.MD` with the restored overview presentation. No data migration, backfill, credential change or new dependency is required. Operational rollback restores the previous application assets together; database contents and audit history require no conversion.
