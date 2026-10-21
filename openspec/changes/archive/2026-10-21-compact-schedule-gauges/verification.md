# Implementation verification

Verified on 2026-10-21 with Chromium 155.0.8059.84. All data, contacts, accounts,
configuration and database files come from the synthetic test helpers. No
production configuration, providers or stored user data were used.

## Layout measurements

The same saved fixture data was rendered with the original template/CSS from
commit `1859deb6fae6d5177113eb8efc4e984e14f36af4` and the implemented assets.
Heights below are collapsed native `summary` bounding boxes in CSS pixels,
rounded to one decimal. `5/8` is the representative partially staffed adult game;
the preparation block has one of three positions occupied. The full measurements
for all 16 cards are in [layout.json](verification/layout.json).

| Viewport | Text size | Adult game before → after | Preparation before → after |
| --- | --- | --- | --- |
| 320 px | 100% | 258.2 → 194.2 | 187.0 → 152.0 |
| 320 px | 200% | 588.6 → 453.2 | 355.2 → 317.0 |
| 360 px | 100% | 210.6 → 174.3 | 187.0 → 152.0 |
| 360 px | 200% | 545.4 → 453.2 | 355.2 → 278.6 |
| 390 px | 100% | 210.6 → 174.3 | 187.0 → 152.0 |
| 390 px | 200% | 545.4 → 429.0 | 355.2 → 278.6 |
| 700 px | 100% | 210.6 → 174.3 | 187.0 → 152.0 |
| 700 px | 200% | 326.2 → 274.7 | 281.2 → 234.6 |
| 900 px | 100% | 210.6 → 174.3 | 187.0 → 152.0 |
| 900 px | 200% | 326.2 → 274.7 | 281.2 → 234.6 |
| 1280 px | 100% | 141.6 → 86.3 | 118.0 → 80.0 |
| 1280 px | 200% | 268.4 → 184.4 | 180.2 → 106.6 |

Every viewport/text combination passed the summary and page overflow checks,
ring/badge overlap checks, Details captions in the lower right corner without staffing overlap,
native keyboard/touch expansion, visible focus,
collapsed initial state and positive-capacity gauge checks without JavaScript.
Badges use their natural text width and flow directly beside each other with
a 6 px horizontal gap. They wrap when the next badge no longer fits; there are
no reserved columns. A one-off geometry check confirmed tight spacing and no
premature wraps for every viewport/text combination.
Each ring has a 6 px outer margin for breathing room above, below and beside
it. A sole badge occupies only its content width; a full container hides that area
without reserving rows. Long identities and independent age/setup warnings grow
naturally. No summary has a percentage row or a separate full-width disclosure control.
On phones, bottom padding reserves room for the lower-right caption beneath
the staffing area; desktop cards reserve room beside it. A one-off geometry
check confirmed a 12 px right inset and 10 px bottom inset, without status
overlap, at every viewport/text combination.

At 360 and 1280 px, delayed application scripts retained the initial counts,
arcs and badges exactly. Opening a card and loading its candidates left the
complete database hash unchanged and sent no POST. Reduced motion removed arc
transitions; responsive filter defaults remained unchanged. The `9999/10000`
cake fraction remained inside its expanded ring without page overflow. Statistics
badge dimensions, colors and text matched the original CSS at both widths.

Screenshots isolate the dated cards, hiding the fixed navigation during capture:

- [360 px: preparation, positive cakes, partial game and cleanup](verification/schedule-360-100.png)
- [1280 px: information and staffing side by side](verification/schedule-1280-100.png)
- [320 px: all adult duties vacant and explicit zero cakes](verification/all-vacant-zero-320-100.png)
- [320 px: full youth game with retained duty and unconfigured cakes](verification/full-unconfigured-320-100.png)
- [320 px / 200% text: long identity, long responsible team and age warning](verification/long-warning-320-200.png)

Reproduce with the preinstalled Playwright/Chromium as documented in
`test/README.md`. For the comparison, export the original `templates/schedule.html`
and `static/style.css` into a temporary directory and pass it as `GAUGE_BEFORE_DIR`.
The opt-in browser runner is `test/browser_schedule_gauges.mjs`; its loopback
fixture is `test/schedule_gauge_fixture.py`.

## Delta scenario review

Every scenario in `specs/schedule-overview/spec.md` was reviewed. The following
regressions and browser observations cover their specified observable behavior;
no ring, vacancy, responsive, privacy or saved-state requirement is omitted.

| Scenario | Evidence |
| --- | --- |
| Open and close a game | `test_webapp.py`, native keyboard/touch browser checks |
| Responsible team remains visible in a collapsed game | `test_schedule_filters.py` (responsible name and open placeholder) |
| Open a day-task block | `test_day_blocks_web.py`, native block expansion browser checks |
| Youth duties and retained assignments | `test_game_duties.py`, full-youth screenshot |
| Compact disclosure stays accessible | `js_assignment_confirmation.mjs` via `test_javascript.py`, browser keyboard/focus/touch checks |
| Optional support does not complete a game | `test_schedule_filters.py` (3/5 with retained Kasse, 40 dash offset) |
| Block coverage | `test_day_blocks_web.py` (1/3, exact arc and grouped preparation vacancy) |
| Assignment changes progress | `js_schedule_progress_events.mjs`, `js_age_replacement.mjs` |
| Adult additional duties count | `test_game_duties.py`, `js_schedule_progress_events.mjs` |
| Retained removed duty released | `js_schedule_progress_events.mjs`, existing retained-duty server regressions |
| Replacement wanted is still occupied | `js_replacement_cancellation.mjs`, `js_age_replacement.mjs` |
| Empty and full gauges | `test_schedule_filters.py`, `js_schedule_progress_events.mjs` (offsets 100/0) |
| Full game lacks an adult seller | `test_age_reporting.py` (8/8, no vacancy chips, separate warning) |
| Timing eligibility is unresolved | `test_age_reporting.py`, long-warning browser fixture |
| Stored data changes compliance | `test_age_reporting.py`, existing saved staffing/eligibility update path |
| Guest sees eligibility status | `test_age_reporting.py`, `test_mobile_schedule_statistics.py` |
| Configured cake block is expanded | `test_cake_blocks_web.py`, `js_cake_blocks.mjs` (four independent positions) |
| Cake block needs setup | `test_cake_blocks_web.py`, unconfigured screenshot (no progressbar/vacancies) |
| Administrator adjusts delivery settings | `test_cake_blocks_web.py`, `js_cake_blocks.mjs` (saved settings and refused/stale reductions) |
| No cakes requested | `test_cake_blocks_web.py`, `js_cake_blocks.mjs`, zero-cake screenshot |
| Remaining cake contributions are grouped | `test_cake_blocks_web.py`, `js_cake_blocks.mjs` (shared saved count/vacancies) |
| Several roles are vacant | `test_game_duties.py` (same ordered grouped labels as statistics) |
| Game type determines offered vacancies | `test_game_duties.py` (adult, youth, GE, SPF and retained Kasse) |
| Fully occupied container | `test_schedule_filters.py`, `test_age_reporting.py`, full-youth screenshot |
| Cancellation and failed claim | `js_assignment_confirmation.mjs` (Cancel, Escape, outside/close), `js_schedule_progress_events.mjs` |
| Staffing refresh reflects another saved change | `js_schedule_progress_events.mjs` (game/block candidates and concurrent block occupancy), `js_cake_blocks.mjs` |
| Guest sees public staffing needs | `test_game_duties.py`, `test_day_blocks_web.py`, `test_cake_blocks_web.py` (initial HTML and private-data exclusions) |
| Phone with several missing roles | one-off badge flow/spacing checks, all-vacant fixture at all widths |
| Wide display | 1280 px badge flow/spacing checks and desktop screenshot |
| Long names and text enlargement | all viewport/text browser bounds and overlap checks; long-warning screenshot |
| Few or no vacancies | hidden empty badge area; single content-sized badge; partial/full screenshots |

## Final validation

- `test/run_tests.sh`: passed after the coherent implementation steps and again
  after the final accessibility and responsive lower-right disclosure alignment, ring spacing and naturally wrapping badge adjustments;
  every test module exited successfully.
- `node test/browser_schedule_gauges.mjs` with the documented local runtime:
  passed all twelve viewport/text combinations plus both delayed-script and
  statistics comparisons.
- `openspec validate compact-schedule-gauges --strict`: passed.
- `git diff --check`: passed.

No schema migration, runtime dependency or domain-writer change is required.
