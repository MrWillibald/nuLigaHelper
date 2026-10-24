# Verification — restore-schedule-progress-layout

Verified locally on 2026-10-24 with synthetic SQLite databases and offline provider
fakes. Browser fixtures bind only to loopback; no production configuration,
contacts, messages or database files are used.

## Result and scope

The shared schedule component renders descriptive German counts, rounded whole
percentages and horizontal bars for games, preparation, cleanup and configured
positive cake quantities. Initial HTML and live saved-state updates use the same
rounding. Zero/unconfigured cakes hide the entire component and expose no
zero-maximum progressbar. Schedule rings and vacancy badges are removed; statistics
badges and existing response fields are retained.

Implementation changes are confined to `templates/schedule.html`,
`static/style.css`, `static/app.js`, documentation and verification/tests. There
are no changes to database models/migrations, domain writers, notifications,
API response contracts, candidate freshness guards, filters or authorization.

## Reference comparison

Inspected the supplied `1000014311.png` reference and compared the new mobile
screenshots with it: time, task/matchup, metadata and optional responsible team
precede the count/percentage row and full-width bar. The navy fill, pale track,
plain cards, task icons and right-aligned Details below progress match that
presentation. Data, device scale and page framing differ; this is a selective
restoration of schedule summaries, not a page-wide pixel comparison.

The harness also renders `1859deb`'s template/CSS (the parent of `98694d3`) against
the same current synthetic data. Current and historical summary padding is
18px 20px, information gap 7px, status gap 8px, bar height 10px and top margin
8px, with 1.5rem time emphasis. Summary columns use a 12px gap and stack through
900px; above 900px the historical `minmax(240px,340px) minmax(0,1fr)` columns
and 22px gap are restored. Automated comparisons confirm padding, gap and column
count at every tested width/text scale. Later page framing, wrapping protections,
explicit time line height and direct-child past disclosure selectors remain.
Small card-height differences from historical screenshots are expected; there is
no height reduction target.

Representative captures (all synthetic):

- [Mobile, 360px](verification/schedule-360-100.png) and [historical reference](verification/reference-360-100.png).
- [Desktop, 1280px](verification/schedule-1280-100.png) and [historical reference](verification/reference-1280-100.png).
- [Two columns at 901px](verification/schedule-901-100.png).
- [320px with 200% text](verification/schedule-320-200.png).
- [Long identity, responsible team, classification and eligibility warnings at 200%](verification/long-warning-320-200.png).
- [Empty staffing and zero cakes](verification/all-vacant-zero-320-100.png).
- [Full SPF staffing, retained duty and unconfigured cakes](verification/full-unconfigured-320-100.png).
- [Past cards with a previous-date preparation qualifier](verification/past-320-100.png).
- [Visible keyboard focus](verification/keyboard-focus-360.png).

## Browser results

`test/browser_schedule_progress.mjs` passed with Chromium 156.0.8078.47.
[The complete layout report](verification/layout.json) records all 14 combinations
of 320, 360, 390, 700, 900, 901 and 1280 CSS pixels with 100% and 200% root text.
The fixture contains 20 cards across four upcoming dates and one past date:
partial/empty/full games, adult seven-position and SPF five-position staffing,
a retained Kasse occupant, preparation/cleanup, positive/zero/unconfigured cakes,
long names, classification/eligibility warnings and a dated time qualifier.

Every combination passed checks for page/summary overflow, count/percentage
intersection, information order, full-width tracks, status/Details separation,
accessible values and saved count/percentage/fill consistency. Each container and
past card was independently expanded by keyboard and collapsed by touch without
JavaScript. Focus outlines remained visible, and opening a card left other cards
closed. Screenshots were visually inspected for clipping and overlap.

Guest views additionally passed at 320, 901 and 1280px with 200% text and no private
identifiers/candidate URLs or assignment controls. Delayed application scripts at
360/1280px preserved initial staffing, native disclosure and the database snapshot;
read-only candidate refreshes caused no POST. Mobile introduction/filter defaults
remain closed with JavaScript, desktop defaults open, and native no-JavaScript
filters remain available. Reduced motion removes fill transitions. Large cake
counts reflow correctly. Statistics badge text, dimensions and colors match the
historical stylesheet.

Reproduction (using already installed Playwright/Chromium, no package installation):

```bash
mkdir -p /tmp/schedule-progress-reference
git show 1859deb:templates/schedule.html > /tmp/schedule-progress-reference/schedule.html
git show 1859deb:static/style.css > /tmp/schedule-progress-reference/style.css
PLAYWRIGHT_MODULE=/absolute/path/to/playwright/index.mjs \
  BROWSER_EXECUTABLE=/absolute/path/to/chromium \
  PROGRESS_REFERENCE_DIR=/tmp/schedule-progress-reference \
  BROWSER_REPORT_PATH=/tmp/progress-layout.json \
  BROWSER_SCREENSHOT_DIR=/tmp/progress-screens \
  node test/browser_schedule_progress.mjs
```

## Regression coverage

The adapted Python regressions cover initial public/signed-in bars, seven adult
and five youth/SPF positions, retained duties excluded from progress, setup/zero
cake states, contextual accessible values and guest privacy. A fully occupied but
age-deficient game retains its 100% bar and visible independent eligibility warning.

All eight affected JavaScript commands from `test/README.md` passed:
`js_assignment_options`, `js_schedule_progress_events`, `js_candidate_loading`,
`js_candidate_toggle`, `js_cake_blocks`, `js_assignment_confirmation`,
`js_age_replacement` and `js_replacement_cancellation` (each under `test/` with
`.mjs`). They verify authoritative claims/candidate refreshes, rounded bars,
concurrent block occupancy, rejection, stale responses, cancellation/dismissal,
unchanged occupancy for replacement search/closure/takeover and cake transitions.

The existing `test/browser_binding_assignments.mjs` also passed with the same
Playwright/Chromium runtime: native confirmation/cancel/focus, replacement search
closure/reopening and atomic takeover across all duty containers, concurrent
Cancel/CAS conflict, immediate late locking, and non-JavaScript review/cancel/
confirmation. The synthetic provider/database boundaries remain intact.

`openspec validate restore-schedule-progress-layout --strict` passed, and
`git diff --check` passed. Final scope review confirmed that schema/domain writers,
notification code and API contracts are unchanged.

`test/run_tests.sh` passed twice after the presentation update; the final run
completed with exit status 0 across 56 test files and 454 PASS results. This
includes binding confirmations, replacement generations and concurrency, fixed
day-before locks, guest/privacy boundaries, current age eligibility, cake handling,
notifications and guarded migrations. Expected synthetic failure logs in negative
tests do not represent test failures.
