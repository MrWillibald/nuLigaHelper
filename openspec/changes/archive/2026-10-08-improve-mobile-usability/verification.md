# Implementation verification

Verified on 2026-10-08 using synthetic SQLite data from `test/helpers.py`,
Flask test clients, and locally served captured responses. No production
configuration, contacts, databases or notification providers were used.
Chrome used the existing bundled Playwright runtime; external browser requests
were blocked. No application dependency was added.

## Initial browser and visual checks

Schedule, helper management and statistics were checked at 320×780, 390×844,
700×900, 701×900, 1280×900 and landscape 844×390 CSS pixels. Initial responsive
defaults and fully expanded content passed at each size. Screenshots of mobile,
desktop-breakpoint, error and overlay states were inspected. At 320px, compact
headings and minimum-width corrections keep helper forms, management controls
and badges within the viewport.

- Native keyboard activation opens statistics sections independently. Deliberate
  choices survive crossing the 700px breakpoint.
- A returned person-validation error opens only the affected maintenance area;
  submitted name/contact values remain available. Editing and rotation preserve
  the draft and focused input.
- Active filter summaries remain visible outside closed mobile forms, including
  unknown-value and empty-result examples.
- The return arrow appears after one viewport, is at least 44×44px, respects the
  CSS safe-area offsets and returns keyboard focus to the navigation. Reduced
  motion uses immediate scrolling.
- Feedback, task-help popovers and membership dialogs suppress the arrow in
  portrait and landscape. Existing overlays remain readable and operable.
- With JavaScript disabled at 320px, introductions and filters used their
  expanded fallback. Maintenance was expanded in this initial check; the later
  maintenance-default follow-up below supersedes that initial state. Team
  statistics starts open; other sections open through native controls. The
  footer anchor remains usable.

## Offline regression checks

`test/js_mobile_presentation.mjs` checks both initial viewport modes, independent
activation, programmatic-toggle handling, dirty/focused fields, validation state,
resize, scroll visibility, focus, reduced motion and overlay suppression. New
disclosures make no candidate requests. Existing candidate-loading, feedback and
task-help JavaScript checks pass.

`test/test_mobile_persons.py` and `test/test_mobile_schedule_statistics.py` cover
actual response contents, tier-restricted filters/actions, complete memberships,
duplicate-name identities, escaped/unknown filter values, past-day filtering,
validation replay, native form fallbacks, statistics section order/defaults,
zero states and one count per affected container. Existing membership, contact,
birth-date, staffing and filtering regressions remain authoritative.

README labels and defaults were checked against rendered pages. Final diff
review found no schema, assignment, staffing-calculation or authorization changes;
backend changes only supply filter presentation and authorized validation replay.

Final checks passed before the caption follow-up below:

- `test/run_tests.sh`: 402 checks across 51 test modules, exit status 0.
- `openspec validate improve-mobile-usability --strict`.
- `git diff --check`.

## Caption follow-up

Disclosure captions now use concise, stable content labels: "Hinweise", "Filter",
"Bearbeitung", "Details", "Vergangene Spieltage" and the existing statistics
section titles/counts. The same caption applies in both expanded and collapsed
states; native disclosure markers and accessibility state still communicate
expansion. README and current change artifacts reflect this caption-only follow-up.
Responsive defaults, disclosure behavior, warnings and permissions are unchanged.
The revised captions were checked in templates and rendered-response regressions.
`test/run_tests.sh` passed again with 402 checks across 51 modules; strict OpenSpec
validation and `git diff --check` also passed. No new interaction or layout behavior
was introduced by this caption-only adjustment.

## Details indicator and maintenance-default follow-up

Schedule game/day cards now pair the stable "Details" caption with a triangle
pointing right when closed and down when open, driven by native disclosure state
so the indicator works without JavaScript. Authorized helper maintenance now
starts closed on every viewport and without JavaScript, while native expansion
keeps actions reachable. Only the affected entry returned with field validation
errors starts open, with its submitted values available for correction.

README, proposal, design, specs and the existing completed task descriptions
reflect these requested presentation changes. All 17 task checkboxes remain
completed. Introductory/filter/statistics defaults and action permissions are
unchanged. Targeted Chrome checks passed at 320px and 1280px: game/day triangles
change direction on native keyboard activation, maintenance starts closed, manual
opening survives resize, and only the affected validation-error entry starts open.
The triangle and maintenance/error behavior also passed without JavaScript;
the mobile triangle placement was inspected visually.

`test/run_tests.sh` passed with 402 checks across 51 modules. Strict OpenSpec
validation and `git diff --check` passed again.
