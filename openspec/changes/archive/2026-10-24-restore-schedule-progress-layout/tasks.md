# Tasks

## 1. Restore progress rendering and saved-state updates

- [x] 1.1 Replace the shared schedule coverage macro's SVG rings and vacancy badges with descriptive German counts, percentages and horizontal bars, retaining contextual accessible values and hidden zero/unconfigured cake states; update the existing schedule/game/day-block/cake/age-reporting Python regressions and verify initial guest and signed-in HTML has correct current totals, no badge/ring presentation and no private-data regression.
- [x] 1.2 Update `setCoverage()`, game staffing and common block summary rendering to set count, rounded percentage, bar width and accessible state from authoritative saved snapshots; remove schedule badge updates while preserving candidate-response freshness and current API fields, and adapt the existing JavaScript progress/candidate/cake/confirmation/replacement regressions to verify claims, takeovers, unchanged request occupancy, cancellation, rejection, stale responses and setup/zero/positive cake transitions.
- [x] 1.3 Update `README.MD` to describe the restored count/percentage/bar display and separate eligibility warnings; verify its wording matches the delta spec and retains seven required adult positions, one Wischer and current age rules.

## 2. Restore responsive card layout and reference spacing

- [x] 2.1 Restore the historical summary padding, row gaps, desktop information/progress columns, at-most-900-pixel stacking and 1.5-rem time emphasis; put right-aligned `Details` beneath progress/status in normal flow, preserve overflow and past-disclosure fixes, and verify game, preparation, cleanup and cake summaries retain their information order and independent keyboard/touch expansion without JavaScript.
- [x] 2.2 Adapt the existing browser layout harness and synthetic schedule fixture to bars, updating names and `test/README.md` commands together if renamed; run at 320, 360, 390, 700, 900, 901 and approximately 1280 CSS pixels, including 200-percent text enlargement, long names, eligibility/classification warnings, partial/empty/full staffing, cake setup/zero/positive states and past cards, and record screenshots/results showing no overflow, clipping, overlap or hidden focus.
- [x] 2.3 Compare representative mobile output with the user's reference screenshot and historical pre-`98694d3` spacing, and desktop output with the historical two-column layout; record the comparison in this change's verification notes and confirm existing filter/introduction defaults, delayed/disabled JavaScript behavior and statistics badges remain intact.

## 3. Verify the integrated change

- [x] 3.1 Run `test/run_tests.sh` and the affected JavaScript regression commands documented in `test/README.md` after implementation, resolve failures and record results; verify binding confirmations, replacement tokens, day-before locks, guest privacy and current eligibility behavior stay green alongside the restored presentation.
- [x] 3.2 Run `openspec validate restore-schedule-progress-layout --strict` and inspect the final diff for presentation-only scope; confirm no schema, domain writer, notification or API-contract rollback was introduced and that the implementation and verification evidence satisfy every delta requirement.
