# Tasks

## 1. Supply consistent saved staffing summaries

- [x] 1.1 Add game vacancy data to the initial schedule view using existing staffing results and stable role order; extend the existing synthetic schedule tests to verify adult, youth/SPF, retained-role, fully occupied and grouped-Verkauf cases agree with statistics.
- [x] 1.2 Add grouped public vacancy data to `block_view()` from saved progress/capacity and return current block summary data on authorized preparation/cleanup candidate refreshes; verify existing block/cake web tests cover three-slot blocks, positive cake capacity, zero and unconfigured cakes, including guest HTML without private identifiers.

## 2. Render accessible rings and grouped badges

- [x] 2.1 Replace the shared schedule coverage bar/percentage markup with a server-rendered SVG ring and central fraction, and render the red vacancy badges using existing chip styling; update schedule/webapp assertions to verify zero, partial and full arcs, readable contextual progress semantics, correct badge grouping and useful initial HTML without JavaScript.
- [x] 2.2 Integrate cake setup/zero states and preserve separate age/classification warnings and the labeled native `Details` disclosure; verify existing cake, eligibility and guest tests show no zero-total gauge or invented vacancy, while a full deficient game retains its warning and ordinary card expansion exposes the existing duties.
- [x] 2.3 Document the ring's occupied/required meaning, grouped missing-duty badges and distinct age/setup warnings in `README.MD`; verify the wording matches the rendered German UI and existing assignment responsibility rules.

## 3. Refresh ring and badges from saved responses

- [x] 3.1 Adapt `setCoverage()` and `updateGameStaffing()` to update fractions, arcs, accessible values and vacancy badges from the same game staffing snapshot; update `test/js_schedule_progress_events.mjs` and its runner coverage to verify saved claims, adult roles, retained-duty retirement, failed/stale writes and candidate refreshes without an unsaved visual improvement.
- [x] 3.2 Add a shared block-summary update path for preparation, cleanup and cake data and replace local occupancy-delta display updates when authoritative block data is returned; extend the existing block/cake JavaScript regressions to verify a concurrent saved claim, positive/zero/unconfigured cake transitions and refused configuration changes keep count, badge and capacity consistent.
- [x] 3.3 Route existing request-open, request-close, takeover and candidate-response handling through the common summary updaters while preserving request freshness and mutation guards; verify confirmation cancellation/Escape/dismissal leave summaries unchanged, replacement search keeps occupancy, and an older response cannot overwrite a newer saved summary using the existing synthetic JavaScript fixtures.

## 4. Reduce collapsed-card height with responsive layout

- [x] 4.1 Rework schedule-scoped summary CSS into compact information, ring/badge and disclosure areas, reducing redundant padding and stacking; verify badges sit directly beside each other and wrap only when the next badge does not fit, and one/no vacancies reserve no empty rows or separate action/percentage line.
- [x] 4.2 Preserve full names, metadata, warnings, visible focus, reduced-motion behavior and native keyboard/touch expansion while allowing narrower or enlarged content to reflow; update affected mobile/disclosure regressions and verify statistics badge layout and responsive filter defaults are unchanged.
- [x] 4.3 Capture representative before/after collapsed-card heights and layout evidence using only synthetic fixtures at 320, 360, 390, 700, 900 and about 1280 CSS pixels plus 200-percent text enlargement; record results with this change, showing reduced height for normal mobile cases and no horizontal overflow, clipping or overlap for long names, all vacant roles, an age warning and each block phase, including disabled/delayed JavaScript.

## 5. Validate the integrated change

- [x] 5.1 Run `test/run_tests.sh` after each coherent implementation change as required by repository instructions and resolve regressions; verify the final complete offline suite passes with synthetic database/configuration data.
- [x] 5.2 Run `openspec validate compact-schedule-gauges --strict` and review the final implementation against every scenario in the delta spec; verify validation succeeds and the review records no missing ring, vacancy, responsive, privacy or saved-state behavior.
