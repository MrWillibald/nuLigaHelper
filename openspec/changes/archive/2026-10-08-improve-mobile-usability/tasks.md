# Tasks

## 1. Shared disclosures and introductory explanations

- [x] 1.1 Add reusable native-disclosure styling and responsive initialization at the existing 700px breakpoint; verify initial defaults, independent toggles and preservation of deliberate choices, dirty forms and focused fields across resize with focused offline JavaScript interaction checks.
- [x] 1.2 Wrap only general introductions in `schedule.html`, `persons.html` and `statistik.html`; verify mobile collapse, desktop expansion and no-JavaScript accessibility, with actionable age/setup warnings, validation errors and task-help controls retaining their existing contexts.
- [x] 1.3 Verify the new disclosure attributes/listeners do not trigger game/day-block candidate requests using the existing candidate-loading regressions; update `README.MD` with the breakpoint, introductory disclosure and expanded introductory/filter no-JavaScript fallback.

## 2. Return-to-top navigation

- [x] 2.1 Add the shared arrow/link in `base.html`, progressive floating placement and scroll visibility updates; verify it appears after one viewport, hides near the top and works as a footer anchor without JavaScript using focused interaction checks and a long synthetic page.
- [x] 2.2 Handle keyboard focus, reduced-motion scrolling, safe areas and coexistence with feedback, task-help popovers and dialogs; verify at least a 44px touch target and no obstruction in narrow portrait/landscape views, and run existing feedback/task-help regressions.
- [x] 2.3 Document the return-to-top action and fallback in `README.MD`; verify the documented behavior against the rendered synthetic page.

## 3. Schedule and roster filters

- [x] 3.1 Wrap both existing GET filter forms in responsive disclosures and render escaped, readable active-filter summaries and clear/reset actions outside; verify mobile and desktop behavior, selected-value retention, unknown values and no submission merely from toggling.
- [x] 3.2 Extend existing schedule/management filtering checks for collapsed active summaries, combined filters, past-date expansion, empty results and role-restricted roster filters; verify existing guest and member response privacy and no-JavaScript submissions remain correct.
- [x] 3.3 Update `README.MD` with mobile filter disclosure and persistent active-filter summaries; verify labels and reset links match the implemented pages.

## 4. Compact helper maintenance

- [x] 4.1 Keep name, every team and already-authorized status outside each maintenance disclosure, with authorized actions inside; verify compact entries with initially closed maintenance on mobile, desktop and without JavaScript, duplicate-name separation by internal identity, MV membership access and absence of empty editing controls.
- [x] 4.2 Force the affected maintenance disclosure open on field-validation failure and preserve submitted values; verify invalid self/admin edits expose the correct errors and values without opening unrelated entries or losing them during viewport changes.
- [x] 4.3 Extend management/privacy regressions for admin, MV and member views and actual response contents; verify other people's contacts/birth dates remain absent for lower tiers, existing action permissions and dialogs work, and management still precedes the roster. Document initially closed mobile/desktop/no-JavaScript maintenance, native action access and error reopening in `README.MD`.

## 5. Statistics disclosure

- [x] 5.1 Wrap all three statistics sections in independent native disclosures, retaining their current order; verify all start closed on enhanced mobile and only "Spiele pro Mannschaft" starts open on desktop, including the no-JavaScript desktop fallback.
- [x] 5.2 Add accurately labeled team/person/affected-container counts and visible zero-count headers with expanded empty states; verify a game with multiple staffing issues contributes one affected container and existing duties, age deficiencies and setup reporting retain their current meaning.
- [x] 5.3 Extend existing statistics response checks and responsive interaction checks for independent expansion and empty datasets; document the section order and mobile/desktop defaults in `README.MD` and verify the documented defaults against rendered pages.

## 6. Integrated verification

- [x] 6.1 Inspect synthetic schedule, helper and statistics pages at 320px, 390px, 700px, 701px and a desktop width, plus landscape, keyboard navigation, reduced motion and JavaScript disabled; verify all spec scenarios, the right/down "Details" triangle, initially closed maintenance except validation errors, form-error recovery and overlay placement, and record results without production contacts/configuration.
- [x] 6.2 Run `test/run_tests.sh`, `openspec validate improve-mobile-usability --strict` and `git diff --check`; verify all pass and inspect the final diff for unintended domain, authorization, data or out-of-scope presentation changes.
