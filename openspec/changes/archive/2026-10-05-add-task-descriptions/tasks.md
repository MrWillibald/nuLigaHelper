# Tasks

## 1. Complete task description catalog

- [x] 1.1 Add a developer-maintained Python description catalog keyed by supported semantic roles/phases, with concise German guidance for every current task; verify catalog coverage and identical wording across numbered positions.
- [x] 1.2 Integrate renamed/new tasks when staffing or cake delivery is adopted without presenting inactive proposal-only roles; verify coverage follows the actual supported role catalog for each application order.
- [x] 1.3 Document where developers maintain task descriptions and how role additions preserve coverage; verify the documented catalog and coverage checks exist.

## 2. Accessible information controls

- [x] 2.1 Render one small information control immediately after each schedule assignment-task label using shared markup and escaped text; verify game/day-block labels, guest visibility and no roster/contact/person-identifier leakage.
- [x] 2.2 Support hover/focus description access and click/tap persistent opening with keyboard dismissal, task-specific accessible names and correct associations; verify pointer, keyboard and touch interactions do not submit assignments or toggle unrelated cards.
- [x] 2.3 Show descriptions as floating popovers over underlying content without shifting assignment fields, cards or the page; keep text wrapped within the viewport and avoid container clipping; inspect representative desktop/mobile renders and verify hover/focus/tap layout remains stable.
- [x] 2.4 Document the information icon, floating presentation and hover/keyboard/tap behavior in README.MD; verify the described controls match the rendered interface.

## 3. Integration verification

- [x] 3.1 Run `openspec validate add-task-descriptions --strict` and the relevant schedule/privacy/interaction checks, followed by `test/run_tests.sh`; resolve integration failures and verify the new description control preserves assignment behavior.
