## 1. Server-Side Candidate Ordering

- [x] 1.1 Add a per-game candidate categorization and sort helper in `webapp.py` using playing-team, responsible-team, Supporter, other-team precedence and normalized name plus person ID ordering; verify focused helper or schedule tests cover every category, absent responsible teams, overlapping roles, and duplicate names.
- [x] 1.2 Apply the ordering only after the existing admin, MV, and member candidate restrictions are calculated for each slot; verify webapp/refusal tests still show only candidates authorized for each access tier.

## 2. Dropdown Rendering and Live Updates

- [x] 2.1 Render the server-provided category, normalized-name, and ID sort metadata on person options while preserving selected values and the existing playing/outside classes, titles, and visible German hints; verify schedule HTML assertions cover ordering, warnings, and exclusion from other slots.
- [x] 2.2 Update `static/app.js` to reinsert released-person options by the rendered category/name/ID metadata; verify a focused client-side test or equivalent DOM-level test demonstrates restoration to the correct category and alphabetical position with warning metadata intact.

## 3. Verification

- [x] 3.1 Add regression coverage for all four ordered categories, alphabetical ordering within categories, already-assigned-person exclusion, and current-occupant visibility; run the focused schedule/webapp test files successfully.
- [x] 3.2 Run `test/run_tests.sh` and verify the complete offline suite passes with both debug switches left `False`.
