## 1. Filter the schedule data

- [x] 1.1 Parse playing-team, responsible-team, and assigned-name GET filters on the existing overview route; verify empty, malformed, and unknown team values behave as specified in focused route tests.
- [x] 1.2 Apply all active criteria before date/month grouping while preserving `db.game_sort_key()` and `common.effective_today()` behavior; verify combined filters, unassigned responsible teams, duplicate helper names, and past-only matches in focused tests.

## 2. Present the filtered overview

- [x] 2.1 Add labeled filter controls, retained values, and a clear-filter link to `schedule.html`; verify guests and signed-in viewers can submit the GET form without JavaScript and that guest HTML contains no roster, person IDs, or contacts.
- [x] 2.2 Render distinct no-games and no-matches messages, suppress empty headings, and show the matching past-day count in the collapsed section; verify all three states and expand/collapse behavior in page tests.

## 3. Finish the visual treatment

- [x] 3.1 Style the overview filters and past section with existing tokens and readable muted colors; verify desktop and narrow-screen layouts visually and check keyboard use of the disclosure control.
- [x] 3.2 Style Impressum and Datenschutzerklärung links in the shared footer with readable default, hover, and focus states; verify both links remain visible and operable on overview and other public pages at desktop and narrow widths.

## 4. Integration checks

- [x] 4.1 Run `test/run_tests.sh` and verify the full offline suite passes without changing assignment authorization or guest privacy behavior.
- [x] 4.2 Update `README.MD` to explain the three overview filters and past-games control; verify the instructions match the rendered German UI.
