## 1. Schedule data and filtering

- [x] 1.1 Add chronological current-season date options and exact date filtering to `build_schedule()` and the GET route; apply person matching independently to games and blocks, suppress blocks for a responsible-team filter, and retain block-only dates. Verify combinations, invalid dates, ordering, and full-date block times in `test/test_schedule_filters.py` and `test/test_day_blocks_web.py`.
- [x] 1.2 Compute game and block staffing counts from occupied slots, excluding optional `Unterstützung` and the responsible team from game progress; verify empty, partial, and full counts in schedule rendering tests.

## 2. Overview presentation

- [x] 2.1 Add the date dropdown to the existing filter form and open the outer past section for a selected past date, including a block-only result; verify the selected value, clear-filter behavior, and past-date visibility in `test/test_schedule_filters.py`.
- [x] 2.2 Render every game and day-task block as a closed native disclosure with identity and accessible progress in the summary and existing assignment fields inside; verify guests, members, and admins see the correct collapsed and expanded content in web tests.
- [x] 2.3 Adapt schedule CSS for compact and expanded cards at desktop and phone widths; verify readable counts, keyboard focus, and no clipped fields through a visual inspection at both widths.

## 3. Live updates and completion

- [x] 3.1 Update `static/app.js` to refresh the affected progress count, percentage, and bar after confirmed game or block assignment changes; verify claim, release, replacement, optional `Unterstützung`, and failed-operation behavior with the existing DOM test setup or an equivalent interaction test.
- [x] 3.2 Document the date filter, compact cards, and required-slot progress meaning in `README.MD`; verify the instructions match the finished overview.
- [x] 3.3 Run `test/run_tests.sh` and inspect the overview as a guest and signed-in viewer; verify the suite passes and the date, disclosure, and assignment flows work together.
