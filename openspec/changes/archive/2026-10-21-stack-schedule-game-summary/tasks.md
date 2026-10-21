# Tasks

## 1. Reorganize overview information rows

- [x] 1.1 Update the shared information layout in `static/style.css` to stack time, matchup/task label, metadata and the game-only responsible team; remove obsolete child flex sizing and increase time text modestly (initial target `1.25rem`). Verify with synthetic rendered game, preparation, cleanup and cake cards that time and title occupy separate rows even on desktop, and run `test/run_tests.sh` after the change.
- [x] 1.2 Adjust `templates/schedule.html` only as needed to place classification warnings after the four primary game rows and preserve natural wrapping of block time/date qualifiers. Verify ordinary and Spielfest game content, the open responsible-team placeholder, all block icons, cake time/quantity update attributes and the absence of any responsible-team row on blocks; run `test/run_tests.sh` after the change.
- [x] 1.3 Update the overview description in `README.MD` to document the ordered information rows and enlarged time for both games and day-task cards. Verify that the text distinguishes four game rows from three block rows and run `test/run_tests.sh` after the change.

## 2. Verify integrated presentation

- [x] 2.1 Use offline synthetic schedule data to visually verify 320, 360 and 390 CSS pixel viewports and desktop, including enlarged text, long names, unset responsibility, classification/eligibility warnings, all block types and positive/zero/unset cake configuration. Confirm separate rows, readable larger times, natural wrapping, no horizontal scrolling or overlap, and usable rings, badges and "Details"; record the checked cases and results.
- [x] 2.2 Verify upcoming and past cards for guests and signed-in viewers, collapsed and expanded summaries, keyboard disclosure, rendering without JavaScript, and saved responsible-team/cake-setting updates. Confirm that the row hierarchy persists, all existing display states remain visible and no layout interaction changes assignments; run `test/run_tests.sh` and record its result.
