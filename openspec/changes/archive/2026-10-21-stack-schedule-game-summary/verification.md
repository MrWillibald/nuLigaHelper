# Implementation verification

Checked on 2026-10-21 using offline synthetic data, Playwright and installed
Chromium 156.0.8078.47. Only a loopback Flask fixture was served; production
configuration, databases and notification providers were not used.

## Layout and disclosure

- The existing `test/browser_schedule_gauges.mjs` passed at 320, 360, 390,
  700, 900 and 1280 CSS pixels with 100% and 200% root text size, including
  ring/badge overflow checks, delayed JavaScript and read-only database snapshots.
- An extended temporary fixture exercised 25 cards: ordinary adult/youth games,
  Spielfest, unresolved classification, long matchup/responsible-team names,
  unset responsibility, eligibility deficiencies, all three day-block types,
  unknown block times and adjacent-date qualifiers. Cake states included positive,
  zero and unset quantities and delivery times.
- All 32 combinations of 320/360/390/1280 CSS pixels, 100%/200% text,
  guest/admin and JavaScript enabled/disabled passed. Every card was checked
  collapsed and expanded, including past cards opened through the past-days
  disclosure. Enter and Space toggled native disclosure with visible keyboard focus.
- Measured primary rows remained separate and in reading order: four game rows
  and three block rows, with larger times. Classification warnings followed the
  responsible team. Icons, cake update attributes, placeholders, eligibility
  warnings, rings and badges remained visible in their applicable states.
- No horizontal page scrolling, clipped primary rows or overlap with Details was
  found. Screenshots of mobile/desktop cards, long warnings, adjacent dates and
  enlarged text were inspected visually; full names and metadata wrapped naturally.
- Opening, closing and focusing cards issued no writes; hashes of every persisted
  application table remained unchanged throughout each read-only case.

The past-days wrapper previously applied its summary padding and margins to nested
game/block summaries, causing a badge/Details overlap in a past cake card. Its
selectors now target only the wrapper's direct summary; the complete matrix above
passed after that correction.

## Saved settings and regression checks

- Through the real admin UI at 360 CSS pixels, an unset responsible team was
  saved and appeared in the fourth summary row after navigation.
- An unconfigured cake block was saved with delivery time `14:45` and quantity
  `5`. Its time and quantity updated in their original rows immediately and
  remained correct after reload. Assignment rows were identical before and after
  both setting changes.
- `test/run_tests.sh` passed after the shared CSS change, the template change,
  the past-card selector correction and the README update: 454 checks across
  56 test modules per complete run, with exit status 0.
- `openspec validate stack-schedule-game-summary --strict` and
  `git diff --check` passed.

## Evidence from this run

The requested follow-up increases the shared time size to `1.5rem`, with
`1.25` line height and `2px 0 6px` margins. Together with the grid gap this
leaves 10px before the title. The complete 32-case browser matrix and real saved
settings checks passed again with these larger times and additional spacing.
`test/run_tests.sh` also passed again: 454 checks across 56 modules, exit status 0.

Temporary detailed reports and screenshots were retained under
`/tmp/nuliga-stack-layout.json`, `/tmp/nuliga-stack-integrated-report.json`,
`/tmp/nuliga-stack-evidence/` and `/tmp/nuliga-stack-integrated/`.
