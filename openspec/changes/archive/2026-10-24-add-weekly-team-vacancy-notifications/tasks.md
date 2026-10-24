# Tasks

## 1. Centralized invitation copy

- [x] 1.1 Add the ordinary-game and Spielfest weekly vacancy templates, explicit e-mail/SMS variants, documented named fields and singular/plural vacancy fragments in `messages.py`, using the copy in `design.md`; verify both event variants render with the correct subject, personal greeting, event details, counts, invitation, thanks and sign-off in `test/test_messages.py`.
- [x] 1.2 Extend message-catalog tests for one versus two vacant Verkauf positions, unknown or missing fields, literal braces in names and Spielfest text without opponents; verify with `./venv/bin/python test/test_messages.py` and ensure vacancy prose is defined only in the catalog.

## 2. Vacancy sender and recipient selection

- [x] 2.1 Add `Notifier.notify_team_vacancies_early(date)` using season/date-scoped games, shared required-vacancy counts and the saved responsible team's active memberships; add synthetic notifier tests covering adult/youth/SPF counts, no responsible team, no appointed MV, qualifying MV, non-members and inactive/pending/rejected accounts, and verify with `./venv/bin/python test/test_notifier.py`.
- [x] 2.2 Exclude every same-game assignment by person identity, including retained removed-duty and Ersatz-gesucht assignments, while retaining members assigned only in other games/day blocks; add tests for each exclusion, duplicate display names, multiple memberships and no additional birth-date filter, and verify each qualifying identity receives one message per game per invocation.
- [x] 2.3 Dispatch through the existing validated renderer and ordinary transport helpers; add notifier tests for e-mail preference, SMS fallback, contactless warning/zero count, both event variants and provider-error propagation, and verify the sender creates no assignment, replacement, cutoff or assignment-audit writes.
- [x] 2.4 Add notifier regressions showing full physical staffing sends no invitation despite age/classification deficiencies or replacement searches, block gaps do not trigger invitations, and assigned-helper/day-before MV reminder conditions remain intact; verify `./venv/bin/python test/test_notifier.py` passes and use synthetic configuration and recorded transports throughout.

## 3. Daily scheduling and operator documentation

- [x] 3.1 Call the new sender from the existing seven-day branch and include its dispatch count; update the daily-job notifier double and add integration coverage proving the effective-today plus seven-day date is used, six/eight-day games are excluded and existing weekly senders remain called, verified with `./venv/bin/python test/test_main.py`.
- [x] 3.2 Document the new invitation in `README.MD`, including physical-vacancy triggering, active responsible-team recipients, same-game exclusion, other-container inclusion, e-mail/SMS preference, per-game delivery and the existing missed-run/rerun behavior; verify documentation matches both delta specs and does not claim delivery deduplication or introduce configuration requirements.

## 4. Integrated verification

- [x] 4.1 Run `test/run_tests.sh` after implementation and resolve failures with all database/configuration/provider data remaining synthetic and offline; verify the full suite passes and committed debug switches remain false.
- [x] 4.2 Run `openspec validate add-weekly-team-vacancy-notifications --strict` and review the implementation diff against the recipient and message scenarios; verify no schema migration, assignment-policy change, team-independent recruitment or unintended notification trigger was introduced.
