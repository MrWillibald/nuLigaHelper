# Tasks

## 1. Release locks that the current game date no longer justifies

- [x] 1.1 Add `release_unbound_game_locks(session, game)` to `nuligahelper/db/locks.py`: for each refreshed assignment with `locked_at` set, release when the current date is open/unparseable, or when it is valid, `duty_is_locked(session, game)` is false and `locked_at` is earlier than the date's cutoff; clear `locked_at`, audit `unlock` with system actor, return the count. Verify `tests/db/test_db_structure.py` still accepts the layer imports
- [x] 1.2 Call it from `sync_games()` in `nuligahelper/db/sync.py` for every stored game of the synchronized season (scraped and no longer scraped) after the scraped fields are applied, keeping the existing `enforce_due_duty_locks` call before the update; verify with `./venv/bin/python -m pytest tests/db -v -k "sync or binding"`
- [x] 1.3 Add postponement tests in `tests/db/test_binding_assignments.py` (patching `nuligahelper.db.clock.duty_now`): unlock after a move to a later not-yet-due date with helper retained and a new replacement request plus takeover possible; closed request not reopened; late-claim lock released; re-lock at the new cutoff with a fresh `lock` audit; later-but-due or marker-saved new date keeps the lock; earlier date and same-date time change keep the lock; day-block locks untouched; exactly one `unlock` audit per released assignment and none on a repeated sync. Verify the tests pass and fail without task 1.2
- [x] 1.4 Add open-date tests: valid → `""` and → `"offen"` releases the lock with helper retained; a game already stored with an unparseable date and a saved lock is released by a sync without a date change, and a second sync adds no audit; open → valid not-yet-due date releases a remaining lock; open → valid already-due date locks again; an unlocked open-date Zeitnehmer takeover is still refused by the invalid-date age rules. Verify the tests pass and fail without tasks 1.1/1.2
- [x] 1.5 Add release-transition tests: a database state built directly as left by the previous release (game date already later and not yet due, assignments still carrying the original date's `locked_at`, request closed, `lock` audits present) is repaired by the first `sync_games()` run with the scraped date unchanged, including a stored game missing from the scrape; a lock on a game whose current cutoff is due stays; a game of another season is untouched; old audits stay unchanged. Verify the tests pass

## 2. Documentation

- [x] 2.1 Update the lock domain rule in `AGENTS.md` (replace "retain any saved lock after game date/category changes" with: saved locks are released when the current date is open/unparseable or a later not-yet-due date, evaluated on every sync) and verify the text matches the delta spec in `specs/task-self-service/spec.md`
- [x] 2.2 Add `unlock` to the audit action list in `docs/benutzer/BENUTZERHANDBUCH.md` and mention in the lock explanations of `BENUTZERHANDBUCH.md` / `MITGLIEDERHANDBUCH.md` that duties of a game postponed to a later date or without a fixed date are unlocked until the day-before cutoff of the new date; verify by reading the rendered sections
- [x] 2.3 Add a one-time transition note to `docs/betrieb/deployment.md` (next to the existing one-time transition section): the first daily run after this release unlocks duties of already postponed/open-date games, nothing is sent for the unlock itself, how to review the `unlock` entries on the audit page, and the rollback behaviour from design.md; verify the steps match the deploy and `resume-timers` commands documented there

## 3. Integration check

- [x] 3.1 Run `tests/run_tests.sh` and verify the whole suite is green; run `openspec validate unlock-duties-of-postponed-games --strict` and verify it passes

## Workflow follow-up

- Archive the change after review and sync the delta specs into `openspec/specs/`.
