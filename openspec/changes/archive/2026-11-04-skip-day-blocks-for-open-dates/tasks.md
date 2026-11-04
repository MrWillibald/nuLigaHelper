# Tasks

## 1. Synchronization

- [x] 1.1 Add `block_date_valid()` to `nuligahelper/db/day_blocks.py` and export it. Verify with tests for a valid date, `""`, `"offen"` and `31.02.2026`, and that `tests/db/test_db_structure.py` accepts the import
- [x] 1.2 Compute calendar dates once in `sync_games()` in `nuligahelper/db/sync.py`, remove the preflight refusal and pass `retained_dates` to `reconcile_day_blocks()`, which keeps its refusal for any other occupied block. Verify in `tests/db/test_day_blocks.py`: no blocks for open dates, empty blocks removed, occupied open-date block retained with sync completing and other games updated, emptied block removed by the next sync, repeated sync idempotent, direct `reconcile_day_blocks()` refusal unchanged, open date receiving a valid date gets its blocks
- [x] 1.3 Check existing open-date sync tests in `tests/db/test_binding_assignments.py` still pass with `./venv/bin/python -m pytest tests/db -v -k "sync or binding"`
- [x] 1.4 Update the day-block and destructive-maintenance rules in `AGENTS.md` and add a transition note to `docs/betrieb/deployment.md`. Verify the text matches the delta specs

## 2. Retained blocks

- [x] 2.1 Refuse `claim_block_slot()` and `configure_cake_block()` on open-date blocks after serialization. Verify in `tests/db/test_day_blocks.py` and `tests/db/test_cake_blocks.py` that nothing is written or audited, and that a replacement request and takeover still succeed
- [x] 2.2 Show retained blocks in `build_schedule()` even without a game and offer no free positions or vacancies in `block_view()` in `nuligahelper/web/views.py`. Verify in `tests/web/test_day_blocks_web.py`, including a retained block whose date no game has
- [x] 2.3 Exclude retained blocks from missing and setup reporting in `build_statistics()` in `nuligahelper/web/statistics.py` while keeping personal counts. Verify in `tests/web/test_day_blocks_web.py`
- [x] 2.4 Describe the behaviour for games without a fixed date in `docs/benutzer/BENUTZERHANDBUCH.md` and verify it against the delta spec

## 3. Cancelled home-game dates

- [x] 3.1 Add `_cancel_date_block_duties()` and `SyncEvents.cancelled_block_duties` in `nuligahelper/db/sync.py` as in design.md, with `request_close` and `date_remove` audits. Verify in `tests/db/test_day_blocks.py` and `tests/db/test_cake_blocks.py`: last game moved and vanished, locked duty and open request removed, one event and one audit per assignment, game update saved in the same sync, date returning gets empty blocks, and rewrite the sync refusal in `test_day_blocks.py` that expects `verbindlichen`
- [x] 3.2 Retain occupied blocks of dates before today without failing the sync. Verify with a test for today (cancelled), yesterday (retained, no event) and unchanged personal statistics
- [x] 3.3 Refuse an empty scrape for a season with stored games before any write, and allow it for an empty season. Verify with tests that nothing changes and no event is produced
- [x] 3.4 Add `block.date_cancelled` with mail and SMS variants to `nuligahelper/jobs/messages.py`. Verify with `tests/jobs/test_messages.py`
- [x] 3.5 Add `Notifier.notify_cancelled_block_duties()` and call it from `_send_notifications()` in `nuligahelper/jobs/daily.py`. Verify in `tests/jobs/test_notifier.py`: mail 1, phone fallback 1, no contact 0, missing person skipped, and a dispatch failure leaves the removal saved
- [x] 3.6 Describe cancelled home-game dates in `docs/benutzer/BENUTZERHANDBUCH.md` and `MITGLIEDERHANDBUCH.md`, including the statements that occupied dates cannot be removed, and add `date_remove` to the audit action list. Verify against the delta specs

## 4. Integration check

- [x] 4.1 Run `tests/run_tests.sh` and verify the whole suite is green; run `openspec validate skip-day-blocks-for-open-dates --strict` and verify it passes

## Workflow follow-up

- Archive the change after review and sync the delta specs into `openspec/specs/`.
