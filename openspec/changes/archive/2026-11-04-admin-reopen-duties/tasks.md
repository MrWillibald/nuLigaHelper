# Tasks

## 1. Domain operation

- [x] 1.1 Add `db.reopen_slot()` in `nuligahelper/db/duties.py` with shared admin authorization in `db/slots.py`, as in design.md; export it from `db/__init__.py`. Verify with new tests in `tests/db/test_binding_assignments.py`: before cutoff, after cutoff, saved lock, open request closed with stale takeover refused, past date refused, open date permitted, non-admin and `actor=None` refused, conflict on a changed occupant, sole adult seller, other locks untouched
- [x] 1.2 Add `db.reopen_block_slot()` in `nuligahelper/db/block_duties.py` with the same rules. Verify with tests in `tests/db/test_day_blocks.py` and `tests/db/test_cake_blocks.py`, including a late claim being immediately locked and a cake reduction succeeding after reopening
- [x] 1.3 Write the `reopen` audit (plus `request_close` first when a request was open) for both writers. Verify exactly one `reopen` entry with actor and affected person, and none on refusal, and that `tests/db/test_db_structure.py` still accepts the imports
- [x] 1.4 Verify ordinary `release_slot()` and `release_block_slot()` behaviour is unchanged by running `./venv/bin/python -m pytest tests/db -v -k "binding or block"`
- [x] 1.5 Update the binding, lock and destructive-maintenance rules in `AGENTS.md` and verify the text matches the delta specs

## 2. Notification

- [x] 2.1 Add `duty.game.reopened`, `duty.spielfest.reopened` and `duty.block.reopened` to `nuligahelper/jobs/messages.py` with mail and SMS variants. Verify with `tests/jobs/test_messages.py`
- [x] 2.2 Add `Notifier.notify_duty_reopened()` using `_dispatch()`. Verify return counts in `tests/jobs/test_notifier.py`: mail 1, phone fallback 1, no contact 0

## 3. Web and user interface

- [x] 3.1 Add `POST /api/assignment/reopen` and `/api/block-assignment/reopen`: admin only, commit, then notify in a guarded block and return `notified`. Verify in `tests/web/test_refusals.py` and `tests/web/test_sticky_web.py`: CSRF, every lower tier refused, conflict response, failed delivery keeps the slot open and returns `notified: false` with no contact data
- [x] 3.2 Add `can_reopen` to `duty_state()` and block slot views in `nuligahelper/web/views.py`, and render `Auf offen setzen` on occupied slots including locked ones. Verify the control appears only for administrators and not on past duties
- [x] 3.3 Add the confirmation to `web/static/app.js` and the failure warning. Verify in `tests/web/test_javascript.py` that cancel and Escape send no request and that focus returns
- [x] 3.4 Add `action=reopen` with mandatory `decision=confirm` to `/dienst` and `assignment_form.html`. Verify in `tests/web/test_action_feedback.py` that a POST without confirmation changes nothing
- [x] 3.5 Adjust the binding confirmation wording in `assignment_form.html` and `schedule.html` to say helpers and MVs cannot change locked duties. Verify against the `action-feedback` delta
- [x] 3.6 Update `docs/benutzer/BENUTZERHANDBUCH.md` (sections on short-notice absence, admin rules, audit actions, troubleshooting) and `MITGLIEDERHANDBUCH.md`, and `README.md` where it describes locks. Verify no statement "kein Admin kann ändern" remains unqualified with `grep -n "Admin" docs/benutzer/*.md`

## 4. Integration check

- [x] 4.1 Run `tests/run_tests.sh` and verify the whole suite is green; run `openspec validate admin-reopen-duties --strict` and verify it passes

## Workflow follow-up

- Archive the change after review and sync the delta specs into `openspec/specs/`.
