# Tasks

## 1. Align explicitly confirmed cutoff and maintenance policy

- [x] 1.1 Record the user's explicit fixed 09:00 Europe/Berlin day-before cutoff, independent of missed jobs/delivery, and refusal of destructive maintenance that would cancel binding commitments. Replace the superseded actual-batch/default-maintenance design throughout all deltas and rerun strict validation.
- [x] 1.2 Specify atomic account/game/day/block maintenance refusal and authorized obsolete-duty retirement before cutoff, without claiming an administrator override or allowing occupied cake reductions.

## 2. Durable assignment and reminder-batch state

- [x] 2.1 Add persistent replacement-request generation, saved commitment lock state and fixed day-before cutoff/date marker for game and block assignments; verify synthetic model tests preserve uniqueness, slot numbers and original occupant while seeking replacement.
- [x] 2.2 Add a reviewed Alembic revision on the single head and update exact fingerprints/postflight verification; verify synthetic upgrade/fresh-init/drift tests preserve all people, assignments, memberships, MV appointments and old audit snapshots without fabricated confirmation or requests.
- [x] 2.3 Document guarded stopped-writers migration and validated-snapshot rollback in `README.MD`, with no real-database operation; verify startup still rejects old/divergent heads and run `test/run_tests.sh`.

## 3. Serialized replacement requests and responsibility transfer

- [x] 3.1 Replace ordinary offered-duty release with occupied `Ersatz gesucht` requests and implement generation-aware atomic takeover for games and each block phase; verify success keeps continuous occupancy and failures preserve the outgoing helper/request in synthetic domain tests.
- [x] 3.2 Revalidate authorization, current account/membership, offeredness/capacity, per-container uniqueness and cutoff after obtaining write access and after retries; verify competing requests/takeovers and stale request generations after a prior transfer have one winner without overwrites or false success audits.
- [x] 3.3 Apply existing age eligibility to the resulting takeover group, excluding the outgoing seller; verify birthdays/unknown dates, a sole adult seller, concurrent Verkauf changes and current saved person/game data in the existing eligibility/concurrency suites.
- [x] 3.4 Record request and first-lock transitions plus linked before/after takeover evidence in one transaction with immutable private-data-safe snapshots; verify no mutation/audit on cancellation/refusal and historical readability after pre-cutoff obsolete-duty retirement and refusal of destructive maintenance.
- [x] 3.5 Apply the explicitly confirmed maintenance-refusal policy consistently to release helpers, CLI, deactivation/deletion, sync/removal and cake settings without compaction or changing required progress; verify destructive maintenance is refused atomically and obsolete retirement remains available only before cutoff through synthetic regression tests and update `README.MD`/`AGENTS.md`, then run `test/run_tests.sh`.

## 4. Fixed day-before cutoff and reminder integration

- [x] 4.1 Enforce 09:00 Europe/Berlin on the day before the scheduled date using a shared aware clock; persist idempotent dated cutoff/occupied lock/request-closure state across games, retained roles, preparation, cakes and cleanup. Verify immediately before/at/after the boundary and daylight-saving date arithmetic.
- [x] 4.2 Keep provider delivery outside the write transaction; due reminder processing persists transitions before dispatch, while early day-before calls do not advance 09:00. Verify missing contacts, missed jobs, provider failure, retries and restart never postpone or reopen locks, and weekly reminders do not lock.
- [x] 4.3 Recheck the fixed clock and saved locks inside every applicable write after serialization and retries; lock late empty claims immediately and close requests retaining their helpers. Verify takeover crossing 09:00, missing marker, stale page and game/block concurrency.
- [x] 4.4 Update README.MD/AGENTS.md for the explicitly confirmed clock cutoff and destructive-maintenance refusal. Verify no 24-hour timer, automatic return, administrator override or delivery-dependent lock is claimed; run test/run_tests.sh.

## 5. Pre-save commitment confirmation backed by the lifecycle

- [x] 5.1 Add German pre-save confirmation with explicit `Bestätigen` / `Abbrechen` to every existing game/preparation/cake/cleanup assignment-creation UI, including authorized assignments for another helper; verify selection, dialog display, Cancel, Escape and dismissal perform no mutation request or compensating release, using synthetic browser/JavaScript coverage.
- [x] 5.2 Preserve keyboard focus/return and add a usable non-JavaScript confirmation path where assignment creation is available; verify confirm submits once with CSRF and the original expected occupant, and cancellation never submits an assignment write.
- [x] 5.3 Cover a slot filled concurrently during confirmation: Cancel preserves the winner and Confirm receives current-state CAS refusal; verify no cancelled-user assignment or successful assignment audit appears for any container.
- [x] 5.4 Document the implemented confirmation behavior in `README.MD` and `AGENTS.md` without claiming the unimplemented lifecycle is active; verify the wording matches implemented paths and run `test/run_tests.sh` after this change group.

## 6. Web lifecycle, candidates and truthful presentation

- [x] 6.1 Add CSRF-protected request/takeover routes and candidate scopes using current tier rights and post-change eligibility; verify member, responsible-team MV, admin, guest/pending and post-lock refusals in synthetic Flask authorization tests.
- [x] 6.2 Show persistent German replacement/fixed state and keep locked occupied duties free of editing controls while genuinely empty offered duties stay claimable; verify game/block/cake cards, retained-duty policy, responsive/keyboard behavior and that guests receive no IDs, roster, contacts, full birth dates or exact ages.
- [x] 6.3 Replace the browser release-then-claim flow with one takeover; reuse pre-save confirmation for it, reconcile conflicts and lost responses, and refresh affected sibling/Verkauf candidates; verify failed/cancelled takeovers never display a committed vacancy or false success.
- [x] 6.4 Preserve occupancy/progress, personal duty ownership and separate age-deficiency/MV reporting across requests, locks and transfers; verify eight/five offered game positions, block capacities and retained-duty exclusions with synthetic statistics/notifier tests.
- [x] 6.5 Update action feedback and user documentation for the final completed behavior, keeping any notification text in `messages.py` if needed; verify no replacement broadcast or unapproved exception appears and run `test/run_tests.sh`.

## 7. Integration review

- [x] 7.1 Review the complete diff against every delta and the explicitly confirmed clock and maintenance decisions, confirm `DEBUG_FLAG` and `CHANGE_DAY` are `False`, and verify no real config/database, notification, commit, push or deployment was used.
- [x] 7.2 Run the complete synthetic offline `test/run_tests.sh` and strict OpenSpec validation after the final change; report actual passing results, remaining limitations and incomplete checklist items without representing partial work as a complete lifecycle.

## 8. End an own replacement search without giving up responsibility

- [x] 8.1 Add serialized game/block cancellation helpers and CSRF-protected routes using expected occupant plus request generation, current actor authority and cutoff checks; verify ownership, stale/reopened requests, racing takeover, cutoff contention and one actual-actor closure audit without changing the assignment.
- [x] 8.2 Show `Ersatzsuche beenden` for the current helper and existing authorized managers across every block/game flow, including non-JavaScript POST and refreshed cake cards; verify cancellation, reopening, stale responses, focus and unchanged physical occupancy in JavaScript/browser and Flask tests.
- [x] 8.3 Update the existing specs/design and README/AGENTS for the approved follow-up, run the complete synthetic offline suite and strict OpenSpec validation in the main checkout, and record actual results without touching real configuration, databases or providers.
