# Tasks

## 1. Canonical role, capacity and guarded upgrade

- [x] 1.1 Rename the current cleaning role value to Wischer and set its offered capacity to one; add offline domain tests proving seven adult required places, five youth/SPF/unresolved places, no new age minimum, and rejection of old-name or second-position writes through serialized helpers.
- [x] 1.2 Add one reviewed Alembic successor and revision-aware preflight/postflight: rename current Reinigung slot 0 and delete exactly current Reinigung slot 1 with one atomic system migration-removal audit per deletion. Verify synthetic current-head upgrades cover empty/occupied first slots, past/future and non-adult games, locks and open requests, preserve every unrelated row and old audit, and do not promote helpers.
- [x] 1.3 Update retained-data fingerprints and supported-source handling; verify older original Reinigung still becomes Kasse, unexpected target roles/slots/collisions refuse before domain writes, and fresh initialization, already-current reruns, startup refusal and snapshot recovery work without duplicate audit evidence. Ensure the chain has one head and the audit table is not exempted from retained-data checks.
- [x] 1.4 Document the requested one-time deletion exception, all-date scope, effect on past assignment counts, backup/recovery and seven adult places in `README.MD` and `AGENTS.md`; verify runtime binding and cutoff guidance remains strict and historical role descriptions retain their meaning.

## 2. Assignment presentation and protection

- [x] 2.1 Update web labels, confirmation flows, JSON validation and JavaScript role lists for the single Wischer; remove the second display entry. Verify synthetic guest/member/MV/admin schedule and candidate responses expose no second Wischer and retain German confirmation, cancellation and privacy behavior.
- [x] 2.2 Add post-migration tests for stale old-name and second-position claim/request/takeover writes, including previously saved replacement tokens; verify no deleted duty is recreated and no successful mutation audit or notification occurs. Verify the remaining first duty still obeys serialization, one-duty-per-game and before/at/after-cutoff protection without a runtime bypass.
- [x] 2.3 Update CLI choices and listings for one Wischer; verify synthetic CLI tests show no second assignment, reject unsupported mutations and retain ordinary obsolete-duty retirement restrictions. Update CLI examples in `README.MD` and check them against the parser.
- [x] 2.4 Change the central German description to singular Wischer wording with its existing work and separate Aufräumen meaning; verify task-description coverage, public help controls and keyboard access, and update current task-help documentation.

## 3. Staffing statistics and notifications

- [x] 3.1 Update progress, statistics and MV follow-up for seven adult positions; verify 5/7 identifies Kasse and one Wischer, 7/7 is complete, and youth retains five positions with independent age deficiencies. Verify deleted second assignments disappear from personal counts, including past counts, and update progress examples and staffing tables in `README.MD`.
- [x] 3.2 Verify recorded weekly/day-before/shift messages use Wischer, notify each remaining saved helper once and never send reminders for deleted second duties; preserve dispatch-count behavior and central validated templates. Verify migration sends no provider notifications and existing audits plus appended removal evidence stay readable in admin activity, with tests for any new action rendering.

## 4. Integration verification

- [x] 4.1 Run `test/run_tests.sh` on synthetic offline data after implementation and resolve failures; verify auth, refusal, concurrency, migration and binding-duty coverage remains green without production config or outbound provider use.
- [x] 4.2 Review application/catalog/docs for obsolete current labels, second-slot controls and eight-position assumptions while retaining historical fixtures and migrations; run `openspec validate rename-cleaning-to-single-wiper --strict` and `git diff --check`, and verify final behavior matches the delta scenarios.
