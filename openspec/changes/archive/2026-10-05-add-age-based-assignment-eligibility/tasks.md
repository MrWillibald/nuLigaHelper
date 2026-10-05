# Tasks

## 1. Birth-Date Storage and Migration

- [x] 1.1 Add date-only person birth-date storage and a reviewed Alembic revision; verify fresh initialization and guarded migration reach one current head with no model drift.
- [x] 1.2 Keep existing birth dates unknown during migration and add postflight checks; verify synthetic active, inactive, pending, contactless, admin, and MV people retain identities, memberships, contacts, assignments, and audit snapshots.
- [x] 1.3 Implement shared date validation using the effective current date; verify valid dates, impossible dates, future dates, leap dates, and debug-date behavior.
- [x] 1.4 Document the nullable legacy state, roster backfill, guarded migration, postflight, and rollback in operator guidance; verify documented commands match CLI help and existing stopped-service recovery procedures.

## 2. Person Collection and Maintenance

- [x] 2.1 Require birth dates before registration-code request and preserve them through the signed two-step flow; verify server-only registration, resubmission, missing/invalid dates, and contact anti-enumeration behavior.
- [x] 2.2 Require birth dates in admin/MV creation forms; verify contactless creation, managed-team creation scope, and complete validation rollback.
- [x] 2.3 Require birth dates in CLI creation and provide identity-safe CLI correction; verify valid and invalid date input, existing person IDs, duplicate-name handling, and refusal of partial creation.
- [x] 2.4 Add self/admin birth-date completion and correction without broadening MV profile permissions; verify legacy unrelated edits remain usable and a known date cannot be cleared.
- [x] 2.5 Expose existing birth dates only to self/admin person-maintenance contexts and add admin visibility of incomplete dates; verify member, MV, guest, candidate, schedule, statistics, notification, log, and audit paths exclude unauthorized dates and exact ages.
- [x] 2.6 Update registration, person-maintenance, CLI, and data-inventory documentation; verify the documented required fields, correction authority, and privacy visibility match the delivered forms and commands.

## 3. Eligibility Domain Rules

- [x] 3.1 Implement the explicit adult/youth/unknown classifier shared with the game-duty proposal; verify adult M/F, youth classes, SPF, GE, malformed classes, and league-prefix variations.
- [x] 3.2 Calculate completed age on the scheduled game date and implement Zeitnehmer 14 youth/18 adult, Sekretär 14 youth/16 adult, and collective Verkauf coverage by at least one person 18 or older; verify birthday-eve/birthday cases, sale qualification from the 18th birthday, leap-day convention, future games, and unknown dates.
- [x] 3.3 Produce separate structured physical-vacancy and eligibility-deficiency results; verify full invalid rosters, incomplete Verkauf groups, missing birth dates, and unknown timing-role classification remain distinguishable.
- [x] 3.4 Document inclusive birthday boundaries, game-date calculation, the leap-day convention, collective adult coverage, and unsupported-class handling; verify examples agree with the age-eligibility acceptance scenarios.

## 4. Assignment Enforcement and Candidates

- [x] 4.1 Enforce individual eligibility and post-claim Verkauf coverage in shared database mutations and all web/CLI entry points; verify members, MVs, admins, system callers, and past-game admin corrections obey the same age rules without changing authority or one-task limits.
- [x] 4.2 Revalidate authoritative group state under write serialization and contention retries; verify concurrent younger-person claims to different sale slots cannot create a full group without an adult and stale same-slot claims retain existing conflict behavior.
- [x] 4.3 Filter per-game candidates using stored eligibility while retaining current occupants, membership ordering, and advisory hints; verify an unknown/younger first seller can be assigned and a final-slot claim requires a qualifying group.
- [x] 4.4 Refresh affected sale-slot eligibility after claims/releases and retain withdrawals of the sole adult seller; verify browser state reflects saved group state and a rejected claim leaves occupants unchanged.
- [x] 4.5 Document claim refusals, candidate refresh, incomplete-sale staffing, and preserved release rights; verify documented behavior through the delivered API and browser regression scenarios.

## 5. Revalidation, Reporting, and Follow-Up

- [x] 5.1 Reevaluate current/future staffing after birth-date corrections and game date/class changes without automatically deleting assignments; verify birthday-crossing moves, youth-to-adult changes, legacy unknown dates, and unchanged historical audits.
- [x] 5.2 Add separate schedule eligibility indicators and statistics gaps while keeping physical occupancy counts truthful; verify full but invalid games are visibly outstanding and raw birth dates never enter schedule output.
- [x] 5.3 Include eligibility deficiencies in existing responsible-team MV follow-up and coordinate German feedback with centralized message templates when available; verify a full invalid roster still triggers the existing MV path, no unassigned adult satisfies coverage, and games without an active responsible MV retain existing routing.
- [x] 5.4 Document how staff resolve eligibility deficiencies and distinguish them from vacancies; verify the guidance matches schedule indications, statistics, and unchanged notification-recipient selection.

## 6. Integration Verification

- [x] 6.1 Run integrated synthetic migration, person-management, authentication, authorization, eligibility, assignment-concurrency, schedule, statistics, and notifier coverage; verify all components agree on saved age and occupancy state without production configuration reads.
- [x] 6.2 Run `test/run_tests.sh` and strict OpenSpec validation; verify the full offline suite passes with both debug switches false and every required planning artifact remains complete.
