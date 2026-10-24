# Tasks

## 1. Shared display label and retained identity

- [x] 1.1 Introduce the dependency-safe shared role-display helper and apply it to position labels, eligibility reason text and human-facing domain errors while retaining `Ordnungsdienst` as the assignment/comparison key. Extend existing synthetic domain/eligibility coverage to verify singular Ordner labels, unchanged role/slot identity, age-18 birthday boundaries and refusal for unknown birth dates or invalid game dates; run the affected tests and `test/run_tests.sh`.
- [x] 1.2 Capture the current display label in every new game-duty audit construction path while rendering historical audit labels literally. Extend audit/lifecycle coverage to verify old Ordnungsdienst entries and snapshots are unchanged, new claim/request/closure/takeover/lock evidence names Ordner, and assignment roles, occupants, locks and request generations retain their existing semantics; run the affected tests and `test/run_tests.sh`. Document the stored-key versus audit-label distinction in `AGENTS.md`.

## 2. Schedule, guidance and statistics

- [x] 2.1 Apply the current name to server-rendered task labels, accessible text, task-help labels, Kasse guidance, age notices, statistics and displayed vacancy/deficiency fragments. Update existing web/task-description tests to verify Ordner for guests and signed-in users, seven adult/five baseline positions, retained occupants and unchanged privacy boundaries; update current role/age documentation in `README.MD` and `AGENTS.md`, then run the affected tests and `test/run_tests.sh`.
- [x] 2.2 Keep display labels consistent in JavaScript refreshes, assignment/replacement controls and both confirmation paths while preserving structured role values and candidate keys. Verify through existing browser/route coverage that Ordner is visible, confirmation still precedes writes, cancellation/dismissal writes nothing, stale state is refused, and successful claims/takeovers retain age, one-duty and cutoff checks; run the affected tests and `test/run_tests.sh`.

## 3. Notifications and CLI

- [x] 3.1 Supply display labels to the authoritative message renderer for assigned-helper reminders, rescheduling, MV gaps/deficiencies and weekly team vacancy invitations. Extend existing captured-message tests for ordinary-game/Spielfest e-mail and SMS to verify Ordner wording with unchanged named fields, recipient selection, dispatch counts and triggers; run the affected tests and `test/run_tests.sh` without real provider calls.
- [x] 3.2 Accept the current CLI argument Ordner and the compatibility alias Ordnungsdienst through one normalization path, and use the current label in help, supported-role guidance and results. Extend existing CLI coverage to verify both inputs target the same saved role and share domain refusals; update `README.MD` CLI guidance to document the current spelling and alias, then run the affected tests and `test/run_tests.sh`.

## 4. Integration verification

- [x] 4.1 Review remaining Ordnungsdienst references and classify them as intended stored/transport identifiers, compatibility aliases, historical evidence or obsolete visible wording. Verify a synthetic existing database needs no migration, keeps assignment/audit state across read-only views, and has consistent Ordner output across schedule, confirmation, statistics, CLI and captured notifications. Run `test/run_tests.sh` and `openspec validate rename-ordnungsdienst-to-ordner --strict`; confirm both debug switches remain false and only intended implementation/documentation changes are present.
