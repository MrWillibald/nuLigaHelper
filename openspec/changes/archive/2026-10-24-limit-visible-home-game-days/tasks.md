# Tasks

## 1. Persistent storage and guarded schema upgrade

- [x] 1.1 Add the constrained singleton visibility-setting model, explicit unrestricted initialization and one successor revision to `0008_single_wischer`; verify fresh initialization and upgrade from a synthetic current database yield the same new single-head schema with exactly one unrestricted setting.
- [x] 1.2 Update `schema_migrations.py` source recognition, revision-specific expected metadata, required head and drift checks without changing the fixed historical Wischer revision; extend offline fixtures/tests to accept exact supported sources and reject near-miss schemas, missing singleton state and implicit startup upgrades.
- [x] 1.3 Add change-specific migration postflight and retention tests covering game/block assignment identities, locks, open replacement tokens, persons, memberships and every historical audit row; verify older supported upgrade chains retain only their already-reviewed changes, repeated migration preserves a configured boundary, and backup/restore recovers the pre-upgrade snapshot.
- [x] 1.4 Document the new required schema revision and stopped-writers `migrate-schema --confirm-stopped` deployment/recovery procedure in `README.MD`; verify commands match the CLI and no downgrade or runtime migration is suggested.

## 2. Visibility policy and protected setting writes

- [x] 2.1 Add a read-only setting lookup and pure shared date-visibility helper using existing date parsing; verify synthetic cases before/on/after the boundary, no boundary, month/year transitions, leap dates, invalid game dates and an effective day/season advancing beyond a retained boundary.
- [x] 2.2 Add the serialized setting-update helper with fresh active-admin checks, expected-boundary comparison, required expected-state input, future-date validation, clear support and identical-value no-op; verify permitted saves persist across separate sessions, invalid/stale saves do not overwrite state and an expired saved date can still be cleared or resubmitted unchanged.
- [x] 2.3 Add the explicit CSRF-protected POST endpoint and non-JavaScript form/validation flow using existing German action feedback and post/redirect/get; add tier, pending-account, missing-field, missing/invalid-CSRF and rights-revoked-after-load tests proving refused requests leave settings and domain records unchanged.
- [x] 2.4 Add a concurrent two-session setting-write regression verifying an outdated expected value is refused after serialization/retry rather than replacing a newer boundary, and contention exits within existing SQLite policy without a successful-save message.
- [x] 2.5 Add the admin-only setting form to the schedule with a labeled date input, explicit save/clear behavior and display-only help; verify initial state, correction feedback, keyboard access and saves without JavaScript, and document the fixed inclusive date, manual lifecycle and empty-value behavior in `README.MD`.

## 3. Schedule date scope and administrator full-date view

- [x] 3.1 Resolve the shared boundary and request-local `show_hidden=1` only from current admin rights; add German full-date toggle links and preserve authorized mode across schedule/statistics navigation, filter submission and filter clearing. Verify copied override links do not widen guest/member/MV results and revoked rights take effect on the next request.
- [x] 3.2 Apply the visibility predicate before schedule date options, ordinary filters and day/month grouping, covering games, preparation, cake delivery and cleanup; extend offline schedule tests for the inclusive edge, all block phases, unknown dates, combined date/team/person filters, a manually supplied hidden date and chronological options without leaked hidden cards.
- [x] 3.3 Retain full-date game input for block-time calculation and the original full-season count for empty-season detection; verify ordinary filtering keeps actual first/last game block times, clearing cannot remove the shared boundary, and a retained past boundary interacts correctly with past-section grouping/counts and admin full-date mode.
- [x] 3.4 Add active-scope notices and distinct no-visible-content feedback, keeping admin controls reachable when everything is hidden; verify native/keyboard operation, German labels, initial HTML without JavaScript and wrapping at 320/360 CSS pixels and 200-percent enlargement without private roster/contact/birth-date data.
- [x] 3.5 Document schedule override usage, notices and the distinction from ordinary filters in `README.MD`; verify examples match the rendered labels and preserve the explanation that hidden duties remain binding and notified.

## 4. Outstanding statistics with complete-season totals

- [x] 4.1 Apply the shared effective date scope only when constructing outstanding game/block rows, retaining existing current/future and invalid-date handling; add synthetic tests for boundary-day vacancies, age deficiencies, preparation/cleanup gaps and unconfigured cake setup, and verify admin full-date mode restores hidden current/future issues without adding past rows.
- [x] 4.2 Update the outstanding header count, scope notice and scoped empty-state wording while leaving team/person sections season-wide; verify each visible affected container counts once, hidden issues count zero, and all season totals, team shares, personal role counts and retained-duty counts are identical before and after changing the boundary.
- [x] 4.3 Preserve existing responsive statistics disclosures and session/privacy requirements; verify keyboard/no-JavaScript behavior, zero-count headings, narrow-width scope controls, guest rejection and absence of birth dates or exact personal ages, and document the differing section scopes in `README.MD`.

## 5. Integration and release readiness

- [x] 5.1 Verify hidden-date candidate reads and claims/takeovers through JSON and non-JavaScript confirmation still follow existing rights, age, expected-state and lock rules; verify preview/filter GETs and setting saves create no duty/lock/request/audit transitions or notifications, while ordinary due-lock and notifier flows still process hidden commitments with synthetic mocked providers.
- [x] 5.2 Run `test/run_tests.sh` offline after the completed changes and confirm the entire suite remains green, including migration, authentication/refusal, stale-write, notification and responsive-presentation coverage; resolve regressions without production configuration or provider traffic.
- [x] 5.3 Run `openspec validate limit-visible-home-game-days --strict` and review the final implementation against all four delta specs and README instructions; verify only the intended feature changes remain and both debug switches are still false before delivery.
