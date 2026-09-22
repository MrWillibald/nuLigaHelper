## 1. Schema and migration

- [x] 1.1 Add day-block and block-assignment ORM models, phase/slot definitions, relationships, and uniqueness constraints; verify focused database tests can create both phases and reject duplicate occupants and duplicate person assignments within one block.
- [x] 1.2 Extend assignment audit storage to identify either a game or day block with a durable target snapshot and a block reference that becomes null on block deletion; verify model tests preserve readable game and block entries after related records are renamed or deleted and never cascade-delete audit history.
- [x] 1.3 Add an Alembic revision that creates and seeds blocks for existing dated games, extends audit storage, preflights role collisions, and renames current `Reinigung` assignments to `Unterstützung` without rewriting audit history; verify migration tests cover empty, populated, collision, and rollback-snapshot workflows.
- [x] 1.4 Update schema fingerprints and synthetic schema fixtures for the new head; verify current, legacy, and near-miss schema tests continue to fail or pass as intended.

## 2. Day-block domain behavior

- [x] 2.1 Reconcile exactly one preparation and cleanup block for every distinct season/date after successful validated game synchronization, retaining same-date assignments when boundaries change and removing both blocks when a date disappears; verify sync tests cover new dates, repeated sync, earlier/later inserted games, disappearing dates, and suppression of deletion after failed or invalid sync.
- [x] 2.2 Implement robust block-time calculation at first start minus 90 minutes and last start plus 60 minutes, including midnight crossings and invalid times; verify focused unit tests cover normal, boundary, and no-time cases.
- [x] 2.3 Implement bounded compare-and-swap claim/release helpers for block slots with active-person checks, one-task-per-block enforcement, atomic audits, and a system removal path for deleted blocks; verify database and concurrent-claim tests exercise success, stale expectations, duplicate tasks, SQLite contention, one removal audit per occupied deleted slot, and one structured log entry per deleted block including empty blocks.
- [x] 2.4 Split assignable game roles from required game roles, rename the UI/domain role to `Unterstützung`, and exclude only that role from missing-slot evaluation; verify database tests show occupied optional duties still work while empty optional slots do not make a game incomplete.

## 3. Authorization and web API

- [x] 3.1 Add CSRF-protected default-deny block claim/release endpoints with member/MV self-only, admin-anyone, active-account, and past-date rules; verify refusal and web tests cover every tier, stale conflicts, inactive people, and past blocks.
- [x] 3.2 Return conflict and success payloads compatible with live schedule correction while keeping block and game target identifiers distinct; verify API tests assert current occupants and ensure failed writes create no audit entry.

## 4. Schedule interface and filtering

- [x] 4.1 Extend schedule construction to create complete date groups with calculated blocks and apply game-, block-, and date-level filter semantics from the specification; verify schedule-filter tests cover block-name matches, combined team filters, middle-game filtering, past dates, and empty results.
- [x] 4.2 Render preparation and cleanup cards around each visible date with three slots, calculated times, alphabetical admin choices, and self-service controls; verify web tests assert ordering, labels, absence of responsible-team controls, and responsive markup.
- [x] 4.3 Extend client-side assignment synchronization to treat each block as its own one-task container without coupling it to adjacent games; verify JavaScript tests cover claims, releases, duplicate suppression within a block, and availability across blocks and games.
- [x] 4.4 Preserve guest privacy for day blocks by rendering assigned names without IDs, roster data, contacts, or controls; verify guest-response and access-control tests scan block markup for forbidden data.
- [x] 4.5 Add schedule styling for the two boundary cards and warning/no-time states consistent with the existing visual language; verify representative desktop and narrow-screen renders remain readable.

## 5. Notifications, statistics, and audit UI

- [x] 5.1 Move the one-week special preparation dispatch from first-game `Verkauf` assignments to occupied preparation slots and send ordinary weekly reminders to all game-level sale helpers; verify notifier tests assert recipients, counts, subjects, partner/context handling, and calculated preparation time.
- [x] 5.2 Add weekly cleanup and day-before block reminders using existing email-first, phone-fallback, skip-count, and debug suppression behavior; verify notifier and daily-job tests cover occupied/empty slots, missing contacts, invalid times, and no MV recipient.
- [x] 5.3 Include block duties in person totals and upcoming block gaps in open-duty statistics while excluding empty `Unterstützung`; verify statistics tests cover all three cases and confirm blocks do not affect responsible-team coverage counts.
- [x] 5.4 Extend the admin audit view and filters to present block targets and snapshots alongside games without making history editable; verify audit tests cover ordering, target filtering, deleted targets, historical `Reinigung` text, and mobile presentation.

## 6. Documentation and integration verification

- [x] 6.1 Update notification configuration examples and production text handling for preparation and cleanup messages without changing existing placeholder contracts unintentionally; verify provider/configuration tests and template-format tests pass.
- [x] 6.2 Update `README.MD` and test documentation with the two day blocks, offsets, authorization, reminder behavior, and optional `Unterstützung` semantics; verify documented commands and role names match the implemented interface.
- [x] 6.3 Run `openspec validate add-game-day-task-blocks --strict` and `test/run_tests.sh`, resolving every validation or regression failure while keeping both debug date switches false.

## 7. Follow-up refinements

- [x] 7.1 Treat numbered preparation and cleanup slots as display positions of the single semantic roles `Vorbereitung` and `Aufräumen`, matching `Verkauf`; verify notifications, statistics, gaps, errors, and audits do not create numbered roles.
- [x] 7.2 Match block card surfaces to ordinary game cards without a colored background fade or side bar, and distinguish only the selected preparation `↗` (`#a0e656`) and cleanup `↘` (`#ffb752`) icon colors; verify schedule markup and responsive CSS tests.
