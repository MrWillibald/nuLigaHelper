# Design

## Context

The existing role map has Zeitnehmer, Sekretär, two Verkauf places, one Ordnungsdienst and optional Unterstützung. An earlier reviewed revision renamed legacy Reinigung to Unterstützung. Current Unterstützung must now become Kasse; newly added Reinigung must not inherit it. Existing revisions and historical audits remain immutable.

## Goals / Non-Goals

**Goals:** Authoritative category-dependent position sets for offered duties, claims, rendering, reporting and progress; retained identity and readable history; release-only maintenance for removed duties.

**Non-Goals:** Configurable roles, task descriptions, changed age thresholds, changed day blocks, relaxed assignment authority or exceptions to one task per person per game.

## Decisions

### Offered positions are all required

Adult M/F games offer and require Zeitnehmer 0, Sekretär 0, Verkauf 0/1, Ordnungsdienst 0, Kasse 0 and Reinigung 0/1: eight positions. Youth mA–mE/wA–wE and SPF offer and require only the baseline first five. Unknown/GE categories retain that same five-position baseline and expose unresolved classification. Exactly one independent Ordnungsdienst remains; the separate Kasse role also carries the additional ordering/security work of the rejected second Ordner.

Expose one shared offered-position calculation and use it for requiredness, candidates, rendering, CLI validation, vacancies, MV follow-up and progress. Global role capacities validate physical positions, while category-specific offered positions decide new claims. All positions in the offered set are mandatory; there are no optional markers or empty removed fields.

### Preserve removed-duty occupants and permit releases

A changed category or migrated youth/unknown Unterstützung assignment may leave an occupied position outside the offered set. Preserve its assignment identity and audits. Show it separately as **Bestehende Einteilung**, with its duty and occupant visible under existing privacy rules. Authorized users may release it using the existing per-slot CAS endpoint; the control offers no replacement candidate or new claim. After release it disappears. The member/MV/admin scopes and past-date rules remain unchanged. A removed assignment still prevents a second duty for that person in the same game and remains in personal statistics, filters and ordinary reminders. It does not affect required progress, vacancies or MV reminders.

Validate the saved game category under the writer lock and on retries before a new claim, so stale category state cannot authorize a removed role. No tier or CLI/system caller can override the offered-position restriction. Historical age deficiencies and day-block behavior remain as implemented.

### Shared explicit classification

Reuse the classifier used by age eligibility rather than UI colors. It recognizes explicit adult M/F and youth mA–mE/wA–wE plus SPF. GE, missing, malformed or conflicting labels remain unknown; the baseline is retained without inferring adult duties. Unknown classes still refuse new timing claims under the independent age rules.

### Resolve every assignment by role and position

Verkauf and Reinigung have numbered display positions of one semantic role; Ordnungsdienst remains unnumbered and singleton. Render occupants, fetch candidates and notify actual assignments by role and position, with no Verkauf-only shortcut. Return authoritative saved required progress in candidate and mutation responses; the client never guesses progress from a role name.

### Guarded rename migration

Append revision `0006_game_duty_staffing` after the existing cake head. Before changing data, refuse Unterstützung/Kasse collisions, unexpected target/new-role rows and invalid old positions rather than merging or discarding them. Rename only current Unterstützung rows, preserving id, game, person and position. Keep historical audits and earlier Reinigung-to-Unterstützung revision unchanged. New Reinigung places start empty, and the old ROLE_CLEANING alias becomes the distinct new Reinigung role. Keep ROLE_SUPPORT only as an import alias for Kasse.

Update known schema fingerprints and revision-aware retained-data verification; fresh initialization stamps head, runtime never upgrades implicitly, and near-miss schemas fail closed. Per-category availability is deliberately not a schema constraint, because retained removed assignments must survive.

## Risks / Trade-offs

- Unknown classes must not gain adult duties: explicit classification and baseline fixtures.
- A repeated-role occupant could disappear or receive duplicate messages: actual-assignment traversal and sparse-position fixtures.
- Legacy data could collide: preflight refusal, backup-first migration and snapshot recovery.
- Removed duties could accept new claims: serialized saved-category validation, release-only candidate/control tests across tiers and CLI.
- Progress and reports could disagree: shared offered/required positions and authoritative API responses.

## Migration Plan

Stop web/daily database users and run the existing guarded backup-first migration. Verify assignment identities, new head, capacities and unchanged audits before restart. Rollback restores the printed validated snapshot and previous application with writers stopped; no downgrade rewrites history. Production migration is operator work outside local implementation.
