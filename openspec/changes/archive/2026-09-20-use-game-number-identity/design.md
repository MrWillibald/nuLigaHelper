## Context

See `proposal.md` for motivation and `specs/game-identity/spec.md` for required
behavior. Today the scraper emits a `source_key` from a nuLiga meeting link or a
fingerprint, `Game` is unique on `(season_year, source_key)`, and `game_nr` is an
integer display field. That allows changing source metadata to create a second row.

Real local data also establishes the special case: an `SPF Mini` Spielfest currently
produces seven rows, and small match numbers are reused at later Spielfeste. These are
not seven independent staffing events. The same database already contains accounts,
assignments and append-only audit records, so recreating it after the schema change is
not acceptable.

## Goals / Non-Goals

**Goals:**

- Keep ordinary game identity stable when nuLiga links or descriptive fields change.
- Convert each SPF date and full age group into one normal task-bearing `Game` object so
  existing assignment, notification and authorization paths need no parallel model.
- Preserve existing non-conflicting data through a reviewable one-time migration.
- Make pre-sync validation and migration failures atomic and actionable.

**Non-Goals:**

- Preserve individual SPF matchups, scores or match numbers in the application database.
- Transfer assignments automatically when a Spielfest moves to another date.
- Infer a survivor for existing duplicate ordinary game numbers.
- Introduce a general-purpose migration framework for unrelated future schema changes.

## Decisions

### D1 - Use one textual canonical game-number field

Change `Game.game_nr` from integer semantics to a bounded text value and enforce
uniqueness on `(season_year, game_nr)`. For an ordinary game, the canonical value is the
decimal representation of the scraped integer without decoration. The scraper no
longer emits `source_key`, and sync maps both incoming and stored ordinary games by the
canonical number.

Season remains part of database identity because nuLiga numbers may be reused in later
seasons. Local `Game.id` remains the relational identifier used by assignments, audits,
web endpoints and administrative mutations.

*Alternative considered:* retain `source_key` but prefer game number during matching.
This leaves two competing identities and permits another code path to recreate the
defect. A single canonical field makes the constraint match the sync rule.

*Alternative considered:* keep an integer and encode Spielfeste as negative numbers or
hashes. Such values are opaque and either require a fragile age-group registry or carry
collision risk.

### D2 - Collapse SPF rows in the scraper's normalization boundary

After the hall filter and basic field cleanup, partition rows whose full `ak` contains
the case-insensitive token `SPF`. Group them by parsed date and normalized full age-group
text. Each group becomes one synthetic game before duplicate identity validation and
before `sync_games()` sees the data. Non-SPF rows pass through one-for-one.

The synthetic canonical number uses an explicit namespaced form based on ISO date and
normalized full age group, for example `SPF:2026-11-28:spf mini`. Normalization trims and
collapses whitespace and uses Unicode-aware case folding. The original full age-group
label remains available for display.

The aggregate takes the shared date, day, hall and age group plus the earliest valid
match time. It carries a Spielfest presentation rather than an invented home/guest
matchup; plan cards, picker labels and notification templates render it as one
Spielfest. Score-dependent referee detection does not combine individual SPF scores.

All rows in a group must have the same hall after normalization. Missing or invalid
dates/times and conflicting shared fields produce a diagnostic containing the group,
and parsing returns no partially normalized result.

*Alternative considered:* store every match and add a separate staffing-event table.
That preserves detail the user explicitly does not need and would duplicate most game
assignment and notification relationships.

*Alternative considered:* group SPF rows during database sync. Keeping source-specific
aggregation in the scraper gives sync one uniform canonical identity contract and makes
the behavior independently testable from HTML fixtures.

### D3 - Treat a moved Spielfest as removal plus addition

Because the date is intentionally part of the pseudo number, a moved Spielfest has a
new identity. Complete-scrape reconciliation reports the old aggregate as removed and
the new aggregate as added. No heuristic copies its team or assignments.

This is an accepted trade-off: SPF shifts are rare, and guessing continuity would
reintroduce an identity independent of the pseudo game number.

### D4 - Keep local IDs in downstream events and mutations

Sync events continue to include `game_id` for exact lookup and the textual canonical
number for readable diagnostics. Web and CLI mutations continue to post internal game
IDs. Sorting treats ordinary decimal numbers numerically while placing Spielfest entries
by the existing date/time sort keys and using their canonical text only as a final
tiebreaker.

This avoids widening the identity change into assignment concurrency, audit or access
control behavior.

### D5 - Use an explicit, backup-first one-time SQLite migration

Add a narrowly scoped migration command/version check rather than silently relying on
`Base.metadata.create_all()`, which cannot replace the existing constraint or remove the
non-null `source_key` column. Normal application startup detects the old schema and
stops with instructions instead of mutating production data implicitly.

The migration runs with the daily job and web service stopped and uses SQLite's backup
API to create a dated recovery copy. It performs a preflight before schema mutation:

1. Identify ordinary rows sharing `(season_year, game_nr)` and require manual survivor
   selection before retrying.
2. Partition SPF rows by season, date and normalized full age group.
3. Reject conflicting responsible teams, role/slot occupants, or person assignments.
4. Select the row carrying compatible task data as survivor, otherwise choose a
   deterministic row, and preserve that `Game.id`.
5. Repoint readable audit references to the survivor while retaining audit snapshots.
6. Rebuild `games` with textual `game_nr`, no `source_key`, and the new uniqueness
   constraint; preserve IDs and all unaffected foreign-key references.
7. Run `foreign_key_check`, row-count and relationship-count verification before commit.

Migration records a schema version only after all validation succeeds. The backup path
and verification results are printed for the operator.

*Alternative considered:* delete and recreate the database. Existing production
accounts, assignments and audit evidence make that incompatible with current data
preservation requirements.

*Alternative considered:* keep the obsolete column and constraint. Its non-null
requirement complicates new inserts, leaves misleading schema semantics and does not
allow the model constraint to express the true identity rule.

## Risks / Trade-offs

- **Age-group text is corrected without changing date** -> the pseudo number changes;
  treat this like a new identity rather than guessing, consistent with the chosen rule.
- **An SPF event spans more than one hall** -> reject the group and expose the source
  rows; do not silently split because the agreed grouping key is date plus age group.
- **Textual game numbers affect numeric filters and ordering** -> normalize ordinary
  input to decimal text and use explicit numeric-aware display sorting.
- **A migration is interrupted** -> use a transaction, leave the schema-version marker
  unchanged, verify the source/backup before retrying, and never start application writes
  against a partially migrated schema.
- **Old duplicate ordinary rows carry useful data** -> abort with IDs and context for an
  operator decision; do not choose newest, oldest or most populated automatically.
- **Collapsed SPF data loses match detail** -> accepted explicitly; nuLiga remains the
  authoritative detailed match plan while nuLigaHelper stores the staffing event.

## Migration Plan

1. Stop the web service and daily timer/job and confirm no SQLite writers remain.
2. Run migration preflight and resolve any reported ordinary duplicates or incompatible
   SPF task data without altering the original database automatically.
3. Create and verify the dated SQLite backup.
4. Run the transactional schema/data migration and its integrity checks.
5. Start the new code, perform one scrape, and verify ordinary counts plus one plan entry
   per SPF date/age group.
6. Verify representative assignments, responsible teams, audit entries, filters and
   notifications before restoring normal scheduling.

Rollback stops all writers, retains the failed database for diagnosis, restores the
verified pre-migration backup, and redeploys the prior application version. A successful
migration is not reversed in place.
