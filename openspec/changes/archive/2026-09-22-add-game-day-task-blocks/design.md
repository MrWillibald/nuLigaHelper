## Context

Assignments currently have a mandatory foreign key to a scraped `Game`, and their two
uniqueness constraints implement one person per game and one occupant per role slot.
Schedule rendering, authorization, compare-and-swap writes, audits, notifications, and
statistics all consume that game-centric model. Dates and times are scraped German
strings, and schema changes require explicit Alembic migrations.

The existing early-sale notification treats the first game's two `Verkauf` helpers as
day-wide preparation staff. The requested blocks separate that concern while leaving
`Verkauf` itself per-game.

## Goals / Non-Goals

**Goals:**

- Model day-level work without fake games or dependence on whichever game is currently
  first or last.
- Reuse the concurrency, privacy, audit, contact-routing, and effective-date guarantees
  already established for game assignments.
- Keep the first version fixed at two phases and three slots per phase.

**Non-Goals:**

- User-configurable block names, slot counts, offsets, or arbitrary task definitions.
- Responsible-team ownership, MV staffing authority, or suitability grouping for blocks.
- Inferring game duration or the actual time at which the last game ends.

## Decisions

### Persist dated blocks separately from games

Add a day-block record identified by `(season_year, date, phase)`, where phase is
`preparation` or `cleanup`, plus a separate block-assignment table. Blocks are ensured by
game synchronization for every distinct season/date and are not represented by synthetic
`Game` rows. Their displayed times are derived on read from all games on the same date,
so inserting a new boundary game changes the time without moving assignments.

Alternative: store block roles on the first and last games. Rejected because assignments
would silently change meaning when boundary games change and the one-task-per-game rule
would incorrectly couple block and game duties.

Alternative: introduce a completely generic polymorphic work-item model and migrate game
assignments into it. Rejected as excessive migration and regression risk for two fixed
block types.

After a successful, validated game synchronization, reconcile the stored blocks against
the distinct season/date pairs still represented by games. If a date has disappeared,
release every occupied slot through a system-level removal path that writes an audit
snapshot, then delete both blocks and their assignments. Emit one structured application
log entry per deleted block containing its season, date, phase, and removed-assignment
count, including when that count is zero. Do not run block deletion after a failed or
partially validated scrape.

Assignments are deliberately not transferred to a different date. A split, merge, or
moved game day does not provide enough information to determine whether the same people
remain available or which new date should inherit the work.

### Use a parallel assignment model with shared service behavior

Block assignments use uniqueness constraints on `(block_id, person_id)` and
`(block_id, slot)`. The block phase determines the three valid zero-based slots and their
numbered display labels. The phase label is the semantic task role used by notifications,
statistics, audit entries, and missing-duty aggregation; the slot number does not create
roles such as `Vorbereitung 1`, matching the established treatment of `Verkauf` slots.
Dedicated claim and release helpers use the same bounded SQLite retry,
fresh compare-and-swap validation, active-person validation, audit atomicity, and conflict
response shape as game helpers.

The web API uses distinct block endpoints and target identifiers rather than overloading
`game_id`. The browser maintains taken-person state independently for every game card and
block card. This permits the same person in different containers while suppressing them
from the other selectors of the current container.

### Keep block authorization independent of teams

Members and MVs can select only themselves; admins can select any active person. Options
are alphabetical with deterministic duplicate-name ordering. Blocks carry no team foreign
key, and the game candidate grouping and warning categories do not apply. Existing CSRF,
default-deny routing, past-date restrictions, and guest redaction apply unchanged.

### Derive times rather than store them

Parse the first token of each game time as `HH:MM`. The earliest valid time minus 90
minutes supplies preparation time; the latest valid time plus 60 minutes supplies cleanup
time. Combine the clock with the parsed game date while calculating so midnight crossings
can be displayed correctly, but keep the block grouped under its home-game date. If no
valid boundary time exists, expose no calculated time and continue safely.

Storing calculated times was rejected because any game-time update would require another
state synchronization step and could leave stale reminders.

### Separate assignable roles from required game roles

Rename the game role constant and UI label from `Reinigung` to `Unterstützung`, while
keeping it in the complete assignable-role map and notification order. Introduce an
explicit required-role collection used by missing-slot statistics and MV reminders;
`Unterstützung` is absent from that collection. This avoids special-case checks scattered
through callers and leaves occupied optional assignments visible everywhere appropriate.

### Extend audit targets without rewriting history

Add an optional block reference and a block snapshot to audit entries while retaining the
existing nullable game reference and game snapshot. Each new entry has exactly one target
kind. The block reference SHALL become null rather than cascading when its block is
deleted, while the snapshot remains intact. Audit rendering and filtering accept either
target. Existing audit role strings and snapshots are not rewritten when current
`Reinigung` assignments are renamed.

### Treat block matches as date-level filter matches

Build complete date groups before filtering. Team filters continue to select games. A
person-name match on a game retains that game; a match on either day block retains the
date and all of its games that satisfy the team filters. Blocks bookend whichever games
remain visible, while their times always use the unfiltered date boundaries. Guest output
contains names only, as for game assignments.

### Route reminders by target type

One week ahead, preparation assignees receive the former special early-preparation
message using the calculated block time. The first game's sale assignees instead pass
through the ordinary weekly per-game notification. Cleanup assignees receive an ordinary
weekly block reminder. One day ahead, all occupied block slots receive a day-level
reminder. Block dispatch uses the existing email-first, phone-fallback and skip-count
contract, but no game number or MV recipient.

### Keep block cards visually aligned with game cards

Preparation and cleanup cards reuse the ordinary game-card surface and border treatment,
without a phase-colored background fade or side bar. Their phase distinction is limited to
icon badges: a `#a0e656` `↗` for preparation and a `#ffb752` `↘` for cleanup.
The missing-time state changes the time text only, preserving the same card surface.

## Risks / Trade-offs

- [A game date disappears, splits, or merges after assignments were made] -> Remove its
  blocks only after successful validated synchronization, audit every occupied-slot
  removal, log each block deletion, and never guess a transfer to another date.
- [A transient scrape failure appears to remove a date] -> Run reconciliation only after
  the existing scrape and identity validation have completed successfully.
- [String times are malformed] -> Calculate from valid times only and render a safe
  no-time state when a boundary cannot be determined.
- [Parallel assignment tables duplicate logic] -> Share validation, retry, audit, and
  response helpers where practical, with concurrency tests for both target types.
- [Role rename collides with unexpected existing data] -> Migration preflight checks for
  both role strings in the same game/slot and fails closed rather than discarding data.
- [Block matches broaden person-filter results to a whole date] -> Keep this explicit in
  the UI/specification and continue applying team filters to individual games.

## Migration Plan

1. Stop all database users and run the normal fingerprint/snapshot migration workflow.
2. Create day-block and block-assignment tables and extend assignment-audit targeting.
3. Seed preparation and cleanup block rows for distinct existing `(season_year, date)`
   game dates with non-empty dates.
4. Preflight and update current assignment roles from `Reinigung` to `Unterstützung`;
   leave assignment-audit rows unchanged.
5. Upgrade the application and verify schema head, block uniqueness, migrated assignments,
   and existing audit readability before restarting the daily job and web application.
6. On rollback, restore the migration snapshot rather than attempting an in-place schema
   downgrade.
