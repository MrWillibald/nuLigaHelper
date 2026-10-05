# Design

## Context

DayBlock is keyed by season, scraped German date and phase. Its constraints accept only preparation/cleanup, and assignments have a fixed three-position constraint. Reconciliation creates both phases and removes vanished-date assignments with audit snapshots. Block timing, candidates, statistics and notification flows assume those phases and the shared three-position count.

See proposal.md for the reason for extending this model. The user explicitly confirmed one volunteer position per cake.

## Goals / Non-Goals

**Goals:** Extend existing day-level assignment guarantees to variable cake quantities, preserve dated settings and assignments through same-date schedule changes, and make quantity edits safe under concurrent claims.

**Non-Goals:** Arbitrary administrator-defined task types, a responsible team for cakes, inferred cake quantities or delivery times, or changing preparation/cleanup capacities and offsets.

## Decisions

### Extend dated blocks with cake metadata and per-kind capacity

Introduce a cake-delivery phase using the existing season/date/phase identity. Add nullable cake quantity and delivery clock time, meaningful only for that phase. Preparation and cleanup retain three positions. Cake position count comes from the saved whole-number quantity, with one independently assignable position per cake; numbers are display positions of one cake task role.

A separate CakeDelivery table would duplicate reconciliation, assignment, audit and authorization machinery. A generic task-builder model would exceed the requested scope. Replace global three-position assumptions with a block-capacity helper and review database constraints so variable capacity is enforced at mutation boundaries.

### Require administrator setup without guessed defaults

Recorded design assumption: new and migrated cake blocks begin with time/count unset. The schedule shows setup needed and claims remain unavailable until the administrator configures both. A configured count is a nonnegative whole number; zero explicitly means no cakes are requested and exposes no volunteer positions. It is distinct from unknown quantity. Delivery clock time is interpreted on the block's home-game date and remains independent of changed game start times.

This avoids silently choosing a club-wide default that the user did not supply. Delivery time entry uses an explicit `HH:MM` field with 24-hour validation so the browser locale cannot substitute an AM/PM picker. Existing preparation/cleanup times continue to be calculated as before.

### Preserve positions and serialize quantity edits with assignment writes

Increasing count appends empty positions. Reducing count rejects any edit that would remove an occupied numbered position, without compacting or silently releasing other assignments. Administrators explicitly release affected positions through the ordinary audited release flow first; then retry the quantity edit. Existing low-numbered positions, occupants and histories retain identity.

Configuration writes and claims must read the current capacity under the same bounded SQLite write/compare-and-swap discipline. A claim racing a reduction cannot succeed for a position removed by the committed configuration. Stale configuration edits are refused with current settings rather than overwriting a newer administrator edit.

### Reuse team-independent block authority and reporting

Members and MVs claim/release only themselves; administrators manage any active person. The one-task limit is per cake block and does not prevent duties in another game/block on that date. Cake reminders use the ordinary contact preference and day-level weekly/day-before flows, with the saved delivery time and one-cake contribution meaning. Empty slots create no helper reminder. Cake blocks have no MV recipient.

Each occupied cake position counts once in the volunteer's season statistics. Configured missing positions aggregate by the cake role. An unconfigured date is shown as needing setup rather than as fully staffed or as an invented count of missing cakes.

### Keep overview deltas independent

Add cake-specific overview requirements rather than replacing the staffing change's expanded-game or progress requirements. Reuse date, team and person filter semantics and guest redaction. Render the cake card in a stable date-level position after preparation and before game cards; its displayed delivery clock is administrator-controlled rather than used to redefine the existing bookends.

## Risks / Trade-offs

- [A quantity edit races an assignment claim] -> Serialize writes, validate current capacity and refuse stale settings.
- [An administrator reduces away an occupied position] -> Reject the reduction until an explicit audited release occurs.
- [A scrape removes a date transiently] -> Reconcile only after successful validated synchronization, retaining existing block-removal safeguards.
- [New phase reaches a fixed-count or preparation/cleanup-only caller] -> Audit every phase/count consumer and verify variable-capacity integration.
- [Unknown settings look like complete coverage] -> Expose a setup-needed status distinct from configured zero or complete staffing.

## Migration Plan

Stop database users and use the existing fingerprint/snapshot migration workflow. Add the cake phase and metadata, adjust slot constraints, seed one unconfigured cake block for every existing dated game group, and verify uniqueness and all retained preparation/cleanup assignments and audits. Do not manufacture cake volunteers, time or quantity. Update application schema fingerprints/head, run postflight integrity checks and restart only after success. Rollback restores the validated snapshot. Rebase the ordered Alembic revision against whichever independently reviewed proposal is applied first.
