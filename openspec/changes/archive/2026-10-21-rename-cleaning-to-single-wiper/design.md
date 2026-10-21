# Design

## Context

See `proposal.md` for motivation. Persisted role values also serve as German API/CLI labels. Current Reinigung has two positions; adult offered requirements derive from that catalog. The current head is 0007 with saved locks and replacement tokens.

The user explicitly requested deletion of existing second assignments after reviewing the original preservation plan. This overrides the repository's commitment-preservation rule solely for this one-time migration. General runtime assignment, lock and destructive-maintenance rules continue to apply. The historical pre-day-block Reinigung role has a different meaning and must still migrate to Kasse.

## Goals / Non-Goals

**Goals:** Keep one canonical Wischer role and offered position; delete exactly current Reinigung slot 1; preserve first assignments, unrelated data and existing audit history; record each deletion accurately.

**Non-Goals:** Create a runtime override, alter other duties or age rules, move helpers between positions, rewrite historical audits or execute a production migration during planning.

## Decisions

### 1. Rename the canonical role and reduce capacity

Change the existing cleaning role constant's value to `Wischer` and its offered capacity to one. Keep the English constant name if convenient. Update web slot catalogs, JSON/CLI validation, JavaScript role lists, notifications and the singular German description together. The normal label is `Wischer`; no second Wischer display entry remains. Reject stale Reinigung mutation inputs and unsupported Wischer slot 1.

A display-only alias would require a new translation layer across callers; the established canonical-role model fits the change. Historical migrations and historical fixture literals retain their recorded spelling.

### 2. Delete second assignments in the guarded migration only

The exact source set is current `assignments.role = 'Reinigung' AND slot = 1` from the game-duty era. Include all dates and categories, even past or locked duties and retained youth duties. Remove those rows and their embedded request state. Do not delete persons, games, other assignments or existing audits. Rename only current Reinigung slot 0 to Wischer, preserving all other fields.

No second occupant is promoted into an empty slot 0. That adult slot remains a required vacancy. Deleted duties stop contributing to personal assignment statistics and reminders, including past assignment counts, while historical audit evidence remains. Saved tokens for deleted duties become stale and cannot recreate a second position.

Keeping a retained obsolete second field was superseded by the explicit user instruction. Automatic promotion was rejected because it would invent responsibility for the remaining position. Runtime release helpers must retain their cutoff checks; use self-contained reviewed migration SQL rather than introducing a bypass flag into application writers.

### 3. Preserve old audits and append removal evidence

Before each targeted deletion, append one system-origin removal entry in the same migration transaction using the existing audit schema. Identify the original Reinigung role, slot 1, affected person and game snapshot, with a readable migration-specific origin/action such as `migration_remove`. The entry must distinguish the deliberate capacity migration from voluntary helper withdrawal. Include no full birth date or exact age. Existing audit rows remain byte-for-byte unchanged, and renaming slot 0 does not itself fabricate an assignment event. No outbound notifications run in the offline migration.

Extend migration verification to compare pre-existing audit rows unchanged and validate the exact appended removal set. Do not exempt the whole audit table from fingerprints. Validate that only the planned assignment rows disappeared and only allowed role values changed. An already-current rerun must append nothing.

### 4. Keep source-revision-aware transformations

Append one successor to the single head and update `schema_migrations.py` source validation and retained-data expectations. Before the day-block revision, original Reinigung maps to Kasse through the existing chain. Unterstützung likewise maps to Kasse where applicable. Only current Reinigung from revision 0006 onward is subject to the new slot-specific rename/delete transformation.

Preflight classifies the exact mutation set and rejects unexpected target Wischer rows, collisions and unsupported slots before domain-data changes. The postflight allowance is exact: slot-0 role renames, deletion of identified slot-1 rows, one audit per deletion and the revision update, plus reviewed earlier pending changes. No broad role-string replacement or generic fingerprint exclusion is acceptable.

## Risks / Trade-offs

- Second helpers lose their saved duty, including past statistics → the explicit request is scoped to slot 1; preserve existing history, append removal evidence and retain the validated full backup.
- Historical Reinigung text can mean Kasse → test upgrades from before day blocks, before game-duty staffing and the current binding head with revision-specific expected rows.
- Open replacement tokens survive in old pages → verify old requests fail without assignment recreation, successful takeover audit or notification.
- A broad deletion or relaxed fingerprint could hide unrelated data loss → precompute exact source ids and compare all other rows and existing audits unchanged.

## Migration Plan

1. Implement and verify using synthetic offline databases; document the single migration exception and seven adult positions in `README.MD` and `AGENTS.md`.
2. Stop web and daily writers and run `manage_db.py migrate-schema --confirm-stopped` using its validated backup, preflight and postflight. Review the exact rename/delete set through private operator diagnostics.
3. Verify the new head, unchanged slot-0 identity/lock/request state, absence of every targeted slot-1 assignment, exact removal evidence and unchanged historical snapshots before restarting writers.
4. On verification or deployment failure, keep writers stopped and restore the printed pre-migration snapshot with the previous application. This restores deleted assignments too; do not downgrade in place.
