# Proposal

## Why

The per-game cleaning duty should be called “Wischer” to describe its actual work during play. Adult games need one person for this duty rather than two, so staffing requirements and all current presentations must reflect that smaller need.

## What Changes

- Rename the current per-game role Reinigung to Wischer across schedule labels, assignment flows, task help, CLI, statistics and ordinary notifications.
- **BREAKING**: Offer and require only Wischer position 0 for adult games, reducing their required total from eight to seven. Youth/SPF and unresolved categories retain their five-position baseline without Wischer.
- **BREAKING**: Delete every existing second current Reinigung assignment (position 1) during the guarded upgrade, across all dates and game categories, including locked duties and open replacement requests. This is the explicitly requested one-time exception to commitment preservation; position 0 is retained and is never filled by moving position 1.
- Migrate current assignment role values through the explicit backup-first Alembic workflow. Preserve first-position identities and binding/replacement state; preserve existing historical audit snapshots verbatim and append one system migration-removal entry for each deleted second assignment. Earlier historical Reinigung duties that the existing migration chain maps to Kasse remain Kasse.
- Update documentation and tests for the new name, totals and deletion of the second position. Day-wide Aufräumen remains separate.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `game-task-staffing`: One offered adult Wischer position, seven required adult positions, consistent current naming and one-time removal of existing second assignments and subsequent reporting.
- `task-descriptions`: Description coverage and singular German wording for Wischer.
- `schema-migrations`: Guarded rename of current first Reinigung assignments and audited deletion of second assignments, preserving other commitments and historical records across supported upgrade paths.

## Impact

Affected areas include `db.py` role/position helpers and descriptions, web slot catalogs and JSON validation, `static/app.js`, CLI role handling, notifier task labels, statistics, the Alembic migration chain and retained-data verification, offline tests, `README.MD` and `AGENTS.md`. Current role values exposed to API/CLI callers change to Wischer; stale Reinigung writes must fail without creating assignments. No new dependency or provider configuration is needed.
