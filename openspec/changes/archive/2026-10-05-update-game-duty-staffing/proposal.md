# Proposal

## Why

Adult home games need additional staffing and a clearer Kasse role. Youth games need only the established five game duties, without offering extra game duties.

## What Changes

- **BREAKING** Rename the current Unterstützung/Unterstützungsdienst role to Kasse while preserving every current assignment and historical audit snapshot.
- Retain exactly one independent Ordnungsdienst position. Kasse also performs the additional ordering/security duties originally planned for a second Ordner; it remains a separate role and never relaxes one task per person per game.
- Adult games offer eight required positions: Zeitnehmer, Sekretär, Verkauf 1/2, Ordnungsdienst, Kasse and Reinigung 1/2.
- Youth games and unsupported categories offer only the baseline five required positions: Zeitnehmer, Sekretär, Verkauf 1/2 and Ordnungsdienst. Kasse and game Reinigung are not offered and refuse new claims there. No game duties are optional.
- Use the same offered/required positions for candidates, UI, CLI, progress, open-duty statistics and MV reminders. Existing age rules remain authoritative.
- Preserve any current assignment in a now-removed duty as a visible existing assignment with release-only maintenance under the existing rights. It remains in personal statistics and helper reminders, without creating required staffing or new candidates. Releases keep append-only audits.
- Keep newly introduced Reinigung distinct from legacy Reinigung assignments previously migrated to Unterstützung. Preparation, day cleanup and cake blocks remain independent.

## Capabilities

### New Capabilities

- `game-task-staffing`: Defines category-dependent offered and required game positions and retained-assignment maintenance.

### Modified Capabilities

- `task-self-service`: Replaces optional Unterstützung with Kasse and preserves existing assignments through migration.
- `schedule-overview`: Displays each game's offered positions and retained removed assignments, and calculates progress from the offered/required set.

## Impact

Role definitions, guarded data migration, schema fingerprints and fixtures, repeated-role occupants and candidates, schedule templates and JavaScript, statistics, notifications, CLI and README. Reuse the shared age-eligibility game-category classifier. Other open changes stay outside this scope.
