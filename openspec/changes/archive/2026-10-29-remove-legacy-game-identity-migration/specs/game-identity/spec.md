# Game Identity Delta

## REMOVED Requirements

### Requirement: Existing databases migrate without silent task-data loss

**Reason**: The source-key to canonical game-number migration has been completed on every deployed database; keeping the converter adds code, tests and operator instructions without remaining use.

**Migration**: A database that still has the legacy `games.source_key` schema must first be converted with an earlier release that contains `manage_db.py migrate-game-identity --confirm-stopped`. The current release refuses such a database unchanged (see `schema-migrations`).
