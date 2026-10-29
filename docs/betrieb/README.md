# Operations documentation

Operator procedures for an installed nuLigaHelper. Behavior rules are specified in
[`openspec/specs/`](../../openspec/specs/); the user manuals are listed in
[docs/benutzer/](../benutzer/README.md).

| Document | Content |
| --- | --- |
| [sqlite-and-backups.md](sqlite-and-backups.md) | SQLite connection profile, WAL, daily lock, Dropbox snapshots and retention, restoring a downloaded snapshot |
| [schema-migration.md](schema-migration.md) | Explicit initialization, guarded `migrate-schema --confirm-stopped` upgrades, recovery; legacy game identity is refused |
| [cli.md](cli.md) | `nuligahelper.cli` commands for persons, memberships, games, assignments and contact preflight |
| [notification-wording.md](notification-wording.md) | `nuligahelper/jobs/messages.py` catalog, legacy `club.texts` preflight and referee recipients |
| [private-deployment.md](private-deployment.md) | Production topology, legal pages, environment variables |
| [deployment.md](deployment.md) | Routine release update on the existing installation |
| [../../release-assets/README.md](../../release-assets/README.md) | Host setup and release command sequence |
| [data-inventory.md](data-inventory.md) | Stored data and its visibility |

---
Back to the [README](../../README.md)
