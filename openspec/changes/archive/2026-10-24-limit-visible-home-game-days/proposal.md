# Proposal

## Why

Showing the entire season makes the next home game days harder to find and spreads attention across distant staffing needs. Administrators need one shared date boundary that focuses the home schedule and outstanding-duty statistics on nearer dates while retaining all games and binding assignments.

## What Changes

- Add an administrator-managed, persistent optional setting labeled "Heimspieltage ausblenden ab" on the schedule. The chosen date is the first hidden date; clearing the setting restores all dates.
- Hide complete dates on or after the boundary for guests, members and MVs, including games, preparation, cake delivery and cleanup. Apply the boundary before ordinary filters, date choices and date/month headings are built.
- Apply the same boundary to "Offene Dienste, Altersanforderungen und Einrichtung" in statistics and its affected-container count. Keep "Spiele pro Mannschaft", "Dienste pro Person" and season totals based on the complete season.
- Let administrators explicitly show hidden dates on both pages without changing the shared setting. Display clear German explanations of the active scope and meaningful empty states.
- Keep this a presentation setting: scraping, synchronization, notifications, binding assignments, locks, existing authorization and direct authorized assignment flows continue independently of the boundary.
- Introduce the setting through the existing guarded Alembic migration workflow, initially unrestricted, without changing existing domain rows or historical audits.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `schedule-overview`: Persist and manage the shared date boundary; filter complete day groups and date options; explain filtered scope; support an administrator full-date view.
- `statistics-overview`: Limit outstanding games, age deficiencies and day-block setup needs to visible dates while retaining complete-season distribution and personal totals.
- `access-control`: Restrict setting writes and full-date overview overrides to current active administrators without treating visibility as assignment authorization.
- `schema-migrations`: Add unrestricted visibility-setting storage through an explicit backup-first successor revision while preserving existing commitments and history.

## Impact

- `db.py`: setting model, explicit setting-write helper and shared pure date-visibility policy used only by presentation callers.
- `webapp.py`, `templates/schedule.html`, `templates/statistik.html` and small shared template/style additions as needed: setting form, view override, visibility notices and filtering integration.
- `migrations/versions/`, `schema_migrations.py`, initialization and offline migration fixtures: reviewed successor to the current `0008_single_wischer` head, schema recognition and retention verification.
- Offline tests for settings, tier/CSRF enforcement, filtering, statistics, migration, preserved notifications and assignment behavior; `README.MD` for usage and deployment/recovery instructions.
- No new external dependency, provider integration or notification wording is required. Existing deployments must run the stopped-writers schema migration before starting the new application.
