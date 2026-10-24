# Design

## Context

See `proposal.md` for motivation and the capability deltas for the behavior contract. `webapp.py` currently builds the complete current-season schedule, derives date options, applies ordinary filters and groups games with their dated blocks. Statistics separately query all season games and blocks, calculate distribution/personal totals, and construct outstanding-container rows. No persistent web-managed global setting currently exists.

Game and block dates remain German strings. Existing parsers and `db.game_sort_key()` must remain authoritative for date interpretation and ordering. The current schema head is `0008_single_wischer`; migrations use explicit offline preflight, validated backups, exact source-schema checks and retained-data fingerprints. Saved assignments, locks, replacement generations and audit history must survive this change unchanged.

## Goals / Non-Goals

**Goals:**

- Read one shared setting snapshot per request and reuse one visibility predicate for schedule dates and outstanding statistics.
- Keep full-season data available for calculations and authorized direct assignment operations.
- Make setting changes explicit, validated, protected and usable without JavaScript.

**Non-Goals:**

- No publication workflow, new assignment rights, rolling time window, automated boundary movement or notification suppression.
- No per-person/team boundary, extra season selector, CLI settings command, generic settings framework or new assignment-audit category.
- No conversion of scraped date columns or changes to staffing and cutoff calculations.

## Decisions

### 1. Store one optional global date in SQLite

Add a dedicated singleton `ScheduleVisibilitySetting` model/table with a constrained fixed key and nullable date-only `hide_from`. Explicit initialization and migration seed one unrestricted row. The setting is global and persists across web workers, restarts and season changes; it has no automatic expiry. A missing row after initialization is an integrity/configuration error rather than a reason for a GET to insert data.

This is the smallest durable setting shared by every viewer. `config.json` is currently operational input and would require file permissions, cross-process file writes and cache invalidation for a web form; a cookie or session value would not create shared visibility. A season-specific table or key/value settings system adds behavior the request does not need.

The form accepts an ISO date through an HTML date input; notices use German formatting. A newly chosen or changed nonempty date must be strictly after `common.effective_today()`. Clearing is always allowed. Resubmitting the identical saved value is a permitted no-op even if that date has since passed. A retained date continues to apply after time advances, including to dates now in the past; the notice and admin full-date view make that state recoverable. These validation/lifecycle details are implementation assumptions for this proposal.

### 2. Isolate the presentation predicate

Add a pure shared helper accepting a scraped date and the saved boundary. A valid date is visible when the boundary is empty or the date is strictly earlier than it. An unparseable scraped date remains visible under existing "Ohne Datum"/unresolved-date handling: the setting must not silently hide a data problem. Parse calendar dates rather than comparing German strings.

Call the helper only while preparing overview results. Do not change general database query helpers, scraping/sync, notifier queries, assignment/candidate routes, duty locks or `staffing_status()`. This preserves authorized direct service operations for hidden dates and prevents an apparently harmless overview filter from suppressing reminders or binding enforcement.

### 3. Apply visibility before schedule choices and grouping

Load season games and the saved setting, resolve the viewer's authorized full-date mode, and select visible games before building date options, ordinary-filter inputs and date groups. Derive block inclusion from the same dated groups, including cake delivery. Existing block-time calculations continue using the full set of that date's games, before ordinary team/person filtering. Use the full-season game count only to distinguish an actually empty season from a limited view with no visible content.

Ordinary query parameters cannot widen the date scope. A manually entered hidden date has no matching cards and is labeled "Unbekannter Spieltag" by the existing fallback rather than confirming hidden game details. Clearing ordinary filters preserves the effective visibility mode. Existing card rendering, privacy, month headings and past-section behavior remain intact within visible dates.

### 4. Give administrators an explicit request-local full-date view

Use `show_hidden=1` on the schedule and statistics, honored only when the freshly derived tier is admin. Both pages initially use the same restricted view as other viewers. Provide "Ausgeblendete Spieltage anzeigen" and "Ausgeblendete Spieltage wieder ausblenden" as GET links; preserve this mode through ordinary admin filter submissions and clear-filter actions. Carry the mode between the two page links when navigating from an admin full-date view. Do not save it as a global setting or grant it to lower tiers through copied URLs. Removing admin rights takes effect on the next request.

The admin setting form lives on the schedule. Use a CSRF-protected POST endpoint with post/redirect/get. Submit the expected saved boundary along with the proposed boundary. The write helper obtains a SQLite writer serialization point before rereading the singleton and active admin record, compares the expected value, validates the proposed date, then stages only the setting update. A conflicting edit returns a German message and refreshed saved value without overwriting it; transaction-busy handling is bounded by existing SQLite policy. The expected value must distinguish an explicitly empty boundary from a missing request field.

An ordinary read, preview toggle or form cancellation never writes settings, assignments, lock markers, audits or notifications. Do not add this settings POST to the assignment-lock preflight endpoint set: it is not a schedule duty mutation.

### 5. Filter only the outstanding part of statistics

Build `team_stats`, `person_stats` and season denominators from the complete season exactly as today. Apply the shared predicate to games and blocks when building `gaps`, in addition to existing current/future rules. Consequently its length, heading count, detail rows and empty state all describe the same visible subset. In admin full-date mode only that subset restriction is lifted; past dates still do not become outstanding rows.

Use notices near the affected content: "Heimspieltage ab dem 01.12.2026 sind derzeit ausgeblendet." and a statistics explanation that distribution/personal totals cover the whole season. When no gaps remain under an active restriction, say "Im angezeigten Zeitraum gibt es keine offenen Dienste, ungeklärten Altersanforderungen oder erforderlichen Einrichtungen." Do not claim that the complete season is staffed. With no visible schedule content and an existing season, explain the date restriction rather than asking the user to load games.

Reuse existing German feedback, native form/link behavior and responsive layout. The setting and override need no custom JavaScript. New controls and notices must wrap at narrow widths and retain labels and keyboard focus.

## Risks / Trade-offs

- A saved fixed date can eventually hide recent/past days as well as future days. Keep its actual value visible, never auto-reset it, and let admins clear/change it or view all dates.
- Hidden duties still generate notifications and may be reachable through authorized direct URLs. Explain the display-only scope beside the setting; test those paths instead of adding a publication barrier.
- Complete-season totals can differ from the outstanding view. Label their scopes and keep the existing count units.
- Concurrent admin edits could overwrite each other. Compare the expected saved boundary under writer serialization and require review after a conflict.
- Adding a new model can make older recognized schemas look like drift. Build exact revision-specific expected metadata for the newly prior `0008_single_wischer` schema and adjust older verifiers to omit only the new table where appropriate; retain rejection of near-miss columns, indexes and constraints.

## Migration Plan

1. Add one reviewed successor revision (suggested `0009_schedule_visibility`) with `down_revision = "0008_single_wischer"`; create/seed only the new unrestricted settings row. Keep `WISCHER_REVISION` tied to its original revision and advance the required head separately.
2. Update exact source-schema inspection, historical expected metadata, head/drift checks, initialization and migration fixtures. Existing `0008` databases become recognized prior sources; older supported chains still apply their separately reviewed transformations once.
3. Add change-specific postflight verification for the singleton key, nullable date and unrestricted initial value, alongside unchanged retained-domain/audit fingerprints. A repeated migration at head must preserve any subsequently configured date.
4. Run offline tests and strict OpenSpec validation. For deployment, stop all web/daily writers, run `manage_db.py migrate-schema --confirm-stopped`, retain the printed validated snapshot and restart only after postflight succeeds. No live deployment or database migration is part of this planning change.
5. For recovery, stop writers, restore the validated pre-migration snapshot and matching prior application. Do not downgrade an active database in place. Document that older pending revisions retain their own already-reviewed effects, including the one-time `0008` second-duty deletion; this revision introduces no further assignment exception.
