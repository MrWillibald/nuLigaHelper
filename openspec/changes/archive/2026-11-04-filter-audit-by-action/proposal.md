# Proposal

## Why

The Änderungsprotokoll can only be filtered by target and person. With system entries such
as `lock` and `unlock` written for whole dates, administrators have to scan long lists to
answer questions such as "which locks were released by the last sync?". The deployment
guide already asks operators to review `unlock` entries.

## What Changes

- The audit page gets a third filter, `Aktion`, with `Alle Aktionen` as the default.
- It lists the actions actually stored in the audit, independently of the other active
  filters, with German labels; an action without a label appears under its stored name.
- It combines with the target and person filters and is reset by `Zurücksetzen`.
- The `Aktion` column uses the same German labels, so a filter choice matches what the
  table shows. Stored audit rows are not changed.

Non-goals: multiple selection, date-range filters, pagination or export.

## Capabilities

### New Capabilities

_None._

### Modified Capabilities

- `assignment-audit`: adds the administrator filter by action.

## Impact

- `nuligahelper/web/statistics.py` (`audit()`), `web/formatting.py` (label map),
  `web/templates/audit.html`, `web/static/style.css` (`.audit-filter-card` grid).
- Read-only: no schema or data change.
- Docs: `docs/benutzer/BENUTZERHANDBUCH.md` audit section.
