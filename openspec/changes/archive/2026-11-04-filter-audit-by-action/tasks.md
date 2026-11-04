# Tasks

## 1. Filter

- [x] 1.1 Add `AUDIT_ACTION_LABELS` to `nuligahelper/web/formatting.py` (`claim` Zusage, `release` Freigabe, `request_open` Ersatzsuche geöffnet, `request_close` Ersatzsuche beendet, `takeover_out` Übernahme abgegeben, `takeover_in` Übernahme angenommen, `lock` Sperre, `unlock` Sperre aufgehoben, `migration_remove` Durch Migration entfernt, `reopen` Auf offen gesetzt, `date_remove` Spieltag entfallen) and a helper that falls back to the stored name. Verify with a unit test for a known and an unknown action
- [x] 1.2 In `audit()` in `nuligahelper/web/statistics.py`, read `action`, filter on it, and pass `selected_action` plus all distinct stored actions from an unfiltered query, sorted by label. Verify in `tests/web/test_management_ui.py`: single action, combination with target and person, list independent of other filters, unlabelled action, unknown value gives an empty list with status 200, non-admin refused
- [x] 1.3 Add the `Aktion` select to `audit.html` and adapt the `.audit-filter-card` grid in `web/static/style.css` for desktop; the existing narrower breakpoints already reflow it. Verify the selection is retained after filtering and `Zurücksetzen` clears it
- [x] 1.4 Render the German label in the `Aktion` column, keeping the `audit-action-<key>` CSS class on the stored key and replacing the inline `migration_remove` special case. Verify existing audit tests, updating only assertions on raw action text
- [x] 1.5 Update the audit section of `docs/benutzer/BENUTZERHANDBUCH.md`: the filter steps and the action list with the German labels, and the note that some names are technical. Verify the labels match `AUDIT_ACTION_LABELS`

## 2. Integration check

- [x] 2.1 Run `tests/run_tests.sh` and verify the whole suite is green; run `openspec validate filter-audit-by-action --strict` and verify it passes

## Workflow follow-up

- Archive the change after review and sync the delta specs into `openspec/specs/`.
