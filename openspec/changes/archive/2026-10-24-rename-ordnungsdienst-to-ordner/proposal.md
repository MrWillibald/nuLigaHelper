# Proposal

## Why

The club wants the game duty currently shown as “Ordnungsdienst” to be called “Ordner”. Current task labels, guidance and notifications must use that name consistently while preserving existing assignments and their binding state.

## What Changes

- Use the singular, unnumbered label “Ordner” throughout current schedule fields, confirmations, accessibility text, task descriptions, eligibility messages, statistics, CLI output and ordinary notifications, including advertised vacancies.
- Update the Kasse description to refer to supporting the Ordner as a second steward; keep Kasse and the independent Ordner assignment separate.
- Preserve the existing stored assignment/API role identifier `Ordnungsdienst`; resolve human-readable labels separately. Accept `Ordner` in CLI role arguments while retaining the old spelling as a compatibility alias.
- Record “Ordner” as the label on new audit transitions while retaining the original labels on existing historical audit entries.
- Preserve one security-duty position per game, category-dependent staffing totals, the age-18 requirement, authorization, confirmation, replacement and cutoff behavior. No assignment or schema migration is needed.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `game-task-staffing`: Name the independent security duty Ordner and require consistent presentation without changing assignment identity or staffing.
- `task-descriptions`: Use Ordner in description coverage and related Kasse guidance.
- `age-eligibility`: Name the existing adult-only security duty Ordner and use the label in public eligibility reasons.
- `schedule-overview`: Use Ordner in progress examples and age guidance/deficiency presentation.
- `task-self-service`: Update the independent duty's name in the Kasse responsibility contract.

## Impact

Implementation will affect role presentation in `db.py`, `age_eligibility.py`, `webapp.py`, templates and JavaScript where raw role values become visible, plus `notifier.py`, CLI argument/output handling in `manage_db.py`, tests, `README.MD` and `AGENTS.md`. Notification wording remains owned by `messages.py`; callers pass display labels into its existing named fields. Stored assignment roles, JSON mutation identifiers, candidate lookup keys, schema revisions and historical audits stay compatible. No new dependency, notification trigger or capacity change is introduced.
