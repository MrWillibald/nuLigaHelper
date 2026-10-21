# Proposal

## Why

The club's revised duty rules require adult Ordnungsdienst helpers, allow Zeitnehmer at adult games from age 16, and remove sale age restrictions at youth games. The shared eligibility checks and staffing reports must reflect these rules consistently across assignments, takeovers, candidate lists and follow-up.

## What Changes

- **BREAKING**: Require a known birth date and age 18 on the game date for Ordnungsdienst at every game, including youth/SPF and unresolved categories.
- Lower the adult-game Zeitnehmer minimum from 18 to 16; retain Sekretär at 16 for adults and both timing roles at 14 for youth/SPF.
- Exempt recognized youth games, including SPF, from all Verkauf age requirements, including collective adult coverage and birth-date requirements for eligibility.
- Retain collective adult Verkauf coverage for adult games. Unresolved categories keep the existing coverage rule; they do not receive the youth exemption.
- Reevaluate current/future staffing under the revised rules without deleting or rewriting assignments, locks, requests or historical audits. Expose new Ordnungsdienst deficiencies and remove obsolete youth-sale/adult-Zeitnehmer deficiencies from current reporting.
- Update German schedule guidance, documented rules and offline regression coverage.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `age-eligibility`: Revise individual duty thresholds, scope Verkauf coverage to non-youth games, and apply the revised rules atomically and to retained staffing.
- `schedule-overview`: Explain and display the revised requirements without showing youth-sale adult-coverage warnings or changing physical occupancy.

## Impact

- Shared policy in `age_eligibility.py`, claim/takeover validation and `staffing_status()` in `db.py`.
- Web candidates and JSON mutations, non-JavaScript confirmations, CLI assignments, schedule/statistics and existing MV follow-up consuming shared policy.
- `templates/schedule.html`, the German task descriptions in `db.py`, `README.MD` and the age-rule sections of `AGENTS.md` during implementation.
- Existing synthetic eligibility, reporting, binding-duty and concurrency tests; full `test/run_tests.sh` regression run.
- No database schema migration, new provider integration or new notification type is needed. Other authorization, commitment, cutoff and privacy guarantees remain in force.
