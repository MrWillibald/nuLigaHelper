# Verification: add-sticky-duty-assignments

Verified on 2026-10-19 in the local worktree, on top of the already integrated
upstream commit `64bc8ce6d5ead91c51b28d750a5aca05615623f4`.

## Final results

| Check | Actual result |
| --- | --- |
| `test/run_tests.sh` | Exit 0; 426 passing tests; 54 Python files invoked |
| `./venv/bin/python -m pytest test/test_provider_env.py -q` | Exit 0; 2 passed; this file has no standalone runner |
| Optional `node test/browser_binding_assignments.mjs` | Exit 0; all 7 native-browser groups passed |
| Existing `node test/browser_disclosures.mjs` | Exit 0; all 36 browser groups passed |
| `openspec validate add-sticky-duty-assignments --strict` | Valid |
| `git diff --check` | Clean |
| Committed debug defaults | `DEBUG_FLAG = False`, `CHANGE_DAY = False` |

The final complete runner log is `/tmp/nuliga-sticky-green-suite.log`. Browser
logs are `/tmp/nuliga-sticky-browser.log` and
`/tmp/nuliga-sticky-disclosures.log`; the latter's structured report is
`/tmp/nuliga-sticky-disclosures.json`. These temporary files are convenience
evidence, not required application files.

The optional browser commands used installed Playwright via `PLAYWRIGHT_MODULE`
and headless Chrome. Both served only synthetic loopback fixtures, blocked page
requests to external origins, and are documented in `test/README.md`.

## Coverage of the confirmed decisions

- Domain tests cover occupied replacement requests, idempotent request opening,
  occupant/token CAS, stale generations, competing takeovers, continuous
  responsibility, current authority and one-duty constraints, resulting Verkauf
  adult coverage, and paired immutable audit evidence.
- Fixed-cutoff tests cover immediately before/at 09:00 Europe/Berlin, calendar-day
  arithmetic across both DST transitions, missing cron markers, delivery failure,
  missing contacts, weekly reminders, late empty claims and a real contended
  writer crossing the cutoff. Saved locks survive metadata changes, and repeated
  enforcement does not duplicate transition audits.
- Management regressions verify refusal of current/future deactivation, deletion
  with any assignments, occupied day/block removal and cake reductions. Tests
  cover stale loaded empty relationships and cached past dates after a committed
  reschedule. Past-only deactivation preserves duties/history. Only authorized
  pre-cutoff retirement of retained unoffered game duties remains available.
- Flask and JavaScript tests cover every game role and all three block phases,
  member/MV/admin scope, CSRF, guest privacy, persistent lifecycle status,
  candidate refresh, conflict recovery, retained duties, physical progress and
  separate age deficiencies. Existing notification/statistics tests remain green;
  opening a replacement request adds no outbound broadcast.
- Real-browser checks exercise native confirmation before saving, Cancel/Escape
  and returned focus, request plus atomic takeover for games/preparation/cakes/
  cleanup, a concurrent winner preserved on Cancel and confirmed CAS refusal,
  immediate late locks, and JavaScript-disabled review/Cancel/Confirm. Existing
  mobile disclosure initialization and fallback regressions also pass.
- Migration tests verify a single Alembic head, fresh-schema agreement, guarded
  upgrades from reviewed historical revisions, retained identities/memberships/
  MV appointments/positions/old audit snapshots, refusal of predecessor drift,
  and no invented confirmation, request or historical lock state.

Existing release-based tests were updated to require the confirmed binding-duty
behavior; they now assert preserved occupants and refusal or eligible atomic
takeover, rather than expecting ordinary withdrawal to create a vacancy.

## Operational boundary

All database/configuration inputs and secrets used for verification were synthetic;
provider transport was mocked. No real database migration, provider send, commit,
push or deployment was performed. The untracked local `venv` symlink is test
runtime infrastructure and must be excluded from any later feature commit.

Before using this version with an existing database, stop every writer and run the
documented guarded `migrate-schema --confirm-stopped` workflow to reach
`0007_binding_duty_assignments`. Startup does not migrate implicitly. The workflow
creates and verifies a snapshot; rollback restores that snapshot with the previous
application and writers stopped. No implementation checklist item remains open.


## Approved follow-up: end an own replacement search (2026-10-20)

Implemented and verified directly in `/home/manu/Projects/nuLigaHelper`, on
`codex/propose-sticky-duty-assignments`, preserving the prior local implementation
and source-worktree recovery state.

The current active helper can select **Ersatzsuche beenden** for their own offered
game duty or preparation/cake/cleanup position before the existing cutoff. The
same assignment, person and responsibility remain saved. Existing admin and game
MV management scope is retained; foreign or former occupants gain no rights.
The new game/block `cancel-replacement` POST routes and non-JavaScript form action
share the existing serialized domain boundary, compare current occupant plus
request token, enforce the fixed cutoff after waits/retries and write exactly one
`request_close` event with the actual actor. A reopened search has a fresh token.
The schema remains `0007_binding_duty_assignments`; no further migration is needed.

Final verification in the main checkout:

- `test/run_tests.sh`: exit 0, **435 passed**, 55 Python files invoked; log
  `/tmp/nuliga-cancel-final-suite.log`.
- `./venv/bin/python -m pytest test/test_provider_env.py -q -p no:cacheprovider`:
  exit 0, **2 passed**; log `/tmp/nuliga-cancel-provider.log`.
- Updated `test/browser_binding_assignments.mjs`: exit 0, **8 real-browser
  groups passed**; log `/tmp/nuliga-cancel-browser-final.log`. These verify native
  search closure/reopening and unchanged responsibility for all four container
  types, focus after cake-card refresh, non-JavaScript closure and confirmation,
  concurrent claim conflicts and immediate late locks. Unchanged disclosure-only
  browser groups retain the earlier evidence and were not rerun for this follow-up.
- New synthetic cancellation tests cover current ownership independently of the
  search opener, admin/MV scope, guest/inactive/unapproved sessions, CSRF, stale
  closed/reopened tokens, former ownership, a real cancellation/takeover race,
  and contended writers crossing 09:00 for games and blocks.
- Strict OpenSpec validation and `git diff --check` passed; all **28/28** tasks
  are complete. Both debug switches remain False.

Only synthetic databases, configuration and mocked providers were used. The
user's migrated database/configuration were not read or changed; no real provider
send, new migration, commit, push or deployment was performed.
