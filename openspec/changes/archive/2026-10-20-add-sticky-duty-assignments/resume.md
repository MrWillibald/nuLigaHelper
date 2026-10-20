# Resume: add-sticky-duty-assignments

Feature implementation resumed on 2026-10-19 after the human supplied the final policy decisions. The previous pause note is superseded. No feature commit, push, deployment, real migration or provider send has been performed.

## Authoritative policy

- Game duties and preparation, cleanup and configured cake positions are binding.
- An occupied `Ersatz gesucht` request retains its helper and responsibility until an eligible, authorized atomic takeover commits. Occupant and request-token compare-and-swap prevent stale transfers; age, collective Verkauf coverage and one-duty rules apply to the resulting group.
- The fixed cutoff is 09:00 Europe/Berlin on the previous calendar day, regardless of cron, contact availability or delivery. Weekly reminders do not lock. Empty offered positions remain claimable and immediately lock when claimed after cutoff.
- Destructive administration refuses to cancel commitments. Deactivation refuses current/future duties; erroneous-record deletion refuses any assigned duty. Removal of occupied dates/blocks and occupied cake reductions are refused atomically.
- Retained unoffered game duties only permit authorized release-only retirement before cutoff. No new claim, replacement request, takeover or lock override is available for them.
- Every UI claim/takeover requires German pre-save Bestätigen/Abbrechen confirmation. Opening, selection, Cancel, Escape and dismissal make no assignment/request write, audit or notification. A full non-JavaScript confirmation path exists.

## Workspace

- Current working directory: `/home/manu/Projects/nuLigaHelper`, branch `codex/propose-sticky-duty-assignments`, HEAD `64bc8ce6d5ead91c51b28d750a5aca05615623f4`, with the complete implementation and follow-up as uncommitted changes.
- The original detached source worktree remains preserved at `/home/manu/.codex/worktrees/4a11/nuLigaHelper`; it does not contain this later self-service search-closure follow-up. The handoff recovery ref remains `refs/codex/recovery/sticky-duty-handoff-20261019`.
- The authorized upstream PR #19 integration is complete; no re-merge is needed. Its mobile disclosure initialization remains intact.
- The source worktree has an untracked test-runtime `venv` symlink; the main directory uses its existing ignored environment. Neither is a feature file.
- Only synthetic throwaway databases/configuration and mocked providers are used for verification. Do not run the live daily job or real migrations as tests.

## Implemented

- Domain request/takeover serialization, fixed cutoff state and audits; current authorization and saved-date rechecks after writer access; late locks; pre-cutoff obsolete retirement; atomic management and synchronization refusal.
- Alembic head `0007_binding_duty_assignments`, reviewed predecessor fingerprints, retained-data postflight and stopped-writers snapshot rollback documentation.
- CSRF-protected Flask lifecycle endpoints/candidates; accessible native dialog and HTML confirmation path; persistent replacement/locked presentation across games and all block phases.
- Updated README, AGENTS, OpenSpec deltas and existing regressions, plus new domain, migration, Flask, JavaScript and optional real-browser coverage.

## Verification status — complete on 2026-10-20

The explicitly approved follow-up adds `Ersatzsuche beenden` for the current helper before cutoff across games and every block phase. JSON and non-JavaScript POST paths preserve the assignment and responsibility, compare occupant/request token, reuse current admin/MV scope, and append one actual-actor closure audit. Reopening uses a fresh token. No new schema revision is required.

Final main-checkout verification: `test/run_tests.sh` exited 0 with 435 passing tests; the two additional provider tests passed. All eight updated binding-flow browser groups passed, including search closure/reopening and cake-card focus. Earlier disclosure-only browser evidence remains valid for unchanged code. Strict OpenSpec validation and `git diff --check` pass; both debug switches are False. All 28/28 implementation checklist items are complete.

See verification.md for commands and evidence. No implementation work remains. No real configuration/database access, migration, notification, commit, push, cleanup, deployment or archival was performed for this follow-up.
