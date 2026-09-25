## Disposable VM rehearsal — 2026-09-25

This was a Debian 13 amd64 QEMU guest with Python 3.13.5, synthetic configuration,
synthetic SQLite data, installed systemd units, and dedicated Caddy ingress.
No production database, secret, hostname, or service was used. The guest started
with the exact production baseline source commit
`ea0ab0067e73a076f64b31fedb8ed9614764ab42` and prepared the merged
protected-master commit `4e4bc06807889dad1410051742ea3b65d2389576`
(tree `bc820f3aff738303b9ec6f826dfc442b4335eb96`). The candidate's full
offline suite passed in its isolated virtual environment. The revised host
deployer used for these fault injections had SHA-256
`95e676b6f237d57fce70cc8c8abbb5ffc63856438a1e2b6099b2a6b11da0e4ff`;
it must still be merged through protected `master` before production use.

The VM database used the approved path `/var/lib/nuligahelper/nuliga_helper.db`
and had one synthetic person, game, assignment, and assignment audit. Each
integrity check returned `ok` and no foreign-key violations. No daily or
cleanup service was invoked as a smoke test.

| Exercise | Observed result |
| --- | --- |
| Failed candidate health / 60-second readiness deadline | Candidate Gunicorn bound the wrong loopback port. Activation record `236fd3a600cb406f9c0385c9c5bba90f` ended `failed`, with snapshot `/var/backups/nuligahelper-deploy/snapshot-236fd3a600cb406f9c0385c9c5bba90f.db`; Caddy and all timers stayed inactive, the database remained at `0003_game_day_task_blocks`, and row counts stayed `1/1/1/1`. |
| Compatible code-only rollback with delayed prior-release start | `rollback-code` for that record waited for old Gunicorn readiness, then returned `rolled_back`; `current` resolved to the `ea0ab00` release, Caddy and web were active, HTTPS `/healthz` returned `ok`, all timers remained inactive, and revision/counts were unchanged. |
| Migration-required gate | A synthetic recognized `0001_current_schema_baseline` database replaced the stopped VM fixture only after a separate validated copy of the head database was retained. Activation record `6b2ebcc338a64f208d16ad54e319f969` stopped as `migration_required` with snapshot `/var/backups/nuligahelper-deploy/snapshot-6b2ebcc338a64f208d16ad54e319f969.db`; the snapshot passed integrity and foreign-key checks and kept counts `1/1/1/1`. The prior release remained selected, with Caddy, web, and timers inactive. |
| Explicit migration and continuation | `manage_db.py migrate-schema --confirm-stopped` upgraded the VM database through `0002` to `0003`, creating its own pre-schema backup. `continue` returned `public_ready`; `current` resolved to the prepared `4e4bc068` release, Caddy and web were active, HTTPS `/healthz` returned `ok`, and all timers remained inactive. Integrity, foreign keys, revision, and row counts passed again. `resume-timers --catchup hold` recorded the deliberate hold and did not start timers. |

Unit tests also inject failed units and both candidate/rollback readiness
timeouts, asserting that ingress and timers stay closed and no automatic
database replacement occurs. The final production gate remains open: publish
the revised deployer through a green PR into protected `master`, then run the
host preflight and obtain operator review before any live cutover.
