# Notification wording and migration

`nuligahelper/jobs/messages.py` is the single editing point for retained notification subjects,
e-mail/SMS bodies and fragments. Its catalog uses explicit keys for day-before
and weekly reminders, ordinary games and Spielfeste, day blocks, account actions,
operator alerts and the dormant newspaper feature. Each entry declares the
required named fields, with their meanings documented beside the catalog.
For example, `{recipient_name}`, `{task_label}`, `{game_date}` and
`{game_start_time}` identify a game duty, while `{block_meeting_time}` identifies
the meeting or delivery time of a day block. Display labels come from domain
data, so numbered positions do not become separate task names in notifications.

Developers edit the relevant catalog entry, preserve the named field contract,
then review synthetic rendered examples and run `tests/run_tests.sh`. When a
contract changes, update its declaration, meanings and callers together.
`messages.render(key, **context)` returns `subject`, `email` and `sms` variants;
`messages.render_fragment(key, **context)` renders a reusable fragment.
An explicitly shared variant is declared in the catalog. Ordinary notifications
prefer e-mail and fall back to SMS; authentication uses the explicitly selected
route. Rendering rejects unknown keys, missing or extra fields, and line breaks
in subjects before dispatch. Values stay plain text, including literal braces.
Rendering diagnostics contain identifiers and field names, never values or codes.

For an existing deployment, run this **read-only** preflight from the candidate
release before activation, using the actual private configuration path:

```bash
./venv/bin/python -m nuligahelper.jobs.message_config --config /path/to/private/config.json
```

The command opens no database or provider connection and sends no notifications.
It reports obsolete keys, their destinations and `prior-default`, `customized`
or `unrecognized` status. Comparison fingerprints cover the last public positional
example and notifier fallback defaults; other historical versions are treated as
customized for review. No message contents or recipient values are printed.
Privately inspect customized or unrecognized entries, map the old positional
values to the documented named fields, and reconcile intended wording into
`messages.py` through normal code review **before deployment**. Referee alerts
now explicitly place age class, game date, start time and notified names; the old
sample's shifted positional values must not be copied into the new contract.

Move `club.texts.mailRefCoordTargets` unchanged to
`club.notifications.referee_targets`; entries retain their `Name` and `Address`
shape. If both are present, the new key takes precedence, including an empty
list; conflicting metadata produces a value-free warning. When only the old key
exists it remains a recipient fallback. Remove `club.texts` once wording and
metadata have been reconciled, and retain a private configuration snapshot for
rollback together with the previous application release. Legacy text presence
alone does not reject startup, but produces key-only deprecation warnings and
never overrides the catalog. Provider/environment settings are unchanged.
No database migration is involved.

The [message inventory](../../openspec/changes/archive/2026-10-05-centralize-message-templates/message-inventory.md)
records every legacy key, channel and caller, including removed unused early-task
definitions. `send_article` remains callable with centralized newspaper content;
its daily-job call site remains disabled.

Copy review covers spelling and grammar, current task labels, greetings and
sign-offs, readable paragraphs, complete date/time information and a clear next
action. Spielfeste have no invented opponents; block meeting/delivery times are
distinct from match start times. Required welcome wording and the 15-minute code
lifetime must remain intact. Existing club-specific arrival instructions are
retained and need club/domain review before anyone changes their policy;
centralizing text introduces no new duty rules.


---
Back to the [README](../../README.md) · [Operations index](README.md)
