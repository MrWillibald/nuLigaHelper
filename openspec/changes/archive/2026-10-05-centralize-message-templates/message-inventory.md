# Message inventory and copy review

All renderable wording lives in `messages.py`. `message_config.py` holds only
migration destinations and comparison fingerprints, never renderable old copy.
Every catalog entry records its purpose, caller, supported channels, declared
fields and active/dormant status. `FIELD_DESCRIPTIONS` documents each field's
meaning. E-mail and SMS are distinct unless `shared_sms=True` explicitly states
that both channels share the same complete body. E-mail-only messages declare
no SMS variant.

## Legacy positional definitions

| Legacy keys | Catalog destination | Channels / caller | Disposition |
| --- | --- | --- | --- |
| `mailErrorSubject`, `mailError` | `game.new` | E-mail / `notify_new_games` | Retained admin new-game alert |
| `mailTask`, `textTask` | `game.day_before` | E-mail + SMS / `notify_game_day` | Named recipient, task, date, age class, opponents and start time |
| `mailPreTask`, `textPreTask` | `game.weekly` | E-mail + SMS / `notify_pre` | Same semantic context with explicit weekly key |
| `mailMVSubject`, `mailMV`, `textMV` | `mv.game.day_before` | E-mail + SMS / `notify_game_day` | Named responsible team, timekeeper, secretary and event values; eligibility suffix remains derived from domain state |
| `mailPreparationTask`, `textPreparationTask` | `block.preparation.weekly` | E-mail + SMS / `notify_service_early` | Named block date, task, other volunteers and meeting time; original sale sender credentials preserved |
| `mailBlockPreTask`, `textBlockPreTask` | `block.weekly` | E-mail + SMS / `_notify_blocks_early` | Cleanup/cake duty time remains distinct from event start |
| `mailBlockTask`, `textBlockTask` | `block.day_before` | E-mail + SMS / `notify_blocks_day_before` | Day-before duty time; one cake per assigned cake position |
| `mailShifted`, `textShifted` | `game.shifted` | E-mail + SMS / `notify_shifts` | Named previous/new event dates and start times |
| `mailRefCoordSubject`, `mailRefCoord` | `referee.missing` | Explicit shared body / `_notify_missing_referee` | Five named values reconcile the old four-placeholder mismatch: recipient, age class, event date, start time, notified persons |
| `mailNewspaperSubject`, `mailNewspaper` | `newspaper.article` | E-mail / `send_article` | Supported callable, dormant in daily job; publication date, weekday, event date and schedule |
| `mailEarlyTask`, `textEarlyTask` | None | No supported caller | Removed: source inventory found only definitions in the example and planning references; active early preparation uses `block.preparation.weekly` |
| `mailRefCoordTargets` | `club.notifications.referee_targets` | Recipient settings / `common.referee_targets` | Metadata, never message wording; legacy fallback only when new key is absent |

## Former caller-local content

| Catalog keys | Caller / status | Semantic contract |
| --- | --- | --- |
| `spielfest.day_before`, `spielfest.weekly` | Helper reminders / active | Recipient, task, event date, full SPF age class, start time; no opponents |
| `mv.spielfest.day_before` | MV follow-up / active | Responsible team, saved timing volunteers, SPF date/start and eligibility suffix |
| `spielfest.shifted` | `notify_shifts` / active | SPF identity and previous/new event timing; no opponents |
| `staffing.deficiencies` | `notify_game_day` / active | Domain explanations without full birth dates or exact ages |
| `staffing.position_unassigned` | `notify_game_day` / active | Central vacancy fallback for explicitly labeled timing positions; no empty-name sentences |
| `block.time_unset`, `block.time_cross_date` | `_block_time_text` / active | Explicit unset time or a task time falling on a different calendar date |
| `preparation.partner_none` | `notify_service_early` / active | Existing no-partner fallback |
| `task.cake` | `_block_reminder_task` / active | Current display task label plus one-cake quantity |
| `auth.registration_code`, `auth.login_code` | `_send_challenge` / active | Recipient and six-digit code; purpose and 15-minute lifetime explicit in each key |
| `auth.existing_account` | Registration flow / active | Existing-account notice and invitation to use login; selected route preserved |
| `account.approval_request` | `_notify_registration_approvers` / active | Administrator name, registrant name, every selected team via the domain membership display label; management action |
| `account.welcome` | `_notify_registration_approved` / active | Exact required welcome greeting, approval confirmation and invitation to take open duties |
| `operations.alert` | `operations.send_alert` / active | Affected components, ISO UTC occurrence time and existing runbook reference; separate operator SMTP recipients |
| `newspaper.game_row`, `newspaper.minis_row`, `newspaper.e_youth_row` | `send_article` / dormant | Ordinary row's start, age-class display name, opponents; retained MI/GE tournament summaries |
| `newspaper.women`, `newspaper.men` | Newspaper helpers / dormant | Existing display fragments; weekdays are supplied event context; no new send trigger |

The former `_spielfest_task_text` optional partner branch had no caller supplying
a partner. Active preparation partner information remains in
`block.preparation.weekly`. Weekly/day-before selection is explicit and never
inferred by comparing template bodies. Notification subjects are part of each
catalog entry. Domain-provided labels and event facts remain caller context.

## Copy review checklist and retained domain instructions

- Check German spelling, grammar, greetings, sign-offs and readable paragraphs.
- Use current semantic task labels; distinguish event start, block meeting and
  cake delivery. Include a calendar date when a calculated block time crosses
  midnight. Preserve each message's identifying data and next action.
- Retain channel-specific information and deliberate shared variants. Preserve
  e-mail preference, selected authentication route, skip/dispatch counts and
  best-effort account delivery after committed transitions.
- Preserve the exact greeting `herzlich Willkommen beim nuLigaHelper des TuS
  Raubling Handball!`, approval confirmation and open-duty invitation. Both
  authentication codes remain valid for 15 minutes.
- Spielfeste identify the full SPF event and timing without invented opponents.
  Referee copy now identifies each distinct semantic value in its proper place.
- Existing club-specific arrival instructions are retained. Any policy change
  requires club/domain review of the arrival instructions.
- Dormant newspaper addressee, newspaper, venue and author retain the reconciled
  club-specific wording and require club review before enabling the feature. Neither
  template extraction nor the inventory changes the disabled daily-job call.

The [synthetic rendering tests](../../../../test/test_messages.py) exercise every
catalog entry and all supported channels. Reviewed club-specific wording stays
in the authoritative catalog; configuration contents are not copied into review
artifacts. Recorder and account/operator integration tests verify semantic placement,
current labels, shared-source edits, timing, routing and privacy. Configuration
preflight itself prints only key names and status; developers privately reconcile
custom wording using the migration instructions in `README.MD`.
