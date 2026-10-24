# Design

## Context

See `proposal.md` for motivation and the two delta specs for the behavior contracts. This design is required because the feature crosses scheduling, recipient selection and the validated message catalog.

`main._send_notifications()` already derives `next_week` from its effective-today argument and calls the weekly game and day-block senders. `Notifier.notify_pre()` reminds existing game assignees. `db.staffing_status()` separates physical `vacancies` from eligibility `deficiencies`; `db.missing_slots()` returns required vacancy counts by semantic role. `Team.persons` exposes memberships including non-active accounts, so dispatch must filter account state explicitly. `messages.py` owns all wording and requires explicit named contexts. Existing delivery has no persistent per-recipient weekly delivery ledger.

## Goals / Non-Goals

**Goals:**

- Add a separate, testable recruitment sender without changing the contract of assigned-helper reminder methods.
- Resolve vacancy counts and recipients from saved game/team/person identities, sharing existing domain and transport helpers.
- Capture reviewable German copy and clear ordinary-game/Spielfest message contracts before implementation.

**Non-Goals:**

- No new delivery queue, schema, URL configuration or provider integration.
- No database writer or lock transaction held during invitation dispatch, and no assignment or request mutations from this sender.
- No reimplementation of age eligibility or required-position classification inside notification code.

## Decisions

### 1. Add a distinct sender in the existing seven-day branch

Introduce `Notifier.notify_team_vacancies_early(date)` and call it from the `next_week` branch in `main._send_notifications()`, adding its sent count to that branch's existing total. Use the notifier's season when selecting games for this new flow. The scheduler alone selects the date; tests can call the sender with a synthetic explicit date. The production date remains based on `common.effective_today()` through the existing main-job argument.

Keep `notify_pre()` focused on reminders for saved assignments. This avoids changing its callers and sent-count expectations. An alternative would append recruitment to `notify_pre()`, but that would blur the difference between reminding an assignee and inviting an unassigned member.

Use exactly seven calendar days, without missed-run catch-up or daily follow-up. Dispatch remains per game per invocation: two qualifying games can generate two messages for the same person. A manual rerun on the same day can resend messages, just as existing weekly reminders can. Durable deduplication would require a separate delivery lifecycle and is outside this change.

### 2. Use physical vacancies and saved memberships

For each selected game, obtain `db.missing_slots(game)` (equivalently `staffing_status(game)["vacancies"]`). Skip empty vacancy maps and games without `game.team`. Do not use `staffing_status(game)["complete"]` as the trigger because it includes age deficiencies.

Construct an exclusion set from **all** `game.assignments` using `person_id`, including retained removed roles and open replacement searches. Traverse `game.team.persons`, retain only `ACCOUNT_ACTIVE`, remove excluded identities and deduplicate by `Person.id`. Include the MV only through this membership traversal. Do not select the playing team via `ak`, union other memberships, or add Supporter/admin recipients unless they are members of this responsible team. Use deterministic person ordering for reproducible output.

Other game and day-block assignments do not enter the exclusion set. Do not prefilter recipients using birth dates or eligibility: the invitation asks them to choose a suitable task, and actual claims keep authoritative domain validation. Compared with age-filtered candidate selection, this preserves the agreed recipient rule and avoids turning invitation delivery into a new eligibility service.

### 3. Render counts and event variants centrally

Add `team.game.vacancies.weekly` and `team.spielfest.vacancies.weekly` entries with explicit e-mail and SMS variants. Use shared named context fields `recipient_name`, `responsible_team_name`, `game_date`, `game_start_time`, `age_class` and `open_tasks_text`; add `home_team_name` and `away_team_name` only for the ordinary-game entry. Extend field documentation so `responsible_team_name` covers team invitations as well as MV follow-up.

Build `open_tasks_text` in required-position display order from catalog fragments taking `task_label` and `vacancy_count`, with correct singular/plural wording: `Verkauf: 1 Platz` / `Verkauf: 2 Plätze`. Join rows with a catalog-owned separator such as `; ` so the same summary works on a separate e-mail line and inline in SMS. Count missing positions, not whole-role capacity; do not enumerate the two Verkauf positions as different tasks. No contact or birth-date values enter the rendering context.

The alternative of composing sentences directly in the notifier would violate the authoritative catalog rule. Selecting the Spielfest key with `db.is_spielfest(game)` follows existing reminders and avoids invented matchups. Missing saved start time remains an empty display value under current notification conventions; it does not become a new dispatch guard.

### 4. Use the agreed copy with a Spielfest-specific opening

Ordinary-game e-mail subject: **Deine Unterstützung beim Heimspiel am {game_date}**

```text
Hallo {recipient_name},

nächste Woche, am {game_date}, ist deine Mannschaft {responsible_team_name} für die Dienste beim Spiel {home_team_name} gegen {away_team_name} ({age_class}) verantwortlich. Das Spiel beginnt um {game_start_time}.

Wir möchten dich freundlich daran erinnern, dass für dieses Spiel noch folgende Pflichtaufgaben zu besetzen sind:

{open_tasks_text}

Vielleicht kannst du einen der offenen Dienste übernehmen? Bitte schau im Heimspielplan nach und trage dich gerne für eine passende Aufgabe ein. Mit deiner Unterstützung trägst du dazu bei, dass unser Heimspieltag gut gelingt.

Vielen Dank für deine Unterstützung!

Viele Grüße
TuS Raubling Handball
```

Ordinary-game SMS:

```text
Hallo {recipient_name}, nächste Woche ({game_date}, {game_start_time}) übernimmt deine Mannschaft {responsible_team_name} die Dienste beim Spiel der {age_class}. Noch offen: {open_tasks_text}. Kannst du uns unterstützen? Bitte schau im Heimspielplan nach und übernimm gerne einen passenden Dienst. Vielen Dank! VG TuS Handball
```

Spielfest e-mail subject: **Deine Unterstützung beim Spielfest am {game_date}**

Use the same e-mail structure and closing, replacing the opening with:

```text
nächste Woche, am {game_date}, ist deine Mannschaft {responsible_team_name} für die Dienste beim Spielfest {age_class} verantwortlich. Das Spielfest beginnt um {game_start_time}.
```

Replace `für dieses Spiel` with `für dieses Spielfest` in the reminder sentence. Spielfest SMS:

```text
Hallo {recipient_name}, nächste Woche ({game_date}, {game_start_time}) übernimmt deine Mannschaft {responsible_team_name} die Dienste beim Spielfest {age_class}. Noch offen: {open_tasks_text}. Kannst du uns unterstützen? Bitte schau im Heimspielplan nach und übernimm gerne einen passenden Dienst. Vielen Dank! VG TuS Handball
```

### 5. Reuse ordinary delivery and preserve failure behavior

Pass each rendered invitation through `_person_receiver()` and `_dispatch_message()`. E-mail preference, phone fallback, debug suppression, skipped-contact warnings and dispatch counting remain shared with existing ordinary notifications. An SMTP failure does not newly trigger SMS fallback. Use the existing daily-job failure reporting rather than adding automatic retries or silently treating provider errors as success. Each recipient gets an individual message with no roster-wide address header.

## Risks / Trade-offs

- [A missed daily run misses the seven-day invitation; a rerun can resend it] -> Keep the existing reminder execution model explicit in README documentation; do not claim exactly-once delivery.
- [Several games can produce several invitations] -> Keep each message specific to its game and vacancy list; a digest is outside this scope.
- [Assignments or team memberships can change during network dispatch] -> Derive the cohort from the saved state read for this invocation and describe vacancies as a reminder to check the current Heimspielplan. Do not hold database write serialization across provider I/O or imply reservation of a listed slot.
- [Some recipients cannot claim all listed tasks] -> Invite a suitable duty and retain existing claim eligibility and binding confirmation; do not expose ages or dates of birth.
- [Long team names and vacancy lists can require multiple SMS segments] -> Keep SMS wording shorter than e-mail while retaining agreed information; do not silently truncate task counts or event identity.

## Migration Plan

1. Implement the message catalog, sender, scheduler integration and offline tests described in `tasks.md`.
2. Update README notification behavior and run the full offline suite with synthetic configuration and providers.
3. Deploy through the existing application update process. No database migration or configuration conversion is needed.
4. Roll back to the prior application version if necessary; there is no new persisted feature state to undo, and already delivered messages cannot be recalled.
