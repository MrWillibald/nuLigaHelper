# Design

## Context

See `proposal.md` for motivation and `specs/message-templates/spec.md` for the behavior contract.

The current message inventory spans several sources:

| Source | Message families and relevant context |
| --- | --- |
| `config_template.json` / legacy `club.texts` | Ordinary game reminders, weekly reminders, responsible-team MV notifications, preparation and cleanup reminders, rescheduling, referee alerts, new-game alerts, newspaper content and unused early-task definitions |
| `notifier.py` | Default preparation/block templates, notification subjects, Spielfest helper/MV/rescheduling variants, timing and optional partner fragments, and newspaper schedule fragments |
| `webapp.py` | Registration/login codes, existing-account notifications, administrator approval requests and approved-user welcome messages |
| `operations.py` | Operator-alert subject and body |
| Legacy text configuration metadata | `mailRefCoordTargets` contains recipient configuration rather than message text |

The repository requires German user-facing messages and English code/comments. Current positional `str.format` order and argument counts are part of the effective contract and must be mapped explicitly during conversion. In particular, the sample referee template contains four positional placeholders while its caller supplies five semantic values; conversion must reconcile the intended content instead of carrying forward shifted values.

Existing account specifications require the exact welcome greeting `herzlich Willkommen beim nuLigaHelper des TuS Raubling Handball!`, automatic e-mail preference with SMS fallback, explicitly selected routes for authentication, and best-effort informational delivery after committed account transitions. The centralized source must preserve those requirements.

## Goals / Non-Goals

**Goals:**

- Give developers one reviewable source for complete notification content and the context required to render it.
- Make placeholders readable and verifiable rather than dependent on argument order.
- Make the configuration transition explicit without silently losing recipient settings or customized wording.
- Keep rendering independent of transport, database access and notification triggering.

**Non-Goals:**

- Introduce runtime text editing, localization infrastructure or additional languages.
- Centralize every page label, validation error, log entry or general interface string.
- Change notification timing, recipient eligibility, delivery preference, retry behavior or provider settings.
- Enable the newspaper path or add new messages because a template exists.
- Invent new arrival instructions, business policies or final club-specific wording.

## Decisions

### Use one developer-maintained Python module

Create `messages.py` containing stable message keys, subjects, e-mail/SMS bodies, optional content fragments and named placeholder contracts. Callers provide context and request the appropriate message variant; message text is not duplicated at call sites or loaded from JSON.

The user selected Python because editing the wording requires developer knowledge of placeholders. A separate JSON catalog would retain the same maintenance problem without making the contracts easier to inspect. Multiple Python files per workflow would also weaken the requested single editing point.

The module performs no provider calls, configuration loading, database queries or notification scheduling. Settings and recipient metadata remain in their existing configuration mechanisms.

### Inventory and convert semantic values before editing wording

For every retained template, document its message purpose, active or dormant status, channels, named fields and caller. Map legacy positional arguments to semantic names before changing the copy.

Representative fields include `recipient_name`, `task_label`, `game_date`, `start_time`, `age_class`, `home_team`, `away_team`, `responsible_team`, `timekeeper_name`, `secretary_name`, `partner_names`, `old_date`, `old_time`, `new_date`, `new_time`, `registrant_name`, `team_names`, `auth_code`, `action`, `expiry_minutes`, `component_names`, `occurred_at`, `runbook_reference`, `article_date`, `weekday` and `schedule_text`. Each message declares only its own required fields.

Use explicit message keys for day-before versus weekly reminders and ordinary games versus Spielfest. Do not infer message timing by comparing template text. Where two channels intentionally share wording, declare that relationship in the catalog; preserve concise SMS variants wherever they already exist or the cleanup introduces a reviewed shorter variant.

The referee alert must place age class, event date/time and notified-person names in their intended positions. Distinct synthetic values must demonstrate that the migration fixes the existing mismatch.

### Render validated plain-text contexts

Use a small rendering entry point associated with the catalog. Validate that each declared placeholder exists and is supplied before dispatch. Unknown message keys or incomplete contexts produce a clear rendering error rather than sending a partial or incorrectly formatted message.

Use named plain-text substitution without expression evaluation or recursive interpretation of values. Braces in a person or team name remain literal content. Keep subject templates free of line breaks and prevent supplied subject values from creating additional mail headers. Existing transport boundaries continue to own mail/SMS construction and delivery.

Keep diagnostic information limited to message keys and field names. Do not log rendered authentication messages, codes, contact data or configuration contents.

### Review copy against a documented checklist

For each active message, review spelling, grammar, task terminology, date/time presentation, greetings, sign-offs, paragraph structure, optional fields and the recipient's next action. Keep the distinction between match start and block meeting/delivery time explicit.

Supply labels and event facts from domain context so messages follow renamed tasks without embedding old labels at call sites. Preserve required account wording, authentication-code lifetime and all existing information needed to identify the duty or event. Spielfest messages must retain the event identity without inventing opponents.

The proposal does not prescribe new final sentences or club-specific arrival policies. Copy cleanup is performed against this checklist during implementation, with the resulting text reviewed together with its rendered examples.

### Separate configuration with explicit compatibility

Move referee recipient metadata to `club.notifications.referee_targets`, preserving the existing recipient entry shape and values. During transition, read `club.texts.mailRefCoordTargets` only as a legacy recipient fallback when the new setting is absent. When both exist, the new setting wins and conflicting metadata is reported without printing contact values.

Render message text exclusively from `messages.py`. A readable legacy `club.texts` section must not overwrite module templates or cause startup rejection solely because it exists. Emit a deprecation warning identifying obsolete setting keys, without printing their values.

Provide a preflight inventory that reports legacy text keys and identifies customized templates against the applicable prior defaults using names/status rather than full message content. Document how developers privately inspect those customizations and carry intended wording and placeholder semantics into the Python source before deployment. The preflight reports work to resolve; it does not automatically rewrite private configuration or source code.

Remove legacy text entries and move recipient metadata as a deployment preparation step. Compatibility handling allows an existing configuration to remain readable during transition, while the final documented format contains settings only. Provider settings and credentials keep their current environment-based contract.

### Keep dormant features dormant

Retain and centralize newspaper templates and schedule fragments while `send_article` remains a supported callable path; preserve the disabled call site in `main.py`. Remove `mailEarlyTask` and `textEarlyTask` only after repository-wide inspection confirms they have no supported consumers. Record the disposition of every legacy key so omissions are deliberate.

A template's presence must never create a new trigger. General log messages and unrelated page labels remain outside the inventory.

## Risks / Trade-offs

- [Customized private text differs from committed examples] → Run preflight and privately reconcile customizations before deployment; never print private configuration into review artifacts.
- [Named-field conversion puts valid data in the wrong place] → Map each positional contract and render templates with distinguishable synthetic values.
- [Centralization changes message timing or delivery paths] → Keep orchestration and transport at existing boundaries and cover resulting sends with offline recorders.
- [Copy changes contradict required account wording or task policy] → Preserve existing account requirements and review contextual instructions against authoritative domain behavior.
- [Legacy text entries remain present and appear authoritative] → Emit key-only deprecation warnings and document that only the Python source controls rendered messages.
- [A dormant feature becomes active during cleanup] → Assert that extraction does not add triggers or enable the newspaper call site.

## Migration Plan

1. Inventory committed templates and callers, document named contracts, and add synthetic rendered examples.
2. Run configuration preflight on the deployment's existing configuration without exposing its contents; reconcile intended customizations into the module through normal developer review.
3. Move referee recipient metadata to the new settings location and remove legacy text entries in the prepared configuration.
4. Deploy the module and converted callers together, with legacy configuration readability and recipient fallback available during transition.
5. Verify offline delivery recordings, authentication behavior and operator-alert content before production use; no migration step sends real member notifications.
6. Rollback restores the previous application release and its compatible configuration snapshot together. No database migration is required.
