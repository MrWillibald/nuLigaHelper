# Proposal

## Why

Notification texts are spread across the JSON configuration and several Python modules, making wording and placeholder contracts difficult to maintain. A single developer-maintained Python source will make messages easier to review and keep their wording and dynamic values consistent.

## What Changes

- Centralize notification subjects, bodies and reusable message fragments in one dedicated Python file, covering game reminders, day-block reminders, rescheduling, referee alerts, new-game alerts, account/authentication messages, Spielfest variants and operator alerts.
- Replace implicit positional formatting with documented named placeholders and explicit message contexts.
- Clean up German wording, spelling, formatting and terminology while preserving each message's purpose, required information and existing behavioral requirements.
- Preserve distinct e-mail and SMS variants where present, and make deliberate shared variants explicit.
- Keep recipient lists, contact routes, provider credentials and operational settings separate from message text.
- Provide configuration preflight and migration guidance for existing customized texts; keep legacy configurations readable during transition without allowing legacy text settings to override the Python source.
- Preserve disabled notification paths and remove only template definitions proven unused.

## Capabilities

### New Capabilities

- `message-templates`: Consistent notification content, named placeholder contracts and migration to one developer-maintained message source.

### Modified Capabilities

- `user-accounts`: Align the agreed welcome greeting capitalization with the final message copy, preserving approval information, delivery and failure handling.

## Impact

- New `messages.py` as the source for notification text and template contracts.
- `notifier.py`, `webapp.py` and `operations.py`: render messages from the shared source while retaining dispatch responsibilities.
- `common.py` and configuration consumers: separate recipient metadata from legacy text configuration and support migration warnings.
- `config_template.json`, `test/helpers.py` and synthetic fixtures: settings-only examples and explicit recipient configuration.
- Offline notifier, authentication, configuration and operator-alert tests.
- `README.MD`, `AGENTS.md` and relevant deployment documentation: message editing, placeholder contracts and configuration migration.
- No database schema, provider configuration, external dependency or new notification trigger is required.
