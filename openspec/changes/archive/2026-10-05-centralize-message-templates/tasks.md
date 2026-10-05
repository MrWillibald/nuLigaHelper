# Tasks

## 1. Inventory and Shared Message Source

- [x] 1.1 Inventory all notification subjects, bodies, optional fragments, active callers and dormant definitions across the JSON example, notifier, web account flows and operator alerts; verify every legacy text key has a documented destination or justified unused disposition without reading real configuration in tests.
- [x] 1.2 Add `messages.py` with stable message keys, channel variants and named placeholder contracts; verify synthetic examples for every retained template render all required values in the correct positions.
- [x] 1.3 Add validated plain-text rendering and focused offline coverage for missing fields, unknown keys, literal braces in values and invalid subject line breaks; verify invalid contexts cannot produce a dispatchable message and diagnostics omit message values.
- [x] 1.4 Document the catalog, channel selection, named contracts and developer editing workflow in `README.MD` and update the positional-template guidance in `AGENTS.md`; verify the documentation identifies one authoritative Python editing point.

## 2. Configuration Transition

- [x] 2.1 Add configuration preflight that reports legacy text keys and customized-template status without exposing template contents or contact values; verify it is non-mutating and works with synthetic prior-default and customized configurations.
- [x] 2.2 Separate referee recipients into `club.notifications.referee_targets`, retain the old metadata key as a fallback during transition, and keep provider settings unchanged; verify settings-only, legacy-only and conflicting-recipient configurations preserve the documented precedence.
- [x] 2.3 Keep legacy `club.texts` readable with key-only deprecation warnings while preventing text values from overriding the shared catalog; verify legacy text presence alone does not reject startup and no legacy text changes rendered output.
- [x] 2.4 Update `config_template.json`, synthetic configuration helpers, fixtures and deployment/setup documentation for settings-only configuration and customization migration; verify documented preflight and migration steps require no real notification send and tests never read private `config.json`.

## 3. Game and Day-Block Notifications

- [x] 3.1 Convert game reminders, weekly reminders, MV messages, preparation/cleanup reminders, rescheduling, referee alerts and new-game alerts to named contexts; verify offline recorder tests preserve recipient sets, channel preference, reminder timing and dispatch/skip counts.
- [x] 3.2 Extract Spielfest variants, timing/partner fragments and notification subjects from call sites; verify ordinary games and Spielfest render the intended information without invented opponents or inference from template text equality.
- [x] 3.3 Reconcile the referee alert's current positional mismatch using explicit age-class, date, time and notified-person fields; verify distinguishable synthetic values appear in their intended places for each supported channel.
- [x] 3.4 Apply the German copy-review checklist to these templates and document any existing instructions that require domain review; verify rendered examples use consistent current task labels and distinguish event start from a block meeting time.

## 4. Account, Operator and Dormant Messages

- [x] 4.1 Convert authentication codes, existing-account notices, administrator approval requests and approved-user welcomes to the shared source; verify explicit authentication-route selection, automatic channel preference, anti-enumeration responses and post-commit best-effort delivery remain unchanged.
- [x] 4.2 Verify the welcome template retains the exact greeting required by `user-accounts` and the administrator message still identifies the registrant and all selected teams; review rendered e-mail/SMS examples for clear actions and unchanged required content.
- [x] 4.3 Convert operator-alert subject/body without routing alerts through member notifications; verify synthetic operator SMTP tests preserve configured recipients, affected components, occurrence time and runbook reference.
- [x] 4.4 Centralize retained newspaper content and schedule fragments while leaving its daily-job call site disabled; remove early-task definitions only when an inventory search confirms no supported consumers, and verify extraction adds no send trigger.
- [x] 4.5 Review account, operator and dormant copy against the same documented checklist and update relevant message examples; verify descriptions of authentication lifetime and operational actions still match existing specifications.

## 5. Integration Verification

- [x] 5.1 Run the repository offline test suite and strict OpenSpec validation for this change; verify both pass and committed debug switches remain disabled.
- [x] 5.2 Audit notification call sites and the committed configuration example for remaining duplicated message definitions or text overrides; verify every retained message resolves through the shared catalog and recipient/provider settings remain separate.
