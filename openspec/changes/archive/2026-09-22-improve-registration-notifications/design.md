## Context

See `proposal.md` for motivation and `specs/user-accounts/spec.md` for the behavioral contract. Registration verification currently commits the verified state and calls a web-layer helper that selects only the first active administrator. Approval commits the active state without notifying the user. `Notifier.send_account_message` already provides the required automatic e-mail-first, SMS-fallback behavior, and the web layer already wraps account delivery so provider failures are logged rather than propagated.

The notification must remain separate from the authentication-code route chosen explicitly by the user: these are automatic account notifications and therefore follow the established stored-contact preference. Notifications also run in the request that performs the state transition, so provider calls must happen only after database commit.

## Goals / Non-Goals

**Goals:**

- Reuse the existing account-message dispatcher and its contact preference.
- Attempt admin delivery independently so one missing contact or provider failure does not stop the remaining recipients.
- Ensure only a successful verified-to-active transition produces a welcome notification.
- Keep committed account state authoritative when delivery is unavailable.

**Non-Goals:**

- Add a durable notification queue, retry scheduler, or delivery ledger.
- Notify users about rejected registrations.
- Introduce a public base URL or clickable deep link in notification messages.
- Change authentication-code delivery limits or the registration state model.
- Move these texts into configuration as part of this change.

## Decisions

### Notify administrators only after successful contact verification

The verification paths will continue to commit the `verified` state before invoking the administrator notification helper. The helper will select all people with `is_admin` set and active account status in stable identifier order, rather than selecting only the first match.

This timing matches the approval queue: unverified registrations cannot be approved and should not create administrator noise. Sending immediately after initial form submission was rejected because it would expose unverified or abandoned attempts and make abuse easier.

### Deliver independently through the existing safe account-message path

Each administrator will receive a personalized e-mail body and concise SMS body through the existing safe wrapper around `send_account_message`. Iteration continues after a skipped recipient or caught provider exception. No new multi-recipient mail is introduced, which avoids exposing administrator addresses to one another and preserves per-person e-mail/SMS fallback.

The messages will include the registrant display name and stable, presentation-safe membership label. The administrative action is described as navigation to "Helfer verwalten" instead of a URL because the application has no configured externally reachable base URL.

### Commit approval before sending the welcome notification

The approval route will perform the verified-to-active transition and commit it before sending the welcome message. Rejected, missing, already active, or otherwise invalid registrations follow their existing response paths and do not reach delivery. This makes the state transition itself the idempotency boundary: a repeated request cannot send a second welcome notification because the person is no longer verified.

Approval remains successful if the user lacks a usable contact or the provider fails. This is consistent with the established best-effort notification behavior and avoids rolling back a valid administrative decision because of an external service.

### Keep account-notification copy close to the account workflow

The registration-specific German message copy will remain in the web account workflow, alongside authentication-code and existing registration-approval copy, while the notifier remains responsible for transport. Moving the text into `config.json` was considered but rejected for this change because it would expand the configuration contract and placeholder compatibility surface for two stable product messages.

The user e-mail and SMS variants will both include the exact greeting "herzlich willkommen beim nuLigaHelper des TuS Raubling Handball!", confirmation that approval succeeded, and an invitation to sign in and take open duties in the Heimspielplan.

## Risks / Trade-offs

- [Synchronous delivery to several administrators increases approval-page latency] -> Keep the expected administrator set small and isolate every attempt; a durable queue can be introduced separately if operational evidence warrants it.
- [No retry means a transient provider failure can lose an informational message] -> Log each failed or skipped delivery while preserving the committed state; operators can still see pending registrations in the management UI.
- [Two verification entry points could diverge] -> Route both the code-confirmation and legacy-link paths through the same administrator notification helper and cover both behaviorally where practical.
- [Personal data is sent to all administrators] -> Limit recipients to active administrators, who can already view pending registrations and their selected teams in the management interface.

## Migration Plan

No schema or data migration is required. Deploy the application changes normally. Rollback restores the former notification behavior without changing stored registration states.
