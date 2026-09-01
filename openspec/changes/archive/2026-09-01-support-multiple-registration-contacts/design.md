## Context

Registration currently validates only the selected route, hides the other contact field,
and persists exactly one contact. The existing person schema already permits nullable,
unique e-mail and phone columns, and login lookup and ordinary notification dispatch are
already route-aware.

## Goals / Non-Goals

**Goals:**

- Collect and canonicalize both contacts during registration when both are supplied.
- Make route availability understandable and prevent invalid contacts from being used.
- Preserve explicit route selection for authentication while retaining mail-first automatic
  notification behavior.
- Keep the existing verification, privacy, uniqueness, rate-limit, and approval boundaries.

**Non-Goals:**

- No database migration or new contact fields.
- No automatic verification of both contacts; only the selected route proves control during
  registration.
- No change to the existing automatic notification message content or fallback semantics.
- No change to how an already registered person adds or edits contacts.

## Decisions

### Validate both fields independently on the server

Registration will normalize e-mail and phone separately, using the submitted calling-code
context for the phone. An empty field is absent; a non-empty invalid field is an error.
At least one normalized contact and a selected normalized route are required. This is
preferred over validating only the selected route because otherwise a supplied contact
could be silently discarded.

### Persist both contacts before verification

The pending `Person` stores every valid supplied contact immediately, while the selected
route is used only to issue the verification challenge. This keeps contact uniqueness
enforced by the existing database constraints and makes both routes available after
verification without a second data-entry step.

### Keep authentication route selection explicit

Login continues to look up the submitted canonical e-mail or phone independently and sends
the code through that same route. It must not inherit the automatic mail-first preference,
because a user with both routes explicitly controls the authentication destination.

### Put both contact panels before the route selector

The registration template will render e-mail and SMS inputs unconditionally, followed by
the route radios. Progressive enhancement will update radio availability from field
validity and keep the route choice usable without JavaScript through server validation.
Invalid non-empty fields remain blocking errors even when the other route is valid.

### Check all supplied contacts for uniqueness without revealing ownership

Before creating a registration, canonicalized e-mail and phone values will both be checked
for conflicts. A conflict on either value prevents creation or modification and preserves
the existing generic registration response. This avoids attaching a second contact to an
existing account or leaking which of two submitted contacts matched.

## Risks / Trade-offs

- [A visitor enters one valid and one invalid contact] -> Block the submission and show the
  field-specific error until the invalid value is corrected or cleared.
- [A submitted e-mail and phone belong to different existing people] -> Create no record,
  send no route-specific account message, and retain the generic response.
- [Client-side validity differs from server validation] -> Treat browser behavior only as
  an affordance; repeat all validation and route checks on the server.
- [Existing pending registrations contain one contact] -> Leave them unchanged; the new
  behavior applies to subsequent registration submissions and does not require migration.

## Migration Plan

No schema migration is required. Deploy the application and templates together. Existing
one-contact records remain valid, and existing login and automatic-notification behavior
continues to work for them. Rollback is application-only; newly created two-contact records
remain compatible with the nullable contact columns.
