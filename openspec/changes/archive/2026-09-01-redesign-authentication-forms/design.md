## Context

See `proposal.md` for motivation and `specs/user-accounts/spec.md` for the changed
behavior contract. The current Flask routes render separate login/registration and
generic SMS-code pages. E-mail challenges contain only signed links, SMS challenges
contain codes, contact lookup compares literal stored strings, and the notifier always
prefers e-mail when both channels exist. Authentication responses must remain resistant
to account enumeration, tokens must remain short-lived and single-use, and the complete
test suite must stay offline.

The shared stylesheet was originally written for the dark person-management form. Its
global form-label rules and incomplete generic card styling are the immediate source of
the visual mismatch. The redesigned components therefore need narrowly scoped auth
styles instead of another extension of those global selectors.

## Goals / Non-Goals

**Goals:**

- Use one server-rendered page for both steps of each authentication flow, with optional
  JavaScript limited to immediate presentation changes.
- Give both delivery channels the same code-entry model while retaining atomic token
  consumption, expiry, CSRF checks, rate limits and enumeration-safe output.
- Canonicalize contact data at every write and lookup boundary.
- Make the auth cards responsive, accessible and visually consistent with the existing
  navy, neutral and orange design tokens.
- Keep deployment compatible with the project's no-migration policy.

**Non-Goals:**

- Password authentication, third-party identity providers or remembered devices.
- Changing registration approval, access tiers, session lifetime or the notification
  channel preference used for ordinary game notifications.
- Checking whether an e-mail mailbox is deliverable or whether a phone number currently
  belongs to the person; the code proves control after syntax validation.
- Introducing a general component framework or redesigning non-authentication pages.

## Decisions

### D1 - Treat both pages as server-rendered two-action state machines

Each page has `request_code` and `confirm_code` actions handled by its existing route.
The initial GET renders step one and a visibly secondary code-request button. A valid
request POST renders the same template in a challenge state with step two enabled, a
generic delivery message and the primary completion button. An "Angaben ändern" action
returns to the initial state. Registration step one includes name, team and consent;
login step one includes only the contact controls.

JavaScript changes the E-Mail/SMS input presentation immediately and can focus the code
field, but it does not own the workflow. Server-rendered POST responses remain complete
and usable without JavaScript.

*Alternative considered:* JSON endpoints and a client-only wizard. Rejected because it
duplicates form error handling, creates a second CSRF surface and makes a small Flask UI
dependent on JavaScript for authentication.

### D2 - Issue one numeric challenge model for both delivery channels

Every new registration or login challenge stores a random nonce and a six-digit code.
The page receives a signed opaque challenge value containing the nonce and purpose; it
does not receive a person ID. Confirmation validates the signature, purpose, code,
expiry and unused state, then marks the token used with the existing compare-and-swap
update. A challenge for an unknown or otherwise ineligible contact renders the same
visible state but cannot authenticate anyone.

Delivery is explicit for account messages in this flow: E-Mail sends the code by mail
and SMS sends it through Twilio, even if the matched person has both addresses. This
does not change `Notifier`'s mail-first fallback for game and helper notifications.

The separate code-entry URLs redirect to their corresponding unified page. The signed
link consumers remain temporarily capable of consuming already-issued pre-deployment
link tokens during their 15-minute lifetime, but newly issued challenges use codes only.

*Alternative considered:* include both a link and a code in e-mail. Rejected because it
keeps two completion paths, undermines the requested consistent interaction, and leaves
more behavior and tests than the club needs.

### D3 - Centralize offline-safe contact validation and normalization

Add a small contact-validation module used by registration, authentication lookup,
admin person creation/editing and member self-editing. Use `email-validator` with DNS
deliverability checks disabled so validation is deterministic and offline; store and
compare the library's normalized address. Use `phonenumbers` to parse the chosen country
calling code plus national input, require a possible and valid number, and store E.164.

The phone control defaults to `+49 Deutschland`, offers the club's common neighboring
calling codes prominently, and provides an "Andere Ländervorwahl" choice so every valid
international number remains representable. Input normalization accepts customary
spaces, punctuation and a domestic trunk prefix where the parser can interpret it. A
full international number entered into the national field must agree with the selected
calling code rather than silently changing the selection.

Invalid syntax produces a field-specific German error before a token is issued or data
is changed. A syntactically valid but unknown login contact proceeds to the same generic
challenge state as a known contact, preserving enumeration resistance. Database unique
constraints remain the final defense against races after canonicalization.

*Alternatives considered:* HTML validation only is bypassable and cannot canonicalize;
hand-written regular expressions cannot reliably validate international phone plans;
online mailbox or carrier checks would make authentication and tests depend on external
services.

### D4 - Preserve challenge context without trusting editable fields

After step one, the server displays a masked destination and includes only the signed
opaque challenge in the confirmation submission. The original registration data is
already held in the pending registration record, and the token nonce identifies the
challenge atomically. Changing a contact requires returning to step one and requesting a
new challenge. Unknown-contact requests receive a same-shaped signed dummy challenge so
markup and timing do not become a useful membership signal.

Unverified registrants may request a replacement code for the same canonical contact;
the newest valid challenge can complete the existing pending registration without
creating a duplicate. Existing active, inactive or rejected contacts do not gain a new
registration record.

*Alternative considered:* repost the contact and registration fields with the code.
Rejected because read-only and hidden browser fields are still attacker-controlled and
would force the server to reconstruct which step-one values were originally accepted.

### D5 - Scope a reusable authentication visual language

Create auth-specific card, field, help-text, radio-group, phone-group, step, challenge
and button classes. The card is centered with a readable maximum width inside the
existing `.section`; inputs and selects share border, focus and error states. Contact
choices use native radio inputs inside a `fieldset`/`legend`. The code remains one input
with `inputmode="numeric"`, `autocomplete="one-time-code"`, a six-digit pattern and
visual letter spacing, which supports paste and assistive technology better than six
independent inputs.

The request button uses a neutral/light-navy secondary style. The final action uses the
existing orange primary color. Disabled styling and step labels reinforce order without
depending on color alone. All field labels and concise descriptions are German UI text.

*Alternative considered:* reuse `.person-form-card` directly. Rejected because that
dark, dense admin card is optimized for a sidebar and its selectors caused the current
auth-page inconsistencies.

### D6 - Keep redirects and account gates explicit

A valid code for an active account clears the old session, establishes the sliding
session and redirects to the Heimspielplan. A verified but unapproved registration
continues to redirect to its status page. Invalid, expired or replayed codes stay on the
same form with a generic code error and do not disclose account state.

## Risks / Trade-offs

- [New normalization makes differently formatted existing contacts collide] -> Before
  applying canonical values, report collisions and leave both records unchanged for an
  admin to resolve; no automatic destructive merge is introduced.
- [Code-only e-mail is slightly less convenient than clicking a link] -> Use browser
  one-time-code autocomplete, paste-friendly input and a concise mail template.
- [Six-digit codes have limited entropy] -> Retain the 15-minute expiry, per-person and
  per-client rate limits, single-use consumption and a signed challenge nonce.
- [A country selector can become unwieldy] -> Prioritize common prefixes and provide an
  explicit custom-prefix option while the server library remains authoritative.
- [Dummy and real challenges can differ internally] -> Test visible content and response
  shape for known and unknown contacts, and avoid identifiers or account-specific text.
- [New validation dependencies increase installation size] -> Pin maintained libraries,
  keep deliverability checks offline and cover their boundary behavior with synthetic tests.

## Migration Plan

1. Add validation dependencies and shared normalization helpers with focused unit tests.
2. Audit synthetic/test fixtures and, before real deployment, run a read-only report of
   existing contacts that are invalid, change under normalization or collide.
3. Update token issuance and explicit-channel delivery while retaining temporary
   consumption of already-issued 15-minute e-mail links.
4. Replace the auth templates and scoped styles, then redirect legacy code-entry pages.
5. Run the full offline suite and manually verify both routes at desktop and narrow
   viewport widths.
6. Deploy after resolving any reported real-data contact collisions. No schema migration
   is required; if implementation unexpectedly changes the schema, stop and ask the owner
   before adding migration machinery.

Rollback consists of restoring the previous application version and dependencies. Old
link-token support remains readable during the short transition; code challenges issued
by the new version may require users to request a fresh login after rollback.
