## Why

The web interface is completely open: anyone who can reach port 8080 can assign,
reassign and delete helpers, and every page exposes the full roster. The club wants to
put the interface on the public internet, which makes an open write surface untenable
and turns "who changed this assignment?" into a question nobody can answer today.

At the same time, staffing home games is currently a pull model: the admin chases
helpers, and the MV mail nags while slots stay empty. Letting helpers sign up for tasks
themselves attacks that problem at the source — but self-service only works if the
person clicking is known to be who they claim to be.

## What Changes

- **Passwordless login.** A person proves control of the e-mail address or phone number
  already stored on their roster entry: magic link by mail, one-time code by SMS. No
  passwords are stored anywhere. Reuses the existing "mail preferred, phone fallback"
  dispatch rule from `notifier.py`.
- **Self-signup with an approval gate.** Newcomers register with name, contact channel
  and desired team. Verifying the channel creates an account that can log in; the MV of
  the desired team — or an admin — must approve the registration before the person
  becomes part of the roster and assignable. Verification proves a channel, approval
  proves club membership.
- **Four access tiers** — guest, member, MV, admin — replacing today's single implicit
  "everyone is an admin" tier:
  - *guest* (not logged in): read-only schedule including helper names. `/personen` and
    `/statistik` are not reachable.
  - *member*: sees the roster without contact data, sees and edits only their own name
    and contact data, and claims or releases tasks **for themselves only**.
  - *MV*: additionally assigns and unassigns members of their own team, and approves
    registrations for their own team.
  - *admin*: everything, including approving registrations for any team, creating
    contactless roster entries, setting the responsible team and the team MV, and
    reading the audit trail.
- **Self-service task assignment.** Logged-in members claim free task slots and release
  their own. There is no release cutoff; the MV notification on the day before the game
  remains the safety net. Existing domain rules apply unchanged (one task per person per
  game; "plays itself" and "outside the responsible team" stay advisory warnings).
- **BREAKING — assignment API reshaped.** `POST /api/assignment` currently means "set the
  occupants of this role" (read-modify-write of the whole role). With several people
  editing at once this silently loses writes. It is replaced by explicit claim/release
  operations that reject the write when the slot is not in the expected state.
- **Audit trail.** Every assignment change (who assigned or released whom, for which task
  of which game, when, and acting as which tier) is recorded in an append-only table and
  is visible to admins in the web interface.
- **Persons are retired by deactivation, not deletion.** An admin deactivates someone who
  has left: past assignments, statistics and audit entries stay intact, upcoming slots are
  freed so nothing looks covered, and the person can neither log in nor be assigned until
  an admin reactivates them. Deletion remains only for entries created in error.
- **Contactless roster entries stay first-class.** Admins can create persons without
  e-mail or phone. Such persons receive no notifications and can never log in, but remain
  fully assignable so the schedule can show that a task is covered.
- **BREAKING — persons are identified by an internal ID, not by their name.** The name
  becomes a mutable display attribute that need not be unique, so members can correct or
  change their own name without losing their history. Every lookup, assignment, audit
  record and API payload refers to the internal identifier, which is never shown in the
  user interface.
- **Sessions expire after one hour** of inactivity, which makes an expired session a
  routine occurrence the interface has to handle gracefully rather than an edge case.
- **BREAKING — `NULIGAHELPER_SECRET` becomes mandatory.** The hardcoded default secret key
  in `webapp.py` is removed; the app refuses to start without a configured secret, because
  a session cookie now carries identity.

## Capabilities

### New Capabilities

- `user-accounts`: registration with contact channel and desired team, admin approval,
  passwordless login by magic link or SMS code, session lifetime, logout, abuse limits
  (per-person and per-IP throttling, enumeration-safe responses, token expiry).
- `access-control`: the guest/member/MV/admin tiers, what each tier may read and write on
  every page and API endpoint, and the rule that server-side checks — not disabled form
  controls — enforce permissions.
- `task-self-service`: members claiming and releasing their own task slots, the
  claim/release semantics that replace the set-style API, and how existing assignment
  rules and warnings apply to self-service.
- `assignment-audit`: what is recorded for every assignment change, that the record is
  append-only, and how admins view it.

### Modified Capabilities

<!-- None: openspec/specs/ is currently empty, so all behaviour is captured as new
     capabilities. Existing assignment rules are restated inside task-self-service
     where self-service changes who may trigger them. -->

## Impact

- **`webapp.py`** — the largest surface. `create_app()` gains authentication routes
  (`/login`, `/logout`, `/registrieren`) and a `before_request` guard; `build_schedule()`
  becomes viewer-dependent (a guest must not receive the roster payload that today is
  embedded in every `<select>`); every existing route gains a tier check; the assignment
  API is reshaped.
- **`db.py`** — new account/registration state and the audit table, plus claim/release
  helpers alongside `set_role_assignments`. Business rules stay here, not in `webapp.py`.
  The unique constraint on the person name is dropped and name-based lookup
  (`get_or_create_person`) is no longer an identity operation.
- **`templates/`, `static/app.js`** — login/registration pages, nav entries for login
  state, per-tier rendering of the schedule, and handling of 401 responses in `postJSON`,
  which today parses every response as JSON and would report a session timeout as
  "Server nicht erreichbar".
- **`manage_db.py`** — bootstrap of the first admin out-of-band, and person selection by
  internal ID now that names are neither unique nor stable.
- **`notifier.py`** — reused for login delivery; optionally confirms a self-assignment.
- **`test/test_webapp.py`** — the click-through scenario runs as an anonymous client
  today and will need an authenticated one, plus new coverage for rejected access.
- **Out of scope, tracked separately**: the database migration scheme (the DB currently
  holds mock data; a migration path will be proposed once real data exists) and
  production hardening for public exposure (WSGI server, TLS, cookie flags at the
  deployment level). Both are prerequisites for going online, neither is part of this
  change.
