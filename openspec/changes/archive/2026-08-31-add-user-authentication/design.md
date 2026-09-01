## Context

See `proposal.md` — Why. Constraints that shape the approach:

- One Flask app (`create_app()`), 3 read pages and 6 write endpoints, all currently open.
  Business rules live in `db.py`; `webapp.py` and `manage_db.py` are thin callers.
- The roster already encodes most of the identity graph: `Person.email` / `Person.phone`
  are the reachable channels, `Team.mv_person_id` names the MV of each team, and
  `notifier._dispatch()` already implements "mail preferred, phone fallback, skip if
  neither".
- Flask depends on `itsdangerous`, and `smtplib` plus Twilio are already wired up. Signed
  tokens and both delivery channels are therefore available without a new dependency.
- The database currently holds mock data, so this design may change the schema freely. A
  migration scheme is a separate, later change (proposal — Impact).
- Deployment today is the Flask dev server on plain HTTP in the LAN. Public exposure with
  a real WSGI server and TLS is planned but out of scope here; the design must not make
  that harder.
- Dates scraped from nuLiga are German `dd.mm.yyyy` strings and must keep using
  `db.game_sort_key()`. That convention applies to scraped game data only — everything
  introduced here (token issue times, audit timestamps, approval times) uses real
  datetimes.

## Goals / Non-Goals

**Goals:**

- Authentication that stores no passwords and adds no runtime dependency.
- A permission model where a missing check fails closed, not open.
- Assignment writes that are safe when several people edit the same game at once.
- An audit trail that survives deletion of the people it refers to.
- Keep contactless roster entries fully assignable.

**Non-Goals:**

- Production hardening for public exposure (WSGI server, TLS, `Secure` cookie flag,
  SQLite WAL). Prerequisites for going online, tracked separately.
- A database migration scheme.
- Notifications triggered by self-assignment (confirmation mails). The existing
  notification calendar in `main.py` is unchanged by this design.
- Fine-grained permissions beyond the four tiers, and any team-level administration
  beyond the MV tier.

## Decisions

### D1 — Identity is a `Person`; an account is an optional attribute of one

The session stores a `person_id`. There is no parallel user entity.

```
                    HAS CONTACT DATA?
                   yes            no
              +---------------+---------------+
         yes  | self-service  |  cannot exist |  passwordless login needs
   ACCOUNT?   | member        |               |  a channel to prove
              +---------------+---------------+
          no  | notified,     | placeholder   |  both stay assignable
              | admin-assigned| (admin-made)  |
              +---------------+---------------+
```

*Alternative considered:* a separate `users` table joined to `persons`. Rejected — it
makes "person" and "user" drift apart in every query, and three of the four quadrants
above have no user row anyway. The cost of the chosen model is that account state lives
as extra columns on `persons`, which is acceptable at club scale.

### D2 — Registration is a two-gate state machine

Verifying a channel and belonging to the club are different facts, so they are separate
transitions. An account can log in as soon as its channel is verified, but sees only its
own registration status until an admin approves it.

```
  [ registered ] --verify channel--> [ verified ] --approver accepts--> [ active ]
        |                                 |                                 |
        | token expires                   | approver rejects                | admin
        v                                 v                                 | deactivates
    (discarded)                      [ rejected ]                           v
                                                                     [ inactive ]

  assignable / visible on the roster:  active only
  may log in:                          verified, active, (inactive -> sees notice)
```

A registration carries the desired team; approval is what actually sets
`Person.team_id`. Until then the person is not in the roster, not in any dropdown, and
cannot be assigned — which is what stops an unapproved signup from making a task slot
look covered.

The approver is the **MV of the requested team**, with the admin as fallback and
override:

```
  registration for team T
        |
        +-- T has an MV      -->  that MV approves (admin may also act)
        +-- T has no MV      -->  admin approves
        +-- T is Supporter   -->  admin approves
```

An MV can only act on registrations for their own team, and approving does not let them
place the person into a different team — the team is fixed by what the registrant asked
for. Changing someone's team afterwards stays an admin action. This keeps the common case
(a player joining the team they already play in) off the admin's desk while leaving one
party responsible for everything else.

Leaving the club is the mirror image of joining it, and it takes the `inactive`
transition rather than a delete. Deactivation is what keeps the club's history readable:
statistics for past seasons, the audit trail and the record of who served a given game all
depend on the person row surviving. Three consequences follow from that, and they are the
reason deactivation is a distinct operation rather than a flag:

```
  deactivate person P
        |
        +-- past assignments   -->  untouched (history, statistics)
        +-- future assignments -->  released, each one audited
        +-- P is MV of a team  -->  team stood down to no MV
```

Freeing the upcoming slots is the point: a person who has left the club still occupying
Saturday's Verkauf is exactly the phantom coverage the approval gate exists to prevent,
arriving from the other direction. Reactivation restores the person but deliberately does
not restore those slots — someone else may have taken them, and the club has moved on.

Deletion survives for roster entries that should never have existed (a typo, a duplicate).
It removes assignments, so the interface names deactivation as the alternative before
confirming. The audit trail is unaffected either way, because it snapshots names (D7).

*Alternative considered:* auto-approve and moderate afterwards. Rejected — a filled slot
is a promise the club plans around, and `db.missing_slots()` (which drives both the
statistics gap list and the MV nag mail) would actively hide a fake one.

### D3 — Tiers are derived, never stored as a role string

```
  guest   := no valid session
  member  := session person, account active
  mv      := member AND exists Team where mv_person_id == person.id
  admin   := member AND person.is_admin
```

Only `is_admin` is new state. MV-ness already lives in `Team.mv_person_id`, and
duplicating it into a role column would create two sources of truth that drift the first
time an MV changes.

*Alternative considered:* a `role` column per person. Rejected for the duplication above;
also, a person can be MV of a team *and* admin, which a single column handles badly.

### D4 — MV authority is scoped by both team membership and game responsibility

An MV may assign and unassign **members of their own team**, on **games whose responsible
team is their own team**. Both conditions must hold.

*Alternatives considered:* either condition alone. "Any member of my team, any game" lets
an MV staff games another team owns; "any person, my team's games" lets an MV commit
people they have no relationship with. The intersection matches what the role actually
is — staffing the games the team was made responsible for.

### D5 — Assignment writes become compare-and-swap claim/release

`POST /api/assignment` today reads the role's current occupants, builds the full desired
list and rewrites it. Concurrent writers lose each other's changes silently:

```
   Anna                          Ben
   reads Verkauf [-, -]          reads Verkauf [-, -]
   POST desired=[Anna, -]        POST desired=[-, Ben]
        |                             |
        v                             v
   set_role_assignments([Anna])  set_role_assignments([Ben])   <-- Anna dropped,
                                                                   both got ok:true
```

It is replaced by two single-slot operations, each stating the occupant it expects to
find:

```
  POST /api/assignment/claim    {game_id, role, slot, expected_person_id: null,
                                 person_id: <self | someone, if mv/admin>}
  POST /api/assignment/release  {game_id, role, slot, expected_person_id: <occupant>}

  server: if actual occupant != expected_person_id  ->  409 + current state
```

Every mutation touches exactly one slot, which also gives the audit trail (D7) a natural
unit of record. Admins and MVs use the same endpoints with a `person_id` argument rather
than a separate set-style API, so there is one code path to authorize and one to audit.

Domain rules are unchanged and stay in `db.py`: one task per person per game is still
enforced (`uq_game_person`, `db.assign_person`), and "plays itself" / "outside the
responsible team" remain advisory warnings for every tier — including self-service.

### D6 — Rendering is viewer-dependent, and permission checks never live in the template

`build_schedule()` takes the viewer and returns per-slot `editable` flags plus a person
payload scoped to the tier:

```
  guest   ->  assigned names as text; NO roster payload
  member  ->  own name offered in free slots; own slots releasable
  mv      ->  + own team's members in free slots of the team's games
  admin   ->  full roster in every select (today's behaviour)
```

The guest case is a real disclosure fix, not just cosmetics: `templates/schedule.html`
currently renders a `<select>` containing **every person in the club** for every slot of
every game, so an anonymous visitor would otherwise receive the entire membership list —
including people who appear nowhere in the visible schedule.

A disabled control is not a permission. The `before_request` guard is default-deny with
an explicit allowlist of public endpoints, so a new route is protected until someone
deliberately opens it; and every endpoint re-checks tier and ownership regardless of what
the page rendered.

*Alternative considered:* a `@require_tier` decorator per route. Rejected — forgetting it
fails open.

### D7 — The audit trail is an append-only table that outlives its subjects

One row per assignment mutation: timestamp, actor person, tier the actor acted as, action
(`claim` / `release` / `assign` / `unassign`), target person, game, role, slot.

`db.delete_person()` cascades a person's assignments, and the admin's most common
question is exactly about someone who has since been removed. So the table stores both
the foreign keys (nullable, cleared on delete) **and** a text snapshot of the actor and
target names and the game identification. History stays readable after the rows it points
at are gone.

*Alternative considered:* `logging` into `helper.log`. Rejected — that file is not in the
Dropbox backup, lives on the Pi's SD card, and cannot answer "who dropped out of
Saturday's Verkauf?" in the UI. The requirement is an admin-visible record, not forensics.

### D8 — Login tokens are signed, short-lived and single-use

Magic links carry an `itsdangerous` signed token (person id + purpose + nonce, ~15 min
TTL); the SMS path sends a 6-digit code with the same lifetime. Consuming a token marks
its nonce used, so a link sitting in a mailbox cannot be replayed.

*Alternative considered:* a deterministic HMAC code derived from person id and a time
window, which needs no stored state. Rejected — no revocation, and a code stays valid for
the whole window no matter how often it is used. Single-use is worth the small amount of
state.

Delivery reuses `notifier._dispatch()` semantics so login inherits the house rule (mail
preferred, phone fallback, skip when neither) that is already tested.

### D9 — Abuse limits sized for a club, enumeration-safe by default

Per-person and per-IP throttling on "send me a login link" and on registration, with an
in-process token bucket. Login and registration responses are identical whether or not
the contact is known, so the form cannot be used to test who is in the club.

The in-process bucket is correct only for a single worker process. That is true of the
current deployment and of a small `waitress` setup with one worker; if the app is ever run
with multiple workers the counters must move into the database. Recorded as a risk rather
than pre-solved.

SMS deserves a tighter cap than mail: an attacker hammering the phone path spends real
money from the Twilio balance.

### D10 — Identity is an internal ID; the name is mutable display data

A person is identified everywhere by an internal numeric identifier that is never shown
in the interface. The name becomes ordinary profile data: a member may change their own
name as well as their e-mail and phone (including clearing the contact fields, which
stops notifications), and the uniqueness constraint on the name is dropped.

```
  identity   :  internal id      stable, invisible, used by every lookup,
                                 assignment, audit row and API payload
  display    :  name             mutable, may repeat, shown in the UI
  login      :  email / phone    proves control of the person's channel
```

Consequences worth stating, because they are not obvious:

- Two people on the roster may legitimately share a name. Wherever a person is chosen or
  listed, the interface must disambiguate — the team name is the natural qualifier, since
  it is already loaded for every dropdown.
- Name-squatting at registration stops being a threat (nobody can block a name), so the
  approval gate carries the impersonation concern on its own.
- Name-based lookup is no longer identity. `get_or_create_person(name, ...)` becomes a
  convenience for seeding and tests, not a way to resolve a person, and the CLI needs an
  ID-based selector.
- The audit trail's name snapshot (D7) changes from a nice-to-have to a requirement: it
  is the only record of what a person was called when the action happened.

*Alternative considered:* an opaque public identifier (UUID) instead of the existing
integer key. Rejected for now — the integer keys already flow through the JSON API, and
they are only ever visible to authenticated members, since the guest view ships names as
text and no identifiers at all. If the app later exposes per-person URLs to the public,
this is worth revisiting.

### D11 — CSRF: `SameSite=Lax` plus a signed token on state-changing forms

Cookie-based sessions on a public site need CSRF protection. `SameSite=Lax` blocks
cross-site form POSTs, and the JSON endpoints additionally require a signed token minted
with `itsdangerous` — no new dependency, and no reliance on the browser alone.

### D12 — The secret key becomes mandatory

`app.secret_key` currently falls back to a hardcoded literal. Once the cookie carries
identity, that default is a public forging key, so the app refuses to start without
`NULIGAHELPER_SECRET`. The first admin is bootstrapped out-of-band via `manage_db.py`,
matching how that CLI is already used.

### D13 — Sessions expire after one hour of inactivity

The session cookie is valid for one hour and is refreshed on every request, so an hour of
*inactivity* ends it rather than an hour of wall time. Being logged out in the middle of
staffing a game day is the failure mode the club would notice; a sliding window avoids it
while still bounding how long a forgotten session on a shared device stays usable.

*Alternative considered:* an absolute one-hour cap regardless of activity. Rejected as
user-hostile for the one workflow that takes longest — an admin working through a season
of games in one sitting.

This makes expiry a routine event rather than an edge case, which is precisely why the
401 handling in `static/app.js` is part of this change and not a follow-up: a page left
open over lunch will fail its next save, and today that surfaces as
"Server nicht erreichbar".

## Risks / Trade-offs

- **In-process rate limiting breaks under multiple workers** → single worker today;
  documented as a deployment constraint, counters move to the DB if that changes.
- **A magic link grants access to whoever reads the mailbox** (shared family address,
  forwarded mail) → short TTL, single use, logout, and the audit trail records what the
  session did.
- **SMS login costs money and can be triggered by strangers** → tighter per-person and
  per-IP caps on the phone path, mail preferred by the dispatch rule.
- **The approval gate adds admin friction**; a helper who signs up on Saturday may wait
  for approval → registrations are few per season, and the admin can be notified on new
  registrations using the existing mail path.
- **Names of self-registered members are published on a public page** → consent has to be
  obtained at registration; a German Verein needs this to be explicit in the signup flow.
- **Session cookies travel in clear over LAN HTTP until TLS is in place** → do not expose
  the app publicly before the deployment change lands; `Secure` is set there.
- **Rotating `NULIGAHELPER_SECRET` logs everyone out** → acceptable, but the value must
  live somewhere persistent rather than a shell one-liner.
- **Duplicate names on the roster are now possible and dropdowns become ambiguous** →
  every person picker and list shows the team alongside the name; the admin can rename to
  disambiguate, since names are free-form.
- **A member can rename themselves after committing to a task**, which makes the schedule
  read differently than when the promise was made → the audit trail records the name in
  force at the time of each action, so the change is visible rather than silent.
- **The webapp test scenario runs as an anonymous client and will break wholesale** →
  the scenario gains a login step, and rejection paths get their own tests; expect this
  to be a real chunk of the work rather than a mechanical fix.
- **Concurrent writes from the web app and the daily job against one SQLite file** →
  volume is tiny; WAL mode belongs to the deployment change.

## Migration Plan

1. Set `NULIGAHELPER_SECRET` in the environment of both `run_webapp.sh` and the cron job.
2. Delete the database file and let the next run recreate it — the current contents are
   mock data (proposal — Impact). This is the last change that may assume that.
3. Bootstrap the first admin via `manage_db.py` and verify login end to end on the LAN.
4. Re-enter the real roster, or import it, once the schema is stable.

Rollback is a code revert: the schema additions are additive and an older revision
ignores them, so a broken login does not cost the assignment data.

## Open Questions

None outstanding. Session lifetime (D13) and MV approval rights (D2) were both decided
while writing this design.
