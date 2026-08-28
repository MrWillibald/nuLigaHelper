## 1. Domain foundations (`db.py`)

- [ ] 1.1 Drop the uniqueness constraint on `Person.name` and make name-based lookup a
      seeding convenience only; verify `test/run_tests.sh` still passes and two persons
      with the same name can be created
- [ ] 1.2 Add account state to `Person` (admin flag, account status, desired team on
      registration, verification and approval timestamps); verify with a unit test that
      walks registered → verified → active
- [ ] 1.3 Add the audit model with actor, actor tier, action, affected person, game, role,
      slot, real datetime and the name/game text snapshots; verify a unit test reads back
      an entry after the referenced person is deleted
- [ ] 1.4 Add `claim_slot()` and `release_slot()` helpers taking the expected occupant and
      raising on mismatch; verify unit tests cover free, occupied and conflicting slots
- [ ] 1.5 Make every assignment mutation write an audit entry, including the cascade in
      `delete_person()`; verify a unit test asserts one entry per removed assignment
- [ ] 1.6 Extend `get_all_persons()` and the team helpers to exclude unapproved and
      deactivated persons; verify neither appears in any roster or selection query
- [ ] 1.7 Add `deactivate_person()` / `reactivate_person()` that free the person's
      assignments on future games, leave past ones untouched and stand the person down as
      MV; verify unit tests assert past assignments survive, future slots are freed with
      one audit entry each, and the team's `mv_person_id` is cleared

## 2. Authentication plumbing

- [ ] 2.1 Remove the hardcoded `app.secret_key` fallback and fail startup without
      `NULIGAHELPER_SECRET`; verify the app refuses to start with the variable unset
- [ ] 2.2 Configure a one-hour sliding session and expose the signed-in person to the
      request context; verify a test asserts expiry after simulated inactivity and
      renewal on activity
- [ ] 2.3 Implement single-use login tokens (signed, short TTL, nonce consumed on use) for
      the mail link and the SMS code; verify tests cover valid, replayed and expired
      tokens
- [ ] 2.4 Deliver login messages through the existing mail-preferred/phone-fallback rule;
      verify a test asserts the channel chosen for mail-only, phone-only and contactless
      persons
- [ ] 2.5 Add per-person and per-client rate limiting with a stricter cap on SMS; verify a
      test drives the limit and asserts the cooling-off refusal sends no message
- [ ] 2.6 Make login and registration responses identical for known and unknown contacts;
      verify a test compares both responses byte for byte

## 3. Registration and approval

- [ ] 3.1 Add the registration page with name, one contact channel, desired team and the
      consent confirmation; verify a test rejects submissions missing consent or a channel
- [ ] 3.2 Implement channel verification that activates login but not roster membership;
      verify a test asserts a verified person can sign in and is still not assignable
- [ ] 3.3 Implement approval and rejection, restricted to the MV of the requested team and
      to admins, with admin as fallback where the team has no MV or is the support team;
      verify tests cover MV approval, cross-team refusal and the fallback
- [ ] 3.4 Notify the responsible approver when a registration becomes pending; verify a
      test asserts the recipient is the team MV, or the admin where no MV exists
- [ ] 3.5 Show a pending registrant their status and nothing else; verify a test asserts a
      verified-but-unapproved session is refused the protected pages

## 4. Authorization layer

- [ ] 4.1 Derive the four tiers from stored facts (active account, MV records, admin flag);
      verify unit tests cover a person who is both admin and MV
- [ ] 4.2 Add a default-deny `before_request` guard with an explicit public allowlist;
      verify a test adds a throwaway route and asserts it is protected without opting in
- [ ] 4.3 Apply tier checks to every existing route per the access-control spec; verify a
      table-driven test walks all endpoints for all four tiers
- [ ] 4.4 Restrict the roster page: names and teams for members, contact data only for
      one's own record, full access for admins; verify a test asserts no foreign contact
      data appears in a member's response
- [ ] 4.5 Let members edit their own name and contact data, and warn when the last channel
      is cleared; verify a test asserts a member cannot edit another person
- [ ] 4.6 Add admin-only deactivate and reactivate actions to the roster page, and gate
      delete behind a warning naming deactivation as the alternative; verify tests cover
      the refusal for members and MVs and the warning text on delete
- [ ] 4.7 Refuse login for deactivated persons without revealing that the account exists;
      verify a test compares the response with the unknown-contact response
- [ ] 4.8 Return a distinguishable authentication failure for expired sessions on the JSON
      endpoints; verify a test asserts the status and payload differ from a validation
      error

## 5. Viewer-dependent schedule

- [ ] 5.1 Give `build_schedule()` a viewer and per-slot editability; verify tests assert
      the slot flags for guest, member, MV and admin
- [ ] 5.2 Ship the guest view without any roster payload or contact data; verify a test
      asserts the anonymous response contains no person not visibly assigned
- [ ] 5.3 Offer members their own name in free slots and release on their own slots only;
      verify a test asserts a member's rendered options contain nobody else
- [ ] 5.4 Offer MVs their own team's members on their team's games; verify a test asserts
      options are empty on a game owned by another team
- [ ] 5.5 Show the team alongside the name wherever a person is listed or offered; verify a
      test with two identically named persons shows both distinguishably

## 6. Claim/release API and client

- [ ] 6.1 Replace `POST /api/assignment` with claim and release endpoints carrying the
      expected occupant; verify tests cover success, conflict and the reported current
      state
- [ ] 6.2 Enforce the tier rules on both endpoints (self only, own team plus own games,
      admin anywhere); verify tests cover each refusal
- [ ] 6.3 Keep the existing assignment rules and advisory warnings for every tier, and
      refuse claims for unapproved persons; verify tests cover the one-task rule and both
      warnings
- [ ] 6.4 Refuse member and MV changes to past games while allowing admin corrections;
      verify tests cover both paths
- [ ] 6.5 Rework `postJSON()` in `static/app.js` to branch on HTTP status before parsing,
      handling 401 with a session-expired prompt and 409 by refreshing the affected card;
      verify a session timeout no longer reports "Server nicht erreichbar"
- [ ] 6.6 Add CSRF protection for the state-changing form posts and JSON endpoints; verify
      a test asserts a request without a valid token is refused

## 7. Audit review

- [ ] 7.1 Add the admin-only audit page listing entries newest first with filters by game
      and by person; verify a test asserts the ordering and both filters
- [ ] 7.2 Confirm no interface or CLI path edits or deletes an entry; verify a test asserts
      the absence of such an endpoint

## 8. CLI and operations

- [ ] 8.1 Add admin bootstrap to `manage_db.py`; verify granting admin to a person enables
      the admin tier on next login
- [ ] 8.2 Switch the CLI's person selection to the internal ID with a name-based search
      helper; verify a test covers two identically named persons
- [ ] 8.3 Add `NULIGAHELPER_SECRET` to `run_webapp.sh` and the daily job; verify both start
      cleanly with the variable set and refuse without it

## 9. Test suite

- [ ] 9.1 Rework the webapp scenario to sign in first and keep the top-to-bottom flow
      intact; verify `test/run_tests.sh` passes
- [ ] 9.2 Add a second scenario covering refusals across tiers; verify the new file passes
      standalone and under pytest
- [ ] 9.3 Add a concurrency test for two claims on one slot; verify the loser is refused
      and the winner's assignment is intact
- [ ] 9.4 Run the full suite and confirm no test touches `config.json` or the real
      database; verify `test/run_tests.sh` is green

## 10. Documentation

- [ ] 10.1 Update `AGENTS.md` with the tiers, the identity rule, the claim/release
      contract and the mandatory secret; verify the invariants section reflects the new
      behaviour
- [ ] 10.2 Update `README.MD` with registration, login, the permission model and the new
      environment requirements; verify the described commands match the code
