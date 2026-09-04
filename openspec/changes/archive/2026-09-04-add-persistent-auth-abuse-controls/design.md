## Context

See `proposal.md` for motivation and `specs/user-accounts/spec.md` for the behavioral contract.

`webapp.create_app()` currently owns a `defaultdict(deque)` and `_rate_allowed()`. Login and registration requests use it for per-client and some per-person limits, while code confirmation uses separate client/person keys. The state disappears at restart, is not shared across processes, and performs an unprotected check-then-append. Client keys use `request.remote_addr` directly. Authentication tokens are already persistent and single-use in SQLite, and each Flask request obtains a SQLAlchemy session against the same database.

The target is one Raspberry Pi running Flask behind a production reverse proxy, with SQLite as the shared local store. Authentication volume is low, outbound SMS costs money, and adding another always-on service would materially increase operational complexity. The repository intentionally has no migration framework; additive tables can still be created by the existing `Base.metadata.create_all()` path.

## Goals / Non-Goals

**Goals:**

- Make every abuse decision shared, restart-safe, atomic under concurrent Flask workers, and fail-closed before an authentication side effect.
- Keep lock duration and database growth predictable for SQLite on Raspberry Pi hardware.
- Use one explicit policy vocabulary for actions, dimensions, channels, limits, windows, and outcomes.
- Keep contacts and credentials out of limiter storage and logs while allowing stable matching of canonical contacts.
- Make trusted proxy attribution auditable and safe even if a caller injects forwarding headers.
- Bound application-side SMS delivery and document Twilio-side cost backstops.

**Non-Goals:**

- Replacing Flask, SQLite, SQLAlchemy, passwordless challenges, sessions, or account states.
- Building a general-purpose distributed rate-limit service or adding Redis for the one-node deployment.
- Solving volumetric denial-of-service before traffic reaches the Raspberry Pi; the reverse proxy/firewall remains responsible for connection-level protection.
- Guaranteeing delivery after an allowance is reserved. Safety favors occasionally consuming an allowance when a process fails before dispatch over sending beyond a cap.
- Redesigning the authentication forms or changing validation messages unrelated to anti-enumeration.

## Decisions

### D1 — Store compact fixed-window counters in SQLite

Add an `AuthAbuseCounter` SQLAlchemy model with fields equivalent to:

- `action`: `login_request`, `registration_request`, `code_confirmation`, or `sms_delivery`;
- `dimension`: `client`, `contact`, `person`, or `global`;
- `subject_digest`: a fixed-length opaque digest (the global subject uses a constant domain value);
- `channel`: `email`, `sms`, or `any`;
- `window_started_at`, `count`, and `expires_at`.

A unique constraint over `(action, dimension, subject_digest, channel, window_started_at)` prevents duplicate counters for one policy bucket. Index `expires_at` for cleanup. One admitted operation creates or increments only the rows for its applicable rules, so storage is proportional to admitted identities and active windows rather than every rejected request.

Fixed windows are chosen over an append-only event table because they use far fewer rows and make bounded cleanup inexpensive. They are chosen over an in-memory token bucket because persistence and worker sharing are requirements. The policy defaults retain today's short-window behavior:

| Action/dimension | Default policy |
| --- | --- |
| Login request, client | 10 per 15 minutes |
| Login request, person and contact | E-mail 3, SMS 2 per 15 minutes |
| Registration request, client | 5 per 15 minutes |
| Registration request, each supplied contact and resolved person | E-mail 3, SMS 2 per 15 minutes |
| Code confirmation, client | 10 per 15 minutes |
| Code confirmation, challenge contact/person | 5 per 15 minutes |
| SMS cost cap, person and contact | 6 per rolling operational day represented as a 24-hour fixed window |
| SMS cost cap, global | 30 per rolling operational day represented as a 24-hour fixed window |

The implementation will define window alignment explicitly in UTC epoch time. The phrase “operational day” in configuration means a configured 24-hour fixed bucket, not local-midnight accounting; this avoids daylight-saving ambiguity. Limits and window lengths are configurable, and startup validation rejects missing types, non-positive limits/windows, retention shorter than the longest window, or unknown policy names. Built-in defaults preserve safe behavior when old `config.json` files omit the new section; `config_template.json` documents all keys.

A fixed window permits a boundary burst of up to twice a nominal limit across adjacent buckets. That trade-off is acceptable for this low-volume deployment because the stricter per-subject, per-client, and global SMS layers overlap. If tests or production evidence show boundary bursts are unacceptable, the model can move to a generic-cell-rate/token-bucket calculation in the same table without adding Redis.

**Alternatives considered:**

- Append every attempt and count a precise sliding window: semantically exact but creates more rows and cleanup pressure under attack.
- Keep process-local counters and enforce one worker: restart bypass remains and production topology becomes a security invariant.
- Redis: unnecessary operational and backup burden for a single host whose SQLite database is already the shared durable authority.

### D2 — Reserve all applicable allowances in one short `BEGIN IMMEDIATE` transaction

Implement a database helper that accepts the complete set of applicable policy keys for one protected action. It opens a dedicated short-lived session/connection, starts SQLite `BEGIN IMMEDIATE`, loads or creates the current counter rows, and checks every rule before incrementing any. If one rule is exhausted, it rolls back without incrementing the other rows and returns a structured refusal naming only coarse dimensions. If all rules allow the action, it increments all rows and commits before the caller performs account mutation, token consumption, notification, or session establishment.

`BEGIN IMMEDIATE` serializes competing writers before the read/check/update sequence, making the multi-row decision all-or-none across workers. A database busy timeout or other storage error is treated as a failed-closed result. The dedicated transaction prevents limiter commits from accidentally committing request-session account changes and keeps the write lock out of SMTP/Twilio calls.

Reservations are not refunded when downstream delivery fails or a worker exits. This is deliberately conservative: refunds would require coupling provider outcomes back into concurrency-sensitive counters and could allow retry storms to exceed spending limits. The application logs a redacted delivery failure separately.

**Alternatives considered:**

- Independent atomic upserts for each dimension: a later failed dimension would leave partial reservations and create order-dependent throttling.
- A normal deferred SQLite transaction: concurrent readers can both observe the final allowance before either writes.
- Hold the transaction through Twilio/SMTP dispatch: prevents overspend but holds SQLite's writer lock across network latency and failure; committing the reservation first provides the same upper bound without that availability cost.

### D3 — Build typed, keyed subject digests from canonical identities

Derive `subject_digest` using HMAC-SHA-256 with `NULIGAHELPER_SECRET` and an explicit domain separator containing schema version, dimension, and channel. Inputs are:

- normalized IP address bytes for `client`;
- canonical normalized e-mail or E.164 phone for `contact`;
- internal `Person.id` for `person`;
- a constant value for `global`.

HMAC, rather than a plain hash, prevents offline recovery of low-entropy phone numbers or predictable e-mail addresses from copied database rows. Domain separation prevents equality comparisons across dimensions. Raw inputs and digests are never logged. Secret rotation already invalidates sessions and challenges; it will also start fresh limiter identities while old rows age out, which must be documented as a security consequence.

Login requests apply client and contact rules for every syntactically valid submission, plus person rules when a person resolves. Registration applies client rules, each supplied canonical contact rule, and every uniquely resolved person rule before creating/updating a pending account or sending an account-exists message. This ensures unknown contacts are throttled without creating person rows and prevents a second registration contact from bypassing controls.

Signed challenge payloads gain opaque contact/person subject digests. Real and dummy challenges carry the same fields, allowing confirmation attempts to use client plus challenge-subject policies without a user-visible distinction. Invalid or malformed challenges have no trusted subject and receive client-scoped enforcement plus the existing generic invalid/expired result.

**Alternatives considered:** storing raw keys (privacy risk), unsalted hashes (dictionary attacks), or only `Person.id` (unknown contacts and pre-registration attempts evade subject limits).

### D4 — Resolve client identity before applying any authentication policy

Add one client-attribution helper used by all protected auth actions. By default it canonicalizes the direct WSGI peer address and ignores `Forwarded`/`X-Forwarded-For`. Proxy-derived attribution is enabled only with explicit production configuration containing trusted proxy CIDRs/addresses and an exact trusted-hop count/topology.

For the intended one-proxy deployment, the reverse proxy must replace, not append user-supplied forwarding metadata, bind the Flask server so clients cannot bypass the proxy, and send exactly one validated origin address. The helper first verifies that the direct peer is trusted, then validates the expected number and syntax of forwarded addresses from the trusted side of the chain. Missing, malformed, extra, or untrusted topology never selects a caller-controlled identity; it falls back to the trusted direct peer or returns a failed-closed attribution result according to configuration.

Do not install a blanket `ProxyFix` based only on hop count while Flask remains directly reachable. If `deploy-production-web-runtime` defines the reverse proxy, this change consumes and documents its concrete bind address, proxy address/range, replacement-header behavior, and hop count. Public exposure is blocked operationally until those settings agree.

**Alternatives considered:** trusting `request.access_route` or the leftmost forwarding value (spoofable), and trusting a hop count without restricting direct peers (safe only if network reachability is perfectly enforced and harder to test in isolation).

### D5 — Centralize policy application around three auth action boundaries

Replace route-local `_rate_allowed()` calls with a small abuse-control service that returns `allowed`, `reason_dimensions`, and `storage_error`, without returning identifiers. The routes evaluate limits at these boundaries:

1. **Login code request:** after contact validation/canonicalization and lookup, before token creation or dispatch.
2. **Registration code request:** after form/contact validation and conflict lookup, before person creation/change, token creation, account-exists messaging, or dispatch.
3. **Code confirmation:** after signed challenge parsing and read-only token lookup, before token consumption, verification, session establishment, or approver notification.

For SMS code requests, the same atomic reservation includes short-window request rules plus the per-person/contact and global `sms_delivery` caps. Only an admitted reservation may call Twilio. E-mail never consumes SMS cost-cap rows.

The existing success-shaped login/registration page and dummy challenge paths remain in place for unknown, ineligible, throttled, and limiter-storage-failure requests. Confirmation refusals use the existing “Code ungültig oder abgelaufen.” result and do not consume a valid code. HTTP status, rendered step, and generic message are asserted equivalent in tests; network timing is not promised.

### D6 — Make cleanup bounded, indexed, and opportunistic

Set each counter's `expires_at` to the end of its policy window plus configured retention padding. Retention defaults to seven days and startup validation requires it to cover the longest active policy/reporting window.

A cleanup helper selects at most a configurable batch (default 200) of expired row IDs ordered by `expires_at`, deletes that batch in a short transaction, and reports the count. Run one batch at application startup and opportunistically no more than once per configured interval (default five minutes) after an abuse-control transaction. Coordination may use a persisted global cleanup marker or tolerate duplicate bounded cleaners; correctness does not depend on cleanup winning. If traffic stops, rows do not grow; the next startup/request resumes cleanup. Repeated runs eventually drain any backlog without an unbounded delete in a request.

Cleanup uses its own transaction and never deletes a row whose `expires_at` is still live. Lock/busy failures skip that cleanup batch and log a redacted warning; they do not undo an already committed abuse reservation.

**Alternative considered:** a mandatory external scheduler. Rejected because the Flask process already has natural startup/request hooks and the Raspberry Pi deployment should not gain another required moving part.

### D7 — Add structured, privacy-safe security logging

Use a dedicated logger namespace and stable event names such as `auth_abuse_allowed`, `auth_abuse_throttled`, `auth_abuse_storage_error`, `auth_abuse_cleanup`, and `auth_delivery_failed`. Include action, channel, coarse dimensions, configured window identifier, and counts only where they do not reveal a subject. Do not include names, raw IPs, e-mail addresses, phone numbers, subject digests, codes, signed challenges, session/cookie values, secrets, full form bodies, or Twilio/SMTP exception text that may echo a destination.

Unexpected provider/storage exceptions are logged by exception class and a stable internal reason code; detailed sensitive provider payloads are not interpolated. Tests capture logs for known sentinel contacts/codes/challenges and assert those values and their digest are absent.

Allowed-event logging can be `INFO` or sampled/debug-level to avoid routine noise; throttles and global SMS refusals are `WARNING`; cleanup summaries are `INFO`; fail-closed storage errors are `ERROR`. Configuration never enables raw-key debugging.

### D8 — Layer application caps with Twilio account controls

Add an `auth_abuse` configuration section to `config_template.json` for policy limits/windows, retention/cleanup bounds, and SMS per-person/contact/global caps. Keep Twilio credentials in the existing section. Document:

- how to choose caps from expected club login volume and current Twilio pricing;
- how to configure Twilio usage triggers/alerts, geographic permissions, and an account/project spending limit or prepaid balance where the account supports it;
- that exact Twilio console features vary by account/region and must be verified by the operator;
- that application reservations bound calls initiated by this app but provider-side controls are the independent final backstop;
- how to monitor redacted global-cap warnings and deliberately reset/adjust a cap without deleting arbitrary auth data.

No network access or live Twilio account is needed for implementation/tests. Provider calls remain mocked in the offline suite.

## Risks / Trade-offs

- [Fixed-window boundaries allow short bursts above a nominal rolling rate] → Overlap client, contact/person, and global SMS controls; document window semantics and revisit the algorithm only if evidence warrants it.
- [SQLite serializes writers] → Use tiny dedicated `BEGIN IMMEDIATE` transactions, commit before network calls, index exact lookup/cleanup columns, and keep cleanup batches small.
- [A busy or damaged database can block legitimate login] → Fail closed for security, emit a redacted high-severity event, and document database health/recovery checks.
- [Reservation can be consumed without a delivered message] → Accept conservative false throttling rather than overspend; windows expire automatically and operators can diagnose provider failures separately.
- [HMAC-key rotation resets effective limits] → Document the coupling to `NULIGAHELPER_SECRET`; retain a stable secret in production and treat rotation as a security operation.
- [Proxy misconfiguration can collapse users onto the proxy identity or permit spoofing] → Default to direct peers, require explicit trust plus exact topology, test malformed chains, and align deployment docs before public exposure.
- [Application caps may block legitimate game-day login spikes] → Keep limits configurable, ship documented club-sized defaults, log coarse refusal dimensions, and require deliberate operator tuning rather than bypass code.
- [Stale rows remain while the app is idle] → They cannot grow without traffic; bounded startup/opportunistic cleanup drains them when service resumes.
- [Twilio console controls differ by plan and region] → Document goals and verification steps without claiming unavailable features; require alerts/limits supported by the actual account.

## Migration Plan

1. Implement the additive counter table and helpers. The existing `create_all()` startup path creates the new table on an existing database, so no migration framework or destructive database recreation is required.
2. Add and validate defaults/configuration, then run focused offline persistence, concurrency, route, proxy, cleanup, logging, and SMS-cap tests followed by `test/run_tests.sh`.
3. Deploy with forwarding headers ignored and direct-peer attribution first. Existing process-local counters disappear at this release; there is intentionally no import because their state was ephemeral.
4. If `deploy-production-web-runtime` is present, verify its Flask bind address, reverse-proxy replacement of forwarding headers, trusted proxy address/range, and hop count. Enable proxy-derived attribution only after direct access to Flask is blocked and tests/health checks show the expected client identity.
5. Configure conservative application SMS caps and independent Twilio alerts/spending protections before enabling public SMS authentication. Observe redacted throttle/global-cap logs and tune only from expected club usage.
6. Rollback by deploying the previous application version. The additive table may remain unused and can be removed only during a deliberate maintenance action; rollback restores process-local limiting and therefore is not acceptable as a long-term publicly exposed state.
