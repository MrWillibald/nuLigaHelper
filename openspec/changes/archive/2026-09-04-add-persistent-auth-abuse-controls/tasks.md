## 1. Policy Configuration and Persistent Model

- [x] 1.1 Define typed abuse-policy defaults and load/validate the new `auth_abuse` configuration without requiring existing private configs to change; verify unit tests cover defaults, overrides, unknown policies, non-positive values, and retention shorter than the longest window.
- [x] 1.2 Add the indexed `AuthAbuseCounter` SQLAlchemy model and composite uniqueness constraint as an additive table; verify a database initialized from scratch and an existing database both gain the table through `db.init_db()` without migration machinery or destructive recreation.
- [x] 1.3 Implement domain-separated HMAC subject derivation for canonical client, contact, person, and global identities; verify deterministic separation tests and database assertions show no raw IP, e-mail, phone, or person name is stored.

## 2. Atomic Enforcement and Retention

- [x] 2.1 Implement the dedicated-session fixed-window policy evaluator with one all-or-none SQLite `BEGIN IMMEDIATE` reservation across every applicable dimension; verify focused database tests cover allow, exact-limit refusal, window rollover, no partial increments, and no accidental commit of a caller's pending ORM work.
- [x] 2.2 Handle SQLite busy/storage failures as structured fail-closed outcomes and keep reservations committed before external side effects; verify fault-injection tests refuse the action without token/account/session mutation or notification dispatch.
- [x] 2.3 Add multi-connection concurrency tests in `test/test_concurrency.py` (or a focused auth-abuse test module) that race for the final client, subject, and global allowance and verify admitted reservations never exceed configured limits.
- [x] 2.4 Add bounded expiry cleanup with an `expires_at` index, startup invocation, and interval-limited opportunistic batches; verify tests preserve live rows, delete no more than the configured batch, drain backlogs over repeated runs, tolerate cleanup contention, and resume cleanup after restart.
- [x] 2.5 Add restart/shared-state tests that construct separate app instances against one synthetic SQLite file and verify both observe the same active counters and remaining cooling-off window.

## 3. Trusted Client Attribution

- [x] 3.1 Implement canonical direct-peer attribution that ignores forwarding headers by default; verify tests show spoofed `Forwarded` and `X-Forwarded-For` values cannot rotate client-scoped limits for direct requests.
- [x] 3.2 Add explicit trusted-proxy address/CIDR and exact-topology parsing with startup validation; verify tests cover the intended one-proxy path, IPv4/IPv6 normalization, untrusted peers, malformed/missing/extra forwarding values, and safe fallback or refusal.
- [x] 3.3 Wire the single client-attribution helper into every login request, registration request, and code-confirmation policy evaluation; verify route tests demonstrate one attributed client shares limits across all supplied contacts for each protected action.

## 4. Authentication Flow Integration

- [x] 4.1 Extend real and dummy signed challenge payloads with same-shaped opaque contact/person limiter subjects while keeping raw contacts and credentials out of the payload; verify decoding, tamper, expiry, known-contact, and unknown-contact tests.
- [x] 4.2 Replace the process-local login-request deque checks with atomic client, canonical-contact, resolved-person, and channel policies before token creation or dispatch; verify e-mail/SMS route tests cover known, unknown, ineligible, and throttled contacts and assert refused requests send nothing.
- [x] 4.3 Replace registration deque checks with one atomic reservation covering the client, every supplied canonical contact, and every resolved person before person creation/change or messaging; verify tests cover new registrations, duplicate contacts, two-contact submissions, replacement codes, account-exists messages, and zero side effects on refusal.
- [x] 4.4 Replace code-confirmation deque checks with atomic client and trusted challenge-subject policies before token consumption, verification, approver notification, or session establishment; verify login and registration tests cover wrong-code exhaustion, a correct code while throttled, malformed/dummy challenges, expiry, and successful confirmation after a new window.
- [x] 4.5 Remove `rate_events`, `_rate_allowed()`, and unused deque/defaultdict imports after all callers move to the persistent service; verify a repository search finds no process-local auth limiter and focused auth tests remain green.
- [x] 4.6 Add anti-enumeration regression assertions comparing HTTP status, rendered form step, challenge shape, and generic message for accepted-known, unknown, ineligible, throttled, and storage-failure requests; verify no comparison relies on delivery timing or reveals account existence.

## 5. SMS Cost Controls and Safe Observability

- [x] 5.1 Include per-person/contact and application-global SMS cost-cap rows in the same atomic reservation as SMS login/registration requests, while leaving e-mail independent; verify tests cover each cap, shared login-plus-registration consumption, window rollover, and concurrent final-allowance races with mocked Twilio dispatch counts.
- [x] 5.2 Add stable structured security events for allowed, throttled, failed-closed, cleanup, global-SMS-cap, and provider-failure outcomes at appropriate levels; verify log-capture tests assert action/channel/coarse dimension fields are present while sentinel names, IPs, contacts, HMAC digests, codes, signed challenges, sessions, secrets, form bodies, and provider exception payloads are absent.
- [x] 5.3 Ensure denied or failed SMS reservations preserve the success-shaped anti-enumeration page and never call Twilio; verify route tests assert identical response shape and zero dispatch for per-subject, global, and storage-failure refusals.

## 6. Operations, Documentation, and Validation

- [x] 6.1 Add every abuse-policy, cleanup, SMS-cap, and trusted-proxy setting with safe defaults and explanatory sample values to `config_template.json`; verify the template parses and matches the keys read by the application.
- [x] 6.2 Document fixed-window semantics, restart persistence, fail-closed behavior, retention/cleanup, HMAC-secret rotation effects, cap monitoring/tuning/reset procedures, and rollback limitations in `README.MD` or the production operations guide; verify every operator-visible setting has a documented purpose and safe deployment sequence.
- [x] 6.3 Coordinate the trusted peer, Flask bind address, forwarding-header replacement, and exact hop count with `deploy-production-web-runtime` when that change is available; verify documentation explicitly keeps proxy-derived attribution disabled and public exposure blocked until the runtime topology and application trust configuration agree.
- [x] 6.4 Document Twilio-side usage alerts/triggers, geographic permissions, and account/project spending-limit or prepaid-balance verification as an independent final backstop, noting account/region differences; verify the procedure requires no credentials in tracked files and does not claim unsupported provider features.
- [x] 6.5 Run the focused auth, refusal, and concurrency tests, then `test/run_tests.sh`, and run `openspec validate add-persistent-auth-abuse-controls --strict`; verify all commands pass and both debug switches in `common.py` remain `False`.
