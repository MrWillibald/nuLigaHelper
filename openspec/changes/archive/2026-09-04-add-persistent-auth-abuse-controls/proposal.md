## Why

The current authentication throttles live in a process-local `rate_events` deque, so restarting the Flask process clears every limit and multiple workers do not share enforcement. Before the web UI is exposed through the production runtime, authentication abuse controls need durable, concurrency-safe enforcement that fits the one-node Raspberry Pi and bounds both account attacks and Twilio spend.

## What Changes

- Replace process-local authentication counters with restart-safe, shared abuse-control state stored through the existing SQLite/SQLAlchemy stack, without introducing Redis unless implementation evidence shows SQLite cannot meet the stated correctness needs.
- Apply configurable throttles to login code requests, code confirmations, and registration attempts across both trusted client identity and person/contact identity, including unknown contacts without weakening anti-enumeration behavior.
- Accept a proxy-derived client IP only when the request passed through explicitly trusted proxy configuration; otherwise use the direct peer and never trust arbitrary forwarding headers.
- Make limit checks and event consumption concurrency-safe so simultaneous requests cannot exceed configured thresholds through race conditions.
- Add bounded retention and opportunistic or scheduled cleanup so abuse-control state cannot grow without limit.
- Add configurable SMS safeguards at per-account/contact and application-global levels, with safe refusal behavior when a cap is reached.
- Add privacy-safe security logging that records decisions and coarse dimensions without raw names, contacts, authentication codes, signed challenges, or reusable secrets.
- Add offline tests covering restart persistence, shared enforcement, concurrency, trusted-proxy handling, each protected auth action and identity dimension, anti-enumeration response equivalence, cleanup, configuration, SMS caps, and redacted logging.
- Document deployment configuration and operational Twilio spending-limit/alert controls. The change may depend on `deploy-production-web-runtime` for the final trusted-proxy topology and production headers.
- Explicitly exclude a FastAPI rewrite, a broad authentication redesign, and unrelated changes to account/session semantics.

## Capabilities

### New Capabilities

_None._

### Modified Capabilities

- `user-accounts`: Strengthen authentication and registration abuse controls with persistent/shared enforcement, trusted client attribution, confirmation-attempt throttling, bounded retention, privacy-safe observability, and configurable SMS cost safeguards while preserving anti-enumeration responses.

## Impact

- Affects authentication and registration request handling in `webapp.py`, persistent models/helpers in `db.py`, auth-related configuration in `config_template.json`, offline auth/concurrency tests, and deployment/operator documentation in `README.MD` or production-runtime documentation.
- Adds SQLite schema state; consistent with the project's no-migration policy, the implementation plan must call out database recreation and defer any persistence-preserving migration machinery to an explicit owner decision.
- Does not require a new network service or Redis for the intended one-node Flask+SQLite Raspberry Pi deployment.
- Coordinates with `deploy-production-web-runtime` where trusted reverse-proxy boundaries, forwarded-header handling, and public deployment settings are defined.
