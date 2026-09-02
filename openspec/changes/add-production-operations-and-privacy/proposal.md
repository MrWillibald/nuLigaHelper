## Why

A public launch needs an operable and privacy-conscious production baseline beyond merely serving the Flask application: least-privilege deployment, protected secrets and data, visible failures, routine host checks, public legal information, and documented lifecycle procedures. Establishing these requirements before launch makes ownership, review inputs, incident response, and rollback explicit instead of relying on ad hoc Raspberry Pi administration.

## What Changes

- Define a dedicated unprivileged service identity, restrictive ownership/modes for `config.json`, `.nuligahelper_secret`, the SQLite database and its containing directories, and systemd-compatible environment/secret handling.
- Define journald logging, retention, access and redaction expectations so operational evidence remains useful without exposing credentials, authentication material, contact data or unnecessary personal data.
- Make web-service, daily-job and backup failures actionable through explicit status, diagnostics and operator-visible alerting, with a minimal non-sensitive health check and an operator monitoring/runbook procedure.
- Add launch checks for time synchronization, disk capacity and TLS certificate validity, plus a repeatable launch and rollback checklist.
- Integrate publicly reachable German `Impressum` and `Datenschutzerklärung` pages/navigation using content supplied by the operator and reviewed by the appropriate legal/privacy reviewer; this change does not author legal text.
- Define an operational retention/deletion policy for authentication tokens, rejected registrations and related personal data, including the deliberate treatment of append-only assignment audits, backups and third-party processors.
- Coordinate sequencing and acceptance boundaries with `deploy-production-web-runtime`, `add-persistent-auth-abuse-controls` and `make-sqlite-production-safe` without absorbing their runtime, abuse-control or SQLite-concurrency scope.

## Capabilities

### New Capabilities

- `production-operations`: Least-privilege production hosting, secret/file protection, redacted journal handling, failure visibility, health monitoring, host checks, runbooks, and launch/rollback controls.
- `privacy-data-lifecycle`: Operator-approved retention and deletion rules for authentication and registration data, with explicit audit, backup and processor implications.
- `public-legal-information`: Public integration and release gating for operator-supplied, legally reviewed Impressum and privacy-notice content.

### Modified Capabilities

- None.

## Impact

- Affects production service/timer definitions, deployment ownership and filesystem layout, environment configuration, logging conventions, health/status behavior, operational scripts or commands, and production documentation/runbooks.
- Affects public Flask routes, templates and navigation only to host supplied legal/privacy content and make it consistently reachable.
- Establishes operational cleanup and verification procedures around existing authentication/registration records, assignment audits and Dropbox database backups; implementation must preserve current identity and append-only audit guarantees unless a separately approved spec changes them.
- Requires operator decisions and legal/privacy review for legal text, retention periods, processor disclosures, alert destinations and final launch approval.
- Depends on or must be reconciled with the planned changes `deploy-production-web-runtime`, `add-persistent-auth-abuse-controls` and `make-sqlite-production-safe`; no network service or new third-party monitoring dependency is assumed by this proposal.
