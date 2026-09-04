# production-web-runtime Specification

## Purpose

Defines a secure, supportable production boundary for serving the existing web interface from a Linux or Raspberry Pi host over the public internet.

## Requirements

### Requirement: Production traffic uses a supervised WSGI service

The production deployment SHALL serve the existing Flask application through a production WSGI process supervised by the operating system. The backend SHALL listen only on a loopback address and SHALL restart after an unexpected process failure without exposing the backend listener directly to the network.

#### Scenario: Production service starts successfully
- **WHEN** the host boots with valid production configuration and an accessible database
- **THEN** the operating system starts the web service
- **AND** the WSGI backend accepts requests only through its loopback listener

#### Scenario: Web process exits unexpectedly
- **WHEN** the production web process terminates unexpectedly
- **THEN** the service supervisor records the failure and restarts it according to a bounded restart policy

#### Scenario: Required startup configuration is absent
- **WHEN** the production service starts without its persistent application secret, trusted public host, or database path
- **THEN** startup fails with an actionable error
- **AND** the service does not fall back to an insecure default

### Requirement: Public traffic terminates at an HTTPS reverse proxy

The production deployment SHALL expose the web interface through a reverse proxy on the configured public hostname. The proxy SHALL automatically obtain and renew a publicly trusted certificate, SHALL redirect plaintext HTTP requests to HTTPS, and SHALL forward accepted traffic only to the loopback backend.

#### Scenario: Visitor uses the public hostname over HTTPS
- **WHEN** a visitor requests the configured public hostname over HTTPS with a valid certificate
- **THEN** the proxy forwards the request to the backend
- **AND** the visitor receives the web application response over HTTPS

#### Scenario: Visitor uses plaintext HTTP
- **WHEN** a visitor requests the configured public hostname over HTTP
- **THEN** the proxy redirects the visitor to the equivalent HTTPS URL

#### Scenario: Certificate prerequisites are not met
- **WHEN** public DNS or inbound HTTP/HTTPS connectivity does not permit certificate issuance
- **THEN** the deployment does not claim successful public readiness
- **AND** the operator documentation identifies the failed prerequisite and diagnostic commands

### Requirement: Proxy-derived request metadata has a strict trust boundary

The production application SHALL accept forwarded client address, scheme, and host metadata only from the single local reverse-proxy hop. The edge proxy SHALL replace, rather than preserve, internet-supplied forwarding metadata, and the application SHALL reject malformed, multi-hop, missing, or otherwise untrusted production proxy metadata instead of using it for security decisions, URL generation, logging, or rate limiting.

#### Scenario: Request arrives through the configured proxy
- **WHEN** the local reverse proxy supplies one valid client address, HTTPS scheme, and trusted host
- **THEN** the application treats those values as the effective request metadata
- **AND** URL generation, secure-request detection, request logging, and client-based rate limiting use that effective metadata

#### Scenario: Internet client spoofs forwarding headers
- **WHEN** an internet client sends its own forwarded client, scheme, or host headers to the public proxy
- **THEN** the proxy discards or overwrites those values
- **AND** the application sees only metadata established by the configured proxy

#### Scenario: Backend receives an untrusted direct request
- **WHEN** a production backend request does not originate from the configured loopback proxy boundary or contains an invalid forwarding chain
- **THEN** the application rejects the request
- **AND** does not treat supplied forwarding values as trusted metadata

### Requirement: Production requests use explicitly trusted hosts

The production service SHALL require an explicit allowlist of public hostnames. It SHALL accept only exact configured hosts needed by the deployment and SHALL reject requests with an absent, malformed, or unlisted host before executing the requested application endpoint.

#### Scenario: Configured hostname is requested
- **WHEN** a proxied request carries a hostname in the production allowlist
- **THEN** normal routing proceeds

#### Scenario: Unlisted hostname is requested
- **WHEN** a request carries a hostname that is not in the production allowlist
- **THEN** the request is rejected with a client error
- **AND** no state-changing endpoint executes

#### Scenario: Production host allowlist is empty
- **WHEN** the application starts in production mode without at least one explicit trusted host
- **THEN** startup fails rather than accepting arbitrary hosts

### Requirement: Production sessions use hardened cookies

Session cookies issued in production SHALL be Secure, HTTP-only, restricted to the application path, host-only unless an explicit domain is required, and SameSite=Lax. The existing one-hour sliding inactivity lifetime SHALL remain unchanged, and production SHALL require a stable externally supplied signing secret.

#### Scenario: Production login succeeds
- **WHEN** an eligible person completes login through HTTPS
- **THEN** the response issues a session cookie with Secure, HttpOnly, SameSite=Lax, and Path=/ attributes
- **AND** the cookie does not set a Domain attribute by default

#### Scenario: Local developer launch is used
- **WHEN** a developer starts the documented local launcher without production mode
- **THEN** the application remains usable over local HTTP
- **AND** production-only secure-cookie and proxy-boundary assumptions are not silently enabled

#### Scenario: Application secret changes
- **WHEN** the configured production signing secret is rotated
- **THEN** existing sessions become invalid
- **AND** the deployment documentation warns the operator of that consequence

### Requirement: Request bodies are bounded at both ingress layers

The production reverse proxy and application SHALL enforce a documented maximum request-body size appropriate for form and JSON requests with no file-upload feature. A body exceeding the limit SHALL be rejected with HTTP 413 before endpoint logic mutates application state, and the application-level limit SHALL remain effective if the edge check is bypassed locally.

#### Scenario: Request body is within the configured limit
- **WHEN** a valid form or JSON request body does not exceed the documented production limit
- **THEN** normal request validation and routing proceed

#### Scenario: Oversized request reaches the public proxy
- **WHEN** a request body exceeds the proxy limit
- **THEN** the proxy rejects it with HTTP 413 without forwarding the complete body to the backend

#### Scenario: Oversized request reaches the backend
- **WHEN** a request body exceeds the application limit despite bypassing the edge check
- **THEN** the application rejects it with HTTP 413
- **AND** no application state is changed

### Requirement: Production responses include browser security headers

Production HTTPS responses SHALL include a documented, application-compatible security-header baseline covering transport security, content-type sniffing, framing, referrer disclosure, browser capabilities, and content sources. The content security policy SHALL allow only the sources required by the current application and SHALL not permit inline script execution; any temporary allowance for existing inline style attributes SHALL be documented.

#### Scenario: Normal HTTPS response is returned
- **WHEN** a visitor receives an application response through the production hostname
- **THEN** the response includes Strict-Transport-Security, X-Content-Type-Options, a frame restriction, Referrer-Policy, Permissions-Policy, and Content-Security-Policy headers
- **AND** the content policy limits scripts to same-origin resources and blocks object embedding and framing

#### Scenario: Proxy returns an error response
- **WHEN** the edge proxy rejects a request before it reaches the application
- **THEN** the proxy response still includes the applicable production security headers

#### Scenario: Current pages render under the policy
- **WHEN** the schedule, authentication, roster, statistics, and audit pages are loaded through the production boundary
- **THEN** their same-origin static assets and existing controlled inline style attributes remain functional
- **AND** no inline script allowance is required

### Requirement: Deployment assets and operations are documented

The repository SHALL provide versioned example configuration and a Raspberry Pi/Linux deployment guide covering prerequisites, installation, service identity and permissions, persistent secret handling, database location, reverse-proxy hostname configuration, startup, validation, logs, updates, backups, and rollback. Examples SHALL contain placeholders rather than live credentials and SHALL ensure the web service and daily job use the same database and stable secret.

#### Scenario: Operator performs a fresh deployment
- **WHEN** an operator follows the deployment guide on a supported Linux/Raspberry Pi host and supplies host-specific values
- **THEN** they can install and start the supervised backend and HTTPS proxy without relying on undocumented steps

#### Scenario: Operator updates the application
- **WHEN** an operator deploys a new revision
- **THEN** the guide provides ordered backup, dependency, validation, restart, and health-check steps
- **AND** identifies how to roll back application/configuration changes while preserving the database

#### Scenario: Example configuration is reviewed or committed
- **WHEN** deployment assets are inspected
- **THEN** they contain no real secret, contact credential, private key, hostname, or personal data

### Requirement: Production hardening is verifiable offline where feasible

Application behavior and static deployment assets SHALL have offline validation that does not contact certificate authorities, DNS providers, package registries, or public hosts. Tests SHALL cover production configuration failures, trusted-host enforcement, secure cookie attributes, proxy metadata acceptance and refusal, request-size rejection, and security headers; deployment documentation SHALL distinguish these checks from host-level checks requiring installed system services.

#### Scenario: Offline application tests run
- **WHEN** the project test suite runs with synthetic configuration and no network access
- **THEN** it verifies the production security behaviors without reading real deployment secrets or `config.json`

#### Scenario: Deployment assets are checked offline
- **WHEN** the documented local validation commands run on a host with the relevant binaries already installed
- **THEN** service, WSGI, and proxy configuration syntax can be checked without issuing a certificate or making an outbound request

#### Scenario: Full end-to-end HTTPS validation is requested
- **WHEN** the operator validates automatic certificate issuance and public reachability
- **THEN** the guide marks those checks as deployment-time network operations rather than offline automated tests
