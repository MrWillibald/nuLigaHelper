## Purpose

Ensures every public visitor can reach current operator-supplied and legally reviewed Impressum and privacy information without requiring an account.

## ADDED Requirements

### Requirement: Public legal pages use supplied and reviewed content

The system SHALL publish a German `Impressum` and `Datenschutzerklärung` using content supplied by the operator and approved by the operator's appropriate legal/privacy reviewer. Repository examples or templates MAY define structure and required placeholders, but SHALL NOT invent names, addresses, contact details, legal bases, retention periods, processor claims or other legal assertions. Production launch SHALL fail its checklist gate while content is missing, still contains placeholders, lacks recorded approval, or conflicts with the enabled production integrations and approved retention schedule.

#### Scenario: Approved content is prepared

- **WHEN** the operator provides versioned content and records the required review approval
- **THEN** the production pages render that supplied content without silently rewriting its legal meaning
- **AND** the launch evidence identifies the reviewed content version

#### Scenario: Content is incomplete or unapproved

- **WHEN** either page is absent, contains unresolved placeholders, or lacks required review approval
- **THEN** the launch checklist remains a no-go
- **AND** no generated legal text is substituted automatically

#### Scenario: Production processing changes

- **WHEN** a release changes data categories, purposes, retention, public operator details or enabled third-party processors
- **THEN** legal/privacy content review is reopened before that release reaches production

### Requirement: Legal information is public and consistently reachable

Both legal pages SHALL be accessible without authentication and without disclosing roster, contact or account data. Every public-facing HTML page, including schedule, login, registration, registration-status and error pages, SHALL provide a consistent link to both pages in the shared navigation or footer. Access SHALL remain available to guests and to verified-but-unapproved registrations.

#### Scenario: Guest follows legal links

- **WHEN** an unauthenticated visitor follows the `Impressum` or `Datenschutzerklärung` link from any public-facing page
- **THEN** the requested page is returned without login
- **AND** it contains no private roster payload, person identifiers or contact data introduced by the legal-page rendering

#### Scenario: Authentication page is displayed

- **WHEN** login, registration or registration status is rendered
- **THEN** links to both public legal pages are present and usable

#### Scenario: Error page is displayed

- **WHEN** the application renders a public error page through its normal HTML error handling
- **THEN** the shared legal links remain available unless doing so would mask a lower-level outage

### Requirement: Legal pages are usable and deployable without application-code edits

Legal content SHALL use the application's established responsive and accessible presentation, with meaningful German headings and links that remain keyboard operable and readable on narrow screens. The deployment approach SHALL keep operator-specific legal text separate enough from application logic that an approved content correction can be deployed through the documented release/configuration process without changing authentication or schedule behavior. Content changes SHALL remain versioned or otherwise auditable with approval evidence and SHALL not expose secret configuration values.

#### Scenario: Legal page is viewed on a narrow screen

- **WHEN** a visitor opens either legal page on a narrow viewport
- **THEN** headings, paragraphs, lists and links remain readable without requiring horizontal scrolling for ordinary text

#### Scenario: Approved correction is released

- **WHEN** the operator supplies and approves a correction to legal content
- **THEN** the documented deployment process can replace the public content without modifying account or schedule data
- **AND** the previous and new content versions and approval evidence can be identified

### Requirement: Privacy information matches operational reality

The published `Datenschutzerklärung` SHALL be checked at launch against the approved data inventory, retention schedule, backup policy and enabled third-party processor inventory. The check SHALL be performed by the operator and appropriate reviewer rather than inferred by application code. Technical documentation SHALL provide the factual inventory needed for that review but SHALL not claim that technical validation constitutes legal advice or legal approval.

#### Scenario: Privacy launch review is completed

- **WHEN** the privacy page is considered ready for public launch
- **THEN** the review evidence maps its supplied statements to the current operational data inventory, retention rules, backup handling and enabled processors
- **AND** discrepancies block launch until the supplied content or production configuration is corrected and re-reviewed

#### Scenario: Technical checks pass without legal approval

- **WHEN** routes, links and rendering pass technical validation but required content approval is absent
- **THEN** technical success does not satisfy the legal/privacy launch gate
