# Access Control Delta

## ADDED Requirements

### Requirement: Only current administrators control shared date visibility and overview overrides

Only currently active administrators SHALL be able to change or clear the shared home-game-day boundary or activate the full-date overview mode on the schedule and outstanding statistics. Both permissions SHALL be checked server-side against the viewer's current account and derived rights on every request. Guests, members, MVs and verified but unapproved accounts SHALL NOT receive setting controls or an effective full-date override; a copied or forged override parameter SHALL NOT widen their overview results. An account that is both admin and MV SHALL retain the administrator capability. Statistics SHALL retain its existing session requirement.

Setting changes SHALL require explicit state-changing submission, an authenticated current administrator and a valid CSRF token. Missing expected-state input, stale state, missing/invalid CSRF or lost administrator rights SHALL leave the saved setting unchanged. Display GETs SHALL remain read-only. The boundary SHALL NOT change authorization for direct duty/candidate operations; those operations SHALL continue to enforce their existing independent rules.

#### Scenario: Lower tier forges a setting save

- **WHEN** a guest, member, MV or verified but unapproved registrant directly submits a setting change or clear request
- **THEN** it is refused or redirected under the existing authentication/account rules
- **AND** the saved boundary, assignments and audit history remain unchanged

#### Scenario: Lower tier copies an administrator view link

- **WHEN** a guest, member or MV supplies the administrator full-date override parameter
- **THEN** the schedule and any authorized statistics response retain the ordinary limited scope
- **AND** the response contains no administrator setting or full-date controls

#### Scenario: Administrator loses rights after loading the page

- **WHEN** an administrator's active status or administrator flag changes before a subsequent setting save or full-date request
- **THEN** the subsequent request uses the updated account rights
- **AND** no stale page or session permits an unauthorized save or overview override

#### Scenario: Administrator submits without CSRF

- **WHEN** an administrator submits a valid date with a missing or invalid CSRF token
- **THEN** the write is refused and the saved boundary remains unchanged

#### Scenario: Administrator is also an MV

- **WHEN** a current active administrator also manages a team
- **THEN** the combined account retains access to the setting and full-date views
- **AND** no new ability to change a locked occupied duty is granted
