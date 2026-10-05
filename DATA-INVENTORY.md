# Application data inventory

This records application storage and visibility for operator and club review.
The private deployment supplies the public privacy text and retention policy.

| Data | Stored purpose and location | Application visibility |
| --- | --- | --- |
| Person identity, display name, account status and team memberships | Stable `persons.id`, mutable name/account fields and `person_teams`; roster, authorization and staffing | Assigned names on the public schedule; authorized signed-in roster; account administration according to tier |
| E-mail address and E.164 phone number | Person contacts; selected-route authentication and ordinary e-mail/SMS notifications | Own/admin person maintenance; absent from schedule and candidate responses |
| Date-only birth date | `persons.birth_date`; eligibility on the scheduled game date, with no persisted age | Only the person themselves and admins in authorized maintenance. MV creation accepts a new person's date but gives no continuing access to it |
| Game and assignment records | Schedule, responsible team, task/slot and person links; reminders and statistics | Public assigned names and schedule details; assignment controls follow existing tier rights |
| Assignment audit snapshots | Append-only actor, person name, duty, game and change history | Admin audit view; no birth dates or exact calculated ages |
| Authentication and abuse records | Purpose-bound expiring challenges and bounded abuse controls | Internal use; routine cleanup reports aggregate counts |

New registration, admin/MV creation and CLI creation require a valid birth date.
The guarded Alembic migration leaves existing dates unknown and retains existing
identities, contacts, memberships, status, MV appointments, assignments and audits.
Unknown legacy dates may be completed by self/admin maintenance or the identity-safe
operator command; a known date may be corrected but cannot be cleared. Unrelated
maintenance remains possible while a legacy date is unknown.

Schedule cards, candidate payloads, statistics and ordinary notifications contain
only necessary duty thresholds or derived eligibility reasons, never full dates or
exact personal ages. Diagnostic logs and assignment audit snapshots exclude them.
Birth dates are self-reported and are not identity-document verification.

Database snapshots, local recovery material and Dropbox database backups include
stored birth dates. Apply the existing restricted access and approved backup
retention policy to those copies. Birth dates follow the containing person record's
existing retention behavior: bounded cleanup may delete eligible unapproved records;
active/inactive approved people and protected audit/assignment references remain.
Backups are not edited individually. Public notice updates belong in the private
deployment's legal content; no provider notification body needs the full date.

See [schema upgrade and recovery](README.MD#initialize-or-migrate-the-application-schema)
and [deployment procedures](DEPLOYMENT.md).
