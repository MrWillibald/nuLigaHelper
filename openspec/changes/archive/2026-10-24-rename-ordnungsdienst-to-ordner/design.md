# Design

## Context

See `proposal.md` for motivation and the delta specs for the observable contract. The application currently uses `db.ROLE_SECURITY = "Ordnungsdienst"` both as an assignment key and as human-facing text. `age_eligibility.py` also compares that literal directly. Assignment writers, candidate lookup keys and browser request payloads depend on it. A constant-only rename would orphan existing assignments from role lookups and could bypass eligibility unless accompanied by a guarded migration.

There are already presentation seams: schedule slots have separate `label` and `role` fields, and `db.position_label()` formats position names. Other paths still render raw roles: vacancy dictionaries, statistics, notification context, assignment-conflict errors and CLI output. Audit entries have a separate `role` text column intended to retain the label in force at the transition, and the audit view renders it literally.

## Goals / Non-Goals

**Goals:** Provide one current display name while retaining the established persisted and transport identities. Make visible text consistent across initial rendering, refreshes, confirmations, reports and messages. Preserve historical evidence literally.

**Non-Goals:** Redesign the role model, rename every stored role, add a schema revision, rewrite archived OpenSpec changes or historical migration scripts, alter duties or change assignment policy.

## Decisions

### Separate display labels from stable assignment keys

Retain the assignment constant and stored role value `Ordnungsdienst`. Add a small shared role-display helper mapping it to `Ordner`, with unchanged fallback for other roles. Keep this helper independent of the database so both `db.py` and `age_eligibility.py` can use it without a circular import. `db.position_label()` must apply the display helper before any required numbering; Ordner remains unnumbered.

An assignment-role rename plus Alembic migration was considered. It adds deployment and stale-client compatibility work for a wording change, so the chosen approach retains the established key. A schedule-only hardcoded label was also considered but leaves notifications, CLI and eligibility text inconsistent.

### Apply labels only where text becomes human-readable

Use the helper for schedule slot labels, generic task guidance and accessible names, confirmation labels, statistics presentation, rendered vacancy lists and assignment-conflict text. Update literal guidance in `templates/schedule.html` and the Kasse task description to use correct German grammar, such as “den Ordner als zweiter Ordner unterstützen”. Keep the Ordner task's existing work and minimum age unchanged.

Structured `role` fields, candidate keys, mutation bodies, HTML form values and `data-role` attributes retain the old key. Preserve existing JSON structures and map vacancy keys at their display boundary rather than changing the dictionary's transport contract. JavaScript must use the supplied current label for visible confirmation, loading, error or accessibility text instead of falling back to an untranslated role. Eligibility reason dictionaries retain their machine identifiers while the human-readable `message` names Ordner. Do not expose birth dates or exact ages.

Map labels before supplying `task_label` values and missing-duty/deficiency fragments to the validated renderer in `messages.py`. Cover assigned-helper reminders, rescheduling, MV follow-up and weekly team vacancy invitations, including ordinary-game and Spielfest variants. Preserve message field contracts, recipients, trigger dates, routing and dispatch counts; do not introduce caller-local notification templates.

### Keep CLI wording and role input usable

Make `Ordner` the documented and displayed CLI role spelling. Normalize it to the existing internal key before any role-taking command invokes the domain helpers; retain `Ordnungsdienst` as an input alias for existing scripts. Both inputs share the same validation and mutation paths. Render the current label in command results and user-facing role errors without rewriting historical audit output. This human-input alias does not add a new API role identifier.

### Capture labels when new audit evidence is appended

For new game transitions, write the current display name to `AssignmentAudit.role`, including claim, replacement opening/closure, takeover, first lock and any other legitimate transition that records this role. Cover direct audit constructors as well as the common audit helper. Assignment rows retain their internal role key. Existing audit rows are untouched and the audit template continues rendering each recorded label literally; a global display mapping in that template would falsify the historical name.

No audit column or marker is needed: the existing audit role column already stores the historical human-readable role. Game/person filters do not depend on joining that text back to the assignment role.

### Verify behavior through existing offline coverage

Extend existing tests at the observable boundaries rather than creating a general role-framework test suite. Use synthetic databases containing old-key assignments and old audit labels, with both locked and replacement-search state. Verify the same identities and state survive schedule and report reads, the minimum age remains 18, confirmed writes still enforce current state and cutoff, cancellation creates no mutation, and new audit labels coexist with old ones.

Check initial and refreshed schedule/eligibility text, non-JavaScript confirmation, CLI aliases/output and captured notification bodies without sending real messages. Assertions must distinguish visible labels from intentionally retained transport values; a global ban on the string Ordnungsdienst would reject required compatibility data.

## Risks / Trade-offs

- Raw roles are rendered in several indirect paths → audit role-to-text boundaries, including vacancy fragments, statistics and generic conflict errors, and verify actual rendered outputs.
- Display text differs from an internal German key → document the compatibility decision and use one shared helper with no duplicate per-caller mapping.
- A display-only check accidentally weakens the security age rule → retain the existing comparison key and exercise underage, unknown-date and eligible claims/takeovers through existing tests.
- Audit display normalization could rewrite historical meaning → map at audit creation, never at historical rendering or through a bulk update.
- Changing vacancy map keys could break candidate refreshes or clients → retain transport keys and format only visible labels.

## Migration Plan

Deploy through the existing application release procedure after offline verification. No database migration or startup rewrite is required. Existing assignments become visible as Ordner through label resolution and keep their IDs, occupants, positions, locks and request tokens.

Rollback restores the preceding application version without restoring or mutating the database for this change. New audit entries bearing Ordner remain readable as historical text under that version; earlier Ordnungsdienst entries also remain intact.
