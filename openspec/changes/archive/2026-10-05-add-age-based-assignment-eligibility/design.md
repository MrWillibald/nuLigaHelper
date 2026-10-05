# Design

## Context

See `proposal.md` for motivation. `Person` has no birth-date field. Game dates remain scraped German strings, while person creation occurs through public registration, admin/MV forms, CLI commands, and shared database helpers.

Assignments use per-slot compare-and-swap operations with bounded SQLite contention retries. Verkauf has two slots. A person can hold only one task per game. Members may release their own assignments at any time before the game, and administrators may correct past games.

`missing_slots` currently describes physical vacancies only. Overview progress, statistics, and MV reminders consume that information, so eligibility must become a separate component of staffing status rather than a fabricated empty slot.

## Goals / Non-Goals

**Goals:**

- Centralize age decisions so web, CLI, candidate loading, statistics, and reminders agree.
- Apply individual and collective rules transactionally without broadening assignment authority.
- Introduce birth dates safely into an existing roster whose dates are unknown.
- Preserve withdrawal rights, past assignments, and append-only assignment audits.
- Expose eligibility decisions without publishing full birth dates or exact ages.

**Non-Goals:**

- Identity-document verification or a claim that self-reported birth dates have been independently verified.
- Age limits for Kasse, Ordnungsdienst, Reinigung, or day-task blocks; the user specified none.
- A generic administrator-configurable rule engine or age-rule overrides.
- Changing notification recipients, permitting MVs to edit another person's profile, or normalizing all scraped game dates into database DATE columns.

## Decisions

### Store birth dates as date-only values and require them for new people

Add a nullable `Person.birth_date` date column. Nullable storage exists to represent honestly unknown legacy dates; new-person and registration boundaries require a valid date. Store no derived age, since age depends on the game date.

Reject impossible dates and dates after `common.effective_today()`. Permit authorized correction and legacy completion, reject clearing an already known date, and permit unrelated maintenance of an unchanged legacy record whose date is still unknown. These defaults avoid blocking contact repairs or account maintenance during backfill.

Collect the field before a registration code is requested and preserve it through the existing signed registration flow and confirmation step. Admins and MVs supply it when creating a person. Existing own-profile/admin profile permissions govern subsequent correction; MV creation does not grant continuing access to another person's birth date.

An alternative was a non-null migration with a synthetic default. That would invent eligibility information and risk losing existing records. A stored age was rejected because it becomes stale and cannot represent future-game eligibility.

### Use completed calendar years on the game date

Calculate completed years using the date of birth and parsed scheduled game date. All minimums are inclusive: qualification starts on the relevant birthday.

The corrected user decision is that Verkauf requires at least one explicitly assigned seller who is **18 or older on the game date**, eligible from their **18th birthday**. The other seller has no individual age minimum introduced by this proposal. Timing-role thresholds remain those in the age-eligibility spec.

Use calendar dates rather than elapsed days divided by 365. The selected leap-day convention is that a 29 February birthday advances on 1 March in a non-leap year, consistent with month/day comparison. This is an implementation convention recorded for review, not a legal-age determination.

### Share an explicit game classification

Use one domain classifier shared with `update-game-duty-staffing`. Recognize adult M/F classes, youth mA–mE/wA–wE classes, and SPF as youth; parse age-group identifiers explicitly rather than borrowing the existing color heuristic. League prefixes do not change the underlying category.

GE and unsupported or missing classes remain unresolved. Refuse new Zeitnehmer/Sekretär claims whose adult/youth threshold cannot be determined and explain the unresolved classification. Other staffing behavior retains the shared classifier's documented fallback. Verkauf uses the same collective adult-coverage rule across game categories.

If a game's date cannot be parsed, refuse new age-restricted claims because age on the relevant date cannot be established. No current-day fallback may authorize such a claim.

### Enforce individual checks and Verkauf coverage in the database mutation boundary

Implement one eligibility evaluator returning structured reasons and required thresholds without putting full birth dates into schedule payloads. Apply it in shared claim helpers and every web/CLI path, after normal authorization and against freshly stored game, person, and assignment state.

Zeitnehmer and Sekretär need a known birth date and the relevant minimum. Administrators obey the same rule, including past-game corrections.

For Verkauf, evaluate the post-claim group:

- A first seller may be younger or have a legacy unknown birth date; the other slot remains open and adult coverage remains outstanding.
- Filling the last slot requires at least one assigned seller with a known birth date who is at least 18 on the game date.
- A younger or unknown-date seller may fill the second slot when the existing seller qualifies.
- Responsible-team membership, an MV appointment, or an unassigned adult in a candidate list never substitutes for an explicit Verkauf assignment.

Revalidate after obtaining write serialization and on every SQLite retry so two concurrent younger-person claims to different slots cannot jointly fill both slots. Recheck person dates and scheduled game data in that same authoritative transaction. Preserve the existing compare-and-swap and one-task-per-game constraints.

An alternative was requiring every seller to be an adult, which would incorrectly replace the requested group rule with an individual rule. Requiring an adult to be assigned first was also rejected because it unnecessarily restricts the existing incremental staffing workflow.

### Preserve releases and distinguish vacancy from eligibility deficiency

A normal authorized release remains allowed even when it removes the only adult seller. Mark adult coverage outstanding, refresh the remaining slot's candidates, and keep existing missing-duty follow-up active.

Introduce shared staffing status with two components: physical vacancies and age-eligibility deficiencies. A full legacy roster can therefore retain its actual occupancy count while visibly failing eligibility. Statistics and existing responsible-team MV reminder routing consider either component outstanding.

Do not invent an extra empty slot or count responsible-team members as assignees. Do not introduce a new fallback recipient when a game has no responsible team or active MV.

### Reevaluate affected assignments without destroying history

Migration does not clear assignments. Correcting a birth date or changing a game's date or age class reevaluates affected current/future assignments, retaining occupants and exposing deficiencies until staff repair them. Current occupants remain visible and releasable even if they are no longer valid candidates.

Keep historical assignment rows and audit snapshots unchanged. A later admin correction to a past game must pass the same age check on the historical game date. Ordinary staffing follow-up continues to concern current/future games.

Persisting an eligibility flag was rejected because it becomes stale after game changes. Evaluate from authoritative data and invalidate cached candidate/staffing results when inputs change.

### Protect birth-date visibility at response boundaries

Only administrators and the person themselves see an existing full birth date in authorized person-maintenance pages. MV creation accepts the new person's date but does not grant subsequent roster visibility of it.

Schedule, statistics, candidate payloads, notification bodies, and assignment-audit snapshots carry eligibility reasons or thresholds when needed, not birth dates or exact calculated ages. Normal diagnostic logs exclude raw dates. Follow existing default-deny authorization, CSRF protection, and profile ownership checks.

### Coordinate related proposals without imposing an implementation order

The staffing proposal owns role counts and game-category completeness, while this proposal adds age compliance to the shared staffing result. Keep physical progress truthful and apply both sets of rules when both changes are installed.

Add German eligibility feedback through the central message catalog if available; otherwise use a single local template definition that the message-template proposal can migrate later. Birth-date schema evolution must join the existing Alembic revision chain in implementation order rather than create a second head.

## Risks / Trade-offs

- [Unknown legacy dates prevent verification] → Preserve records, expose an admin backfill queue, refuse restricted individual claims, and never count an unknown-date seller as a proven adult.
- [Concurrent changes bypass collective coverage] → Revalidate authoritative group state inside the serialized assignment transaction and repeat after retries.
- [A game changes after staff were assigned] → Recompute eligibility and keep affected duties outstanding without silently discarding assignments or audits.
- [A self-reported correction is treated as proof of identity] → Describe eligibility as based on stored dates and leave independent identity verification outside scope.
- [Physical progress is mistaken for valid coverage] → Keep occupancy counts truthful and show a separate eligibility warning; use shared complete-staffing status in statistics and reminders.
- [Birth dates leak through lazy candidate loading or diagnostics] → Return derived reasons only and test every schedule/candidate tier and normal log/audit path.
- [Parallel feature migrations produce divergent heads] → Sequence revisions against the actual implementation head and verify schema drift and guarded upgrades.

## Migration Plan

1. Add the reviewed Alembic birth-date revision and model support; keep existing dates null and verify retained person identities, contacts, memberships, assignments, and audits.
2. Add validated collection/correction to all creation and maintenance entry points and update synthetic fixtures to provide explicit dates.
3. Add the shared game classifier and date/eligibility evaluator, then enforce claims and refresh candidates.
4. Integrate eligibility deficiencies with overview, statistics, and existing MV follow-up.
5. Verify guarded migration, legacy backfill, privacy exclusions, role boundaries, birthday cases, and concurrent Verkauf claims; run the full offline suite.
6. Deploy through the existing stopped-services, backup-first schema migration. Complete unknown roster dates through authorized maintenance and review current/future deficiencies.

Rollback restores the validated pre-migration backup while services are stopped and reinstalls the prior application version. No implicit startup migration or in-place destructive downgrade is introduced.
