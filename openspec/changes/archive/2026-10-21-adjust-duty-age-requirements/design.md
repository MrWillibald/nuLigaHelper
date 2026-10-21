# Design

## Context

See `proposal.md` for motivation and the delta specs for behavior. Calendar calculations and game classification already live in `age_eligibility.py`. Its individual-duty helper currently covers only Zeitnehmer/Sekretär, while `sale_coverage()` always demands an adult. `db.claim_eligibility()` separately rejects invalid sale dates before checking the group. `db.staffing_status()` evaluates only the two timing roles and Verkauf. Web candidates, serialized assignment writers, statistics and MV follow-up consume those shared results.

The change crosses validation, staffing reads and schedule guidance, which warrants a design despite requiring no schema revision. Existing assignments are binding, and age deficiencies cannot authorize destructive cleanup or a cutoff bypass.

## Goals / Non-Goals

**Goals:** Keep one policy for candidates, confirmed writes and staffing reporting; preserve transactional checks and public-safe deficiency payloads; apply the revised rules to current/future staffing without data rewrites.

**Non-Goals:** Changing registration birth-date requirements, game classification, role capacity, day-block age rules, authorization, replacement lifecycle, locking, notification routing or historical assignment/audit records. This change introduces no legal or federation-rule lookup; it implements the club policy requested here.

## Decisions

### 1. Extend shared individual-duty evaluation

Generalize the existing individual eligibility helper to cover Ordnungsdienst as well as timing roles, updating its name and callers if needed for clarity. Ordnungsdienst selects a category-independent threshold of 18; timing roles retain explicit category resolution, with both adult timing thresholds set to 16 and youth thresholds retained at 14. Reuse completed-calendar-year, birthday, leap-day and public deficiency behavior.

Update both hypothetical claims and staffing evaluation to use this policy, including existing Ordnungsdienst occupants. A web-only filter was considered and rejected because CLI, direct domain writes and existing staffing would disagree. Do not duplicate calendar arithmetic or add an independent age check in the browser.

### 2. Make recognized youth the explicit sale exemption

Use the existing classifier: only `YOUTH` (including SPF) exempts Verkauf from age checks. `ADULT` and `UNKNOWN` retain the existing group rule. Keeping unknown categories unchanged limits the relaxation to the requested youth scope; classifying unknown as youth could silently exempt adult games.

Have shared sale coverage return no age deficiency for youth before any date or birth-date evaluation. Update the early invalid-date guard in `db.claim_eligibility()` consistently: youth-sale age checking must not retain a hidden date-based age restriction. General assignment/date/cutoff guards remain authoritative and unchanged. Preserve non-youth invalid-date refusals even for the first seller, and permit incomplete non-youth groups without individual minima as today.

### 3. Preserve serialized claims and resulting-group takeovers

All writers continue through existing claim/takeover helpers, refreshing authoritative game, person and group state under serialization and after retries. Replacing a seller evaluates the resulting group without the outgoing helper. Youth takeovers may remove the only adult; non-youth takeovers may not leave a full nonqualifying group. A stale youth candidate must fail when a concurrent saved category change makes the final group subject to adult coverage.

Ordnungsdienst candidates and writes use the same minimum. Preserve current-occupant visibility, compare-and-swap occupant/request-token checks, explicit confirmation, one-duty uniqueness and fixed-duty protection. Failed writes preserve occupants, open requests and successful-mutation audits.

### 4. Reevaluate staffing on reads with existing reporting contracts

Add occupied Ordnungsdienst positions to individual deficiency evaluation. Use category-aware sale coverage in the same staffing summary. This automatically propagates new Ordnungsdienst deficiencies and resolved youth-sale/adult-Zeitnehmer deficiencies to schedule, statistics and existing MV follow-up without storing derived ages or a policy migration.

Do not change occupancy counts or fabricate audit transitions for policy reevaluation. Saved locked assignments and replacement tokens remain untouched. Reads remain free of assignment/request/lock writes. Reuse existing public-safe reason codes and threshold fields; expose no full birth dates or exact personal ages.

### 5. Align guidance and verification with the shared policy

Update the German schedule notice and the task descriptions in `db.py`: Zeitnehmer currently names 18 for adult games, Verkauf currently requires an adult without a category exception, and Ordnungsdienst needs its new minimum stated. Update `README.MD` and the repository's documented age rules in `AGENTS.md` during implementation. Existing generic reason rendering can display the new role/minimum without browser-side thresholds. Verify candidates and authoritative responses refresh together after claims and takeovers.

Use synthetic offline tests for 16/18 birthday boundaries, leap-day handling, missing dates, adult/youth/SPF/unknown categories, existing occupants, all mutation paths and category changes during contention. Adapt existing youth-sale tests to the exemption while keeping their non-youth coverage and race guarantees. Reusable fixture helpers may currently assign young people to unrestricted Ordnungsdienst; update such fixtures consciously rather than weakening production checks.

## Risks / Trade-offs

- Existing young or unknown-date Ordnungsdienst occupants become deficient → retain their duties and locks, report the requirement through existing staffing/MV paths, and permit only existing authorized resolution before cutoff.
- One forgotten sale-date guard or reporting call keeps youth restrictions alive → cover both final-slot claims and staffing reports, including SPF and legacy unknown birth dates.
- Relaxing sale checks for unknown categories unintentionally broadens scope → test GE, absent and unsupported categories explicitly and retain current behavior.
- Stale candidates or cached game/person data bypass rules → exercise claim/takeover revalidation after concurrent corrections and retries.
- Changed thresholds invalidate old test assumptions → use explicit synthetic birth dates and non-youth fixtures for tests intended to verify adult sale coverage.

## Migration Plan

No Alembic revision or database rewrite is required. Deploy the updated application consistently to web and cron/CLI processes using the existing operational workflow. On subsequent reads, current/future staffing reflects the new rules; new confirmed writes use them immediately. Do not emit assignment audits or notifications merely because the policy was deployed.

Rolling back the application restores the previous eligibility policy for subsequent reads and writes while preserving all assignments created under the revised policy. Such retained assignments may then display deficiencies under the old rules; rollback must not remove them or modify historical audits.
