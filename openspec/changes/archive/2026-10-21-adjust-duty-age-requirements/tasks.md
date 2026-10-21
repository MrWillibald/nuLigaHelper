# Tasks

## 1. Individual duty thresholds

- [x] 1.1 Extend shared individual eligibility to Ordnungsdienst at 18 for every category and lower adult Zeitnehmer to 16; connect claims, takeovers and staffing reporting to the same evaluator. Verify synthetic tests for exact birthdays, the preceding day, leap-day handling, unknown birth dates, invalid game dates and adult/youth/SPF/unknown categories; retain existing youth timing and adult Sekretär thresholds.
- [x] 1.2 Cover the revised individual rules across member/MV/admin web paths, non-JavaScript confirmation, direct domain helpers and CLI assignments; verify stale person/game corrections are revalidated, underage Ordnungsdienst cannot bypass the policy and rejected writes preserve occupants, requests and successful audit history. Update synthetic fixtures where the new Ordnungsdienst rule makes existing setup invalid and run the affected test files.
- [x] 1.3 Update the individual duty table and explanations in `README.MD` and the age-rule sections of `AGENTS.md`; verify the documented values agree with passing threshold tests and retain calendar, privacy and binding-duty guarantees.

## 2. Youth-sale exemption and atomic group checks

- [x] 2.1 Exempt recognized youth/SPF Verkauf in both shared coverage evaluation and the early claim date guard, preserving non-youth collective coverage and invalid-date refusals. Verify synthetic first/final-slot tests for younger and unknown-date sellers, no youth-sale age deficiency even with an invalid game date, and retained adult/GE/missing/unsupported-category behavior; keep general assignment date guards unchanged.
- [x] 2.2 Update takeover and concurrency regression coverage to prove youth-sale transfers may remove the only adult, while adult/unresolved-category transfers validate the resulting group without the outgoing seller. Verify claims/takeovers revalidate category and group changes under serialization and after retries, concurrent non-youth younger sellers cannot bypass coverage, and stale refusals preserve saved request tokens and responsibilities.
- [x] 2.3 Document the youth/SPF exemption, retained adult/unresolved-category coverage and resulting-group takeover rules in `README.MD` and `AGENTS.md`; verify there is no remaining claim that all sale groups require an adult.

## 3. Staffing feedback and German guidance

- [x] 3.1 Update the German schedule age notice and the Zeitnehmer, Verkauf and Ordnungsdienst task descriptions in `db.py` to the revised policy. Verify task-description tests, schedule rendering and loaded candidate responses agree on Ordnungsdienst 18, adult Zeitnehmer 16 and unrestricted youth/SPF Verkauf, including the non-JavaScript path and saved-state refresh after claims/takeovers.
- [x] 3.2 Add reporting regressions for retained underage/unknown-date Ordnungsdienst and newly compliant adult Zeitnehmer/youth-sale staffing. Verify schedule cards, statistics and existing MV follow-up share the expected deficiencies, full rings stay physically accurate, category changes reevaluate coverage, locked commitments/request state/history remain intact and responses contain no full birth dates or exact personal ages.

## 4. Integration verification

- [x] 4.1 Run `test/run_tests.sh` and `openspec validate adjust-duty-age-requirements --strict`; verify both pass with synthetic offline data and resolve any regression without weakening authorization, confirmation, replacement-token, cutoff, uniqueness or day-block rules.
