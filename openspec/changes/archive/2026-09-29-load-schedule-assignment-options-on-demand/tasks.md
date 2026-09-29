## 1. Server-side candidate reads

- [x] 1.1 Extract shared game and block slot editability and candidate-scope rules from schedule rendering; verify existing tier, past-date, and occupied-slot tests still pass.
- [x] 1.2 Add protected per-game and per-block candidate read endpoints with a contact-free, deduplicated roster and slot eligibility; verify direct requests for guest, pending, member, MV, and admin, including a changed MV record or responsible team.
- [x] 1.3 Limit candidate queries to the opened card and eager-load active memberships; verify a synthetic query-count check does not grow one query per roster person or task slot.

## 2. Compact overview and client loading

- [x] 2.1 Remove unassigned candidate options and roster construction from initial schedule rendering, while keeping current occupants and lightweight team metadata; verify response content and size with small and large synthetic rosters.
- [x] 2.2 Load and populate only an opened editable card, disable controls until completion, and show retryable German errors for failed or expired requests; verify browser-side expansion and failure tests.
- [x] 2.3 Preserve candidate grouping, labels, warnings, selected inactive occupants, and same-card exclusion through successful claim and release; verify game and block DOM regression tests plus the existing compare-and-swap tests.

## 3. Documentation and integration

- [x] 3.1 Update `README.MD` to describe on-demand candidate loading and verify the instructions match the implemented page behavior.
- [x] 3.2 Run `test/run_tests.sh` and inspect the final diff; verify the full offline suite passes and no schema or provider configuration changed.
