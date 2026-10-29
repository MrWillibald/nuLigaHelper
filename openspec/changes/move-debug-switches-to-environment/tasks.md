# Tasks

## 1. Overrides from the environment

- [ ] 1.1 Replace `DEBUG_FLAG`, `CHANGE_DAY` and `DEBUG_TODAY` in `common.py` with call-time parsing of `NULIGAHELPER_TODAY` inside `effective_today()` and a `dry_run()` helper for `NULIGAHELPER_DRY_RUN`; verify invalid values raise a German error and unset values keep today's behavior (Europe/Berlin date in production, local date otherwise).
- [ ] 1.2 Change `Notifier.send_Mail` / `send_SMS` to call `common.dry_run()` at send time and remove the import-time `from common import DEBUG_FLAG`; verify dispatch counts are unchanged in dry-run mode and providers are never constructed.
- [ ] 1.3 Add one shared guard (e.g. `common.refuse_overrides_in_production()`) called by `webapp.create_app()`, `main.main()` and `operations.py preflight|launch-check`, replacing the `common.DEBUG_FLAG is False and common.CHANGE_DAY is False` check; verify each entry point refuses before opening the database or contacting providers.

## 2. Tests and documentation

- [ ] 2.1 Migrate tests that assign `common.DEBUG_*` or patch `notifier.DEBUG_FLAG` to environment patching; add regressions for valid/invalid values, independent use of both variables and the production refusal of every entry point.
- [ ] 2.2 Document both variables in `README.MD`, `AGENTS.md` and `CLAUDE.md`, remove the "reset both to False before committing" rule and describe how to reproduce the old `DEBUG_FLAG` behavior.
- [ ] 2.3 Run `test/run_tests.sh` and `openspec validate move-debug-switches-to-environment --strict`.
