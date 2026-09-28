## 1. Reorder the Person-Management Page

- [x] 1.1 Move the existing tier-gated team-management section above the roster heading, filters, and person entries in `templates/persons.html`; verify the rendered admin and MV pages show the available controls first, while the member page shows the roster without that section.

## 2. Verify Presentation and Rights

- [x] 2.1 Add focused assertions to `test/test_management_ui.py` for admin, MV, member, and filtered-roster section order and control visibility; verify the file passes with `./venv/bin/python test/test_management_ui.py` or the available pytest command.
- [x] 2.2 Run `test/run_tests.sh` and verify the complete offline suite passes, including existing authorization and contact-privacy tests.
