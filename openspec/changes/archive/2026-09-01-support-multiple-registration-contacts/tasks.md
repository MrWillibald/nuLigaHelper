## 1. Registration Validation And Persistence

- [x] 1.1 Update registration contact parsing to normalize e-mail and phone independently, apply the selected calling code to phone validation, require at least one valid contact, reject every supplied invalid contact, and verify focused unit tests cover empty, one-contact, both-contact, and invalid-contact submissions
- [x] 1.2 Update registration persistence and conflict handling so both canonical contacts are stored atomically, either contact conflict prevents creation or modification, and verify uniqueness and generic-response tests pass

## 2. Authentication And Notifications

- [x] 2.1 Verify or adjust login lookup and challenge delivery so an eligible person with both contacts can request and complete login through either route, and add tests for both independent paths and selected-route delivery
- [x] 2.2 Verify automatic notification dispatch remains e-mail-first with SMS fallback regardless of the authentication or registration route, and add or update tests for both-contact and phone-only recipients

## 3. Registration Form And Progressive Enhancement

- [x] 3.1 Reorder the registration markup to render e-mail and SMS fields before the route selector, keep both values visible, preserve field-specific errors, and verify server-rendered markup order and accessibility tests
- [x] 3.2 Update client-side route availability to enable only routes with valid non-empty values, keep invalid supplied fields blocking, and verify behavior with JavaScript plus usable server-side behavior without JavaScript
- [x] 3.3 Ensure styles remain responsive for the expanded contact section and verify the authentication-form tests and manual desktop/mobile check pass

## 4. Documentation And Integration Verification

- [x] 4.1 Update user-facing registration or authentication documentation to describe optional e-mail and SMS contacts, selected verification route, dual-route login, and e-mail-first automatic notifications, and verify documented behavior matches the implementation
- [x] 4.2 Run `test/run_tests.sh` and verify the complete offline suite passes without regressions
