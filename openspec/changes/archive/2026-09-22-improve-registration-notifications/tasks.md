## 1. Administrator Approval Notifications

- [x] 1.1 Extend authentication tests with multiple active administrators, an inactive administrator, and an administrator without contact data; verify a successfully contact-verified registration attempts delivery to every and only active administrator, includes the registrant and all selected teams, and leaves the registration verified when one delivery fails.
- [x] 1.2 Replace the single-approver lookup with stable iteration over all active administrators and personalized German e-mail/SMS copy directing them to "Helfer verwalten"; verify the focused registration tests pass for both normal code confirmation and the supported legacy verification path.

## 2. Approved-User Welcome Notification

- [x] 2.1 Add approval-flow tests covering successful approval, exact TuS Raubling Handball welcome wording, e-mail-first/SMS-fallback dispatch, rejected registration, repeated or stale approval, missing contact, and provider failure; verify only one successful verified-to-active transition attempts the welcome delivery and committed active state survives delivery failure.
- [x] 2.2 Send the welcome notification only after the approval transaction commits, with confirmation and an invitation to sign in and take open Heimspielplan duties; verify the focused approval and notifier tests pass.

## 3. Documentation and Verification

- [x] 3.1 Update `README.MD` to describe notification of all active administrators after contact verification and the welcome notification after approval; verify the documentation distinguishes contact verification from admin approval and retains the automatic e-mail-first/SMS-fallback rule.
- [x] 3.2 Run `test/run_tests.sh` and `openspec validate improve-registration-notifications --strict`; verify both complete successfully and committed debug switches remain disabled.
