## Context

See `proposal.md` for the motivation and `specs/access-control/spec.md` for the required behavior. `templates/persons.html` currently renders the roster before one tier-gated management block. The page route already supplies the same roster, pending registrations, and team options to each viewer according to their rights.

## Goals / Non-Goals

**Goals:**

- Change the document order of existing page sections while retaining the current forms, element IDs, tier checks, and roster filters.
- Keep the page title and flash messages at the top so action feedback remains easy to find.

**Non-Goals:**

- Change authorization, data selection, form submission, or visual styling beyond what the new order requires.

## Decisions

- Move the existing management divider and card grid together above the roster heading in the template. Keeping the existing admin/MV conditions on the complete block and its cards preserves control visibility. Reordering cards individually would change more behavior than requested.
- Keep the roster heading, filter form, and people grid together after management. Their relationship remains clear when a filter is active or returns no people. A CSS-only visual reorder was considered, but document order should match screen-reader and keyboard navigation order.
- Verify rendered section order and tier-specific visibility in `test/test_management_ui.py`; existing endpoint tests continue to cover authorization. Changing the route or access checks would add unnecessary risk.

## Risks / Trade-offs

- [Moving a large template block separates a dialog from its trigger or changes nesting] -> Move the full management block intact and check the rendered page and form behavior in focused tests.
- [A filtered roster appears below unchanged management controls] -> Assert this order for a filtered admin or MV view as required by the spec.

## Migration Plan

No data migration is needed. Deploy the template and test change together; reverting that template change restores the previous order.
