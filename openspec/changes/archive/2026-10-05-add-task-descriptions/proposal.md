# Proposal

## Why

Helpers currently see task names without an explanation of the work involved. Descriptions beside every task will help them understand a duty before volunteering.

## What Changes

- Add a concise German description for every supported game and day-level task.
- Show a small information icon immediately after each task label on the schedule.
- Make descriptions available on pointer hover and keyboard focus, and on click or tap for touch users.
- Float descriptions over the underlying schedule without pushing other page content down.
- Reuse the same description across numbered positions of one semantic role.
- Make descriptions available to guests as well as signed-in viewers while preserving existing schedule privacy.
- Maintain the description catalog in Python role metadata without requiring the separate message-template change to be implemented first.

## Capabilities

### New Capabilities

- `task-descriptions`: Defines complete task-description coverage and accessible information controls beside schedule task labels.

### Modified Capabilities

## Impact

Developer-maintained role metadata, schedule view construction and templates, shared information-control styling and behavior, accessibility and narrow-screen checks, catalog-coverage tests and README documentation. No database migration or new external dependency is needed.
