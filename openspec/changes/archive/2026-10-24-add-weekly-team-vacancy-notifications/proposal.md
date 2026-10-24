# Proposal

## Why

The existing one-week reminders reach assigned helpers, while unfilled game duties are only followed up with the responsible team's MV on the day before the game. Active team members who could help need an earlier, friendly invitation that identifies the remaining work without asking already assigned volunteers to take another duty in the same game.

## What Changes

- Extend the existing daily job's seven-day reminder window with a personal recruitment message for games that still have physically vacant required positions.
- Contact active members of the saved responsible team, including its MV when applicable, but exclude everyone holding any assignment in that game. Assignments in other games or day blocks do not exclude a recipient.
- List missing task labels and vacancy counts using the shared category-dependent required positions. Replacement searches and age deficiencies alone do not trigger this invitation.
- Provide polite German e-mail and SMS variants with a greeting, friendly reminder, invitation to take a suitable duty in the Heimspielplan, thanks and sign-off. Use a Spielfest variant without invented opponents.
- Reuse ordinary e-mail-first/SMS-fallback delivery and skipped-contact reporting. Keep existing helper reminders and day-before MV follow-up.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `game-task-staffing`: Define the seven-day vacancy invitation, responsible-team recipient selection and same-game assignment exclusion.
- `message-templates`: Define the centralized German invitation content, vacancy counts and explicit ordinary-game/Spielfest channel variants.

## Impact

- `notifier.py`: Vacancy-based recipient selection and dispatch through existing message rendering and transport helpers.
- `main.py`: An additional call in the existing seven-day notification branch.
- `messages.py`: Validated templates, named fields and vacancy-list fragments.
- Offline notifier, message-catalog and daily-job tests; `README.MD` notification documentation.
- No new provider, dependency, schema revision, public API or configuration setting is required.

## Non-goals

- No recruitment for team-independent preparation, cleanup or cake blocks.
- No replacement broadcast, automatic assignment, eligibility override or change to binding duties and cutoff rules.
- No daily escalation, cross-game message digest, missed-run catch-up or persistent delivery-deduplication mechanism. The new dispatch follows the existing daily reminder execution model.
