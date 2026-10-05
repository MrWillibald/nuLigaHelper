# Design

## Context

Schedule task labels come from game roles and day-block phases and are rendered beside assignment controls or guest-visible names. Repeated positions such as Verkauf and preparation share a semantic role. The UI currently has no task-help control. See proposal.md for motivation.

## Goals / Non-Goals

**Goals:** Complete description coverage with one text per semantic role; support hover, keyboard and touch in the existing card layout; preserve guest privacy and assignment interactions.

**Non-Goals:** An administrator description editor, a database content model, externally hosted help or changes to assignment rules. This proposal does not depend on the message-template proposal.

## Decisions

### Keep one developer-maintained Python role-description catalog

Use role/phase identifiers rather than numbered UI labels as description keys. Store concise German guidance in Python role metadata and expose resolved descriptions through the schedule view. Avoid duplicated template/JavaScript texts. This can later share a role metadata module with other changes, but remains independently applicable.

Describe supported tasks at the current revision, including Kasse, Reinigung and Kuchenlieferung. Keep coverage complete when roles are introduced or renamed, without adding proposal-only roles or obsolete Unterstützung labels. Coverage verification compares the actual supported role catalog against descriptions.

### Use confirmed club task definitions

The description catalog uses the club's confirmed task definitions below. Timing and
training guidance belongs in the descriptions; it does not introduce new assignment
checks or change notification schedules.

**Zeitnehmer (confirmed):**

- Minimum age is 14 for youth games and 18 for adult games, on the scheduled game date.
- Arrive at least one hour before the scheduled game start.
- Participate in the technical meeting with the referees at least 30 minutes before
  the game and directly after the game.
- Be able to operate the game clock and scoreboard.
- Handle two-minute suspensions and watch for correct player substitutions during the game.
- No official timekeeper instruction or training is required.

**Sekretär (confirmed):**

- Minimum age is 14 for youth games and 16 for adult games, on the scheduled game date.
- Arrive at least one hour before the scheduled game start.
- Participate in the technical meeting with the referees at least 30 minutes before
  the game and directly after the game.
- Be able to operate nuScore, the electronic game report tool.
- Watch for correct player substitutions during the game.
- No official secretary instruction or training is required.

**Verkauf (confirmed):**

- Arrive at least one hour before the scheduled game start.
- At least one assigned seller must be 18 or older on the scheduled game date;
  this is a collective requirement rather than an individual minimum for every seller.
- Help prepare food and beverages for the snack bar and sell them during the game.
- Both numbered sale positions use the same description.

**Ordnungsdienst (confirmed):**

- Arrive at least one hour before the scheduled game start.
- Participate in the technical meeting with the referees at least 30 minutes before
  the game and directly after the game.
- Ensure the game takes place safely.
- Remove spectators from the game site if they behave improperly or insult players
  or referees.
- Remain visibly identifiable during the game through a badge or vest.

**Kasse (confirmed):**

- Arrive at least one hour before the scheduled game start.
- Sell spectator tickets for adult games at the entrance.
- Assist Ordnungsdienst as the second steward (Ordner).
- Participation in the technical meetings with the referees is not required.

**Reinigung (confirmed):**

- Assist with cleaning individual spots of the playing court when needed during the game.
- Stay close to the court, but not on the substitution bench.
- Participation in the technical meetings with the referees is not required;
  the referees note the cleaning helpers' presence.
- This is an in-game duty, distinct from the day-wide cleanup block.
- Both numbered cleaning positions use the same description.

**Vorbereitung (confirmed):**

- Arrive at the game site at the displayed preparation time.
- Help prepare the playing court and snack bar.
- All numbered preparation positions use the same description.

**Aufräumen (confirmed):**

- Arrive at the game site at the displayed cleanup time.
- Help dismantle the snack bar and clean the playing court.
- All numbered cleanup positions use the same description.

**Kuchenlieferung (confirmed):**

- Bring one cake to the snack bar at the displayed delivery time.
- Every numbered cake-delivery position uses the same description.

### Place an accessible information control after the label

Use a shared render helper for a small information control immediately after each assignment-task label. Keep the control outside the assignment select's label activation behavior so opening help cannot focus/change a select or submit a form. Associate the control with its description and a task-specific accessible name.

Hover/focus exposes the description in a floating popover over the underlying page; click/tap toggles persistent opening. Showing or dismissing help must not resize the assignment field, card or page, or move the other task controls. Position the popover beside its information control, choose above or below according to available space, and keep it within the viewport. Escape dismisses an opened popover without losing the information control's keyboard position. Native title-only tooltips were rejected because touch access, dismissal and assistive behavior are inconsistent. An always-visible paragraph would make every expanded card needlessly dense.

### Reuse layout and preserve public scope

Descriptions are generic task text, so all viewers can read them. Rendering must not serialize roster, contact, person identifiers or assignment data into help attributes. Use plain text escaping and responsive wrapping rather than HTML authored in task text. The floating popover may cover underlying content while open; dismissing it restores access to that content. Keep its text readable within the viewport, avoid clipping by card containers, and retain dismissal and keyboard/touch access on narrow screens. Bound long descriptions to the available height and allow scrolling within the bubble. Repositioning after page scroll or resize must preserve its reading position.

## Risks / Trade-offs

- [Repeated roles diverge in wording] -> Resolve by semantic role and verify numbered positions share the same description.
- [Help activation changes an assignment or toggles the card] -> Isolate the information control and verify interaction behavior.
- [Hover-only help excludes touch/keyboard users] -> Support focus, click/tap and keyboard dismissal.
- [Task catalog evolves after adoption] -> Make description coverage part of role-addition/rename validation.

## Migration Plan

No database migration is needed. Deploy the description catalog and shared rendering/interaction assets together, and verify current task coverage on desktop and narrow screens. Rollback restores the prior application assets.
