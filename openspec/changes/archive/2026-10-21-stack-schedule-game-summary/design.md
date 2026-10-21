# Design

## Context

See `proposal.md` for motivation and `specs/schedule-overview/spec.md` for observable behavior. `templates/schedule.html` already renders separate `.game-time`, `.game-teams`, `.game-meta` and `.game-responsible` elements. Preparation, cleanup and cake-delivery cards reuse the first three classes. The shared `.game-info` flex layout currently allows time and title to share a row; `.game-responsible` and task-block metadata also have flex rules that reserve room alongside other content. The summary's outer grid independently arranges information and staffing, with an absolutely positioned "Details" caption.

An unresolved-classification warning currently precedes the responsible team in markup. Day-task time elements can contain a date qualifier, and cake time/quantity elements are updated using existing data attributes. These details require a small layout decision before implementation.

## Goals / Non-Goals

**Goals:**

- Use the existing shared information structure to produce a consistent vertical hierarchy for games and every day-task type.
- Keep DOM reading order aligned with visual order and retain wrapping, native disclosure and existing dynamic update hooks.

**Non-Goals:**

- Changes to staffing calculations, game identity, assignment actions, access rights, filtering or block configuration behavior.
- Fixed card heights, truncated names or a fabricated responsible team for blocks.

## Decisions

### Stack the information area, retain the outer summary layout

Use a single-column layout for `.game-info`, with each primary content element on its own row. Remove or adapt obsolete flex sizing on its children so the time, title, metadata and responsible-team elements cannot share a row. Both game and day-task macros use this same rule; retain the outer `.game-summary` and `.coverage` responsive layouts. Keep the existing "Details" position and reserve enough room for it.

Alternatives considered: assigning full flex bases to each child can also create rows but preserves unnecessary interactions with existing flex overrides. Separate game-only selectors would fail to include day-task cards. Adding line-break characters would make content and layout harder to maintain.

### Use a modest shared time-size increase

Start with `1.25rem` for `.game-time`, up from `1.1rem`, retaining the existing strong weight and color treatment. This is a minor implementation assumption for "etwas größer"; adjust modestly during visual verification while keeping the time larger than its current size and than the title. Date qualifiers remain smaller supporting text within the first information row. Permit time placeholders and qualifiers to wrap within that row on narrow screens.

Alternatives considered: a much larger heading would dominate compact summaries; a font-weight-only change would not satisfy the requested larger text.

### Preserve semantic rows and supplemental states

Keep matchup/Spielfest content, color markers, task icons, metadata wording and all cake update attributes. Place any classification warning after the four primary game rows so it remains supplemental without interrupting the requested sequence. Eligibility warnings remain in the existing staffing status area. Blocks render only their three applicable rows, including unset time and cake configuration states, without an empty fourth row.

Alternatives considered: hiding warnings would lose actionable information; using CSS ordering alone would risk divergence between reading and visual order; introducing a block responsible-team placeholder would misrepresent the domain.

## Risks / Trade-offs

- [Cards become taller] → Accept the height required by distinct rows; retain small gaps and the existing compact staffing area.
- [Shared styles affect all block variants, including placeholders and date qualifiers] → Verify preparation, cleanup and cake cards with configured, zero and unset cake settings, including past dates.
- [Long text or "Details" overlaps at small widths] → Retain natural wrapping and minimum-width safeguards; verify 320, 360 and 390 CSS pixels, desktop, and enlarged text with long synthetic names and warnings.
- [Template edits disrupt live updates] → Preserve all existing data attributes and confirm that saved responsible-team and cake-setting updates still target the same elements.

## Migration Plan

Deploy the template and stylesheet with the normal application release; no schema or data migration is required. Rollback restores the prior template and stylesheet. Update the overview description in `README.MD` alongside the implementation.
