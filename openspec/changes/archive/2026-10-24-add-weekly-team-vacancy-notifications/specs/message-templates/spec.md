# Message Templates Delta

## ADDED Requirements

### Requirement: Team vacancy invitations use friendly German channel variants

The authoritative message catalog SHALL provide explicit ordinary-game and Spielfest vacancy-invitation templates with named, documented values and e-mail and SMS variants. Each invitation SHALL personally greet the recipient, gently remind them of their team's responsibility next week, identify the responsible team, date, start time and event, list the currently vacant required task labels with counts, and politely invite them to review the Heimspielplan and take a suitable open duty. It SHALL thank the recipient and end with the established club sign-off: Viele Grüße / TuS Raubling Handball for e-mail and VG TuS Handball for SMS. E-mail SHALL include a descriptive subject containing the game date. SMS SHALL preserve the actionable information in shorter wording. The ordinary-game e-mail SHALL identify the age class and both opponents; the ordinary-game SMS SHALL identify at least the age class, date and start time. Spielfest variants SHALL identify the full Spielfest age-class label and SHALL NOT invent opponents.

Vacancy wording, including task/count fragments, SHALL come from the authoritative catalog and use current semantic task labels. It SHALL describe free places rather than imply that existing helpers have withdrawn. Invitations SHALL contain no other person's contact data, full birth dates or exact ages. They SHALL direct the recipient to a suitable open duty without promising exemption from assignment eligibility or confirmation rules.

#### Scenario: An ordinary-game e-mail follows the established tone

- **WHEN** a weekly vacancy e-mail is rendered for an ordinary game
- **THEN** it opens with Hallo and the recipient's name, names the team's responsibility next week and gives the date, time, age class and opponents
- **AND** it lists the missing tasks and counts, asks politely for support through the Heimspielplan, thanks the recipient and ends with the established e-mail sign-off

#### Scenario: SMS preserves the invitation and next action

- **WHEN** a recipient receives the SMS variant
- **THEN** it includes a personal greeting, responsible team, event date, start time, age class and missing tasks with counts
- **AND** it politely invites a suitable claim through the Heimspielplan, thanks the recipient and ends with VG TuS Handball

#### Scenario: Repeated tasks report vacancies rather than total capacity

- **WHEN** one of the two Verkauf positions is occupied and the other is empty
- **THEN** the invitation reports one free Verkauf position
- **AND** when both are empty it reports two free Verkauf positions under one shared Verkauf label

#### Scenario: A Spielfest invitation preserves event identity

- **WHEN** an invitation is rendered for a Spielfest
- **THEN** both channel variants identify the Spielfest by its full age-class label, date and start time
- **AND** neither variant introduces an ordinary home-versus-away matchup

#### Scenario: Invitation rendering obeys the catalog's named-value contract

- **WHEN** an invitation is requested with incomplete or unknown named context
- **THEN** rendering refuses the message before dispatch under the existing catalog validation rules

#### Scenario: Supplied names are plain text

- **WHEN** an otherwise valid invitation context contains literal braces in a person or team name
- **THEN** the rendered invitation preserves those braces as ordinary text without additional substitutions
