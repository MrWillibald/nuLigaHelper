## Purpose

Defines the shared visual treatment of public page footer links so legal information remains easy to find and read across the site.

## ADDED Requirements

### Requirement: Legal footer links fit the shared design

The shared footer SHALL present the Impressum and Datenschutzerklärung links in a color and style that is readable against the footer background and consistent with the site's visual language. Both links SHALL remain visually identifiable as links and SHALL have visible hover and keyboard focus states on desktop and narrow screens.

#### Scenario: Navigate legal links from the footer

- **WHEN** a visitor views the footer on the overview or another public page
- **THEN** both legal links are readable and identifiable as interactive links
- **AND** activating each link opens its corresponding legal page

#### Scenario: Keyboard focus

- **WHEN** a visitor tabs to either legal footer link
- **THEN** its focus state is clearly visible against the footer background
