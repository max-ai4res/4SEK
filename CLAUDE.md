# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

4SEK is an interactive mockup of a Security Operations Center (SOC) web dashboard for monitoring a shopping mall via AI-equipped cameras. It is a **static front-end only** project — no build system, no bundler, no server-side code.

The UI language is Italian.

## How to Run

Open `index.html` directly in a browser (desktop SOC dashboard) or `mobile.html` (mobile guard app showcase). No build step required.

```bash
open index.html      # macOS
```

## Architecture

### Single-file HTML apps with inline CSS and JS

- **`index.html`** (~1260 lines) — Desktop SOC dashboard. Contains all CSS, HTML, and JS in a single file. The JS section starts around line 993.
- **`mobile.html`** (~770 lines) — Mobile guard app rendered inside phone-frame mockups for showcase purposes. Also fully self-contained.

Both files use the same design system (CSS custom properties in `:root`, Inter + JetBrains Mono fonts from Google Fonts).

### Data model (index.html JS)

All data is hardcoded in JS arrays at the top of the `<script>` block:

- `CAMERAS[]` — camera definitions with id, zone, area, image, status (`ok`/`warning`/`alert`), AI detection state
- `ALERTS[]` — security events with severity (`critical`/`high`/`medium`/`low`), associated camera, description
- `GUARDS[]` — guard personnel with status (`patrol`/`standby`/`responding`/`break`), zone assignment

### View system (index.html)

Navigation uses a sidebar with `data-view` attributes. `switchView(viewId)` toggles `.view` panels by id:
- `dashboard` — metrics, 3x3 camera grid, alert list
- `cameras` — full camera grid with area/status filters
- `alerts` — full alert list with severity filters
- `map` — SVG floor plan with camera positions
- `guards` — guard roster and assignment panel
- `analytics` — chart cards (hourly bar chart)
- `comms` — speaker controls, emergency contacts, comms log

### Key UI patterns

- Camera cells open a detail modal (`openCamModal`)
- Alert items open an incident side panel (`openIncident`)
- `simulateAlert()` fires toast notifications every 12s to simulate live events
- Filter chips use event delegation on click with `data-filter` attributes

### Assets

- `assets/img/` — stock photos used as camera feed placeholders
- `img_centro/` — real photos of a shopping mall (reference material)
- `img_ai/` — AI detection screenshot references

## Feature Spec

`funzionalita.md` describes the intended capabilities: AI-powered cameras detecting threats (weapons, fights, theft, loitering, etc.), audio detection via onboard microphones, speaker announcements, guard coordination, and emergency services integration.
