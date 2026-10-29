# Mad Apple Winter Workouts

Friday winter practice plan for the Mad Apple baseball team (6–8 pm, Mad Apple North, 4141 N. Clinton St.).

- **Plan:** a tab for each Friday with the minute-by-minute schedule, station drills, conditioning and the closing competition. Open to anyone with the link.
- **Coach tools (team PIN):** attendance, auto-built groups, a pitch-by-pitch bullpen charter, test results, coach notes, and a Season tab with progress.

## Where the data lives

Nothing about players is stored in this repo. Roster, attendance, bullpens, tests and notes live in the private Google Sheet **Mad Apple Winter Workouts — Data**. The page talks to it through the Apps Script in `apps-script/Code.gs`.

## Setup

1. Open the Google Sheet → **Extensions → Apps Script**. Paste `apps-script/Code.gs`, save.
2. **Deploy → New deployment → Web app.** Execute as *Me*, access *Anyone*. Copy the URL ending in `/exec`.
3. Put that URL in `config.js`.
4. Codes live in the sheet's **Config** tab: **Coach PIN** (full access) and **Player code** (leaderboards + logging tests and bullpens; players can only fill empty boxes).

The backend is generic: it reads every tab and saves rows by key, adding new columns as needed. New features ship by pushing to this repo; the Apps Script shouldn't need to change again.
