# Developer Guide — document-generator

## Prerequisites

- Node.js (for clasp)
- `npm install -g @google/clasp`
- `clasp login` (authenticate with the Google account that owns the spreadsheet)
- Access to the document-generator Google Sheet and Drive folder

## Repository Layout

```
gas/                  Apps Script source — this is what gets pushed to Google
  appsscript.json     Script manifest
  Config.js           Centralized IDs (spreadsheets, templates, folders)
  Menu.js             onOpen() hook — registers the Worship menu
  Utils.js            Date helpers shared across features
  Worship.js          Live reads from the worship spreadsheet
  StaffAgenda.js      Feature 1 logic + dialog handler
  DialogStaffAgenda.html  HTML modal for Feature 1
web/                  GitHub Pages static site (Feature 3 — not yet built)
docs/
  dev.md              This file
  operator.md         Volunteer/operator guide
CLAUDE.md             AI assistant handoff — full project context
```

## Pushing Changes

Always push from the `gas/` directory:

```bash
cd gas && clasp push --force
```

After pushing, reload the bound Google Sheet to pick up `onOpen` changes. For dialog/HTML changes the reload is sufficient. For logic changes you can test immediately via the Worship menu.

## Local Development Workflow

1. Edit files in `gas/` using any editor
2. `cd gas && clasp push --force`
3. Open the Google Sheet → Worship menu → test the feature
4. Check Apps Script execution logs: [script.google.com](https://script.google.com) → your project → Executions

## Adding a New Feature

1. Add any new template/folder IDs to `Config.js`
2. Create a new `FeatureName.js` file for the logic
3. Create a `DialogFeatureName.html` if a modal is needed
4. Wire the menu item in `Menu.js` under `onOpen()`
5. Push

Keep each feature in its own file. Shared utilities go in `Utils.js`.

## Key IDs

| Resource | ID |
|----------|----|
| Bound spreadsheet | `1Af78xyGiWgEoGA7MI5F5H9Se3AHKlP7GBO3iZEZlLBY` |
| Worship spreadsheet (read-only) | `19CUEXrS4gP7O7dSqInmbkUzWpJV0W_s9HLlQtDyqC_c` |
| Staff agenda template | `1F78aHRRVv8OoIJIOM33rJaJYS2klqC5Yk45wCK7Q8Xk` |
| Staff agenda output folder | `1pQ74NSmKgbwhSqTWMTak_Gvwpp7PMvfg` |
| Council packet template | **TBD** |
| Business meeting packet template | **TBD** |

## External API

Weekly events are fetched from:
```
https://lexcentral.com/event_calendar/bulletin?date=YYYY-MM-DD
```
Parse the response JSON and use `data.weekly_events`. Each element has `day` (full day name) and `events` (array of `{ title, time }`).

## Spreadsheet Tabs

The bound sheet must have these tabs with these exact names:

| Tab | Purpose |
|-----|---------|
| `meetings` | One row per generated document — date, type, URL, notes |
| `_committees` | Committee roster — name, chair, co-chair(s) |
| `_people` | Staff/member directory — name, email |
| `reports` | Submitted committee reports |

Header rows (row 1) for each tab:

**meetings:** `date | type | docURL | notes`
**_committees:** `committeeName | chair | coChair | coChair2`
**_people:** `name | email`
**reports:** `submittedDate | committee | authorName | authorEmail | councilTarget | bodyHTML`

## Feature 3 — Web App Setup (when building)

The GAS doPost endpoint must be deployed as a Web App:
- Execute as: **Me**
- Who has access: **Anyone**

Add to `appsscript.json`:
```json
"webapp": {
  "executeAs": "USER_DEPLOYING",
  "access": "ANYONE_ANONYMOUS"
}
```

Deploy from the Apps Script editor (Deploy → New deployment → Web app). The deployment URL goes in `web/index.html` as the POST target.

## Roadmap

| Feature | Status |
|---------|--------|
| Staff Meeting Agenda | ✅ Complete |
| Council & Business Packet | 🔲 Not started — needs template IDs from Aaron |
| Committee Report Form (web + doPost) | 🔲 Not started |
