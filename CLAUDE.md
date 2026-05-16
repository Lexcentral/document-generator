# document-generator — AI Handoff

Church administration tools built in Google Apps Script, bound to a Google Sheet, with a companion GitHub Pages static site.

## Hard Rules

- No build pipeline — plain HTML/CSS/JS at repo root, Alpine.js CDN acceptable
- Simplicity first — volunteer-operated
- `gas/` is source of truth — always `cd gas && clasp push --force` after changes
- Never duplicate worship data — always read live from the worship spreadsheet
- All generated docs are set to `ANYONE_WITH_LINK VIEW` on creation
- User refers to table columns with 1-based indexing (col 1 = first column)

## Repo Layout

```
gas/                          Google Apps Script source (clasp root)
  appsscript.json             Manifest — timezone: America/New_York, runtimeVersion: V8
  Config.js                   All IDs (SPREADSHEET_IDS, TEMPLATE_IDS, FOLDER_IDS)
  Menu.js                     onOpen() — builds the "Worship" menu
  Utils.js                    parseDateStr, formatMeetingDate, formatSundayDate, getWeekStart
  Worship.js                  getWorshipSundays() — reads worship sheet live
  StaffAgenda.js              Feature 1 — staff agenda logic + logMeeting()
  DialogStaffAgenda.html      Modal: date picker → duplicate check → generate → link
  MeetingPacket.js            Feature 2 — council & business packet logic
  DialogMeetingPacket.html    Modal: meeting dropdown → generate → link
  DeaconAgenda.js             Feature 2.5 — deacon agenda logic
  DialogDeaconAgenda.html     Modal: meeting dropdown → generate → link
  LeadershipMeetings.js       Leadership schedule, reminders, attendee email dialog
  DialogLeadershipSchedule.html  Year picker for schedule generation
  DialogAttendeeReminder.html Editable email fields + send button for attendee reminders
  DialogTestReminder.html     Meeting type picker → sends test reminder to self
  Benevolence.js              Live benevolence calculation, verification, document lines
  CommitteeReports.js         Feature 3 — doGet/doPost web app + Quill HTML parser
index.html                    GitHub Pages static site — committee report submission form
docs/
  dev.md                      Developer guide
  operator.md                 Operator/volunteer guide
```

## Spreadsheets

### Bound sheet (document-generator)
ID: `1Af78xyGiWgEoGA7MI5F5H9Se3AHKlP7GBO3iZEZlLBY` (from `.clasp.json` parentId)
Access via `SpreadsheetApp.getActiveSpreadsheet()` — no ID needed in bound scripts.

| Tab | Columns |
|-----|---------|
| meetings | A: date, B: type (Staff only now), C: docURL, D: notes |
| leadership_meetings | A: date, B: time, C: type (Deacon/Council/Business), D: agenda (doc URL written on generation) |
| _committees | A: committeeName, B: chair, C: coChair, D: coChair2 — rows prefixed with `~` (e.g. `~Moderator`, `~Treasurer`, `~Clerk`) are officer contacts only; excluded from committee report lists but used in reminders/attendee emails |
| _people | A: name, B: email — single source of truth for all emails |
| _staff | A: name, B: meetings (comma-separated: "Deacon, Council, Business" — determines which reminder email lists they appear on) |
| _deacons | A: name — all deacons; emails looked up in `_people` |
| reminders | A: meeting (Deacon/Council/Business), B: days_before, C: zoom_link — canonical source of Zoom URLs |
| reports | A: submittedDate, B: committee, C: authorName, D: authorEmail, E: councilTarget (e.g. "June 2026"), F: bodyHTML |

### Benevolence sheet (read-only)
ID: `1cSijYiHP27xVqsXVKHcjARnnjqZhj5QURYmU-Qge_Tc`
Access via `SpreadsheetApp.openById(SPREADSHEET_IDS.benevolence)`.

**Checkbook tab** columns (0-indexed): 0=Check#, 1=Date, 2=Description, 3=Category, 4=Debit, 5=Credit, 6=Balance, 7=Transacted (TRUE/FALSE — FALSE rows are excluded from all calculations).
- Category `"Deposit"` = income; all other non-empty categories = expenditure
- Benevolence figures are always calculated live from Checkbook, not from Stats tab formulas

**Stats tab** — A3:B20; col A = label, col B = value. Cross-checked (not used as source) by `runBenevolenceCheck()`.

### Worship sheet (read-only)
ID: `19CUEXrS4gP7O7dSqInmbkUzWpJV0W_s9HLlQtDyqC_c`
Access via `SpreadsheetApp.openById(SPREADSHEET_IDS.worship)`.

**planning tab** (0-indexed columns):
- 0=date, 1=time, 2=dateTime, 3=lectionary, 4=theme, 5=sermonTitle, 6=notes, 7=texts, 8–11=readings set1, 12–15=readings set2, 16=scripture1Text, 17=scripture2Text

**_document_links tab** — aligned by row with planning:
- 0=date, 1=bulletinURL, 2=readerSheetURL, 3=leaderSlidesURL

## External API

```
GET https://lexcentral.com/event_calendar/bulletin?date=YYYY-MM-DD
```
Returns JSON. Use `data.weekly_events` — array of `{ day: "Monday", events: [{ title, time }] }`.
Day names are full English ("Monday" etc). `getWeekStart(agendaDate)` computes the Monday of that week; add 0–6 days to get Mon–Sun dates.

## Template & Folder IDs

| Key | ID |
|-----|----|
| TEMPLATE_IDS.staffAgenda | `1F78aHRRVv8OoIJIOM33rJaJYS2klqC5Yk45wCK7Q8Xk` |
| TEMPLATE_IDS.councilPacket | `1x575pj1kHduvSvuhojTfmYz91hhYPGS1e6fsAf8Tjgc` |
| TEMPLATE_IDS.businessPacket | `1x575pj1kHduvSvuhojTfmYz91hhYPGS1e6fsAf8Tjgc` (same template) |
| TEMPLATE_IDS.deaconAgenda | `14aUIHY8BLqW45qYiZBfNr-Hgn30Ab5Pyvu3GYEu2xOk` |
| FOLDER_IDS.staffAgenda | `1pQ74NSmKgbwhSqTWMTak_Gvwpp7PMvfg` |
| FOLDER_IDS.councilPacket | `1DgXdZzldQ-Gd_jMM88nH9xar-_YxWdvL` |
| FOLDER_IDS.businessPacket | `16Fib6pYRvFCcV0KIOYY0euJhrzsLOz-H` |
| FOLDER_IDS.deaconAgenda | `18nMv23KgTHboihXvSYBxclOwUCVCKvLc` |
| FOLDER_IDS.committeeReports | `1ZS78XK-VdGgrS-kD9GO81pZE0pQwcEf1` |
| FOLDER_IDS.editedMinutes | `1XRsETYMOPCE-rQahrNW5r8EQfnqIJrJu` |

## Leadership Meeting Schedule

**Council months**: March, June, September, December
- Council date: 1st Monday of month (shifts to 2nd Monday if US federal holiday — Labor Day always shifts September)
- Business date: Council date + 9 days = Wednesday of the following week
- Non-council months: 2nd Monday = Deacon meeting

**Generate schedule**: `Worship → Generate Leadership Schedule…` — prompts for a year, fills `leadership_meetings` tab, confirms before overwriting existing rows.

**Zoom links**: stored in `reminders` tab col C (`zoom_link`), read via `readMeetingConfig()`. NOT hardcoded anywhere.

## Duplicate Check & Logging

- **Staff agenda**: checks `meetings` tab; logs to `meetings` tab on generation
- **Council / Business / Deacon**: checks `leadership_meetings.agenda` (col D); writes URL to that column on generation via `updateLeadershipMeetingAgenda(dateStr, type, url)`

## Reminder System

`Worship → Install Reminder Trigger` — run once to set up a daily 8 AM time-based trigger calling `sendLeadershipReminders()`.

**`reminders` tab** (`meeting | days_before | zoom_link`): defines timing and Zoom URL per meeting type.

**Recipients** (from `getReminderRecipients()`):
- Deacon: chairs/co-chairs of "Deacons" row in `_committees` + staff with "Deacon" in meetings column
- Council/Business: chairs/co-chairs of "~Moderator" row in `_committees` + relevant staff

**Attendee list in email body** (from `getAttendeeList()`):
- Deacon: all `_deacons` names → `_people` emails
- Council/Business: all `_deacons` + all `_committees` chairs/co-chairs (including `~` rows) → `_people` emails

**Email body includes**: date, time, Zoom link, agenda URL (if generated), next 4 upcoming meetings across all types.

`Worship → Send Test Reminder Email…` — sends a real-format reminder to `Session.getActiveUser().getEmail()` only.

`Worship → Send Attendee Reminder…` — dialog shows all meetings; generates editable To/Subject/Body fields and a Send button that emails all attendees directly.

## Key Patterns

**Date parsing** — always use `parseDateStr(dateStr)` (local noon) to avoid UTC/Eastern rollover:
```js
function parseDateStr(dateStr) {
  var p = dateStr.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2], 12, 0, 0);
}
```

**Time from Sheets** — time cells come back as Date objects (epoch Dec 30, 1899). Use `formatTimeDisplay(val)` for email strings and `to24h(val)` for `<input type="time">` values. Never use `String(timeCell)` directly.

**Placeholder insertion** — `replaceText` only for simple scalars. For structured content, find the placeholder paragraph, record its index, remove it, then insert elements incrementing the index:
```js
var placeholder = found.getElement().getParent();
var insertIdx   = body.getChildIndex(placeholder);
body.removeChild(placeholder);
// insertParagraph / insertListItem / insertTable at insertIdx++
```

**List glyph type** — always set via `setAttributes` (not separate `setNestingLevel` + `setGlyphType` calls) to prevent inheriting ordered list styles from adjacent template content:
```js
var attrs = {};
attrs[DocumentApp.Attribute.NESTING_LEVEL] = level;
attrs[DocumentApp.Attribute.GLYPH_TYPE]    = glyphType;
item.setAttributes(attrs);
```
`GLYPH_BY_LEVEL` = [BULLET, HOLLOW_BULLET, SQUARE_BULLET].

**Hyperlink replacement** — use `replaceWithHyperlink(element, tokenRegex, linkText, url)` for `{zoomLink}` and similar; `replaceText` cannot create hyperlinks.

**Footer replacement** — `doc.getFooter()` must be called BEFORE structural insertions (`insertBulletSection`, `insertWorshipSection`, etc.) or it may not find tokens.

**Cross-document element copy** — `elem.copy()` then `body.insertParagraph(idx, copy)` / `insertListItem` / `insertTable`. Silently skips unsupported element types in a try/catch.

## Token Names in Templates

**Staff agenda**: `{meetingDate}`, `{worshipTable}`, `{weeklyEvents}`

**Council & Business packets**:
- `{meetingType}` → "Church Council" or "Business Meeting"
- `{dateTime}` → "March 2, 2026, 7 p.m." (minutes omitted when :00)
- `{zoomLink}` → hyperlink "Zoom Link" from `reminders` tab
- `{committeesAndChairs}` → bullet list from `_committees` (skips `~` rows)
- `{staff}` → bullet list from `_staff` col A
- `{benevolenceData}` → live from Checkbook: top-level items at placeholder indent, category items one level deeper (always BULLET glyph at all levels)
- `{previousBusinessMeetingMinutes}` → copied from clerk-edited file (see below)
- `{previousChurchCouncilMinutes}` → copied from clerk-edited file (see below)
- `{committeeReports}` → business packets: H1 sections from council clerk file; council packets: replaced with placeholder text (Feature 3)

**Deacon agenda**: `{meetingDate}`, `{dateTime}`, `{zoomLink}` (hyperlink), `{staff}` (Deacon-tagged staff only), `{worshipTable}`

## Clerk-Edited Minutes Files

After each meeting, the clerk submits an edited Google Doc to the `editedMinutes` folder.
File naming: `YYYY-MM-DD_council_EDITED` or `YYYY-MM-DD_business_EDITED` (any suffix after the type word is ignored; the date prefix is what matters).

**Structure within each file**:
- Meeting sections: plain paragraphs; each starts with a paragraph beginning `"Central Baptist Church"` (not H1)
- Committee reports: H1 headings (`# Report Name`)

**Council file** sections in order: (1) current council minutes, (2) previous business meeting minutes, (3) previous council meeting minutes → then H1 committee reports

**Business file** sections in order: (1) current business minutes, (2) previous business meeting minutes, (3) previous council meeting minutes → then H1 committee reports

**Sources when generating a new Council packet**:
- `{previousBusinessMeetingMinutes}` → section 1 of most recent `_business` file before the council date
- `{previousChurchCouncilMinutes}` → section 1 of most recent `_council` file before the council date

**Sources when generating a new Business packet**:
- `{previousBusinessMeetingMinutes}` → section 2 of most recent `_council` file (clerk embeds prev business there)
- `{previousChurchCouncilMinutes}` → section 1 of most recent `_council` file
- `{committeeReports}` → all H1 sections from the same council file

Sections are identified by keyword: "Business" → business section; "Council" or "Deacons" → council section.
Date validation: section 1 heading must contain the date from the filename (e.g. "March 11") or an error is thrown before the document is created.

## Benevolence Fund

`Worship → Check Benevolence Figures` runs `runBenevolenceCheck()`:
1. Gets last Business meeting date from `leadership_meetings` (not from benevolence sheet)
2. Reads Checkbook, calculates beginning balance (last transacted balance ≤ meeting date), deposited (Credits after), expended (Debits after, non-Deposit categories)
3. Internal check: beginning + deposited − expended = last checkbook balance
4. Cross-check: compares against Stats tab A3:B20 formula values
5. Alerts with results or mismatch details

`getBenevolenceLines(lastMeetingDate)` in `Benevolence.js` returns structured `{ text, level }` entries used by `insertBenevolenceSection()` in `MeetingPacket.js`.

## Feature Status & Roadmap

### ✅ Feature 1 — Staff Meeting Agenda
Menu: `Worship → Generate Staff Agenda`
- Date picker modal; fetches weekly events from lexcentral.com
- Pulls 6 worship Sundays; fills `{meetingDate}`, `{worshipTable}`, `{weeklyEvents}`
- Writes to `meetings` tab; returns doc URL

### ✅ Feature 2 — Council & Business Meeting Packets
Menu: `Worship → Generate Council Packet`, `Worship → Generate Business Meeting Packet`
- Dropdown of upcoming meetings from `leadership_meetings` (pre-fills time)
- Fills all tokens including live benevolence data and previous meeting minutes from clerk files
- Sets doc to ANYONE_WITH_LINK VIEW; writes URL to `leadership_meetings.agenda`
- Duplicate check on `leadership_meetings.agenda` column

### ✅ Feature 2.5 — Deacon Agenda
Menu: `Worship → Generate Deacon Agenda`
- Dropdown of upcoming Deacon meetings from `leadership_meetings`
- Fills `{meetingDate}`, `{dateTime}`, `{zoomLink}`, `{staff}` (Deacon-tagged only), `{worshipTable}`
- Sets doc to ANYONE_WITH_LINK VIEW; writes URL to `leadership_meetings.agenda`

### ✅ Leadership Schedule & Reminders
- Schedule generator fills a full year into `leadership_meetings` (handles federal holidays)
- Daily trigger sends reminder emails N days before each meeting (per `reminders` tab)
- Test reminder, attendee reminder dialog with editable fields and send button

### ✅ Benevolence Reporting
- Live calculation from Checkbook; cross-checks Stats tab formulas
- Inserted live into council/business packets with proper indentation

### ✅ Feature 3 — Committee Report Submission Form
- `index.html` at repo root — plain HTML, Alpine.js + Quill (CDN), no build step
- Deployed via GitHub Pages (lexcentral/document-generator), source: root of main
- On load: fetches committees and upcoming Council targets from GAS `doGet`
- Submit: POSTs JSON as `text/plain` to GAS `doPost`; stores row in `reports` tab
- `appsscript.json` webapp stanza: `executeAs: USER_DEPLOYING, access: ANYONE_ANONYMOUS`
- `{committeeReports}` in council packets reads from `reports` tab filtered by `councilTarget`
- Business packets pull committee reports from the clerk-edited council file (unchanged)
- **GAS_URL** in `index.html` must be updated after each new GAS deployment

## Push Command

```bash
cd gas && clasp push --force
```
