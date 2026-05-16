# Operator Guide — document-generator

This guide is for church staff and volunteers who use the document generator tools. No technical background is needed.

---

## How to Access

Open the **document-generator Google Sheet**. You'll see a **Worship** menu in the top menu bar (it may take a moment to appear when you first open the sheet).

If the Worship menu doesn't appear, try refreshing the page.

---

## Feature 1 — Generate Staff Meeting Agenda

**Menu:** Worship → Generate Staff Agenda

1. Click **Worship → Generate Staff Agenda**
2. A dialog box will appear. Select the date of the staff meeting.
3. Click **Generate Agenda**.
4. Wait a few seconds — the tool is fetching this week's events and worship data.
5. When finished, a link to the new document will appear. Click **Open Agenda ↗** to view it.

**What gets created:**
- A Google Doc in the Staff Agendas folder with the meeting date in its name
- A row added to the `meetings` tab of the spreadsheet with the date and doc link

**If you generate an agenda for a date that already exists:**
A confirmation prompt will appear asking if you want to overwrite the existing sheet entry. Clicking OK overwrites the row; clicking Cancel adds a new row.

**What's in the agenda:**
- The meeting date
- This week's events from the church calendar (organized by day)
- A list of the 6 most relevant worship Sundays (the most recent past Sunday + the next 5), showing the lectionary, theme, notes, and a link to the bulletin if one has been generated

---

## Feature 2 — Generate Council or Business Meeting Packet *(coming soon)*

**Menu:** Worship → Generate Council Packet / Worship → Generate Business Meeting Packet

Similar flow to the Staff Agenda:
1. Select the meeting date
2. The tool assembles a packet including:
   - Previous meeting minutes (linked from the sheet)
   - All committee reports submitted for that council month
3. A link to the new packet document will appear

---

## Feature 3 — Submit a Committee Report *(coming soon)*

A web form (accessible from a link, no login required) where committee chairs submit their reports before each council meeting.

**Form fields:**
- Committee name
- Council meeting target (which month's packet to include it in)
- Your name and email
- Report body (you can paste from Word)

Submitted reports are saved automatically and pulled into the next council packet.

---

## The Spreadsheet Tabs

You generally don't need to edit these directly, but here's what each tab does:

| Tab | What it tracks |
|-----|----------------|
| `meetings` | Every generated agenda and packet — date, type, and a link to the doc |
| `_committees` | Committee names and chair/co-chair names |
| `_people` | Staff and member names and emails |
| `reports` | Submitted committee reports (populated by the web form) |

---

## Troubleshooting

**The Worship menu doesn't appear**
Refresh the page and wait a few seconds. If it still doesn't appear, contact your developer.

**"Events unavailable" appears in the agenda**
The church calendar website couldn't be reached at generation time. Try regenerating, or manually paste the events into the document afterward.

**The agenda was generated but the wrong date is shown**
Date formatting uses Eastern Time. If you're generating from outside the US, contact your developer.

**A bulletin shows "(not yet generated)"**
The bulletin for that Sunday hasn't been created yet in the worship system. Once it's generated there, regenerate the agenda and it will appear as a link.
