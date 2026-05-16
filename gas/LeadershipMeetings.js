var COUNCIL_MONTHS = [3, 6, 9, 12];

var MEETING_TIMES = {
  Deacon:   '7:00 PM',
  Council:  '7:00 PM',
  Business: '6:30 PM'
};

// Reads the reminders tab and returns a config map keyed by meeting type.
// reminders tab columns: meeting | days_before | zoom_link
function readMeetingConfig() {
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('reminders').getDataRange().getValues();
  var config = {};
  data.slice(1).forEach(function(r) {
    if (r[0]) config[r[0]] = { daysBefore: parseInt(r[1], 10), zoomLink: r[2] || '' };
  });
  return config;
}

// ── Schedule generation ───────────────────────────────────────────────────────

function showLeadershipScheduleDialog() {
  var tmpl = HtmlService.createTemplateFromFile('DialogLeadershipSchedule');
  tmpl.nextYear = new Date().getFullYear() + 1;
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(360).setHeight(210),
    'Generate Leadership Schedule'
  );
}

// Returns the count of existing rows in leadership_meetings for the given year.
function checkLeadershipScheduleYear(year) {
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('leadership_meetings').getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (yearOf(data[i][0]) === year) count++;
  }
  return count;
}

// Deletes any existing rows for the year, then appends the generated schedule.
function generateLeadershipSchedule(year) {
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('leadership_meetings');
  var data   = sheet.getDataRange().getValues();
  var config = readMeetingConfig();

  // Delete existing rows for this year in reverse order so row numbers stay valid.
  for (var i = data.length - 1; i >= 1; i--) {
    if (yearOf(data[i][0]) === year) sheet.deleteRow(i + 1);
  }

  buildScheduleRows(year, config).forEach(function(row) { sheet.appendRow(row); });
}

function buildScheduleRows(year, config) {
  var rows = [];
  for (var month = 1; month <= 12; month++) {
    if (COUNCIL_MONTHS.indexOf(month) !== -1) {
      var councilDate = nthMondayOfMonth(year, month, 1);
      if (isUSFederalHoliday(councilDate)) councilDate = nthMondayOfMonth(year, month, 2);
      var businessDate = new Date(councilDate);
      businessDate.setDate(councilDate.getDate() + 9);
      rows.push([councilDate,  MEETING_TIMES.Council,  'Council',  '']);
      rows.push([businessDate, MEETING_TIMES.Business, 'Business', '']);
    } else {
      rows.push([nthMondayOfMonth(year, month, 2), MEETING_TIMES.Deacon, 'Deacon', '']);
    }
  }
  return rows;
}

// Returns the nth Monday (n=1 is first) of the given month (1-based).
function nthMondayOfMonth(year, month, n) {
  var first = new Date(year, month - 1, 1);
  var daysToFirstMonday = (1 - first.getDay() + 7) % 7;
  return new Date(year, month - 1, 1 + daysToFirstMonday + (n - 1) * 7, 12, 0, 0);
}

// Returns true if date is a US federal holiday (or its Monday observation).
function isUSFederalHoliday(date) {
  var y = date.getFullYear(), m = date.getMonth() + 1, d = date.getDate(), dow = date.getDay();

  // Fixed-date holidays — check actual date and Monday observation (when holiday falls on Sunday).
  var fixed = [[1,1],[6,19],[7,4],[11,11],[12,25]];
  for (var i = 0; i < fixed.length; i++) {
    var hm = fixed[i][0], hd = fixed[i][1];
    var hDay = new Date(y, hm - 1, hd);
    var obs  = new Date(hDay);
    if (hDay.getDay() === 0) obs.setDate(hd + 1);       // Sunday → observed Monday
    else if (hDay.getDay() === 6) obs.setDate(hd - 1);  // Saturday → observed Friday
    if (date.toDateString() === obs.toDateString()) return true;
  }

  // Floating Monday holidays.
  if (dow === 1) {
    if (m === 1  && d >= 15 && d <= 21) return true; // MLK Day
    if (m === 2  && d >= 15 && d <= 21) return true; // Presidents Day
    if (m === 5  && d >= 25)            return true; // Memorial Day (last Monday of May)
    if (m === 9  && d <= 7)             return true; // Labor Day
    if (m === 10 && d >= 8  && d <= 14) return true; // Columbus Day
  }
  return false;
}

// Returns the next N upcoming meetings across all types after (not including) afterDate.
// Each entry is formatted as "MMMM d, yyyy — Type".
function getNextMeetings(afterDate, count) {
  var tz   = Session.getScriptTimeZone();
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('leadership_meetings').getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length && results.length < count; i++) {
    var rowDate = data[i][0], rowType = data[i][2];
    if (!rowType || !(rowDate instanceof Date)) continue;
    var d = new Date(rowDate);
    d.setHours(0, 0, 0, 0);
    if (d > afterDate) {
      results.push(Utilities.formatDate(rowDate, tz, 'MMMM d, yyyy') + ' — ' + rowType);
    }
  }
  return results;
}

// Formats a time value from a Sheets cell (Date object or string) for display in emails.
function formatTimeDisplay(timeVal) {
  if (timeVal instanceof Date) {
    return Utilities.formatDate(timeVal, Session.getScriptTimeZone(), 'h:mm a');
  }
  return String(timeVal);
}

function yearOf(val) {
  if (val instanceof Date) return parseInt(Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy'), 10);
  if (typeof val === 'string' && val.length >= 4) return parseInt(val.substring(0, 4), 10);
  return null;
}

// ── Reminders ─────────────────────────────────────────────────────────────────

// Called daily by a time-based trigger. Sends reminder emails for meetings
// whose date is exactly N days away, per the reminders sheet config.
function sendLeadershipReminders() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var meetingConfig = readMeetingConfig(); // { Deacon: { daysBefore, zoomLink }, … }

  var peopleMap      = buildPeopleMap(ss.getSheetByName('_people').getDataRange().getValues());
  var committeesData = ss.getSheetByName('_committees').getDataRange().getValues();
  var staffData      = ss.getSheetByName('_staff').getDataRange().getValues();
  var deaconsData    = ss.getSheetByName('_deacons').getDataRange().getValues();

  ss.getSheetByName('leadership_meetings').getDataRange().getValues().slice(1).forEach(function(row) {
    var rawDate = row[0], time = row[1], type = row[2], agendaUrl = row[3] || '';
    if (!rawDate || !type || !meetingConfig[type]) return;

    var meetingDate = new Date(rawDate);
    meetingDate.setHours(0, 0, 0, 0);
    var daysOut = Math.round((meetingDate - today) / 86400000);
    if (daysOut !== meetingConfig[type].daysBefore) return;

    var zoomLink = meetingConfig[type].zoomLink || '';
    sendReminderForMeeting(type, meetingDate, formatTimeDisplay(time), zoomLink, agendaUrl, peopleMap, committeesData, staffData, deaconsData);
  });
}

function sendReminderForMeeting(type, meetingDate, time, zoomLink, agendaUrl, peopleMap, committeesData, staffData, deaconsData) {
  var recipients   = getReminderRecipients(type, committeesData, staffData, peopleMap);
  if (!recipients.length) return;

  var attendeeList = getAttendeeList(type, committeesData, deaconsData, peopleMap);
  var dateLabel    = Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  var typeLabel    = type === 'Council' ? 'Church Council'
                   : type === 'Business' ? 'Business Meeting'
                   : 'Deacons';

  var upcomingDates = getNextMeetings(meetingDate, 4);

  var lines = [
    'This is a reminder for the upcoming ' + typeLabel + ' meeting.',
    '',
    'Date: ' + dateLabel,
    'Time: ' + time,
    'Zoom: ' + zoomLink
  ];
  if (agendaUrl) lines.push('Agenda: ' + agendaUrl);
  if (upcomingDates.length) {
    lines.push('', 'Upcoming meetings:');
    upcomingDates.forEach(function(d) { lines.push('  ' + d); });
  }
  lines.push('', 'Attendee email list (copy to send group reminder):', attendeeList.join('\n'));

  MailApp.sendEmail({
    to:      recipients.join(','),
    subject: typeLabel + ' Meeting Reminder — ' + dateLabel,
    body:    lines.join('\n')
  });
}

// Reminder recipients: chair/co-chairs of Deacons (Deacon meetings) or ~Moderator
// (Council/Business), plus any staff whose meetings list includes the meeting type.
function getReminderRecipients(type, committeesData, staffData, peopleMap) {
  var emails     = [];
  var targetName = type === 'Deacon' ? 'Deacons' : '~Moderator';

  committeesData.slice(1).forEach(function(r) {
    if (r[0] !== targetName) return;
    [r[1], r[2], r[3]].forEach(function(name) {
      if (name && peopleMap[name]) emails.push(peopleMap[name]);
    });
  });

  staffData.slice(1).forEach(function(r) {
    var name     = r[0];
    var meetings = (r[1] || '').split(',').map(function(m) { return m.trim(); });
    if (name && meetings.indexOf(type) !== -1 && peopleMap[name]) emails.push(peopleMap[name]);
  });

  return emails;
}

// Full attendee list for the email body: all deacons (Deacon meetings) or all
// committee members including ~-prefixed officers plus deacons (Council/Business).
function getAttendeeList(type, committeesData, deaconsData, peopleMap) {
  var lines = [], seen = {};
  function add(name) {
    if (!name || seen[name] || !peopleMap[name]) return;
    seen[name] = true;
    lines.push(name + ' <' + peopleMap[name] + '>');
  }
  deaconsData.slice(1).forEach(function(r) { add(r[0]); });
  if (type !== 'Deacon') {
    committeesData.slice(1).forEach(function(r) { add(r[1]); add(r[2]); add(r[3]); });
  }
  return lines;
}

function buildPeopleMap(peopleData) {
  var map = {};
  peopleData.slice(1).forEach(function(r) { if (r[0] && r[1]) map[r[0]] = r[1]; });
  return map;
}

// ── Attendee reminder mailto ──────────────────────────────────────────────────

function showAttendeeReminderDialog() {
  var tmpl = HtmlService.createTemplateFromFile('DialogAttendeeReminder');
  tmpl.meetings = JSON.stringify(getAllLeadershipMeetings());
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(480).setHeight(520),
    'Send Attendee Reminder'
  );
}

// Returns all meetings from leadership_meetings sorted by date, for the dropdown.
function getAllLeadershipMeetings() {
  var tz   = Session.getScriptTimeZone();
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('leadership_meetings').getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0], rowType = data[i][2];
    if (!rowDate || !rowType || !(rowDate instanceof Date)) continue;
    items.push({
      dateStr: Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd'),
      label:   Utilities.formatDate(rowDate, tz, 'MMMM d, yyyy') + ' — ' + rowType,
      type:    rowType
    });
  }
  items.sort(function(a, b) { return a.dateStr < b.dateStr ? -1 : 1; });
  return items;
}

// Returns { to, subject, body } for the client to assemble into a mailto: link.
function getAttendeeReminderData(dateStr, type) {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var tz          = Session.getScriptTimeZone();
  var meetingDate = parseDateStr(dateStr);

  // Find meeting row for time + agenda URL.
  var rows = ss.getSheetByName('leadership_meetings').getDataRange().getValues();
  var time = '', agendaUrl = '';
  for (var i = 1; i < rows.length; i++) {
    var rowDate = rows[i][0], rowType = rows[i][2];
    if (rowType === type && rowDate instanceof Date &&
        rowDate.toDateString() === meetingDate.toDateString()) {
      time      = formatTimeDisplay(rows[i][1]);
      agendaUrl = rows[i][3] || '';
      break;
    }
  }

  var zoomLink  = (readMeetingConfig()[type] || {}).zoomLink || '';
  var typeLabel = type === 'Council' ? 'Church Council'
                : type === 'Business' ? 'Business Meeting'
                : 'Deacons';
  var dateLabel = Utilities.formatDate(meetingDate, tz, 'MMMM d, yyyy');

  // All attendee emails for this meeting type (same logic as getAttendeeList).
  var peopleMap      = buildPeopleMap(ss.getSheetByName('_people').getDataRange().getValues());
  var committeesData = ss.getSheetByName('_committees').getDataRange().getValues();
  var deaconsData    = ss.getSheetByName('_deacons').getDataRange().getValues();
  var emails = [], seen = {};
  function addEmail(name) {
    if (!name || seen[name] || !peopleMap[name]) return;
    seen[name] = true;
    emails.push(peopleMap[name]);
  }
  deaconsData.slice(1).forEach(function(r) { addEmail(r[0]); });
  if (type !== 'Deacon') {
    committeesData.slice(1).forEach(function(r) { addEmail(r[1]); addEmail(r[2]); addEmail(r[3]); });
  }

  var upcomingDates = getNextMeetings(meetingDate, 4);
  var bodyLines = [
    'This is a reminder for our upcoming ' + typeLabel + ' meeting.',
    '',
    'Date: ' + dateLabel,
    'Time: ' + time,
    'Zoom: ' + zoomLink
  ];
  if (agendaUrl) bodyLines.push('Agenda: ' + agendaUrl);
  if (upcomingDates.length) {
    bodyLines.push('', 'Upcoming meetings:');
    upcomingDates.forEach(function(d) { bodyLines.push('  ' + d); });
  }

  return {
    to:      emails,
    subject: typeLabel + ' Meeting Reminder — ' + dateLabel,
    body:    bodyLines.join('\n')
  };
}

function sendAttendeeReminder(to, subject, body) {
  MailApp.sendEmail({ to: to.join(','), subject: subject, body: body });
}

// ── Test reminder ─────────────────────────────────────────────────────────────

function showTestReminderDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('DialogTestReminder').setWidth(360).setHeight(180),
    'Send Test Reminder Email'
  );
}

// Sends a real reminder email for the next upcoming meeting of the given type,
// but addressed only to the current user rather than the full recipient list.
function sendTestReminderEmail(type) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var peopleMap      = buildPeopleMap(ss.getSheetByName('_people').getDataRange().getValues());
  var committeesData = ss.getSheetByName('_committees').getDataRange().getValues();
  var staffData      = ss.getSheetByName('_staff').getDataRange().getValues();
  var deaconsData    = ss.getSheetByName('_deacons').getDataRange().getValues();

  // Find the next upcoming meeting of this type.
  var meetingsData = ss.getSheetByName('leadership_meetings').getDataRange().getValues();
  var target = null;
  for (var i = 1; i < meetingsData.length; i++) {
    var rowDate = meetingsData[i][0], rowType = meetingsData[i][2];
    if (rowType !== type || !(rowDate instanceof Date)) continue;
    var d = new Date(rowDate);
    d.setHours(0, 0, 0, 0);
    if (d >= today) { target = meetingsData[i]; break; }
  }
  if (!target) return 'No upcoming ' + type + ' meeting found.';

  var meetingDate = new Date(target[0]);
  var time        = formatTimeDisplay(target[1]);
  var zoomLink    = readMeetingConfig()[type].zoomLink || '';
  var agendaUrl   = target[3] || '';

  var attendeeList = getAttendeeList(type, committeesData, deaconsData, peopleMap);
  var dateLabel    = Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  var typeLabel    = type === 'Council' ? 'Church Council'
                   : type === 'Business' ? 'Business Meeting'
                   : 'Deacons';

  var upcomingDates = getNextMeetings(meetingDate, 4);

  var lines = [
    '[This is a test — in production this goes to the meeting chair(s) and staff]',
    '',
    'This is a reminder for the upcoming ' + typeLabel + ' meeting.',
    '',
    'Date: ' + dateLabel,
    'Time: ' + time,
    'Zoom: ' + zoomLink
  ];
  if (agendaUrl) lines.push('Agenda: ' + agendaUrl);
  if (upcomingDates.length) {
    lines.push('', 'Upcoming meetings:');
    upcomingDates.forEach(function(d) { lines.push('  ' + d); });
  }
  lines.push('', 'Attendee email list (copy to send group reminder):', attendeeList.join('\n'));

  var to = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to:      to,
    subject: '[TEST] ' + typeLabel + ' Meeting Reminder — ' + dateLabel,
    body:    lines.join('\n')
  });

  return 'Test email sent to ' + to + ' for the ' + type + ' meeting on ' + dateLabel + '.';
}

// ── Trigger setup ─────────────────────────────────────────────────────────────

// Run once from the menu to install (or reinstall) the daily reminder trigger.
function installReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendLeadershipReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendLeadershipReminders')
    .timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getUi().alert('Done — reminder trigger will run daily at 8 AM.');
}
