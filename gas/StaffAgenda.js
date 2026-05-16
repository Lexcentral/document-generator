function showStaffAgendaDialog() {
  var html = HtmlService.createHtmlOutputFromFile('DialogStaffAgenda')
    .setWidth(420)
    .setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Staff Agenda');
}

function generateStaffAgenda(dateStr, overwriteRow) {
  var meetingDate = parseDateStr(dateStr);
  var eventDays   = fetchWeeklyEventsData(dateStr);
  var sundays     = getWorshipSundays(meetingDate);

  var folder  = DriveApp.getFolderById(FOLDER_IDS.staffAgenda);
  var docFile = DriveApp.getFileById(TEMPLATE_IDS.staffAgenda).makeCopy(folder);
  docFile.setName(dateStr + '_staff-agenda');
  docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var body = DocumentApp.openById(docFile.getId()).getBody();
  body.replaceText('\\{meetingDate\\}', formatMeetingDate(meetingDate));
  insertWorshipSection(body, sundays);
  insertWeeklyEventsSection(body, eventDays, meetingDate);

  var docUrl = docFile.getUrl();
  logMeeting(meetingDate, 'Staff', docUrl, overwriteRow || 0);
  return docUrl;
}

function fetchWeeklyEventsData(dateStr) {
  var url = 'https://lexcentral.com/event_calendar/bulletin?date=' + dateStr;
  try {
    var data = JSON.parse(UrlFetchApp.fetch(url).getContentText());
    return data.weekly_events || [];
  } catch (e) {
    return [];
  }
}

// Replaces {worshipTable} with an H5 heading + bullet list for each Sunday.
// Heading links to the bulletin if one exists. Theme/notes only shown if non-empty.
function insertWorshipSection(body, sundays) {
  var found = body.findText('\\{worshipTable\\}');
  if (!found) return;

  var placeholder = found.getElement().getParent();
  var insertIdx   = body.getChildIndex(placeholder);
  body.removeChild(placeholder);

  sundays.forEach(function(s) {
    var headingText = s.dateFormatted + (s.lectionary ? ' | ' + s.lectionary : '');
    var h = body.insertParagraph(insertIdx, headingText);
    h.setHeading(DocumentApp.ParagraphHeading.HEADING5);
    if (s.bulletinUrl) {
      h.editAsText().setLinkUrl(0, headingText.length - 1, s.bulletinUrl);
    }
    insertIdx++;

    if (s.theme) {
      body.insertListItem(insertIdx, 'Theme: ' + s.theme)
        .setGlyphType(DocumentApp.GlyphType.BULLET);
      insertIdx++;
    }
    if (s.notes) {
      body.insertListItem(insertIdx, 'Notes: ' + s.notes)
        .setGlyphType(DocumentApp.GlyphType.BULLET);
      insertIdx++;
    }
  });
}

// Replaces {weeklyEvents} with an H5 heading + bullet list for each day that has events.
// Days with no events are skipped entirely.
function insertWeeklyEventsSection(body, days, agendaDate) {
  var found = body.findText('\\{weeklyEvents\\}');
  if (!found) return;

  var placeholder = found.getElement().getParent();
  var insertIdx   = body.getChildIndex(placeholder);
  body.removeChild(placeholder);

  var weekStart  = getWeekStart(agendaDate);
  var dayOffsets = {
    monday: 0, tuesday: 1, wednesday: 2, thursday: 3, friday: 4, saturday: 5, sunday: 6,
    mon: 0,    tue: 1,     wed: 2,       thu: 3,      fri: 4,    sat: 5,      sun: 6
  };

  days.forEach(function(day) {
    if (!day.events || day.events.length === 0) return;

    var offset  = dayOffsets[(day.day || '').toLowerCase()];
    var dayDate = offset !== undefined
      ? new Date(weekStart.getTime() + offset * 86400000)
      : null;

    var headingText = dayDate
      ? Utilities.formatDate(dayDate, Session.getScriptTimeZone(), 'EEE, MMM d')
      : day.day;

    var h = body.insertParagraph(insertIdx, headingText);
    h.setHeading(DocumentApp.ParagraphHeading.HEADING5);
    insertIdx++;

    day.events.forEach(function(evt) {
      body.insertListItem(insertIdx, evt.title + ' — ' + evt.time)
        .setGlyphType(DocumentApp.GlyphType.BULLET);
      insertIdx++;
    });
  });
}

// Returns the 1-based sheet row of an existing Staff meeting for this date, or 0 if none.
function checkStaffAgendaDuplicate(dateStr) {
  var meetingDate = parseDateStr(dateStr);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('meetings');
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0];
    var rowType = data[i][1];
    if (rowType === 'Staff' && rowDate instanceof Date &&
        rowDate.toDateString() === meetingDate.toDateString()) {
      return i + 1; // 1-based row number
    }
  }
  return 0;
}

function logMeeting(date, type, docUrl, overwriteRow) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('meetings');
  if (overwriteRow) {
    sheet.getRange(overwriteRow, 1, 1, 4).setValues([[date, type, docUrl, '']]);
  } else {
    sheet.appendRow([date, type, docUrl, '']);
  }
}
