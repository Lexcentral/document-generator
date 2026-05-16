function showDeaconAgendaDialog() {
  var tmpl = HtmlService.createTemplateFromFile('DialogDeaconAgenda');
  tmpl.meetings = JSON.stringify(getLeadershipMeetings('Deacon'));
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(420).setHeight(200),
    'Generate Deacon Agenda'
  );
}

// Returns the existing agenda URL if one already exists for this Deacon meeting, or ''.
function checkDeaconAgendaDuplicate(dateStr) {
  return getLeadershipMeetingAgenda(dateStr, 'Deacon');
}

function generateDeaconAgenda(dateStr, timeStr) {
  var meetingDate = parseDateStr(dateStr);
  var sundays     = getWorshipSundays(meetingDate);
  var zoomLink    = getMeetingZoomLink(dateStr, 'Deacon');
  var dateTimeStr = formatPacketDateTime(dateStr, timeStr);
  var staffNames  = getStaffForMeeting('Deacon');

  var folder  = DriveApp.getFolderById(FOLDER_IDS.deaconAgenda);
  var docFile = DriveApp.getFileById(TEMPLATE_IDS.deaconAgenda).makeCopy(folder);
  docFile.setName(dateStr + '_deacon-agenda');
  docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var doc  = DocumentApp.openById(docFile.getId());
  var body = doc.getBody();

  // Simple text replacements first — structural inserts below can interfere with footer ops.
  body.replaceText('\\{meetingDate\\}', formatMeetingDate(meetingDate));
  body.replaceText('\\{dateTime\\}',    dateTimeStr);
  replaceWithHyperlink(body, '\\{zoomLink\\}', 'Zoom Link', zoomLink);
  var footer = doc.getFooter();
  if (footer) {
    footer.replaceText('\\{meetingDate\\}', formatMeetingDate(meetingDate));
    footer.replaceText('\\{dateTime\\}',    dateTimeStr);
    replaceWithHyperlink(footer, '\\{zoomLink\\}', 'Zoom Link', zoomLink);
  }

  // Structural inserts after all simple replacements.
  insertBulletSection(body, '\\{staff\\}', staffNames);
  insertWorshipSection(body, sundays);
  doc.saveAndClose();

  var docUrl = docFile.getUrl();
  updateLeadershipMeetingAgenda(dateStr, 'Deacon', docUrl);
  return docUrl;
}

// Returns the existing agenda URL (column E) for a given date+type, or ''.
function getLeadershipMeetingAgenda(dateStr, type) {
  var meetingDate = parseDateStr(dateStr);
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('leadership_meetings').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0], rowType = data[i][2];
    if (rowType === type && rowDate instanceof Date &&
        rowDate.toDateString() === meetingDate.toDateString()) {
      return data[i][3] ? String(data[i][3]) : ''; // column D = agenda
    }
  }
  return '';
}

// Returns names of staff whose meetings column includes the given meeting type.
function getStaffForMeeting(type) {
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('_staff').getDataRange().getValues();
  var names = [];
  for (var i = 1; i < data.length; i++) {
    var name     = data[i][0];
    var meetings = (data[i][1] || '').split(',').map(function(m) { return m.trim(); });
    if (name && meetings.indexOf(type) !== -1) names.push(String(name));
  }
  return names;
}

// Writes the generated doc URL into the agenda column (E) of the matching leadership_meetings row.
function updateLeadershipMeetingAgenda(dateStr, type, url) {
  var meetingDate = parseDateStr(dateStr);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('leadership_meetings');
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0], rowType = data[i][2];
    if (rowType === type && rowDate instanceof Date &&
        rowDate.toDateString() === meetingDate.toDateString()) {
      sheet.getRange(i + 1, 4).setValue(url); // column D = agenda
      return;
    }
  }
}
