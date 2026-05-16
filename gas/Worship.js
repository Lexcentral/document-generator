// Returns 6 worship rows: the most recent Sunday on or before meetingDate, plus the next 5.
// Each row: { dateFormatted, lectionary, theme, bulletinUrl }
function getWorshipSundays(meetingDate) {
  var worship = SpreadsheetApp.openById(SPREADSHEET_IDS.worship);
  var planData = worship.getSheetByName('planning').getDataRange().getValues();
  var linkData = worship.getSheetByName('_document_links').getDataRange().getValues();

  var meetingTime = meetingDate.getTime();
  var rows = [];

  // Row 0 is headers; start at 1
  for (var i = 1; i < planData.length; i++) {
    var rowDate = planData[i][0]; // col A = date
    if (!(rowDate instanceof Date) || isNaN(rowDate.getTime())) continue;
    rows.push({
      date:       rowDate,
      lectionary: planData[i][3] || '',  // col D
      theme:      planData[i][4] || '',  // col E
      notes:      planData[i][6] || '',  // col G
      bulletinUrl: (linkData[i] && linkData[i][1]) ? linkData[i][1] : ''  // _document_links col B
    });
  }

  rows.sort(function(a, b) { return a.date - b.date; });

  // Find the last row whose date is on or before the meeting date
  var prevIdx = 0;
  for (var j = 0; j < rows.length; j++) {
    if (rows[j].date.getTime() <= meetingTime) {
      prevIdx = j;
    } else {
      break;
    }
  }

  return rows.slice(prevIdx, prevIdx + 6).map(function(row) {
    return {
      dateFormatted: formatSundayDate(row.date),
      lectionary:    row.lectionary,
      theme:         row.theme,
      notes:         row.notes,
      bulletinUrl:   row.bulletinUrl
    };
  });
}
