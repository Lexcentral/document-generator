function showCouncilPacketDialog() {
  var tmpl = HtmlService.createTemplateFromFile('DialogMeetingPacket');
  tmpl.meetingType = 'Council';
  tmpl.meetings    = JSON.stringify(getLeadershipMeetings('Council'));
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(420).setHeight(260),
    'Generate Church Council Packet'
  );
}

function showBusinessPacketDialog() {
  var tmpl = HtmlService.createTemplateFromFile('DialogMeetingPacket');
  tmpl.meetingType = 'Business';
  tmpl.meetings    = JSON.stringify(getLeadershipMeetings('Business'));
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(420).setHeight(260),
    'Generate Business Meeting Packet'
  );
}

// Returns all meetings of the given type from leadership_meetings, sorted by date.
// Each entry: { dateStr: 'yyyy-MM-dd', label: 'March 2, 2026', time24h: '19:00' }
function getLeadershipMeetings(type) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var data  = ss.getSheetByName('leadership_meetings').getDataRange().getValues();
  var tz    = Session.getScriptTimeZone();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0], rowTime = data[i][1], rowType = data[i][2];
    if (rowType !== type || !(rowDate instanceof Date)) continue;
    items.push({
      dateStr: Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd'),
      label:   Utilities.formatDate(rowDate, tz, 'MMMM d, yyyy'),
      time24h: to24h(rowTime)
    });
  }
  items.sort(function(a, b) { return a.dateStr < b.dateStr ? -1 : 1; });
  return items;
}

// Converts a time value to "HH:MM" for use in <input type="time">.
// Handles Date objects (returned by Sheets) and "7:00 PM" strings.
function to24h(timeVal) {
  if (timeVal instanceof Date) {
    return Utilities.formatDate(timeVal, Session.getScriptTimeZone(), 'HH:mm');
  }
  var m = String(timeVal).match(/(\d+):(\d+)\s*(AM|PM)/i);
  if (!m) return '19:00';
  var h = parseInt(m[1], 10), min = m[2], ampm = m[3].toUpperCase();
  if (ampm === 'PM' && h !== 12) h += 12;
  if (ampm === 'AM' && h === 12) h = 0;
  return (h < 10 ? '0' : '') + h + ':' + min;
}

// Returns the existing agenda URL if one already exists for this meeting, or ''.
function checkMeetingPacketDuplicate(dateStr, type) {
  return getLeadershipMeetingAgenda(dateStr, type);
}

function generateMeetingPacket(dateStr, timeStr, meetingType) {
  var meetingDate    = parseDateStr(dateStr);
  var committees     = getCommitteesAndChairs();
  var staffNames     = getStaffNames();
  var lastBizMeeting = getLastBusinessMeetingDate();

  // Fetch clerk-edited documents from the folder before creating the new doc.
  var councilClerkDoc = fetchClerkDoc('council',  meetingDate);
  var bizClerkDoc     = fetchClerkDoc('business', meetingDate);

  var folderId   = meetingType === 'Council' ? FOLDER_IDS.councilPacket   : FOLDER_IDS.businessPacket;
  var templateId = meetingType === 'Council' ? TEMPLATE_IDS.councilPacket : TEMPLATE_IDS.businessPacket;

  var folder  = DriveApp.getFolderById(folderId);
  var docFile = DriveApp.getFileById(templateId).makeCopy(folder);
  docFile.setName(dateStr + '_' + meetingType.toLowerCase() + '-packet');
  docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var doc  = DocumentApp.openById(docFile.getId());
  var body = doc.getBody();

  var typeLabel   = meetingType === 'Council' ? 'Church Council' : 'Business Meeting';
  var dateTimeStr = formatPacketDateTime(dateStr, timeStr);
  var zoomLink    = getMeetingZoomLink(dateStr, meetingType);

  // Simple token replacements first (must precede structural insertions).
  body.replaceText('\\{meetingType\\}', typeLabel);
  body.replaceText('\\{dateTime\\}',    dateTimeStr);
  replaceWithHyperlink(body, '\\{zoomLink\\}', 'Zoom Link', zoomLink);
  var footer = doc.getFooter();
  if (footer) {
    footer.replaceText('\\{meetingType\\}', typeLabel);
    footer.replaceText('\\{dateTime\\}',    dateTimeStr);
    replaceWithHyperlink(footer, '\\{zoomLink\\}', 'Zoom Link', zoomLink);
  }

  // Structural insertions.
  insertBulletSection(body, '\\{committeesAndChairs\\}', committees);
  insertBulletSection(body, '\\{staff\\}',               staffNames);
  insertBenevolenceSection(body, lastBizMeeting);

  if (meetingType === 'Council') {
    // Business minutes from the most recent _business clerk file (section 1).
    // Council minutes from the most recent _council clerk file (section 1).
    var prevBiz     = bizClerkDoc     ? findSectionByKeyword(bizClerkDoc.sections,     'Business') : null;
    var prevCouncil = councilClerkDoc ? findSectionByKeyword(councilClerkDoc.sections, 'Council')  : null;
    insertPreviousMinutes(body, '\\{previousBusinessMeetingMinutes\\}', prevBiz,     'Business Meeting');
    insertPreviousMinutes(body, '\\{previousChurchCouncilMinutes\\}',   prevCouncil, 'Council');
    insertCommitteeReportsFromSheet(body, dateStr);
  } else {
    // Business packet: everything comes from the most recent _council clerk file.
    // Section 1 = council minutes; section 2 = previous business minutes embedded by clerk.
    var prevCouncil = councilClerkDoc ? findSectionByKeyword(councilClerkDoc.sections, 'Council')  : null;
    var prevBiz     = councilClerkDoc ? findSectionByKeyword(councilClerkDoc.sections, 'Business') : null;
    insertPreviousMinutes(body, '\\{previousBusinessMeetingMinutes\\}', prevBiz,     'Business Meeting');
    insertPreviousMinutes(body, '\\{previousChurchCouncilMinutes\\}',   prevCouncil, 'Council');
    insertCommitteeReports(body, councilClerkDoc ? councilClerkDoc.reports : null);
  }

  doc.saveAndClose();

  var docUrl = docFile.getUrl();
  updateLeadershipMeetingAgenda(dateStr, meetingType, docUrl);
  return docUrl;
}

// Looks up the Zoom link for a meeting type from the reminders tab.
function getMeetingZoomLink(dateStr, meetingType) {
  var config = readMeetingConfig();
  return (config[meetingType] || {}).zoomLink || '';
}

// Formats "2026-03-02" + "19:00" → "March 2, 2026, 7 p.m."
function formatPacketDateTime(dateStr, timeStr) {
  var date     = parseDateStr(dateStr);
  var datePart = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  var parts    = timeStr.split(':');
  var hours    = parseInt(parts[0], 10);
  var mins     = parseInt(parts[1] || '0', 10);
  var ampm     = hours >= 12 ? 'p.m.' : 'a.m.';
  var h        = hours > 12 ? hours - 12 : (hours === 0 ? 12 : hours);
  var timePart = mins === 0
    ? h + ' ' + ampm
    : h + ':' + (mins < 10 ? '0' : '') + mins + ' ' + ampm;
  return datePart + ', ' + timePart;
}

// Returns array of "Committee Name: Chair[, Co-Chair[, Co-Chair2]]" strings, in sheet order.
// Skips rows whose name begins with ~ (those are officer contacts, not committees).
function getCommitteesAndChairs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_committees');
  var data  = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    if (!name || String(name).charAt(0) === '~') continue;
    var line = name + ': ' + data[i][1];
    if (data[i][2]) line += ', ' + data[i][2];
    if (data[i][3]) line += ', ' + data[i][3];
    result.push(line);
  }
  return result;
}

// Returns array of staff names from _staff tab col A.
function getStaffNames() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_staff');
  var data  = sheet.getDataRange().getValues();
  var names = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) names.push(String(data[i][0]));
  }
  return names;
}

// Inserts the benevolence summary at {benevolenceData}, with category items indented one level.
function insertBenevolenceSection(body, lastMeetingDate) {
  var found = body.findText('\\{benevolenceData\\}');
  if (!found) return;
  var placeholder = found.getElement().getParent();
  var insertIdx   = body.getChildIndex(placeholder);
  var baseLevel   = placeholder.getType() === DocumentApp.ElementType.LIST_ITEM
    ? placeholder.getNestingLevel() : 0;
  body.removeChild(placeholder);

  var lines = lastMeetingDate ? getBenevolenceLines(lastMeetingDate) : [];
  lines.forEach(function(line) {
    var level = baseLevel + line.level;
    var glyph = GLYPH_BY_LEVEL[level] || DocumentApp.GlyphType.BULLET;
    var item  = body.insertListItem(insertIdx, line.text);
    var attrs = {};
    attrs[DocumentApp.Attribute.NESTING_LEVEL] = level;
    attrs[DocumentApp.Attribute.GLYPH_TYPE]    = glyph;
    item.setAttributes(attrs);
    var text = item.editAsText();
    text.setForegroundColor(null);
    text.setBackgroundColor(null);
    text.setLinkUrl(null);
    insertIdx++;
  });
}

// Finds a token in a body/footer, replaces it with linkText, and makes it a hyperlink.
// Falls back to plain-text replacement if url is empty.
function replaceWithHyperlink(element, tokenRegex, linkText, url) {
  var found = element.findText(tokenRegex);
  if (!found) return;
  if (!url) {
    element.replaceText(tokenRegex, linkText);
    return;
  }
  var textEl = found.getElement();
  var start  = found.getStartOffset();
  var end    = found.getEndOffsetInclusive();
  textEl.deleteText(start, end);
  textEl.insertText(start, linkText);
  textEl.setLinkUrl(start, start + linkText.length - 1, url);
}

// ── Clerk-edited minutes helpers ───────────────────────────────────────────────
//
// Clerk submits files named YYYY-MM-DD_council or YYYY-MM-DD_business to the
// editedMinutes folder. Each file contains:
//   • Meeting sections (plain paragraphs, each starting with "Central Baptist Church")
//   • Committee report sections (H1 headings, after the meeting sections)

// Finds the most recent file of the given type whose date is before `beforeDate`.
// Returns { file, dateStr:'YYYY-MM-DD' } or null.
function findMostRecentEditedFile(type, beforeDate) {
  var cutoff = Utilities.formatDate(beforeDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var iter   = DriveApp.getFolderById(FOLDER_IDS.editedMinutes).getFiles();
  var best = null, bestDate = '';
  while (iter.hasNext()) {
    var f    = iter.next();
    var name = f.getName().replace(/\.(docx?|pdf)$/i, '').toLowerCase();
    var dateStr = name.substring(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) continue;
    if (name.indexOf('_' + type) === -1) continue; // type must appear after the date
    if (dateStr >= cutoff) continue;
    if (!best || dateStr > bestDate) { bestDate = dateStr; best = f; }
  }
  return best ? { file: best, dateStr: bestDate } : null;
}

// Opens a clerk doc and splits it into meeting sections and committee report sections.
// Validates that section 1's heading contains the expected date from the filename.
// Returns { sections:[{heading,elements}], reports:[{heading,elements}] } or throws.
function parseClerkDoc(docId, expectedDateStr) {
  var body      = DocumentApp.openById(docId).getBody();
  var n         = body.getNumChildren();
  var sections  = []; // meeting minutes blocks
  var reports   = []; // H1 committee report blocks
  var curSect   = null, curReport = null, inReports = false;

  for (var i = 0; i < n; i++) {
    var elem = body.getChild(i);
    var type = elem.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      var para    = elem.asParagraph();
      var text    = para.getText().trim();
      var heading = para.getHeading();

      if (heading === DocumentApp.ParagraphHeading.HEADING1) {
        inReports  = true;
        curReport  = { heading: text, elements: [elem] };
        reports.push(curReport);
        continue;
      }
      if (inReports) { if (curReport) curReport.elements.push(elem); continue; }

      if (text.startsWith('Central Baptist Church')) {
        curSect = { heading: text, elements: [elem] };
        sections.push(curSect);
        continue;
      }
    } else {
      if (inReports) { if (curReport) curReport.elements.push(elem); continue; }
    }
    if (curSect) curSect.elements.push(elem);
  }

  // Date verification on section 1.
  if (sections.length && expectedDateStr) {
    var d = parseDateStr(expectedDateStr);
    var expectedShort = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM d');
    var h1 = sections[0].heading.replace(/[\n]/g, ' ');
    if (h1.indexOf(expectedShort) === -1) {
      throw new Error(
        'Date mismatch in clerk document.\n' +
        'Expected date: ' + expectedShort + '\n' +
        'Section 1 heading: "' + h1.substring(0, 120) + '"'
      );
    }
  }
  return { sections: sections, reports: reports };
}

// Fetches and parses the most recent clerk-edited file of `type` before `beforeDate`.
// Returns parsed doc data or null if no file exists yet.
function fetchClerkDoc(type, beforeDate) {
  var found = findMostRecentEditedFile(type, beforeDate);
  if (!found) return null;
  return parseClerkDoc(found.file.getId(), found.dateStr);
}

// Returns the first section whose heading (case-insensitive) contains `keyword`.
function findSectionByKeyword(sections, keyword) {
  keyword = keyword.toLowerCase();
  for (var i = 0; i < sections.length; i++) {
    if (sections[i].heading.toLowerCase().indexOf(keyword) !== -1) return sections[i];
  }
  return null;
}

// Inserts a meeting section's elements at the given placeholder token.
// Inserts an italic placeholder note if section is null.
function insertPreviousMinutes(body, tokenRegex, section, label) {
  var found = body.findText(tokenRegex);
  if (!found) return;
  var placeholder = found.getElement().getParent();
  var insertIdx   = body.getChildIndex(placeholder);
  body.removeChild(placeholder);

  if (!section) {
    body.insertParagraph(insertIdx, '[Previous ' + label + ' minutes not yet available]')
      .editAsText().setItalic(true);
    return;
  }

  section.elements.forEach(function(elem) {
    var t = elem.getType();
    try {
      if (t === DocumentApp.ElementType.PARAGRAPH)  body.insertParagraph(insertIdx, elem.copy());
      else if (t === DocumentApp.ElementType.LIST_ITEM) body.insertListItem(insertIdx, elem.copy());
      else if (t === DocumentApp.ElementType.TABLE)  body.insertTable(insertIdx, elem.copy());
      else return;
      insertIdx++;
    } catch (e) { /* skip elements that can't be copied cross-document */ }
  });
}

// Inserts committee report sections (H1 blocks) at the {committeeReports} placeholder.
// Used for business meeting packets; council packets leave the token as a Feature 3 placeholder.
function insertCommitteeReports(body, reports) {
  var found = body.findText('\\{committeeReports\\}');
  if (!found) return;
  var placeholder = found.getElement().getParent();
  var insertIdx   = body.getChildIndex(placeholder);
  body.removeChild(placeholder);

  if (!reports || !reports.length) {
    body.insertParagraph(insertIdx, '[No committee reports available]')
      .editAsText().setItalic(true);
    return;
  }

  reports.forEach(function(report) {
    report.elements.forEach(function(elem) {
      var t = elem.getType();
      try {
        if (t === DocumentApp.ElementType.PARAGRAPH)  body.insertParagraph(insertIdx, elem.copy());
        else if (t === DocumentApp.ElementType.LIST_ITEM) body.insertListItem(insertIdx, elem.copy());
        else if (t === DocumentApp.ElementType.TABLE)  body.insertTable(insertIdx, elem.copy());
        else return;
        insertIdx++;
      } catch (e) { /* skip */ }
    });
  });
}

// Replaces {committeeReports} in a council packet with reports submitted via the web form,
// filtered to the council month matching the packet date (e.g. "June 2026").
function insertCommitteeReportsFromSheet(body, councilDateStr) {
  var found = body.findText('\\{committeeReports\\}');
  if (!found) return;
  var placeholder = found.getElement().getParent();
  var insertIdx   = body.getChildIndex(placeholder);
  body.removeChild(placeholder);

  var d       = parseDateStr(councilDateStr);
  var target  = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM yyyy');
  var reports = getCommitteeReportsForTarget(target);

  if (!reports.length) {
    body.insertParagraph(insertIdx, '[No committee reports submitted for ' + target + ']')
      .editAsText().setItalic(true);
    return;
  }

  reports.forEach(function(report) {
    var h = body.insertParagraph(insertIdx++, report.committee);
    h.setHeading(DocumentApp.ParagraphHeading.HEADING1);

    var byline = body.insertParagraph(insertIdx++, 'Submitted by: ' + report.authorName);
    byline.editAsText().setItalic(true);

    insertIdx = insertQuillHtmlIntoDoc(body, report.bodyHTML, insertIdx);

    body.insertParagraph(insertIdx++, '');
  });
}

var GLYPH_BY_LEVEL = [
  DocumentApp.GlyphType.BULLET,
  DocumentApp.GlyphType.HOLLOW_BULLET,
  DocumentApp.GlyphType.SQUARE_BULLET
];

// Finds a placeholder token in the body, removes it, and inserts lines as bullets at that index,
// preserving the nesting level the placeholder had in the template.
function insertBulletSection(body, tokenRegex, lines) {
  var found = body.findText(tokenRegex);
  if (!found) return;
  var placeholder  = found.getElement().getParent();
  var insertIdx    = body.getChildIndex(placeholder);
  var nestingLevel = placeholder.getType() === DocumentApp.ElementType.LIST_ITEM
    ? placeholder.getNestingLevel()
    : 0;
  var glyphType = GLYPH_BY_LEVEL[nestingLevel] || DocumentApp.GlyphType.BULLET;
  body.removeChild(placeholder);
  lines.forEach(function(line) {
    var item  = body.insertListItem(insertIdx, line);
    var attrs = {};
    attrs[DocumentApp.Attribute.NESTING_LEVEL] = nestingLevel;
    attrs[DocumentApp.Attribute.GLYPH_TYPE]    = glyphType;
    item.setAttributes(attrs);
    var text = item.editAsText();
    text.setForegroundColor(null);
    text.setBackgroundColor(null);
    text.setLinkUrl(null);
    insertIdx++;
  });
}
