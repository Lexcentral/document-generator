// Parses a YYYY-MM-DD string into a local-noon Date to avoid UTC/timezone edge cases.
function parseDateStr(dateStr) {
  var p = dateStr.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2], 12, 0, 0);
}

function formatMeetingDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
}

function formatSundayDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM d');
}

// Returns the Monday of the week containing the given date, at noon.
function getWeekStart(date) {
  var d = new Date(date.getTime());
  var daysFromMonday = (d.getDay() + 6) % 7; // Mon=0 … Sun=6
  d.setDate(d.getDate() - daysFromMonday);
  d.setHours(12, 0, 0, 0);
  return d;
}
