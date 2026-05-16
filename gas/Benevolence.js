// ── Benevolence fund reporting & verification ─────────────────────────────────

function runBenevolenceCheck() {
  var ui = SpreadsheetApp.getUi();

  // 1. Last Business meeting date from leadership_meetings (source of truth).
  var lastMeetingDate = getLastBusinessMeetingDate();
  if (!lastMeetingDate) {
    ui.alert('No past Business meeting found in leadership_meetings tab.');
    return;
  }

  // 2. Calculate figures from the Checkbook.
  var calc = calculateFromCheckbook(lastMeetingDate);

  // 3. Internal consistency check: beginning + deposited - expended = ending = last balance.
  var calcEnding = round2(calc.beginning + calc.deposited - calc.expended);
  var errors = [];

  if (calc.lastBalance !== null && Math.abs(calcEnding - calc.lastBalance) > 0.02) {
    errors.push(
      'Internal check failed:\n' +
      '  Beginning + Deposited - Expended = ' + fmt(calcEnding) + '\n' +
      '  Last checkbook balance            = ' + fmt(calc.lastBalance) + '\n' +
      '  Difference: ' + fmt(Math.abs(calcEnding - calc.lastBalance))
    );
  }

  // 4. Cross-check against Stats sheet formulas.
  var sheet = getStatsSheetValues();
  function crossCheck(label, gasVal, sheetVal) {
    if (sheetVal === null) return;
    if (Math.abs(gasVal - sheetVal) > 0.02) {
      errors.push(label + ': GAS=' + fmt(gasVal) + '  Sheet=' + fmt(sheetVal) +
        '  Diff=' + fmt(Math.abs(gasVal - sheetVal)));
    }
  }
  crossCheck('Beginning Balance', calc.beginning,  sheet.beginning);
  crossCheck('Deposited',         calc.deposited,  sheet.deposited);
  crossCheck('Expended',          calc.expended,   sheet.expended);
  crossCheck('Ending Balance',    calcEnding,      sheet.ending);

  // 5. Build summary text.
  var tz        = Session.getScriptTimeZone();
  var dateLabel = Utilities.formatDate(lastMeetingDate, tz, 'MMMM d, yyyy');

  var catLines = Object.keys(calc.categories).sort().map(function(cat) {
    return '    ' + cat + ': ' + fmt(calc.categories[cat]);
  });

  var summary =
    'Period since last Business Meeting: ' + dateLabel + '\n\n' +
    'Beginning Balance:  ' + fmt(calc.beginning)  + '\n' +
    'Deposited:          ' + fmt(calc.deposited)  + '\n' +
    'Expended:           ' + fmt(calc.expended)   + '\n' +
    'Ending Balance:     ' + fmt(calcEnding)      + '\n\n' +
    'Assistance Detail:\n' + catLines.join('\n');

  if (errors.length > 0) {
    ui.alert('⚠️  Benevolence figures do not reconcile\n\n' +
      errors.join('\n\n') + '\n\n──────────────\n' + summary);
  } else {
    ui.alert('✓  All benevolence figures reconcile\n\n' + summary);
  }
}

// Returns the date of the most recent past Business meeting from leadership_meetings.
function getLastBusinessMeetingDate() {
  var today = new Date();
  today.setHours(23, 59, 59, 0);
  var data  = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('leadership_meetings').getDataRange().getValues();
  var last  = null;
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0], rowType = data[i][2];
    if (rowType !== 'Business' || !(rowDate instanceof Date)) continue;
    var d = new Date(rowDate);
    if (d <= today && (!last || d > last)) last = d;
  }
  if (last) last.setHours(12, 0, 0, 0);
  return last;
}

// Reads the Checkbook and calculates figures for the period after lastMeetingDate.
function calculateFromCheckbook(lastMeetingDate) {
  var ss     = SpreadsheetApp.openById(SPREADSHEET_IDS.benevolence);
  var data   = ss.getSheetByName('Checkbook').getDataRange().getValues();
  // Columns: 0=Check#, 1=Date, 2=Description, 3=Category,
  //          4=Debit, 5=Credit, 6=Balance, 7=Transacted

  var cutoff     = new Date(lastMeetingDate);
  cutoff.setHours(12, 0, 0, 0);
  var beginning  = 0;
  var deposited  = 0;
  var expended   = 0;
  var categories = {};
  var lastBalance = null;

  for (var i = 1; i < data.length; i++) {
    var row        = data[i];
    var date       = row[1];
    var category   = String(row[3] || '').trim();
    var debit      = Number(row[4]) || 0;
    var credit     = Number(row[5]) || 0;
    var balance    = row[6];
    var transacted = row[7];

    if (!(date instanceof Date) || transacted !== true) continue;

    var d = new Date(date);
    d.setHours(12, 0, 0, 0);

    if (d <= cutoff) {
      // Track last transacted balance on or before the meeting date.
      if (typeof balance === 'number') beginning = balance;
    } else {
      // This period: after the last business meeting.
      if (category === 'Deposit') {
        deposited += credit;
      } else if (category) {
        expended += debit;
        categories[category] = (categories[category] || 0) + debit;
      }
      if (typeof balance === 'number') lastBalance = balance;
    }
  }

  var rounded = {};
  Object.keys(categories).forEach(function(k) { rounded[k] = round2(categories[k]); });

  return {
    beginning:   round2(beginning),
    deposited:   round2(deposited),
    expended:    round2(expended),
    categories:  rounded,
    lastBalance: lastBalance !== null ? round2(lastBalance) : null
  };
}

// Reads the Stats sheet A3:B20 for the values calculated by sheet formulas.
function getStatsSheetValues() {
  var data   = SpreadsheetApp.openById(SPREADSHEET_IDS.benevolence)
    .getSheetByName('Stats').getRange('A3:B20').getValues();
  var result = { beginning: null, deposited: null, expended: null, ending: null };
  for (var i = 0; i < data.length; i++) {
    var label = String(data[i][0]).trim().toLowerCase();
    var val   = data[i][1];
    if (typeof val !== 'number') continue;
    if (label === 'beginning balance') result.beginning = val;
    else if (label === 'deposited')    result.deposited = val;
    else if (label === 'expended')     result.expended  = val;
    else if (label === 'ending balance') result.ending  = val;
  }
  return result;
}

// Returns structured benevolence lines for document insertion.
// Each entry: { text: string, level: 0|1 } where level 1 = indented under Assistance Detail.
function getBenevolenceLines(lastMeetingDate) {
  var calc      = calculateFromCheckbook(lastMeetingDate);
  var calcEnding = round2(calc.beginning + calc.deposited - calc.expended);
  var lines = [
    { text: 'Beginning Balance: ' + fmt(calc.beginning), level: 0 },
    { text: 'Deposited: '         + fmt(calc.deposited), level: 0 },
    { text: 'Expended: '          + fmt(calc.expended),  level: 0 },
    { text: 'Ending Balance: '    + fmt(calcEnding),     level: 0 },
    { text: 'Assistance Detail',                         level: 0 }
  ];
  Object.keys(calc.categories).sort().forEach(function(cat) {
    lines.push({ text: cat + ': ' + fmt(calc.categories[cat]), level: 1 });
  });
  return lines;
}

function fmt(n) {
  return '$' + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function round2(n) {
  return Math.round(n * 100) / 100;
}
