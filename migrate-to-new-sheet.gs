/* ═══════════════════════════════════════════════════════════════
   ██  AL RASA — ONE-OFF MIGRATION into a fresh spreadsheet
   ██  Paste as a SECOND file in the NEW Al Rasa Apps Script project
   ██  (Code.gs must already be pasted there too), then run migrate().
   ██  Safe to re-run: it skips any date already present.
   ██  Delete this file once the numbers check out.
   ═══════════════════════════════════════════════════════════════ */

// Everything the old database actually held. Sales, collection, cheques and
// the invoice audit log were all empty, so there is nothing else to carry.
var OLD_INVOICES = [
  // date          al        an     ho         updatedAt
  ['2026-07-31', 1203.75,    0, 40575.86, '2026-07-31 18:09'],
  ['2026-08-01', 2586,     315,  7201,    '2026-08-03 18:13'],
  ['2026-08-02', 1575,       0,     0,    '2026-08-03 18:16'],
  ['2026-08-03', 1270.75,    0, 16065,    '2026-08-03 18:14'],
  ['2026-08-04', 1667.5,     0,  5172,    '2026-08-04 18:14']
];

var OLD_CONFIG = { openingBalance: 500000, openingDate: '2026-08-01' };

function migrate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheets(ss);                     // creates the tabs and seeds the 13 logins

  // Guard: never run this against the Krew database.
  if (ss.getSheetByName('entries') && !ss.getSheetByName('invoices')) {
    Logger.log('STOP — this looks like the Krew WorkLog database. Nothing written.');
    return;
  }

  var sh = ss.getSheetByName('invoices');
  var headers = headersOf(sh, 'invoices');
  var col = function (h) { return headers.indexOf(h); };
  var tz = ss.getSpreadsheetTimeZone();

  var existing = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    existing[dateStr(r[col('date')], tz)] = true;
  });

  var added = 0;
  OLD_INVOICES.forEach(function (inv) {
    if (existing[inv[0]]) { Logger.log('skip %s — already there', inv[0]); return; }
    var obj = {
      id: 'inv_mig_' + inv[0].replace(/-/g, ''),
      date: inv[0], al: inv[1], an: inv[2], ho: inv[3],
      note: '', updatedBy: 'Neethu', updatedAt: inv[4]
    };
    var r = sh.getLastRow() + 1;
    headers.forEach(function (h, i) {
      var cell = sh.getRange(r, i + 1);
      // Dates and timestamps must be stored as text or Sheets turns them into
      // serial numbers and the app reads back 1899 dates.
      if (h === 'date' || h === 'updatedAt') cell.setNumberFormat('@');
      cell.setValue(obj[h] !== undefined ? obj[h] : '');
    });
    added++;
  });

  // config — opening balance and start date
  var cfg = ss.getSheetByName('config');
  var vals = cfg.getDataRange().getValues();
  Object.keys(OLD_CONFIG).forEach(function (k) {
    var row = -1;
    for (var i = 1; i < vals.length; i++) if (String(vals[i][0]) === k) { row = i + 1; break; }
    if (row < 0) { row = cfg.getLastRow() + 1; cfg.getRange(row, 1).setValue(k); }
    var c = cfg.getRange(row, 2);
    if (k === 'openingDate') c.setNumberFormat('@');
    c.setValue(OLD_CONFIG[k]);
  });

  Logger.log('Invoices added: %s', added);
  Logger.log('Config set: openingBalance=%s, openingDate=%s',
             OLD_CONFIG.openingBalance, OLD_CONFIG.openingDate);
  Logger.log('Tabs now: %s', ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
  Logger.log('Check the To Be Collected page: opening 500,000 + invoiced - collected.');
}
