/**
 * Crown Superior - quotes from the website into the CROWN PHONE sheet.
 *
 * The website sends one quote at a time to this script. The script finds
 * the customer's row and fills in the payment columns the phone system
 * reads.
 *
 * ----------------------------------------------------------------------
 * HOW TO PUT THIS ON YOUR SHEET  (about five minutes, once)
 *
 *  1. Open the CROWN PHONE sheet
 *  2. Extensions  ->  Apps Script
 *  3. DO NOT delete or change anything already in there. Whatever is
 *     already on this sheet is your phone system and it must be left
 *     alone. Instead click the + next to "Files" at the top left and
 *     choose Script, name it CrownSheet, and paste this file into the
 *     NEW empty file.
 *  4. Change SECRET below to any word you like, and change SHEET_NAME
 *     to the name on the tab if it is not "quotes"
 *  5. Click Deploy  ->  New deployment
 *       - the gear next to "Select type", choose  Web app
 *       - Execute as:        Me
 *       - Who has access:    Anyone
 *       - Deploy, then Authorize access and allow it
 *  6. Copy the Web app URL it gives you (it starts
 *     https://script.google.com/macros/s/... and ends /exec)
 *
 * Then on the website: Get a quote -> Customer Quotes, and paste that
 * address and the same secret word into "The Google sheet the phone
 * system reads". That is it.
 *
 * If you ever change this script, Deploy -> Manage deployments -> the
 * pencil -> Version: New version -> Deploy. The address stays the same.
 * ----------------------------------------------------------------------
 */

/** Any word you like. It must match the one saved on the website. */
var SECRET = 'change-me';

/** The tab the quotes are on. */
var SHEET_NAME = 'quotes';

/**
 * Which heading each value goes under.
 *
 * Left side is what the website sends, right side is the exact heading
 * text in row 1 of your sheet. Change the right side to match your
 * headings; anything whose heading is not found is simply skipped, so a
 * wrong name here loses one column rather than breaking the whole thing.
 *
 * To stop a value being written at all, delete its line.
 */
var COLUMNS = {
  down_payment:    'down_payment',
  monthly_payment: 'monthly_payment',
  msg_to_customer: 'msg_to_customer',
  total:           'total_quoted',      // <- tell me your heading for this
  carrier:         'carrier',           // <- and this, if you want it
  result:          'quote_result',      // <- and this
  quote_link:      'quote_link',        // <- and this
  request_id:      'website_request_id' // used to find the row next time
};

/**
 * The column the website's request number is written into.
 *
 * The first push finds the row by e-mail and name and stamps the number
 * here. Every push after that finds the row by the number instead, which
 * cannot be confused by two customers sharing an e-mail address.
 */
var ID_HEADING = 'website_request_id';


function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (String(body.secret || '') !== SECRET) {
      return reply({ ok: false, error: 'wrong secret word' });
    }

    var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

    if (!sheet) {
      return reply({ ok: false, error: 'no tab called ' + SHEET_NAME });
    }

    var headings = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                        .getValues()[0]
                        .map(function (h) { return String(h).trim().toLowerCase(); });

    var row = findRow(sheet, headings, body);

    if (!row) {
      return reply({ ok: false, error: 'no row for ' + (body.email || body.request_id) });
    }

    var wrote = 0;

    for (var name in COLUMNS) {
      if (!COLUMNS.hasOwnProperty(name)) { continue; }

      var at = headings.indexOf(String(COLUMNS[name]).trim().toLowerCase());
      if (at < 0) { continue; }

      var value = body[name];
      if (value === undefined || value === null || value === '') { continue; }

      sheet.getRange(row, at + 1).setValue(value);
      wrote++;
    }

    // Stamp the request number so the next push finds this row exactly.
    var idAt = headings.indexOf(ID_HEADING.trim().toLowerCase());

    if (idAt >= 0 && body.request_id) {
      sheet.getRange(row, idAt + 1).setValue(body.request_id);
    }

    return reply({ ok: true, row: row, wrote: wrote });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  }
}


/**
 * The customer's row.
 *
 * By the request number if it has been stamped before - that is exact.
 * Otherwise by e-mail, and if several rows share one e-mail the last one
 * wins, because that is the most recent request. Name is used only to
 * break a tie, never on its own.
 */
function findRow(sheet, headings, body) {
  var last = sheet.getLastRow();
  if (last < 2) { return 0; }

  var idAt = headings.indexOf(ID_HEADING.trim().toLowerCase());

  if (idAt >= 0 && body.request_id) {
    var ids = sheet.getRange(2, idAt + 1, last - 1, 1).getValues();

    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === String(body.request_id).trim()) {
        return i + 2;
      }
    }
  }

  var mailAt = headings.indexOf('email');
  if (mailAt < 0 || !body.email) { return 0; }

  var mails = sheet.getRange(2, mailAt + 1, last - 1, 1).getValues();
  var wanted = String(body.email).trim().toLowerCase();
  var found = 0;

  for (var j = 0; j < mails.length; j++) {
    if (String(mails[j][0]).trim().toLowerCase() === wanted) {
      found = j + 2;                       // keep going: the last one wins
    }
  }

  return found;
}


function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Run this from the editor to check the script can see your sheet and
 * your headings before you wire the website up to it. Look at the log
 * underneath (View -> Logs).
 */
function testMe() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log('There is no tab called "' + SHEET_NAME + '". The tabs are: '
      + SpreadsheetApp.getActive().getSheets().map(function (s) { return s.getName(); }).join(', '));
    return;
  }

  var headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('Tab "' + SHEET_NAME + '" has ' + (sheet.getLastRow() - 1) + ' rows.');
  Logger.log('Headings: ' + headings.join(' | '));

  var lower = headings.map(function (h) { return String(h).trim().toLowerCase(); });

  for (var name in COLUMNS) {
    if (!COLUMNS.hasOwnProperty(name)) { continue; }
    var at = lower.indexOf(String(COLUMNS[name]).trim().toLowerCase());
    Logger.log((at < 0 ? 'NOT FOUND  ' : 'ok         ') + name + '  ->  ' + COLUMNS[name]);
  }
}
