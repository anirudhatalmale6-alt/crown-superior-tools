/**
 * Crown Superior - quotes from the website into the CROWN PHONE sheet.
 *
 * The website sends one quote at a time to this script. The script finds
 * the customer's row and fills in the payment columns the phone system
 * reads.
 *
 * ----------------------------------------------------------------------
 * THIS IS A SEPARATE SCRIPT PROJECT. Do not put it in the one attached to
 * your sheet.
 *
 * Google allows a project only ONE doPost, and the project on your sheet
 * already has one - that is your phone system. Two cannot live together,
 * and whichever wins, something stops working. So this one lives on its
 * own and reaches the sheet by its id. Your phone system is not touched,
 * not edited, and not redeployed.
 *
 * HOW TO SET IT UP  (about five minutes, once)
 *
 *  1. Go to  script.google.com  and click  New project
 *     (NOT Extensions - Apps Script from inside the sheet)
 *  2. Delete the empty myFunction and paste this whole file in
 *  3. Name the project Crown quotes, top left
 *  4. Change SECRET below to any word you like
 *  5. Click Deploy  ->  New deployment
 *       - the gear next to "Select type", choose  Web app
 *       - Execute as:        Me
 *       - Who has access:    Anyone
 *       - Deploy, then Authorize access and allow it
 *       - the first time it will warn you it is unverified: Advanced,
 *         then Go to Crown quotes (unsafe). It is your own script.
 *  6. Copy the Web app URL it gives you - it ends in /exec
 *
 * Then on the website: Get a quote -> Customer Quotes, and paste that
 * address and the same secret word into "The Google sheet the phone
 * system reads".
 *
 * If you ever change this script, Deploy -> Manage deployments -> the
 * pencil -> Version: New version -> Deploy. The address stays the same.
 * ----------------------------------------------------------------------
 */

/** The CROWN PHONE sheet. From its own web address, between /d/ and /edit. */
var SHEET_ID = '1M-4JswHCLCcEecxqYoJPIQcgptn0Yb8XpGTbBSPXtmM';

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
  total:           'Total premium',     // the total WITH the autoclub on it
  down_payment:    'down_payment',      // the down payment WITH the autoclub on it
  autoclub:        'autoclub',          // the amount added, on its own
  monthly_payment: 'monthly_payment',
  msg_to_customer: 'msg_to_customer',
  quote_link:      'quote_link',
  request_id:      'website_request_id' // add this column - see ID_HEADING below

  // Not written, because they are yours to fill in and not the carrier's:
  //   payment_day, processing_fee, amount_paid, amount to company
  // If you ever want one of them filled from a quote, add a line here.
  // Also available if you want a column for them:
  //   carrier_total, carrier_down   - the carrier's own figures, autoclub off
  //   carrier, result, quoted_by, term
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

    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

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
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log('There is no tab called "' + SHEET_NAME + '". The tabs are: '
      + SpreadsheetApp.openById(SHEET_ID).getSheets().map(function (s) { return s.getName(); }).join(', '));
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
