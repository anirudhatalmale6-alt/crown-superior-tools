/**
 * Crown Superior - quotes from the website into the CROWN PHONE sheet.
 *
 * The website sends one quote at a time to this script. The script finds
 * the customer's row and fills in the payment columns the phone system
 * reads.
 *
 * ----------------------------------------------------------------------
 * THIS MUST GO IN A BRAND NEW, EMPTY PROJECT.
 *
 * Not "Crown Customer lookup" and not the one attached to the sheet.
 * Google allows a project only ONE doPost, and Crown Customer lookup
 * already has one - that is your phone system. Two cannot live together.
 * While my file sits in there it is being ignored, which is exactly what
 * was happening.
 *
 * HOW TO SET IT UP  (five minutes, once)
 *
 *  1. First tidy up: open Crown Customer lookup, click the three dots
 *     next to CrownSheet.gs in the Files list, and Delete. Nothing else
 *     in that project gets touched.
 *
 *  2. Go to  script.google.com  and click  New project
 *     It opens with one file, Code.gs, holding an empty myFunction.
 *
 *  3. SELECT ALL of it and delete it, so the editor is completely empty,
 *     THEN paste this file in.
 *     Pasting without deleting first leaves my file inside myFunction,
 *     and then nothing in it exists as a function of its own - which is
 *     why the dropdown only offered myFunction.
 *
 *  4. Name the project Crown quotes, top left.
 *
 *  5. Fill in SHEET_ID below - see the note on it, it is a copy and paste
 *     from your other script, not something to type out.
 *
 *  6. Change SECRET to the word already saved on the website. It is
 *     shown on the website at Get a quote -> Customer Quotes, in the box
 *     "The Google sheet the phone system reads" - copy it exactly,
 *     capital letter and all. If they do not match, nothing is written
 *     and the website says so rather than failing quietly.
 *
 *  7. Deploy -> New deployment
 *       - gear next to "Select type", choose  Web app
 *       - Execute as:        Me
 *       - Who has access:    Anyone
 *       - Deploy, then Authorize access and allow it
 *       - first time it warns you it is unverified: Advanced, then
 *         "Go to Crown quotes (unsafe)". It is your own script.
 *
 *  8. Copy the Web app URL - it ends in /exec - and paste it on the
 *     website: Get a quote, Customer Quotes, "The Google sheet the phone
 *     system reads".
 *
 * Before step 7, run testMe: pick it in the dropdown between Debug and
 * Execution log, press Run, and read what it prints.
 *
 * If you ever change this script: Deploy -> Manage deployments -> pencil
 * -> Version: New version -> Deploy. The address stays the same.
 * ----------------------------------------------------------------------
 */

/**
 * Which spreadsheet to write to.
 *
 * DO NOT TYPE THIS OUT and do not copy it off a screenshot - it is forty
 * random characters and a lower-case L looks exactly like a capital i.
 *
 * Copy it from a script you already know works: open Crown Customer
 * lookup, look at line 18 of Code.gs, and copy the text between the
 * quote marks on that line. Paste it between the quote marks here.
 */
var SHEET_ID = 'PASTE THE ID FROM Code.gs LINE 18 HERE';

/**
 * The password between the website and this script.
 *
 * It must be character for character the one already saved on the
 * website - Get a quote, Customer Quotes, "The Google sheet the phone
 * system reads". Copy it from there rather than typing it.
 */
var SECRET = 'PUT THE WORD FROM THE WEBSITE HERE';

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
  if (SHEET_ID.indexOf('PASTE') === 0) {
    Logger.log('SHEET_ID has not been filled in yet.');
    Logger.log('Open Crown Customer lookup, Code.gs, line 18, and copy the text');
    Logger.log('between the quote marks. Paste it into SHEET_ID at the top of this file.');
    return;
  }

  var book;

  try {
    book = SpreadsheetApp.openById(SHEET_ID);
  } catch (err) {
    Logger.log('Could not open that spreadsheet: ' + err);
    Logger.log('The id is almost certainly mistyped. Copy it from Code.gs line 18');
    Logger.log('rather than typing it - a lower-case L and a capital i look identical.');
    return;
  }

  Logger.log('Opened: ' + book.getName());

  var sheet = book.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log('There is no tab called "' + SHEET_NAME + '". The tabs are:');
    Logger.log('   ' + book.getSheets().map(function (s) { return s.getName(); }).join(', '));
    Logger.log('Set SHEET_NAME at the top of this file to whichever of those it should be.');
    return;
  }

  var headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('Tab "' + SHEET_NAME + '" has ' + (sheet.getLastRow() - 1) + ' rows.');

  var lower = headings.map(function (h) { return String(h).trim().toLowerCase(); });
  var missing = 0;

  for (var name in COLUMNS) {
    if (!COLUMNS.hasOwnProperty(name)) { continue; }

    var at = lower.indexOf(String(COLUMNS[name]).trim().toLowerCase());

    if (at < 0) { missing++; }

    Logger.log((at < 0 ? 'NOT FOUND  ' : 'ok         ') + name + '  ->  ' + COLUMNS[name]);
  }

  if (missing) {
    Logger.log('');
    Logger.log(missing + ' column(s) were not found. Either add them to row 1 of the');
    Logger.log('tab, or change the right-hand side in COLUMNS above to match what you');
    Logger.log('already call them. Anything not found is simply skipped - it will not');
    Logger.log('stop the rest working.');
    Logger.log('');
    Logger.log('Your headings are: ' + headings.join(' | '));
  } else {
    Logger.log('');
    Logger.log('Every column found. Deploy it and paste the address on the website.');
  }
}
