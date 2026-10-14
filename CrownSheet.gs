/**
 * Crown Superior - quotes from the website into the CROWN PHONE sheet.
 *
 * The website sends one quote at a time to this script. The script finds
 * the customer's row and fills in the payment columns the phone system
 * reads. A customer who is not on the sheet yet gets a new row at the
 * bottom, filled in from their website request under your own headings.
 *
 * UPDATING FROM AN EARLIER VERSION
 *   1. In the editor, copy your two lines SHEET_ID = ... and SECRET = ...
 *      somewhere safe (Notepad).
 *   2. Select all, delete, paste this file in.
 *   3. Put your SHEET_ID and SECRET values back in.
 *   4. Save, then Deploy -> Manage deployments -> pencil ->
 *      Version: New version -> Deploy. The web address stays the same,
 *      so nothing on the website needs changing.
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
 * The headings below are the ones already on his sheet, read off the
 * screenshot he sent - total_premium and auto_club, not the tidier names
 * I had guessed. The sheet is right; the script was wrong.
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
  total:           'total_premium',    // the total WITH the autoclub on it
  down_payment:    'down_payment',      // the down payment WITH the autoclub on it
  autoclub:        'auto_club',         // the amount added, on its own
  monthly_payment: 'monthly_payment',
  msg_to_customer: 'msg_to_customer',
  quote_link:      'quote_link',
  request_id:      'website_request_id' // already on the sheet, column BG

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
    var added = false;

    // A customer who is not on the sheet yet - a request straight off the
    // website. Their row is added at the bottom under YOUR headings (none
    // are renamed or added), then the quote is filled in like any other.
    if (row) {
      // Already on the sheet: fill in any of its EMPTY cells we have an
      // answer for (a vehicle, a second driver). Nothing typed is changed.
      fillBlanks(sheet, row, headings, body);
    } else {
      row = addRow(sheet, headings, body);
      added = true;

      if (!row) {
        return reply({ ok: false, error: 'no row for ' + (body.email || body.request_id) +
                                         ' and nothing to start a new one with' });
      }
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

    return reply({ ok: true, row: row, wrote: wrote, added: added });
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


/**
 * Which website question goes under which of your headings, where the two
 * are not simply the same words. Left: your heading. Right: the question's
 * name on the website's Get a quote form.
 *
 * Anything not listed is matched by its words - "Gender" fills "gender",
 * "Marital Status" fills "marital_status" - so this list only needs the
 * ones worded differently. Add a line here if a column stays empty that
 * should not.
 */
var NEW_ROW_FROM = {
  quote_type:           'Coverage',
  how_did_you_hear:     'Source',
  phone:                'Phone number',
  dob:                  'Date of Birth',
  city:                 'CIty',
  zip:                  'Zip_Code',
  drivers_license:      'Drivers License Number',
  drivers_license_number: 'Drivers License Number',
  license_state:        'DL_state',
  drivers_license_state: 'DL_state',
  industry:             'Industry_',
  job_title:            'JOB_TITLE_',
  previous_insurance:   'Previous insurance'
};

/** "Marital Status " -> "maritalstatus": the words, nothing else. */
function flat(text) {
  return String(text || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

/** The list entry for a heading: exact, or one the heading starts with. */
function aliasFor(aliases, want) {
  if (aliases[want]) { return aliases[want]; }
  for (var a in aliases) {
    // Only for a heading cut short on the sheet, like "how_did_you_he...".
    // A wider rule let drivers_license_state take the licence NUMBER.
    if (aliases.hasOwnProperty(a) && a === 'howdidyouhear' && want.indexOf(a) === 0) { return aliases[a]; }
  }
  return '';
}

/**
 * Add a row for a customer the sheet has not seen.
 *
 * Every heading is looked up in what the website sent: first the list
 * above, then the same words, then - for a heading that was cut short,
 * like "how_did_you_he" - the one question whose name starts with it.
 * A heading nothing matches is left empty. Nothing is ever renamed.
 */
function addRow(sheet, headings, body) {
  var line = rowValues(headings, body);

  if (!line) { return 0; }

  sheet.appendRow(line);

  return sheet.getLastRow();
}


/** Fill the empty cells of an existing row - never one that has something in it. */
function fillBlanks(sheet, row, headings, body) {
  var line = rowValues(headings, body);
  if (!line) { return 0; }

  var range = sheet.getRange(row, 1, 1, headings.length);
  var have = range.getValues()[0];
  var filled = 0;

  for (var c = 0; c < headings.length; c++) {
    if ((have[c] === '' || have[c] === null) && line[c] !== '' && flat(headings[c]) !== 'savedat') {
      have[c] = line[c];
      filled++;
    }
  }

  if (filled) { range.setValues([have]); }

  return filled;
}


/**
 * The numbered questions. The website calls vehicle 1's year just "Year"
 * and vehicle 2's "Year_2nd", so "vehicle_2_year" can never match by its
 * words alone. Each entry: a word looked for in your heading, then the
 * website question for number 1, 2, 3, 4. Checked top to bottom, so
 * "purchase" is tried before "year".
 */
var VEHICLE_QUESTIONS = [
  ['purchase', ['Purchase_Year', 'Purchase year veh 2', 'Purchase year veh 3', 'Purchase year veh 4']],
  ['year',     ['Year', 'Year_2nd', 'Year_3rd', 'Year_4th']],
  ['make',     ['Make', 'Make_2nd', 'Make_3rd', 'Make_4th']],
  ['model',    ['Model', 'Model_2nd', 'Model_3rd', 'Model_4th']],
  ['vin',      ['Vin_Number', 'Vin_Number_2nd', 'Vin_Number_3rd', 'Vin_Number_4th']],
  ['trim',     ['Trim_veh_1', 'Trim_veh_2', 'Trim_veh_3', 'Trim_veh_4']],
  ['milesdrive', ['How_many_miles_do_you_drive_', 'How_many_miles_do_you_drive__veh_2_each_year',
                  'How_many_miles_do_you_drive_veh_3_each_year', 'How_many_miles_do_you_drive_veh_4_each_year']],
  ['mile',     ['Mileage', '', 'Mileage veh 3', 'Mileage veh 4']],
  ['ride',     ['vehicle 1 ridehsare', 'RIdesharing vehicle 2', 'RIdesharing vehicle 3', 'RIdesharing vehicle 4']],
  ['comm',     ['vehicle 1 commercial use', 'Vehicle 2 Commercial use', 'Vehicle 3 Commercial use', 'Vehicle 4 Commercial']],
  ['deduct',   ['Vehicle 1 Deductibles', 'Vehicle 2 Deductibles', 'Vehicle 3 Deductibles', 'Vehicle 4 Deductibles']],
  ['type',     ['Vehicle_type_', 'Vehicle 2 type', 'Vehicle 3 type', 'Vehicle 4 type']],
  ['value',    ['Vehicle_Value_', 'Vehicle 2 Value', 'Vehicle 3 Value', 'Vehicle 4 Value']],
  ['owned',    ['How_long_vehicle_owned_', '', '', '']],
  ['new',      ['New Purchase', '', '', '']]
];

var DRIVER_QUESTIONS = [
  ['first',    ['First Name', 'Driver 2 First', 'Driver 3 First', 'Driver 4 First']],
  ['middle',   ['Middle name', '', '', '']],
  ['last',     ['Last Name', 'Driver 2 Last', 'Driver 3 Last', 'Driver 4 Last']],
  ['dob',      ['Date of Birth', 'Driver 2 dob', 'Driver 3 dob', 'Driver 4 dob']],
  ['birth',    ['Date of Birth', 'Driver 2 dob', 'Driver 3 dob', 'Driver 4 dob']],
  ['gender',   ['Gender', 'driver 2 Gender', 'driver 3 Gender', 'driver 4 Gender']],
  ['marital',  ['Marital Status', 'Driver 2 Marital Status', 'Driver 3 Marital Status', 'Driver 4 Marital Status']],
  ['state',    ['DL_state', 'DL_state_Drv 2', 'DL_state_Drv 3_copy', 'DL_state_Drv 4_copy_copy']],
  ['licen',    ['Drivers License Number', 'Driver 2 DL', 'Driver 3 DL', 'Driver 4 DL']],
  ['dl',       ['Drivers License Number', 'Driver 2 DL', 'Driver 3 DL', 'Driver 4 DL']],
  ['industry', ['Industry_', 'Driver 2 industry', 'Driver 3 industry', 'Driver 4 industry']],
  ['occupa',   ['JOB_TITLE_', 'Driver 2 JOB TITLE', 'Driver  3 JOB TITLE', 'Driver 4 JOB TITLE']],
  ['job',      ['JOB_TITLE_', 'Driver 2 JOB TITLE', 'Driver  3 JOB TITLE', 'Driver 4 JOB TITLE']],
  ['relation', ['', 'Driver_2_relation', 'Driver_3_relation', 'Driver_4_relation_copy']],
  ['exclu',    ['', 'Exclude', 'Exclude  driver 3', 'Exclude  driver 4']],
  ['riding',   ['Years_riding_experience', 'Years riding experience driver 2',
                'Years riding experience driver 3', 'Years riding experience driver 4']]
];

/** "vehicle_2_year" -> the answer to "Year_2nd", or undefined. */
function numbered(want, byFlat) {
  var m = /^(vehicle|veh|car|driver|drv)([1-4])(.*)$/.exec(want);
  if (!m) { return undefined; }

  var list = (m[1] === 'driver' || m[1] === 'drv') ? DRIVER_QUESTIONS : VEHICLE_QUESTIONS;
  var n = parseInt(m[2], 10) - 1;
  var rest = m[3];

  for (var i = 0; i < list.length; i++) {
    if (rest.indexOf(list[i][0]) >= 0) {
      var q = list[i][1][n];
      if (!q) { return undefined; }
      return byFlat[flat(q)];
    }
  }

  return undefined;
}


/**
 * One row's worth of values, in the order of your headings.
 *
 * Every heading is looked up in what the website sent: the list above,
 * then the same words, then the numbered vehicle/driver questions, then -
 * for a heading that was cut short - the one question whose name starts
 * with it. A heading nothing matches is left empty. Nothing is renamed.
 */
function rowValues(headings, body) {
  var fields = body.fields || {};
  var byFlat = {};

  for (var name in fields) {
    if (fields.hasOwnProperty(name)) { byFlat[flat(name)] = fields[name]; }
  }

  if (!body.email && !Object.keys(byFlat).length) { return null; }

  var aliases = {};
  for (var h in NEW_ROW_FROM) {
    if (NEW_ROW_FROM.hasOwnProperty(h)) { aliases[flat(h)] = flat(NEW_ROW_FROM[h]); }
  }

  var line = [];

  for (var c = 0; c < headings.length; c++) {
    var want = flat(headings[c]);
    var value = '';
    var num;

    if (!want) {
      value = '';
    } else if (want === 'savedat') {
      value = body.saved_at || new Date();
    } else if (aliasFor(aliases, want) && byFlat[aliasFor(aliases, want)] !== undefined) {
      value = byFlat[aliasFor(aliases, want)];
    } else if (byFlat[want] !== undefined) {
      value = byFlat[want];
    } else if ((num = numbered(want, byFlat)) !== undefined) {
      value = num;
    } else if (want.length >= 5) {
      // A heading cut short: the one question that starts with it.
      var hits = Object.keys(byFlat).filter(function (k) { return k.indexOf(want) === 0; });
      if (hits.length === 1) { value = byFlat[hits[0]]; }
    }

    if (value === '' && want === 'email') { value = body.email || ''; }
    if (value === '' && want === 'firstname') { value = body.first_name || ''; }
    if (value === '' && want === 'lastname') { value = body.last_name || ''; }

    line.push(value);
  }

  return line;
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
