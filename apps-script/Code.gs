/**
 * Sojourn website form handler.
 *
 * Receives POST requests from the "Keep Me Posted" (home page) and
 * "Join Us" (recruiting page) forms, appends each submission to its own
 * tab in this bound spreadsheet, and emails a notification.
 *
 * SETUP (one time):
 * 1. Open the "Sojourn Website Submissions" Google Sheet.
 * 2. Extensions -> Apps Script.
 * 3. Delete any starter code and paste this whole file in.
 * 4. Update NOTIFY_EMAILS below with who should get notified.
 * 5. Deploy -> New deployment -> select type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Click Deploy, authorize the requested permissions, and copy the
 *    "Web app URL" it gives you (ends in /exec).
 * 7. Paste that URL into SUBMIT_ENDPOINT in index.html and recruiting.html.
 */

// Who should get an email every time someone submits a form.
var NOTIFY_EMAILS = ['glenn.kelman@gmail.com'];

var NEWSLETTER_SHEET = 'Newsletter signups';
var APPLICATION_SHEET = 'Applications';

var NEWSLETTER_HEADERS = ['Timestamp', 'Email'];
var APPLICATION_HEADERS = [
  'Timestamp',
  'Product built / how customers liked it',
  'LinkedIn or portfolio link',
  'Why interested in Sojourn'
];

function doPost(e) {
  try {
    var params = e.parameter;
    var formType = params.formType;

    if (formType === 'newsletter') {
      handleNewsletter(params);
    } else if (formType === 'application') {
      handleApplication(params);
    } else {
      return jsonResponse({ ok: false, error: 'Unknown form type' });
    }

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function handleNewsletter(params) {
  var sheet = getOrCreateSheet(NEWSLETTER_SHEET, NEWSLETTER_HEADERS);
  var timestamp = new Date();
  sheet.appendRow([timestamp, params.email || '']);

  notify(
    'New Sojourn newsletter signup',
    'Email: ' + (params.email || '(none)') + '\nTime: ' + timestamp
  );
}

function handleApplication(params) {
  var sheet = getOrCreateSheet(APPLICATION_SHEET, APPLICATION_HEADERS);
  var timestamp = new Date();
  sheet.appendRow([
    timestamp,
    params.product || '',
    params.linkedin || '',
    params.why || ''
  ]);

  notify(
    'New Sojourn application',
    'Product: ' + (params.product || '(none)') +
    '\n\nLinkedIn/portfolio: ' + (params.linkedin || '(none)') +
    '\n\nWhy Sojourn: ' + (params.why || '(none)') +
    '\n\nTime: ' + timestamp
  );
}

function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function notify(subject, body) {
  if (!NOTIFY_EMAILS || NOTIFY_EMAILS.length === 0) return;
  MailApp.sendEmail(NOTIFY_EMAILS.join(','), subject, body);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
