/**
 * Nexspace Realty - enquiry form + owner property registration receiver
 * Paste this into Extensions > Apps Script of the "Nexspace Enquiries" Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 *
 * Every submission:
 *  1. is saved as a new row in the "Enquiries" tab of this sheet (nothing is ever lost),
 *  2. has its PDF saved in the Drive folder "Nexspace Enquiry PDFs",
 *  3. is emailed with the PDF attached to the addresses below.
 */
const TO = 'nathita.t@nexspace-realty.com';
const CC = 'piriya.t@nexspace-realty.com';
const SHEET_NAME = 'Enquiries';
const FOLDER_NAME = 'Nexspace Enquiry PDFs';

const OWNER_SHEET_NAME = 'Owner Listings';
const OWNER_FOLDER_NAME = 'Nexspace Owner Listing PDFs';
const OWNER_COLUMNS = [
  'Submitted', 'Name', 'Phone', 'Email', 'LINE / WhatsApp', 'Best way to reach',
  'Property type', 'Property name', 'Unit / Address no.', 'Area (sq m)', 'Floor',
  'Bedroom', 'Bathroom', 'Maid', 'Parking', 'Listing', 'Resale price (THB)', 'Rental price (THB / month)',
  'Net price (THB)', 'Furnished', 'Appliances', 'Direction', 'Key', 'Pet', 'TF expenses',
  'Additional conditions', 'Agent / referred by', 'Source', 'Form language', 'Signed', 'PDF', 'Email sent'
];

const ENQUIRY_COLUMNS = [
  'Submitted', 'Name', 'Phone', 'Email', 'LINE / WhatsApp', 'Best way to reach',
  'Property type', 'Buy', 'Rent', 'Additional conditions',
  'Agent / referred by', 'Source', 'Form language', 'Signed', 'PDF', 'Email sent'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const f = data.fields || {};
    // form: "owner" comes from /enquiry/owner/ (property registration); anything else is a client enquiry.
    const isOwner = data.form === 'owner';
    const COLUMNS = isOwner ? OWNER_COLUMNS : ENQUIRY_COLUMNS;

    const sheet = isOwner ? getSheet_(OWNER_SHEET_NAME, OWNER_COLUMNS) : getSheet_(SHEET_NAME, ENQUIRY_COLUMNS);
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, COLUMNS.length).setValues([COLUMNS.map(c => c === 'PDF' || c === 'Email sent' ? '' : (f[c] || ''))]);

    let pdfBlob = null, pdfUrl = '';
    if (data.pdfBase64) {
      pdfBlob = Utilities.newBlob(Utilities.base64Decode(data.pdfBase64), 'application/pdf', data.filename || 'Nexspace Enquiry.pdf');
      const file = getFolder_(isOwner ? OWNER_FOLDER_NAME : FOLDER_NAME).createFile(pdfBlob);
      pdfUrl = file.getUrl();
      sheet.getRange(row, COLUMNS.indexOf('PDF') + 1).setValue(pdfUrl);
    }

    const rows = COLUMNS.filter(c => c !== 'PDF' && c !== 'Email sent')
      .map(c => '<tr><td style="padding:6px 12px;color:#5d6584;white-space:nowrap;vertical-align:top">' + esc_(c) +
                '</td><td style="padding:6px 12px;color:#1a2140">' + esc_(f[c] || '—').replace(/\n/g, '<br>') + '</td></tr>')
      .join('');
    const html = '<div style="font-family:Arial,sans-serif">' +
      '<h2 style="color:#1f2f8f;margin:0 0 12px">' + (isOwner ? 'New owner listing' : 'New enquiry') + '</h2>' +
      '<table style="border-collapse:collapse;font-size:14px">' + rows + '</table>' +
      (pdfUrl ? '<p style="font-size:13px;color:#5d6584">PDF also saved in Drive: <a href="' + pdfUrl + '">open</a></p>' : '') +
      '</div>';

    const opts = { cc: CC, htmlBody: html, name: isOwner ? 'Nexspace Property Registration' : 'Nexspace Enquiry Form' };
    if (pdfBlob) opts.attachments = [pdfBlob];
    if (f['Email'] && /@/.test(f['Email'])) opts.replyTo = f['Email'];
    MailApp.sendEmail(TO, data.subject || ('New enquiry: ' + (f['Name'] || '')), 'A new enquiry was submitted. See the HTML version or the attached PDF.', opts);
    sheet.getRange(row, COLUMNS.indexOf('Email sent') + 1).setValue('Yes');

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Lets you open the web app URL in a browser to check it is running.
function doGet() {
  return json_({ ok: true, service: 'Nexspace enquiry receiver' });
}

// Run once from the editor (select "setup", press Run) to create the sheet tab and Drive folder and grant permissions.
function setup() {
  getSheet_(SHEET_NAME, ENQUIRY_COLUMNS);
  getFolder_(FOLDER_NAME);
  getSheet_(OWNER_SHEET_NAME, OWNER_COLUMNS);
  getFolder_(OWNER_FOLDER_NAME);
}

function getSheet_(name, columns) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, columns.length).setValues([columns]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
