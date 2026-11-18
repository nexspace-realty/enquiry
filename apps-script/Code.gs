/**
 * Nexspace Realty - receiver for the enquiry form and the owner property registration form.  Version 2.
 * Lives in Extensions > Apps Script of the "Nexspace Enquiries" Google Sheet, deployed as a
 * Web app (Execute as: Me, Who has access: Anyone).
 *
 * To update: paste this whole file over Code.gs, Save, run "setup" once, then
 * Deploy > Manage deployments > (pencil) Edit > Version: New version > Deploy.  The web app URL stays the same.
 *
 * Client enquiries (enquiry form):
 *   row in the "Enquiries" tab, PDF in Drive folder "Nexspace Enquiry PDFs",
 *   email with the PDF to nathita.t@ (cc piriya.t@).
 * Owner property registrations (form: "owner"):
 *   row in the "Owner Listings" tab, PDF + the owner's photos in their own folder inside
 *   "Nexspace Owner Listings", email to info@ (cc nathita.t@, piriya.t@) with the PDF and photos attached
 *   (photos over ~20 MB in total stay in Drive and the email links to the folder).
 */
const TO = 'nathita.t@nexspace-realty.com';
const CC = 'piriya.t@nexspace-realty.com';
const SHEET_NAME = 'Enquiries';
const FOLDER_NAME = 'Nexspace Enquiry PDFs';

const OWNER_TO = 'info@nexspace-realty.com';
const OWNER_CC = 'nathita.t@nexspace-realty.com,piriya.t@nexspace-realty.com';
const OWNER_SHEET_NAME = 'Owner Listings';
const OWNER_FOLDER_NAME = 'Nexspace Owner Listings';
const MAX_ATTACH_BYTES = 20 * 1024 * 1024;
const OWNER_COLUMNS = [
  'Submitted', 'Name', 'Phone', 'Email', 'LINE / WhatsApp', 'Best way to reach',
  'Property type', 'Property name', 'Unit / Address no.', 'Area (sq m)', 'Floor',
  'Bedroom', 'Bathroom', 'Maid', 'Parking', 'Listing', 'Resale price (THB)', 'Rental price (THB / month)',
  'Net price (THB)', 'Furnished', 'Appliances', 'Direction', 'Key', 'Door code / password', 'Pet', 'Rent includes',
  'TF expenses', 'Additional conditions', 'Agent / referred by', 'Source', 'Form language', 'Signed',
  'Photos', 'Drive folder', 'PDF', 'Email sent'
];
const OWNER_AUTO = ['Photos', 'Drive folder', 'PDF', 'Email sent'];

const ENQUIRY_COLUMNS = [
  'Submitted', 'Name', 'Phone', 'Email', 'LINE / WhatsApp', 'Best way to reach',
  'Property type', 'Buy', 'Rent', 'Additional conditions',
  'Agent / referred by', 'Source', 'Form language', 'Signed', 'PDF', 'Email sent'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = JSON.parse(e.postData.contents);
    return data.form === 'owner' ? owner_(data) : enquiry_(data);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function enquiry_(data) {
  const f = data.fields || {};
  const COLUMNS = ENQUIRY_COLUMNS;
  const sheet = getSheet_(SHEET_NAME, COLUMNS);
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, COLUMNS.length).setValues([COLUMNS.map(c => c === 'PDF' || c === 'Email sent' ? '' : (f[c] || ''))]);

  let pdfBlob = null, pdfUrl = '';
  if (data.pdfBase64) {
    pdfBlob = Utilities.newBlob(Utilities.base64Decode(data.pdfBase64), 'application/pdf', data.filename || 'Nexspace Enquiry.pdf');
    pdfUrl = getFolder_(FOLDER_NAME).createFile(pdfBlob).getUrl();
    sheet.getRange(row, COLUMNS.indexOf('PDF') + 1).setValue(pdfUrl);
  }
  const html = table_('New enquiry', COLUMNS.filter(c => c !== 'PDF' && c !== 'Email sent'), f) +
    (pdfUrl ? '<p style="font-family:Arial,sans-serif;font-size:13px;color:#5d6584">PDF also saved in Drive: <a href="' + pdfUrl + '">open</a></p>' : '');
  const opts = { cc: CC, htmlBody: html, name: 'Nexspace Enquiry Form' };
  if (pdfBlob) opts.attachments = [pdfBlob];
  if (f['Email'] && /@/.test(f['Email'])) opts.replyTo = f['Email'];
  MailApp.sendEmail(TO, data.subject || ('New enquiry: ' + (f['Name'] || '')), 'A new enquiry was submitted. See the HTML version or the attached PDF.', opts);
  sheet.getRange(row, COLUMNS.indexOf('Email sent') + 1).setValue('Yes');
  return json_({ ok: true, v: 2 });
}

function owner_(data) {
  const f = data.fields || {};
  const COLUMNS = OWNER_COLUMNS;
  const sheet = getSheet_(OWNER_SHEET_NAME, COLUMNS);
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, COLUMNS.length).setValues([COLUMNS.map(c => OWNER_AUTO.indexOf(c) >= 0 ? '' : (f[c] || ''))]);
  // Door codes are kept as plain text so a code like 0123 keeps its leading zero.
  sheet.getRange(row, COLUMNS.indexOf('Door code / password') + 1).setNumberFormat('@').setValue(f['Door code / password'] || '');

  const stamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HHmm');
  const title = [f['Property name'], f['Unit / Address no.']].filter(String).join(' ') || 'Property';
  const folder = getFolder_(OWNER_FOLDER_NAME).createFolder(clean_(stamp + ' ' + title + ' - ' + (f['Name'] || 'Owner')));
  const folderUrl = folder.getUrl();
  sheet.getRange(row, COLUMNS.indexOf('Drive folder') + 1).setValue(folderUrl);

  const attachments = [];
  let bytes = 0;
  if (data.pdfBase64) {
    const pdf = Utilities.newBlob(Utilities.base64Decode(data.pdfBase64), 'application/pdf', data.filename || 'Nexspace Property Registration.pdf');
    sheet.getRange(row, COLUMNS.indexOf('PDF') + 1).setValue(folder.createFile(pdf).getUrl());
    attachments.push(pdf); bytes += pdf.getBytes().length;
  }
  const photos = (data.photos || []).slice(0, 30);
  let attachedPhotos = 0;
  photos.forEach((p, i) => {
    const blob = Utilities.newBlob(Utilities.base64Decode(p.base64), 'image/jpeg', clean_(p.name || ('photo ' + (i + 1) + '.jpg')));
    folder.createFile(blob);
    const n = blob.getBytes().length;
    if (bytes + n <= MAX_ATTACH_BYTES) { attachments.push(blob); bytes += n; attachedPhotos++; }
  });
  sheet.getRange(row, COLUMNS.indexOf('Photos') + 1).setValue(photos.length);

  let note = '<p style="font-family:Arial,sans-serif;font-size:13px;color:#5d6584">PDF' +
    (photos.length ? ' and ' + photos.length + ' owner photo' + (photos.length === 1 ? '' : 's') : '') +
    ' saved in Drive: <a href="' + folderUrl + '">open folder</a></p>';
  if (attachedPhotos < photos.length) {
    note += '<p style="font-family:Arial,sans-serif;font-size:13px;color:#b3261e">' + (photos.length - attachedPhotos) +
      ' photo(s) were too large to attach; they are in the Drive folder above.</p>';
  }
  const html = table_('New owner listing', COLUMNS.filter(c => OWNER_AUTO.indexOf(c) < 0 || c === 'Photos'), Object.assign({}, f, { 'Photos': String(photos.length) })) + note;
  const opts = { cc: OWNER_CC, htmlBody: html, name: 'Nexspace Property Registration', attachments: attachments };
  if (f['Email'] && /@/.test(f['Email'])) opts.replyTo = f['Email'];
  MailApp.sendEmail(OWNER_TO, data.subject || ('New owner listing: ' + title), 'A new property was registered. See the HTML version or the attached PDF.', opts);
  sheet.getRange(row, COLUMNS.indexOf('Email sent') + 1).setValue('Yes');
  return json_({ ok: true, v: 2, photos: photos.length });
}

function table_(heading, cols, f) {
  const rows = cols.map(c => '<tr><td style="padding:6px 12px;color:#5d6584;white-space:nowrap;vertical-align:top">' + esc_(c) +
    '</td><td style="padding:6px 12px;color:#1a2140">' + esc_(f[c] || '—').replace(/\n/g, '<br>') + '</td></tr>').join('');
  return '<div style="font-family:Arial,sans-serif"><h2 style="color:#1f2f8f;margin:0 0 12px">' + esc_(heading) + '</h2>' +
    '<table style="border-collapse:collapse;font-size:14px">' + rows + '</table></div>';
}

function clean_(s) {
  return String(s).replace(/[\\/:*?"<>|]+/g, '').trim().slice(0, 120);
}

// Opening the web app URL in a browser shows it is running; the owner form also uses this to check for version 2.
function doGet() {
  return json_({ ok: true, v: 2, service: 'Nexspace enquiry + property registration receiver' });
}

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
    sh.setFrozenRows(1);
  }
  // Keep the header row in step with the columns (new columns are added on the right).
  sh.getRange(1, 1, 1, columns.length).setValues([columns]).setFontWeight('bold');
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
