/**
 * Alice Public School - enquiry form backend (Google Apps Script)
 *
 * SETUP
 * 1. Create a Google Sheet (use a school Google account). Name the first tab "Enquiries".
 *    Add headers in row 1: Time | Parent name | Phone | Class | Message
 * 2. In the Sheet: Extensions > Apps Script. Paste this file in, replacing any existing code.
 * 3. Set NOTIFY_EMAIL below to the school office email.
 * 4. Deploy > New deployment > type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 * 5. Copy the Web app URL (ends in /exec) and paste it into ENQUIRY_URL in index.html.
 * 6. After any code change, use Deploy > Manage deployments > Edit > New version.
 */

const SHEET_NAME = 'Enquiries';
const NOTIFY_EMAIL = 'school@example.com'; // change to the real school email

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Honeypot: real visitors leave this hidden field empty, bots often fill it.
  if (p.website) return reply('ok');

  const name = clean(p.parentName, 80);
  const phone = clean(p.phone, 15);
  const cls = clean(p.classApplying, 20);
  const message = clean(p.message, 500);

  if (!name || !/^[0-9+\-\s]{10,15}$/.test(phone) || !cls) {
    return reply('invalid');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    sheet.appendRow([new Date(), name, phone, cls, message]);
  } finally {
    lock.releaseLock();
  }

  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New admission enquiry: ' + cls,
      body: 'Parent: ' + name + '\nPhone: ' + phone + '\nClass: ' + cls +
            '\nMessage: ' + (message || '-')
    });
  } catch (err) {
    // Enquiry is already saved; ignore email failures.
  }

  return reply('ok');
}

// Strips formula-injection characters and trims length so the sheet stays safe.
function clean(value, max) {
  let s = String(value || '').trim().slice(0, max);
  if (/^[=+\-@]/.test(s) && !/^\+?[0-9\s\-]+$/.test(s)) s = "'" + s;
  return s;
}

function reply(status) {
  return ContentService.createTextOutput(JSON.stringify({ status: status }))
    .setMimeType(ContentService.MimeType.JSON);
}