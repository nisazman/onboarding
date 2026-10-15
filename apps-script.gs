// ============================================================
// AA Group FSC Onboarding — Google Apps Script Web App
// Sheet ID: 1d3rUMgf-js7vGikOc9rRXKjW7Jokw59zyqH7gAd8GOQ
// Deploy: Web App > Execute as Me > Anyone can access
// ============================================================

const SHEET_ID    = '1d3rUMgf-js7vGikOc9rRXKjW7Jokw59zyqH7gAd8GOQ';
const NOTIFY_EMAIL = 'nisa.azman@aagroup.sg';

function doGet() {
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var type = body.type || '';

    if (type === 'details')     return handleDetails(body);
    if (type === 'certs')       return handleCerts(body);
    if (type === 'scibooking')  return handleSciBooking(body);
    if (type === 'namecard')    return handleFormUpload(body, 'Namecard Submissions', '🪪 New Namecard Submission');
    if (type === 'officepass')  return handleFormUpload(body, 'Office Pass Submissions', '🏢 New Office Pass Submission');

    return respond(false, 'Unknown type: ' + type);
  } catch(err) {
    return respond(false, 'Error: ' + err.message);
  }
}

// ── 1. Candidate Details ─────────────────────────────────────
function handleDetails(data) {
  var sheet = getOrCreateSheet('Candidate Details',
    ['Timestamp','Full Name','Preferred Name','NRIC (last 4)','Mobile','Personal Email']);
  sheet.appendRow([
    data.timestamp || now(),
    data.fullName || '', data.preferredName || '',
    data.nric || '', data.mobile || '', data.email || ''
  ]);
  sendEmail('🆕 New FSC Candidate — ' + (data.fullName || ''),
    'Name: '           + data.fullName +
    '\nPreferred: '    + data.preferredName +
    '\nMobile: '       + data.mobile +
    '\nEmail: '        + data.email +
    '\nSubmitted: '    + data.timestamp +
    '\n\n' + sheetLink());
  return respond(true, 'Details saved.');
}

// ── 2. CMFAS Cert Upload Notification ───────────────────────
function handleCerts(data) {
  var sheet = getOrCreateSheet('Cert Uploads',
    ['Timestamp','Full Name','Notes']);
  sheet.appendRow([data.timestamp || now(), data.certName || '', data.certNotes || '']);
  sendEmail('📎 CMFAS Certs Ready — ' + (data.certName || ''),
    'Name: '       + data.certName +
    '\nNotes: '    + (data.certNotes || '—') +
    '\nSubmitted: '+ data.timestamp +
    '\n\nCandidate will WhatsApp/email certs directly.' +
    '\n\n' + sheetLink());
  return respond(true, 'Cert notification saved.');
}

// ── 3. SCI Exam Booking Request ──────────────────────────────
function handleSciBooking(data) {
  var sheet = getOrCreateSheet('SCI Exam Bookings',
    ['Timestamp','Full Name','SCI Student ID','Exam Paper(s)','Preferred Date','Preferred Time','Alternative Date']);
  sheet.appendRow([
    data.timestamp || now(),
    data.fullName || '', data.sciStudentId || '',
    data.examPapers || '', data.examDate || '',
    data.examTime || '', data.altDate || '—'
  ]);
  sendEmail('🎓 Exam Booking Request — ' + (data.fullName || ''),
    'Name: '            + data.fullName +
    '\nSCI Student ID: '+ data.sciStudentId +
    '\nExam Paper(s): ' + data.examPapers +
    '\nPreferred Date: '+ data.examDate +
    '\nPreferred Time: '+ data.examTime +
    '\nAlternative: '   + (data.altDate || '—') +
    '\nSubmitted: '     + data.timestamp +
    '\n\nPlease log in to SCI and book the slot.' +
    '\n\n' + sheetLink());
  return respond(true, 'Booking request saved.');
}

// ── 4. Namecard / Office Pass Submission ─────────────────────
function handleFormUpload(data, sheetName, subject) {
  var sheet = getOrCreateSheet(sheetName, ['Timestamp','Full Name']);
  sheet.appendRow([data.timestamp || now(), data.fullName || '']);
  sendEmail(subject + ' — ' + (data.fullName || ''),
    'Name: '      + data.fullName +
    '\nSubmitted: '+ data.timestamp +
    '\n\nCandidate will WhatsApp/email the signed form directly.' +
    '\n\n' + sheetLink());
  return respond(true, 'Submission saved.');
}

// ── Helpers ──────────────────────────────────────────────────
function getOrCreateSheet(name, headers) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function sendEmail(subject, body) {
  MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: subject, body: body });
}

function sheetLink() {
  return 'View sheet:\nhttps://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit';
}

function now() {
  return new Date().toLocaleString('en-SG', { timeZone: 'Asia/Singapore' });
}

function respond(ok, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
