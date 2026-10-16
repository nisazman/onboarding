// ============================================================
// AA Group FSC Onboarding — Google Apps Script Web App
// Sheet ID: 1d3rUMgf-js7vGikOc9rRXKjW7Jokw59zyqH7gAd8GOQ
// Deploy: Web App > Execute as Me > Anyone can access
// ============================================================

const SHEET_ID     = '1d3rUMgf-js7vGikOc9rRXKjW7Jokw59zyqH7gAd8GOQ';
const NOTIFY_EMAIL = 'nisa.azman@aagroup.sg';

// Leader email lookup — add new leaders here
const LEADER_EMAILS = {
  'Ray Tan':    'ray.advisorsalliance@gmail.com',
  'Lucas Chan': '',  // add when ready
};

// ── CORS preflight (OPTIONS) ─────────────────────────────────
function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── GET ──────────────────────────────────────────────────────
// Handles both ?payload= form submissions AND JSONP dashboard
// calls (?payload=…&callback=fn or payload containing callback).
function doGet(e) {
  if (e && e.parameter && e.parameter.payload) {
    try {
      var body = JSON.parse(decodeURIComponent(e.parameter.payload));

      // JSONP callback — comes inside the payload JSON
      var cb = body.callback || null;

      if (body.type === 'dashboard') {
        var result = getDashboard(body.leader || '');
        if (cb) {
          return ContentService
            .createTextOutput(cb + '(' + JSON.stringify(result) + ');')
            .setMimeType(ContentService.MimeType.JAVASCRIPT);
        }
        return respondJson(result);
      }

      return route(body);

    } catch(err) {
      return respond(false, 'Parse error: ' + err.message);
    }
  }
  return respond(true, 'OK');
}

// ── POST ─────────────────────────────────────────────────────
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    return route(body);
  } catch(err) {
    return respond(false, 'Error: ' + err.message);
  }
}

function route(body) {
  var type = body.type || '';
  if (type === 'details')    return handleDetails(body);
  if (type === 'certs')      return handleCerts(body);
  if (type === 'scibooking') return handleSciBooking(body);
  if (type === 'namecard')   return handleFormUpload(body, 'Namecard Submissions', '🪪 New Namecard Submission');
  if (type === 'officepass') return handleFormUpload(body, 'Office Pass Submissions', '🏢 New Office Pass Submission');
  return respond(false, 'Unknown type: ' + type);
}

// ── 1. Candidate Details ─────────────────────────────────────
function handleDetails(data) {
  var sheet = getOrCreateSheet('Candidate Details',
    ['Timestamp','Leader','Full Name','Preferred Name','NRIC / FIN','Mobile','Personal Email','Mailing Address']);
  sheet.appendRow([
    data.timestamp || now(),
    data.leader || '',
    data.fullName || '', data.preferredName || '',
    data.nric || '', data.mobile || '', data.email || '',
    data.mailingAddress || ''
  ]);
  var leaderEmail = LEADER_EMAILS[data.leader || ''] || '';
  var subject = '🆕 New FSC Candidate — ' + (data.fullName || '');
  var body =
    'Leader: '            + (data.leader || '—') +
    '\nName: '            + data.fullName +
    '\nPreferred: '       + data.preferredName +
    '\nNRIC / FIN: '      + (data.nric || '—') +
    '\nMobile: '          + data.mobile +
    '\nEmail: '           + data.email +
    '\nMailing Address: ' + (data.mailingAddress || '—') +
    '\nSubmitted: '       + data.timestamp +
    '\n\n' + sheetLink();

  Logger.log('Sending to admin: ' + NOTIFY_EMAIL);
  sendEmail(subject, body);

  if (leaderEmail) {
    Logger.log('Sending to leader: ' + leaderEmail);
    MailApp.sendEmail({ to: leaderEmail, subject: subject, body: body });
    Logger.log('Leader email sent successfully.');
  } else {
    Logger.log('No leader email found for: ' + (data.leader || '—'));
  }

  return respond(true, 'Details saved.');
}

// ── 2. CMFAS Cert Notification ───────────────────────────────
function handleCerts(data) {
  var sheet = getOrCreateSheet('Cert Uploads',
    ['Timestamp','Full Name','Notes']);
  sheet.appendRow([data.timestamp || now(), data.certName || '', data.certNotes || '']);
  sendEmail('📎 CMFAS Certs Ready — ' + (data.certName || ''),
    'Name: '       + data.certName +
    '\nNotes: '    + (data.certNotes || '—') +
    '\nSubmitted: '+ data.timestamp +
    '\n\nCandidate will WhatsApp/email certs directly.\n\n' + sheetLink());
  return respond(true, 'Cert notification saved.');
}

// ── 3. SCI Exam Booking ──────────────────────────────────────
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
    'Name: '             + data.fullName +
    '\nSCI Student ID: ' + data.sciStudentId +
    '\nExam Paper(s): '  + data.examPapers +
    '\nPreferred Date: ' + data.examDate +
    '\nPreferred Time: ' + data.examTime +
    '\nAlternative: '    + (data.altDate || '—') +
    '\nSubmitted: '      + data.timestamp +
    '\n\nPlease log in to SCI and book the slot.\n\n' + sheetLink());
  return respond(true, 'Booking request saved.');
}

// ── 4. Namecard / Office Pass ────────────────────────────────
function handleFormUpload(data, sheetName, subject) {
  var sheet = getOrCreateSheet(sheetName, ['Timestamp','Full Name']);
  sheet.appendRow([data.timestamp || now(), data.fullName || '']);
  sendEmail(subject + ' — ' + (data.fullName || ''),
    'Name: '       + data.fullName +
    '\nSubmitted: ' + data.timestamp +
    '\n\nCandidate will WhatsApp/email the signed form directly.\n\n' + sheetLink());
  return respond(true, 'Submission saved.');
}

// ── 5. Leader Dashboard ──────────────────────────────────────
// Returns all rows from Candidate Details, Cert Uploads,
// Namecard Submissions and Office Pass Submissions
// filtered by the given leader name.
function getDashboard(leader) {
  if (!leader) return { ok: false, message: 'No leader specified.' };

  var ss = SpreadsheetApp.openById(SHEET_ID);

  // ── Read Candidate Details (leader in col B, name in col C)
  var detailsSheet = ss.getSheetByName('Candidate Details');
  var candidates = {};  // keyed by Full Name
  if (detailsSheet && detailsSheet.getLastRow() > 1) {
    var dData = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 7).getValues();
    dData.forEach(function(row) {
      // [Timestamp, Leader, Full Name, Preferred Name, NRIC, Mobile, Email]
      if ((row[1] || '').toString().toLowerCase() === leader.toLowerCase()) {
        var name = row[2] || '';
        if (name) {
          candidates[name] = {
            name: name,
            preferred: row[3] || '',
            mobile: row[5] || '',
            email: row[6] || '',
            certs: 'No',
            namecard: 'No',
            officepass: 'No'
          };
        }
      }
    });
  }

  // ── Cert Uploads (col B = Full Name)
  var certSheet = ss.getSheetByName('Cert Uploads');
  if (certSheet && certSheet.getLastRow() > 1) {
    var cData = certSheet.getRange(2, 1, certSheet.getLastRow() - 1, 2).getValues();
    cData.forEach(function(row) {
      var name = (row[1] || '').toString();
      if (candidates[name]) candidates[name].certs = 'Yes';
    });
  }

  // ── Namecard Submissions (col B = Full Name)
  var ncSheet = ss.getSheetByName('Namecard Submissions');
  if (ncSheet && ncSheet.getLastRow() > 1) {
    var ncData = ncSheet.getRange(2, 1, ncSheet.getLastRow() - 1, 2).getValues();
    ncData.forEach(function(row) {
      var name = (row[1] || '').toString();
      if (candidates[name]) candidates[name].namecard = 'Yes';
    });
  }

  // ── Office Pass Submissions (col B = Full Name)
  var opSheet = ss.getSheetByName('Office Pass Submissions');
  if (opSheet && opSheet.getLastRow() > 1) {
    var opData = opSheet.getRange(2, 1, opSheet.getLastRow() - 1, 2).getValues();
    opData.forEach(function(row) {
      var name = (row[1] || '').toString();
      if (candidates[name]) candidates[name].officepass = 'Yes';
    });
  }

  var rows = Object.values(candidates);
  return { ok: true, leader: leader, rows: rows };
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

function sendEmail(subject, body, cc) {
  var opts = { to: NOTIFY_EMAIL, subject: subject, body: body };
  if (cc) opts.cc = cc;
  MailApp.sendEmail(opts);
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

function respondJson(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
