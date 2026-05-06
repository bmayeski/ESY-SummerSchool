// ============================================================
//  Code.gs  —  ESY / Summer School Application Web App
//  Grossmont Union High School District  | 2026
//
//  SETUP STEPS:
//  1. Create a new Google Spreadsheet and paste its ID into
//     SPREADSHEET_ID below.
//  2. Add admin email addresses to ADMIN_EMAILS below.
//  3. Run setupSpreadsheet() once from the Apps Script editor
//     (Run > Run function > setupSpreadsheet).
//  4. Deploy as Web App:
//     Deploy > New deployment > Web App
//     Execute as: Me  | Who has access: Anyone in [your domain]
//
//  ADMIN PANEL ACCESS:
//  Navigate to your web app URL with ?admin=true appended.
//  Only emails listed in ADMIN_EMAILS will see the panel.
// ============================================================


// ── CONFIGURATION ───────────────────────────────────────────
const SPREADSHEET_ID = '1EjG_CPKGR9j9jQPAcz0AbNxMaPh6NYskonaJXS96noc'; // ← Paste your Google Spreadsheet ID here

// Add the Google account emails of anyone who should have
// access to the admin panel. Comparison is case-insensitive.
const ADMIN_EMAILS = [
  'bmayeski@guhsd.net',
  'sbrooks@guhsd.net'
  // 'another.admin@guhsd.net',
];

// ── SHEET NAMES ─────────────────────────────────────────────
const SHEET_NAMES = {
  CERT_ESY:   'Certificated ESY',
  CERT_SS:    'Certificated Summer School',
  CLASS_ESY:  'Classified ESY',
  ADMIN_INFO: 'Admin Info'
};

// ── COLUMN HEADERS ───────────────────────────────────────────
// Location uses Option 3: two columns — full ranked list + top choice.
// Session uses two separate checkbox columns instead of a ranking.

const HEADERS = {

  CERT_ESY: [
    'Timestamp',
    'Full Name',
    'Employee ID',
    'Phone',
    'Email',
    'Current School/Site',
    'Subject/Grades Taught in 25/26',
    'Credential Type',
    'Credential Authorization',
    'Authorized Subjects on Credential',
    'Preferred Subject to Teach During ESY',
    'Previous ESY Years',
    'Previous ESY Location',
    'Session Availability',           // e.g. "Session 1, Session 2"
    'Location Preference (Ranked)',   // e.g. "Grossmont, El Cajon, Monte Vista"
    'Top Location Choice',            // e.g. "Grossmont"
    'Interested in Substituting',
    'Signature Acknowledgment'
  ],

  CERT_SS: [
    'Timestamp',
    'Full Name',
    'Employee ID',
    'Phone',
    'Email',
    'Current School/Site',
    'Subject/Grades Taught in 25/26',
    'Credential Type',
    'Credential Authorization',
    'Authorized Subjects on Credential',
    'Previous Summer School Years',
    'Previous Summer School Location',
    'Session Availability',           // e.g. "Session 1, Session 2"
    'Location Preference (Ranked)',
    'Top Location Choice',
    'Interested in Substituting',
    'Signature Acknowledgment'
  ],

  CLASS_ESY: [
    'Timestamp',
    'Full Name',
    'Employee ID',
    'Job Title',
    'Current School/Site',
    'Phone',
    'Email',
    '25/26 Position(s)',
    'Session Availability',           // e.g. "Session 1"
    'Location Preference (Ranked)',
    'Top Location Choice',
    'Interested in Substituting',
    'Signature Acknowledgment'
  ]

};

// ── WEB APP ENTRY POINT ─────────────────────────────────────
function doGet(e) {
  const params      = e && e.parameter ? e.parameter : {};
  const isAdminReq  = params.admin === 'true';
  const userEmail   = Session.getActiveUser().getEmail().toLowerCase();
  const isAdmin     = ADMIN_EMAILS.map(a => a.toLowerCase()).includes(userEmail);

  const template = HtmlService.createTemplateFromFile('Index');
  template.isAdmin     = isAdmin;
  template.isAdminView = isAdminReq && isAdmin;

  return template
    .evaluate()
    .setTitle('ESY & Summer School Staff Applications — GUHSD 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ── HTML INCLUDE HELPER ──────────────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ── ADMIN AUTH CHECK ─────────────────────────────────────────
// Called from JS_Admin to verify access server-side
function checkAdminAccess() {
  const email   = Session.getActiveUser().getEmail().toLowerCase();
  const isAdmin = ADMIN_EMAILS.map(a => a.toLowerCase()).includes(email);
  return { isAdmin, email };
}


// ── ONE-TIME SPREADSHEET SETUP ───────────────────────────────
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  _ensureSheet(ss, SHEET_NAMES.CERT_ESY,  HEADERS.CERT_ESY);
  _ensureSheet(ss, SHEET_NAMES.CERT_SS,   HEADERS.CERT_SS);
  _ensureSheet(ss, SHEET_NAMES.CLASS_ESY, HEADERS.CLASS_ESY);
  _ensureAdminSheet(ss);

  SpreadsheetApp.flush();
  Logger.log('✅ Spreadsheet setup complete.');
}

function _ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a3a5c')
      .setFontColor('#ffffff')
      .setWrap(true);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 180);
  }
}

function _ensureAdminSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.ADMIN_INFO);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAMES.ADMIN_INFO);

  if (sheet.getLastRow() > 0) return; // Already set up

  const defaults = [
    ['Key', 'Value', 'Description'],

    // ── Home screen ──────────────────────────────────────
    ['announcementMsg',
     'Applications are now open for the 2026 ESY and Summer School Credit Recovery programs. Please review all program information carefully before submitting your interest application.',
     'Announcement banner on home screen'],

    ['applicationDeadline',
     'April 22, 2026',
     'Deadline shown in the header chip'],

    // ── Shared schedule (Comprehensive / Chap/IDEA / MERIT) ──
    ['sharedDates',
     '6/8/26 – 6/26/26 (no school 6/19)',
     'Dates for Comprehensive, Chap/IDEA, and MERIT sites'],

    ['sharedSessionI',
     '8:00 am – 11:00 am',
     'Session I hours for Comprehensive/Chap/IDEA/MERIT'],

    ['sharedSessionII',
     '11:45 am – 2:45 pm',
     'Session II hours for Comprehensive/Chap/IDEA/MERIT'],

    // ── ELITE ────────────────────────────────────────────
    ['eliteSessIDates',   '6/8/26 – 6/26/26',        'ELITE Session I dates'],
    ['eliteSessIIDates',  '6/29/26 – 7/17/26',       'ELITE Session II dates'],
    ['eliteHours',        '9:30 am – 1:30 pm',        'ELITE daily hours'],
    ['eliteNoSchool',     'No school 6/19 or 7/3',    'ELITE no-school note'],

    // ── Helix ────────────────────────────────────────────
    ['helixDates',        '6/10/26 – 7/1/26',         'Helix Charter dates'],
    ['helixSessIHours',   '8:00 am – 11:15 am',       'Helix Session I hours'],
    ['helixSessIIHours',  '11:45 am – 3:00 pm',       'Helix Session II hours'],
    ['helixNoSchool',     'No school 6/19',           'Helix no-school note'],

    // ── Steele Canyon ────────────────────────────────────
    ['steeleDates',       '6/8/26 – 6/26/26',         'Steele Canyon dates'],
    ['steeleSessIHours',  '8:00 am – 12:45 pm',       'Steele Canyon Session I hours'],
    ['steeleNoSchool',    'No school 6/19',            'Steele no-school note'],

    // ── Contact ──────────────────────────────────────────
    ['contactName',       'Special Education Department', 'Contact name/title'],
    ['contactEmail',      'sped@guhsd.net',               'Contact email'],
    ['contactPhone',      '(619) 644-8000',               'Contact phone'],

    ['additionalInfo',
     'Positions are contingent upon sufficient student enrollment. Placement decisions will be communicated after the application deadline. You will be contacted via email with next steps.',
     'Additional info paragraph on home screen'],
  ];

  sheet.getRange(1, 1, defaults.length, 3).setValues(defaults);
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#1a3a5c')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 480);
  sheet.setColumnWidth(3, 320);
  sheet.getRange(2, 2, defaults.length - 1, 1).setWrap(true);
}


// ── ADMIN INFO: READ ─────────────────────────────────────────
function getAdminInfo() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ADMIN_INFO);
    if (!sheet) return { success: false, error: 'Admin Info sheet not found. Run setupSpreadsheet() first.' };

    const rows = sheet.getDataRange().getValues();
    const info = {};
    for (let i = 1; i < rows.length; i++) {
      const key = String(rows[i][0]).trim();
      const val = String(rows[i][1]).trim();
      if (key) info[key] = val;
    }
    return { success: true, data: info };

  } catch (e) {
    Logger.log('getAdminInfo error: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ── ADMIN INFO: WRITE ────────────────────────────────────────
// Receives an object of { key: newValue } pairs and updates
// matching rows in the Admin Info sheet.
function saveAdminInfo(updates) {
  try {
    // Re-verify admin server-side before writing
    const email   = Session.getActiveUser().getEmail().toLowerCase();
    const isAdmin = ADMIN_EMAILS.map(a => a.toLowerCase()).includes(email);
    if (!isAdmin) return { success: false, error: 'Access denied.' };

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ADMIN_INFO);
    const rows  = sheet.getDataRange().getValues();

    let saved = 0;

    for (let i = 1; i < rows.length; i++) {
      const key = String(rows[i][0]).trim();
      if (key && updates.hasOwnProperty(key)) {
        sheet.getRange(i + 1, 2).setValue(updates[key]);
        saved++;
      }
    }

    SpreadsheetApp.flush();
    return { success: true, saved };

  } catch (e) {
    Logger.log('saveAdminInfo error: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ── FORM SUBMISSION HANDLERS ─────────────────────────────────
function submitCertESY(formData) {
  return _submitForm(formData, SHEET_NAMES.CERT_ESY, HEADERS.CERT_ESY, 'Certificated ESY');
}
function submitCertSS(formData) {
  return _submitForm(formData, SHEET_NAMES.CERT_SS, HEADERS.CERT_SS, 'Certificated Summer School Credit Recovery');
}
function submitClassESY(formData) {
  return _submitForm(formData, SHEET_NAMES.CLASS_ESY, HEADERS.CLASS_ESY, 'Classified ESY / Summer School');
}


// ── CORE SUBMIT LOGIC ────────────────────────────────────────
function _submitForm(formData, sheetName, headers, programLabel) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('Sheet "' + sheetName + '" not found. Run setupSpreadsheet() first.');

    const row = [new Date()];
    headers.slice(1).forEach(header => {
      const val = formData[header];
      row.push(val !== undefined && val !== null ? val : '');
    });

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    try {
      _sendConfirmationEmail(formData['Email'], formData['Full Name'], programLabel);
    } catch (emailErr) {
      Logger.log('Email error (non-fatal): ' + emailErr.message);
    }

    return { success: true };
  } catch (e) {
    Logger.log('_submitForm error [' + sheetName + ']: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ── CONFIRMATION EMAIL ───────────────────────────────────────
function _sendConfirmationEmail(toEmail, name, programLabel) {
  if (!toEmail || !toEmail.includes('@')) return;

  const displayName = name || 'Applicant';
  const subject     = `Application Received — ${programLabel} (GUHSD 2026)`;

  const htmlBody = `
<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f0f2f5;font-family:'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f2f5;padding:32px 16px;">
  <tr><td align="center">
    <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;border-radius:12px;overflow:hidden;box-shadow:0 4px 14px rgba(0,0,0,.10);">
      <tr>
        <td style="background:linear-gradient(135deg,#1a3a5c 0%,#2563a8 100%);padding:28px 32px;">
          <p style="margin:0 0 4px;color:rgba(255,255,255,.6);font-size:12px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;">Grossmont Union High School District</p>
          <h1 style="margin:0;color:#fff;font-size:22px;font-weight:700;">Application Received</h1>
          <p style="margin:6px 0 0;color:rgba(255,255,255,.75);font-size:14px;">${programLabel} — 2026 Staff Interest Application</p>
        </td>
      </tr>
      <tr>
        <td style="background:#fff;padding:28px 32px;">
          <p style="margin:0 0 16px;font-size:16px;color:#1c2333;">Dear <strong>${displayName}</strong>,</p>
          <p style="margin:0 0 16px;font-size:15px;color:#4a5568;line-height:1.6;">Thank you for submitting your interest application for the <strong>${programLabel}</strong> program. Your submission has been received.</p>
          <table width="100%" cellpadding="0" cellspacing="0" style="margin:20px 0;">
            <tr>
              <td style="background:#f0f4fa;border-left:4px solid #2563a8;border-radius:0 8px 8px 0;padding:14px 18px;">
                <p style="margin:0 0 8px;font-size:14px;font-weight:700;color:#1a3a5c;">What happens next?</p>
                <ul style="margin:0;padding-left:18px;font-size:14px;color:#4a5568;line-height:1.75;">
                  <li>Your application will be reviewed by the Special Education Department.</li>
                  <li>Selection is based on rotation and seniority.</li>
                  <li>You will be contacted at this email address with placement information.</li>
                  <li>Once you receive an offer, you have <strong>5 business days</strong> to accept or decline.</li>
                </ul>
              </td>
            </tr>
          </table>
          <p style="margin:0 0 16px;font-size:14px;color:#4a5568;line-height:1.6;">If you have questions, please contact the Special Education Department directly. Do not reply to this automated email.</p>
          <p style="margin:0;font-size:14px;color:#1c2333;">Warm regards,<br><strong>Special Education Department</strong><br>Grossmont Union High School District</p>
        </td>
      </tr>
      <tr>
        <td style="background:#e8edf2;padding:14px 32px;text-align:center;">
          <p style="margin:0;font-size:11px;color:#999;">This is an automated confirmation — please do not reply. | GUHSD 2026</p>
        </td>
      </tr>
    </table>
  </td></tr>
</table>
</body></html>`;

  const plainBody =
    `Dear ${displayName},\n\nYour application for ${programLabel} has been received.\n\n` +
    `What happens next:\n- Your application will be reviewed by the Special Education Department.\n` +
    `- Selection is based on rotation and seniority.\n- You will be contacted with placement information.\n` +
    `- Once you receive an offer, you have 5 business days to accept or decline.\n\n` +
    `Warm regards,\nSpecial Education Department\nGrossmont Union High School District`;

  GmailApp.sendEmail(toEmail, subject, plainBody, {
    htmlBody,
    name: 'GUHSD Summer Programs'
  });
}
