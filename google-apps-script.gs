// ─────────────────────────────────────────────────────────────────────────────
// Assorted Group — Job Application → Google Sheets
//
// SETUP INSTRUCTIONS:
//  1. Open Google Sheets and create a new sheet (name it "Applications").
//  2. Go to Extensions → Apps Script.
//  3. Paste this entire file into the editor (replace the default function).
//  4. Click "Deploy" → "New deployment" → Type: Web app.
//     - Execute as: Me
//     - Who has access: Anyone
//  5. Click "Deploy" and copy the Web App URL.
//  6. In index.html, replace 'YOUR_GOOGLE_APPS_SCRIPT_URL_HERE' with that URL.
// ─────────────────────────────────────────────────────────────────────────────

var SHEET_NAME = 'Applications';

var HEADERS = [
  'Timestamp',
  'Full Name',
  'Email',
  'Phone',
  'Location',
  'Department / Vertical',
  'Years of Experience',
  'Position Applied For',
  'LinkedIn Profile',
  'Portfolio / CV Link',
  'How They Found Us',
  'Cover Letter'
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet();

    sheet.appendRow([
      data.timestamp    || new Date().toISOString(),
      data.fullName     || '',
      data.email        || '',
      data.phone        || '',
      data.location     || '',
      data.department   || '',
      data.experience   || '',
      data.position     || '',
      data.linkedin     || '',
      data.portfolio    || '',
      data.source       || '',
      data.coverLetter  || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Assorted Group Job Application endpoint is live.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var headerRow = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRow.setValues([HEADERS]);
    headerRow.setFontWeight('bold');
    headerRow.setBackground('#1C4585');
    headerRow.setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, HEADERS.length, 200);
  }

  return sheet;
}
