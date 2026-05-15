// ================================================================
//  CA Ashish M Chandak — Contact Form → Google Sheets
//  Google Apps Script
//
//  PASTE THIS ENTIRE FILE into your Apps Script project,
//  replacing ALL existing code, then follow the steps below.
// ================================================================
//
//  ── ONE-TIME SETUP ──────────────────────────────────────────
//
//  Step 1: Paste this code (replace everything in the editor)
//
//  Step 2: Create sheet headers (run ONCE manually):
//    · Click the function dropdown → select "setupSheet"
//    · Click ▶ Run
//    · Click "Review Permissions" → Advanced → Allow
//    · Check your Google Sheet — a styled header row appears
//
//  Step 3: Deploy as Web App:
//    · Click Deploy → New Deployment
//    · Click gear icon → select "Web App"
//    · Execute as   → Me (your Gmail)
//    · Who can access → Anyone  (NOT "Anyone with Google Account")
//    · Click Deploy → Authorise → Allow
//    · COPY the /exec URL
//
//  Step 4: Paste /exec URL into index.html:
//    const GOOGLE_SCRIPT_URL = 'YOUR_EXEC_URL_HERE';
//
//  Step 5: Test in browser:
//    Open your /exec URL in a browser tab.
//    You must see: {"status":"active","message":"Endpoint is live"}
//    If you see a Google login page: access is wrong, redo Step 3.
//
//  Step 6: Test real submission:
//    Select "testSubmission" → click Run → check Sheet for new row.
//
//  EVERY TIME YOU EDIT THIS CODE:
//    Deploy → Manage Deployments → Edit → Version: New version → Deploy
//
// ================================================================

var SHEET_TAB_NAME = ''; // Leave blank = active sheet. Or set e.g. 'Form Responses'

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status  : 'active',
      message : 'Endpoint is live. Send POST requests to submit form data.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    lock.tryLock(10000);

    var p = e.parameter || {};

    // Fallback: also try JSON body
    if (!p.fullName && !p.email && e.postData && e.postData.contents) {
      try { p = JSON.parse(e.postData.contents); } catch(_) {}
    }

    if (!p.fullName && !p.email) {
      lock.releaseLock();
      return ContentService
        .createTextOutput(JSON.stringify({ result: 'error', message: 'No data received' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SHEET_TAB_NAME ? ss.getSheetByName(SHEET_TAB_NAME) : ss.getActiveSheet();

    if (!sheet) {
      lock.releaseLock();
      return ContentService
        .createTextOutput(JSON.stringify({ result: 'error', message: 'Sheet not found: ' + SHEET_TAB_NAME }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    sheet.appendRow([
      new Date(),
      p.fullName  || '',
      p.email     || '',
      p.phone     || '',
      p.company   || '',
      p.service   || '',
      p.office    || '',
      p.message   || '',
      p.consent   || 'No',
      p.timestamp || ''
    ]);

    lock.releaseLock();
    Logger.log('Row saved: ' + p.fullName + ' / ' + p.email);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', message: 'Saved' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    try { lock.releaseLock(); } catch(_) {}
    Logger.log('ERROR: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function setupSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SHEET_TAB_NAME ? ss.getSheetByName(SHEET_TAB_NAME) : ss.getActiveSheet();

  if (sheet.getRange(1, 1).getValue() !== '') {
    Logger.log('Headers already exist. Delete row 1 first to re-run.');
    return;
  }

  var headers = [
    'Timestamp', 'Full Name', 'Email Address', 'Phone Number',
    'Company / Organisation', 'Service Required', 'Preferred Office',
    'Message', 'Consent', 'Client Timestamp'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#0a1628')
    .setFontColor('#c9a84c')
    .setHorizontalAlignment('center');

  [180, 200, 230, 150, 210, 270, 150, 420, 100, 200].forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  sheet.setFrozenRows(1);
  Logger.log('Headers created successfully.');
}

function testSubmission() {
  var fakeEvent = {
    parameter: {
      fullName : 'Test User',
      email    : 'test@example.com',
      phone    : '+91 9876543210',
      company  : 'Test Company',
      service  : 'Forensic Audit',
      office   : 'Mumbai',
      message  : 'This is a test submission. Please delete this row.',
      consent  : 'Yes',
      timestamp: new Date().toISOString()
    },
    postData: null
  };

  var result = doPost(fakeEvent);
  Logger.log('Result: ' + result.getContent());
  Logger.log('Check your Sheet for the new test row.');
}
