// ═══════════════════════════════════════════════════════════════
// EDUCART BDM ASSESSMENT — Google Apps Script Backend
// ═══════════════════════════════════════════════════════════════
//
// SETUP INSTRUCTIONS (takes ~5 minutes):
//
// 1. Go to https://script.google.com
// 2. Click "New Project"
// 3. Delete all existing code and paste this entire file
// 4. Click "Save" (Ctrl+S)
// 5. Click "Deploy" → "New Deployment"
// 6. Type: "Select type" → choose "Web app"
// 7. Set:
//      Description: Educart BDM Assessment
//      Execute as: Me
//      Who has access: Anyone
// 8. Click "Deploy"
// 9. Click "Authorize access" and follow the steps
// 10. COPY the Web App URL shown — it looks like:
//     https://script.google.com/macros/s/AKfy.../exec
// 11. Paste that URL into index.html where it says:
//     const GOOGLE_SCRIPT_URL = 'YOUR_GOOGLE_APPS_SCRIPT_URL_HERE';
//
// That's it! Every test submission will create a new row in your Google Sheet.
// ═══════════════════════════════════════════════════════════════

// The script auto-creates a Google Sheet named this:
const SHEET_NAME = 'Educart BDM Responses';

// One sheet tab per test:
const TAB_NAMES = {
  1: 'Test 1 — QB Readiness',
  2: 'Test 2 — Scenarios',
  3: 'Test 3 — Academic Structure',
  4: 'Test 4 — NEP & Policy',
};

// Questions per test (for column headers):
const QUESTIONS = {
  1: [
    'Why is 2026-27 CBSE pattern different?',
    'What does 50% competency-based mean?',
    'Textbook vs Reference Book vs QB?',
    "Why is it called a Question 'Bank'?",
    'Why do traditional books fail in 2026-27?',
    'Response to NCERT is enough objection',
    'Difference: NCERT vs RD Sharma vs Educart QB',
    '1-minute pitch on Educart QB',
    'Parameters to evaluate a good book',
    'Problem QB solves for board prep',
  ],
  2: [
    '5 questions for RD Sharma teacher',
    'How to reposition QB vs RD Sharma',
    'Unique features / USPs of QB',
    'How to review topper answers',
    'Response to too many books objection',
    'How to join specimen program',
    'Response to price objection',
    '5-minute HoD pitch',
  ],
  3: [
    'Teachers & schools as real customers',
    'NCERT vs CBSE difference',
    'New 2026-27 exam pattern',
    'Academic calendar of CBSE school',
    'Best month to maximise sales',
    'QB vs One-Shot difference',
    'Why CBSE is core market',
    'GTM strategies of Educart',
    'Summer vacation rapport plan',
    'Teacher recommendation + samples vs content',
    'Teacher engagement initiatives',
    'Last 4 months strategy',
  ],
  4: [
    'Full form and year of NEP and NCF',
    'Primary aim of NEP 2020',
    '5+3+3+4 model explanation',
    'NIPUN Bharat Mission',
    'CPD Training and targets',
    'Using NEP/NCF in school conversations',
    'Policy → Classroom → Book → Sale chain',
    'How NEP changes teacher role',
    'How Educart books align with NEP',
  ],
};

// ───────────────────────────────────────────────
// Main entry point — handles POST requests
// ───────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const testId = parseInt(data.testId);
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateTab(ss, testId);
    appendRow(sheet, data, testId);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Allow GET for testing (visit URL in browser to confirm it works)
function doGet(e) {
  return ContentService
    .createTextOutput('Educart BDM Assessment API is running ✓')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ───────────────────────────────────────────────
// Get or create the Google Spreadsheet
// ───────────────────────────────────────────────
function getOrCreateSpreadsheet() {
  const files = DriveApp.getFilesByName(SHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  const ss = SpreadsheetApp.create(SHEET_NAME);
  return ss;
}

// ───────────────────────────────────────────────
// Get or create a sheet tab for the test
// ───────────────────────────────────────────────
function getOrCreateTab(ss, testId) {
  const tabName = TAB_NAMES[testId] || `Test ${testId}`;
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    writeHeaders(sheet, testId);
  }
  return sheet;
}

// ───────────────────────────────────────────────
// Write header row
// ───────────────────────────────────────────────
function writeHeaders(sheet, testId) {
  const questions = QUESTIONS[testId] || [];
  const headers = [
    'Timestamp',
    'Name',
    'Email',
    'Phone',
    'Test',
    'Time Taken (sec)',
    'Auto Submit?',
    'Tab Switches',
  ];
  questions.forEach((q, i) => headers.push(`Q${i + 1}: ${q}`));
  sheet.appendRow(headers);

  // Style the header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1A3A2A');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  sheet.setFrozenRows(1);

  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, i <= 8 ? 140 : 300);
  }
}

// ───────────────────────────────────────────────
// Append a response row
// ───────────────────────────────────────────────
function appendRow(sheet, data, testId) {
  const questions = QUESTIONS[testId] || [];
  const row = [
    data.submittedAt || new Date().toLocaleString(),
    data.name || '',
    data.email || '',
    data.phone || '',
    data.testTitle || `Test ${testId}`,
    data.timeTaken || 0,
    data.autoSubmit ? 'Yes' : 'No',
    data.tabSwitches || 0,
  ];
  questions.forEach((_, i) => {
    row.push(data[`Q${i + 1}`] || '');
  });
  sheet.appendRow(row);

  // Highlight flagged rows (tab switches > 0)
  if (parseInt(data.tabSwitches) > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, row.length).setBackground('#FFF3E0');
  }
}
