// =================================================================
// EDUCART BDM ASSESSMENT -- Google Apps Script (v3)
// Handles: registrations, logins, test submissions, admin reads
// Passwords are stored in plain text in the Candidates sheet.
// =================================================================
//
// SETUP INSTRUCTIONS:
// 1. Go to https://script.google.com -> New Project
// 2. Delete all existing code, paste this entire file
// 3. Save (Ctrl+S)
// 4. Deploy -> New Deployment -> Web App
//      Execute as: Me
//      Who has access: Anyone
// 5. Deploy -> Authorize -> Copy the Web App URL
// 6. Paste URL into index.html: const API = 'YOUR_URL_HERE';
//
// After editing this file always re-deploy with a New Version.
// =================================================================

var SHEET_NAME = 'Educart BDM Responses';

var TAB = {
  USERS: 'Candidates',
  T1: 'Test 1 - QB Readiness',
  T2: 'Test 2 - Scenarios',
  T3: 'Test 3 - Academic Structure',
  T4: 'Test 4 - NEP and Policy',
};

var Q_HEADERS = {
  1: [
    'Why is 2026-27 CBSE pattern different',
    'What does 50 percent competency-based mean',
    'Textbook vs Reference Book vs QB',
    'Why called a Question Bank',
    'Why traditional books fail in 2026-27',
    'Response to NCERT is enough objection',
    'NCERT vs RD Sharma vs Educart QB',
    '1-minute pitch on Educart QB',
    'Parameters to evaluate a good book',
    'Problem QB solves for board prep',
  ],
  2: [
    '5 questions for RD Sharma teacher',
    'Reposition QB vs RD Sharma',
    'Unique features and USPs of QB',
    'How to review topper answers',
    'Response to too many books objection',
    'How to join specimen program',
    'Response to price objection',
    '5-minute HoD pitch',
  ],
  3: [
    'Teachers and schools as real customers',
    'NCERT vs CBSE difference',
    'New 2026-27 exam pattern',
    'Academic calendar of CBSE school',
    'Best month to maximise sales',
    'QB vs One-Shot difference',
    'Why CBSE is core market',
    'GTM strategies of Educart',
    'Summer vacation rapport plan',
    'Teacher recommendation and samples vs content',
    'Teacher engagement initiatives',
    'Last 4 months strategy',
  ],
  4: [
    'Full form and year of NEP and NCF',
    'Primary aim of NEP 2020',
    '5+3+3+4 model explanation',
    'NIPUN Bharat Mission',
    'CPD Training and targets',
    'Using NEP and NCF in school conversations',
    'Policy to Classroom to Book to Sale chain',
    'How NEP changes teacher role',
    'How Educart books align with NEP',
  ],
};

// =================================================================
// ROUTER
// =================================================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'submit_test';
    var result;
    if      (action === 'register')    { result = handleRegister(data); }
    else if (action === 'login')       { result = handleLogin(data); }
    else if (action === 'submit_test') { result = handleSubmitTest(data); }
    else if (action === 'save_notes')      { result = handleSaveNotes(data); }
    else if (action === 'forgot_password') { result = handleForgotPassword(data); }
    else if (action === 'reset_password')  { result = handleResetPassword(data); }
    else                                   { result = { status: 'error', message: 'Unknown action' }; }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  try {
    var action = e.parameter.action || 'ping';
    var result;
    if      (action === 'ping')            { result = { status: 'ok', message: 'Educart BDM API running' }; }
    else if (action === 'get_submissions') { result = handleGetSubmissions(); }
    else if (action === 'get_candidates')  { result = handleGetCandidates(); }
    else if (action === 'get_detail')      { result = handleGetDetail(e.parameter); }
    else if (action === 'check_completed') { result = handleCheckCompleted(e.parameter); }
    else                                   { result = { status: 'error', message: 'Unknown action' }; }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =================================================================
// SPREADSHEET HELPERS
// =================================================================
function getOrCreateSS() {
  var files = DriveApp.getFilesByName(SHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  return SpreadsheetApp.create(SHEET_NAME);
}

function getOrCreateTab(ss, tabName, headers) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(headers);
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#1A3A2A');
    hRange.setFontColor('#FFFFFF');
    hRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    for (var i = 0; i < headers.length; i++) {
      sheet.setColumnWidth(i + 1, i < 8 ? 150 : 320);
    }
  }
  return sheet;
}

function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

// Adds Password column if missing (fixes sheets created by older versions)
function ensurePasswordColumn(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var found = false;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'Password') { found = true; break; }
  }
  if (!found) {
    var newCol = lastCol + 1;
    sheet.getRange(1, newCol).setValue('Password');
    sheet.getRange(1, newCol)
      .setBackground('#1A3A2A')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sheet.setColumnWidth(newCol, 150);
  }
}

// =================================================================
// REGISTER
// Stores password in plain text -- visible in Candidates sheet
// Columns: ID | Name | Email | Phone | Password | Registered At
// =================================================================
function handleRegister(data) {
  var ss = getOrCreateSS();
  var headers = ['ID', 'Name', 'Email', 'Phone', 'Password', 'Registered At'];
  var sheet = getOrCreateTab(ss, TAB.USERS, headers);

  ensurePasswordColumn(sheet); // Safely adds Password column if old sheet lacks it

  var rows = sheetToObjects(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i]['Email'] && rows[i]['Email'].toString().toLowerCase() === data.email.toLowerCase()) {
      return { status: 'error', message: 'An account with this email already exists.' };
    }
  }

  var id = Date.now().toString();
  sheet.appendRow([
    id,
    data.name,
    data.email.toLowerCase(),
    data.phone || '',
    data.pass,                           // Plain text password stored here
    new Date().toLocaleString('en-IN'),
  ]);

  return {
    status: 'ok',
    user: { id: id, name: data.name, email: data.email.toLowerCase(), phone: data.phone || '' },
  };
}

// =================================================================
// LOGIN
// =================================================================
function handleLogin(data) {
  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(TAB.USERS);
  if (!sheet) {
    return { status: 'error', message: 'No accounts found. Please register first.' };
  }

  var rows = sheetToObjects(sheet);
  var user = null;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (
      r['Email'] && r['Email'].toString().toLowerCase() === data.email.toLowerCase() &&
      r['Password'] && r['Password'].toString() === data.pass
    ) {
      user = r;
      break;
    }
  }

  if (!user) return { status: 'error', message: 'Incorrect email or password.' };

  var completed = getCompletedTests(ss, data.email.toLowerCase());
  return {
    status: 'ok',
    user: { id: user['ID'], name: user['Name'], email: user['Email'], phone: user['Phone'] || '' },
    completed: completed,
  };
}

// =================================================================
// CHECK COMPLETED TESTS
// =================================================================
function handleCheckCompleted(params) {
  var email = (params.email || '').toLowerCase();
  if (!email) return { status: 'error', message: 'No email provided' };
  var ss = getOrCreateSS();
  return { status: 'ok', completed: getCompletedTests(ss, email) };
}

function getCompletedTests(ss, email) {
  var completed = [];
  var tids = [1, 2, 3, 4];
  for (var t = 0; t < tids.length; t++) {
    var tid = tids[t];
    var ts = ss.getSheetByName(TAB['T' + tid]);
    if (!ts) continue;
    var rows = sheetToObjects(ts);
    for (var i = 0; i < rows.length; i++) {
      if (rows[i]['Email'] && rows[i]['Email'].toString().toLowerCase() === email) {
        completed.push(tid);
        break;
      }
    }
  }
  return completed;
}

// =================================================================
// SUBMIT TEST
// =================================================================
function handleSubmitTest(data) {
  var testId = parseInt(data.testId);
  var ss = getOrCreateSS();
  var tabName = TAB['T' + testId];

  var existing = ss.getSheetByName(tabName);
  if (existing) {
    var rows = sheetToObjects(existing);
    for (var i = 0; i < rows.length; i++) {
      if (rows[i]['Email'] && rows[i]['Email'].toString().toLowerCase() === data.email.toLowerCase()) {
        return { status: 'error', message: 'This test has already been submitted.' };
      }
    }
  }

  var qHeaders = Q_HEADERS[testId] || [];
  var headers = ['Row ID', 'Submitted At', 'Name', 'Email', 'Phone', 'Test', 'Time Taken (sec)', 'Auto Submit', 'Tab Switches'];
  for (var i = 0; i < qHeaders.length; i++) {
    headers.push('Q' + (i + 1) + ': ' + qHeaders[i]);
  }
  headers.push('Notes');

  var sheet = getOrCreateTab(ss, tabName, headers);

  var row = [
    Date.now().toString(),
    data.submittedAt || new Date().toLocaleString('en-IN'),
    data.name || '',
    data.email || '',
    data.phone || '',
    data.testTitle || 'Test ' + testId,
    data.timeTaken || 0,
    data.autoSubmit ? 'Yes' : 'No',
    data.tabSwitches || 0,
  ];
  for (var j = 0; j < qHeaders.length; j++) {
    row.push(data['Q' + (j + 1)] || '');
  }
  row.push('');

  sheet.appendRow(row);

  if (parseInt(data.tabSwitches) > 0) {
    sheet.getRange(sheet.getLastRow(), 1, 1, row.length).setBackground('#FFF3E0');
  }

  return { status: 'ok' };
}

// =================================================================
// ADMIN -- GET ALL SUBMISSIONS
// =================================================================
function handleGetSubmissions() {
  var ss = getOrCreateSS();
  var all = [];
  var tids = [1, 2, 3, 4];
  for (var t = 0; t < tids.length; t++) {
    var tid = tids[t];
    var sheet = ss.getSheetByName(TAB['T' + tid]);
    if (!sheet) continue;
    var rows = sheetToObjects(sheet);
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      all.push({
        rowId:       r['Row ID'] || '',
        testId:      tid,
        testTitle:   r['Test'] || '',
        name:        r['Name'] || '',
        email:       r['Email'] || '',
        phone:       r['Phone'] || '',
        submittedAt: r['Submitted At'] || '',
        timeTaken:   r['Time Taken (sec)'] || 0,
        autoSubmit:  r['Auto Submit'] === 'Yes',
        tabSwitches: parseInt(r['Tab Switches']) || 0,
      });
    }
  }
  all.sort(function(a, b) { return String(b.rowId).localeCompare(String(a.rowId)); });
  return { status: 'ok', submissions: all };
}

// =================================================================
// ADMIN -- GET ALL CANDIDATES (includes passwords)
// =================================================================
function handleGetCandidates() {
  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(TAB.USERS);
  if (!sheet) return { status: 'ok', candidates: [] };

  var users = sheetToObjects(sheet);
  var result = [];
  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    var email = (u['Email'] || '').toLowerCase();
    result.push({
      id:           u['ID'] || '',
      name:         u['Name'] || '',
      email:        email,
      phone:        u['Phone'] || '',
      password:     u['Password'] || '',    // Shown in admin Candidates tab
      registeredAt: u['Registered At'] || '',
      completed:    getCompletedTests(ss, email),
    });
  }
  return { status: 'ok', candidates: result };
}

// =================================================================
// ADMIN -- GET FULL DETAIL FOR ONE SUBMISSION
// =================================================================
function handleGetDetail(params) {
  var testId = parseInt(params.testId);
  var email  = (params.email || '').toLowerCase();
  if (!testId || !email) return { status: 'error', message: 'Missing testId or email' };

  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(TAB['T' + testId]);
  if (!sheet) return { status: 'error', message: 'No submissions for this test yet.' };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { status: 'error', message: 'No data found.' };

  var headers = data[0];
  var emailIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    if (headers[h] === 'Email') { emailIdx = h; break; }
  }

  var rowData = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][emailIdx] && data[i][emailIdx].toString().toLowerCase() === email) {
      rowData = data[i];
      break;
    }
  }
  if (!rowData) return { status: 'error', message: 'Submission not found.' };

  var obj = {};
  for (var k = 0; k < headers.length; k++) { obj[headers[k]] = rowData[k]; }

  var answers = {};
  var qh = Q_HEADERS[testId] || [];
  for (var q = 0; q < qh.length; q++) {
    answers['Q' + (q + 1)] = obj['Q' + (q + 1) + ': ' + qh[q]] || '';
  }

  return {
    status: 'ok',
    detail: {
      name:        obj['Name'],
      email:       obj['Email'],
      phone:       obj['Phone'],
      testId:      testId,
      testTitle:   obj['Test'],
      submittedAt: obj['Submitted At'],
      timeTaken:   obj['Time Taken (sec)'],
      autoSubmit:  obj['Auto Submit'] === 'Yes',
      tabSwitches: parseInt(obj['Tab Switches']) || 0,
      answers:     answers,
      notes:       obj['Notes'] || '',
    },
  };
}

// =================================================================
// ADMIN -- SAVE EVALUATOR NOTES
// =================================================================
function handleSaveNotes(data) {
  var testId = parseInt(data.testId);
  var email  = (data.email || '').toLowerCase();
  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(TAB['T' + testId]);
  if (!sheet) return { status: 'error', message: 'Sheet not found.' };

  var allData  = sheet.getDataRange().getValues();
  var headers  = allData[0];
  var emailIdx = -1, notesIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    if (headers[h] === 'Email') emailIdx = h;
    if (headers[h] === 'Notes') notesIdx = h;
  }
  if (notesIdx < 0) return { status: 'error', message: 'Notes column not found.' };

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][emailIdx] && allData[i][emailIdx].toString().toLowerCase() === email) {
      sheet.getRange(i + 1, notesIdx + 1).setValue(data.notes || '');
      return { status: 'ok' };
    }
  }
  return { status: 'error', message: 'Submission not found.' };
}

// =================================================================
// FORGOT PASSWORD — looks up email, sends password via Gmail
// =================================================================
function handleForgotPassword(data) {
  var email = (data.email || '').toLowerCase().trim();
  if (!email) return { status: 'error', message: 'Please enter your email address.' };

  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(TAB.USERS);
  if (!sheet) return { status: 'error', message: 'No accounts found. Please register first.' };

  var rows = sheetToObjects(sheet);
  var user = null;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i]['Email'] && rows[i]['Email'].toString().toLowerCase() === email) {
      user = rows[i];
      break;
    }
  }

  if (!user) return { status: 'error', message: 'No account found with this email address.' };

  // Send password via email
  try {
    var subject = 'Educart BDM Portal — Your Password';
    var body =
      'Hello ' + (user['Name'] || 'there') + ',\n\n' +
      'You requested your password for the Educart BDM Assessment Portal.\n\n' +
      'Your password is: ' + (user['Password'] || '(not set)') + '\n\n' +
      'If you did not request this, please ignore this email.\n\n' +
      '— Educart Assessment Team';
    MailApp.sendEmail(email, subject, body);
    return { status: 'ok', message: 'Password has been sent to your email address.' };
  } catch (err) {
    return { status: 'error', message: 'Could not send email. Please contact the admin.' };
  }
}

// =================================================================
// RESET PASSWORD — verifies email + phone, then sets new password
// =================================================================
function handleResetPassword(data) {
  var email = (data.email || '').toLowerCase().trim();
  var phone = (data.phone || '').trim();
  var newPass = data.newPass || '';

  if (!email) return { status: 'error', message: 'Please enter your email address.' };
  if (!phone) return { status: 'error', message: 'Please enter your phone number.' };
  if (newPass.length < 6) return { status: 'error', message: 'New password must be at least 6 characters.' };

  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(TAB.USERS);
  if (!sheet) return { status: 'error', message: 'No accounts found. Please register first.' };

  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var emailIdx = -1, phoneIdx = -1, passIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    if (headers[h] === 'Email') emailIdx = h;
    if (headers[h] === 'Phone') phoneIdx = h;
    if (headers[h] === 'Password') passIdx = h;
  }
  if (passIdx < 0) return { status: 'error', message: 'Password column not found in sheet.' };

  for (var i = 1; i < allData.length; i++) {
    var rowEmail = (allData[i][emailIdx] || '').toString().toLowerCase();
    var rowPhone = (allData[i][phoneIdx] || '').toString().trim();
    // Normalize phone: strip spaces, dashes, +91 prefix for comparison
    var normInput = phone.replace(/[\s\-\+]/g, '').replace(/^91/, '');
    var normRow = rowPhone.replace(/[\s\-\+]/g, '').replace(/^91/, '');

    if (rowEmail === email && normRow === normInput) {
      sheet.getRange(i + 1, passIdx + 1).setValue(newPass);
      return { status: 'ok', message: 'Password has been reset successfully. You can now sign in.' };
    }
  }

  return { status: 'error', message: 'Email and phone number do not match any account.' };
}
