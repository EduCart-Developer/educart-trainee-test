// =================================================================
// EDUCART BDM ASSESSMENT -- Google Apps Script Backend (v4)
// =================================================================
// Protocol (matches index.html frontend):
//
// POST body { type: ... }
//   signup             -> add user to Candidates sheet
//   submission         -> append row to Test N sheet (with answers Q1..Qn)
//   evaluation         -> upsert eval row in Evaluations sheet
//   reattempt          -> mark reattempt grant active
//   revoke_reattempt   -> mark reattempt grant inactive
//   reset_test         -> delete a candidate's submission row for one test
//   delete_users       -> bulk-delete users + all their submissions/evals
//
// GET ?action=sync     -> { users, submissions (with answers), evaluations, reattempts }
// GET ?action=ping     -> health check
//
// SETUP
// 1. https://script.google.com -> New Project, paste this whole file, save.
// 2. Deploy -> New Deployment -> Web App (Execute as: Me, Access: Anyone).
// 3. Copy the /exec URL, paste into worker (1).js as GAS_URL, redeploy worker.
// Re-deploy as NEW VERSION after every edit.
// =================================================================

var SHEET_NAME = 'Educart BDM Responses';

var TAB = {
  USERS:       'Candidates',
  EVALS:       'Evaluations',
  REATTEMPTS:  'Reattempts',
  T1: 'Test 1 - QB Readiness',
  T2: 'Test 2 - Scenarios',
  T3: 'Test 3 - Academic Structure',
  T4: 'Test 4 - NEP and Policy',
};

// Question id list per testId — MUST match TESTS array order in index.html.
// These are stored as the answer column header suffix and are how the frontend
// re-maps answers back to questions during sync.
var Q_IDS = {
  1: ['1a','1b','1c','1d','1e','1f','1g','1h','1i','1j'],
  2: ['2a','2b','2c','2d','2e','2f','2g','2h'],
  3: ['3a','3b','3c','3d','3e','3f','3g','3h','3i','3j','3k','3l'],
  4: ['4a','4b','4c','4d','4e','4f','4g','4h','4i'],
};

// =================================================================
// ROUTER
// =================================================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var type = data.type || data.action || '';
    var result;
    if      (type === 'signup')           { result = handleSignup(data); }
    else if (type === 'submission')       { result = handleSubmission(data); }
    else if (type === 'evaluation')       { result = handleEvaluation(data); }
    else if (type === 'reattempt')        { result = handleReattempt(data, 'active'); }
    else if (type === 'revoke_reattempt') { result = handleReattempt(data, 'inactive'); }
    else if (type === 'reset_test')       { result = handleResetTest(data); }
    else if (type === 'delete_users')     { result = handleDeleteUsers(data); }
    else                                  { result = { status: 'error', message: 'Unknown type: ' + type }; }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'ping';
    var result;
    if      (action === 'ping') { result = { status: 'ok', message: 'Educart BDM API running' }; }
    else if (action === 'sync') { result = handleSync(); }
    else                        { result = { status: 'error', message: 'Unknown action: ' + action }; }
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
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return SpreadsheetApp.create(SHEET_NAME);
}

function getOrCreateTab(ss, tabName, headers) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(headers);
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#1A3A2A').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
    for (var i = 0; i < headers.length; i++) {
      sheet.setColumnWidth(i + 1, i < 8 ? 150 : 320);
    }
  } else {
    // Ensure existing tab has all headers (add any missing ones at the end)
    var lastCol = sheet.getLastColumn();
    var existing = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    for (var j = 0; j < headers.length; j++) {
      if (existing.indexOf(headers[j]) < 0) {
        var newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(headers[j])
          .setBackground('#1A3A2A').setFontColor('#FFFFFF').setFontWeight('bold');
      }
    }
  }
  return sheet;
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    result.push(obj);
  }
  return result;
}

function findRowByMatch(sheet, predicate) {
  if (!sheet) return -1;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return -1;
  var headers = data[0];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    if (predicate(obj)) return i + 1; // 1-based sheet row
  }
  return -1;
}

function questionHeaders(testId) {
  var ids = Q_IDS[testId] || [];
  return ids.map(function (qid, i) { return 'Q' + (i + 1) + ':' + qid; });
}

// =================================================================
// SIGNUP
// =================================================================
function handleSignup(data) {
  var ss = getOrCreateSS();
  var sheet = getOrCreateTab(ss, TAB.USERS, ['Name', 'Email', 'Phone', 'Created At']);

  var email = String(data.email || '').toLowerCase().trim();
  if (!email) return { status: 'error', message: 'Email required' };

  // Skip if a row with this email already exists
  var existingRow = findRowByMatch(sheet, function (r) {
    return String(r['Email'] || '').toLowerCase().trim() === email;
  });
  if (existingRow > 0) return { status: 'ok', message: 'Already registered' };

  sheet.appendRow([
    data.name || '',
    email,
    data.phone || '',
    data.createdAt || new Date().toLocaleString('en-IN'),
  ]);
  return { status: 'ok' };
}

// =================================================================
// SUBMIT TEST
// =================================================================
function handleSubmission(data) {
  var testId = parseInt(data.testId);
  if (!testId || !TAB['T' + testId]) return { status: 'error', message: 'Invalid testId' };

  var ss = getOrCreateSS();
  var qHeaders = questionHeaders(testId);
  var headers = ['Row ID', 'Submitted At', 'Name', 'Email', 'Phone', 'Test',
                 'Time Taken (sec)', 'Auto Submit', 'Tab Switches'].concat(qHeaders);

  var sheet = getOrCreateTab(ss, TAB['T' + testId], headers);
  var email = String(data.email || '').toLowerCase().trim();

  // De-dup: if this email already has a row in this test sheet, replace it
  // (acts as upsert so a re-submission after a reset overwrites cleanly).
  var existingRow = findRowByMatch(sheet, function (r) {
    return String(r['Email'] || '').toLowerCase().trim() === email;
  });
  if (existingRow > 0) sheet.deleteRow(existingRow);

  var row = [
    data.id || Date.now().toString(),
    data.submittedAt || new Date().toLocaleString('en-IN'),
    data.name || '',
    email,
    data.phone || '',
    data.testTitle || 'Test ' + testId,
    data.timeTaken || 0,
    data.autoSubmit ? 'Yes' : 'No',
    data.tabSwitches || 0,
  ];
  for (var i = 0; i < qHeaders.length; i++) {
    row.push(data['Q' + (i + 1)] || '');
  }

  sheet.appendRow(row);
  if (parseInt(data.tabSwitches) > 0) {
    sheet.getRange(sheet.getLastRow(), 1, 1, row.length).setBackground('#FFF3E0');
  }
  return { status: 'ok' };
}

// =================================================================
// EVALUATION (upsert by submissionId)
// =================================================================
function handleEvaluation(data) {
  var ss = getOrCreateSS();
  var headers = ['Submission ID', 'Evaluated At', 'Name', 'Email', 'Phone',
                 'Test ID', 'Test Title', 'Total Score', 'Max Score',
                 'Percentage', 'Grade', 'Overall Comment', 'Scores JSON',
                 'Comments JSON', 'Tab Switches', 'Auto Submit'];
  var sheet = getOrCreateTab(ss, TAB.EVALS, headers);

  var subId = String(data.submissionId || '').trim();
  if (!subId) return { status: 'error', message: 'submissionId required' };

  var testId = parseInt(data.testId);
  var qIds = Q_IDS[testId] || [];
  var scores = {}, comments = {};
  qIds.forEach(function (qid, i) {
    if (data['Q' + (i + 1) + '_score']   !== undefined && data['Q' + (i + 1) + '_score']   !== '') scores[qid]   = data['Q' + (i + 1) + '_score'];
    if (data['Q' + (i + 1) + '_comment'] !== undefined && data['Q' + (i + 1) + '_comment'] !== '') comments[qid] = data['Q' + (i + 1) + '_comment'];
  });

  var row = [
    subId,
    data.evaluatedAt || new Date().toLocaleString('en-IN'),
    data.name || '',
    String(data.email || '').toLowerCase().trim(),
    data.phone || '',
    testId,
    data.testTitle || '',
    data.totalScore || 0,
    data.maxScore || 0,
    data.percentage || 0,
    data.grade || '',
    data.overallComment || '',
    JSON.stringify(scores),
    JSON.stringify(comments),
    data.tabSwitches || 0,
    data.autoSubmit || 'No',
  ];

  var existingRow = findRowByMatch(sheet, function (r) {
    return String(r['Submission ID'] || '').trim() === subId;
  });
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { status: 'ok' };
}

// =================================================================
// REATTEMPT GRANT
// =================================================================
function handleReattempt(data, status) {
  var ss = getOrCreateSS();
  var headers = ['Email', 'Test ID', 'Status', 'Updated At'];
  var sheet = getOrCreateTab(ss, TAB.REATTEMPTS, headers);

  var email = String(data.email || '').toLowerCase().trim();
  var testId = parseInt(data.testId);
  if (!email || !testId) return { status: 'error', message: 'email + testId required' };

  var row = [email, testId, status, new Date().toLocaleString('en-IN')];
  var existingRow = findRowByMatch(sheet, function (r) {
    return String(r['Email'] || '').toLowerCase().trim() === email
      && parseInt(r['Test ID']) === testId;
  });
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { status: 'ok' };
}

// =================================================================
// RESET TEST — delete one candidate's submission for one test
// =================================================================
function handleResetTest(data) {
  var ss = getOrCreateSS();
  var testId = parseInt(data.testId);
  var email = String(data.email || '').toLowerCase().trim();
  if (!email || !testId) return { status: 'error', message: 'email + testId required' };

  var sheet = ss.getSheetByName(TAB['T' + testId]);
  if (sheet) {
    var r = findRowByMatch(sheet, function (row) {
      return String(row['Email'] || '').toLowerCase().trim() === email;
    });
    if (r > 0) sheet.deleteRow(r);
  }

  var evalSheet = ss.getSheetByName(TAB.EVALS);
  if (evalSheet) {
    var er = findRowByMatch(evalSheet, function (row) {
      return String(row['Email'] || '').toLowerCase().trim() === email
        && parseInt(row['Test ID']) === testId;
    });
    if (er > 0) evalSheet.deleteRow(er);
  }
  return { status: 'ok' };
}

// =================================================================
// BULK DELETE USERS (+ all their submissions + their evaluations)
// =================================================================
function handleDeleteUsers(data) {
  var emails = (data.emails || []).map(function (e) {
    return String(e || '').toLowerCase().trim();
  }).filter(function (e) { return !!e; });
  if (!emails.length) return { status: 'error', message: 'No emails provided' };

  var emailSet = {};
  emails.forEach(function (e) { emailSet[e] = true; });

  var ss = getOrCreateSS();
  var sheetsToClean = [TAB.USERS, TAB.EVALS, TAB.REATTEMPTS, TAB.T1, TAB.T2, TAB.T3, TAB.T4];
  var deleted = 0;

  sheetsToClean.forEach(function (tabName) {
    var sh = ss.getSheetByName(tabName);
    if (!sh) return;
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    var headers = data[0];
    var emailIdx = -1;
    for (var h = 0; h < headers.length; h++) if (headers[h] === 'Email') { emailIdx = h; break; }
    if (emailIdx < 0) return;
    // Walk bottom-up so row indexes stay stable while deleting
    for (var i = data.length - 1; i >= 1; i--) {
      var em = String(data[i][emailIdx] || '').toLowerCase().trim();
      if (emailSet[em]) { sh.deleteRow(i + 1); deleted++; }
    }
  });
  return { status: 'ok', deleted: deleted };
}

// =================================================================
// SYNC — return everything the admin needs
// =================================================================
function handleSync() {
  var ss = getOrCreateSS();

  // Users
  var users = [];
  var usersSheet = ss.getSheetByName(TAB.USERS);
  sheetToObjects(usersSheet).forEach(function (u) {
    var em = String(u['Email'] || '').toLowerCase().trim();
    if (!em) return;
    users.push({
      name:      u['Name'] || '',
      email:     em,
      phone:     u['Phone'] || '',
      createdAt: u['Created At'] || '',
    });
  });

  // Submissions (with answers re-keyed by q.id so the frontend slots them correctly)
  var submissions = [];
  [1, 2, 3, 4].forEach(function (testId) {
    var sheet = ss.getSheetByName(TAB['T' + testId]);
    if (!sheet) return;
    var qIds = Q_IDS[testId] || [];

    // Build a position -> header-key map ONCE per sheet, robust to any header
    // format the spreadsheet might have ('Q1', 'Q1:1a', 'Q1: Question text', ...).
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var keyByQNum = {};
    headerRow.forEach(function (h) {
      var m = String(h || '').match(/^Q(\d+)\b/);
      if (m) keyByQNum[parseInt(m[1])] = h;
    });

    sheetToObjects(sheet).forEach(function (r) {
      var email = String(r['Email'] || '').toLowerCase().trim();
      if (!email) return;
      var answers = {};
      qIds.forEach(function (qid, i) {
        var headerKey = keyByQNum[i + 1];
        var v = headerKey ? r[headerKey] : '';
        answers[qid] = v == null ? '' : String(v);
      });
      submissions.push({
        id:          r['Row ID'] || '',
        name:        r['Name'] || '',
        email:       email,
        phone:       r['Phone'] || '',
        testId:      testId,
        testTitle:   r['Test'] || ('Test ' + testId),
        submittedAt: r['Submitted At'] || '',
        timeTaken:   parseInt(r['Time Taken (sec)']) || 0,
        autoSubmit:  String(r['Auto Submit']) === 'Yes',
        tabSwitches: parseInt(r['Tab Switches']) || 0,
        answers:     answers,
      });
    });
  });

  // Evaluations
  var evaluations = [];
  var evalSheet = ss.getSheetByName(TAB.EVALS);
  sheetToObjects(evalSheet).forEach(function (e) {
    var scores = {}, comments = {};
    try { scores = JSON.parse(e['Scores JSON'] || '{}'); } catch (_) {}
    try { comments = JSON.parse(e['Comments JSON'] || '{}'); } catch (_) {}
    evaluations.push({
      submissionId:   String(e['Submission ID'] || ''),
      scores:         scores,
      comments:       comments,
      overallComment: e['Overall Comment'] || '',
    });
  });

  // Reattempts
  var reattempts = [];
  var raSheet = ss.getSheetByName(TAB.REATTEMPTS);
  sheetToObjects(raSheet).forEach(function (r) {
    reattempts.push({
      email:  String(r['Email'] || '').toLowerCase().trim(),
      testId: parseInt(r['Test ID']) || 0,
      status: r['Status'] || 'inactive',
    });
  });

  return {
    status: 'ok',
    users: users,
    submissions: submissions,
    evaluations: evaluations,
    reattempts: reattempts,
  };
}
