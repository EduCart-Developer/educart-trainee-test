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

// PASTE YOUR SPREADSHEET ID HERE.
// You can find it in the URL: https://docs.google.com/spreadsheets/d/<THIS_PART>/edit
// If left empty, the script will fall back to finding a file by SHEET_NAME, which
// is fragile (it may create a new empty file if the existing one is owned by a
// different account). Setting SHEET_ID is the reliable way to bind the script
// to the exact sheet your data lives in.
var SHEET_ID   = '';
var SHEET_NAME = 'Educart BDM Responses';

// Tab name candidates. The first existing one wins. This makes the script
// tolerant of old vs. new naming conventions (em-dash vs. hyphen, "&" vs. "and").
var TAB_CANDIDATES = {
  USERS:      ['Candidates', 'Users', 'Registrations'],
  EVALS:      ['Evaluations', 'Evaluation', 'Evals'],
  REATTEMPTS: ['Reattempts', 'Reattempt Grants'],
  T1: ['Test 1 - QB Readiness',         'Test 1 — QB Readiness',         'Test 1'],
  T2: ['Test 2 - Scenarios',            'Test 2 — Scenarios',            'Test 2'],
  T3: ['Test 3 - Academic Structure',   'Test 3 — Academic Structure',   'Test 3'],
  T4: ['Test 4 - NEP and Policy',       'Test 4 — NEP & Policy',         'Test 4 — NEP and Policy', 'Test 4'],
};

// Resolve actual tab name (first existing candidate, or the first candidate as fallback for creation).
function tabName(key) {
  var ss = getOrCreateSS();
  var list = TAB_CANDIDATES[key] || [];
  for (var i = 0; i < list.length; i++) if (ss.getSheetByName(list[i])) return list[i];
  return list[0];
}

// Convenience: { USERS, EVALS, REATTEMPTS, T1..T4 } resolved on demand.
var TAB = {
  get USERS()      { return tabName('USERS'); },
  get EVALS()      { return tabName('EVALS'); },
  get REATTEMPTS() { return tabName('REATTEMPTS'); },
  get T1()         { return tabName('T1'); },
  get T2()         { return tabName('T2'); },
  get T3()         { return tabName('T3'); },
  get T4()         { return tabName('T4'); },
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
    if      (action === 'ping')  { result = { status: 'ok', message: 'Educart BDM API running' }; }
    else if (action === 'sync')  { result = handleSync(); }
    else if (action === 'debug') { result = handleDebug(); }
    else                         { result = { status: 'error', message: 'Unknown action: ' + action }; }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// =================================================================
// DEBUG — visit ?action=debug in the browser to see which spreadsheet
// GAS is bound to + its tab structure. Useful when sync returns empty.
// =================================================================
function handleDebug() {
  var ss = getOrCreateSS();
  var tabs = ss.getSheets().map(function (s) {
    return {
      name:    s.getName(),
      rows:    s.getLastRow(),
      cols:    s.getLastColumn(),
      headers: s.getLastColumn() ? s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0] : [],
    };
  });
  return {
    status:   'ok',
    sheetId:  ss.getId(),
    sheetUrl: ss.getUrl(),
    name:     ss.getName(),
    resolved: {
      USERS:      TAB.USERS,
      EVALS:      TAB.EVALS,
      REATTEMPTS: TAB.REATTEMPTS,
      T1: TAB.T1, T2: TAB.T2, T3: TAB.T3, T4: TAB.T4,
    },
    tabs: tabs,
  };
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
  if (SHEET_ID) {
    try { return SpreadsheetApp.openById(SHEET_ID); }
    catch (e) { throw new Error('SHEET_ID is set but unreachable: ' + e.message); }
  }
  var files = DriveApp.getFilesByName(SHEET_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  // Falling back to creating a fresh sheet is dangerous (it will appear empty
  // and silently shadow your real data). Surface that loudly instead.
  throw new Error(
    'Cannot find a spreadsheet named "' + SHEET_NAME + '" in this account\'s Drive. ' +
    'Set SHEET_ID at the top of this script to your existing sheet ID ' +
    '(from its URL: /spreadsheets/d/<ID>/edit), then redeploy.'
  );
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
    obj.__cells = data[i]; // raw cell values for shape-based fallback
    result.push(obj);
  }
  return result;
}

// Read header row of an existing sheet
function readHeaders(sheet) {
  if (!sheet) return [];
  var lastCol = sheet.getLastColumn();
  if (!lastCol) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

// Build a row array matching the given header order from a value map.
// Headers that aren't in valueMap end up as ''.
function rowFromMap(headers, valueMap) {
  return headers.map(function (h) {
    return valueMap.hasOwnProperty(h) ? valueMap[h] : '';
  });
}

// Heuristic classifier for one cell value
function classify(v) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return 'empty';
  if (s.indexOf('@') > 0 && /\.[a-z]{2,}/i.test(s)) return 'email';
  var digits = s.replace(/\D/g, '');
  if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)) return 'date';
  if (/^\+?\d[\d\s\-]{6,14}$/.test(s) && digits.length >= 7 && digits.length <= 15) return 'phone';
  if (/^\d{10,}$/.test(s)) return 'id_or_phone'; // pure long digits (could be timestamp ID OR a phone w/o spaces)
  return 'text';
}

// Pull a candidate's { name, email, phone, createdAt } out of a sheet row no
// matter which columns its data ended up in. We first trust the standard
// headers, and fall back to scanning all cells when a value looks wrong.
function normalizeUserRow(obj) {
  var cells = obj.__cells || [];
  var hdrEmail     = String(obj['Email']         || obj['email']         || '').trim();
  var hdrName      = String(obj['Name']          || obj['name']          || '').trim();
  var hdrPhone     = String(obj['Phone']         || obj['phone']         || '').trim();
  var hdrCreatedAt = String(obj['Created At']    || obj['Registered At'] || obj['createdAt'] || '').trim();

  // EMAIL — must contain @. Scan all cells if header column doesn't have one.
  var email = (hdrEmail.indexOf('@') > 0) ? hdrEmail : '';
  if (!email) {
    for (var i = 0; i < cells.length; i++) {
      if (classify(cells[i]) === 'email') { email = String(cells[i]).trim(); break; }
    }
  }
  if (!email) return null; // no email anywhere — skip this row

  // PHONE — look for phone-shaped cell
  var phone = '';
  if (classify(hdrPhone) === 'phone' || classify(hdrPhone) === 'id_or_phone') {
    phone = hdrPhone;
  } else {
    for (var p = 0; p < cells.length; p++) {
      var cls = classify(cells[p]);
      var s = String(cells[p]).trim();
      if ((cls === 'phone' || cls === 'id_or_phone') && s !== email) {
        // Skip values that look like a millisecond timestamp ID (13+ digits)
        if (cls === 'id_or_phone' && s.length >= 13) continue;
        phone = s;
        break;
      }
    }
  }

  // CREATED AT — look for date-shaped cell
  var createdAt = (classify(hdrCreatedAt) === 'date') ? hdrCreatedAt : '';
  if (!createdAt) {
    for (var d = 0; d < cells.length; d++) {
      if (classify(cells[d]) === 'date') { createdAt = String(cells[d]).trim(); break; }
    }
  }

  // NAME — anything text-shaped that isn't the email/phone/createdAt
  var name = '';
  if (hdrName && hdrName !== email && hdrName !== phone && classify(hdrName) === 'text') {
    name = hdrName;
  } else {
    for (var n = 0; n < cells.length; n++) {
      var sv = String(cells[n]).trim();
      if (!sv) continue;
      if (sv === email || sv === phone || sv === createdAt) continue;
      if (classify(sv) === 'text') { name = sv; break; }
    }
  }

  return { name: name, email: email.toLowerCase(), phone: phone, createdAt: createdAt };
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
// SIGNUP — writes by header name so it works with any existing layout
// (v3's "ID | Name | Email | Phone | Password | Registered At" or fresh).
// =================================================================
function handleSignup(data) {
  var ss = getOrCreateSS();
  // If the tab doesn't exist yet, create with our preferred header order.
  var sheet = getOrCreateTab(ss, TAB.USERS, ['ID', 'Name', 'Email', 'Phone', 'Password', 'Registered At']);
  var headers = readHeaders(sheet);

  var email = String(data.email || '').toLowerCase().trim();
  if (!email) return { status: 'error', message: 'Email required' };

  // Already registered? Use normalizeUserRow so shifted rows are still matched.
  var data2D = sheet.getDataRange().getValues();
  for (var r = 1; r < data2D.length; r++) {
    var obj = { __cells: data2D[r] };
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = data2D[r][c];
    var n = normalizeUserRow(obj);
    if (n && n.email === email) return { status: 'ok', message: 'Already registered' };
  }

  var nowStamp = data.createdAt || new Date().toLocaleString('en-IN');
  // Map a value into every header variant we might encounter.
  var valueMap = {
    'ID':             Date.now().toString(),
    'Name':           data.name || '',
    'name':           data.name || '',
    'Email':          email,
    'email':          email,
    'Phone':          data.phone || '',
    'phone':          data.phone || '',
    'Password':       '', // empty — passwords aren't sent from frontend
    'Registered At':  nowStamp,
    'Created At':     nowStamp,
    'createdAt':      nowStamp,
  };
  sheet.appendRow(rowFromMap(headers, valueMap));
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
    // Walk bottom-up so row indexes stay stable while deleting. Use
    // normalizeUserRow so we can delete rows whose email landed in the wrong
    // column (e.g. the Name column for shifted rows).
    for (var i = data.length - 1; i >= 1; i--) {
      var obj = { __cells: data[i] };
      for (var c = 0; c < headers.length; c++) obj[headers[c]] = data[i][c];
      var n = normalizeUserRow(obj);
      var em = n ? n.email : '';
      // Also accept literal Email-column match as a fallback for sheets that
      // never had shift issues (e.g. Test sheets, which always populate Email
      // correctly via handleSubmission).
      if (!em) em = String(obj['Email'] || '').toLowerCase().trim();
      if (em && emailSet[em]) { sh.deleteRow(i + 1); deleted++; }
    }
  });
  return { status: 'ok', deleted: deleted };
}

// =================================================================
// SYNC — return everything the admin needs
// =================================================================
function handleSync() {
  var ss = getOrCreateSS();

  // Users — tolerate v3 / shifted / mixed layouts. normalizeUserRow scans the
  // row by cell shape (@ = email, digits = phone, dd/mm/yyyy = date) so a
  // row whose data landed in the wrong columns still gets picked up.
  var users = [];
  var usersSheet = ss.getSheetByName(TAB.USERS);
  var seenEmails = {};
  sheetToObjects(usersSheet).forEach(function (u) {
    var n = normalizeUserRow(u);
    if (!n || !n.email || seenEmails[n.email]) return;
    seenEmails[n.email] = true;
    users.push(n);
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
