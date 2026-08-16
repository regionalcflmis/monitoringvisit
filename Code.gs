/**
 * CODE.GS — MoneyWise CFL Monitoring Visit App
 * ------------------------------------------------------------------
 * Covers BOTH checklists from Monitoring_Visit_Data_Form.xlsx:
 *   - CFL Reporting format  -> 17 indicators  (CFL_ITEMS in Items.gs)
 *   - Session Monitoring Report -> 14 indicators (SESSION_ITEMS in Items.gs)
 * Every indicator's full wording is preserved verbatim (see Items.gs).
 *
 * SETUP (do this once):
 *   1. Open Extensions > Apps Script on a NEW Google Sheet.
 *   2. Paste these files in exactly (Code.gs, Items.gs, MasterData.gs,
 *      Index.html, PdfTemplate.html), plus replace appsscript.json
 *      content (View > Show manifest file).
 *   3. Run the function `setupProject` once (choose it from the
 *      function dropdown and click Run). Approve the permissions asked.
 *   4. Deploy > New deployment > Web app > Execute as "Me",
 *      Who has access "Anyone with the link". Copy the URL and share it.
 */

var SHEET_CFL_MASTER   = 'CFL Master';
var SHEET_EMP_MASTER   = 'Employee Master';
var SHEET_CFL_SUB      = 'CFL Submissions';
var SHEET_SESSION_SUB  = 'Session Submissions';
var SHEET_PINS         = 'Employee PINs';
var SHEET_EDIT_REQUESTS = 'Edit Requests';
var DRIVE_FOLDER_NAME  = 'MoneyWise CFL Monitoring Reports';
var TAT_LIMIT_DAYS     = 7;   // <=7 days => Green, else Red (per your instruction)

// The officer filling the form NEVER sees or types these — report is
// auto-emailed to exactly this fixed list on every submit, silently.
// Edit this list (max 5) whenever the recipients need to change.
var RECIPIENT_EMAILS = [
  'vivek.mishra@ext-crisil.com',
  'regionalcflmis@gmail.com','vinay.sharma1@ext-crisil.com'
];

// ---------------------------------------------------------------------
// ONE-TIME SETUP
// ---------------------------------------------------------------------
function setupProject() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  writeMasterSheet_(ss, SHEET_CFL_MASTER,
    ['CFL Name', 'BCC Code', 'Phase', 'State', 'District', 'Base Block', 'Adjacent Block 1', 'Adjacent Block 2', 'Bank Name'],
    CFL_MASTER_DATA);

  // Employee Master is NEVER overwritten if it already exists — you may
  // have added an "Email" column (or edited rows) directly in the sheet,
  // and re-running setup must never wipe that out.
  if (!ss.getSheetByName(SHEET_EMP_MASTER)) {
    writeMasterSheet_(ss, SHEET_EMP_MASTER,
      ['Name', 'Designation', 'State', 'District', 'Zone', 'CFL Name', 'Email'],
      EMPLOYEE_MASTER_DATA.map(function (r) { return r.concat(['']); }));
  }

  ensurePinsSheet_(ss);

  ensureSubmissionSheet_(ss, SHEET_CFL_SUB, CFL_ITEMS, 'CFL');
  ensureSubmissionSheet_(ss, SHEET_SESSION_SUB, SESSION_ITEMS, 'SESSION');

  getOrCreateFolder_(); // pre-create the Drive folder

  var msg = 'Setup complete. Master data loaded (' + CFL_MASTER_DATA.length +
    ' CFLs, ' + EMPLOYEE_MASTER_DATA.length + ' employees). Now deploy as a Web App.';
  Logger.log(msg);
  // SpreadsheetApp.getUi() only works when the Sheet itself is open in a
  // browser tab — running this from the Apps Script editor's "Run" button
  // without the Sheet open throws an error there. That failure is
  // harmless (setup above already completed) — swallow it so it never
  // surfaces as a scary "unknown error".
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* fine — see Logger.log above */ }
}

function writeMasterSheet_(ss, name, headers, rows) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.setFrozenRows(1);
}

/**
 * Creates the "Employee PINs" sheet with one row per unique officer name
 * and a blank PIN column, WITHOUT touching it if it already exists (so
 * PINs you've already typed in are never wiped by re-running setupProject).
 * Fill in a 4-digit PIN for whichever officers should be required to enter
 * one — leave it blank for anyone who shouldn't need a PIN yet.
 */
function ensurePinsSheet_(ss) {
  var sh = ss.getSheetByName(SHEET_PINS);
  if (sh) return sh; // never overwrite PINs already set

  sh = ss.insertSheet(SHEET_PINS);
  var seen = {};
  var rows = [];
  EMPLOYEE_MASTER_DATA.forEach(function (r) {
    var name = r[0];
    if (seen[name]) return;
    seen[name] = true;
    rows.push([name, '']);
  });
  sh.getRange(1, 1, 1, 2).setValues([['Officer Name', 'PIN (4 digit, blank = no PIN required)']]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 260);
  return sh;
}

/**
 * Checks a PIN against the Employee PINs sheet.
 * - No row for this officer, or PIN cell blank → not required, always ok.
 * - PIN set and matches what was entered → ok.
 * - PIN set and does NOT match → fails.
 */
function verifyPin_(officerName, pin) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PINS);
  if (!sh || sh.getLastRow() < 2) return { ok: true, required: false };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(officerName)) {
      var stored = String(data[i][1] || '').trim();
      if (!stored) return { ok: true, required: false };
      return { ok: stored === String(pin || '').trim(), required: true };
    }
  }
  return { ok: true, required: false };
}

// Client-callable pre-check, so the officer gets immediate feedback before
// photos are uploaded. submitReport() ALSO re-checks server-side — never
// trust a client-only check for something used as a security gate.
function checkOfficerPin(officerName, pin) {
  return verifyPin_(officerName, pin);
}

function ensureSubmissionSheet_(ss, name, items, reportType) {
  var expected = baseSubmissionHeaders_(reportType).concat(itemHeaders_(items)).concat(followUpHeaders_());
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, expected.length).setValues([expected]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // Sheet already exists (may have data from an older version of this script).
  // Never wipe it — just append any NEW columns to the end so old rows survive.
  var lastCol = Math.max(1, sh.getLastColumn());
  var current = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var missing = expected.filter(function (h) { return current.indexOf(h) === -1; });
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  sh.setFrozenRows(1);
  return sh;
}

function baseSubmissionHeaders_(reportType) {
  var h = ['Submission ID', 'Timestamp', 'Report Type', 'CFL Name', 'District', 'State',
    'BCC Code', 'Bank Name', 'Phase', 'Base Block', 'Adjacent Block', 'Visit Date',
    'Officer Name', 'Officer Designation'];
  if (reportType === 'SESSION') {
    h = h.concat(['Consultant Staff Designation', 'Session Observer / Stakeholders',
      'Type of Session', 'Venue of Session']);
  }
  return h;
}

// Sheet column name for a given indicator + field — uses the FULL
// indicator name (not "Sr2") so the sheet is organized indicator-wise.
function itemColName_(it, field) {
  return it.head + ' - ' + field;
}

function itemHeaders_(items) {
  var h = [];
  items.forEach(function (it) {
    h.push(itemColName_(it, 'Score'));
    h.push(itemColName_(it, 'Observation'));
    h.push(itemColName_(it, 'Suggestion'));
    h.push(itemColName_(it, 'Timeline Date'));
    h.push(itemColName_(it, 'Completion Date'));
    if (it.staffing) h.push(itemColName_(it, 'Staffing Details'));
  });
  return h;
}

function followUpHeaders_() {
  return ['Total Score', 'Max Score', 'Score %', 'Grade', 'Photo URLs',
    'TAT Green Count', 'TAT Red Count', 'TAT Pending Count',
    'Sign-off Name', 'Sign-off Designation', 'Sign-off District', 'Sign-off Zone',
    'Recipient Emails', 'PDF Link', 'Excel Link', 'Logged-in Google Account', 'Logged-in Email',
    'Form Opened At', 'Fill Duration (min)', 'Photo Location', 'Photo Location Map Link',
    'Submit Location', 'Submit Location Map Link'];
}

function formatLocation_(loc) {
  if (!loc || loc === 'denied' || typeof loc !== 'object') return 'Not available (permission not given)';
  return loc.lat.toFixed(6) + ', ' + loc.lng.toFixed(6) + ' (±' + (loc.accuracy || '?') + 'm)';
}

function mapLink_(loc) {
  if (!loc || loc === 'denied' || typeof loc !== 'object') return '';
  return 'https://maps.google.com/?q=' + loc.lat + ',' + loc.lng;
}

/**
 * ONE-TIME CLEANUP: if your "CFL Submissions" sheet was created before the
 * fix that keeps session-only fields (Type of Session, Venue of Session,
 * etc.) out of it, those columns may still be sitting there from the old
 * version — this never happens automatically (existing columns are never
 * auto-deleted, only new ones get added). Run this once from the function
 * dropdown to remove them safely. Never touches "Session Submissions".
 */
function cleanupCflSessionOnlyColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_CFL_SUB);
  if (!sh) { Logger.log('"' + SHEET_CFL_SUB + '" sheet not found.'); try { SpreadsheetApp.getUi().alert('"' + SHEET_CFL_SUB + '" sheet not found.'); } catch (e) {} return; }

  var toRemove = ['Consultant Staff Designation', 'Session Observer / Stakeholders',
    'Type of Session', 'Venue of Session'];
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var removed = [];
  for (var c = headers.length - 1; c >= 0; c--) {
    if (toRemove.indexOf(headers[c]) > -1) {
      sh.deleteColumn(c + 1);
      removed.push(headers[c]);
    }
  }
  var cleanupMsg = removed.length ?
    ('Removed from CFL Submissions: ' + removed.join(', ')) :
    'Nothing to remove — CFL Submissions is already clean.';
  Logger.log(cleanupMsg);
  try { SpreadsheetApp.getUi().alert(cleanupMsg); } catch (e) { /* fine — see Logger.log above */ }
}

/**
 * DIAGNOSTIC: lists the exact current header row of both submission
 * sheets, in order, with column letters — run this if a column (like
 * Timestamp) seems to be missing, to see exactly what's there and where.
 */
/**
 * DIAGNOSTIC: checks "Employee Master" specifically for the login-email
 * problem — shows the exact header row, and lists every row's Name +
 * Email so you can see typos/missing emails at a glance. Run this from
 * the function dropdown, then check View → Logs.
 */
/**
 * DIAGNOSTIC: isolates whether Drive access itself is blocked (org policy,
 * account restriction) vs something specific to uploadPhoto's logic. Run
 * this from the function dropdown, then check View → Logs.
 */
function testDriveAccess() {
  try {
    var f = DriveApp.createFolder('CFL-test-delete-me-' + new Date().getTime());
    Logger.log('✅ Drive access WORKS — created folder: ' + f.getUrl());
    Logger.log('Owner account: ' + Session.getEffectiveUser().getEmail());
    f.setTrashed(true); // clean up
    Logger.log('(test folder trashed automatically)');
  } catch (e) {
    Logger.log('❌ Drive access BLOCKED: ' + e.message);
    Logger.log('Effective user: ' + Session.getEffectiveUser().getEmail());
    Logger.log('This is almost always a Google Workspace admin policy blocking ' +
      'Apps Script / API-based Drive access for this account — not a code bug. ' +
      'Ask your Workspace admin to allow Drive API access for Apps Script, or ' +
      'deploy this project from a personal (non-organization) Google account instead.');
  }
  try { SpreadsheetApp.getUi().alert('Test complete — check View > Logs for the result.'); } catch (e) {}
}

/**
 * ONE-TIME FIX: sets "Login Exempt" to FALSE for every row in Employee
 * Master — use this if it got set to TRUE for everyone by mistake (e.g.
 * when importing/copying data into a new sheet). After running this,
 * only officers the admin explicitly exempts via the Admin Panel will
 * be able to log in without email.
 */
function resetAllLoginExemptToFalse() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMP_MASTER);
  if (!sh) { Logger.log('❌ "Employee Master" sheet not found.'); return; }

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var exemptCol = headers.indexOf('Login Exempt');
  if (exemptCol === -1) {
    Logger.log('❌ No "Login Exempt" column found — nothing to reset.');
    return;
  }

  var lastRow = sh.getLastRow();
  if (lastRow < 2) { Logger.log('No data rows found.'); return; }

  var range = sh.getRange(2, exemptCol + 1, lastRow - 1, 1);
  var values = range.getValues().map(function () { return ['FALSE']; });
  range.setValues(values);

  var msg = 'Done — "Login Exempt" set to FALSE for all ' + (lastRow - 1) + ' rows.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* fine — see Logger.log above */ }
}

function diagnoseEmployeeEmails() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMP_MASTER);
  if (!sh) { Logger.log('❌ "Employee Master" sheet not found.'); return; }

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Logger.log('=== Employee Master headers (exact, in order) ===');
  headers.forEach(function (h, i) { Logger.log((i + 1) + '. "' + h + '"'); });

  var emailColIdx = headers.indexOf('Email');
  if (emailColIdx === -1) {
    Logger.log('❌ PROBLEM FOUND: no column is named exactly "Email" (capital E). ' +
      'Rename your email column header to exactly: Email');
  } else {
    Logger.log('✅ "Email" column found at position ' + (emailColIdx + 1));
  }

  Logger.log('=== Rows (Name -> Email) ===');
  var live = getEmployeeMasterLive_();
  live.forEach(function (e) {
    Logger.log(e.name + ' -> "' + (e.email || '(EMPTY)') + '"');
  });

  try { SpreadsheetApp.getUi().alert('Diagnostic likh diya — Apps Script editor me "View > Logs" (Ctrl+Enter) se dekh lein.'); } catch (e) { /* fine — check View > Logs directly */ }
}

function listSheetHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_CFL_SUB, SHEET_SESSION_SUB].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { Logger.log(name + ': sheet not found'); return; }
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    Logger.log('=== ' + name + ' (' + headers.length + ' columns) ===');
    headers.forEach(function (h, i) {
      Logger.log((i + 1) + '. ' + h);
    });
  });
  try { SpreadsheetApp.getUi().alert('Header list likh diya gaya hai — Apps Script editor me "View > Logs" (ya Ctrl+Enter) se dekh lein.'); } catch (e) { /* fine — check View > Logs directly */ }
}

function getOrCreateFolder_() {
  var it = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function getOrCreateSubfolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ---------------------------------------------------------------------
// WEB APP ENTRY
// ---------------------------------------------------------------------
function getWebAppUrl_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || 'form';
  var appUrl = getWebAppUrl_();

  if (page === 'dashboard') {
    try {
      var dash = HtmlService.createTemplateFromFile('Dashboard');
      dash.webAppUrl = appUrl;
      return dash.evaluate()
        .setTitle('MoneyWise CFL Monitoring Dashboard')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:Arial;padding:24px;max-width:520px;margin:0 auto;">' +
        '<h3 style="color:#cf222e;">Dashboard load nahi ho paya</h3>' +
        '<p>Apps Script project me <b>"Dashboard"</b> naam ki HTML file honi chahiye ' +
        '(exact naam, .html extension nahi likhna file naam field me — sirf "Dashboard" ' +
        'type karna, type dropdown me "HTML" select karna).</p>' +
        '<p>File add/paste karne ke baad <b>Deploy → Manage deployments → Edit → ' +
        'New version → Deploy</b> zaroor karein.</p>' +
        '<p style="color:#6e7781;font-size:12px;">Technical error: ' + err.message + '</p>' +
        '<p><a href="' + appUrl + '">← Form par wapas jaayein</a></p></div>'
      ).setTitle('Dashboard Error');
    }
  }

  if (page === 'form') {
    var form = HtmlService.createTemplateFromFile('Index');
    form.webAppUrl = appUrl;
    return form.evaluate()
      .setTitle('MoneyWise CFL Monitoring Visit Form')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'admin') {
    var admin = HtmlService.createTemplateFromFile('AdminPanel');
    admin.webAppUrl = appUrl;
    return admin.evaluate()
      .setTitle('MoneyWise CFL Admin Panel')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'mysubmissions') {
    var mysub = HtmlService.createTemplateFromFile('MySubmissions');
    mysub.webAppUrl = appUrl;
    return mysub.evaluate()
      .setTitle('My Submissions')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // Default: the Home portal — ONE link that lists the form, dashboard,
  // and any future forms added later, so nobody needs a separate link
  // for each tool.
  try {
    var home = HtmlService.createTemplateFromFile('Home');
    home.webAppUrl = appUrl;
    return home.evaluate()
      .setTitle('MoneyWise CFL Monitoring System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    // Home.html not added yet — fall back straight to the form so the
    // link still works while it's being set up.
    var fallback = HtmlService.createTemplateFromFile('Index');
    fallback.webAppUrl = appUrl;
    return fallback.evaluate()
      .setTitle('MoneyWise CFL Monitoring Visit Form')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------
// DATA APIS FOR THE CLIENT
// ---------------------------------------------------------------------
function getLoggedInUser() {
  try { return Session.getActiveUser().getEmail() || 'Not detected'; }
  catch (e) { return 'Not detected'; }
}

// Reads "Employee Master" LIVE from the sheet (by header name, not fixed
// position) so any column you add manually there — like "Email" — is
// picked up automatically without needing a code change.
function getEmployeeMasterLive_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMP_MASTER);
  if (!sh || sh.getLastRow() < 2) return [];
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = function (name) { return headers.indexOf(name); };
  var nameCol = col('Name'), desigCol = col('Designation'), stateCol = col('State'),
    distCol = col('District'), zoneCol = col('Zone'), cflCol = col('CFL Name'),
    emailCol = col('Email'), exemptCol = col('Login Exempt');

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  return data.filter(function (r) { return r[nameCol]; }).map(function (r) {
    return {
      name: r[nameCol], designation: desigCol > -1 ? r[desigCol] : '',
      state: stateCol > -1 ? r[stateCol] : '', district: distCol > -1 ? r[distCol] : '',
      zone: zoneCol > -1 ? r[zoneCol] : '', cfl: cflCol > -1 ? r[cflCol] : '',
      email: emailCol > -1 ? String(r[emailCol] || '').trim().toLowerCase() : '',
      loginExempt: exemptCol > -1 && String(r[exemptCol] || '').trim().toUpperCase() === 'TRUE'
    };
  });
}

// Lets someone marked "Login Exempt" in Employee Master get in WITHOUT an
// email — they select their name instead. Admin controls who gets this
// via the Admin Panel.
function checkLoginExempt(name) {
  var match = getEmployeeMasterLive_().filter(function (e) { return e.name === name; })[0];
  return { ok: !!(match && match.loginExempt), name: match ? match.name : '' };
}

// ---------------------------------------------------------------------
// ADMIN PANEL — edit an employee's email, or exempt them from needing one.
// Only the Admin Panel UI itself is admin-gated (client + page routing);
// these two write functions are simple by design, matching the rest of
// this system's PIN/email checks (no deeper server-side admin auth layer).
// ---------------------------------------------------------------------
function getAdminEmployeeList() {
  return getEmployeeMasterLive_();
}

function adminUpdateEmployee(name, newEmail, loginExempt) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMP_MASTER);
  if (!sh) throw new Error('Employee Master sheet not found.');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var nameCol = headers.indexOf('Name');
  var emailCol = headers.indexOf('Email');
  var exemptCol = headers.indexOf('Login Exempt');

  if (emailCol === -1) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue('Email');
    emailCol = sh.getLastColumn() - 1;
    headers.push('Email');
  }
  if (exemptCol === -1) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue('Login Exempt');
    exemptCol = sh.getLastColumn() - 1;
    headers.push('Login Exempt');
  }

  // A single Area Manager has ONE ROW PER ASSIGNED CFL in this sheet —
  // Email/Login Exempt are person-level, so every one of their rows must
  // be updated, not just the first match (that was a real bug before).
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var updated = 0;
  for (var i = 0; i < data.length; i++) {
    if (data[i][nameCol] === name) {
      var row = i + 2;
      sh.getRange(row, emailCol + 1).setValue(String(newEmail || '').trim().toLowerCase());
      sh.getRange(row, exemptCol + 1).setValue(loginExempt ? 'TRUE' : 'FALSE');
      updated++;
    }
  }
  if (updated === 0) throw new Error('Employee "' + name + '" not found in Employee Master.');
  return { ok: true, rowsUpdated: updated };
}

/**
 * Adds a NEW row to Employee Master — for a brand-new officer, OR to give
 * an existing officer an additional/transferred CFL assignment (their
 * Email/Login Exempt should match their existing rows if they have any).
 */
function adminAddEmployeeRow(data) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMP_MASTER);
  if (!sh) throw new Error('Employee Master sheet not found.');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var row = headers.map(function (h) {
    var key = { 'Name': 'name', 'Designation': 'designation', 'State': 'state',
      'District': 'district', 'Zone': 'zone', 'CFL Name': 'cfl', 'Email': 'email' }[h];
    if (h === 'Login Exempt') return data.loginExempt ? 'TRUE' : 'FALSE';
    return key ? (data[key] || '') : '';
  });
  sh.appendRow(row);
  return { ok: true };
}

/**
 * Removes one CFL assignment row (Name + CFL Name pair) — e.g. when an
 * officer transfers away from a CFL/district, or a duplicate/wrong row
 * needs cleanup. Never touches other rows for the same person.
 */
function adminDeleteEmployeeRow(name, cflName) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMP_MASTER);
  if (!sh) throw new Error('Employee Master sheet not found.');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var nameCol = headers.indexOf('Name'), cflCol = headers.indexOf('CFL Name');
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][nameCol] === name && data[i][cflCol] === cflName) {
      sh.deleteRow(i + 2);
      return { ok: true };
    }
  }
  throw new Error('Row not found for ' + name + ' / ' + cflName);
}

// Checks if an email belongs to someone in Employee Master. Used both for
// the login gate (client pre-check) and inside submitReport (server-side,
// never trust the client alone).
// Admin emails always get login access, regardless of what's in Employee
// Master — a fallback so access never depends solely on that sheet being
// kept up to date. Admins also see the Dashboard link; everyone else
// (Area Managers) only sees Form Entry.
var ADMIN_EMAILS = ['regionalcflmis@gmail.com'];

function verifyLoginEmail_(email) {
  if (!email) return { ok: false, isAdmin: false };
  var norm = String(email).trim().toLowerCase();
  if (ADMIN_EMAILS.indexOf(norm) > -1) {
    return { ok: true, name: 'Admin', isAdmin: true };
  }
  var match = getEmployeeMasterLive_().filter(function (e) { return e.email && e.email === norm; })[0];
  return { ok: !!match, name: match ? match.name : '', isAdmin: false };
}

function checkLoginEmail(email) {
  return verifyLoginEmail_(email);
}

// Returns the list of CFL names a specific officer is assigned to (from
// their Employee Master rows — one row per assigned CFL). Used to scope
// a non-admin officer's CFL dropdown to only their own CFLs.
function getEmployeeAssignedCFLs(name) {
  return getEmployeeMasterLive_()
    .filter(function (e) { return e.name === name; })
    .map(function (e) { return e.cfl; })
    .filter(Boolean);
}

/**
 * MY SUBMISSIONS + EDIT REQUEST SYSTEM
 * ------------------------------------------------------------------
 * Officers can see only their OWN past submissions (matched by login
 * email, or by officer name for "no-email" exempt logins), and can
 * request permission to edit one. The request gets emailed to admin and
 * logged in an "Edit Requests" sheet; admin approves/rejects from the
 * Admin Panel. Approval currently just flips the status and notifies the
 * officer — reopening the actual form pre-filled for editing is a
 * follow-up feature, not built yet.
 */
function getMySubmissions(loginId, isAdmin) {
  if (!loginId) return [];
  var isNoLogin = loginId.indexOf('NOLOGIN:') === 0;
  var officerName = isNoLogin ? loginId.substring(8) : '';

  var results = [];
  [SHEET_CFL_SUB, SHEET_SESSION_SUB].forEach(function (sheetName) {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var col = function (name) { return headers.indexOf(name); };
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    data.forEach(function (d) {
      // Admin sees every submission from every officer — everyone else
      // only sees their own (matched by login email, or by officer name
      // for "no-email" exempt logins).
      var mine = isAdmin || (isNoLogin ?
        d[col('Officer Name')] === officerName :
        String(d[col('Logged-in Email')] || '').toLowerCase() === loginId.toLowerCase());
      if (!mine) return;
      results.push({
        submissionId: d[col('Submission ID')], reportType: sheetName === SHEET_CFL_SUB ? 'CFL' : 'SESSION',
        cflName: d[col('CFL Name')], visitDate: d[col('Visit Date')] ? formatDate_(d[col('Visit Date')]) : '',
        officerName: d[col('Officer Name')] || '',
        scorePct: d[col('Score %')], grade: d[col('Grade')],
        pdfLink: d[col('PDF Link')], excelLink: d[col('Excel Link')],
        timestamp: d[col('Timestamp')] ? formatDate_(d[col('Timestamp')]) : ''
      });
    });
  });
  results.sort(function (a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });
  return results;
}

function ensureEditRequestsSheet_(ss) {
  var sh = ss.getSheetByName(SHEET_EDIT_REQUESTS);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_EDIT_REQUESTS);
  sh.getRange(1, 1, 1, 7).setValues([['Request ID', 'Submission ID', 'Report Type', 'Requested By',
    'Reason', 'Status', 'Requested At']]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function requestEditPermission(submissionId, reportType, reason, loginId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureEditRequestsSheet_(ss);
  var requestId = 'REQ-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  var now = new Date();
  sh.appendRow([requestId, submissionId, reportType, loginId, reason || '', 'Pending', now]);

  var recipients = RECIPIENT_EMAILS.slice(0, 5);
  var subject = '[Edit Request] ' + submissionId + ' — approval needed';
  var body = 'Namaste,\n\n' + loginId + ' ne ek submission edit karne ki request bheji hai.\n\n' +
    'Submission ID: ' + submissionId + '\n' +
    'Report Type: ' + reportType + '\n' +
    'Reason: ' + (reason || '(not given)') + '\n\n' +
    'Admin Panel me jaakar "Edit Requests" section se approve/reject karein.';
  try { GmailApp.sendEmail(recipients.join(','), subject, body); } catch (e) { /* non-fatal */ }

  return { ok: true, requestId: requestId };
}

function getEditRequests() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EDIT_REQUESTS);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  return data.map(function (d, i) {
    return { row: i + 2, requestId: d[0], submissionId: d[1], reportType: d[2],
      requestedBy: d[3], reason: d[4], status: d[5], requestedAt: formatDate_(d[6]) };
  }).reverse();
}

function respondEditRequest(row, approve) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EDIT_REQUESTS);
  if (!sh) throw new Error('Edit Requests sheet not found.');
  var status = approve ? 'Approved' : 'Rejected';
  sh.getRange(row, 6).setValue(status);
  var requestedBy = sh.getRange(row, 4).getValue();
  var submissionId = sh.getRange(row, 2).getValue();
  if (requestedBy && requestedBy.indexOf('@') > -1) {
    try {
      GmailApp.sendEmail(requestedBy, '[Edit Request] ' + submissionId + ' — ' + status,
        'Aapki edit request ' + status.toLowerCase() + ' ho gayi hai. Submission ID: ' + submissionId);
    } catch (e) { /* non-fatal */ }
  }
  return { ok: true };
}

// Checks if this logged-in person has an APPROVED (and not yet used) edit
// request waiting — shown as a banner right when they log in, so they
// don't have to go hunting for it.
function getApprovedEditForUser(loginId) {
  if (!loginId) return null;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EDIT_REQUESTS);
  if (!sh || sh.getLastRow() < 2) return null;
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][3] === loginId && data[i][5] === 'Approved') {
      return { requestRow: i + 2, submissionId: data[i][1], reportType: data[i][2] };
    }
  }
  return null;
}

// Marks an edit request as consumed once the officer has actually gone
// through and resubmitted — prevents the same approval being reused
// endlessly.
function markEditRequestUsed_(row) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EDIT_REQUESTS);
  if (sh) sh.getRange(row, 6).setValue('Edited');
}

// Fetches a past submission's full data (header fields + every
// indicator's Score/Observation/Suggestion/Timeline) so the form can be
// re-opened pre-filled for editing.
function getSubmissionForEdit(submissionId, reportType) {
  var sheetName = reportType === 'CFL' ? SHEET_CFL_SUB : SHEET_SESSION_SUB;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) throw new Error('Submission sheet not found.');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idCol = headers.indexOf('Submission ID');
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  for (var r = 0; r < data.length; r++) {
    if (data[r][idCol] === submissionId) {
      var row = data[r];
      var col = function (name) { var i = headers.indexOf(name); return i > -1 ? row[i] : ''; };
      var items = reportType === 'CFL' ? CFL_ITEMS : SESSION_ITEMS;
      var itemsData = items.map(function (it) {
        return {
          sr: it.sr,
          score: col(itemColName_(it, 'Score')),
          observation: col(itemColName_(it, 'Observation')),
          suggestion: col(itemColName_(it, 'Suggestion')),
          timeline: col(itemColName_(it, 'Timeline Date')) ? formatDateForInput_(col(itemColName_(it, 'Timeline Date'))) : ''
        };
      });
      return {
        submissionId: submissionId,
        cflName: col('CFL Name'), bccCode: col('BCC Code'),
        visitDate: col('Visit Date') ? formatDateForInput_(col('Visit Date')) : '',
        officerName: col('Officer Name'),
        consultantDesignation: col('Consultant Staff Designation') || '',
        sessionObserver: col('Session Observer / Stakeholders') || '',
        sessionType: col('Type of Session') || '',
        sessionVenue: col('Venue of Session') || '',
        signOffName: col('Sign-off Name'),
        items: itemsData
      };
    }
  }
  throw new Error('Submission ' + submissionId + ' not found.');
}

function getBootstrapData() {
  return {
    cflList: CFL_MASTER_DATA.map(function (r) {
      return { name: r[0], bccCode: r[1], phase: r[2], state: r[3], district: r[4],
        baseBlock: r[5], adjBlock1: r[6], adjBlock2: r[7], bankName: r[8] };
    }),
    employeeList: getEmployeeMasterLive_(),
    cflItems: CFL_ITEMS,
    sessionItems: SESSION_ITEMS,
    cflMaxTotal: CFL_MAX_TOTAL,
    sessionMaxTotal: SESSION_MAX_TOTAL,
    tatLimitDays: TAT_LIMIT_DAYS
  };
}

/**
 * Finds the most recent PRIOR submission for this CFL + report type and
 * returns, for EVERY indicator, that visit's Observation / Suggestion /
 * Timeline Date — so the form's Follow-up TABLE can auto-populate one row
 * per question instead of the officer typing it again. Also used
 * server-side (submitReport) as the single source of truth for TAT calc.
 */
function buildFollowUpMap_(reportType, cflName) {
  var sheetName = reportType === 'CFL' ? SHEET_CFL_SUB : SHEET_SESSION_SUB;
  var items = reportType === 'CFL' ? CFL_ITEMS : SESSION_ITEMS;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  var result = { lastVisitDate: '', items: {} };
  items.forEach(function (it) { result.items[it.sr] = { observation: '', suggestion: '', timeline: '' }; });
  if (!sh || sh.getLastRow() < 2 || !cflName) return result;

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var colCfl = headers.indexOf('CFL Name');
  var colTs = headers.indexOf('Timestamp');
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  var best = null, bestTs = null;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][colCfl]) === String(cflName)) {
      var ts = new Date(data[i][colTs]);
      if (!bestTs || ts > bestTs) { bestTs = ts; best = data[i]; }
    }
  }
  if (!best) return result;

  var visitDateCol = headers.indexOf('Visit Date');
  result.lastVisitDate = best[visitDateCol] ? formatDate_(best[visitDateCol]) : '';

  items.forEach(function (it) {
    var obsCol = headers.indexOf(itemColName_(it, 'Observation'));
    var sugCol = headers.indexOf(itemColName_(it, 'Suggestion'));
    var tlCol = headers.indexOf(itemColName_(it, 'Timeline Date'));
    result.items[it.sr] = {
      observation: obsCol > -1 ? (best[obsCol] || '') : '',
      suggestion: sugCol > -1 ? (best[sugCol] || '') : '',
      timeline: (tlCol > -1 && best[tlCol]) ? formatDateForInput_(best[tlCol]) : ''
    };
  });
  return result;
}

function getFollowUpData(reportType, cflName) {
  return buildFollowUpMap_(reportType, cflName);
}

// ---------------------------------------------------------------------
// PHOTO UPLOAD
// ---------------------------------------------------------------------
function getExistingPhotoUrls_(sh, submissionId) {
  if (!submissionId) return '';
  try {
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var idCol = headers.indexOf('Submission ID');
    var urlCol = headers.indexOf('Photo URLs');
    if (idCol === -1 || urlCol === -1 || sh.getLastRow() < 2) return '';
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][idCol] === submissionId) return data[i][urlCol] || '';
    }
  } catch (e) { /* fall through */ }
  return '';
}

function uploadPhoto(base64Data, filename, mimeType, officerName) {
  var root, file;

  try {
    root = getOrCreateFolder_();
  } catch (e) { throw new Error('Root folder create/access fail hui: ' + e.message); }

  // NOTE: photos are saved directly in the root report folder (same place
  // PDF/Excel already save successfully) rather than in nested Photos/
  // <Officer>/<Month> subfolders. Creating NEW FOLDERS appears to be
  // blocked by a Drive policy on this account/organization even though
  // creating FILES is allowed — so instead of a folder structure, the
  // officer name and month are baked into the FILENAME itself, which
  // achieves the same "organize/search by officer + month" goal without
  // needing folder-creation permission at all.
  var monthLabel = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'MMM-yyyy'); // e.g. "Aug-2026"
  var safeName = (officerName ? String(officerName).trim() : 'Unknown-Officer').replace(/[^a-zA-Z0-9 _-]/g, '');
  var finalFilename = safeName + ' - ' + monthLabel + ' - ' + filename;

  try {
    var bytes = Utilities.base64Decode(base64Data.split(',').pop());
    var blob = Utilities.newBlob(bytes, mimeType, finalFilename);
    file = root.createFile(blob);
  } catch (e) { throw new Error('Photo file create fail hui: ' + e.message); }

  // Setting "Anyone with the link" sharing is a SEPARATE permission from
  // basic create/write access — many Google Workspace organizations block
  // external/link sharing via policy even when file creation itself is
  // allowed. Never let that policy block the whole upload — the file is
  // already saved at this point; sharing is just for convenient viewing.
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log('⚠️ Could not set "Anyone with link" sharing (org policy?) — file is still saved. ' + e.message);
  }

  return { url: file.getUrl(), directUrl: 'https://drive.google.com/uc?id=' + file.getId(), id: file.getId() };
}

// ---------------------------------------------------------------------
// SUBMIT
// ---------------------------------------------------------------------
function submitReport(payload) {
  // payload = {
  //   reportType: 'CFL'|'SESSION', cflName, bccCode, district, state, bankName, phase,
  //   baseBlock, adjBlock, visitDate, officerName, officerDesignation,
  //   consultantDesignation, sessionObserver, sessionType, sessionVenue,
  //   items: [{sr, score, observation, suggestion, timeline, completionDate}],
  //   photos: [{url, directUrl}], signOff: {name, designation, district, zone}
  // }
  var emails = RECIPIENT_EMAILS.slice(0, 5); // fixed, never comes from the client
  if (!payload) throw new Error('Invalid request — no data received.');

  var loginId = payload.loggedInEmail;
  if (!loginId) {
    throw new Error('Login zaroori hai. Email daal kar ya "bina email" option se login karein.');
  }
  if (loginId.indexOf('NOLOGIN:') === 0) {
    var exemptName = loginId.substring(8);
    var exemptCheck = checkLoginExempt(exemptName);
    if (!exemptCheck.ok) {
      throw new Error('Ye officer login-exempt nahi hai. Admin se apna email add karwayein.');
    }
  } else {
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(loginId)) {
      throw new Error('Login email zaroori hai. Sahi email daal kar dobara try karein.');
    }
    var emailCheck = verifyLoginEmail_(loginId);
    if (!emailCheck.ok) {
      throw new Error('Ye email Employee Master me registered nahi hai. Sahi email se login karein.');
    }
  }

  var isEdit = !!payload.editSubmissionId;
  if (!isEdit && (!payload.photos || payload.photos.length < 2)) {
    throw new Error('Kam se kam 2 photo upload karna zaroori hai.');
  }

  var items = payload.reportType === 'CFL' ? CFL_ITEMS : SESSION_ITEMS;
  // Score is mandatory — an empty/undefined score means the officer never
  // selected one (the client's dropdown starts on a disabled placeholder).
  // Observation/Suggestion remain optional.
  items.forEach(function (it) {
    var match = (payload.items || []).filter(function (p) { return Number(p.sr) === it.sr; })[0];
    if (!match || match.score === '' || match.score === null || typeof match.score === 'undefined') {
      throw new Error('Sr ' + it.sr + ' (' + it.head + ') ka Score select nahi kiya gaya — Score bharna zaroori hai.');
    }
  });
  var maxTotal = payload.reportType === 'CFL' ? CFL_MAX_TOTAL : SESSION_MAX_TOTAL;

  var totalScore = 0;
  payload.items.forEach(function (it) { totalScore += Number(it.score) || 0; });
  var pct = maxTotal ? Math.round((totalScore / maxTotal) * 1000) / 10 : 0;
  var grade = gradeFromPct_(pct);

  // Server is the single source of truth for "last month" values and TAT —
  // never trust these back from the client.
  var followUpMap = buildFollowUpMap_(payload.reportType, payload.cflName);
  var tatGreen = 0, tatRed = 0, tatPending = 0;
  var itemsWithTat = items.map(function (it) {
    var match = payload.items.filter(function (p) { return Number(p.sr) === it.sr; })[0] || {};
    var prior = followUpMap.items[it.sr] || { observation: '', suggestion: '', timeline: '' };
    var tat = computeTat_(prior.timeline, match.completionDate);
    if (tat.status === 'Green') tatGreen++;
    else if (tat.status === 'Red') tatRed++;
    else tatPending++;
    return {
      sr: it.sr, head: it.head, text: it.text, max: it.max,
      score: Number(match.score) || 0,
      observation: match.observation || '', suggestion: match.suggestion || '',
      timeline: match.timeline || '', completionDate: match.completionDate || '',
      staffingNote: match.staffingNote || '',
      lastObservation: prior.observation, lastSuggestion: prior.suggestion, lastTimeline: prior.timeline,
      tat: tat
    };
  });

  var submissionId = payload.editSubmissionId ||
    ((payload.reportType === 'CFL' ? 'CFL-' : 'SESS-') + Utilities.getUuid().slice(0, 8).toUpperCase());
  var timestamp = new Date();

  // --- Security/audit trail: when the form was opened vs submitted, and
  // where the photos / final submit happened — helps confirm the assigned
  // officer actually filled this in personally, not someone else via a
  // forwarded link.
  var openedAt = payload.formOpenedAt ? new Date(payload.formOpenedAt) : null;
  var fillDurationMin = openedAt ? Math.round(((timestamp - openedAt) / 60000) * 10) / 10 : '';

  // The web app is deployed with "Anyone with a Google account" access, so
  // whoever opens the link must already be signed into a Google account —
  // this captures WHICH account, for free, no OTP/SMS gateway needed.
  var loggedInAs = '';
  try { loggedInAs = Session.getActiveUser().getEmail() || 'Not detected'; }
  catch (e) { loggedInAs = 'Not detected'; }

  // --- 1. Save row to the submissions sheet ---
  var sheetName = payload.reportType === 'CFL' ? SHEET_CFL_SUB : SHEET_SESSION_SUB;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureSubmissionSheet_(ss, sheetName, items, payload.reportType);

  var vals = {
    'Submission ID': submissionId, 'Timestamp': timestamp, 'Report Type': payload.reportType,
    'CFL Name': payload.cflName, 'District': payload.district, 'State': payload.state,
    'BCC Code': payload.bccCode, 'Bank Name': payload.bankName, 'Phase': payload.phase,
    'Base Block': payload.baseBlock, 'Adjacent Block': payload.adjBlock, 'Visit Date': payload.visitDate,
    'Officer Name': payload.officerName, 'Officer Designation': payload.officerDesignation,
    'Consultant Staff Designation': payload.consultantDesignation || '',
    'Session Observer / Stakeholders': payload.sessionObserver || '',
    'Type of Session': payload.sessionType || '', 'Venue of Session': payload.sessionVenue || '',
    'Total Score': totalScore, 'Max Score': maxTotal, 'Score %': pct, 'Grade': grade,
    'Photo URLs': (payload.photos && payload.photos.length) ? payload.photos.map(function (p) { return p.url; }).join(', ') : getExistingPhotoUrls_(sh, submissionId),
    'TAT Green Count': tatGreen, 'TAT Red Count': tatRed, 'TAT Pending Count': tatPending,
    'Sign-off Name': payload.signOff.name, 'Sign-off Designation': payload.signOff.designation,
    'Sign-off District': payload.signOff.district, 'Sign-off Zone': payload.signOff.zone,
    'Recipient Emails': emails.join(', '), 'PDF Link': '', 'Excel Link': '',
    'Logged-in Google Account': loggedInAs,
    'Logged-in Email': (payload && payload.loggedInEmail) || '',
    'Form Opened At': openedAt || '', 'Fill Duration (min)': fillDurationMin,
    'Photo Location': formatLocation_(payload.photoLocation), 'Photo Location Map Link': mapLink_(payload.photoLocation),
    'Submit Location': formatLocation_(payload.submitLocation), 'Submit Location Map Link': mapLink_(payload.submitLocation)
  };
  itemsWithTat.forEach(function (it) {
    vals[itemColName_(it, 'Score')] = it.score;
    vals[itemColName_(it, 'Observation')] = it.observation;
    vals[itemColName_(it, 'Suggestion')] = it.suggestion;
    vals[itemColName_(it, 'Timeline Date')] = it.timeline;
    vals[itemColName_(it, 'Completion Date')] = it.completionDate;
    if (it.staffing) vals[itemColName_(it, 'Staffing Details')] = it.staffingNote || '';
  });

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var row = headers.map(function (h) { return vals.hasOwnProperty(h) ? vals[h] : ''; });

  // Multiple officers can submit at almost the same moment — a lock makes
  // sure only ONE write happens at a time, so two simultaneous submissions
  // can never clash or overwrite each other's row. Waits up to 30s for its
  // turn (invisible to the officer — just a brief delay).
  var lock = LockService.getScriptLock();
  var newRowIndex;
  try {
    lock.waitLock(30000);
    if (payload.editSubmissionId) {
      // EDIT MODE: overwrite the existing row for this Submission ID
      // instead of appending a new one.
      var idCol = headers.indexOf('Submission ID');
      var allIds = sh.getRange(2, idCol + 1, Math.max(1, sh.getLastRow() - 1), 1).getValues();
      var foundRow = -1;
      for (var i = 0; i < allIds.length; i++) {
        if (allIds[i][0] === payload.editSubmissionId) { foundRow = i + 2; break; }
      }
      if (foundRow === -1) throw new Error('Original submission ' + payload.editSubmissionId + ' not found — cannot edit.');
      sh.getRange(foundRow, 1, 1, row.length).setValues([row]);
      newRowIndex = foundRow;
    } else {
      sh.appendRow(row);
      newRowIndex = sh.getLastRow();
    }
  } finally {
    lock.releaseLock();
  }
  if (payload.editSubmissionId && payload.editRequestRow) {
    try { markEditRequestUsed_(payload.editRequestRow); } catch (e) { /* non-fatal */ }
  }

  // --- 2. Generate PDF + Excel ---
  var tatSummary = { green: tatGreen, red: tatRed, pending: tatPending };
  var audit = {
    loggedInAs: loggedInAs,
    loginEmail: payload.loggedInEmail,
    formOpenedAt: openedAt ? formatDate_(openedAt) + ' ' + Utilities.formatDate(openedAt, 'Asia/Kolkata', 'HH:mm') : 'N/A',
    fillDuration: fillDurationMin !== '' ? fillDurationMin + ' min' : 'N/A',
    photoLocation: formatLocation_(payload.photoLocation),
    submitLocation: formatLocation_(payload.submitLocation)
  };
  var reportData = buildReportData_(payload, itemsWithTat, maxTotal, totalScore, pct, grade,
    tatSummary, submissionId, timestamp, followUpMap.lastVisitDate, audit);
  var pdfBlob = generatePdf_(reportData);

  var folder;
  try {
    folder = getOrCreateFolder_();
  } catch (e) { throw new Error('Report folder create/access fail hui: ' + e.message); }

  var excelFile;
  try {
    excelFile = generateExcel_(reportData, folder); // created directly inside folder — no risky "move" step
  } catch (e) { throw new Error('Excel file create fail hui: ' + e.message); }

  var pdfFile;
  try {
    pdfFile = folder.createFile(pdfBlob);
  } catch (e) { throw new Error('PDF file create fail hui: ' + e.message); }

  // "Anyone with link" sharing is a separate Drive permission from basic
  // create/write access — some organizations block it via policy even
  // when file creation itself is allowed. Never let that block the whole
  // report — both files are already saved by this point.
  try { pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { Logger.log('⚠️ Could not share PDF link (org policy?) — file still saved. ' + e.message); }
  try { excelFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { Logger.log('⚠️ Could not share Excel link (org policy?) — file still saved. ' + e.message); }

  // write links back (by header name, not fixed position)
  var pdfCol = headers.indexOf('PDF Link');
  var xlsCol = headers.indexOf('Excel Link');
  if (pdfCol > -1) sh.getRange(newRowIndex, pdfCol + 1).setValue(pdfFile.getUrl());
  if (xlsCol > -1) sh.getRange(newRowIndex, xlsCol + 1).setValue(excelFile.getUrl());

  // --- 3. Email ---
  sendReportEmail_(emails, reportData, pdfBlob, excelFile.getBlob());

  return {
    submissionId: submissionId,
    totalScore: totalScore, maxTotal: maxTotal, pct: pct, grade: grade,
    tatSummary: tatSummary, pdfUrl: pdfFile.getUrl(), excelUrl: excelFile.getUrl(),
    previewHtml: renderPdfHtml_(reportData)
  };
}

function gradeFromPct_(pct) {
  if (pct >= 90) return 'A';
  if (pct >= 75) return 'B';
  if (pct >= 60) return 'C';
  return 'D';
}

function computeTat_(suggestionDateStr, completionDateStr) {
  if (!suggestionDateStr) {
    return { days: null, overDays: null, status: 'Pending' };
  }
  var d1 = new Date(suggestionDateStr);
  // If Completion Date isn't filled in yet, calculate LIVE against today —
  // so an overdue item's Red day-count keeps growing day by day until it's
  // actually completed, instead of just sitting as "Pending" forever.
  var d2 = completionDateStr ? new Date(completionDateStr) : new Date();
  var days = Math.round((d2 - d1) / (1000 * 60 * 60 * 24));

  if (days < 0) return { days: days, overDays: null, status: 'Pending' }; // timeline date is in the future
  if (days <= TAT_LIMIT_DAYS) return { days: days, overDays: 0, status: 'Green' };
  return { days: days, overDays: days - TAT_LIMIT_DAYS, status: 'Red' };
}

function formatDate_(v) {
  try {
    return Utilities.formatDate(new Date(v), 'Asia/Kolkata', 'dd-MMM-yyyy');
  } catch (e) { return String(v); }
}

function formatDateForInput_(v) {
  try {
    return Utilities.formatDate(new Date(v), 'Asia/Kolkata', 'yyyy-MM-dd');
  } catch (e) { return ''; }
}

// ---------------------------------------------------------------------
// REPORT DATA (shared by PDF + Excel)
// ---------------------------------------------------------------------
function buildReportData_(payload, itemsWithTat, maxTotal, totalScore, pct, grade, tatSummary, submissionId, timestamp, lastVisitDate, audit) {
  // Common header fields (both report types)
  var meta = [
    ['CFL Name', payload.cflName], ['District', payload.district], ['State', payload.state],
    ['BCC Code', payload.bccCode], ['Bank Name', payload.bankName], ['Phase', payload.phase],
    ['Base Block', payload.baseBlock], ['Adjacent Block', payload.adjBlock],
    ['Visit Date', payload.visitDate], ['Officer Name', payload.officerName],
    ['Designation', payload.officerDesignation]
  ];
  // Session-only fields — these must NOT appear in the CFL report.
  if (payload.reportType === 'SESSION') {
    meta = meta.concat([
      ['Designation of Consultant Staff', payload.consultantDesignation || '-'],
      ['Session Observer / Stakeholders', payload.sessionObserver || '-'],
      ['Type of Session', payload.sessionType || '-'],
      ['Venue of Session', payload.sessionVenue || '-']
    ]);
  }

  return {
    submissionId: submissionId,
    timestamp: formatDate_(timestamp),
    reportTitle: payload.reportType === 'CFL' ?
      'Checklist for CFL Visit - MoneyWise CFL Project' :
      'Checklist for Session Quality Monitoring - MoneyWise CFL Project',
    reportType: payload.reportType,
    cflName: payload.cflName,
    visitDate: payload.visitDate,
    lastVisitDate: lastVisitDate || '',
    audit: audit || {},
    meta: meta,
    items: itemsWithTat,
    totalScore: totalScore, maxTotal: maxTotal, pct: pct, grade: grade,
    tatSummary: tatSummary,
    signOff: payload.signOff,
    photos: payload.photos
  };
}

// ---------------------------------------------------------------------
// DASHBOARD — indicator-wise & score(compliance)-wise
// ---------------------------------------------------------------------
function getDashboardData(reportType, filters) {
  filters = filters || {};
  var sheetName = reportType === 'CFL' ? SHEET_CFL_SUB : SHEET_SESSION_SUB;
  var items = reportType === 'CFL' ? CFL_ITEMS : SESSION_ITEMS;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  var empty = {
    rows: [], indicatorStats: items.map(function (it) { return { sr: it.sr, text: shortText_(it.text), avgPct: 0, count: 0, max: it.max }; }),
    gradeTrend: [],
    filterOptions: { officers: [], cfls: [], districts: [] },
    summary: { totalVisits: 0, avgScorePct: 0, grades: { A: 0, B: 0, C: 0, D: 0 }, tat: { Green: 0, Red: 0, Pending: 0 } }
  };
  if (!sh || sh.getLastRow() < 2) return empty;

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = function (name) { return headers.indexOf(name); };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  // Dropdown options always reflect the FULL dataset (not the filtered
  // subset) so the officer/CFL/district lists never shrink after you pick one.
  var officerSet = {}, cflSet = {}, districtSet = {};
  data.forEach(function (d) {
    if (d[col('Officer Name')]) officerSet[d[col('Officer Name')]] = true;
    if (d[col('CFL Name')]) cflSet[d[col('CFL Name')]] = true;
    if (d[col('District')]) districtSet[d[col('District')]] = true;
  });
  var filterOptions = {
    officers: Object.keys(officerSet).sort(),
    cfls: Object.keys(cflSet).sort(),
    districts: Object.keys(districtSet).sort()
  };

  var indicatorSums = items.map(function (it) { return { sr: it.sr, text: shortText_(it.text), max: it.max, sum: 0, count: 0 }; });
  var rows = [];
  var grades = { A: 0, B: 0, C: 0, D: 0 };
  var tat = { Green: 0, Red: 0, Pending: 0 };
  var pctTotal = 0;
  var matchedCount = 0;

  for (var r = 0; r < data.length; r++) {
    var d = data[r];
    if (filters.officerName && d[col('Officer Name')] !== filters.officerName) continue;
    if (filters.cflName && d[col('CFL Name')] !== filters.cflName) continue;
    if (filters.district && d[col('District')] !== filters.district) continue;
    matchedCount++;

    var pct = Number(d[col('Score %')]) || 0;
    var grade = d[col('Grade')] || 'D';
    var rowGreen = Number(d[col('TAT Green Count')]) || 0;
    var rowRed = Number(d[col('TAT Red Count')]) || 0;
    var rowPending = Number(d[col('TAT Pending Count')]) || 0;

    pctTotal += pct;
    if (grades.hasOwnProperty(grade)) grades[grade]++;
    tat.Green += rowGreen; tat.Red += rowRed; tat.Pending += rowPending;

    items.forEach(function (it, i) {
      var scoreCol = col(itemColName_(it, 'Score'));
      if (scoreCol > -1) {
        indicatorSums[i].sum += Number(d[scoreCol]) || 0;
        indicatorSums[i].count++;
      }
    });

    rows.push({
      submissionId: d[col('Submission ID')],
      timestamp: formatDate_(d[col('Timestamp')]),
      cflName: d[col('CFL Name')],
      district: d[col('District')],
      visitDate: d[col('Visit Date')] ? formatDate_(d[col('Visit Date')]) : '',
      officerName: d[col('Officer Name')],
      totalScore: d[col('Total Score')],
      maxScore: d[col('Max Score')],
      pct: pct,
      grade: grade,
      tatGreen: rowGreen, tatRed: rowRed, tatPending: rowPending,
      pdfLink: d[col('PDF Link')],
      excelLink: d[col('Excel Link')]
    });
  }

  rows.sort(function (a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });

  // Grading trend — grade counts grouped by month (using Visit Date),
  // sorted chronologically, so the dashboard can show how grading has
  // moved over time. Respects the same filters as everything else.
  var trendMap = {};
  for (var t = 0; t < data.length; t++) {
    if (filters.officerName && data[t][col('Officer Name')] !== filters.officerName) continue;
    if (filters.cflName && data[t][col('CFL Name')] !== filters.cflName) continue;
    if (filters.district && data[t][col('District')] !== filters.district) continue;
    var visitDateVal = data[t][col('Visit Date')];
    if (!visitDateVal) continue;
    var monthKey = Utilities.formatDate(new Date(visitDateVal), 'Asia/Kolkata', 'yyyy-MM');
    var monthLabel = Utilities.formatDate(new Date(visitDateVal), 'Asia/Kolkata', 'MMM yyyy');
    if (!trendMap[monthKey]) trendMap[monthKey] = { key: monthKey, label: monthLabel, A: 0, B: 0, C: 0, D: 0 };
    var g = data[t][col('Grade')] || 'D';
    if (trendMap[monthKey].hasOwnProperty(g)) trendMap[monthKey][g]++;
  }
  var gradeTrend = Object.keys(trendMap).sort().map(function (k) { return trendMap[k]; });

  var indicatorStats = indicatorSums.map(function (s) {
    return { sr: s.sr, text: s.text, max: s.max, count: s.count,
      avgPct: s.count ? Math.round((s.sum / (s.count * s.max)) * 1000) / 10 : 0 };
  });

  return {
    rows: rows,
    indicatorStats: indicatorStats,
    gradeTrend: gradeTrend,
    filterOptions: filterOptions,
    summary: {
      totalVisits: matchedCount,
      avgScorePct: matchedCount ? Math.round((pctTotal / matchedCount) * 10) / 10 : 0,
      grades: grades,
      tat: tat
    }
  };
}

function shortText_(t) {
  var firstLine = (t.split('\n')[0] || t).trim();
  return firstLine.length > 60 ? firstLine.substring(0, 60) + '...' : firstLine;
}


function renderPdfHtml_(reportData) {
  var t = HtmlService.createTemplateFromFile('PdfTemplate');
  t.data = reportData;
  return t.evaluate().getContent();
}

function generatePdf_(reportData) {
  var html = renderPdfHtml_(reportData);
  var blob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
  blob.setName(reportData.reportType + '_' + reportData.submissionId + '.pdf');
  return blob;
}

// ---------------------------------------------------------------------
// EXCEL GENERATION (matches original CFL Reporting format / Session
// Monitoring Report sheet layout: Sr No | Title | Score | Observation)
// ---------------------------------------------------------------------
function generateExcel_(reportData, targetFolder) {
  var tempName = 'TEMP_' + reportData.submissionId;
  var tempSs = SpreadsheetApp.create(tempName);
  var sh = tempSs.getSheets()[0];
  sh.setName(reportData.reportType === 'CFL' ? 'CFL Reporting format' : 'Session Monitoring Report');

  var BLACK = '#000000';
  var HEADER_BLUE = '#0b5cab';
  var border = function (range) {
    range.setBorder(true, true, true, true, true, true, BLACK, SpreadsheetApp.BorderStyle.SOLID);
  };

  var row = 1;
  var titleRange = sh.getRange(row, 1, 1, 7).merge().setValue(reportData.reportTitle)
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center')
    .setBackground(HEADER_BLUE).setFontColor('#ffffff');
  border(titleRange);
  row += 2;

  reportData.meta.forEach(function (m) {
    var labelCell = sh.getRange(row, 1).setValue(m[0]).setFontWeight('bold').setBackground('#f0f2f4');
    var valCell = sh.getRange(row, 2, 1, 6).merge().setValue(m[1]);
    border(labelCell); border(valCell);
    row++;
  });
  row++;

  var headerRange = sh.getRange(row, 1, 1, 6);
  headerRange.setValues([['Sr. No', 'Titles', 'Score Obtained', 'Observation', 'Suggestion Recommended', 'Time Line Date']])
    .setFontWeight('bold').setBackground(HEADER_BLUE).setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  border(headerRange);
  row++;

  var itemsStartRow = row;
  reportData.items.forEach(function (it) {
    sh.getRange(row, 1).setValue(it.sr).setHorizontalAlignment('center');
    var headText = it.head + '\n' + it.text;
    var rich = SpreadsheetApp.newRichTextValue()
      .setText(headText)
      .setTextStyle(0, it.head.length, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#0b5cab').build())
      .build();
    sh.getRange(row, 2).setRichTextValue(rich).setWrap(true);
    sh.getRange(row, 3).setValue(it.score + ' / ' + it.max).setHorizontalAlignment('center').setFontWeight('bold');
    var obsCombined = it.observation || '';
    if (it.staffingNote) obsCombined += (obsCombined ? '\n\n' : '') + 'Staffing —\n' + it.staffingNote;
    sh.getRange(row, 4).setValue(obsCombined).setWrap(true);
    sh.getRange(row, 5).setValue(it.suggestion || '').setWrap(true);
    sh.getRange(row, 6).setValue(it.timeline || '').setHorizontalAlignment('center');
    row++;
  });
  border(sh.getRange(itemsStartRow, 1, reportData.items.length, 6));
  row++;

  var summaryStartRow = row;
  sh.getRange(row, 1).setValue('Total Score').setFontWeight('bold');
  sh.getRange(row, 2).setValue(reportData.totalScore + ' / ' + reportData.maxTotal);
  row++;
  sh.getRange(row, 1).setValue('Score %').setFontWeight('bold');
  sh.getRange(row, 2).setValue(reportData.pct + '%');
  row++;
  sh.getRange(row, 1).setValue('Grade').setFontWeight('bold');
  sh.getRange(row, 2).setValue(reportData.grade);
  border(sh.getRange(summaryStartRow, 1, 3, 2));
  row += 2;

  // ---- Follow-up table: one row per indicator ----
  var fuTitleRange = sh.getRange(row, 1, 1, 7).merge()
    .setValue('Follow-up on Previous Visit Observations & Stakeholder Feedback Status' +
      (reportData.lastVisitDate ? ' (Last Visit: ' + reportData.lastVisitDate + ')' : ' (Pehli visit)'))
    .setFontWeight('bold').setBackground(HEADER_BLUE).setFontColor('#ffffff');
  border(fuTitleRange);
  row++;
  var fuHeaderRange = sh.getRange(row, 1, 1, 7);
  fuHeaderRange.setValues([['Sr. No', 'Titles', 'Last Month Observation', 'Last Month Suggestion',
    'Last Month Time Line Date', 'Completion Date', 'TAT']])
    .setFontWeight('bold').setBackground('#f0f2f4').setHorizontalAlignment('center');
  border(fuHeaderRange);
  row++;
  var fuStartRow = row;
  var fuItems = reportData.items.filter(function (it) { return it.lastObservation || it.lastSuggestion; });
  if (fuItems.length === 0) {
    sh.getRange(row, 1, 1, 7).merge().setValue('Pichli visit me koi Observation/Suggestion darj nahi tha.')
      .setFontColor('#6e7781').setHorizontalAlignment('center');
    row++;
  } else {
    fuItems.forEach(function (it) {
      sh.getRange(row, 1).setValue(it.sr).setHorizontalAlignment('center');
      sh.getRange(row, 2).setValue(it.head).setWrap(true);
      sh.getRange(row, 3).setValue(it.lastObservation || '-').setWrap(true);
      sh.getRange(row, 4).setValue(it.lastSuggestion || '-').setWrap(true);
      sh.getRange(row, 5).setValue(it.lastTimeline || '-').setHorizontalAlignment('center');
      sh.getRange(row, 6).setValue(it.completionDate || '-').setHorizontalAlignment('center');
      var tatCell = sh.getRange(row, 7).setValue(it.tat.status).setHorizontalAlignment('center').setFontWeight('bold');
      if (it.tat.status === 'Green') tatCell.setFontColor('#1a7f37');
      else if (it.tat.status === 'Red') tatCell.setFontColor('#cf222e');
      else tatCell.setFontColor('#6e7781');
      row++;
    });
  }
  border(sh.getRange(fuStartRow, 1, Math.max(fuItems.length, 1), 7));
  row += 2;

  var signOffStartRow = row;
  sh.getRange(row, 1).setValue('Sign-off Name').setFontWeight('bold');
  sh.getRange(row, 2).setValue(reportData.signOff.name);
  row++;
  sh.getRange(row, 1).setValue('Designation').setFontWeight('bold');
  sh.getRange(row, 2).setValue(reportData.signOff.designation);
  row++;
  sh.getRange(row, 1).setValue('District').setFontWeight('bold');
  sh.getRange(row, 2).setValue(reportData.signOff.district);
  row++;
  sh.getRange(row, 1).setValue('Zone').setFontWeight('bold');
  sh.getRange(row, 2).setValue(reportData.signOff.zone);
  row++;
  sh.getRange(row, 1).setValue('Digital Signature').setFontWeight('bold');
  sh.getRange(row, 2).setValue(reportData.signOff.name + ' · digitally submitted on ' +
    reportData.timestamp + ' (Submission ID ' + reportData.submissionId + ')');
  border(sh.getRange(signOffStartRow, 1, 5, 2));

  sh.setColumnWidth(1, 45);
  sh.setColumnWidth(2, 260);
  sh.setColumnWidth(3, 80);
  sh.setColumnWidth(4, 200);
  sh.setColumnWidth(5, 170);
  sh.setColumnWidth(6, 100);
  sh.setColumnWidth(7, 80);

  SpreadsheetApp.flush();
  var tempId = tempSs.getId();
  var token = ScriptApp.getOAuthToken();
  var url = 'https://docs.google.com/spreadsheets/d/' + tempId + '/export?format=xlsx';
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  var blob = resp.getBlob().setName(reportData.reportType + '_' + reportData.submissionId + '.xlsx');
  var excelFile = targetFolder ? targetFolder.createFile(blob) : DriveApp.createFile(blob);

  // Cleaning up the temp Google Sheet is nice-to-have, not critical — if
  // deleting is blocked by org policy, don't let that fail the whole report.
  try {
    DriveApp.getFileById(tempId).setTrashed(true);
  } catch (e) {
    Logger.log('⚠️ Could not trash temp sheet (org policy?) — harmless, ignoring. ' + e.message);
  }
  return excelFile;
}

// ---------------------------------------------------------------------
// EMAIL
// ---------------------------------------------------------------------
function sendReportEmail_(emails, reportData, pdfBlob, excelBlob) {
  var district = '';
  reportData.meta.forEach(function (m) { if (m[0] === 'District') district = m[1]; });

  var subject = '[' + reportData.reportType + ' Monitoring] ' + reportData.cflName +
    ' - ' + reportData.visitDate + ' - Score ' + reportData.pct + '% (' + reportData.grade + ')';
  var body = 'Namaste,\n\n' +
    (reportData.reportType === 'CFL' ? 'CFL Monitoring Visit' : 'Session Quality Monitoring') +
    ' report attached hai.\n\n' +
    'CFL: ' + reportData.cflName + '\n' +
    'District: ' + district + '\n' +
    'Visit Date: ' + reportData.visitDate + '\n' +
    'Score: ' + reportData.totalScore + '/' + reportData.maxTotal + ' (' + reportData.pct + '%) - Grade ' + reportData.grade + '\n' +
    'Follow-up TAT: 🟢 ' + reportData.tatSummary.green + ' Green · 🔴 ' + reportData.tatSummary.red +
    ' Red · ⏳ ' + reportData.tatSummary.pending + ' Pending\n\n' +
    'PDF aur Excel dono attach hain. Dhanyawad.';

  GmailApp.sendEmail(emails.join(','), subject, body, {
    attachments: [pdfBlob, excelBlob],
    name: 'MoneyWise CFL Monitoring System'
  });
}
