/**
 * ============================================================
 *  ICAN SCHOOL — STAFF ATTENDANCE BACKEND  (v4.0)
 *  Can be hosted either on GitHub Pages (index.html calls this
 *  script via POST) OR served directly by this Apps Script project
 *  — in that case the HTML file in this project MUST be named
 *  "index" (i.e. index.html) so doGet() can find it.
 * ============================================================
 *
 *  SETUP — do these 3 things:
 *
 *  1) SHEET_ID: paste your spreadsheet ID between the quotes.
 *     It is the long part of the sheet URL:
 *     https://docs.google.com/spreadsheets/d/ <THIS PART> /edit
 *     (If this script lives INSIDE the sheet via Extensions → Apps
 *     Script, you can leave it blank and it will still work.)
 *
 *  2) SHEET_NAME / STAFF_SHEET: must match the tab names at the bottom
 *     of the spreadsheet EXACTLY (capitals matter). SHEET_NAME is the
 *     punch-log tab; STAFF_SHEET is the roster tab (names in column A,
 *     one per row, section-header rows may be bold and are skipped).
 *
 *  3) Deploy: Deploy → New deployment → Web app
 *        Execute as:  Me
 *        Who has access:  Anyone        ← must be "Anyone"
 *     Copy the /exec URL — open it directly in a browser to load the
 *     kiosk (if hosting the HTML here), or paste it in the app's
 *     Settings screen (if hosting index.html on GitHub Pages instead).
 *
 *  IMPORTANT — every time you edit this file you must publish a
 *  NEW VERSION:  Deploy → Manage deployments → ✏️ → Version:
 *  "New version" → Deploy.  Saving alone does NOT update the URL.
 *
 *  Test in a browser tab:   <your /exec URL>?action=debug
 * ============================================================
 */

var SHEET_ID    = '1uryQqA5WvUPwargh_luUcn_Uo3Lzo9yNS76c4S-VSbw';             // ← paste sheet ID here (or leave blank if bound)
var SHEET_NAME  = 'Attendance';   // ← punch log tab name, exact match
var STAFF_SHEET = 'Staff';       // ← roster tab name, exact match (column A = names)
var ADMIN_PIN   = '2789';         // ← dashboard / settings PIN

// Sheet columns:
// A Timestamp | B Date | C Guard | D Staff Name | E Action | F Time | G Signature

// ---------------------------------------------------------- sheet access
function getSpreadsheet() {
  var ss = SHEET_ID
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet not found — paste your SHEET_ID at the top of Code.gs');
  return ss;
}

function getSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Tab "' + SHEET_NAME + '" not found — check SHEET_NAME in Code.gs');
  return sheet;
}

function getStaffSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(STAFF_SHEET);
  if (!sheet) throw new Error('Tab "' + STAFF_SHEET + '" not found — check STAFF_SHEET in Code.gs');
  return sheet;
}

// ---------------------------------------------------------- entry points
function doPost(e) {
  var out;
  try {
    var req    = JSON.parse(e.postData.contents);
    var fn     = req.function;
    var p      = req.parameters || [];

    switch (fn) {
      case 'getKioskData':    out = getKioskData();                       break;
      case 'syncPunches':     out = syncPunches(p[0]);                    break;
      case 'getPunchesByDate':out = getPunchesByDate(p[0]);               break;
      case 'verifyPin':       out = { ok: p[0] === ADMIN_PIN };           break;
      case 'getDashboard':    out = getDashboard(p[0], p[1]);             break;
      case 'updatePunchTime': out = updatePunchTime(p[0], p[1], p[2]);    break;
      case 'deletePunch':     out = deletePunch(p[0], p[1]);              break;
      case 'saveSettings':    out = saveSettings(p[0], p[1]);             break;
      default:                out = { ok: false, error: 'Unknown function: ' + fn };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'debug') {
      var sheet = getSheet();
      var staffSheet = getStaffSheet();
      var kiosk = getKioskData();
      return jsonOut({
        ok: true,
        sheetIdConfigured: !!SHEET_ID,
        tabName: SHEET_NAME,
        tabFound: true,
        totalRows: sheet.getLastRow(),
        staffTabName: STAFF_SHEET,
        staffTabRows: staffSheet.getLastRow(),
        staffCount: kiosk.staff.length,
        staffSample: kiosk.staff.slice(0, 5),
        settings: kiosk.settings
      });
    }
    // Default: serve the kiosk web app itself (index.html must be an
    // HTML file in this Apps Script project named "index").
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ICAN Attendance')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------- settings
var DEFAULT_SETTINGS = {
  lateTime:  '08:35',   // an "In" after this counts as late
  earlyTime: '16:00',   // an "Out" before this counts as early leave
  kioskTitle: 'ICAN Attendance',
  kioskSubtitle: 'Tap your name to sign in or out'
};

function getSettings() {
  var raw = PropertiesService.getScriptProperties().getProperty('ICAN_SETTINGS');
  var stored = raw ? JSON.parse(raw) : {};
  var merged = {};
  for (var k in DEFAULT_SETTINGS) merged[k] = stored[k] || DEFAULT_SETTINGS[k];
  return merged;
}

function saveSettings(pin, settings) {
  if (pin !== ADMIN_PIN) return { ok: false, error: 'Invalid PIN' };
  var clean = {};
  for (var k in DEFAULT_SETTINGS) {
    if (settings && settings[k] != null && settings[k] !== '') clean[k] = String(settings[k]);
  }
  PropertiesService.getScriptProperties().setProperty('ICAN_SETTINGS', JSON.stringify(clean));
  return { ok: true, settings: getSettings() };
}

// ---------------------------------------------------------- kiosk data
// Reads names from column A of the STAFF_SHEET roster tab, skipping the
// header row, blank rows, and bold section-header rows (e.g. "Administration").
function getKioskData() {
  var sheet = getStaffSheet();
  var last  = sheet.getLastRow();
  var staff = [];
  if (last >= 2) {
    var range   = sheet.getRange(2, 1, last - 1, 1);
    var values  = range.getValues();
    var weights = range.getFontWeights();
    var seen = {};
    for (var i = 0; i < values.length; i++) {
      var name = String(values[i][0] || '').trim();
      if (!name || weights[i][0] === 'bold') continue;
      if (!seen[name]) { seen[name] = true; staff.push(name); }
    }
    staff.sort(function (a, b) { return a.localeCompare(b); });
  }
  return { ok: true, staff: staff, settings: getSettings() };
}

// ---------------------------------------------------------- write punches
function syncPunches(punches) {
  if (!Array.isArray(punches) || punches.length === 0) return { ok: true, added: 0 };
  var sheet = getSheet();
  var rows = punches.map(function (p) {
    return [
      p.timestamp ? new Date(p.timestamp) : new Date(),
      p.date  || dateKey(new Date()),
      p.guard || 'KIOSK',
      String(p.name || ''),
      p.action === 'Out' ? 'Out' : 'In',
      String(p.time || timeKey(new Date())),
      ''
    ];
  });
  // single batched write — faster and atomic
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  return { ok: true, added: rows.length };
}

// ---------------------------------------------------------- read punches
function readAll(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 7).getValues();
}

function getPunchesByDate(dateStr) {
  var sheet  = getSheet();
  var target = normDate(dateStr) || dateKey(new Date());
  var out = [];
  readAll(sheet).forEach(function (row, i) {
    var name = String(row[3] || '').trim();
    if (!name) return;
    if (normDate(row[1]) !== target) return;
    out.push(punchFromRow(row, i + 2));
  });
  return { ok: true, date: target, punches: out };
}

function getDashboard(pin, dateStr) {
  if (pin !== ADMIN_PIN) return { ok: false, error: 'Invalid PIN' };
  var res = getPunchesByDate(dateStr);
  res.settings = getSettings();
  res.staff = getKioskData().staff;
  return res;
}

function punchFromRow(row, rowNumber) {
  var ts = row[0];
  return {
    row: rowNumber,
    timestamp: (ts instanceof Date) ? ts.getTime() : null,
    date:   normDate(row[1]),
    guard:  String(row[2] || ''),
    name:   String(row[3] || '').trim(),
    action: String(row[4] || 'In'),
    time:   normTime(row[5])
  };
}

// ---------------------------------------------------------- edit / delete
function updatePunchTime(pin, rowNumber, newTime) {
  if (pin !== ADMIN_PIN) return { ok: false, error: 'Invalid PIN' };
  if (!/^\d{1,2}:\d{2}(:\d{2})?$/.test(String(newTime))) {
    return { ok: false, error: 'Time must look like 08:30' };
  }
  var sheet = getSheet();
  var r = Number(rowNumber);
  if (!r || r < 2 || r > sheet.getLastRow()) return { ok: false, error: 'Row not found' };
  var t = String(newTime).split(':');
  var hh = ('0' + t[0]).slice(-2), mm = t[1], ss = t[2] || '00';
  sheet.getRange(r, 6).setNumberFormat('@').setValue(hh + ':' + mm + ':' + ss);
  return { ok: true };
}

function deletePunch(pin, rowNumber) {
  if (pin !== ADMIN_PIN) return { ok: false, error: 'Invalid PIN' };
  var sheet = getSheet();
  var r = Number(rowNumber);
  if (!r || r < 2 || r > sheet.getLastRow()) return { ok: false, error: 'Row not found' };
  sheet.deleteRow(r);
  return { ok: true };
}

// ---------------------------------------------------------- date/time utils
// Accepts a Date object, "yyyy-MM-dd", or "dd/MM/yyyy" — returns "yyyy-MM-dd".
function normDate(v) {
  if (v instanceof Date) return dateKey(v);
  var s = String(v || '').trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad(m[2]) + '-' + pad(m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + pad(m[2]) + '-' + pad(m[1]);
  return s || null;
}

// Accepts a Date object or "H:mm[:ss]" text — returns "HH:mm:ss".
function normTime(v) {
  if (v instanceof Date) return timeKey(v);
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (m) return pad(m[1]) + ':' + m[2] + ':' + (m[3] || '00');
  return s;
}

function dateKey(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
function timeKey(d) { return pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds()); }
function pad(n) { return ('0' + n).slice(-2); }