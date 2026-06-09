/*** ============================================================
 *  Code.gs  —  Google Apps Script (server-side)
 *  Unit Entry Tracker -> Google Sheet updater
 *
 *  DEPLOYMENT STEPS
 *  1. Go to https://script.google.com  ->  New project.
 *  2. Paste this file into Code.gs.
 *  3. Add an HTML file named exactly "Index" (File > New > HTML file)
 *     and paste the front-end HTML (second code block) into it.
 *  4. (Recommended) Store the Sheet ID in Script Properties instead
 *     of source:  Project Settings > Script Properties >
 *        add  SHEET_ID = 1ksOJ5B_NCHhkWFqV4n-1AsIWQ3KSOp5xA3FBBMteidk
 *  5. Deploy > New deployment > type "Web app".
 *        Execute as: Me
 *        Who has access: <your organization>  (NOT "Anyone" if data is sensitive)
 *  6. Authorize when prompted, copy the /exec URL, open it in a browser.
 * ============================================================ ***/

// --- Configuration ---------------------------------------------------------
// Prefer Script Properties; fall back to constant only if not set.
function getSheetId_() {
  var fromProps = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return fromProps || '1ksOJ5B_NCHhkWFqV4n-1AsIWQ3KSOp5xA3FBBMteidk';
}

var SHEET_TAB_NAME = ''; // '' = use the first/active sheet. Set a name to target a specific tab.

// Allow-lists (server-side enforcement; never trust the client) --------------
var ALLOWED_UNITS = ['HHT BDE', '6-9CAV', '2-7CAV', '215 BSB', '3 BEB', '3-8', '1-12CAV', '2-82FA'];
var ALLOWED_SHIPS = ['At port', 'Green Bay', 'Green Ocean', 'ARC Integrity', 'Boat 4'];

// Header names exactly as they appear in row 1 of the sheet ------------------
var HDR_UNIT   = 'UNIT';            // expected column A
var HDR_BUMPER = 'bumper no space'; // expected column M
var HDR_STATUS = 'status';          // expected column J
var HDR_TIME   = 'time';            // expected column K

// --- Web app entry point ---------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Unit Entry Tracker')
}

/**
 * Called from the client via google.script.run.
 * @param {Object} payload {unit, bumper, ship, timestamp}
 * @return {Object} {ok:boolean, status:'UPDATED'|'NO_MATCH'|'INVALID'|'ERROR', message:string}
 */
function recordEntry(payload) {
  var lock = LockService.getScriptLock();
  try {
    // Serialize writes to avoid race conditions / double updates.
    lock.waitLock(15000);

    // ---- 1. Validate & sanitize input (defense in depth) ----
    if (!payload || typeof payload !== 'object') {
      return result_(false, 'INVALID', 'Malformed request.');
    }
    var unit   = sanitize_(payload.unit);
    var bumper = sanitize_(payload.bumper);
    var ship   = sanitize_(payload.ship);
    var ts     = sanitize_(payload.timestamp);

    if (ALLOWED_UNITS.indexOf(unit) === -1) {
      return result_(false, 'INVALID', 'Unrecognized unit.');
    }
    if (ALLOWED_SHIPS.indexOf(ship) === -1) {
      return result_(false, 'INVALID', 'Unrecognized ship.');
    }
    if (!/^[A-Za-z0-9]+$/.test(bumper)) {
      return result_(false, 'INVALID', 'Bumper must be letters/numbers only.');
    }
    // Validate ISO-8601 timestamp; if bad, regenerate server-side.
    if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(ts)) {
      ts = new Date().toISOString();
    }

    // ---- 2. Open the sheet ----
    var ss = SpreadsheetApp.openById(getSheetId_());
    var sheet = SHEET_TAB_NAME ? ss.getSheetByName(SHEET_TAB_NAME) : ss.getSheets()[0];
    if (!sheet) {
      return result_(false, 'ERROR', 'Target sheet/tab not found.');
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) {
      return result_(false, 'NO_MATCH', 'Sheet has no data rows.');
    }

    // ---- 3. Resolve columns by HEADER NAME (safer than fixed letters) ----
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var cUnit   = indexOfHeader_(headers, HDR_UNIT);
    var cBumper = indexOfHeader_(headers, HDR_BUMPER);
    var cStatus = indexOfHeader_(headers, HDR_STATUS);
    var cTime   = indexOfHeader_(headers, HDR_TIME);

    var missing = [];
    if (cUnit   < 0) missing.push(HDR_UNIT);
    if (cBumper < 0) missing.push(HDR_BUMPER);
    if (cStatus < 0) missing.push(HDR_STATUS);
    if (cTime   < 0) missing.push(HDR_TIME);
    if (missing.length) {
      return result_(false, 'ERROR', 'Missing header column(s): ' + missing.join(', '));
    }

    // ---- 4. Find the matching row (UNIT + bumper no space) ----
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var wantUnit   = unit.toLowerCase();
    var wantBumper = bumper.toLowerCase();
    var foundRow = -1;

    for (var i = 0; i < data.length; i++) {
      var rowUnit   = String(data[i][cUnit]   == null ? '' : data[i][cUnit]).trim().toLowerCase();
      var rowBumper = String(data[i][cBumper] == null ? '' : data[i][cBumper]).trim().toLowerCase();
      if (rowUnit === wantUnit && rowBumper === wantBumper) {
        foundRow = i + 2; // +2: array is 0-based and starts at sheet row 2
        break;
      }
    }

    if (foundRow === -1) {
      return result_(false, 'NO_MATCH',
        'No row where UNIT="' + unit + '" and bumper="' + bumper + '".');
    }

    // ---- 5. Write status (col J) and time (col K) ----
    // Plain strings only -> avoids any formula injection into the sheet.
    sheet.getRange(foundRow, cStatus + 1).setValue("'" /*force text*/ ? ship : ship);
    sheet.getRange(foundRow, cStatus + 1).setValue(ship);
    sheet.getRange(foundRow, cTime + 1).setValue(ts);
    SpreadsheetApp.flush();

    return result_(true, 'UPDATED',
      'Row ' + foundRow + ' updated: status="' + ship + '", time="' + ts + '".');

  } catch (err) {
    // Log full detail server-side; return a generic message to the client.
    console.error('recordEntry error: ' + (err && err.stack ? err.stack : err));
    return result_(false, 'ERROR', 'Server error while updating the sheet.');
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// --- Helpers ---------------------------------------------------------------
function sanitize_(v) {
  // Coerce to string, trim, strip control chars.
  return String(v == null ? '' : v).replace(/[\u0000-\u001F\u007F]/g, '').trim();
}

function indexOfHeader_(headers, name) {
  var target = String(name).trim().toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === target) return i;
  }
  return -1;
}

function result_(ok, status, message) {
  return { ok: !!ok, status: status, message: message };
}
