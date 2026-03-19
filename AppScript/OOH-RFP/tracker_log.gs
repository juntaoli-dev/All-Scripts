/******************************************************
 * AUTOMATION TRACKER – Logs automation runs to the
 * central "Automation Tracker" tab in the OOH Process
 * tracking sheet.
 *
 * How it works:
 * 1. Reads automation metadata from the local
 *    "Automation Registry" tab (gid 1954082249) in
 *    THIS spreadsheet.
 * 2. Upserts a row into the external Automation Tracker
 *    sheet (by Automation Name).
 * 3. Called at the end of each LIVE automation run.
 ******************************************************/

/* ============ CONFIG ================================ */
var TRACKER_SHEET_ID   = '17Tp-e3e1i0_nw52_qxGdkrMRNv9_nSeejfOr2fwRQrs';
var TRACKER_TAB_NAME   = 'Automation Tracker';
var REGISTRY_GID       = 1954082249;   // gid of the local metadata tab
/* ==================================================== */

/**
 * Expected headers in Automation Tracker (target):
 *   Automation Name | Owner(s) | OOH Process Category |
 *   Manual Process Replaced | Process Stage |
 *   Implementation Date | Time Saved (hrs/week) |
 *   Annual Time Saved | Impact Type | Impact Notes |
 *   Links to Resources
 */
var TRACKER_HEADERS = [
  'Automation Name',
  'Owner(s)',
  'OOH Process Category',
  'Manual Process Replaced',
  'Process Stage',
  'Implementation Date',
  'Time Saved (hrs/week)',
  'Annual Time Saved',
  'Impact Type',
  'Impact Notes',
  'Links to Resources'
];

/* ============ PUBLIC API ============================ */

/**
 * Call this at the end of any LIVE automation run.
 *
 * @param {string} automationName  – Must match the
 *        "Automation Name" value in the local registry tab.
 *        e.g. "Step 1: Prepare File Copies"
 * @param {Object} [overrides]  – Optional key/value pairs
 *        to override or supplement registry data.
 *        Keys should match TRACKER_HEADERS (case-insensitive).
 */
function logToAutomationTracker_(automationName, overrides) {
  if (!automationName) {
    logLine_('TRACKER: skipped – no automationName provided');
    return;
  }

  try {
    // 1. Read metadata from local registry tab
    var meta = getRegistryRow_(automationName);
    if (!meta) {
      logLine_('TRACKER WARNING: No registry row found for "' + automationName + '". Logging with defaults.');
      meta = {};
    }

    // Apply overrides
    if (overrides) {
      for (var k in overrides) {
        if (overrides.hasOwnProperty(k)) {
          meta[normalizeKey_(k)] = overrides[k];
        }
      }
    }

    // Ensure Automation Name is set
    meta[normalizeKey_('Automation Name')] = automationName;

    // Auto-calculate Annual Time Saved if not set
    var weeklyKey = normalizeKey_('Time Saved (hrs/week)');
    var annualKey = normalizeKey_('Annual Time Saved');
    if (meta[weeklyKey] && !meta[annualKey]) {
      var weekly = parseFloat(meta[weeklyKey]);
      if (!isNaN(weekly)) {
        meta[annualKey] = Math.round(weekly * 52 * 10) / 10; // 1 decimal
      }
    }

    // 2. Open external tracker sheet
    var extSs  = SpreadsheetApp.openById(TRACKER_SHEET_ID);
    var extTab  = extSs.getSheetByName(TRACKER_TAB_NAME);
    if (!extTab) {
      logLine_('TRACKER ERROR: Tab "' + TRACKER_TAB_NAME + '" not found in external sheet.');
      return;
    }

    // 3. Read existing headers from tracker (row 1)
    var extHeaders = extTab.getRange(1, 1, 1, extTab.getLastColumn()).getValues()[0];
    var extIdx = {};
    for (var h = 0; h < extHeaders.length; h++) {
      var nk = normalizeKey_(extHeaders[h]);
      if (nk) extIdx[nk] = h;
    }

    // 4. Build the row values array aligned to external headers
    var rowVals = [];
    for (var c = 0; c < extHeaders.length; c++) {
      var colKey = normalizeKey_(extHeaders[c]);
      rowVals.push(meta[colKey] || '');
    }

    // 5. Upsert: find existing row by Automation Name (column A assumed)
    var nameColIdx = extIdx[normalizeKey_('Automation Name')];
    if (nameColIdx == null) nameColIdx = 0; // default col A

    var lastRow = extTab.getLastRow();
    var existingRow = -1;

    if (lastRow >= 2) {
      var nameCol = extTab.getRange(2, nameColIdx + 1, lastRow - 1, 1).getValues();
      for (var r = 0; r < nameCol.length; r++) {
        if (String(nameCol[r][0]).trim().toLowerCase() === automationName.trim().toLowerCase()) {
          existingRow = r + 2; // 1-indexed, skip header
          break;
        }
      }
    }

    if (existingRow > 0) {
      // UPDATE existing row
      extTab.getRange(existingRow, 1, 1, rowVals.length).setValues([rowVals]);
      logLine_('TRACKER: Updated row ' + existingRow + ' for "' + automationName + '"');
    } else {
      // APPEND new row
      extTab.appendRow(rowVals);
      logLine_('TRACKER: Appended new row for "' + automationName + '"');
    }

  } catch (e) {
    logLine_('TRACKER ERROR: ' + e.message);
  }
}

/* ============ REGISTRY READER ====================== */

/**
 * Finds the local tab with gid = REGISTRY_GID and
 * returns a flat {normalizedKey: value} object for
 * the row whose "Automation Name" matches.
 */
function getRegistryRow_(automationName) {
  var ss = SpreadsheetApp.getActive();
  var regTab = findSheetByGid_(ss, REGISTRY_GID);

  if (!regTab) {
    // Fallback: try common tab names
    var fallbackNames = ['Automation Registry', 'Registry', 'Automations'];
    for (var i = 0; i < fallbackNames.length; i++) {
      regTab = ss.getSheetByName(fallbackNames[i]);
      if (regTab) break;
    }
  }

  if (!regTab) return null;

  var lastRow = regTab.getLastRow();
  var lastCol = regTab.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;

  var headers = regTab.getRange(1, 1, 1, lastCol).getValues()[0];
  var data    = regTab.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Build header index (normalized)
  var hIdx = {};
  for (var h = 0; h < headers.length; h++) {
    var nk = normalizeKey_(headers[h]);
    if (nk) hIdx[nk] = h;
  }

  // Find matching row
  var nameKey = normalizeKey_('Automation Name');
  var nameCol = hIdx[nameKey];
  if (nameCol == null) {
    // Try first column as name
    nameCol = 0;
  }

  for (var r = 0; r < data.length; r++) {
    var rowName = String(data[r][nameCol] || '').trim().toLowerCase();
    if (rowName === automationName.trim().toLowerCase()) {
      // Build result object
      var result = {};
      for (var key in hIdx) {
        if (hIdx.hasOwnProperty(key)) {
          var val = data[r][hIdx[key]];
          result[key] = (val != null && val !== '') ? String(val).trim() : '';
        }
      }
      return result;
    }
  }

  return null;
}

/**
 * Finds a sheet by its gid (getSheetId()).
 */
function findSheetByGid_(ss, gid) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  return null;
}

/* ============ HELPERS =============================== */

/**
 * Normalize a header string for matching:
 * lowercase, strip non-alphanumeric, collapse spaces.
 */
function normalizeKey_(s) {
  if (!s) return '';
  return String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
}
