/* AUTOMATION TRACKER */
var TRACKER_SHEET_ID = '17Tp-e3e1i0_nw52_qxGdkrMRNv9_nSeejfOr2fwRQrs';
var TRACKER_TAB_NAME = 'Automation Tracker';

var AUTOMATION_REGISTRY = {

  'Step 1: Prepare File Copies': {
    'Owner(s)': 'PMG Technology',
    'OOH Process Category': 'RFP',
    'Manual Process Replaced': 'Manually creating Drive folder structures and copying media plan templates for each vendor/market combination',
    'Process Stage': 'Deployed',
    'Implementation Date': '2026-03-18',
    'Time Saved (hrs/week)': 2,
    'Impact Type': 'Time Savings',
    'Impact Notes': 'Automates folder creation, template copying, SuperGrid setup, and vendor link tracking'
  },

  'Step 2: Send Vendor Emails': {
    'Owner(s)': 'PMG Technology',
    'OOH Process Category': 'RFP',
    'Manual Process Replaced': 'Manually composing and sending individual RFP emails to each vendor with correct attachments and folder links',
    'Process Stage': 'Deployed',
    'Implementation Date': '2026-03-18',
    'Time Saved (hrs/week)': 3,
    'Impact Type': 'Time Savings',
    'Impact Notes': 'Automates email template rendering, token replacement, Drive sharing permissions, and send tracking'
  }

};

// Helper: find tab by trimmed name (handles trailing spaces)
function findTrackerTab_(ss) {
  var tab = ss.getSheetByName(TRACKER_TAB_NAME);
  if (tab) return tab;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getName().trim() === TRACKER_TAB_NAME.trim()) return all[i];
  }
  return null;
}

function authorizeTrackerAccess() {
  try {
    var ext = SpreadsheetApp.openById(TRACKER_SHEET_ID);
    var tab = findTrackerTab_(ext);
    if (tab) { SpreadsheetApp.getActive().toast('Access granted. Tab found with ' + tab.getLastRow() + ' rows.', 'Tracker Auth', 6); }
    else { SpreadsheetApp.getActive().toast('Connected but tab not found.', 'Tracker Auth', 6); }
  } catch (e) { SpreadsheetApp.getActive().toast('Auth failed: ' + e.message, 'Tracker Auth Error', 10); }
}

function logToAutomationTracker_(automationName, overrides) {
  if (!automationName) { logLine_('TRACKER: skipped'); return; }
  try {
    var reg = AUTOMATION_REGISTRY[automationName] || {};
    var meta = {};
    for (var k in reg) { if (reg.hasOwnProperty(k)) meta[k] = reg[k]; }
    if (overrides) { for (var ok in overrides) { if (overrides.hasOwnProperty(ok)) meta[ok] = overrides[ok]; } }
    meta['Automation Name'] = automationName;
    var weekly = parseFloat(meta['Time Saved (hrs/week)']);
    if (!isNaN(weekly) && !meta['Annual Time Saved']) { meta['Annual Time Saved'] = Math.round(weekly * 52 * 10) / 10; }
    var extSs = SpreadsheetApp.openById(TRACKER_SHEET_ID);
    var extTab = findTrackerTab_(extSs);
    if (!extTab) { logLine_('TRACKER ERROR: Tab not found'); return; }
    var lastCol = extTab.getLastColumn();
    if (lastCol < 1) lastCol = 11;
    var extHeaders = extTab.getRange(1, 1, 1, lastCol).getValues()[0];
    var rowVals = [];
    for (var c2 = 0; c2 < extHeaders.length; c2++) {
      var header = String(extHeaders[c2]).trim();
      rowVals.push(meta[header] != null ? meta[header] : '');
    }
    var lastRow = extTab.getLastRow();
    var existingRow = -1;
    var searchName = automationName.trim().toLowerCase();
    if (lastRow >= 2) {
      var nameCol = extTab.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var r = 0; r < nameCol.length; r++) {
        if (String(nameCol[r][0]).trim().toLowerCase() === searchName) { existingRow = r + 2; break; }
      }
    }
    if (existingRow > 0) {
      extTab.getRange(existingRow, 1, 1, rowVals.length).setValues([rowVals]);
      logLine_('TRACKER: Updated row ' + existingRow + ' for ' + automationName);
    } else {
      // Find first empty row in column A (skip header)
      var colA = extTab.getRange(2, 1, extTab.getMaxRows() - 1, 1).getValues();
      var insertRow = -1;
      for (var ri = 0; ri < colA.length; ri++) {
        if (!colA[ri][0] || String(colA[ri][0]).trim() === '') { insertRow = ri + 2; break; }
      }
      if (insertRow > 0) {
        extTab.getRange(insertRow, 1, 1, rowVals.length).setValues([rowVals]);
        logLine_('TRACKER: Wrote new row ' + insertRow + ' for ' + automationName);
      } else {
        extTab.appendRow(rowVals);
        logLine_('TRACKER: Appended row for ' + automationName);
      }
    }
  } catch (e) { logLine_('TRACKER ERROR: ' + e.message); }
}
