/* AUTOMATION LOG — appends a row to the external "Automation Log" tab
   Columns: Campaign Name | Date | Time | Number of RFPs Sent              */

var LOG_SHEET_ID  = '1liw1F0anynQzz-0aUpuqYKQjwxS9ScZ5hupc32_Yblo';
var LOG_TAB_NAME  = 'OOH Automation Log';

// Helper: find tab by trimmed name (handles trailing spaces)
function findLogTab_(ss) {
  var tab = ss.getSheetByName(LOG_TAB_NAME);
  if (tab) return tab;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getName().trim() === LOG_TAB_NAME.trim()) return all[i];
  }
  return null;
}

function authorizeLogAccess() {
  try {
    var ext = SpreadsheetApp.openById(LOG_SHEET_ID);
    var tab = findLogTab_(ext);
    if (tab) { SpreadsheetApp.getActive().toast('Access granted. Tab found with ' + tab.getLastRow() + ' rows.', 'Log Auth', 6); }
    else { SpreadsheetApp.getActive().toast('Connected but tab "' + LOG_TAB_NAME + '" not found.', 'Log Auth', 6); }
  } catch (e) { SpreadsheetApp.getActive().toast('Auth failed: ' + e.message, 'Log Auth Error', 10); }
}

/**
 * Append one row to the Automation Log.
 * @param {string} campaignName  Value from Inputs!B6
 * @param {number} rfpCount      Number of data rows in the RFP List tab
 */
function logToAutomationLog_(campaignName, rfpCount, automationType) {
  if (!campaignName) { logLine_('LOG: skipped — no campaign name'); return; }
  try {
    var now  = new Date();
    var date = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var time = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    var file_name = SpreadsheetApp.getActive().getName();
    var file_url = SpreadsheetApp.getActive().getUrl();
    var extSs  = SpreadsheetApp.openById(LOG_SHEET_ID);
    var extTab  = findLogTab_(extSs);
    if (!extTab) { logLine_('LOG ERROR: Tab "' + LOG_TAB_NAME + '" not found'); return; }

    var emailSender = (automationType && automationType.toLowerCase().indexOf('send') !== -1) ? Session.getActiveUser().getEmail() : '';
    var row = [date, time, '', campaignName, automationType, rfpCount, emailSender];

    // Find first empty row in column A (skip header row 1)
    var colA = extTab.getRange(2, 1, extTab.getMaxRows() - 1, 1).getValues();
    var insertRow = -1;
    for (var i = 0; i < colA.length; i++) {
      if (!colA[i][0] || String(colA[i][0]).trim() === '') { insertRow = i + 2; break; }
    }
if (insertRow > 0) {
  extTab.getRange(insertRow, 1, 1, row.length).setValues([row]);
  extTab.getRange(insertRow, 3).setFormula('=HYPERLINK("' + file_url + '","' + file_name + '")');
  logLine_('LOG: Wrote row ' + insertRow + ' — ' + campaignName);
} else {
  extTab.appendRow(row);
  var lastRow = extTab.getLastRow();
  extTab.getRange(lastRow, 3).setFormula('=HYPERLINK("' + file_url + '","' + file_name + '")');
  logLine_('LOG: Appended row for ' + campaignName);
}
  } catch (e) { logLine_('LOG ERROR: ' + e.message); }
}


