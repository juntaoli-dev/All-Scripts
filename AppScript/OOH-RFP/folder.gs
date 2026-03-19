/******************************************************
 * RFP Automation – STEP 1 (Vendor Links-only, with Debug)
 * Reads "Holiday RFP List" → copies files into Parent→Market→Vendor
 * Writes ONLY to "Vendor Links" (no edits to Holiday RFP List)
 * Includes: Debug scanner, Automation Log, Preview that always writes rows
 * Now reads PARENT_FOLDER_ID from "Inputs" tab (Cell B2)
 *
 * UPDATE (SuperGrid):
 * - Reads SuperGrid template Sheet ID from Inputs!B7
 * - Copies ONE SuperGrid sheet into the Parent folder (not Market/Vendor)
 * - Includes de-dup/overwrite behavior using OVERWRITE_EXISTING
 * - Writes Campaign Name to G4 and Parent Folder ID to G5
 ******************************************************/

/* =============== CONFIG ============================ */
// PARENT_FOLDER_ID is now dynamic (from Inputs tab)
var SHARE_ANYONE_VIEW = false;   // true → new copies “Anyone with link: Viewer”
var OVERWRITE_EXISTING = true;   // true → replace old links / replace existing SuperGrid in Parent
var SOURCE_TAB = 'RFP List';
var TARGET_TAB = 'Vendor Links';
var INPUTS_TAB_NAME = 'Inputs';
var PARENT_FOLDER_CELL = 'B2';
var CAMPAIGN_NAME_CELL = 'B6';

// NEW: SuperGrid template lives in Inputs!B7
var SUPERGRID_TEMPLATE_CELL = 'B7';

// Naming for the copied SuperGrid file in Parent
var SUPERGRID_NAME_SUFFIX = ' | SuperGrid Template';
/* =================================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RFP Automation (Market - Vendor)')
    .addItem('Step 1: Prepare Copies (Preview)', 'prepareCopiesPreview')
    .addItem('Step 1: Prepare Copies (Live)', 'prepareCopiesLive')
    .addSeparator()
    // ▼▼ Step 2 buttons (defined elsewhere in your project) ▼▼
    .addItem('Step 2: Send Deck + Plan (Preview)', 'sendDeckAndPlanPreview')
    .addItem('Step 2: Send Deck + Plan (Live)', 'sendDeckAndPlanLive')
    // ▲▲ ▲▲
    .addSeparator()
    .addItem('Debug: Scan Source', 'debugScanSource_')
    .addToUi();
}

function prepareCopiesPreview() { prepareCopies_({ live:false }); }
function prepareCopiesLive()    { prepareCopies_({ live:true  }); }

function prepareCopies_(opts) {
  var live = !!(opts && opts.live);
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName(SOURCE_TAB);

  if (!src) { toast_('Source tab "' + SOURCE_TAB + '" not found.'); logLine_('ERROR: source not found'); return; }

  var lastRow = src.getLastRow();
  if (lastRow <= 1) { toast_('No data rows in "' + SOURCE_TAB + '".'); logLine_('INFO: no data rows'); return; }

  var headers = src.getRange(1,1,1,src.getLastColumn()).getValues()[0];
  var idx = headerIndexMap_(headers);

  // must have market + vendor
  var missing = [];
  ['market','vendor'].forEach(function(h){ if (!(h in idx) && !headerLooseHas_(idx,h)) missing.push(h); });
  if (missing.length) { toast_('Missing headers: ' + missing.join(', ')); logLine_('ERROR: missing ' + missing.join(',')); return; }

  var data = src.getRange(2,1,lastRow-1,src.getLastColumn()).getValues();
  if (!data || !data.length) { toast_('Read 0 rows from source.'); logLine_('ERROR: empty data read'); return; }

  // Parent folder (only needed in LIVE)
  var parent = null;
  var parentId = null; // Stored at this scope so we can write it to SuperGrid
  var campaignName = '';
  
  if (live) {
    try {
      parentId = getParentFolderIdFromInputs_();
      if (!parentId) {
        toast_('Parent Folder ID missing in ' + INPUTS_TAB_NAME + '!' + PARENT_FOLDER_CELL);
        logLine_('ERROR: Parent Folder ID missing in Inputs tab');
        return;
      }
      parent = DriveApp.getFolderById(parentId);
    }
    catch (e) { toast_('Parent folder not accessible.'); logLine_('ERROR: parent folder ' + e.message); return; }

    campaignName = getCampaignNameFromInputs_();
    if (!campaignName) {
      toast_('Campaign Name missing in ' + INPUTS_TAB_NAME + '!' + CAMPAIGN_NAME_CELL);
      logLine_('ERROR: Campaign Name missing in Inputs tab');
      return;
    }

    // =========================
    // NEW: SuperGrid copy ONCE into Parent (LIVE only)
    // =========================
    try {
      var superGridTemplateIdOrUrl = getSuperGridTemplateIdFromInputs_();
      if (!superGridTemplateIdOrUrl) {
        logLine_('INFO: SuperGrid template ID missing in Inputs!' + SUPERGRID_TEMPLATE_CELL);
      } else {
        var superGridName = safeName_(campaignName + SUPERGRID_NAME_SUFFIX);

        // De-dupe / overwrite behavior
        var existing = findFileByNameInFolder_(parent, superGridName);
        if (existing) {
          if (OVERWRITE_EXISTING) {
            try {
              existing.setTrashed(true);
              logLine_('INFO: Trashed existing SuperGrid in Parent: ' + superGridName);
            } catch (eTrash) {
              logLine_('ERROR: Could not trash existing SuperGrid (' + superGridName + '): ' + eTrash.message);
            }
          } else {
            logLine_('INFO: SuperGrid already exists in Parent (skipping because OVERWRITE_EXISTING=false): ' + existing.getUrl());
            // If you want to skip copying entirely, just do nothing here.
            // We intentionally do NOT copy another in this case.
          }
        }

        // Copy if none exists OR overwrite is enabled
        if (!existing || OVERWRITE_EXISTING) {
          var superGridUrl = copySpreadsheetTemplateToFolder_(superGridTemplateIdOrUrl, parent, superGridName);
          
          // NEW: Write Campaign Name to G4 and Parent Folder ID to G5
          try {
            var copiedSg = SpreadsheetApp.openByUrl(superGridUrl);
            var sgFirstTab = copiedSg.getSheets()[0];
            sgFirstTab.getRange('G4').setValue(campaignName);
            sgFirstTab.getRange('G5').setValue(parentId);
          } catch (eWrite) {
            logLine_('WARNING: SuperGrid copied, but failed to write G4/G5 data: ' + eWrite.message);
          }

          logLine_('SUCCESS: SuperGrid created in Parent: ' + superGridUrl);
        }
      }
    } catch (eSG) {
      logLine_('ERROR: SuperGrid copy failed: ' + eSG.message);
    }
  } else {
    // PREVIEW mode: no Drive ops
    var sgPrev = getSuperGridTemplateIdFromInputs_();
    if (sgPrev) logLine_('PREVIEW: SuperGrid would be copied into Parent (Inputs!' + SUPERGRID_TEMPLATE_CELL + ')');
    else logLine_('PREVIEW: SuperGrid template missing (Inputs!' + SUPERGRID_TEMPLATE_CELL + ')');
  }

  var tz = Session.getScriptTimeZone();
  var nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var processed = 0;

  // Build records for Vendor Links
  var linkMap = {}; // key: market|vendor  -> { market, vendor, deck, plan, updatedAt, notes }

  for (var i = 0; i < data.length; i++) {
    var row = data[i];

    var market  = val(row, idx, 'market');
    var vendor  = val(row, idx, 'vendor');
    var deckSrc  = val(row, idx, ['deck template link','deck link','deck']);

    // Accept "RFP Link" as Media Plan source too
    var planSrc  = val(row, idx, ['media plan template link','media plan link','plan link','media plan','plan','rfp link']);
    var planCameFromRfp = (!val(row, idx, ['media plan template link','media plan link','plan link','media plan','plan']) &&
                           !!val(row, idx, 'rfp link'));

    var copiedDeck = '';
    var copiedPlan = '';
    var notes = [];

    if (!market || !vendor) {
      notes.push('Missing Market or Vendor');
      linkMap[(market + '|' + vendor).toLowerCase()] = {
        market: market, vendor: vendor,
        deck: live ? '' : '(preview) missing Market/Vendor',
        plan: live ? '' : '(preview) missing Market/Vendor',
        updatedAt: live ? nowStr : '',
        notes: notes.join(' | ')
      };
      continue;
    }

    try {
      if (live) {
        // Create/locate destination folders only in LIVE
        var marketFolder = getOrCreateSubFolder_(parent, market);
        var vendorFolderName = safeName_(vendor + ' - ' + campaignName);
        var vendorFolder = getOrCreateSubFolder_(marketFolder, vendorFolderName);
        var specsFolder = getOrCreateSubFolder_(vendorFolder, campaignName + ' Specs');
        var photosheetsFolder = getOrCreateSubFolder_(vendorFolder, campaignName + ' Photosheets');

        // ---- Deck (DISABLED) ----
        copiedDeck = '';                  // always blank in Vendor Links
        // notes.push('Deck copy disabled'); // optional

        // ---- Plan (LIVE) ----
        if (planSrc) {
          var planId = parseGoogleFileId_(planSrc);
          if (planId) {
            try {
              var pFile = DriveApp.getFileById(planId);
              var pCopy = pFile.makeCopy(safeName_(campaignName + ' | ' + vendor + ' | ' + market + ' | Grid'), vendorFolder);
              if (SHARE_ANYONE_VIEW) setAnyoneWithLinkView_(pCopy);
              copiedPlan = 'https://drive.google.com/file/d/' + pCopy.getId() + '/view';
              if (planCameFromRfp) notes.push('Plan source = RFP Link');
            } catch (eP) { notes.push('Plan copy error: ' + eP.message); }
          } else { notes.push('Plan src not Drive URL'); }
        } else { notes.push('No plan src'); }

      } else {
        // ---- PREVIEW: never touch Drive ----
        if (planSrc) {
          copiedPlan = '(preview) ' + planSrc;
          if (planCameFromRfp) notes.push('Plan source = RFP Link (preview)');
        } else {
          copiedPlan = '(preview) no plan src';
        }
      }

    } catch (eFolder) {
      notes.push('Folder error: ' + eFolder.message);
    }

    processed++;

    var key = (market + '|' + vendor).toLowerCase();
    linkMap[key] = {
      market: market,
      vendor: vendor,
      vendorFolderLink: live ? vendorFolder.getUrl() : '(preview) vendor folder link',
      deck: copiedDeck,
      plan: copiedPlan,
      updatedAt: live ? nowStr : '',
      notes: notes.join(' | ')
    };
  }

  // Write to Vendor Links
  upsertVendorLinks_(linkMap, live);

  toast_((live ? 'LIVE' : 'PREVIEW') + ' processed rows: ' + processed);
  logLine_((live ? 'LIVE' : 'PREVIEW') + ' processed=' + processed + ', vendors=' + Object.keys(linkMap).length);

  // Log to external Automation Tracker (LIVE only)
  if (live) {
    logToAutomationTracker_('Step 1: Prepare File Copies', {
      'Links to Resources': ss.getUrl()
    });
  }
}

/* -------- Vendor Links write-back -------- */
function upsertVendorLinks_(map, live){
  var ss=SpreadsheetApp.getActive();
  var tab=ss.getSheetByName(TARGET_TAB);
  if(!tab) tab=ss.insertSheet(TARGET_TAB);

  var headers=['Market','Vendor', 'Vendor Folder Link','Copied Media Plan Link','Last Updated','Notes'];
  ensureExactHeaders_(tab,headers);

  var lastRow=tab.getLastRow();
  var idx={};
  if(lastRow>=2){
    var vals=tab.getRange(2,1,lastRow-1,headers.length).getValues();
    for(var i=0;i<vals.length;i++){
      idx[(String(vals[i][0])+'|'+String(vals[i][1])).toLowerCase()]=2+i;
    }
  }

  var count=0;
  Object.keys(map).forEach(function(key){
    var r=map[key];
    // FIXED: Removed r.deck||'' so the array matches the 6 columns defined in headers.
    var rowVals=[[r.market||'', r.vendor||'', r.vendorFolderLink||'', r.plan||'', r.updatedAt||'', r.notes||'']];
    if(idx[key] && OVERWRITE_EXISTING){
      tab.getRange(idx[key],1,1,rowVals[0].length).setValues(rowVals);
      count++;
    } else if(!idx[key]) {
      tab.appendRow(rowVals[0]);
      count++;
    }
  });

  logLine_('Vendor Links upserted: ' + count + ' rows (live=' + live + ').');
}

/* ---------------- DEBUG ------------------ */
function debugScanSource_(){
  var ss=SpreadsheetApp.getActive();
  var src=ss.getSheetByName(SOURCE_TAB);
  if(!src){ toast_('Debug: source tab "'+SOURCE_TAB+'" not found.'); return; }
  var lastRow=src.getLastRow(), lastCol=src.getLastColumn();
  var headers= lastCol ? src.getRange(1,1,1,lastCol).getValues()[0] : [];
  var rows = (lastRow>1) ? src.getRange(2,1,Math.min(10,lastRow-1),lastCol).getValues() : [];

  var dbg=ss.getSheetByName('Debug Snapshot');
  if(!dbg) dbg=ss.insertSheet('Debug Snapshot'); else dbg.clear();

  dbg.getRange(1,1,1,4).setValues([['Timestamp','LastRow','LastCol','Headers']]);
  dbg.getRange(2,1,1,4).setValues([[new Date(), lastRow, lastCol, JSON.stringify(headers)]]);
  dbg.getRange(4,1,1,1).setValue('First up to 10 data rows:');
  if(rows.length){
    dbg.getRange(5,1,rows.length,rows[0].length).setValues(rows);
  } else {
    dbg.getRange(5,1,1,1).setValue('(no data rows)');
  }

  logLine_('DEBUG: lastRow='+lastRow+', lastCol='+lastCol+', headers='+JSON.stringify(headers));
  toast_('Debug complete. Check "Debug Snapshot" sheet.');
}

/* --------------- Helpers ------------------ */

// Fetch Parent Folder ID from "Inputs" tab (fixed cell)
function getParentFolderIdFromInputs_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(INPUTS_TAB_NAME);
  if (!sheet) return null;
  var v = sheet.getRange(PARENT_FOLDER_CELL).getValue();
  return v ? String(v).trim() : null;
}

// Fetch Campaign Name from "Inputs" tab (fixed cell)
function getCampaignNameFromInputs_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(INPUTS_TAB_NAME);
  if (!sheet) return null;
  var v = sheet.getRange(CAMPAIGN_NAME_CELL).getValue();
  return v ? String(v).trim() : null;
}

// NEW: Fetch SuperGrid template Sheet ID/URL from "Inputs" tab (fixed cell B7)
function getSuperGridTemplateIdFromInputs_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(INPUTS_TAB_NAME);
  if (!sheet) return null;
  var v = sheet.getRange(SUPERGRID_TEMPLATE_CELL).getValue();
  return v ? String(v).trim() : null;
}

/**
 * Copies a Spreadsheet template into a folder and returns the new Spreadsheet URL.
 * Accepts either an ID or a Drive URL.
 */
function copySpreadsheetTemplateToFolder_(templateIdOrUrl, destinationFolder, newName) {
  var templateId = parseGoogleFileId_(templateIdOrUrl) || String(templateIdOrUrl || '').trim();
  if (!templateId) throw new Error('Template ID is blank/invalid.');

  var templateFile = DriveApp.getFileById(templateId);
  var copy = templateFile.makeCopy(safeName_(newName || templateFile.getName()), destinationFolder);

  if (SHARE_ANYONE_VIEW) setAnyoneWithLinkView_(copy);

  // Return spreadsheet URL (nicer than /file/d/ for Sheets)
  return SpreadsheetApp.openById(copy.getId()).getUrl();
}

/**
 * Finds the first file with a given name in a folder.
 */
function findFileByNameInFolder_(folder, name) {
  var it = folder.getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}

function headerIndexMap_(h){var m={};for(var i=0;i<h.length;i++){var n=String(h[i]||'').toLowerCase().trim();if(n)m[n]=i;}return m;}
function headerLooseHas_(idx,k){
  k=String(k||'').toLowerCase().replace(/\s+/g,'');
  for (var key in idx){ if (key.replace(/\s+/g,'')===k) return true; }
  return false;
}
function val(row,idx,keys){
  if(!row) return '';
  if(typeof keys==='string') keys=[keys];
  if(!keys||!keys.length) return '';
  for(var i=0;i<keys.length;i++){
    var k=String(keys[i]||'').toLowerCase();
    if(k in idx){ var v=row[idx[k]]; if(v!==''&&v!=null) return String(v).trim(); }
    for(var key in idx){
      if(key.replace(/\s+/g,'')===k.replace(/\s+/g,'')){
        var v2=row[idx[key]]; if(v2!==''&&v2!=null) return String(v2).trim();
      }
    }
  }
  return '';
}
function getOrCreateSubFolder_(p,n){var it=p.getFoldersByName(n);return it.hasNext()?it.next():p.createFolder(n);}
function parseGoogleFileId_(u){
  if(!u) return null;
  var s = String(u);
  var m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/); if(m) return m[1];
  var m2 = s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/); if(m2) return m2[1];
  // If user pasted a raw ID:
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s.trim())) return s.trim();
  return null;
}
function safeName_(s){return (s||'file').replace(/[\/\\:*?"<>]+/g,'_').slice(0,120);}
function setAnyoneWithLinkView_(f){try{f.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);}catch(e){}}
function ensureExactHeaders_(tab,h){
  if(tab.getLastRow()===0){tab.getRange(1,1,1,h.length).setValues([h]);return;}
  var cur=tab.getRange(1,1,1,h.length).getValues()[0];
  var mismatch=false;
  for(var i=0;i<h.length;i++){if(String(cur[i]||'')!==h[i]){mismatch=true;break;}}
  if(mismatch){tab.clear();tab.getRange(1,1,1,h.length).setValues([h]);}
}
function toast_(msg){SpreadsheetApp.getActive().toast(msg,'RFP Automation',6);}

/**
 * Writes a line to the "Automation Log" sheet.
 */
function logLine_(msg){
  var ss = SpreadsheetApp.getActive();
  var t = ss.getSheetByName('Automation Log');
  if(!t) t = ss.insertSheet('Automation Log');

  // Ensure headers
  if (t.getLastRow() === 0) t.appendRow(['Timestamp', 'Message']);

  t.appendRow([new Date(), String(msg || '')]);
}