/******************************************************
 * RFP Automation – STEP 1 (Partner Links-only, with Debug)
 * Reads "Holiday RFP List" → copies files into Parent→Channel→Partner
 * Writes ONLY to "Partner Links" (no edits to Holiday RFP List)
 * Includes: Debug scanner, Automation Log, Preview that always writes rows
 * Now reads PARENT_FOLDER_ID from "Inputs" tab (Cell B2)
 ******************************************************/

/* =============== CONFIG ============================ */
// PARENT_FOLDER_ID is now dynamic (from Inputs tab)
var SHARE_ANYONE_VIEW = false;   // true → new copies “Anyone with link: Viewer”
var OVERWRITE_EXISTING = true;   // true → replace old links in Partner Links
var SOURCE_TAB = 'RFP List';
var TARGET_TAB = 'Partner Links';
var INPUTS_TAB_NAME = 'Inputs';
var PARENT_FOLDER_CELL = 'B2';
/* =================================================== */

function onOpen() {
  // Auto-capture this spreadsheet's ID so the daily trigger knows which sheet to open
  try { saveSpreadsheetId_(); } catch(e) {}

  SpreadsheetApp.getUi()
    .createMenu('RFP Automation')
    .addItem('Step 1: Prepare Copies (Preview)', 'prepareCopiesPreview')
    .addItem('Step 1: Prepare Copies (Live)', 'prepareCopiesLive')
    .addSeparator()
    // ▼▼ Step 2 buttons ▼▼
    .addItem('Step 2: Send Deck + Plan (Preview)', 'sendDeckAndPlanPreview')
    .addItem('Step 2: Send Deck + Plan (Live)', 'sendDeckAndPlanLive')
    .addSeparator()
    // ▼▼ Step 3: Archive to Progo Partner Hub ▼▼
    .addItem('Step 3: Archive to Hub (Preview)', 'archiveToHubPreview')
    .addItem('Step 3: Archive to Hub (Live)', 'archiveToHubLive')
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

  // must have channel + partner
  var missing = [];
  ['channel','partner'].forEach(function(h){ if (!(h in idx) && !headerLooseHas_(idx,h)) missing.push(h); });
  if (missing.length) { toast_('Missing headers: ' + missing.join(', ')); logLine_('ERROR: missing ' + missing.join(',')); return; }

  var data = src.getRange(2,1,lastRow-1,src.getLastColumn()).getValues();
  if (!data || !data.length) { toast_('Read 0 rows from source.'); logLine_('ERROR: empty data read'); return; }

  // Parent folder (only needed in LIVE)
  var parent = null;
  if (live) {
    try { 
      var parentId = getParentFolderIdFromInputs_();
      if (!parentId) {
        toast_('Parent Folder ID missing in ' + INPUTS_TAB_NAME + '!' + PARENT_FOLDER_CELL);
        logLine_('ERROR: Parent Folder ID missing in Inputs tab');
        return;
      }
      parent = DriveApp.getFolderById(parentId); 
    }
    catch (e) { toast_('Parent folder not accessible.'); logLine_('ERROR: parent folder ' + e.message); return; }
  }

  var tz = Session.getScriptTimeZone();
  var nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var processed = 0;

  // Build records for Partner Links
  var linkMap = {}; // key: channel|partner  -> { channel, partner, deck, plan, updatedAt, notes }

  for (var i = 0; i < data.length; i++) {
    var row = data[i];

    var channel  = val(row, idx, 'channel');
    var partner  = val(row, idx, 'partner');
    var deckSrc  = val(row, idx, ['deck template link','deck link','deck']);

    // Accept "RFP Link" as Media Plan source too
    var planSrc  = val(row, idx, ['media plan template link','media plan link','plan link','media plan','plan','rfp link']);
    var planCameFromRfp = (!val(row, idx, ['media plan template link','media plan link','plan link','media plan','plan']) &&
                           !!val(row, idx, 'rfp link'));

    var copiedDeck = '';
    var copiedPlan = '';
    var notes = [];

    if (!channel || !partner) {
      notes.push('Missing Channel or Partner');
      linkMap[(channel + '|' + partner).toLowerCase()] = {
        channel: channel, partner: partner,
        deck: live ? '' : '(preview) missing Channel/Partner',
        plan: live ? '' : '(preview) missing Channel/Partner',
        updatedAt: live ? nowStr : '',
        notes: notes.join(' | ')
      };
      continue;
    }

    try {
      if (live) {
        // Create/locate destination folders only in LIVE
        var channelFolder = getOrCreateSubFolder_(parent, channel);
        var partnerFolder = getOrCreateSubFolder_(channelFolder, partner);

        // ---- Deck (LIVE) ----
        if (deckSrc) {
          var deckId = parseGoogleFileId_(deckSrc);
          if (deckId) {
            try {
              var dFile = DriveApp.getFileById(deckId);
              var dCopy = dFile.makeCopy(safeName_(partner + ' - ' + channel + ' - Deck'), partnerFolder);
              if (SHARE_ANYONE_VIEW) setAnyoneWithLinkView_(dCopy);
              copiedDeck = 'https://drive.google.com/file/d/' + dCopy.getId() + '/view';
            } catch (eD) { notes.push('Deck copy error: ' + eD.message); }
          } else { notes.push('Deck src not Drive URL'); }
        } else { notes.push('No deck src'); }

        // ---- Plan (LIVE) ----
        if (planSrc) {
          var planId = parseGoogleFileId_(planSrc);
          if (planId) {
            try {
              var pFile = DriveApp.getFileById(planId);
              var pCopy = pFile.makeCopy(safeName_(partner + ' - ' + channel + ' - Media Plan'), partnerFolder);
              if (SHARE_ANYONE_VIEW) setAnyoneWithLinkView_(pCopy);
              copiedPlan = 'https://drive.google.com/file/d/' + pCopy.getId() + '/view';
              if (planCameFromRfp) notes.push('Plan source = RFP Link');
            } catch (eP) { notes.push('Plan copy error: ' + eP.message); }
          } else { notes.push('Plan src not Drive URL'); }
        } else { notes.push('No plan src'); }

      } else {
        // ---- PREVIEW: never touch Drive ----
        copiedDeck = deckSrc ? '(preview) ' + deckSrc : '(preview) no deck src';
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

    var key = (channel + '|' + partner).toLowerCase();
    linkMap[key] = {
      channel: channel,
      partner: partner,
      deck: copiedDeck,
      plan: copiedPlan,
      updatedAt: live ? nowStr : '',
      notes: notes.join(' | ')
    };
  }

  // Write to Partner Links
  upsertPartnerLinks_(linkMap, live);

  toast_((live ? 'LIVE' : 'PREVIEW') + ' processed rows: ' + processed);
  logLine_((live ? 'LIVE' : 'PREVIEW') + ' processed=' + processed + ', partners=' + Object.keys(linkMap).length);
}

/* -------- Partner Links write-back -------- */
function upsertPartnerLinks_(map, live){
  var ss=SpreadsheetApp.getActive();
  var tab=ss.getSheetByName(TARGET_TAB);
  if(!tab) tab=ss.insertSheet(TARGET_TAB);

  var headers=['Channel','Partner','Copied Deck Link','Copied Media Plan Link','Last Updated','Notes'];
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
    var rowVals=[[r.channel||'', r.partner||'', r.deck||'', r.plan||'', r.updatedAt||'', r.notes||'']];
    if(idx[key] && OVERWRITE_EXISTING){
      tab.getRange(idx[key],1,1,rowVals[0].length).setValues(rowVals);
      count++;
    } else if(!idx[key]) {
      tab.appendRow(rowVals[0]);
      count++;
    }
  });

  logLine_('Partner Links upserted: ' + count + ' rows (live=' + live + ').');
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

// NEW: Helper to fetch Parent Folder ID from "Inputs" tab
function getParentFolderIdFromInputs_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(INPUTS_TAB_NAME);
  if (!sheet) return null;
  var val = sheet.getRange(PARENT_FOLDER_CELL).getValue();
  return val ? String(val).trim() : null;
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
function parseGoogleFileId_(u){if(!u)return null;var m=String(u).match(/\/d\/([a-zA-Z0-9_-]{20,})/);if(m)return m[1];var m2=String(u).match(/[?&]id=([a-zA-Z0-9_-]{20,})/);return m2?m2[1]:null;}
function safeName_(s){return (s||'file').replace(/[\/\\:*?"<>|]+/g,'_').slice(0,120);}
function setAnyoneWithLinkView_(f){try{f.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);}catch(e){}}
function ensureExactHeaders_(tab,h){if(tab.getLastRow()===0){tab.getRange(1,1,1,h.length).setValues([h]);return;}var cur=tab.getRange(1,1,1,h.length).getValues()[0];var mismatch=false;for(var i=0;i<h.length;i++){if(String(cur[i]||'')!==h[i]){mismatch=true;break;}}if(mismatch){tab.clear();tab.getRange(1,1,1,h.length).setValues([h]);}}
function toast_(msg){SpreadsheetApp.getActive().toast(msg,'RFP Automation',6);}
function logLine_(msg){var ss=SpreadsheetApp.getActive();var t=ss.getSheetByName('Automation Log');if(!t)t=ss.insertSheet('Automation Log');t.appendRow([new Date(),msg]);}