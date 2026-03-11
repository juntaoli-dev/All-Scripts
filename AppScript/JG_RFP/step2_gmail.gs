/************** STEP 2: Send Deck + Plan (Dynamic Template ID from Inputs) **************/

// ---- Config (adjust if needed) ----
var TARGET_TAB = 'Partner Links';                                    // Where Step 1 wrote copied links
var CONTACT_TAB_CANDIDATES = ['RFP List','Spring RFP List'];         // Where Email + Contact live
var EMAIL_PREVIEW_TAB = 'Step 2 Email Preview';
var EMAIL_LOG_TAB = 'Step 2 Email Log';

// NEW: Configuration for where to find the Template ID
var INPUTS_TAB_NAME = 'Inputs';
var TEMPLATE_ID_CELL = 'B3';

function sendDeckAndPlanPreview() { sendDeckAndPlan_({dryRun:true}); }
function sendDeckAndPlanLive()    { sendDeckAndPlan_({dryRun:false}); }

function sendDeckAndPlan_(opts) {
  var dry = !!(opts && opts.dryRun);
  var ss = SpreadsheetApp.getActive();

  // 1) Partner Links
  var linksTab = ss.getSheetByName(TARGET_TAB);
  if (!linksTab) { toast_('Partner Links tab not found. Run Step 1 first.'); return; }
  var linksHeaders = linksTab.getRange(1,1,1,linksTab.getLastColumn()).getValues()[0];
  var lh = headerIndexMap_(linksHeaders);
  
  var required = ['channel','partner','copied deck link','copied media plan link'];
  for (var i=0;i<required.length;i++){
    if (!(required[i] in lh) && !headerLooseHas_(lh, required[i])) {
      toast_('Missing column in "'+TARGET_TAB+'": ' + required[i]); return;
    }
  }

  // Check for an override Subject column in the Partner Links tab
  var subjectColIdx = lh['subject'] || lh['subject line'] || findLoose_(lh, 'subject');

  var linksData = (linksTab.getLastRow() > 1)
    ? linksTab.getRange(2,1,linksTab.getLastRow()-1,linksTab.getLastColumn()).getValues()
    : [];

  // ✅ also read rich text so we can extract real hyperlink URLs
  var linksRich = (linksTab.getLastRow() > 1)
    ? linksTab.getRange(2,1,linksTab.getLastRow()-1,linksTab.getLastColumn()).getRichTextValues()
    : [];

  if (!linksData.length) { toast_('No rows in Partner Links to send.'); return; }

  // ✅ column indexes for the link columns
  var deckColIdx = lh['copied deck link'];
  if (deckColIdx == null) deckColIdx = findLoose_(lh, 'copied deck link');

  var planColIdx = lh['copied media plan link'];
  if (planColIdx == null) planColIdx = findLoose_(lh, 'copied media plan link');

  // 2) Contact sheet (RFP List) - Lookup Email, Contact AND Extra Notes (Col H)
  var contactSheet = getContactSheet_();
  if (!contactSheet) { toast_('Could not find a contact sheet matching: ' + CONTACT_TAB_CANDIDATES.join(', ')); return; }
  
  // Ensure we fetch enough columns to include Column H (Index 7)
  var lastCol = Math.max(contactSheet.getLastColumn(), 8);
  var chHeaders = contactSheet.getRange(1,1,1,lastCol).getValues()[0];
  var ch = headerIndexMap_(chHeaders);
  
  var needContacts = ['channel','partner','email','contact'];
  for (var j=0;j<needContacts.length;j++){
    if (!(needContacts[j] in ch) && !headerLooseHas_(ch, needContacts[j])) {
      toast_('Contact sheet missing column: ' + needContacts[j]); return;
    }
  }
  
  var chRows = (contactSheet.getLastRow() > 1)
    ? contactSheet.getRange(2,1,contactSheet.getLastRow()-1,lastCol).getValues()
    : [];

  // Build (channel|partner) -> {email, contact, notes}
  var contactMap = {};
  for (var r=0;r<chRows.length;r++){
    var row = chRows[r];
    var channelKey = safeLower_(val(row, ch, 'channel'));
    var partnerKey = safeLower_(val(row, ch, 'partner'));
    
    // Grab text from Column H (Index 7) directly
    var notesText = (row.length > 7 && row[7] != null) ? String(row[7]).trim() : '';

    if (!channelKey || !partnerKey) continue;
    contactMap[channelKey + '|' + partnerKey] = {
      email: val(row, ch, 'email'),
      contact: val(row, ch, 'contact'),
      notes: notesText
    };
  }

  // 3) Load Email Template 
  var tpl = getRichEmailTemplate_(); 

  // 4) Preview + Log sheets
  var previewTab = ss.getSheetByName(EMAIL_PREVIEW_TAB);
  if (dry) {
    if (!previewTab) previewTab = ss.insertSheet(EMAIL_PREVIEW_TAB); else previewTab.clear();
    previewTab.getRange(1,1,1,8).setValues([[
      'Channel','Partner','To','Subject','Deck Link','Plan Link','Notes','Final Body (HTML Preview)'
    ]]);
  }
  var logTab = ensureEmailLogTab_();

  var tz = Session.getScriptTimeZone();
  var nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

  // global CC list from Inputs (rows with ALIAS_EMAIL in col A, email in col B)
  var aliasCcList = getAliasCcEmailsFromInputs_();

  var processed = 0;
  for (var i2=0;i2<linksData.length;i2++){
    var lr = linksData[i2];
    var rr = linksRich[i2]; // ✅ rich row
    var channel = val(lr, lh, 'channel');
    var partner = val(lr, lh, 'partner');

    // ✅ pull the REAL URL if the cell is a hyperlink-with-text (or contains "(preview) https://...")
    var deckLink = (deckColIdx != null) ? cellLinkOrValue_(lr[deckColIdx], rr && rr[deckColIdx]) : val(lr, lh, 'copied deck link');
    var planLink = (planColIdx != null) ? cellLinkOrValue_(lr[planColIdx], rr && rr[planColIdx]) : val(lr, lh, 'copied media plan link');

    // Check for row-specific override subject in the sheet
    var rowSubject = '';
    if (subjectColIdx != null && subjectColIdx >= 0) {
       rowSubject = String(lr[subjectColIdx] || '').trim();
    }

    var key = (safeLower_(channel) + '|' + safeLower_(partner));
    var rec = contactMap[key] || {email:'', contact:'', notes:''};
    var recipients = splitEmails_(rec.email);
    var logNotes = [];

    if (!recipients.length) { logNotes.push('No email found'); writeEmailLog_(logTab, channel, partner, '', 'Skipped', nowStr, logNotes.join(' | ')); continue; }
    if (!deckLink && !planLink) { logNotes.push('No links found'); writeEmailLog_(logTab, channel, partner, recipients.join(','), 'Skipped', nowStr, logNotes.join(' | ')); continue; }

    // Use notes from Column H found in contactMap
    var extraNotesRaw = rec.notes;
    var additionalNotesHtml = additionalNotesHtml_(extraNotesRaw); // Formats as bullets if multiple lines

    // Map of tokens for replacement
    var tokenMap = {
      Partner: partner,
      Channel: channel,
      CONTACT: rec.contact,
      Contact: rec.contact,
      DECK_LINK: deckLink,
      PLAN_LINK: planLink
    };

    // Priority: 1. Sheet Subject Column -> 2. Doc Extracted Subject -> 3. Default
    var baseSubject = rowSubject ? rowSubject : (tpl.subject || 'Materials for {Partner}');
    var finalSubject = fillTemplateCurly_(baseSubject, tokenMap);
    
    // Render Body: Preserve HTML, just swap tokens
    var bodyHtml = fillTemplateCurly_(tpl.bodyHtml, tokenMap);

    // Replace {insert extra notes here} with our HTML (or nothing)
    bodyHtml = bodyHtml.replace(/\{insert extra notes here\}/gi, additionalNotesHtml || '');

    if (dry) {
      // Show CC info in preview notes so you can sanity-check it
      if (aliasCcList.length) {
        logNotes.push('Alias CC: ' + aliasCcList.join(','));
      }

      var previewText = htmlToPreviewText_(bodyHtml);
      previewTab.appendRow([
        channel,
        partner,
        recipients.join(','),
        finalSubject,
        deckLink,
        planLink,
        logNotes.join(' | '),
        previewText
      ]);
      writeEmailLog_(logTab, channel, partner, recipients.join(','), 'Simulated', nowStr, logNotes.join(' | '));
      processed++;
    } else {
      try {
        // Grant access ONLY to the specific Deck/Plan files being sent, for only the recipients (+ CC)
        var shareEmails = recipients.slice();
        for (var ccI = 0; ccI < aliasCcList.length; ccI++) shareEmails.push(aliasCcList[ccI]);
        shareEmails = dedupeEmails_(shareEmails);
        grantAccessToLinks_(deckLink, planLink, shareEmails, logNotes);

        var options = { htmlBody: bodyHtml };
        if (aliasCcList.length) {
          options.cc = aliasCcList.join(',');
        }

        GmailApp.sendEmail(recipients.join(','), finalSubject, stripHtml_(bodyHtml), options);
        writeEmailLog_(logTab, channel, partner, recipients.join(','), 'Sent', nowStr, logNotes.join(' | '));
        processed++;
      } catch(eSend) {
        logNotes.push('Send error: ' + eSend.message);
        writeEmailLog_(logTab, channel, partner, recipients.join(','), 'Error', nowStr, logNotes.join(' | '));
      }
    }
  }

  toast_((dry ? 'Preview' : 'Live') + ': processed ' + processed + ' emails.');
}

/************** Rich Template Helper (Drive API + Structure Parsing) **************/

// Fetch ID from the "Inputs" tab
function getTemplateIdFromInputs_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(INPUTS_TAB_NAME);
  if (!sheet) return null;
  var val = sheet.getRange(TEMPLATE_ID_CELL).getValue();
  return val ? String(val).trim() : null;
}

function getRichEmailTemplate_() {
  // 1. Get ID from Sheet
  var docId = getTemplateIdFromInputs_();
  
  if (!docId) {
    toast_('Error: Template ID missing in ' + INPUTS_TAB_NAME + '!' + TEMPLATE_ID_CELL);
    return { subject: 'Missing Template ID', bodyHtml: '<p>Please add Doc ID to Inputs tab.</p>' };
  }
  
  try {
    // --- STEP A: Extract Subject from Plain Text ---
    // Logic: Find "Title:", take the very next line.
    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();
    var paras = body.getParagraphs();
    var subject = '';
    
    for (var i = 0; i < paras.length; i++) {
      var txt = paras[i].getText().trim();
      // Case insensitive check for "Title:"
      if (txt.match(/^Title\s*:/i)) {
        // The subject is likely the NEXT paragraph
        if (i + 1 < paras.length) {
          subject = paras[i+1].getText().trim();
        }
        break;
      }
    }

    // --- STEP B: Extract Body from HTML ---
    // Logic: Fetch HTML, find "Copy:", and remove everything before it.
    var url = "https://www.googleapis.com/drive/v3/files/" + docId + "/export?mimeType=text/html";
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
    });
    var fullHtml = response.getContentText();

    // Extract content inside <body>
    var bodyMatch = fullHtml.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    var contentHtml = bodyMatch ? bodyMatch[1] : fullHtml;

    // Strip out everything up to and including the paragraph containing "Copy:"
    var splitRegex = /([\s\S]*?)(<p[^>]*>.*?Copy\s*:.*?<\/p>)/i;
    var match = contentHtml.match(splitRegex);
    
    if (match) {
      contentHtml = contentHtml.replace(match[0], '').trim();
    }

    return {
      subject: subject,
      bodyHtml: contentHtml
    };

  } catch (e) {
    console.log("Error fetching rich template: " + e.message);
    toast_('Error loading Doc: ' + e.message);
    return { subject: 'Error Loading Template', bodyHtml: '<p>Error loading template. Ensure Drive API is enabled and ID is correct.</p>' };
  }
}

/************** Template Helpers **************/

// Replace {Token} placeholders in HTML safely
function fillTemplateCurly_(text, map) {
  if (!text) return '';
  
  return text.replace(/\{([^}]+)\}/g, function(_, raw){
    var inner = String(raw || '').trim();
    
    // LINK tokens: {link Proposal here with name "Proposal"} / {link Media Plan here with name "Media Plan"}
    var low = inner.toLowerCase();
    if (low.indexOf('link') === 0) {
      var isDeck = /link\s+proposal\b/i.test(inner) || /link\s+deck\b/i.test(inner);
      var isPlan = /link\s+media\s*plan\b/i.test(inner) || /link\s+plan\b/i.test(inner);

      // Try to parse 'with name "X"'
      var nameMatch = inner.match(/name\s+"([^"]+)"/i) || inner.match(/name\s+'([^']+)'/i);
      var anchor = nameMatch ? nameMatch[1] : (isDeck ? 'Proposal Deck' : (isPlan ? 'Media Plan' : 'Link'));

      var href = '';
      if (isDeck) href = map.DECK_LINK;
      else if (isPlan) href = map.PLAN_LINK;
      
      if (!href) return anchor; 
      return '<a href="' + htmlEscape_(href) + '">' + htmlEscape_(anchor) + '</a>';
    }

    // Standard Token: {Partner}
    var key = inner;
    var k2 = key.replace(/\s+/g,'');
    var candidates = [key, key.toUpperCase(), key.toLowerCase(), k2, k2.toUpperCase(), k2.toLowerCase()];
    
    for (var i=0;i<candidates.length;i++){
      var k = candidates[i];
      if (map.hasOwnProperty(k)) {
        return htmlEscape_(map[k]);
      }
    }
    return '{' + inner + '}'; 
  });
}

function htmlEscape_(s){
  return String(s)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

// ✅ CHANGED: if the cell is a hyperlink-with-text, return the real URL;
// otherwise extract first https://... from the display text (ex: "(preview) https://...")
function cellLinkOrValue_(cellValue, rich) {
  var s = (cellValue == null) ? '' : String(cellValue).trim();

  try {
    if (rich && typeof rich.getLinkUrl === 'function') {
      var u = rich.getLinkUrl();
      if (u) return String(u).trim();
    }
  } catch (e) {}

  // Extract first URL from the string if present
  var m = s.match(/https?:\/\/\S+/i);
  if (m && m[0]) return m[0].replace(/[)\],.]+$/g, '');

  return s;
}

// Turn multi-line notes into an HTML bullet list
function notesToBulletsHtml_(notes) {
  if (!notes) return '';
  var lines = String(notes).split(/\r?\n/);
  var cleaned = [];
  for (var i = 0; i < lines.length; i++) {
    var t = lines[i] ? String(lines[i]).trim() : '';
    if (t) cleaned.push(t);
  }
  if (!cleaned.length) return '';
  var html = '<ul>';
  for (var j = 0; j < cleaned.length; j++) {
    html += '<li>' + htmlEscape_(cleaned[j]) + '</li>';
  }
  html += '</ul>';
  return html;
}

// Build the full "Additional Notes" section (header + bullets)
function additionalNotesHtml_(notes) {
  var list = notesToBulletsHtml_(notes);
  if (!list) return '';
  return '<p>&nbsp;</p><p><strong>Additional Notes:</strong></p>' + list;
}

// Fetch ALIAS_EMAIL CC list from the "Inputs" tab (col A = key, col B = email)
function getAliasCcEmailsFromInputs_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(INPUTS_TAB_NAME);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  var rows = sheet.getRange(1, 1, lastRow, 2).getValues();
  var ccList = [];

  for (var i = 0; i < rows.length; i++) {
    var key = rows[i][0];
    var emailVal = rows[i][1];

    if (!key || !emailVal) continue;

    if (String(key).toUpperCase().trim() === 'ALIAS_EMAIL') {
      var pieces = splitEmails_(emailVal);
      for (var j = 0; j < pieces.length; j++) {
        if (ccList.indexOf(pieces[j]) === -1) {
          ccList.push(pieces[j]);
        }
      }
    }
  }

  return ccList;
}

/************** Step 2 Helpers **************/
function getContactSheet_() {
  var ss = SpreadsheetApp.getActive();
  for (var i=0;i<CONTACT_TAB_CANDIDATES.length;i++){
    var sh = ss.getSheetByName(CONTACT_TAB_CANDIDATES[i]);
    if (sh) return sh;
  }
  var all = ss.getSheets();
  for (var j=0;j<all.length;j++){
    var name = all[j].getName().toLowerCase();
    if (name.indexOf('rfp') !== -1 && name.indexOf('list') !== -1) return all[j];
  }
  return null;
}

function ensureEmailLogTab_() {
  var ss = SpreadsheetApp.getActive();
  var tab = ss.getSheetByName(EMAIL_LOG_TAB);
  if (!tab) tab = ss.insertSheet(EMAIL_LOG_TAB);
  var headers = ['Timestamp','Channel','Partner','Recipients','Status','Notes'];
  if (tab.getLastRow() === 0) {
    tab.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return tab;
}

function writeEmailLog_(tab, channel, partner, recipients, status, timestamp, notes) {
  tab.appendRow([timestamp, channel, partner, recipients, status, notes || '']);
}

/************** Generic helpers **************/
function safeLower_(s){ return (s==null||s==='') ? '' : String(s).toLowerCase().trim(); }
function splitEmails_(s){ if(!s) return []; return String(s).split(/[,;]+/).map(function(x){return x.trim();}).filter(function(x){return x;}); }
function stripHtml_(html){ return String(html||'').replace(/<[^>]*>/g,' ').replace(/\s+/g,' ').trim(); }

function headerIndexMap_(h){ var m={}; for (var i=0;i<h.length;i++){ var n=String(h[i]||'').toLowerCase().trim(); if(n) m[n]=i; } return m; }
function headerLooseHas_(idx,k){ k=String(k||'').toLowerCase().replace(/\s+/g,''); for (var key in idx){ if(key.replace(/\s+/g,'')===k) return true; } return false; }

function val(row,idx,keys){
  if(!row) return '';
  if(typeof keys==='string') keys=[keys];
  for(var i=0;i<keys.length;i++){
    var k=String(keys[i]||'').toLowerCase();
    if(k in idx){ var v=row[idx[k]]; if(v!==''&&v!=null) return String(v).trim(); }
    for(var key in idx){ if(key.replace(/\s+/g,'')===k.replace(/\s+/g,'')){ var v2=row[idx[key]]; if(v2!==''&&v2!=null) return String(v2).trim(); } }
  }
  return '';
}
function findLoose_(h,key){ var target=String(key||'').toLowerCase().replace(/\s+/g,''); for (var k in h){ if(k.replace(/\s+/g,'')===target) return h[k]; } return null; }
function toast_(msg){ SpreadsheetApp.getActive().toast(msg,'RFP Automation',6); }

/************** Preview helper (preserve layout-ish) **************/
function htmlToPreviewText_(html) {
  if (!html) return '';
  var txt = String(html);

  txt = txt.replace(/<li[^>]*>/gi, '\n• ');

  txt = txt
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/h[1-6]>/gi, '\n\n')
    .replace(/<\/li>/gi, '');

  txt = txt.replace(/<[^>]*>/g, '');

  txt = txt.replace(/&nbsp;/gi, ' ');
  txt = txt.replace(/[ \t]+/g, ' ');
  txt = txt.replace(/\n{3,}/g,'\n\n');

  return txt.trim();
}

/************** Drive sharing helpers **************/
function extractDriveFileId_(url) {
  if (!url) return '';
  var u = String(url).trim();

  var m = u.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];

  m = u.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];

  return '';
}

function dedupeEmails_(emails) {
  var out = [];
  for (var i = 0; i < (emails || []).length; i++) {
    var e = String(emails[i] || '').trim();
    if (!e) continue;
    if (out.indexOf(e) === -1) out.push(e);
  }
  return out;
}

function grantAccessToFileId_(fileId, emails, logNotes) {
  if (!fileId) return;
  if (!emails || !emails.length) return;

  try {
    var file = DriveApp.getFileById(fileId);
    for (var i = 0; i < emails.length; i++) {
      try {
        file.addEditor(emails[i]);
      } catch (eOne) {
        if (logNotes) logNotes.push('Share failed for ' + emails[i] + ': ' + eOne.message);
      }
    }
  } catch (eFile) {
    if (logNotes) logNotes.push('Share failed (file): ' + eFile.message);
  }
}

function grantAccessToLinks_(deckLink, planLink, emails, logNotes) {
  var deckId = extractDriveFileId_(deckLink);
  var planId = extractDriveFileId_(planLink);

  if (deckId) grantAccessToFileId_(deckId, emails, logNotes);
  if (planId) grantAccessToFileId_(planId, emails, logNotes);

  if (deckLink && !deckId && logNotes) logNotes.push('Could not parse Deck file ID');
  if (planLink && !planId && logNotes) logNotes.push('Could not parse Plan file ID');
}
