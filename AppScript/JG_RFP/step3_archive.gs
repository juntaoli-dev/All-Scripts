/******************************************************
 * RFP Automation – STEP 3 (Auto-Archive to Progo Partner Hub)
 *
 * After a copied Sheet or Deck has not been edited for 30 days,
 * uploads a copy to the shared "Progo Partner Hub" Google Drive.
 * Files are placed in a subfolder matching the partner name
 * (scraped automatically from the file title).
 *
 * Safety guarantees:
 *   1. "Archive Log" tab tracks every file ID ever processed —
 *      a file will NEVER be archived twice.
 *   2. Before copying, we check the destination folder for a file
 *      with the same name — prevents silent duplicates.
 *   3. Existing folders in the Hub are never renamed, moved, or
 *      deleted — we only READ them or CREATE a new one if no match.
 *
 * Can be run manually from the menu or via a daily time-driven
 * trigger (use archiveToHubTriggered as the trigger function).
 *
 * ▸▸▸ REQUIRED SETUP (one-time) ◂◂◂
 *   The Hub lives on a Shared Drive. Normal DriveApp can't
 *   reliably access Shared Drives, so we use the Advanced
 *   Drive Service. You MUST enable it:
 *
 *   1. In Apps Script editor: Extensions ▸ Apps Script
 *   2. Left sidebar → Services (+ icon)
 *   3. Find "Drive API" → Add → version v3
 *      (the object name must be "Drive")
 *   4. Your Google account needs at least Contributor access
 *      on the Progo Partner Hub Shared Drive.
 ******************************************************/

/* =============== STEP 3 CONFIG ===================== */
var PROGO_HUB_FOLDER_ID   = '0AIWj8uZTfwsrUk9PVA'; // Progo Partner Hub root
var ARCHIVE_STALE_DAYS    = 30;                       // days since last edit before archiving
var ARCHIVE_LOG_TAB       = 'Archive Log';
var ARCHIVE_PREVIEW_TAB   = 'Archive Preview';
// PARTNER_LINKS_TAB already declared in folder.gs as TARGET_TAB
/* =================================================== */

/* ---------- Menu wrappers (called from onOpen in folder.gs) ---------- */
function archiveToHubPreview() { archiveToHub_({ live: false }); }
function archiveToHubLive()    { archiveToHub_({ live: true  }); }

/**
 * Trigger-friendly wrapper — point a daily time-driven trigger at
 * this function so archiving happens automatically.
 *
 * Background triggers have no "active" spreadsheet, so we read the
 * spreadsheet ID from Script Properties (auto-saved the first time
 * anyone runs Step 3 from the menu).
 */
function archiveToHubTriggered() {
  var ssId = PropertiesService.getScriptProperties().getProperty('ARCHIVE_SS_ID');
  if (!ssId) {
    console.log('ARCHIVE ERROR: Spreadsheet ID not found in Script Properties. ' +
                'Run Step 3 once from the menu first so the ID gets captured.');
    return;
  }
  var ss = SpreadsheetApp.openById(ssId);
  archiveToHub_({ live: true, ss: ss });
}

/**
 * Auto-saves the current spreadsheet ID to Script Properties.
 * Called every time a menu button is clicked (interactive context).
 * This way, when the sheet is copied, the first menu click
 * captures the NEW sheet's ID automatically — no code edits needed.
 */
function saveSpreadsheetId_() {
  var id = SpreadsheetApp.getActive().getId();
  PropertiesService.getScriptProperties().setProperty('ARCHIVE_SS_ID', id);
}

/* ================================================================== */
/*  MAIN ARCHIVE ROUTINE                                              */
/* ================================================================== */
function archiveToHub_(opts) {
  var live = !!(opts && opts.live);
  // Menu clicks → getActive(); triggers → passed in via opts.ss
  var ss   = (opts && opts.ss) ? opts.ss : SpreadsheetApp.getActive();
  var tz   = Session.getScriptTimeZone();
  var now  = new Date();
  var nowStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm');

  /* ---- 1. Read Partner Links ---- */
  var linksTab = ss.getSheetByName(TARGET_TAB);
  if (!linksTab) {
    safeToast_(ss, 'Partner Links tab not found. Run Step 1 first.');
    safeLog_(ss, 'ARCHIVE ERROR: Partner Links tab missing');
    return;
  }

  var headers   = linksTab.getRange(1, 1, 1, linksTab.getLastColumn()).getValues()[0];
  var lh        = headerIndexMap_(headers);
  var linksData = (linksTab.getLastRow() > 1)
    ? linksTab.getRange(2, 1, linksTab.getLastRow() - 1, linksTab.getLastColumn()).getValues()
    : [];

  if (!linksData.length) {
    safeToast_(ss, 'No rows in Partner Links to archive.');
    return;
  }

  /* ---- 2. Archive Log (de-duplication ledger) ---- */
  var archiveLog      = ensureArchiveLogTab_(ss);
  var alreadyArchived = loadArchivedFileIds_(archiveLog);

  /* ---- 3. Validate Hub access (live only) ---- */
  if (live) {
    var accessErr = validateHubAccess_();
    if (accessErr) {
      safeToast_(ss, accessErr);
      safeLog_(ss, 'ARCHIVE ERROR: ' + accessErr);
      return;
    }
  }

  /* ---- 4. Sub-folder cache (filled lazily on first live use) ---- */
  var hubSubFolders = {};   // lowercase partner name → { id, name }

  /* ---- 5. Preview sheet ---- */
  var previewTab = null;
  if (!live) {
    previewTab = ss.getSheetByName(ARCHIVE_PREVIEW_TAB);
    if (!previewTab) previewTab = ss.insertSheet(ARCHIVE_PREVIEW_TAB);
    else previewTab.clear();
    previewTab.getRange(1, 1, 1, 7).setValues([[
      'Partner (Scraped)', 'Channel', 'File Name', 'File Type',
      'Last Edited', 'Days Since Edit', 'Status'
    ]]);
  }

  /* ---- 6. Column indices for link columns ---- */
  var deckColIdx = lh['copied deck link'];
  if (deckColIdx == null) deckColIdx = findLoose_(lh, 'copied deck link');

  var planColIdx = lh['copied media plan link'];
  if (planColIdx == null) planColIdx = findLoose_(lh, 'copied media plan link');

  /* ---- 7. Iterate rows ---- */
  var checked  = 0;
  var archived = 0;
  var skipped  = 0;

  for (var i = 0; i < linksData.length; i++) {
    var row     = linksData[i];
    var channel = val(row, lh, 'channel');
    var partner = val(row, lh, 'partner');
    var deckUrl = (deckColIdx != null) ? String(row[deckColIdx] || '') : '';
    var planUrl = (planColIdx != null) ? String(row[planColIdx] || '') : '';

    // Build list of files to evaluate for this partner row
    var files = [];
    if (deckUrl) files.push({ url: deckUrl, type: 'Deck' });
    if (planUrl) files.push({ url: planUrl, type: 'Media Plan' });

    for (var f = 0; f < files.length; f++) {
      checked++;
      var entry   = files[f];
      var fileId  = extractDriveFileId_(entry.url);

      /* -- Bad URL -- */
      if (!fileId) {
        writePreviewOrSkip_(previewTab, live, [partner, channel, '(no parseable URL)', entry.type, '', '', 'Skipped – bad URL']);
        skipped++;
        continue;
      }

      /* -- Already archived (in our log) -- */
      if (alreadyArchived[fileId]) {
        writePreviewOrSkip_(previewTab, live, [partner, channel, '', entry.type, '', '', 'Skipped – already archived']);
        skipped++;
        continue;
      }

      /* -- Try to open the file -- */
      var driveFile;
      try {
        driveFile = DriveApp.getFileById(fileId);
      } catch (eAccess) {
        writePreviewOrSkip_(previewTab, live, [partner, channel, fileId, entry.type, '', '', 'Skipped – cannot access file']);
        skipped++;
        continue;
      }

      var lastEdited = driveFile.getLastUpdated();
      var daysSince  = Math.floor((now.getTime() - lastEdited.getTime()) / (1000 * 60 * 60 * 24));
      var fileName   = driveFile.getName();

      /* -- Not stale enough -- */
      if (daysSince < ARCHIVE_STALE_DAYS) {
        writePreviewOrSkip_(previewTab, live, [
          partner, channel, fileName, entry.type,
          Utilities.formatDate(lastEdited, tz, 'yyyy-MM-dd'),
          daysSince,
          'Skipped – only ' + daysSince + ' days old (need ' + ARCHIVE_STALE_DAYS + ')'
        ]);
        skipped++;
        continue;
      }

      /* -- Scrape partner name from file title -- */
      var scrapedPartner = scrapePartnerFromTitle_(fileName);
      if (!scrapedPartner) scrapedPartner = partner; // fallback to sheet value

      /* ========= LIVE: actually archive ========= */
      if (live) {
        var destFolderId = getOrCreateHubPartnerFolder_(hubSubFolders, scrapedPartner);

        // Final duplicate check: same file name already in the hub folder?
        if (fileExistsInFolder_(destFolderId, fileName)) {
          recordArchived_(archiveLog, fileId, fileName, scrapedPartner, nowStr, 'Skipped – duplicate name in hub');
          alreadyArchived[fileId] = true;
          skipped++;
          continue;
        }

        try {
          copyFileToSharedDrive_(fileId, fileName, destFolderId);
          recordArchived_(archiveLog, fileId, fileName, scrapedPartner, nowStr, 'Archived');
          alreadyArchived[fileId] = true;
          archived++;
        } catch (eCopy) {
          recordArchived_(archiveLog, fileId, fileName, scrapedPartner, nowStr, 'Error – ' + eCopy.message);
          skipped++;
        }

      /* ========= PREVIEW: just report ========= */
      } else {
        previewTab.appendRow([
          scrapedPartner, channel, fileName, entry.type,
          Utilities.formatDate(lastEdited, tz, 'yyyy-MM-dd'),
          daysSince,
          'Will archive → Hub/' + scrapedPartner + '/'
        ]);
        archived++;
      }
    }
  }

  var summary = (live ? 'LIVE' : 'PREVIEW') +
    ' | Checked: ' + checked +
    ' | Archived: ' + archived +
    ' | Skipped: ' + skipped;
  safeToast_(ss, summary);
  safeLog_(ss, 'ARCHIVE ' + summary);
}


/* ================================================================== */
/*  PARTNER-NAME SCRAPER                                               */
/* ================================================================== */

/**
 * Extracts the partner name from a file title.
 *
 * Step 1 names copies like:
 *   "PartnerName - Channel - Deck"
 *   "PartnerName - Channel - Media Plan"
 *
 * So we split on " - " and take the FIRST segment.
 * Falls back to the full title if there are no delimiters.
 */
function scrapePartnerFromTitle_(title) {
  if (!title) return '';
  var parts = String(title).split(/\s*-\s*/);
  // If there are at least 2 segments, the first one is the partner
  if (parts.length >= 2) {
    return parts[0].trim();
  }
  // No delimiter found — return the whole title as a last resort
  return String(title).trim();
}


/* ================================================================== */
/*  ADVANCED DRIVE SERVICE HELPERS  (Shared Drive compatible)           */
/* ================================================================== */

/**
 * Pre-flight check: can we reach the Hub and does the user have
 * write access? Returns an error string or null if everything is OK.
 */
function validateHubAccess_() {
  try {
    // This will throw if the Advanced Drive Service isn't enabled
    var meta = Drive.Files.get(PROGO_HUB_FOLDER_ID, {
      supportsAllDrives: true,
      fields: 'id,name,capabilities'
    });
    // Check write permission (capabilities.canAddChildren)
    if (meta.capabilities && meta.capabilities.canAddChildren === false) {
      return 'You do not have write access to the Hub folder ("' +
             (meta.name || PROGO_HUB_FOLDER_ID) + '"). Ask a Hub admin to add you as a Contributor.';
    }
    return null; // all good
  } catch (e) {
    if (String(e).indexOf('Drive is not defined') !== -1 ||
        String(e).indexOf('is not a function') !== -1) {
      return 'Advanced Drive Service is not enabled. ' +
             'Go to Extensions ▸ Apps Script ▸ Services (+) ▸ Drive API v3 ▸ Add.';
    }
    return 'Cannot access Hub folder: ' + e.message +
           '\nMake sure your account has access to the Progo Partner Hub Shared Drive.';
  }
}

/**
 * Lists all sub-folders inside a Shared Drive folder.
 * Returns an array of { id, name }.
 */
function listSubFoldersInHub_() {
  var folders = [];
  var pageToken = null;
  do {
    var resp = Drive.Files.list({
      q: '"' + PROGO_HUB_FOLDER_ID + '" in parents ' +
         'and mimeType = "application/vnd.google-apps.folder" ' +
         'and trashed = false',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      fields: 'nextPageToken, files(id, name)',
      pageSize: 200,
      pageToken: pageToken
    });
    var items = resp.files || [];
    for (var i = 0; i < items.length; i++) {
      folders.push({ id: items[i].id, name: items[i].name });
    }
    pageToken = resp.nextPageToken;
  } while (pageToken);
  return folders;
}

/**
 * Checks whether a file with `fileName` already exists inside
 * the given folder on the Shared Drive.
 */
function fileExistsInFolder_(folderId, fileName) {
  var escaped = fileName.replace(/'/g, "\\'");
  var resp = Drive.Files.list({
    q: '"' + folderId + '" in parents ' +
       'and name = \'' + escaped + '\' ' +
       'and trashed = false',
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    fields: 'files(id)',
    pageSize: 1
  });
  return (resp.files && resp.files.length > 0);
}

/**
 * Copies a file into a Shared Drive folder.
 */
function copyFileToSharedDrive_(sourceFileId, newName, destFolderId) {
  Drive.Files.copy(
    { name: newName, parents: [destFolderId] },
    sourceFileId,
    { supportsAllDrives: true }
  );
}

/**
 * Creates a new folder inside the Hub on the Shared Drive.
 * Returns the new folder's ID.
 */
function createFolderInHub_(folderName) {
  var meta = Drive.Files.create(
    {
      name: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [PROGO_HUB_FOLDER_ID]
    },
    null,
    { supportsAllDrives: true }
  );
  return meta.id;
}


/* ================================================================== */
/*  HUB FOLDER MATCHING (case-insensitive + substring fuzzy match)     */
/* ================================================================== */

/**
 * Finds a subfolder in the Hub whose name matches partnerName
 * (case-insensitive). If a fuzzy/substring match is found, we use
 * the EXISTING folder (never rename it). A new folder is only
 * created when there is truly no match at all.
 *
 * Returns the folder **ID** (string).
 * Results are cached in `cache` so we only list the Hub once per run.
 */
function getOrCreateHubPartnerFolder_(cache, partnerName) {
  var key = partnerName.toLowerCase().trim();

  // Fast path: already resolved
  if (cache[key]) return cache[key];

  // Lazy-load: enumerate every subfolder in the Hub ONCE
  if (!cache.__loaded__) {
    var subs = listSubFoldersInHub_();
    for (var i = 0; i < subs.length; i++) {
      cache[subs[i].name.toLowerCase().trim()] = subs[i].id;
    }
    cache.__loaded__ = true;
  }

  // Exact match (case-insensitive)
  if (cache[key]) return cache[key];

  // Fuzzy: check if the scraped name is a substring of an existing
  // folder name, or vice-versa (e.g. "Hulu" matches "Hulu Inc")
  for (var existingKey in cache) {
    if (existingKey === '__loaded__') continue;
    if (typeof cache[existingKey] !== 'string') continue; // skip non-ID values
    if (existingKey.indexOf(key) !== -1 || key.indexOf(existingKey) !== -1) {
      cache[key] = cache[existingKey];   // cache under the new key too
      return cache[key];
    }
  }

  // No match anywhere — create a brand-new folder
  var newId = createFolderInHub_(partnerName.trim());
  cache[key] = newId;
  return newId;
}


/* ================================================================== */
/*  ARCHIVE LOG (de-duplication ledger)                                */
/* ================================================================== */

function ensureArchiveLogTab_(ss) {
  var tab = ss.getSheetByName(ARCHIVE_LOG_TAB);
  if (!tab) {
    tab = ss.insertSheet(ARCHIVE_LOG_TAB);
    tab.getRange(1, 1, 1, 5).setValues([[
      'File ID', 'File Name', 'Partner', 'Archived At', 'Status'
    ]]);
  }
  return tab;
}

/**
 * Reads column A of the Archive Log and returns a lookup object
 * { fileId: true } for every file we've already touched.
 */
function loadArchivedFileIds_(logTab) {
  var map = {};
  var lastRow = logTab.getLastRow();
  if (lastRow <= 1) return map;

  var col = logTab.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    var id = String(col[i][0] || '').trim();
    if (id) map[id] = true;
  }
  return map;
}

/**
 * Appends one row to the Archive Log.
 */
function recordArchived_(logTab, fileId, fileName, partner, timestamp, status) {
  logTab.appendRow([fileId, fileName, partner, timestamp, status]);
}


/* ================================================================== */
/*  TINY HELPERS                                                       */
/* ================================================================== */

/**
 * Writes a row to the preview tab (dry-run) or silently skips (live).
 */
function writePreviewOrSkip_(previewTab, isLive, rowArray) {
  if (!isLive && previewTab) {
    previewTab.appendRow(rowArray);
  }
}

/**
 * Toast that won't crash when run from a background trigger
 * (where there is no active spreadsheet UI).
 */
function safeToast_(ss, msg) {
  try {
    ss.toast(msg, 'RFP Automation', 6);
  } catch (e) {
    // Background trigger — no UI to toast to, just log it
    console.log('ARCHIVE: ' + msg);
  }
}

/**
 * Log helper that writes to the "Automation Log" tab using the
 * provided spreadsheet object (works in background triggers).
 */
function safeLog_(ss, msg) {
  try {
    var t = ss.getSheetByName('Automation Log');
    if (!t) t = ss.insertSheet('Automation Log');
    t.appendRow([new Date(), msg]);
  } catch (e) {
    console.log('ARCHIVE LOG: ' + msg);
  }
}
