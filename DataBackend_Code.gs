/**
 * Workshop Portal — Data Backend
 * -------------------------------
 * A generic key/value store backed by a Google Sheet, with automatic overflow
 * to Google Drive for any value that would be too large for a single Sheet
 * cell (Google Sheets' hard limit is 50,000 characters per cell).
 *
 * Contract (unchanged from before, the portal itself needs NO changes):
 *   GET  ?month=<key>          -> { found: true, json: "<string>" }  or { found: false }
 *   POST { month:<key>, json:<string> }  -> { ok: true }
 *
 * Large values (e.g. a month with many signature images) are automatically
 * written to a Drive file instead of the cell, with just a small pointer
 * ("drive:<fileId>") kept in the Sheet. Retrieval is transparent — the
 * caller always just gets back the original string either way.
 */

var SIZE_THRESHOLD = 35000; // stay well under Sheets' 50,000-char cell limit
var SHEET_NAME = 'PortalData'; // reuse the existing tab so previously-saved data is visible again
var DRIVE_FOLDER_NAME = 'WorkshopPortalData';
var SPREADSHEET_ID = '1NwWgdI9gxY5Magr94WxVQSy__dnKan0Z5MKfCzGc6SQ'; // Workshop Jadwal Sheet

function getSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Key', 'Value', 'UpdatedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getDriveFolder_() {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function findRow_(sh, key) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  var keys = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) return i + 2; // 1-based row number, +1 for header
  }
  return -1;
}

function safeFileName_(key) {
  return key.replace(/[^a-zA-Z0-9_-]/g, '_') + '.json';
}

function doGet(e) {
  var key = e && e.parameter && e.parameter.month;
  if (!key) return jsonOut_({ found: false });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet_();
    var row = findRow_(sh, key);
    if (row === -1) return jsonOut_({ found: false });

    var value = sh.getRange(row, 2).getValue();
    if (typeof value === 'string' && value.indexOf('drive:') === 0) {
      var fileId = value.substring(6);
      try {
        value = DriveApp.getFileById(fileId).getBlob().getDataAsString('UTF-8');
      } catch (err) {
        return jsonOut_({ found: false, error: 'drive file missing' });
      }
    }
    return jsonOut_({ found: true, json: value });
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'bad request' });
  }
  var key = body.month;
  var value = body.json;
  if (!key) return jsonOut_({ ok: false, error: 'missing key' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet_();
    var row = findRow_(sh, key);
    var toStore = value;

    if (typeof value === 'string' && value.length > SIZE_THRESHOLD) {
      // Too big for one cell — park the full content in a Drive file instead,
      // and keep only a small pointer in the Sheet.
      var folder = getDriveFolder_();
      var fileName = safeFileName_(key);
      var existing = folder.getFilesByName(fileName);
      var fileId;
      if (existing.hasNext()) {
        var f = existing.next();
        f.setContent(value);
        fileId = f.getId();
      } else {
        fileId = folder.createFile(fileName, value, MimeType.PLAIN_TEXT).getId();
      }
      toStore = 'drive:' + fileId;
    }

    if (row === -1) {
      sh.appendRow([key, toStore, new Date()]);
    } else {
      sh.getRange(row, 2, 1, 2).setValues([[toStore, new Date()]]);
    }
    return jsonOut_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
