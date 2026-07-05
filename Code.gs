/*
  UPDATED Apps Script for the AA Scooters spreadsheet.
  This REPLACES your current Code.gs entirely.

  What changed vs. the previous version:
  - Fixed a bug where new customer rows saved "Renting from" / "Return date"
    in yyyy-MM-dd format (e.g. 2026-07-04) instead of dd/MM/yyyy like the
    rest of the sheet. Added formatIsoDateToDMY() and used it in doPost()
    when appending the row.
  - Added Bike Photos support, backing the new bikephotos.html page:
      * uploadPhoto (doPost)  — saves a photo into that bike's Drive subfolder
      * deletePhoto (doPost)  — trashes a photo by file ID
      * bikePhotos  (doGet)   — lists a bike's photos
    All photos live under one Drive folder (PHOTOS_ROOT_FOLDER_ID below),
    with one auto-created subfolder per bike name.
  - Everything else (doGet, getPartsData, getOperationStatusRows,
    updateBikeRow) is unchanged from before.

  ACTION NEEDED: set PHOTOS_ROOT_FOLDER_ID below to your real Drive folder
  ID before deploying (see the comment next to it).

  After pasting this in, click Deploy > Manage deployments > Edit (pencil) >
  select "New version" > Deploy. The web app URL stays the same.
*/

var PARTS_SHEET_NAME = 'Parts and Oil change';
var OPERATION_SHEET_NAME = 'Operation';

// ID of the Drive folder that holds one subfolder per bike, full of that
// bike's photos. Get this from the folder's URL:
// https://drive.google.com/drive/folders/THIS_PART_IS_THE_ID
// The Google account running this script must have at least Editor access
// to this folder (owning it themselves is simplest).
var PHOTOS_ROOT_FOLDER_ID = '1E11bBgY5BeohoSiDCffA1Uz4-YQJOt7U';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'updateBike') {
      return updateBikeRow(data);
    }
    if (data.action === 'uploadPhoto') {
      return uploadBikePhoto(data);
    }
    if (data.action === 'deletePhoto') {
      return deleteBikePhoto(data);
    }

    // ---- Customer-intake behavior ----
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('customer');
    if (!sheet) {
      throw new Error('Sheet named "customer" not found in this spreadsheet.');
    }

    sheet.appendRow([
      Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy'),
      data.contact || '',
      data.name || '',
      data.nationality || '',
      data.passport || '',
      data.bikeModel || '',
      '',
      formatIsoDateToDMY(data.rentingDateFrom),
      formatIsoDateToDMY(data.returnDate),
      data.returnTime || '',
      data.deliverToHotel || '',
      data.totalPrice || '',
      data.paidBy || ''
    ]);

    var newRow = sheet.getLastRow();
    var numCols = 13;
    var newRange = sheet.getRange(newRow, 1, 1, numCols);
    newRange.setBorder(true, true, true, true, true, true);

    var fromDate = data.rentingDateFrom ? new Date(data.rentingDateFrom + 'T00:00:00') : null;
    var toDate = data.returnDate ? new Date(data.returnDate + 'T00:00:00') : null;
    if (fromDate && toDate) {
      var dayCount = Math.round((toDate - fromDate) / (1000 * 60 * 60 * 24));
      var fillColor = dayCount >= 30 ? '#00ffff' : '#93C47D';
      sheet.getRange(newRow, 2, 1, 3).setBackground(fillColor);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Converts a yyyy-MM-dd string (what <input type="date"> sends) into
// dd/MM/yyyy, matching the format already used everywhere else in the sheet.
// If the string doesn't look like yyyy-MM-dd, it's returned unchanged.
function formatIsoDateToDMY(isoStr) {
  if (!isoStr) return '';
  var m = String(isoStr).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return isoStr;
  var y = m[1];
  var mo = m[2].length === 1 ? '0' + m[2] : m[2];
  var d = m[3].length === 1 ? '0' + m[3] : m[3];
  return d + '/' + mo + '/' + y;
}

// ---- Update one row in the Parts and Oil change tab (unchanged) ----
function updateBikeRow(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PARTS_SHEET_NAME);
    if (!sheet) {
      throw new Error('Sheet named "' + PARTS_SHEET_NAME + '" not found in this spreadsheet.');
    }

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    var headerToCol = {};
    headers.forEach(function(h, i) {
      var key = (h || '').toString().trim();
      if (!key) key = 'Column ' + columnLetter(i + 1);
      headerToCol[key] = i + 1;
    });

    var bikeCol = 1;
    var lastRow = sheet.getLastRow();
    var bikeNames = sheet.getRange(1, bikeCol, lastRow, 1).getValues();

    var targetRow = -1;
    for (var r = 0; r < bikeNames.length; r++) {
      if ((bikeNames[r][0] || '').toString().trim() === (data.bike || '').toString().trim()) {
        targetRow = r + 1;
        break;
      }
    }

    if (targetRow === -1) {
      throw new Error('Bike "' + data.bike + '" not found in "' + PARTS_SHEET_NAME + '".');
    }

    var fields = data.fields || {};
    Object.keys(fields).forEach(function(headerName) {
      var col = headerToCol[headerName];
      if (col) {
        sheet.getRange(targetRow, col).setValue(fields[headerName]);
      }
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Bike Photos: Drive-backed storage ----
// One subfolder per bike, named exactly as the bike appears in the Parts
// and Oil change tab, sitting inside PHOTOS_ROOT_FOLDER_ID. Created lazily
// the first time a photo is uploaded for that bike.

function getOrCreateBikeFolder(bikeName) {
  var root = DriveApp.getFolderById(PHOTOS_ROOT_FOLDER_ID);
  var name = (bikeName || '').toString().trim();
  if (!name) throw new Error('No bike name given.');

  var existing = root.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();

  return root.createFolder(name);
}

function fileToPhotoObject(file) {
  return {
    id: file.getId(),
    name: file.getName(),
    url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000'
  };
}

// ---- doPost: action = 'uploadPhoto' ----
// Expects: { action, bike, filename, mimeType, base64 }
function uploadBikePhoto(data) {
  try {
    if (!data.bike) throw new Error('No bike specified.');
    if (!data.base64) throw new Error('No image data received.');

    var folder = getOrCreateBikeFolder(data.bike);
    var bytes = Utilities.base64Decode(data.base64);
    var blob = Utilities.newBlob(bytes, data.mimeType || 'image/jpeg', data.filename || 'photo.jpg');

    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, photo: fileToPhotoObject(file) }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- doPost: action = 'deletePhoto' ----
// Expects: { action, fileId }
// Moves the file to Trash rather than permanently deleting it, so an
// accidental tap can still be recovered from Drive's Trash if needed.
function deleteBikePhoto(data) {
  try {
    if (!data.fileId) throw new Error('No file specified.');
    var file = DriveApp.getFileById(data.fileId);
    file.setTrashed(true);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- doGet: action = 'bikePhotos', param bike=<name> ----
// Returns { success, photos: [{id, name, url}, ...] }. If the bike has no
// folder yet (no photos ever uploaded), returns an empty list rather than
// an error.
function getBikePhotos(bikeName) {
  try {
    var name = (bikeName || '').toString().trim();
    if (!name) throw new Error('No bike specified.');

    var root = DriveApp.getFolderById(PHOTOS_ROOT_FOLDER_ID);
    var folders = root.getFoldersByName(name);
    if (!folders.hasNext()) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, photos: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var folder = folders.next();
    var files = folder.getFiles();
    var photos = [];
    while (files.hasNext()) {
      var f = files.next();
      photos.push(fileToPhotoObject(f));
    }
    // Most recently added first.
    photos.reverse();

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, photos: photos }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function columnLetter(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ---- Used by the "Search" screen on the Customer Record page (unchanged) ----

var HEADER_ROWS = 1;

function doGet(e) {
  try {
    if (e.parameter.action === 'parts') {
      return getPartsData();
    }
    if (e.parameter.action === 'bikePhotos') {
      return getBikePhotos(e.parameter.bike);
    }

    // ---- Customer-search behavior ----
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('customer');
    if (!sheet) {
      throw new Error('Sheet named "customer" not found in this spreadsheet.');
    }

    var values = sheet.getDataRange().getValues();
    var keys = ['timestamp','contact','name','nationality','passport','bikeModel',
                'status','rentingDateFrom','returnDate','returnTime','deliverToHotel',
                'totalPrice','paidBy'];
    var tz = ss.getSpreadsheetTimeZone();

    function cellToString(key, val) {
      if (val instanceof Date) {
        if (key === 'returnTime') {
          return Utilities.formatDate(val, tz, 'HH:mm');
        }
        return Utilities.formatDate(val, tz, 'dd/MM/yyyy');
      }
      return val !== undefined && val !== null ? String(val) : '';
    }

    var rows = values.slice(HEADER_ROWS).map(function(row) {
      var obj = {};
      keys.forEach(function(k, i) {
        obj[k] = cellToString(k, row[i]);
      });
      return obj;
    }).filter(function(r) { return r.name !== ''; });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, rows: rows }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Serve the Parts and Oil change tab, PLUS the Operation tab's Bike +
// Status columns (operationRows), so pages can use whichever Status is
// actually correct without needing a separate round trip. ----
function getPartsData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PARTS_SHEET_NAME);
    if (!sheet) {
      throw new Error('Sheet named "' + PARTS_SHEET_NAME + '" not found in this spreadsheet.');
    }

    var values = sheet.getDataRange().getValues();
    var tz = ss.getSpreadsheetTimeZone();
    var rawHeaders = values[0];

    var headers = rawHeaders.map(function(h, i) {
      var key = (h || '').toString().trim();
      return key || ('Column ' + columnLetter(i + 1));
    });

    function cellToString(val) {
      if (val instanceof Date) {
        return Utilities.formatDate(val, tz, 'dd/MM/yyyy');
      }
      return val !== undefined && val !== null ? String(val) : '';
    }

    var rows = values.slice(1).map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) {
        obj[h] = cellToString(row[i]);
      });
      return obj;
    }).filter(function(r) { return (r[headers[0]] || '').toString().trim() !== ''; });

    var operationRows = getOperationStatusRows();

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        headers: headers,
        rows: rows,
        operationRows: operationRows
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Read just the Bike + Status columns from the Operation tab.
// Looks up the columns by header text ("Bike" / "Status") rather than fixed
// column letters, so it keeps working if columns get added or reordered.
// Returns [] (rather than throwing) if the tab or columns are missing, so a
// problem here never breaks the rest of the Oil Change page. ----
function getOperationStatusRows() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(OPERATION_SHEET_NAME);
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    if (!values.length) return [];

    var headerRow = values[0].map(function(h) {
      return (h || '').toString().trim().toLowerCase();
    });
    var bikeCol = headerRow.indexOf('bike');
    var statusCol = headerRow.indexOf('status');
    if (bikeCol === -1 || statusCol === -1) return [];

    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var bike = (values[i][bikeCol] || '').toString().trim();
      var status = (values[i][statusCol] || '').toString().trim();
      if (bike) rows.push({ bike: bike, status: status });
    }
    return rows;

  } catch (err) {
    return [];
  }
}
