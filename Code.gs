/*
  Apps Script for the AA Scooters spreadsheet.

  Includes: customer intake, bike parts/oil data, bike photos (Drive-backed),
  and Bike Tax category lookup (used by Available Bikes to price by category).
*/

var PARTS_SHEET_NAME = 'Parts and Oil change';
var OPERATION_SHEET_NAME = 'Operation';
var BIKE_TAX_SHEET_NAME = 'Bike Tax';

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
    if (data.action === 'markReturned') {
      return markBikeReturned(data);
    }
    if (data.action === 'extendBike') {
      return extendBikeRow(data);
    }
    if (data.action === 'closeBikeForExtend') {
      return closeBikeForExtend(data);
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
    var dayCount = null;
    if (fromDate && toDate) {
      dayCount = Math.round((toDate - fromDate) / (1000 * 60 * 60 * 24));
      var fillColor = dayCount >= 30 ? '#00ffff' : '#93C47D';
      sheet.getRange(newRow, 2, 1, 3).setBackground(fillColor);
    }

    // Also log this rental as a row on the current month's income sheet
    // (e.g. "July"). Wrapped so a problem here never breaks customer intake.
    try {
      appendMonthlyIncomeRow(ss, data, dayCount);
    } catch (incomeErr) {
      // Swallow -- customer record is already saved; the income row is best-effort.
    }

    // If (and only if) paid in cash, also log it on the "cash" sheet.
    try {
      if ((data.paidBy || '').toString().trim().toLowerCase() === 'cash') {
        appendCashSheetRow(ss, data, dayCount);
      }
    } catch (cashErr) {
      // Swallow -- customer record is already saved; the cash row is best-effort.
    }

    // If paid by Wise or Revolut, add the amount into that method's
    // deposit-tracking cell on the current month's sheet (self-locating via
    // the label in column K), as a running "=X+Y" formula. If the label
    // can't be verified/found anywhere, nothing is written to column L --
    // instead a warning is logged and returned in the response, rather than
    // silently guessing at the wrong cell.
    var depositWarning = null;
    try {
      var paidByLower = (data.paidBy || '').toString().trim().toLowerCase();
      if (paidByLower === 'wise' || paidByLower === 'revolut') {
        processDepositForPayment(ss, paidByLower, data.totalPrice);
      }
    } catch (depositErr) {
      depositWarning = depositErr.message;
      Logger.log('Deposit update warning: ' + depositErr.message);
    }

    var responsePayload = { success: true };
    if (depositWarning) responsePayload.warning = depositWarning;

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
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

// ---- Builds the "<bike> rent <N> day(s)" text used both on the monthly
// income sheet and the cash sheet, so the two stay in the same format. ----
function buildRentalIncomeText(data, dayCount) {
  var bikeName = (data.bikeModel || '').toString().trim();
  var text = bikeName;
  if (dayCount !== null && dayCount !== undefined && !isNaN(dayCount)) {
    text += ' rent ' + dayCount + (dayCount === 1 ? ' day' : ' days');
  } else {
    text += ' rent';
  }
  return text;
}

// ---- Returns the sheet tab matching the current month's name (e.g.
// "July"), or null if there isn't one. Shared by anything that logs against
// the current month's sheet, so the lookup stays in one place. ----
function getCurrentMonthSheet(ss) {
  var monthName = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'MMMM'); // e.g. "July"
  return ss.getSheetByName(monthName);
}

// ---- Log a new rental as a row on the current month's income sheet (e.g.
// "July"). Columns: E Date, F Income, G PAX name, H Amount, I paid. Finds
// the next empty row by looking only at column E (the Date column), since
// unrelated columns further left/right (e.g. the "expense" block, or notes
// off to the side) can have data further down and shouldn't shift where a
// new income row lands. If no sheet matches the current month name, this
// silently does nothing -- it's not created automatically, since it should
// already exist as a normal monthly tab. ----
function appendMonthlyIncomeRow(ss, data, dayCount) {
  var DATE_COL = 5;   // E
  var INCOME_COL = 6; // F
  var NAME_COL = 7;   // G
  var AMOUNT_COL = 8; // H
  var PAID_COL = 9;   // I

  var sheet = getCurrentMonthSheet(ss);
  if (!sheet) return; // No tab for the current month -- nothing to log against.

  var maxRow = sheet.getMaxRows();
  var dateColValues = sheet.getRange(1, DATE_COL, maxRow, 1).getValues();
  var lastFilledRow = 1; // Assume row 1 is the header row.
  for (var i = 0; i < dateColValues.length; i++) {
    var v = dateColValues[i][0];
    if (v !== '' && v !== null) lastFilledRow = i + 1;
  }
  var targetRow = lastFilledRow + 1;

  var incomeText = buildRentalIncomeText(data, dayCount);

  var paidByRaw = (data.paidBy || '').toString().trim().toLowerCase();
  var paidDisplay = paidByRaw === 'scan' ? 'QR scan' : paidByRaw;

  var amountValue = data.totalPrice !== '' && data.totalPrice !== undefined && !isNaN(Number(data.totalPrice))
    ? Number(data.totalPrice)
    : '';

  sheet.getRange(targetRow, DATE_COL, 1, 5).setValues([[
    new Date(), incomeText, data.name || '', amountValue, paidDisplay
  ]]);

  // Match the formatting (currency style, borders, banding) of the row
  // directly above, so the new row looks consistent with the rest.
  if (lastFilledRow >= 2) {
    sheet.getRange(lastFilledRow, DATE_COL, 1, 5)
      .copyFormatToRange(sheet, DATE_COL, PAID_COL, targetRow, targetRow);
  }
}

// ---- Log a new rental as a row on the "cash" sheet, but ONLY when it was
// paid in cash. Unlike the monthly sheet, "cash" is one running log for the
// whole year (columns A income date, B income, C amount), so rows are just
// appended at the very bottom rather than looked up per month. Finds the
// next empty row via column A only, since column D ("expense"/"Tax" labels)
// can have its own unrelated entries further down. ----
function appendCashSheetRow(ss, data, dayCount) {
  var DATE_COL = 1;   // A
  var INCOME_COL = 2; // B
  var AMOUNT_COL = 3; // C

  var sheet = ss.getSheetByName('cash');
  if (!sheet) return; // No "cash" tab -- nothing to log against.

  var maxRow = sheet.getMaxRows();
  var dateColValues = sheet.getRange(1, DATE_COL, maxRow, 1).getValues();
  var lastFilledRow = 1; // Assume row 1 is the header row.
  for (var i = 0; i < dateColValues.length; i++) {
    var v = dateColValues[i][0];
    if (v !== '' && v !== null) lastFilledRow = i + 1;
  }
  var targetRow = lastFilledRow + 1;

  var incomeText = buildRentalIncomeText(data, dayCount);

  var amountValue = data.totalPrice !== '' && data.totalPrice !== undefined && !isNaN(Number(data.totalPrice))
    ? Number(data.totalPrice)
    : '';

  sheet.getRange(targetRow, DATE_COL, 1, 3).setValues([[
    new Date(), incomeText, amountValue
  ]]);

  // Match the formatting (currency style, borders) of the row directly
  // above, so the new row looks consistent with the rest.
  if (lastFilledRow >= 2) {
    sheet.getRange(lastFilledRow, DATE_COL, 1, 3)
      .copyFormatToRange(sheet, DATE_COL, AMOUNT_COL, targetRow, targetRow);
  }
}

// ---- Adds an amount into one of the fixed deposit-tracking cells on the
// current month's sheet -- L11 "wise(less deposit)" for Wise, L12
// "revolut(less deposit)" for Revolut. These are fixed reference cells (not
// something that grows with new rental rows), and the goal is to keep a
// visible running total as a formula, e.g. "=100+300", rather than just
// silently replacing the number. If the cell is empty, the formula becomes
// "=amount". If it already holds a formula, "+amount" is appended to it. If
// it holds a plain (non-formula) number -- as these cells currently do --
// that number becomes the first term of a new "=existing+amount" formula,
// so nothing already there is lost. ----
function addAmountToDepositCell(sheet, row, col, rawAmount) {
  var amount = Number(rawAmount);
  if (rawAmount === '' || rawAmount === null || rawAmount === undefined || isNaN(amount)) return;

  var range = sheet.getRange(row, col);
  var formula = range.getFormula();

  if (formula && formula.charAt(0) === '=') {
    range.setFormula(formula + '+' + amount);
    return;
  }

  var currentValue = range.getValue();
  if (currentValue === '' || currentValue === null || isNaN(Number(currentValue))) {
    range.setFormula('=' + amount);
  } else {
    range.setFormula('=' + Number(currentValue) + '+' + amount);
  }
}

// ---- Finds the row where a given label (e.g. "wise(less deposit)") lives
// in a label column (column K). Checks the expected row first; if the label
// there doesn't match (case/whitespace-insensitive), searches the whole
// column for it instead, in case the sheet's layout shifted. Returns the
// row number if found, or null if the label isn't anywhere in the column. ----
function findDepositRow(sheet, expectedRow, labelCol, expectedLabel) {
  function norm(s) { return (s || '').toString().trim().toLowerCase(); }
  var target = norm(expectedLabel);

  var atExpected = norm(sheet.getRange(expectedRow, labelCol).getValue());
  if (atExpected === target) return expectedRow;

  var maxRow = sheet.getMaxRows();
  var colValues = sheet.getRange(1, labelCol, maxRow, 1).getValues();
  for (var i = 0; i < colValues.length; i++) {
    if (norm(colValues[i][0]) === target) return i + 1;
  }
  return null; // Label not found anywhere in the column.
}

// ---- Entry point used by doPost for Wise/Revolut payments. Verifies the
// label in column K actually matches before touching column L, self-heals
// if the row shifted, and throws a clear error (rather than writing to the
// wrong cell) if the label can't be found anywhere on the sheet. ----
function processDepositForPayment(ss, paidByLower, rawAmount) {
  var LABEL_COL = 11;  // K
  var VALUE_COL = 12;  // L

  var sheet = getCurrentMonthSheet(ss);
  if (!sheet) {
    throw new Error('No sheet found for the current month -- could not update the ' + paidByLower + ' deposit total.');
  }

  var expectedRow = paidByLower === 'wise' ? 11 : 12;
  var expectedLabel = paidByLower === 'wise' ? 'wise(less deposit)' : 'revolut(less deposit)';

  var row = findDepositRow(sheet, expectedRow, LABEL_COL, expectedLabel);
  if (row === null) {
    throw new Error('Could not find a "' + expectedLabel + '" row in column K of the "' +
      sheet.getName() + '" sheet -- the ' + paidByLower + ' deposit total was NOT updated.');
  }

  addAmountToDepositCell(sheet, row, VALUE_COL, rawAmount);
}

// ---- Update one row in the Parts and Oil change tab ----
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

// ---- Update one row in the customer tab: called from the "Return" button
// on the Bikes Status page. Sets the return date to the picked date, makes
// sure that date's font color is black (it's sometimes red from earlier
// conditional/manual formatting), and flips "situation" to "Returned". ----
function markBikeReturned(data) {
  try {
    var rowNumber = parseInt(data.rowNumber, 10);
    if (!rowNumber || rowNumber < 2) {
      throw new Error('Invalid row number.');
    }
    if (!data.returnDate) {
      throw new Error('No return date given.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('customer');
    if (!sheet) {
      throw new Error('Sheet named "customer" not found in this spreadsheet.');
    }

    var CUSTOMER_RETURN_DATE_COL = 9;  // I: Return date
    var CUSTOMER_RETURN_TIME_COL = 10; // J: Return time
    var CUSTOMER_SITUATION_COL = 14;   // N: situation

    var m = String(data.returnDate).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (!m) {
      throw new Error('Return date must be in yyyy-MM-dd format.');
    }
    var y = parseInt(m[1], 10), mo = parseInt(m[2], 10) - 1, d = parseInt(m[3], 10);
    var dateValue = new Date(y, mo, d);

    var dateCell = sheet.getRange(rowNumber, CUSTOMER_RETURN_DATE_COL);
    dateCell.setValue(dateValue);
    dateCell.setFontColor('#000000');

    // The return time cell is often left red by the same earlier
    // conditional/manual formatting as the return date, so it's normalized
    // to black here too, not just the date.
    sheet.getRange(rowNumber, CUSTOMER_RETURN_TIME_COL).setFontColor('#000000');

    sheet.getRange(rowNumber, CUSTOMER_SITUATION_COL).setValue('Returned');

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Update one row in the customer tab: called from the "Extend" flow
// on the Bikes Status page for SHORT extensions only (under 30 days, and
// the "Extend 1 month" checkbox not ticked). Adds the given number of days
// onto whatever return date is currently in the sheet, and appends the
// newly paid amount onto the total price as a "=oldValue+amountPaid"
// formula. Longer extensions (1 month, or 30+ days) instead go through
// closeBikeForExtend() and a brand-new customer-intake row — see the
// "Extend" button's client-side logic in bikes.html. ----
function extendBikeRow(data) {
  try {
    var rowNumber = parseInt(data.rowNumber, 10);
    if (!rowNumber || rowNumber < 2) {
      throw new Error('Invalid row number.');
    }
    var daysToExtend = parseInt(data.daysToExtend, 10);
    if (!daysToExtend || daysToExtend <= 0) {
      throw new Error('Days to extend must be a positive number.');
    }
    var amountPaid = parseFloat(data.amountPaid);
    if (isNaN(amountPaid) || amountPaid < 0) {
      throw new Error('Amount paid must be a number.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('customer');
    if (!sheet) {
      throw new Error('Sheet named "customer" not found in this spreadsheet.');
    }

    var CUSTOMER_RETURN_DATE_COL = 9;  // I: Return date
    var CUSTOMER_TOTAL_PRICE_COL = 12; // L: total price

    var dateCell = sheet.getRange(rowNumber, CUSTOMER_RETURN_DATE_COL);
    var currentDateValue = dateCell.getValue();
    var currentDate = currentDateValue instanceof Date ? new Date(currentDateValue.getTime()) : null;
    if (!currentDate) {
      // Fall back to parsing a dd/MM/yyyy string, in case the cell isn't a real Date.
      var m = String(currentDateValue).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (!m) {
        throw new Error('Could not read the current return date to extend from.');
      }
      currentDate = new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    }
    currentDate.setDate(currentDate.getDate() + daysToExtend);
    dateCell.setValue(currentDate);

    var priceCell = sheet.getRange(rowNumber, CUSTOMER_TOTAL_PRICE_COL);
    var currentPrice = Number(priceCell.getValue()) || 0;
    priceCell.setFormula('=' + currentPrice + '+' + amountPaid);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Update one row in the customer tab: called from the "Extend" flow
// on the Bikes Status page when the extension is long (1-month checkbox
// ticked, or 30+ days typed in). Rather than pushing the due date on this
// row, that flow closes this booking out — using its current due date,
// left untouched, as the point it ended — and starts a brand-new rental
// record for the extension period instead. This just flips "situation" to
// "Returned" and normalizes the return date's and return time's font color
// (they're sometimes red from earlier conditional/manual formatting, same
// fix markBikeReturned applies); it deliberately does NOT change the
// return date's value, unlike markBikeReturned(). ----
function closeBikeForExtend(data) {
  try {
    var rowNumber = parseInt(data.rowNumber, 10);
    if (!rowNumber || rowNumber < 2) {
      throw new Error('Invalid row number.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('customer');
    if (!sheet) {
      throw new Error('Sheet named "customer" not found in this spreadsheet.');
    }

    var CUSTOMER_RETURN_DATE_COL = 9;  // I: Return date
    var CUSTOMER_RETURN_TIME_COL = 10; // J: Return time
    var CUSTOMER_SITUATION_COL = 14;   // N: situation

    sheet.getRange(rowNumber, CUSTOMER_RETURN_DATE_COL).setFontColor('#000000');
    sheet.getRange(rowNumber, CUSTOMER_RETURN_TIME_COL).setFontColor('#000000');
    sheet.getRange(rowNumber, CUSTOMER_SITUATION_COL).setValue('Returned');

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

// ---- doGet: action = 'photoFolders' ----
// Returns { success, folders: [{name, count}] } — every subfolder under
// PHOTOS_ROOT_FOLDER_ID (one per bike that's ever had a photo uploaded)
// with how many files are inside. Lets a page check photo coverage across
// every bike in one round trip instead of one bikePhotos call per bike.
function getPhotoFolders() {
  try {
    var root = DriveApp.getFolderById(PHOTOS_ROOT_FOLDER_ID);
    var folders = root.getFolders();
    var result = [];
    while (folders.hasNext()) {
      var folder = folders.next();
      var count = 0;
      var files = folder.getFiles();
      while (files.hasNext()) { files.next(); count++; }
      result.push({ name: folder.getName(), count: count });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, folders: result }))
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

// ---- Used by the "Search" screen on the Customer Record page ----

var HEADER_ROWS = 1;

function doGet(e) {
  try {
    if (e.parameter.action === 'parts') {
      return getPartsData();
    }
    if (e.parameter.action === 'bikePhotos') {
      return getBikePhotos(e.parameter.bike);
    }
    if (e.parameter.action === 'photoFolders') {
      return getPhotoFolders();
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
                'totalPrice','paidBy','situation'];
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

    var rows = values.slice(HEADER_ROWS).map(function(row, i) {
      var obj = {};
      keys.forEach(function(k, ki) {
        obj[k] = cellToString(k, row[ki]);
      });
      obj.rowNumber = HEADER_ROWS + i + 1; // 1-indexed sheet row this record lives on
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
// Status columns (operationRows), PLUS the Bike Tax tab's Bike model +
// category columns (categoryRows), so pages can price/status bikes without
// needing separate round trips. ----
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

    // Read strikethrough formatting on the bike-name column (col 1) so the
    // client can tell "sold, struck through on purpose" bikes apart from
    // bikes that are just missing/mismatched data. getFontLines() returns
    // "line-through" or "none" per cell.
    var lastRow = sheet.getLastRow();
    var strikeArray = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getFontLines() : [];

    var rows = values.slice(1).map(function(row, i) {
      var obj = {};
      headers.forEach(function(h, i2) {
        obj[h] = cellToString(row[i2]);
      });
      obj.__struck = !!(strikeArray[i] && strikeArray[i][0] === 'line-through');
      return obj;
    }).filter(function(r) { return (r[headers[0]] || '').toString().trim() !== ''; });

    var operationRows = getOperationStatusRows();
    var categoryRows = getBikeTaxCategories();

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        headers: headers,
        rows: rows,
        operationRows: operationRows,
        categoryRows: categoryRows
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

// ---- Read Bike model + category + make/model/cc/key from the Bike Tax tab.
// Looks columns up by header text so it keeps working if columns move.
// make/model/cc/key are optional — any that are blank or missing just come
// back as empty strings, since not every bike has them filled in yet.
// Returns [] (never throws) so a problem here never breaks the rest of the
// Available Bikes page. ----
function getBikeTaxCategories() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(BIKE_TAX_SHEET_NAME);
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    if (!values.length) return [];

    var headerRow = values[0].map(function(h) {
      return (h || '').toString().trim().toLowerCase();
    });
    var bikeCol = headerRow.indexOf('bike model');
    if (bikeCol === -1) bikeCol = headerRow.indexOf('bike');
    var catCol = headerRow.indexOf('category');
    if (bikeCol === -1 || catCol === -1) return [];

    var makeCol = headerRow.indexOf('make');
    // Look for model/cc/key/deposit starting from the make column onward, so
    // these always resolve to the group of columns sitting next to "make" —
    // not some unrelated column elsewhere in the sheet that happens to
    // share a name.
    var searchFrom = makeCol > -1 ? makeCol : 0;
    var modelCol = headerRow.indexOf('model', searchFrom);
    var ccCol = headerRow.indexOf('cc', searchFrom);
    var keyCol = headerRow.indexOf('key', searchFrom);
    var depositCol = headerRow.indexOf('deposit', searchFrom);
    if (depositCol === -1) depositCol = headerRow.indexOf('deposit');
    var boxCol = headerRow.indexOf('box', searchFrom);
    if (boxCol === -1) boxCol = headerRow.indexOf('box');
    var absCol = headerRow.indexOf('abs', searchFrom);
    if (absCol === -1) absCol = headerRow.indexOf('abs');
    var tractionCol = headerRow.indexOf('traction control', searchFrom);
    if (tractionCol === -1) tractionCol = headerRow.indexOf('traction control');

    // Plate number lives near the front of the sheet (e.g. "Plate No."),
    // not next to make/model, so it's looked up across the whole header
    // row rather than from searchFrom onward. Matched by "contains 'plate'"
    // rather than an exact string so header punctuation (e.g. the period
    // in "Plate No.") doesn't break the lookup.
    var plateCol = -1;
    for (var pc = 0; pc < headerRow.length; pc++) {
      if (headerRow[pc].indexOf('plate') !== -1) { plateCol = pc; break; }
    }

    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var bike = (values[i][bikeCol] || '').toString().trim();
      if (!bike) continue;
      var cat = (values[i][catCol] || '').toString().trim();
      rows.push({
        bike: bike,
        category: cat,
        make: makeCol > -1 ? (values[i][makeCol] || '').toString().trim() : '',
        model: modelCol > -1 ? (values[i][modelCol] || '').toString().trim() : '',
        cc: ccCol > -1 ? (values[i][ccCol] || '').toString().trim() : '',
        key: keyCol > -1 ? (values[i][keyCol] || '').toString().trim() : '',
        deposit: depositCol > -1 ? (values[i][depositCol] || '').toString().trim() : '',
        box: boxCol > -1 ? (values[i][boxCol] || '').toString().trim() : '',
        abs: absCol > -1 ? (values[i][absCol] || '').toString().trim() : '',
        tractionControl: tractionCol > -1 ? (values[i][tractionCol] || '').toString().trim() : '',
        plate: plateCol > -1 ? (values[i][plateCol] || '').toString().trim() : ''
      });
    }
    return rows;

  } catch (err) {
    return [];
  }
}

// =====================================================================
// ONE-TIME MIGRATION: import photos from the "Cosmetic Damage" Google Doc
// into each bike's Drive photo folder. Each Doc tab (e.g. "RAX Red") is
// matched to a real bike name from the Parts and Oil change sheet, and
// every image embedded in that tab is copied into that bike's folder.
//
// HOW TO RUN:
//   1. In the Apps Script editor, pick "importCosmeticDamagePhotos" from
//      the function dropdown at the top (next to Run/Debug).
//   2. Click Run. First time, you'll get a permissions prompt for Google
//      Docs access — Advanced > Go to (project) (unsafe) > Allow.
//   3. When it finishes, go to View > Logs (or Executions) to see a
//      summary of what was imported per tab, and which tabs (if any)
//      couldn't be confidently matched to a bike name.
//   4. This only needs to be run once. Re-running it will import the
//      same photos again as duplicates, so don't run it twice unless
//      you've deleted the previous import first.
//
// Safe to delete this whole section afterward if you don't need it again.
// =====================================================================

var COSMETIC_DAMAGE_DOC_ID = '10YMe4YqkJHT94STf40J9fQqu2L01j2Qg0Skr3gLRcPY';

function normalizeBikeNameForImport(s) {
  return (s || '').toString()
    .toLowerCase()
    .replace(/\([^)]*\)/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function bikeNamesMatchForImport(a, b) {
  var na = normalizeBikeNameForImport(a);
  var nb = normalizeBikeNameForImport(b);
  if (!na || !nb) return false;
  if (na === nb) return true;
  return na.indexOf(nb) !== -1 || nb.indexOf(na) !== -1;
}

function getAllBikeNamesForImport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PARTS_SHEET_NAME);
  if (!sheet) throw new Error('Sheet named "' + PARTS_SHEET_NAME + '" not found.');
  var values = sheet.getDataRange().getValues();
  var names = [];
  for (var i = 1; i < values.length; i++) {
    var name = (values[i][0] || '').toString().trim();
    if (name) names.push(name);
  }
  return names;
}

// Picks the best bike-name match for a Doc tab title. If several bike
// names loosely match, prefers whichever is closest in length to the tab
// title (i.e. the tightest match), rather than guessing at random.
function findBestBikeMatchForImport(tabTitle, bikeNames) {
  var candidates = bikeNames.filter(function(n) {
    return bikeNamesMatchForImport(n, tabTitle);
  });
  if (!candidates.length) return null;
  if (candidates.length === 1) return candidates[0];

  var targetLen = normalizeBikeNameForImport(tabTitle).length;
  candidates.sort(function(a, b) {
    var da = Math.abs(normalizeBikeNameForImport(a).length - targetLen);
    var db = Math.abs(normalizeBikeNameForImport(b).length - targetLen);
    return da - db;
  });
  return candidates[0];
}

// Recursively walks a Doc element tree (paragraphs, tables, table cells,
// etc.) collecting every inline image found anywhere inside it.
function collectInlineImagesForImport(element, out) {
  var type = element.getType();
  if (type === DocumentApp.ElementType.INLINE_IMAGE) {
    out.push(element.asInlineImage());
    return;
  }
  var numChildren;
  try {
    numChildren = element.getNumChildren();
  } catch (e) {
    return; // Not a container element (e.g. plain text run) — nothing to recurse into.
  }
  for (var i = 0; i < numChildren; i++) {
    collectInlineImagesForImport(element.getChild(i), out);
  }
}
