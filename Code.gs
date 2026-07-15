// Meowney sync script v2
// Cara pakai: buka Google Sheet kamu -> Extensions -> Apps Script
// Hapus isi default, tempel semua kode ini, lalu Deploy > Manage deployments > edit (pensil) > New version > Deploy
// (kalau ini pertama kali deploy: Deploy > New deployment > Web app)
// Execute as: Me | Who has access: Anyone
// Salin "Web app URL" hasil deploy, tempel ke app Meowney di tab Atur

var SHEET_NAME = 'Transaction History';
var FIRST_DATA_ROW = 8;
var TX_COL_START = 1;   // A
var TX_NUM_COLS = 7;    // A:G -> Date, Category, Sub-category, Channel, Detail, Amount, Type
var TR_COL_START = 9;   // I
var TR_NUM_COLS = 5;    // I:M -> Date, From Account, To Account, Amount, Notes

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'list';
  if (action === 'list') {
    return jsonOut({ transactions: listTransactions(), transfers: listTransfers() });
  }
  return jsonOut({ status: 'ok', message: 'Meowney sync endpoint aktif' });
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'add';
    if (action === 'add') return addTransaction(data);
    if (action === 'update') return updateTransaction(data);
    if (action === 'delete') return deleteTransaction(data);
    if (action === 'transfer') return addTransfer(data);
    if (action === 'updateTransfer') return updateTransfer(data);
    if (action === 'deleteTransfer') return deleteTransfer(data);
    return jsonOut({ success: false, error: 'Aksi tidak dikenal: ' + action });
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function addTransaction(data) {
  var sheet = getSheet();
  if (!sheet) return jsonOut({ success: false, error: 'Sheet "' + SHEET_NAME + '" tidak ditemukan' });
  var targetRow = findNextEmptyRow(sheet, TX_COL_START, TX_NUM_COLS);
  var dateVal = data.date ? new Date(data.date + 'T00:00:00') : new Date();
  sheet.getRange(targetRow, TX_COL_START, 1, TX_NUM_COLS).setValues([[
    dateVal, data.category || '', data.subcategory || '', data.channel || '',
    data.detail || '', Number(data.amount) || 0, data.type || ''
  ]]);
  return jsonOut({ success: true, row: targetRow });
}

function updateTransaction(data) {
  var sheet = getSheet();
  if (!sheet) return jsonOut({ success: false, error: 'Sheet tidak ditemukan' });
  var row = Number(data.row);
  if (!row || row < FIRST_DATA_ROW) return jsonOut({ success: false, error: 'Baris tidak valid' });
  var dateVal = data.date ? new Date(data.date + 'T00:00:00') : new Date();
  sheet.getRange(row, TX_COL_START, 1, TX_NUM_COLS).setValues([[
    dateVal, data.category || '', data.subcategory || '', data.channel || '',
    data.detail || '', Number(data.amount) || 0, data.type || ''
  ]]);
  return jsonOut({ success: true, row: row });
}

function deleteTransaction(data) {
  var sheet = getSheet();
  if (!sheet) return jsonOut({ success: false, error: 'Sheet tidak ditemukan' });
  var row = Number(data.row);
  if (!row || row < FIRST_DATA_ROW) return jsonOut({ success: false, error: 'Baris tidak valid' });
  sheet.getRange(row, TX_COL_START, 1, TX_NUM_COLS).clearContent();
  return jsonOut({ success: true });
}

function addTransfer(data) {
  var sheet = getSheet();
  if (!sheet) return jsonOut({ success: false, error: 'Sheet tidak ditemukan' });
  var targetRow = findNextEmptyRow(sheet, TR_COL_START, TR_NUM_COLS);
  var dateVal = data.date ? new Date(data.date + 'T00:00:00') : new Date();
  sheet.getRange(targetRow, TR_COL_START, 1, TR_NUM_COLS).setValues([[
    dateVal, data.fromAccount || '', data.toAccount || '', Number(data.amount) || 0, data.notes || ''
  ]]);
  return jsonOut({ success: true, row: targetRow });
}

function updateTransfer(data) {
  var sheet = getSheet();
  if (!sheet) return jsonOut({ success: false, error: 'Sheet tidak ditemukan' });
  var row = Number(data.row);
  if (!row || row < FIRST_DATA_ROW) return jsonOut({ success: false, error: 'Baris tidak valid' });
  var dateVal = data.date ? new Date(data.date + 'T00:00:00') : new Date();
  sheet.getRange(row, TR_COL_START, 1, TR_NUM_COLS).setValues([[
    dateVal, data.fromAccount || '', data.toAccount || '', Number(data.amount) || 0, data.notes || ''
  ]]);
  return jsonOut({ success: true, row: row });
}

function deleteTransfer(data) {
  var sheet = getSheet();
  if (!sheet) return jsonOut({ success: false, error: 'Sheet tidak ditemukan' });
  var row = Number(data.row);
  if (!row || row < FIRST_DATA_ROW) return jsonOut({ success: false, error: 'Baris tidak valid' });
  sheet.getRange(row, TR_COL_START, 1, TR_NUM_COLS).clearContent();
  return jsonOut({ success: true });
}

function listTransactions() {
  var sheet = getSheet();
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return [];
  var numRows = lastRow - FIRST_DATA_ROW + 1;
  var vals = sheet.getRange(FIRST_DATA_ROW, TX_COL_START, numRows, TX_NUM_COLS).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (!r[0] && !r[1] && !r[2]) continue;
    out.push({
      row: FIRST_DATA_ROW + i,
      date: formatDate(r[0]),
      category: r[1],
      subcategory: r[2],
      channel: r[3],
      detail: r[4],
      amount: r[5],
      type: r[6]
    });
  }
  return out;
}

function listTransfers() {
  var sheet = getSheet();
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return [];
  var numRows = lastRow - FIRST_DATA_ROW + 1;
  var vals = sheet.getRange(FIRST_DATA_ROW, TR_COL_START, numRows, TR_NUM_COLS).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (!r[0] && !r[1] && !r[2]) continue;
    out.push({
      row: FIRST_DATA_ROW + i,
      date: formatDate(r[0]),
      fromAccount: r[1],
      toAccount: r[2],
      amount: r[3],
      notes: r[4]
    });
  }
  return out;
}

function findNextEmptyRow(sheet, colStart, numCols) {
  var lastRow = Math.max(sheet.getLastRow(), FIRST_DATA_ROW - 1);
  var numRows = lastRow - FIRST_DATA_ROW + 1;
  if (numRows <= 0) return FIRST_DATA_ROW;
  var colVals = sheet.getRange(FIRST_DATA_ROW, colStart, numRows, 1).getValues();
  for (var i = 0; i < colVals.length; i++) {
    if (colVals[i][0] === '' || colVals[i][0] === null) return FIRST_DATA_ROW + i;
  }
  return lastRow + 1;
}

function formatDate(val) {
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val)) {
    var y = val.getFullYear();
    var m = String(val.getMonth() + 1).padStart(2, '0');
    var d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return val ? String(val) : '';
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
