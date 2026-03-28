// === PASTE THIS INTO YOUR GOOGLE SHEET'S APPS SCRIPT ===
// Go to Extensions > Apps Script in your spreadsheet
// Replace the content of Code.gs with this code
// Then: Deploy > Manage Deployments > Edit > New Version > Deploy
// (If first time: Deploy > New Deployment > Web App > Execute as "Me" > Anyone)

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');

    // Route by action
    if (data.action === 'update') {
      return handleUpdate(sheet, data);
    } else if (data.action === 'delete') {
      return handleDelete(sheet, data);
    } else {
      return handleAdd(sheet, data);
    }

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleAdd(sheet, data) {
  var lastRow = sheet.getLastRow();
  var amount = data.amount;
  var amountFormatted = amount >= 0
    ? '$' + Math.abs(amount).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',')
    : '-$' + Math.abs(amount).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');

  var balanceFormula = '=F' + lastRow + '+VALUE(SUBSTITUTE(SUBSTITUTE(E' + (lastRow + 1) + ',"$",""),",",""))';

  var row = [
    data.bank || 'BOFA',
    data.date,
    data.category || 'Other',
    data.description || '',
    amountFormatted,
    balanceFormula,
    data.type || 'Expense',
    data.method || 'Card',
    data.notes || ''
  ];

  sheet.appendRow(row);

  var newRow = lastRow + 1;
  sheet.getRange(newRow, 2).setNumberFormat('MMM d, yyyy');

  SpreadsheetApp.flush();
  var balance = sheet.getRange(newRow, 6).getValue();

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    action: 'add',
    row: newRow,
    balance: balance
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleUpdate(sheet, data) {
  var row = data.row; // 1-indexed row number in the sheet

  // Update individual cells
  if (data.date !== undefined) {
    sheet.getRange(row, 2).setValue(data.date);
    sheet.getRange(row, 2).setNumberFormat('MMM d, yyyy');
  }
  if (data.category !== undefined) {
    sheet.getRange(row, 3).setValue(data.category);
  }
  if (data.description !== undefined) {
    sheet.getRange(row, 4).setValue(data.description);
  }
  if (data.amount !== undefined) {
    var amount = data.amount;
    var amountFormatted = amount >= 0
      ? '$' + Math.abs(amount).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',')
      : '-$' + Math.abs(amount).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    sheet.getRange(row, 5).setValue(amountFormatted);
  }
  if (data.type !== undefined) {
    sheet.getRange(row, 7).setValue(data.type);
  }
  if (data.method !== undefined) {
    sheet.getRange(row, 8).setValue(data.method);
  }
  if (data.notes !== undefined) {
    sheet.getRange(row, 9).setValue(data.notes);
  }

  SpreadsheetApp.flush();

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    action: 'update',
    row: row
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleDelete(sheet, data) {
  var row = data.row;
  sheet.deleteRow(row);
  SpreadsheetApp.flush();

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    action: 'delete',
    row: row
  })).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    message: 'LightNet Finance API v2'
  })).setMimeType(ContentService.MimeType.JSON);
}
