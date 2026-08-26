// === Culdcept Begins - Google Sheets Cloud Storage ===
// Setup:
// 1. Create a new Google Sheet at https://sheets.google.com
// 2. Extensions → Apps Script
// 3. Delete any existing code, paste this entire file
// 4. Save project (name it anything)
// 5. Deploy → New deployment
//    - Type: Web app
//    - Execute as: Me
//    - Who has access: Anyone (so the HTML page can read/write without login)
// 6. Copy the Web App URL
// 7. Paste the URL into the "クラウドURL" field in the HTML page
// 8. Click クラウド保存 to upload data to Google Sheets
// 9. Click クラウド読込 to download data from Google Sheets

const SHEET_NAME = 'CuldceptBooks';
const MAX_BOOKS = 50;

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify(getAllBooks())
  ).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const books = data.books || {};
    saveAllBooks(books);
    return ContentService.createTextOutput(
      JSON.stringify({status: 'ok', count: Object.keys(books).length})
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({status: 'error', message: err.toString()})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doOptions(e) {
  return ContentService.createTextOutput('OK');
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['bookName', 'data']);
    sheet.getRange('A1:B1').setFontWeight('bold');
  }
  return sheet;
}

function getAllBooks() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const books = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      try {
        books[row[0]] = JSON.parse(row[1] || '{}');
      } catch (e) {
        books[row[0]] = {cards: [], created: Date.now()};
      }
    }
  }
  return books;
}

function saveAllBooks(books) {
  const sheet = getSheet();
  const names = Object.keys(books);
  if (names.length > MAX_BOOKS) {
    throw new Error('Maximum ' + MAX_BOOKS + ' books allowed');
  }
  const rows = names.map(name => [name, JSON.stringify(books[name])]);
  sheet.clearContents();
  sheet.appendRow(['bookName', 'data']);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}
