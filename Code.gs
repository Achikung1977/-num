const SHEET_ID = '11YbsYCzTPau9Qw-_CGvbxE3Y8Ue3RDikuFB74AksNAo';
const SHEET_NAME = 'data';
const PAGE_SIZE = 20;

// เปิดหน้า Web App
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Dashboard Number');
}

// ดึงข้อมูลจาก Google Sheet
function getNumbers(page) {

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName(SHEET_NAME);

  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    return {
      data: [],
      totalPages: 0
    };
  }

  const values = sheet
    .getRange(1, 1, lastRow, 2)
    .getValues();

  const start = (page - 1) * PAGE_SIZE;
  const end = start + PAGE_SIZE;

  const pageData = values.slice(start, end).map((row, index) => {
    return {
      row: start + index + 1,
      number: row[0],
      status: row[1] || 'Active'
    };
  });

  return {
    data: pageData,
    totalPages: Math.ceil(values.length / PAGE_SIZE)
  };
}

// เปลี่ยนสถานะเป็น Inactive
function setInactive(rowNumber) {

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName(SHEET_NAME);

  sheet.getRange(rowNumber, 2).setValue('Inactive');

  return true;
}
