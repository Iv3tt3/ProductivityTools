// Define here your cells

// Cell with initial quote number
const cellInitNumber = "A1";

// Cell with full quote number (when format is applied)
const cellQuoteNum = "A2";

// Columns of the table to store name & url of new spreadsheets
const columnName = "E";
const columnNameNum = 5;
const columnURL = "F";
const columnURLNum = 6;

// Columns of the table whith a check & sheet name
const checkColumn = "C";
const sheetNameColumn = "D";

function getNewQuoteNumber(actualSheet) {
  const cell = actualSheet.getRange(cellInitNumber);
  const actualValue = cell.getValue();
  cell.setValue(actualValue + 1);

  return actualSheet.getRange(cellQuoteNum).getValue();
}

function createNewSpreadsheet(actualSheet, name) {
  const newSpreadsheet = SpreadsheetApp.create(name);

  // Add quote number to the table
  const columnNameRange = `${columnName}:${columnName}`;
  const cellToSetQuote =
    actualSheet.getRange(columnNameRange).getValues().filter(String).length + 1;
  actualSheet.getRange(cellToSetQuote, columnNameNum).setValue(name);

  // Add new spreadsheet url to the table
  const columnURLRange = `${columnURL}:${columnURL}`;
  const cellToSetURL =
    actualSheet.getRange(columnURLRange).getValues().filter(String).length + 1;
  actualSheet
    .getRange(cellToSetURL, columnURLNum)
    .setValue(newSpreadsheet.getUrl());

  return newSpreadsheet;
}

function getListOfSheetsToCopy(actualSheet) {
  let sheetsToCopy = [];
  const columnCheckRange = `${checkColumn}:${checkColumn}`;
  const lengthFor = actualSheet
    .getRange(columnCheckRange)
    .getValues()
    .filter(String).length;

  for (let i = 1; i <= lengthFor; i++) {
    const valueC = actualSheet.getRange(checkColumn + i).getValue();
    if (Boolean(valueC)) {
      sheetsToCopy.push(actualSheet.getRange(sheetNameColumn + i).getValue());
    }
  }
  return sheetsToCopy;
}

function copySheetsToSpreadsheet(
  sheetsList,
  actualSpreadsheet,
  newSpreadsheet
) {
  sheetsList.forEach((sheetName) => {
    const sheet = actualSpreadsheet.getSheetByName(sheetName);
    if (sheet) {
      const copiedSheet = sheet.copyTo(newSpreadsheet);
      copiedSheet.setName(sheetName);
    }
  });
}

function newquote() {
  const actualSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const actualSheet = actualSpreadsheet.getActiveSheet();

  const quoteNum = getNewQuoteNumber(actualSheet);

  const newSpreadsheet = createNewSpreadsheet(actualSheet, quoteNum);

  const sheetsToCopy = getListOfSheetsToCopy(actualSheet);

  copySheetsToSpreadsheet(sheetsToCopy, actualSpreadsheet, newSpreadsheet);
}
