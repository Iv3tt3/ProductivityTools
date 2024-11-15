// Template Spreadsheet and Index Spreadsheet
const temSpreadsheetID = "PUT-HERE-THE-ID";
const indexSpreadsheetID = "PUT-HERE-THE-ID";

// Generate new quote number
function newQuoteNum(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("scriptENV");
  const itemToSearch = "quoteNumOutput";
  setNewQuoteNum(sheet);
  const cell = searchCellInSheet(sheet, itemToSearch);
  return sheet.getRange(cell).getValue();
}

// Set new quote url and name to All Quotes Doc

function setIndexData(activeSpreadsheet, quoteNum) {
  const indexSpreadsheet = SpreadsheetApp.openById(indexSpreadsheetID);
  const indexSheet = indexSpreadsheet.getSheetByName("NUEVO");

  // Add quote number to the table
  const cellToSetQuote =
    indexSheet.getRange("A:A").getValues().filter(String).length + 1;
  indexSheet.getRange(cellToSetQuote, 1).setValue(quoteNum);

  // Add spreadsheet url to the table
  const cellToSetURL =
    indexSheet.getRange("B:B").getValues().filter(String).length + 1;
  indexSheet.getRange(cellToSetURL, 2).setValue(activeSpreadsheet.getUrl());
}

// Set templates sheets to new spreadsheet
function setSheets(activeSpreadsheet, temSpreadsheet) {
  // copy templates
  const temSheet = temSpreadsheet.getSheetByName("scriptENV");
  const sheetsListToCopy = searchvInSheet(temSheet, "temSheetsList").split(",");
  sheetsListToCopy.forEach((sheetName) => {
    const sheet = temSpreadsheet.getSheetByName(sheetName);
    if (sheet) {
      //change name if sheet exists
      const existingSheet = activeSpreadsheet.getSheetByName(sheetName);
      if (existingSheet) {
        existingSheet.setName(("DELETE?" + sheetName).substring(0, 35)); //Limit sheet name in case is too long
      }
      const copiedSheet = sheet.copyTo(activeSpreadsheet);
      copiedSheet.setName(sheetName);
    }
  });

  // delete default page
  const sheet = activeSpreadsheet.getSheetByName("NUEVO");
  if (sheet) {
    activeSpreadsheet.deleteSheet(sheet);
  }
}

function createQuoteNEW() {
  // Get template needed sheets
  const temSpreadsheet = SpreadsheetApp.openById(temSpreadsheetID);
  const temSheet = temSpreadsheet.getSheetByName("scriptENV");

  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = activeSpreadsheet.getSheetByName("NUEVO");

  // Get new quoteNum
  const quoteNum = newQuoteNum(temSpreadsheet);

  // Set quoteNum and url to index spreadsheet
  setIndexData(activeSpreadsheet, quoteNum);

  // Set templates sheets to new spreadsheet
  setSheets(activeSpreadsheet, temSpreadsheet);
}
