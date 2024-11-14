// Template Spreadsheet
const temSpreadsheetID = "PUT_HERE_THE_ID";

// GENERATE NEW QUOTE NUMBER
function newQuoteNum(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("scriptENV");
  const itemToSearch = "quoteNumOutput";
  setNewQuoteNum(sheet);
  const cell = searchCellInSheet(sheet, itemToSearch);
  return sheet.getRange(cell).getValue();
}
// SET URL AND LINK TO TEMPLATE
// SET QUOTE NUM TO NEWQUOTE IN DOC NAME AND IN DATA SHEET
// DELETE FIRST PAGE (WITH BUTTON)
// CREATE NEW PAGES (COPIED FROM ENV DOC)
// AFTER PAGE BUTTON ACTION:
// 1) CREATE NEW PAGES
// 2) ADD PAGE NUMBER
// AFTER PRINT BUTTON ACTION: GENERATE PDF

function oldCode() {
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
      actualSheet.getRange(columnNameRange).getValues().filter(String).length +
      1;
    actualSheet.getRange(cellToSetQuote, columnNameNum).setValue(name);

    // Add new spreadsheet url to the table
    const columnURLRange = `${columnURL}:${columnURL}`;
    const cellToSetURL =
      actualSheet.getRange(columnURLRange).getValues().filter(String).length +
      1;
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

  function getRowNum(sheet, itemToSearch, searchColumn = "A") {
    const columnToSearch = `${searchColumn}:${searchColumn}`;
    const findInColumn = sheet
      .getRange(columnToSearch)
      .createTextFinder(itemToSearch);
    const cell = findInColumn.findNext();
    return cell.getRow();
  }

  function searchCellInSheet(
    sheet,
    itemToSearch,
    searchColumn = "A",
    returnCellColumn = "B"
  ) {
    const rowNum = getRowNum(sheet, itemToSearch, searchColumn);
    return returnCellColumn + rowNum;
  }

  function searchvInSheet(
    sheet,
    itemToSearch,
    searchColumn = "A",
    returnValueColumn = "B"
  ) {
    const cell = searchCellInSheet(
      sheet,
      itemToSearch,
      searchColumn,
      returnValueColumn
    );
    const value = sheet.getRange(cell).getValue();
    return value;
  }

  function addPageToSpreadsheet(
    actualSpreadsheet,
    sheetName,
    pageNumber,
    scriptSheet
  ) {
    const sheetToCopy = actualSpreadsheet.getSheetByName(sheetName);
    if (sheetToCopy) {
      const copiedSheet = sheetToCopy.copyTo(actualSpreadsheet);

      //Change sheet name & page Number
      copiedSheet.setName("Pag" + pageNumber);
      const sheetPageCell = searchvInSheet(scriptSheet, "sheetPageCell");
      copiedSheet.getRange(sheetPageCell).setValue(pageNumber);

      const sheetFormulaCell = searchvInSheet(scriptSheet, "sheetFormulaCell");
      const quotepageFormula = copiedSheet
        .getRange(sheetFormulaCell)
        .getFormula();

      // Convert formulas to values
      const range = copiedSheet.getDataRange();
      range.copyTo(range, { contentsOnly: true });

      // Mantain  formulas for page and quote number
      copiedSheet.getRange(sheetFormulaCell).setFormula(quotepageFormula);

      // Move sheet to its position
      copiedSheet.activate();
      actualSpreadsheet.moveActiveSheet(copiedSheet.getIndex() - 2);
    }
  }

  function addPage() {
    const actualSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const scriptSheet = actualSpreadsheet.getSheetByName("scriptData");
    const activeSheetName = actualSpreadsheet.getActiveSheet().getName();

    const templateSheetName = searchvInSheet(scriptSheet, activeSheetName);
    const pageNumber = searchvInSheet(scriptSheet, "nextpage");

    addPageToSpreadsheet(
      actualSpreadsheet,
      templateSheetName,
      pageNumber,
      scriptSheet
    );

    const pageNumCell = searchCellInSheet(scriptSheet, "nextpage");

    scriptSheet.getRange(pageNumCell).setValue(pageNumber + 1);
  }
}

function createQuote() {
  // Get template spreadsheet and active spreadsheet
  const temSpreadsheet = SpreadsheetApp.openById(envSpreadsheetID);
  const temSheet = envSpreadsheet.getSheetByName("scriptENV");
  const activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = activeSpreadSheet.getSheetByName("NUEVO");

  // Get new quoteNum
  const quoteNum = newQuoteNum(temSpreadsheet);
}
