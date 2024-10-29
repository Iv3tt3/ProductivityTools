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
