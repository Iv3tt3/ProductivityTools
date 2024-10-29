// Define here your cells of scriptData
const cellSheetName = "B4";
const cellCurrentPageNumber = "B1";
const cellSheetPageNumber = "A1";

function copySheetsToSpreadsheet(actualSpreadsheet, sheetName, pageNumber) {
  const sheet = actualSpreadsheet.getSheetByName(sheetName);
  if (sheet) {
    const copiedSheet = sheet.copyTo(actualSpreadsheet);

    //Change sheet name & page Number
    copiedSheet.setName("Pag" + pageNumber);
    copiedSheet.getRange(cellSheetPageNumber).setValue(pageNumber);

    const pageFormula = copiedSheet.getRange("G24").getFormula();
    const quoteFormula = copiedSheet.getRange("B1").getFormula();

    // Convert formulas to values
    const range = copiedSheet.getDataRange();
    range.copyTo(range, { formatOnly: true });
    range.copyTo(range, { contentsOnly: true });

    // Mantain  formulas for page and quote number
    copiedSheet.getRange("G24").setFormula(pageFormula);
    copiedSheet.getRange("B1").setFormula(quoteFormula);

    // Move sheet to its position
    copiedSheet.activate();
    actualSpreadsheet.moveActiveSheet(copiedSheet.getIndex() - 2);
  }
}

function addMH() {
  const actualSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const scriptSheet = actualSpreadsheet.getSheetByName("scriptData");
  const sheetName = scriptSheet.getRange(cellSheetName).getValue();
  const pageNumber = scriptSheet.getRange(cellCurrentPageNumber).getValue();

  copySheetsToSpreadsheet(actualSpreadsheet, sheetName, pageNumber);

  scriptSheet.getRange(cellCurrentPageNumber).setValue(pageNumber + 1);
}
