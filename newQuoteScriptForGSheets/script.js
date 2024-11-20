// Template Spreadsheet and Index Spreadsheet
const temSpreadsheetID = "PUT-HERE-THE-ID";
const indexSpreadsheetID = "PUT-HERE-THE-ID";

// Folder for output
const folderID = "PUT-HERE-THE-ID";

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

function setQuoteData(activeSpreadSheet, quoteNum) {
  const sheet = activeSpreadSheet.getSheetByName("quoteData");
  const cell = searchCellInSheet(sheet, "quoteNum");
  sheet.getRange(cell).setValue(quoteNum);
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

  // Set spreadSheet name
  activeSpreadsheet.setName(quoteNum);

  // Set templates sheets to new spreadsheet
  setSheets(activeSpreadsheet, temSpreadsheet);

  // Set quote data
  setQuoteData(activeSpreadsheet, quoteNum);
}

function getSheetList(activeSpreadsheet) {
  const sheet = activeSpreadsheet.getSheetByName("quoteData");
  const sheetsListNames = searchvInSheet(sheet, "outputPagesList")
    .split(",")
    .map((name) => name.trim());
  const sheetsList = sheetsListNames
    .map((sheetName) => activeSpreadsheet.getSheetByName(sheetName))
    .filter((hoja) => hoja);
  // Get IDs
  const sheetIds = sheetsList
    .map((sheet) => `gid=${sheet.getSheetId()}`)
    .join("&");
  return sheetIds;
}

function setParamsToExport() {
  // Set params to export
  const params = {
    format: "pdf",
    portrait: true, // Orientación vertical
    size: "A4", // Tamaño del papel
    top_margin: 0.05, // Margen superior
    bottom_margin: 0.05, // Margen inferior
    left_margin: 0.1, // Margen izquierdo
    right_margin: 0.1, // Margen derecho
    gridlines: false, // Ocultar líneas de cuadrícula
    printtitle: false, // No incluir títulos
    sheetnames: false, // No incluir nombres de hojas
    fitw: true,
  };
  return params;
}
function getExportUrl(activeSpreadsheet, sheetsIDs) {
  // Get Export URL
  const spreadsheetId = activeSpreadsheet.getId();
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?`;

  // Set params to export
  const params = setParamsToExport();

  // Generate URL with params
  const queryParams = Object.keys(params)
    .map((key) => `${key}=${params[key]}`)
    .join("&");

  // Crear la URL final
  const finalUrl = `${exportUrl}${queryParams}&${sheetsIDs}`;

  return finalUrl;
}

function fetchURL(exportUrl) {
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: `Bearer ${token}` },
  });
  return response;
}

function savePdfToDrive(activeSpreadsheet, response) {
  const quoteName = activeSpreadsheet.getName();

  const fileName = `${quoteName}_${new Date().toISOString().slice(0, 10)}.pdf`;
  const blob = response.getBlob().setName(fileName);
  const folder = DriveApp.getFolderById(folderID);
  folder.createFile(blob);
}

function exportPdf() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get IDs of sheets to export
  const sheetsIDs = getSheetList(activeSpreadsheet);

  // Get URL to export
  const exportUrl = getExportUrl(activeSpreadsheet, sheetsIDs);

  // Fetch URL
  const response = fetchURL(exportUrl);

  // Save PDF to Output Folder in Drive
  savePdfToDrive(activeSpreadsheet, response);
}
