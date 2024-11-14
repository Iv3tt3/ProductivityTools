// Find an item in a column and return the row number

function getRowNum(sheet, itemToSearch, searchColumn = "A") {
  const columnToSearch = `${searchColumn}:${searchColumn}`;
  const findInColumn = sheet
    .getRange(columnToSearch)
    .createTextFinder(itemToSearch);
  const cell = findInColumn.findNext();
  return cell.getRow();
}

// Find an item in a range and return the cell number of the output value

function searchCellInSheet(
  sheet,
  itemToSearch,
  searchColumn = "A",
  returnCellColumn = "B"
) {
  const rowNum = getRowNum(sheet, itemToSearch, searchColumn);
  return returnCellColumn + rowNum;
}

// Find an item in a range and return the output value

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

// Generate a dict with all the keys and values of a range

function getDict(sheet, keyColumn = "A", valueColumn = "B") {
  let newDict = {};
  const columnKeyRange = `${keyColumn}:${keyColumn}`;
  const lengthFor = sheet
    .getRange(columnKeyRange)
    .getValues()
    .filter(String).length;

  for (let i = 1; i <= lengthFor; i++) {
    const key = sheet.getRange(keyColumn + i).getValue();
    const value = sheet.getRange(valueColumn + i).getValue();
    newDict[key] = value;
  }
  return newDict;
}
