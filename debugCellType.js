function debugCellType() {
  const sh = SpreadsheetApp.getActive().getSheetByName('People');
  const val = sh.getRange(129, 3).getValue();  // Row 2, column 3 (DateOfBirth)
  Logger.log('Value: %s', val);
  Logger.log('Typeof: %s', typeof val);
  Logger.log('Instanceof Date: %s', val instanceof Date);
}

