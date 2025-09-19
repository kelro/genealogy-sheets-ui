function listAllFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let formulasList = [];
  
  sheets.forEach(sheet => {
    const formulas = sheet.getDataRange().getFormulas();
    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        if (formulas[r][c]) {
          formulasList.push(
            `${sheet.getName()}!R${r+1}C${c+1} → ${formulas[r][c]}`
          );
        }
      }
    }
  });
  
  // Output to a new sheet called "Formula_List"
  let outSheet = ss.getSheetByName("Formula_List");
  if (!outSheet) {
    outSheet = ss.insertSheet("Formula_List");
  } else {
    outSheet.clear();
  }
  
  // Write them as a single column list
  formulasList.forEach((f, i) => {
    outSheet.getRange(i+1, 1).setValue(f);
  });
}
