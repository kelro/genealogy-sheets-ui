function exportNamedRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ranges = ss.getNamedRanges();
  var sheet = ss.insertSheet("NamedRangesExport"); // creates new sheet
  sheet.appendRow(["Name", "Range"]);

  ranges.forEach(function(nr) {
    sheet.appendRow([nr.getName(), nr.getRange().getA1Notation()]);
  });
}
