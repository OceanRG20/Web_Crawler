function OpenRecordView() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AF1').activate();
};