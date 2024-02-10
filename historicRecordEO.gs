function populateHistoricRecordEO() {

  var eoMonth = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Current Month's Spending Allowance").getRange("B35").getValue();

  if (eoMonth == true) {
    populateHistoricRecord()
  }

}
