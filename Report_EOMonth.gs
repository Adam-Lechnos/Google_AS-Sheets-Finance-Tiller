function sendEmailBudgetReportEOMonth() {

  var eoMonth = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Current Month's Spending Allowance").getRange("B35").getValue();
  var monthName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Current Month's Spending Allowance").getRange("D36").getValue();
  var spendWarn = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Current Month's Spending Allowance").getRange("B36").getValue();
  var spendWarnTh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Current Month's Spending Allowance").getRange("B38").getValue();

  if (eoMonth == true) {
    sendEmailBudgetReportArgs(monthName, spendWarn, spendWarnTh)
  }
  
}
