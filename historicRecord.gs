function populateHistoricRecord() {
  
  var sheetHistoricData = SpreadsheetApp.getActive().getSheetByName('Historic Data')
  var sheetCurrMonthSpend = SpreadsheetApp.getActive().getSheetByName("Current Month's Spending Allowance")

  var dateRecord = Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yyyy");
  var spendAllMonth = sheetCurrMonthSpend.getRange('D2').getValue();
  var spendAllYear = sheetCurrMonthSpend.getRange('D1').getValue();
  var numTotalAll = sheetCurrMonthSpend.getRange('B11').getValue();
  var numTotalUsed = sheetCurrMonthSpend.getRange('C11').getValue();
  var numTotalRem = sheetCurrMonthSpend.getRange('D11').getValue();
  var spendWarn = sheetCurrMonthSpend.getRange('B36').getValue();
  var spendWarnLimit = sheetCurrMonthSpend.getRange('D21').getValue();
  
  sheetHistoricData.appendRow([dateRecord,spendAllMonth,spendAllYear,numTotalAll,numTotalUsed,numTotalRem,spendWarn,spendWarnLimit]);

  var range = sheetHistoricData;
  range.sort(1, false);

}
