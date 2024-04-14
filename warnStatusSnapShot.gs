function warnStatusHalfMonthSpendSnapShot() {
  
  var sheetCurrMonthSpend = SpreadsheetApp.getActive().getSheetByName("Current Month's Spending Allowance")

  var warnStatus = sheetCurrMonthSpend.getRange('B21').getValue()
  var warnStatusFreezeCell = sheetCurrMonthSpend.getRange('B36')
  warnStatusFreezeCell.setValue(warnStatus)
}
