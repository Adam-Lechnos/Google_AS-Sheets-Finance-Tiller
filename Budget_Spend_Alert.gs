function sendEmailAlert() {
  // Fetch error status
  var errorRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Current Month's Spending Allowance").getRange("B21"); 
  var errorStatus = errorRange.getValue();
  var spendLimit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Current Month's Spending Allowance").getRange("D21").getValue()*100;

  // Check totals sales
  if (errorStatus == true){
    // Fetch the email address
    var emailAddress = 'adam.lechnos@gmail.com';
  
    // Send Alert Email.
    var message = `${spendLimit}% or more of your monthly budget has been spent within the first 50% of the current month. Monitor your spend for the current month: https://docs.google.com/spreadsheets/d/1NJYEI72r9MGutftoybLrUrDk49mVxeTJcWrQngZ3tho/edit#gid=469816328`; // Second column
    var subject = 'Monthly Spending Allowance Warning!!';
    MailApp.sendEmail(emailAddress, subject, message);
    }
}