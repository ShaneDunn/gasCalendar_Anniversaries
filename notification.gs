function UpcomingBirthdaysEmail() {
  // Get Active sheet
  var BirthdaysSpreadsheet = SpreadsheetApp.getActive();
  // Sort Data in table
  var BirthdaysTable = BirthdaysSpreadsheet.getRangeByName("Birthdays");
  // Sorts by the values in the days till birthday column
  //BirthdaysTable.sort(6);

  // Get cells that contain the information for the email.
  var EmailSheet = BirthdaysSpreadsheet.getSheetByName('Emails')
  var LastRow = BirthdaysSpreadsheet.getRangeByName("SendUpcomingEmailLines").getValue();
  var UpcomingBirthdaysTable = EmailSheet.getRange(1,2,LastRow,4).getDisplayValues();

  // Translate the cells into a HTML table
  var HtmlTemplate = HtmlService.createTemplateFromFile('UpcomingBirthdaysEmail');
  HtmlTemplate.Column1 = "Person";
  HtmlTemplate.Column2 = "Date";
  HtmlTemplate.Column3 = "Turning";
  HtmlTemplate.Column4 = "Days Remaining";
  HtmlTemplate.table = UpcomingBirthdaysTable;
  var EmailTable = HtmlTemplate.evaluate().getContent();
  
  // Send Alert Email.
  var Email = 'user.name@gmail.com';
  MailApp.sendEmail({
  to: Email,
  subject: "Upcoming Birthdays",
  htmlBody:EmailTable
    });
}

function BirthdayAlertEmail() {
  // Get active sheet
  var BirthdaysSpreadsheet = SpreadsheetApp.getActive();
  // Sort Data in table
  var BirthdaysTable = BirthdaysSpreadsheet.getRangeByName("Birthdays");
  // Sorts by the values in the days till birthday column
  // BirthdaysTable.sort(6);

  // Get cells that contain the information for the email.
  var EmailSheet = BirthdaysSpreadsheet.getSheetByName('Emails')
  var SendBirthdayAlert = BirthdaysSpreadsheet.getRangeByName("SendBirthdayAlert").getValue()
  var LastRow = BirthdaysSpreadsheet.getRangeByName("SendBirthdayAlertLines").getValue();
  var BirthdaysTodayTable = EmailSheet.getRange(1,7,LastRow,2).getDisplayValues();

  // Translate the cells into a HTML table
  var HtmlTemplate = HtmlService.createTemplateFromFile('BirthdayAlertEmail');
  HtmlTemplate.Column1 = "Person";
  HtmlTemplate.Column2 = "Turning";
  HtmlTemplate.table = BirthdaysTodayTable;
  var EmailTable = HtmlTemplate.evaluate().getContent();
  
  // Send email if there is a need to
  var Email = 'user.name@gmail.com';
  if (SendBirthdayAlert){
  MailApp.sendEmail({
  to: Email,
  subject: "Birthday Alert!!!",
  htmlBody:EmailTable
    });
  }
}
