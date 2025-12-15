// == ------------------------------------------------------------
function get_config(configID, configs) {
  var i, config, configName;
  for (i = 0; config = configs[i]; ++i) {
    configName = config.configurationID;
    if (config['configurationID'] === configID) {
      log_('Using configuration from: ' + configName);
      return config;
    }
  }
  log_('No configuration found: ' + configID);
}

function UpcomingBirthdaysEmail(e) {
  setupLog_();
  var i, config, configName;
  log_('Running on: ' + now);
  
  var configs = getConfigs_(getOrCreateSheet_(CONFIG_SHEET));
  configName = "UpcomingBirthdaysEmail";
  if (!configs.length) {
    log_('No configurations found');
  } else {
    log_('Found ' + configs.length + ' configurations.');
    config = get_config(configName, configs);
    Logger.log(config);
    // Get Active sheet
    var BirthdaysSpreadsheet = SpreadsheetApp.getActive();

    // Get sheet that contains the information for the email.
    var EmailSheet = BirthdaysSpreadsheet.getSheetByName(config.email_sheet)
    var LastRow = BirthdaysSpreadsheet.getRangeByName(config.ubT_last_row).getValue();
    var UpcomingBirthdaysTable = EmailSheet.getRange(config.ubT_start_row,config.ubT_start_column,LastRow,config.ubT_last_column).getDisplayValues();

    // Translate the cells into a HTML table
    var HtmlTemplate = HtmlService.createTemplateFromFile(config.email_template);
    HtmlTemplate.Column1 = "Person";
    HtmlTemplate.Column2 = "Date";
    HtmlTemplate.Column3 = "Turning";
    HtmlTemplate.Column4 = "Days Remaining";
    HtmlTemplate.table = UpcomingBirthdaysTable;
    var EmailTable = HtmlTemplate.evaluate().getContent();
  
    try {
      log_('Sending Email to: ' + config['email_address']);
      // Send Alert Email.
      MailApp.sendEmail({
        to: config.email_address,
        subject: config.email_subject,
        htmlBody: EmailTable
      });
    } catch (error) {
      log_('Error sending email to ' + config['email_address'] + ': ' + error.message);
    }
  }
  log_('Script done');
    
  // Update the user about the status of the function.
  displayLog_();
  dumpLog_(getOrCreateSheet_(LOG_SHEET));
  // dumpError_(getOrCreateSheet_(ERROR_SHEET));
  
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



function test() {
// function getAllEventsFromAllCalendars() {
  // Define a broad time range (adjust as needed)
  var startDate = new Date("January 1, 2024 00:00:00 UTC");
  var endDate = new Date("December 31, 2025 23:59:59 UTC");

  // Get all calendars the user owns or is subscribed to
  var calendars = CalendarApp.getAllCalendars();
  var allEvents = [];

  for (var i = 0; i < calendars.length; i++) {
    var calendar = calendars[i];
    Logger.log('Processing calendar: ' + calendar.getName() + ' (ID: ' + calendar.getId() + ')');

    // Get events for the defined time range
    //var events = calendar.getEvents(startDate, endDate);
    //allEvents = allEvents.concat(events); // Combine all events into one list
  }

  //Logger.log('Total number of events across all calendars: ' + allEvents.length);
  // You can then process the 'allEvents' array further, for example, write them to a Google Sheet
// }
}