/**
 * 
 * 
 * 
 * 
 * @author dunn.shane@gmail.com (Shane Dunn)
 * 
*/

/* =========== Globals ======================= */
/**
 * Globals found in logging.gs

  var logArray_;
  var LOG_SHEET = 'Log';
  var ERROR_SHEET = 'Errors';


 * Globals found in config.gs

  var CONFIG_SHEET = 'Configuration';
*/

var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
var now = new Date();
var sDate = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

/* =========== Initialisation functions ======================= */
/**
 * Initialise the Menu and anything else needed
 *  for the succesful operation of the spreadsheet
 */
function onOpen() {
  //  add a menu when the spreadsheet is opened
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Calendar Menu')
      .addItem('Update Calendar', 'pushToCalendar')
      // .addItem('Update Calendar', 'v2pushToCalendar')
      .addToUi();

  // Add basic Configuration and Logging sheets if not setup
  var sheet = getOrCreateSheet_(CONFIG_SHEET);
  if (sheet.getRange(1, 1).getValue() == ""){
    loadNewConfiguration(sheet);
  }
  sheet = getOrCreateSheet_(LOG_SHEET);
  if (sheet.getRange(1, 1).getValue() == ""){
    loadNewLog(sheet);
  }
  sheet = getOrCreateSheet_(ERROR_SHEET);
  if (sheet.getRange(1, 1).getValue() == ""){
    loadNewError(sheet);
  }
}

/**
 * Sychonise changes in spreadsheet to the calendar.
 */
function pushToCalendar(e) {
  setupLog_();
  var i, config, configName, sheet;
  log_('pushToCalendar: Running on: ' + now);
  
  var configs = getConfigs_(getOrCreateSheet_(CONFIG_SHEET));
  
  if (!configs.length) {
    log_('No configurations found');
  } else {
    log_('Found ' + configs.length + ' configurations.');
    run_sync(configs);
  }
  log_('Script done');
    
  // Update the user about the status of the queries.
  if( e === undefined ) {
    displayLog_();
  } 
}

/**
 * Do-nothing method to trigger the authorization dialog if not already done.
 */
function checkAuthorization() {
}
 