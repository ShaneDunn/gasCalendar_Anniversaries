//add a menu when the spreadsheet is opened
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Calendar Menu')
      .addItem('Update Calendar', 'pushToCalendar')
      .addToUi();
}

//push new events to calendar
function pushToCalendar() {

  //below are the column ids of that represents the values used in the spreadsheet (these are zero indexed)
  // 
  //Column containg the round to be played
  //Rnd	Date	Start Time	Seniors	vs	Ground	Home	Away	Player	Loaded	Event ID
  var vRound = 0;
  //Column containg the date of the game
  var vDate = 1;
  //Column containing the start time of the U17's
  var vStartTime = 2;
  //Column containing the Start time of 1st Grade
  var v1stStartTime = 3;
  //Column containg the opposition
  var vVs = 4;
  //Column containg the location of the game
  var vlocation = 5;
  //Column containing the notification status
  var vHome = 6;
  //Column containing the Guests
  var vAway = 7;
  //Column containing the Player
  var vPlayer = 8;
  //Column containing the loaded status
  var vLoaded = 9;
  //Column containing the event ID
  var vEventID = 10
  ;

  //spreadsheet variables
  var sheet = SpreadsheetApp.getActive().getSheetByName("Draw");
  var firstRow = 3; 
  var lastRow = sheet.getLastRow(); 
  var numRow = lastRow - firstRow + 1; 
  var firstCol = 31; 
  var numCol = 11; 
  var range = sheet.getRange(firstRow,firstCol,lastRow,numCol);
  var values = range.getValues();   
  var crlf = String.fromCharCode(10);

  //calendar variables
  var calendar = CalendarApp.getCalendarById('user.name@gmail.com')

  var numValues = 0;
  for (var i = 0; i < values.length; i++) {     
    //check to see if round and vs are filled out - date is left off because length is "undefined"
    if ((values[i][vRound].toString().length > 0) && (values[i][vVs].length > 0)) {

      //check if it has not been entered before          
      if (values[i][vLoaded] != 'y') {
        var vDesc = '';
        var newEventTitle = '';
        switch (values[i][vPlayer]) {
          case "Ryan":
            vDesc = 'RFNL Round ' + values[i][vRound] + crlf;
            if (values[i][vlocation] != "") {
              vDesc = vDesc + 'Vs ' + values[i][vVs] + crlf;
              vDesc = vDesc + "U17's Starting at " + Utilities.formatDate(values[i][vStartTime], 'Australia/Sydney', 'hh:mm a');
              vDesc = vDesc + "  - 1st's Starting at " + Utilities.formatDate(values[i][v1stStartTime], 'Australia/Sydney', 'hh:mm a');
            } else {
              vDesc = vDesc + values[i][vVs];
            }
            newEventTitle = 'Ryan - RFNL Rnd: ' + values[i][vRound] + ' - ' + values[i][vVs];
            break;
          case "Emma":
            vDesc = 'ACT Womens Round ' + values[i][vRound] + crlf;
            if (values[i][vlocation] != "") {
              vDesc = vDesc + 'Vs ' + values[i][vVs] + crlf;
              vDesc = vDesc + " Starting at " + Utilities.formatDate(values[i][v1stStartTime], 'Australia/Sydney', 'hh:mm a');
            } else {
              vDesc = vDesc + values[i][vVs];
            }
            newEventTitle = 'Emma - ACTW Rnd: ' + values[i][vRound] + ' - ' + values[i][vVs];
            break;
          case "Mum":
            vDesc = 'Griffith Hockey Round ' + values[i][vRound] + crlf;
            if (values[i][vlocation] != "") {
              vDesc = vDesc + 'Vs ' + values[i][vVs] + crlf;
              vDesc = vDesc + " Starting at " + Utilities.formatDate(values[i][v1stStartTime], 'Australia/Sydney', 'hh:mm a');
            } else {
              vDesc = vDesc + values[i][vVs];
            }
            newEventTitle = 'Mum - GHA Rnd: ' + values[i][vRound] + ' - ' + values[i][vVs];
            break;
          default:
            var vDesc = 'Unkown';
            var newEventTitle = 'Unkown';
            break;
        }

        //create event https://developers.google.com/apps-script/class_calendarapp#createEvent
        var options = {description: vDesc, location: values[i][vlocation]};
        Logger.log(newEventTitle);
        Logger.log(options);
        var newEvent = calendar.createAllDayEvent(newEventTitle, values[i][vDate], options);
        newEvent.removeAllReminders();

        var newEventId = newEvent.getId(); //get ID
        //mark as entered, enter ID
        sheet.getRange(firstRow+i,firstCol+vLoaded).setValue('y');
        sheet.getRange(firstRow+i,firstCol+vEventID).setValue(newEventId);

      }
      else {
        var vDesc = '';
        var newEventTitle = '';
        switch (values[i][vPlayer]) {
          case "Ryan":
            vDesc = 'RFNL Round ' + values[i][vRound] + crlf;
            if (values[i][vlocation] != "") {
              vDesc = vDesc + 'Vs ' + values[i][vVs] + crlf;
              vDesc = vDesc + "U17's Starting at " + Utilities.formatDate(values[i][vStartTime], 'Australia/Sydney', 'hh:mm a');
              vDesc = vDesc + "  - 1st's Starting at " + Utilities.formatDate(values[i][v1stStartTime], 'Australia/Sydney', 'hh:mm a');
            } else {
              vDesc = vDesc + values[i][vVs];
            }
            newEventTitle = 'Ryan - RFNL Rnd: ' + values[i][vRound] + ' - ' + values[i][vVs];
            break;
          case "Emma":
            vDesc = 'ACT Womens Round ' + values[i][vRound] + crlf;
            if (values[i][vlocation] != "") {
              vDesc = vDesc + 'Vs ' + values[i][vVs] + crlf;
              vDesc = vDesc + " Starting at " + Utilities.formatDate(values[i][v1stStartTime], 'Australia/Sydney', 'hh:mm a');
            } else {
              vDesc = vDesc + values[i][vVs];
            }
            newEventTitle = 'Emma - ACTW Rnd: ' + values[i][vRound] + ' - ' + values[i][vVs];
            break;
          case "Mum":
            vDesc = 'Griffith Hockey Round ' + values[i][vRound] + crlf;
            if (values[i][vlocation] != "") {
              vDesc = vDesc + 'Vs ' + values[i][vVs] + crlf;
              vDesc = vDesc + " Starting at " + Utilities.formatDate(values[i][v1stStartTime], 'Australia/Sydney', 'hh:mm a');
            } else {
              vDesc = vDesc + values[i][vVs];
            }
            newEventTitle = 'Mum - GHA Rnd: ' + values[i][vRound] + ' - ' + values[i][vVs];
            break;
          default:
            var vDesc = 'Unkown';
            var newEventTitle = 'Unkown';
            break;
        }
        //update event
        var id = values[i][vEventID];
        var event = calendar.getEventSeriesById(id);
        if (event !== null && typeof event != 'undefined' ) {
          var vOldDesc = event.getDescription();
          if (vDesc != vOldDesc) {
            Logger.log(vDesc);
            event.setDescription(vDesc);
          }
        }
      }
    }
    numValues++;
  } 
}
 