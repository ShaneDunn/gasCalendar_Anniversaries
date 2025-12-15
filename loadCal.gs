/**
 * 
 * 
 * 
 * 
 * @author dunn.shane@gmail.com (Shane Dunn)
*/
/* =========== Globals ======================= */

/* =========== Calendar functions ======================= */

/**
 * Creates an event in the user's default calendar.
 */
function createEvent() {
  // var calendarId = "user.name@gmail.com";
  var calendarId = 'primary';

  //below are the column ids of variables that represent the values used in the spreadsheet (these are zero indexed)
  //Column containg the Event Title
  var vTitle = 0;
  //Column containing the date of the event
  var vDate = 1;
  //Column containing the start time
  var vStartTime = 2;
  //Column containing the end time
  var vEndTime = 3;
  //Column containing the location of the event
  var vLocation = 4;
  //Column containing the notification status
  var vNotify = 5;
  //Column containing the Guests
  var vGuests = 6;
  //Column containing the loaded status
  var vLoaded = 7;
  //Column containing the event ID
  var vEventID = 8;
  //Column containing the Meet URL
  var vMeetURL = 9;
  //Column containing the description
  var vDescription = 10;

  //spreadsheet variables
  var sheet = SpreadsheetApp.getActive().getSheetByName("Hangouts");
  var lastRow = sheet.getLastRow(); 
  var lastCol = sheet.getLastColumn(); 
  var range = sheet.getRange(2,1,lastRow,lastCol);
  var values = range.getValues();   
  var cDate = new Date();
  var eDate = new Date();
  var sTime = '';
  var eTime = '';

  //calendar variables
  var calendar = CalendarApp.getCalendarById('user.name@gmail.com');
   
  for (var i = 0; i < values.length; i++) {     
    // check to see if title is filled out
    if (values[i][vTitle].length > 0) {
       
      // check if it's been entered before          
      if (values[i][vLoaded] != 'y') {                       
         
        // create event https://developers.google.com/apps-script/class_calendarapp#createEvent
        var newEventTitle = values[i][vTitle];
        if (values[i][vTitle].length > 0) {
          // set up start date and time
          cDate = new Date(values[i][vDate]);
          sTime = values[i][vStartTime];
          var Sres = sTime.split(":");
          cDate.setHours(Sres[0], Sres[1]);
          // set up end date and time
          eDate = new Date(values[i][vDate]);
          eTime = values[i][vEndTime];
          var Eres = eTime.split(":");
          eDate.setHours(Eres[0], Eres[1]);
          // Logger.log([cDate, eDate]);
          var newEvent = {
             summary: newEventTitle,
             location: values[i][vLocation],
             start: {
               dateTime: cDate.toISOString()
               },
             end: {
               dateTime: eDate.toISOString()
               },
             // Pale Blue background. Use Calendar.Colors.get() for the full list.
             colorId: 1, // PALE_BLUE,
             conferenceData: {
               createRequest: {
                 requestId: Utilities.getUuid(),
                 conferenceSolutionKey: { type: "hangoutsMeet" },
               },
             },
          };
          // Load Guests
          if (values[i][vGuests].length > 0) {
            var inviteList = values[i][vGuests].split(',');//assuming your guestlist is comma separated
            if (inviteList.length>0){
              newEvent.attendees = [];
              for(var n in inviteList){
                newEvent.attendees[n] = {email: inviteList[n]};
                // console.log(inviteList[n]);
              }
            }
          }
          // console.log(newEvent.attendees)
          newEvent = Calendar.Events.insert(newEvent, calendarId, {conferenceDataVersion: 1},);
        } else {
          newEvent = calendar.createAllDayEvent(newEventTitle, values[i][vDate], {location: values[i][vLocation]});
        }

        console.log(newEvent)

        // Check / Add / Delete Notifications
        /*
        addEmailReminder(minutesBefore)
        addPopupReminder(minutesBefore)
        addSmsReminder(minutesBefore)
        */
        
        
        //get ID
        var newEventId = newEvent.getId();
        // mark as entered, store ID
        sheet.getRange(i+2,vLoaded+1).setValue('y');
        sheet.getRange(i+2,vEventID+1).setValue(newEventId);
        sheet.getRange(i+2,vMeetURL+1).setValue(newEvent.hangoutLink);
        if (newEvent.hangoutLink){
                console.log (newEvent.summary + ' - ' + newEvent.hangoutLink + ' - ' + newEvent.description);
            }
      }
    }
  } 
}

/**
 * Previous version - Not used
 */
//push new events to calendar
function v2pushToCalendar() {
  //below are the column ids of variables that represent the values used in the spreadsheet (these are zero indexed)
  //Column containg the Event Title
  var vTitle = 0;
  //Column containing the date of the event
  var vDate = 1;
  //Column containing the start time
  var vStartTime = 2;
  //Column containing the end time
  var vEndTime = 3;
  //Column containing the location of the event
  var vLocation = 4;
  //Column containing the notification status
  var vNotify = 5;
  //Column containing the Guests
  var vGuests = 6;
  //Column containing the loaded status
  var vLoaded = 7;
  //Column containing the event ID
  var vEventID = 8;
  //spreadsheet variables
  var sheet = SpreadsheetApp.getActive().getSheetByName("Hangouts");
  var lastRow = sheet.getLastRow(); 
  var lastCol = sheet.getLastColumn(); 
  var range = sheet.getRange(2,1,lastRow,lastCol);
  var values = range.getValues();   
  var cDate = new Date();
  var eDate = new Date();
  var sTime = '';
  var eTime = '';

  //calendar variables
  var calendar = CalendarApp.getCalendarById('user.name@gmail.com');
   
  for (var i = 0; i < values.length; i++) {     
    // check to see if title is filled out
    if (values[i][vTitle].length > 0) {
       
      // check if it's been entered before          
      if (values[i][vLoaded] != 'y') {                       
         
        // create event https://developers.google.com/apps-script/class_calendarapp#createEvent
        var newEventTitle = values[i][vTitle];
        if (values[i][vTitle].length > 0) {
          // set up start date and time
          cDate = new Date(values[i][vDate]);
          sTime = values[i][vStartTime];
          var Sres = sTime.split(":");
          cDate.setHours(Sres[0], Sres[1]);
          // set up end date and time
          eDate = new Date(values[i][vDate]);
          eTime = values[i][vEndTime];
          var Eres = eTime.split(":");
          eDate.setHours(Eres[0], Eres[1]);
          // Logger.log([cDate, eDate]);
          var newEvent = calendar.createEvent(newEventTitle,
                                              cDate,
                                              eDate,
                                              {location: values[i][vLocation]});
        } else {
          newEvent = calendar.createAllDayEvent(newEventTitle, values[i][vDate], {location: values[i][vLocation]});
        }
        // Load Guests
        if (values[i][vGuests].length > 0) {
          loadGuests(newEvent, values[i][vGuests]);
        }
        // Check / Add / Delete Notifications
        /*
        addEmailReminder(minutesBefore)
        addPopupReminder(minutesBefore)
        addSmsReminder(minutesBefore)
        */
        
        // Check / Add / Delete Hangouts
        
        //get ID
        var newEventId = newEvent.getId();
        // mark as entered, store ID
        sheet.getRange(i+2,vLoaded+1).setValue('y');
        sheet.getRange(i+2,vEventID+1).setValue(newEventId);
      } 
    }
  } 
}

function loadGuests(vEvent, vGuests) {
  var guestList = vEvent.getGuestList();
  var inviteList = vGuests.split(',');//assuming your guestlist is comma separated
  if (inviteList.length>0){ 
    for(var n in inviteList){
      if (inGuestList(guestList,inviteList[n])) {
        continue;
      } else {
        vEvent.addGuest(inviteList[n]);
      }
    }
  }
  // delete guests not in list ??
}

function inGuestList(pGuestList,vemail) {
  var count=pGuestList.length;
  for(var i=0;i<count;i++) {
    if(pGuestList[i].getEmail() === vemail){return true;}
  }
  return false;
}

function formatDate(date) {
  if(Object.prototype.toString.call(date) !== '[object Date]') return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function checkSessions() {
  var vTitle = 0;
  var vDate = 1;
  var vStartTime = 2;
  var vEndTime = 3;
  var vLocation = 4;
  var vNotify = 5;
  var vGuests = 6;
  var vLoaded = 7;
  var vEventID = 8;
  var cDate = new Date();
  var eDate = new Date();

  var sheet = SpreadsheetApp.getActive().getSheetByName("Hangouts");
  var lastRow = sheet.getLastRow(); 
  var lastCol = sheet.getLastColumn(); 
  var range = sheet.getRange(2,1,lastRow,lastCol);
  var values = range.getValues();   
  var calendar = CalendarApp.getCalendarById('user.name@gmail.com');
   
  for (var i = 0; i < values.length; i++) {     
    if (values[i][vTitle].length > 0) {
      if (values[i][vLoaded] == 'y') {
        var session = values[i][vEventID]; 
        var eventId = session.toString().replace("@google.com","");// is that really the event Id ?
        cDate = new Date(values[i][vDate]);
        cDate.setHours(0, 0);
        eDate = new Date(values[i][vDate]);
        eDate.setHours(23, 59);
        var events = calendar.getEvents(cDate, eDate);
        for  ( var j in events ) {
          // var event = calendar.getEventSeriesById(session);
          var event = events[j];
          Logger.log([event.getId(), session, event.getStartTime()]);
          Logger.log('%s (%s)', event.summary, event.hangoutLink);
          if ( event.getId() == session ) {
            var guestList = event.getGuestList();
            for(var n in guestList){
              Logger.log([guestList[n].getEmail(), guestList[n].getGuestStatus()]);
            }
            var reminders = event.getEmailReminders();
            for(var n in reminders){
              Logger.log("Email Reminders - Minutes: " + reminders[n]);
            }
            var reminders = event.getSmsReminders();
            for(var n in reminders){
              Logger.log("SMS Reminders - Minutes: " + reminders[n]);
            }
            var reminders = event.getPopupReminders();
            for(var n in reminders){
              Logger.log("PopUp Reminders - Minutes: " + reminders[n]);
            }
          }
        }
      }
    }
  }
}

function moveHangoutLinks() {
  // https://gist.github.com/wheelertom/c3c7e2db34d290cf1ef1c9724ca17e48
    var calendarId = 'YOUR CALENDAR ID';
    var now = new Date();
    var events = Calendar.Events.list(calendarId, {
        timeMin: now.toISOString(),
        singleEvents: true,
        orderBy: 'startTime',
        maxResults: 10
    });
    if (events.items && events.items.length > 0) {
        for (var i = 0; i < events.items.length; i++) {
            var event = events.items[i];
            var d = event.description;
            if (!d)
                d = '';
            if (event.hangoutLink && (d.indexOf('Hangout: ')== -1)){
                //Logger.log (event.summary + ' - ' + event.hangoutLink + ' - ' + event.description);
                event.description = 'Hangout: ' + event.hangoutLink + '\n\n' + d;
                Calendar.Events.update(event, calendarId, event.id);
            }

        }
    } else {
        Logger.log('No events found.');
    }
}

function delete_events()
{
    var calendarName = 'Test';
    // for month 0 = Jan, 1 = Feb etc
    // below delete from Jul 13 2020 to Jul 18 2020
    var fromDate = new Date("2020-07-13"); 
    var toDate = new Date("2020-07-18");
    var calendar = CalendarApp.getCalendarsByName(calendarName)[0];
    var events = calendar.getEvents(fromDate, toDate);
    for(var i=0; i<events.length;i++){
        var ev = events[i];
        if(ev.getTitle()=="EventX" & ev.getCreators()=="xyz@gmail.com"){
        // show event name in log
        Logger.log(ev.getTitle()); 
        ev.deleteEvent();
     }
}
}

function delete_eventsv2() {
  var calendarName = 'Test';
  var myEmail = "YOUR EMAIL";
  var myTitle = "Hello";
  // for month 0 = Jan, 1 = Feb etc
  // below delete from now to Jul 18 2020
  var now = new Date(); 
  var toDate = new Date(2020,6,18,0,0,0);
  var calendar = CalendarApp.getCalendarsByName(calendarName)[0];
  var events = calendar.getEvents(now, toDate);
  for(var i=0; i<events.length;i++){
    var ev = events[i];
    // show event name in log
    Logger.log(ev.getTitle()); 
    var creators = ev.getCreators();
//check if you are the calendar creator and the event title matches
    if(creators.indexOf(myEmail) >-1 && ev.getTitle() == myTitle){
      ev.deleteEvent();
    }
  }
}


function sync() {
  var calendarId = 'myemailid@gmail.com'; // Please set your calendar ID.

  var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1');
  var calendar = CalendarApp.getCalendarById(calendarId);
  var startRow = 2;  // First row from which data should process > 2 exempts my header row
  var numRows = sheet.getLastRow();   // Number of rows to process
  var numColumns = sheet.getLastColumn();
  var dataRange = sheet.getRange(startRow, 1, numRows - 1, numColumns);
  var data = dataRange.getValues();
  var done = "Done";  // It seems that this is not used.
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var name = row[1];  //Name of Guest
    var place = row[4];  //Add2
    var room = row[9]; //Room Number    
    var inDate = new Date(row[10]);  //Check-In Date
    var outDate = new Date(row[11]); //Check-Out Date
    var check1 = row[23];  //Booked/Blocked/Cancelled
    var check2 = row[24]; //Event created and EventID (iCalUID) populated 
    
    // I modified below script.
    if (check1 != "Cancelled" && check2 == "") {
      var currentCell = sheet.getRange(startRow + i, numColumns);
      var event = calendar.createEvent(room, inDate, outDate, {
        description: 'Booked by: ' + name + ' / ' + place + '\nFrom: ' + inDate + '\nTo: ' + outDate
      });
      var eventId = event.getId();
      currentCell.setValue(eventId);
    } else if (check1 == "Cancelled" && check2 != "") {
      var status = Calendar.Events.get(calendarId, check2.split("@")[0]).status;
      if (status != "cancelled") {
        calendar.getEventById(check2).deleteEvent();
      }
    }
  }
}
