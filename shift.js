//For google apps script, read calendar events to sheet and allow user to add attendances to event titles

var spreadsheet = SpreadsheetApp.getActiveSheet();
var yearCell = "N1";
var monthCell = "N2";
var calendarIdCell = "N3"
var eventCountCell = "N5"
var calendarId = spreadsheet.getRange(calendarIdCell).getValue();
var eventCal = CalendarApp.getCalendarById(calendarId);
var initalRow = 4;
var dateCol = 1;
var titleCol = 2;
var idCol = 8;
var startTCol = 9;
var endTCol = 10;
var noOfPerson = 5;
var names = []
for (let i = 0; i < noOfPerson; i++) {
  names.push(spreadsheet.getRange(3, 3+i).getValue())
};

function initialise() {
  
  var monthNumber = spreadsheet.getRange(monthCell).getValue();
  var year = spreadsheet.getRange(yearCell).getValue();
  var month = "";
  var date = new Date();
  date.setFullYear(year);
  date.setMonth(monthNumber -1);
  var firstDay = new Date(date.getFullYear(), monthNumber - 1, 1);
  var lastDay = new Date(date.getFullYear(), monthNumber, 1)
  var title = spreadsheet.getRange(1,1);

  switch(monthNumber) {
    case 1:
      month = "JAN"
      break;
    case 2:
      month = "FEB"
      break;
    case 3:
      month = "MAR"
      break;
    case 4:
      month = "APR"
      break;
    case 5:
      month = "MAY"
      break;
    case 6:
      month = "JUN"
      break;
    case 7:
      month = "JUL"
      break;
    case 8:
      month = "AUG"
      break;
    case 9:
      month = "SEP"
      break;
    case 10:
      month = "OCT"
      break;
    case 11:
      month = "NOV"
      break;
    case 12:
      month = "DEC"
      break;
  }

  var confirm = Browser.msgBox('You will initialise the sheet.','Execute script?', Browser.Buttons.OK_CANCEL);
  if(confirm!=='ok'){
    return null;
  };

  title.setValue(month + " " + year);
  var events = eventCal.getEvents(firstDay, lastDay);
  
  spreadsheet.getRange("A4:B").setValue("");
  spreadsheet.getRange("H4:J").setValue("");

  var row = initalRow;

  events.forEach(function (item) {   
    if ((item.getTitle().search("HKP") !== -1) && (item.getTitle().search("取消") < 0)) {
      spreadsheet.getRange(row, idCol).setValue(item.getId());
      spreadsheet.getRange(row, titleCol).setValue(item.getTitle());
      var startT = item.getStartTime();
      var endT = item.getEndTime();
      spreadsheet.getRange(row, dateCol).setValue(startT.toDateString().slice(4,10));
      spreadsheet.getRange(row, startTCol).setValue(startT);
      spreadsheet.getRange(row, endTCol).setValue(endT);
    row += 1;
    }
  });

}

function updateAttendance() {
  var n = spreadsheet.getRange(eventCountCell).getValue() - 1;
  var row = initalRow;

  var confirm = Browser.msgBox('You will uploud attendance.','Execute script?', Browser.Buttons.OK_CANCEL);
  if(confirm!=='ok'){
    return null;
  };

  for (let i = 0; i < n; i++) {
    var id = spreadsheet.getRange(row, idCol).getValue();

    var event = eventCal.getEventById(id);
    if (event === null) {
      var confirm = Browser.msgBox('It seems that one or more event has an missing id, please contact IT support.', Browser.Buttons.OK_CANCEL);
      return null;
    };

    if (event != null) {
      var title = spreadsheet.getRange(row, titleCol).getValue().toString().split('[')[0];
      var attendance =[]
      var noOneAttend = true
      for (let x = 0; x < noOfPerson; x++) {
        input = spreadsheet.getRange(row, 3+x).getValue()
        noOneAttend = ((input == "") && (noOneAttend == true)) ? true : false 
        attendance.push(input)
      }
      if (noOneAttend == false) {
        title = title + " [";
        for (let x = 0; x < noOfPerson; x++) {
          if (attendance[x] != "") {
            title = title + names[x] + ", "
          }
        }
        title = title + "]";
        event.setTitle(title);
      };
    };
    row += 1;
  };
}

function deleteAttendance () {

  var confirm = Browser.msgBox('You will DETELE attendance','Execute script?', Browser.Buttons.OK_CANCEL);
  if(confirm!=='ok'){
    return null;
  };

  var n = spreadsheet.getRange(eventCountCell).getValue() - 1;
  var row = initalRow;
  for (let i = 0; i < n; i++) {
    var id = spreadsheet.getRange(row, idCol).getValue();

    var event = eventCal.getEventById(id);
    if (event === null) {
      var confirm = Browser.msgBox('It seems that one or more event has an missing id, please contact IT support.', Browser.Buttons.OK_CANCEL);
      return null;
    };

    if (event != null) {
      var title = spreadsheet.getRange(row, titleCol).getValue().toString().replace(/\[.*\]/, '');
      event.setTitle(title);
    };
    row += 1;
  };
}
