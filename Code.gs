var sheet = SpreadsheetApp.getActiveSheet();  // define the sheet
var time_zone = 'T16';  // cell where time zone is specified
var calendar_list = 'S7:S13';  // cell range where calendars are listed
var hours_list = 'T7:T13';  // cell range where total hours are shown
var last_column = 'P3:P50';  // last column of calendar
var copied = 'D3:P50';  // region of calendar to copy
var pasted = 'C3:O50';  // one column to the left of COPIED
var calendar_region = 'C3:P50';  // where events are entered and displayed

/** the main function */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // options update calendar or spreadsheet
  ui.createMenu('Calendar-Connect')
    .addItem('Update Calendar', 'calendarize')
    .addSeparator()
    .addItem('Update Sheets', 'sheetify')
    .addToUi();
}

/** selects calendar by color */
function calendarSelector(calNames, colorArray, cellColor) {
  var calendar;
  switch(cellColor) {
    case colorArray[0][0]: calendar = calNames[0][0];
      break;
    case colorArray[1][0]: calendar = calNames[1][0];
      break;
    case colorArray[2][0]: calendar = calNames[2][0];
      break;
    case colorArray[3][0]: calendar = calNames[3][0];
      break;
    case colorArray[4][0]: calendar = calNames[4][0];
      break;
    case colorArray[5][0]: calendar = calNames[5][0];
      break;
    case colorArray[6][0]: calendar = calNames[6][0];
      break;
    default: Logger.log('Invalid calendar');
  }
  return calendar;
}

/** handles the actual calendar selection and input */
function inputEvent(cal, name, start, finish) {
  //Logger.log("cal: " + cal);
  Logger.log("name: " + name);
  /*Logger.log("start: " + start);
  Logger.log("finish: " + finish); */
  var calendar = CalendarApp.getCalendarsByName(cal)[0];
  //Logger.log("calendar: " + calendar.getName());
  if (calendar.getEvents(start, finish)[0] != undefined &&
      calendar.getEvents(start,finish)[0].getTitle() == name) {
    Logger.log('Event already exists');
  } else {
    calendar.createEvent(name, start, finish);
  }
}

/** creates Date objects for calendar input */
function newDate(zone, date, time) {
  var date = new Date(date + ' ' + time + ' ' + zone);
  Logger.log("date: " + date);
  return date;
}


/** traverses spreadsheet to return corresponding time of an "event cell" */
function getTime(sheet, cell) {
  var range = "A" + cell.substring(1);
  return sheet.getRange(range.toString()).getValue().toString();
}


function getDate(sheet, cell) {
  var range = cell.substring(0,1) + "1";
  var day = sheet.getRange(range.toString()).getValue().toString().substring(0,15);
  Logger.log("day: " + day);
  return day;
}


/** iterates thru the highlighted range and inputs events into
 * calendars selected by background color */
function calendarize() {
  var range = sheet.getRange(calendar_region);
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  var name = range.getCell(1,1).getValue();
  var startT = getTime(sheet, "A3");
  var startD = getDate(sheet, "C1");
  var color = range.getCell(1,1).getBackground();
  var newColor = color;
  var endT, endD, calendar;
  // iterate through the eligible range
  for (var col=0; col < range.getLastColumn()-2; col++) {
    for (var row=0; row < range.getLastRow()-2; row++) {
      var newName = values[row][col];
      color = newColor; // color takes on previous value
      newColor = backgrounds[row][col];
      // check for event change and input into calendar if changed 
      if (name != newName || color != newColor) {
        endT = getTime(sheet, range.getCell(row+1,col+1).getA1Notation());
        endD = getDate(sheet, range.getCell(row+1,col+1).getA1Notation());
        calendar = calendarSelector(sheet.getRange(calendar_list).getValues(),
                                    sheet.getRange(calendar_list).getBackgrounds(),
                                    color);
        var zone = sheet.getRange(time_zone).getValue();
        // send packaged event to Google Calendar
        var begin = newDate(zone,startD,startT);
        var end = newDate(zone,endD,endT);
        if (name != "") { inputEvent(calendar,name,begin,end); }
        // reassign variables
        name = newName;
        color = newColor;
        startT = getTime(sheet, range.getCell(row+1,col+1).getA1Notation());
        startD = getDate(sheet, range.getCell(row+1,col+1).getA1Notation()); 
      }
    }
  }
}


/** deletes first day and moves all events over to the left */
function sheetify() {
  var shifted = sheet.getRange(copied);
  var destination = sheet.getRange(pasted);
  shifted.copyTo(destination);
  sheet.getRange(last_column).clear();
  var dayRange = sheet.getRange("C2:P2");
  var names = getWeekDays(dayRange);
  dayRange.setValues([names]);
}


function getWeekDays(dayRange) {
  var days = dayRange.getValues()[0];
  // iterate through days array and reassign
  for (var i=0; i < days.length; i++) {
    Logger.log("day: " + days[i]);
    switch (days[i]) {
      case 1: days[i] = "Sunday";
        break;
      case 2: days[i] = "Monday";
        break;
      case 3: days[i] = "Tuesday";
        break;
      case 4: days[i] = "Wednesday";
        break;
      case 5: days[i] = "Thursday";
        break;
      case 6: days[i] = "Friday";
        break;
      case 7: days[i] = "Saturday";
        break;
      default: Logger.log("Invalid day?");
    }
  }
  return days;
}


function halvsies(d2array) {
  for (var i=0; i < d2array.length; i++) {
    d2array[i][0] = d2array[i][0] * 0.5;
  }
  return d2array;
}

function calTotals() {
  var colors = sheet.getRange(calendar_region).getBackgrounds();
  var calendars = sheet.getRange(calendar_list).getBackgrounds();
  var counter = [[0],[0],[0],[0],[0],[0],[0]];
  // iterate through spreadsheet
  for (var i=0; i < colors.length; i++) {
    for (var j=0; j < colors[i].length; j++) {
      if (colors[i][j] != '#ffffff') { // if background is white, ignore
        switch (colors[i][j]) {  // compare color with calendar color
          case calendars[0][0]: counter[0][0]++;
            break;
          case calendars[1][0]: counter[1][0]++;
            break;
          case calendars[2][0]: counter[2][0]++;
            break;
          case calendars[3][0]: counter[3][0]++;
            break;
          case calendars[4][0]: counter[4][0]++;
            break;
          case calendars[5][0]: counter[5][0]++;
            break;
          case calendars[6][0]: counter[6][0]++;
            break;
          default: Logger.log("Invalid calendar");
        }
      }
    }
  }
  // print updated sums to spreadsheet
  sheet.getRange(hours_list).setValues(halvsies(counter));
}

// calculating total hours by event title and calendar
function onEdit(e) {
  calTotals();  // calculate calendar totals
}
