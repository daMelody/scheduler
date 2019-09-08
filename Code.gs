var time_zone = 'T16';  // cell where time zone is specified
var calendar_list = 'S7:S13';  // cell range where calendars are listed
var hours_list = 'T7:T13';  // cell range where total hours are shown
var last_column = 'P3:P50';  // last column of calendar
var bottom_row = 50;
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
function inputEvent(cal, name, start, finish, loc) {
  var calendar = CalendarApp.getCalendarsByName(cal)[0];
  if (calendar.getEvents(start, finish)[0] != undefined &&
      calendar.getEvents(start,finish)[0].getTitle() == name) {
    Logger.log('Event already exists');
  } else {
    var reminder = SpreadsheetApp.getUi().prompt('Reminder before event (minutes)?').getResponseText();
    calendar.createEvent(name, start, finish, {location: loc}).addEmailReminder(reminder);
  }
}

/** creates Date objects for calendar input */
function newDate(zone, date, time) {
 return new Date(date + ' ' + time + ' ' + zone);
}

/** gets beginning and end cells from active range */
function getCells(sheet) {
  var begin, end; var rangeNotation = [];
  var ranges = sheet.getActiveRangeList().getRanges();
  for (var i=0; i < ranges.length; i++) { // build array of ranges in A1Notation
    rangeNotation[i] = ranges[i].getA1Notation();
  }
  if (rangeNotation.length > 1) {
    // multiple days
    begin = rangeNotation[0].toString().split(':')[0];
    end = rangeNotation[rangeNotation.length-1].toString().split(':')[1];
  } else if (rangeNotation[0].toString().indexOf(':') > 0) {
    // multiple cells
    begin = rangeNotation[0].toString().split(':')[0];
    end = rangeNotation[0].toString().split(':')[1];
  } else {
    // only one cell
    begin = rangeNotation[0].toString();
    end = rangeNotation[0].toString();
  }
  return [begin, end];
}

/** traverses spreadsheet to return corresponding time of an "event cell" */
function getTime(sheet, cell) {
  var range = "A" + cell.substring(1);
  return sheet.getRange(range.toString()).getValue().toString();
}

/** produces array of Date objects which correspond to begin, end of event */
function getDates(sheet, cells) {
  var beginDate, endDate;
  beginDate = cells[0].substring(0,1) + "1";
  beginDate = sheet.getRange(beginDate.toString()).getValue().toString().substring(0,15);
  endDate = cells[1].substring(0,1) + "1";
  endDate = sheet.getRange(endDate.toString()).getValue().toString().substring(0,15);
  // find times, move end time up by 30mins
  var endRow = parseInt(cells[1].substring(1)) + 1;
  if (endRow > bottom_row) { // if event ends at 12am, increment dates[1] by 1 day
    var today = endDate.substring(8,10);
    var tomorrow = parseInt(today) + 1;
    endDate = endDate.replace(today, tomorrow.toString());
    endRow = 3;
  }
  cells[1] = cells[1].replace(cells[1].substring(1), endRow);
  var start = getTime(sheet, cells[0]);   // start time
  var finish = getTime(sheet, cells[1]);  // finish time
  // making the Date objects
  var zone = sheet.getRange(time_zone).getValue();  // get time zone
  start = newDate(zone, beginDate, start);
  finish = newDate(zone, endDate, finish);
  return [start, finish]
}

/** inputs a highlighted range as an event into
 * a calendar selected by background color */
function calendarize() {
  var sheet = SpreadsheetApp.getActive();
  var cells = getCells(sheet);
  // process info from spreadsheet
  var name = sheet.getRange(cells[0]).getValue();
  var dates = getDates(sheet, cells);
  // select a calendar for the event
  var cal = calendarSelector(sheet.getRange(calendar_list).getValues(),
                             sheet.getRange(calendar_list).getBackgrounds(),
                             sheet.getRange(cells[0]).getBackground());

  // get location
  var location = SpreadsheetApp.getUi().prompt('Event location:').getResponseText();
  // actually put the event into the calendar
  inputEvent(cal, name, dates[0], dates[1], location);
}

/** deletes first day and moves all events over to the left */
function sheetify() {
  var sheet = SpreadsheetApp.getActive();
  var shifted = sheet.getRange(copied);
  var destination = sheet.getRange(pasted);
  shifted.copyTo(destination);
  sheet.getRange(last_column).clear();
}

function halvsies(d2array) {
  for (var i=0; i < d2array.length; i++) {
    d2array[i][0] = d2array[i][0] * 0.5;
  }
  return d2array;
}

function calTotals() {
  var sheet = SpreadsheetApp.getActive()
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
