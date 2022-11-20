const STARTROW = 2;
const DK = 0, MHP = 1, HP = 2, TC = 3, MLHP = 4, SSV = 5, GVTG = 6, Thu = 7, Tiet = 8, GD = 9, Nhom = 10, GC = 11, ID = 12;

function gSheet(name) {

  var sheet = SpreadsheetApp.getActive().getSheetByName(name);

  this.getSheet = function() {
    return sheet;
  }

  // Clear sheet's "Dang ky" options
  this.clearSelection = function() {
    sheet.getRange("A" + STARTROW.toString() + ":A").clearContent();
  };

  // Clear ID column
  this.clearID = function() {
    sheet.getRange("M" + STARTROW.toString() + ":M").clearContent();
  };

  // Clear Calendar ID 
  this.clearCalID = function() {
    sheet.getRange('B1').clearContent();;
  };
}

function gCalendar(infoSh, scheduleSh) {
  
  var infoSheet = infoSh.getSheet();
  var calendarId = infoSheet.getRange('B1').getValue().toString(); 
  var calendar;
  var startDate = infoSheet.getRange('B2').getValue();
  var endWeek = infoSheet.getRange('B3').getValue();

  var scheduleSheet = scheduleSh.getSheet();
  var table = scheduleSheet.getRange("A" + STARTROW.toString() + ":M").getValues();

  // init calendar ID
  this.initCal = function() {

    // init calendar ID
    if (calendarId == "") {
      calendar = CalendarApp.createCalendar('UET Calendar', {summary: 'Made by ThanksBinh', hidden: false, selected: true, color: "#7cd2fd"});
      infoSheet.getRange('B1').setValue(calendar.getId());
    }
    else {
      calendar = CalendarApp.getCalendarById(calendarId);
    }

    // validate start date -> Monday
    startDate.setDate(startDate.getDate() + (1 + 7 - startDate.getDay()) % 7);
    infoSheet.getRange('B2').setValue(startDate);

    // init endDate
    if (endWeek == "") {
      endWeek = 15;
      infoSheet.getRange('B3').setValue(endDate);
    }
  };

  // Delete calendar and clear sheet id
  this.deleteCal = function() {

    if (infoSheet.getRange('B1').getValue().toString() == "") {
      SpreadsheetApp.getUi().alert("No Calendar ID found");
      return;
    }

    CalendarApp.getCalendarById(calendarId).deleteCalendar();
  };

  // Add "Tuần 1 .. endWeek"
  this.addWeekNumb = function() {

    //First and Last day of the week
    var firstDay = new Date(startDate);
    var lastDay = new Date(startDate);
    lastDay.setDate(firstDay.getDate() + 7);

    // Add endWeek times
    for (i=0; i<endWeek; i++) {
      calendar.createEvent("Tuần " + (i+1).toString(), firstDay, lastDay);

      firstDay.setDate(firstDay.getDate() + 7);
      lastDay.setDate(lastDay.getDate() + 7);
    }
  };

  // Find weeks noticed
  this.findWeeks = function(row) {
    var weeks = [];

    var afterBracket = row[HP].split('(')[1];

    // Every weeks
    if (afterBracket == null) return weeks;
    if (afterBracket.match(/\d+/g) == null) return weeks;

    // Weeks between "đến tuần"
    if (afterBracket.search("đến tuần") != -1) {
      var numbs = afterBracket.match(/\d+/g);
      if (numbs.length != 2) {
        SpreadsheetApp.getUi().alert("Error: Weird week format");
      }
      else {
        for (var j=parseInt(numbs[0]); j<=parseInt(numbs[1]); j++) weeks.push(j);
      }
      return weeks;
    }
    
    // Week between '-' or ','
    var arr = afterBracket.split(' ');
    for (var i=0; i<arr.length;i++) {
      if (!/\d/.test(arr[i])) continue;

      var numbs = arr[i].match(/\d+/g);

      if (numbs.length > 2) {
        for (var j=0; j<numbs.length; j++) weeks.push(numbs[j]);
      }
      else if (arr[i].search('-') != -1) {
        for (var j=parseInt(numbs[0]); j<=parseInt(numbs[1]); j++) weeks.push(j);
      }
      else {
        weeks.push(parseInt(numbs[0]));
      }
    }

    return weeks;
  };

  this.deleteEventSeries = function(row, rowNumb) {
    var events = calendar.getEventSeriesById(row[ID].toString());
    events.deleteEventSeries();

    scheduleSheet.getRange(rowNumb, ID+1).clearContent();
  };

  this.setStartTime = function(row) {
    var startTime = new Date(startDate);
    var dayOfWeek = parseInt(row[Thu])-2;
    var startHour = parseInt(row[Tiet].toString().split("-")[0]) + 6;
    startTime.setDate(startTime.getDate()+dayOfWeek);
    startTime.setHours(startHour);
    return startTime;
  };

  this.setEndTime = function(row) {
    var endTime = new Date(startDate);
    var dayOfWeek = parseInt(row[Thu])-2;
    var endHour = parseInt(row[Tiet].toString().split("-")[1]) + 7;
    endTime.setDate(endTime.getDate()+dayOfWeek);
    endTime.setHours(endHour);
    return endTime;
  };

  // Export from sheet to calendar
  this.exportToCalendar = function() {

    // Add subject if Dang ky and ID == ''
    for (var i=0; i<table.length; i++) {
      var rowNumb = i + STARTROW;
      var row = table[i];

      if ((row[DK] == "") == (row[ID] == ""))
        continue;
      if (row[DK] == "" && row[ID] != "") {
        this.deleteEventSeries(row, rowNumb);
        continue;
      }
      if (row[HP] == "" || row[Thu] == "" || row[Tiet] == "") {
        SpreadsheetApp.getUi().alert("Invalid selection");
        continue;
      }

      var title = row[Nhom] + " - " + row[HP];
      var des = "Mã lớp học phần: "+row[MLHP] + "\nTín chỉ: "+row[TC] + "\nGiảng viên/ Trợ giảng: "+row[GVTG] + "\nGhi chú: "+row[GC];
      var loca = row[GD];
      var weeks = this.findWeeks(row);
      var startTime = this.setStartTime(row);
      var endTime = this.setEndTime(row);

      var times = 0;
      if(weeks.length == 0 && row[Nhom] == "CL") {
        times = endWeek;
      }
      else if (weeks.length == 0 && row[Nhom] != "CL") {
        times = endWeek-1;
        startTime.setDate(startTime.getDate()+7);
        endTime.setDate(endTime.getDate()+7);
      }
      else {
        times = weeks[weeks.length-1]-weeks[0]+1;
        startTime.setDate(startTime.getDate()+(weeks[0]-1)*7);
        endTime.setDate(endTime.getDate()+(weeks[0]-1)*7);
      }

      var eventSeries = calendar.createEventSeries( title, 
                                                    startTime, 
                                                    endTime, 
                                                    CalendarApp.newRecurrence().addWeeklyRule().times(times),  //.onlyOnWeeks(weeks)
                                                    {description: des, location: loca});
      scheduleSheet.getRange(rowNumb, ID+1).setValue(eventSeries.getId());
      Logger.log(title + " " + startTime + " " + endTime + " " + weeks);

      // Remove redundant events
      let weekNumb = weeks[0];
      var removeStartTime = startTime;
      var removeEndTime = endTime;

      for (let i = 0; i < weeks.length; i++) {
        if (weekNumb != weeks[i]) {
          i--;
          var events = calendar.getEvents(removeStartTime,removeEndTime);
          for (let event of events) {
            if (event.getTitle() == title) {
              Logger.log("Delete event: " + event.getTitle());
              event.deleteEvent();
            }
          }
        }

        weekNumb++;
        removeStartTime.setDate(removeStartTime.getDate()+7);
        removeEndTime.setDate(removeEndTime.getDate()+7);
      }
    }
  };
}

/* --------------------------------------------------------------------------------- */

function onOpen() {
  "use strict";
  var menuEntries = [{
    name: "Make Calendar",
    functionName: "makeCalendar"
  }, {
    name: "Add Week Number",
    functionName: "addWeekNumb"
  }, {
    name: "Delete Calendar",
    functionName: "deleteCalendar"
  }, {
    name: "Clear Selection",
    functionName: "clearSelection"
  }], activeSheet;

  activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  activeSheet.addMenu('Code', menuEntries);
}

var infoSh = new gSheet("info");
var scheduleSh = new gSheet(infoSh.getSheet().getRange('B4').getValue().toString());
var gCal = new gCalendar(infoSh, scheduleSh);

function makeCalendar() {
  gCal.initCal();
  gCal.exportToCalendar();
}

function addWeekNumb() {
  gCal.initCal();
  gCal.addWeekNumb();
}

function deleteCalendar() {
  gCal.deleteCal();
  infoSh.clearCalID();
  scheduleSh.clearID();
}

function clearSelection() {
  scheduleSh.clearSelection();
}
