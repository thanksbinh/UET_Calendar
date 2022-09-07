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

  // Clear End Date 
  this.clearEndDate = function() {
    sheet.getRange('B3').clearContent();;
  }
}

function gCalendar(infoSh, scheduleSh) {
  
  var infoSheet = infoSh.getSheet();
  var calendarId = infoSheet.getRange('B1').getValue().toString(); 
  var calendar;
  var startDate = infoSheet.getRange('B2').getValue();
  var endDate = infoSheet.getRange('B3').getValue();

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

    // validate start date
    startDate.setDate(startDate.getDate() + (1 + 7 - startDate.getDay()) % 7);
    infoSheet.getRange('B2').setValue(startDate);

    // init endDate
    if (endDate == "") {
      endDate = new Date(startDate);
      endDate.setDate(endDate.getDate() + 15*7);
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

  // Add "Tuần 1 .. 15"
  this.addWeekNumb = function() {

    //First and Last day of the week
    var firstDay = new Date(startDate);
    var lastDay = new Date(startDate);
    lastDay.setDate(firstDay.getDate() + 7);

    // Add 15 times
    for (i=0; i<15; i++) {
      calendar.createEvent("Tuần " + (i+1).toString(), firstDay, lastDay);

      firstDay.setDate(firstDay.getDate() + 7);
      lastDay.setDate(lastDay.getDate() + 7);
    }
  };

  // Find weeks noticed
  this.findWeeks = function(row) {
    var weeks = [];

    var afterBracket = row[HP].split('(')[1];

    // All weeks
    if (afterBracket == null) return weeks;
    if (afterBracket.match(/\d+/g) == null) return weeks;

    // Weeks between "đến tuần"
    if (afterBracket.search("đến tuần") != -1) {
      var numbs = afterBracket.match(/\d+/g);
      if (numbs.length != 2) {
        Logger.log("Error: Weird week format");
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
        Logger.log("Error: Weird week format");
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

  // Export from sheet to calendar
  this.exportToCalendar = function() {

    // Add subject if Dang ky
    for (var i=0; i<table.length; i++) {
      var rowNumb = i + STARTROW;
      var row = table[i];

      if (row[DK] == "") {
        if (row[ID] != "") {
          var events = calendar.getEventSeriesById(row[ID].toString());
          events.deleteEventSeries();

          scheduleSheet.getRange(rowNumb, ID+1).clearContent();
        }
        continue;
      }
      else {
        if (row[ID] != "") continue;
        else if (row[HP] == "" || row[Thu] == "" || row[Tiet] == "") {
          SpreadsheetApp.getUi().alert("Invalid selection");
          continue;
        }
      }

      var title = row[Nhom] + " - " + row[HP];
      var des = "Mã lớp học phần: " + row[MLHP] + "\nTín chỉ: " + row[TC] + "\nGiảng viên/ Trợ giảng: " + row[GVTG] + "\nGhi chú: " + row[GC];
      var loca = row[GD];

      var startTime = new Date(startDate);
      var endTime = new Date(startDate);

      var weeks = this.findWeeks(row);
      var dayOfWeek = parseInt(row[Thu]);

      var startHour = parseInt(row[Tiet].toString().split("-")[0]) + 6;
      var endHour = parseInt(row[Tiet].toString().split("-")[1]) + 7;

      startTime.setDate(startTime.getDate()+dayOfWeek-2);
      endTime.setDate(endTime.getDate()+dayOfWeek-2);

      startTime.setHours(startHour);
      endTime.setHours(endHour);

      var valid = true;
      // Add something here . . .
      
      if (valid) {
        if(weeks.length == 0) {
          var eventSeries = calendar.createEventSeries( title, 
                                      startTime, 
                                      endTime, 
                                      CalendarApp.newRecurrence().addWeeklyRule().times(15), 
                                      {description: des, location: loca});

          Logger.log(title + " " + startTime + " " + endTime);
          scheduleSheet.getRange(rowNumb, ID+1).setValue(eventSeries.getId());
        }
        else {
          startTime.setDate(startTime.getDate()+(weeks[0]-1)*7);
          endTime.setDate(endTime.getDate()+(weeks[0]-1)*7);
          
          var eventSeries = calendar.createEventSeries( title, 
                                      startTime, 
                                      endTime, 
                                      CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeeks(weeks).times(weeks.length), 
                                      {description: des, location: loca});

          Logger.log(title + " " + startTime + " " + endTime + " " + weeks);
          scheduleSheet.getRange(rowNumb, ID+1).setValue(eventSeries.getId());
        }
      }
    }
  };
}

/* --------------------------------------------------------------------------------- */

function onOpen() {
  "use strict";
  var menuEntries = [{
    name: "Make Calendar",
    functionName: "make"
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

function make() {
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
  infoSh.clearEndDate();
  scheduleSh.clearID();
}

function clearSelection() {
  scheduleSh.clearSelection();
}