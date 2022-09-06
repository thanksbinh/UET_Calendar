function gSheet() {

  var sheet = SpreadsheetApp.getActiveSheet();
  const STARTROW = 2;
  var table = sheet.getRange("A" + STARTROW.toString() + ":M").getValues();

  // Clear sheet's "Dang ky" options
  this.clearSheet = function() {
    for (var i=0; i<table.length; i++) {
      var rownNumb = i + STARTROW;
      sheet.getRange(rownNumb, 1).setValue("");
    }
  };
}

function gCalendar() {

  var infoSheet = SpreadsheetApp.getActive().getSheetByName('Info');
  var calendarId = infoSheet.getRange('B1').getValue().toString(); 
  var calendar;
  if (calendarId == "") {
    calendar = CalendarApp.createCalendar('UET Calendar', {summary: 'Made by ThanksBinh', hidden: false, selected: true});
    infoSheet.getRange('B1').setValue(calendar.getId());
  }
  else {
    calendar = CalendarApp.getCalendarById(calendarId);
  }

  var startDate = infoSheet.getRange('B2').getValue();
  var endDate = infoSheet.getRange('B3').getValue();
  if (endDate == "") {
    endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + 15*7);
    infoSheet.getRange('B3').setValue(endDate);
  }

  var sheet = SpreadsheetApp.getActiveSheet();
  const STARTROW = 2;
  var table = sheet.getRange("A" + STARTROW.toString() + ":M").getValues();
  const DK = 0, MHP = 1, HP = 2, TC = 3, MLHP = 4, SSV = 5, GVTG = 6, Thu = 7, Tiet = 8, GD = 9, Nhom = 10, GC = 11, ID = 12;

  // Delete calendar and clear sheet id
  this.deleteCal = function() {
    calendar.deleteCalendar();
    sheet.getRange('B1').setValue('');
    
    for (var i=0; i<table.length; i++) {
      var rownNumb = i + STARTROW;
      sheet.getRange(rownNumb, ID+1).setValue("");
    }
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
      if (table[i][DK] == "") {
        if (table[i][ID] != "") {
          var events = calendar.getEventSeriesById(table[i][ID].toString());
          events.deleteEventSeries();

          sheet.getRange(rowNumb, ID+1).setValue("");
        }
        continue;
      }
      else {
        if (table[i][ID] != "") continue;
      }

      var row = table[i];

      var title = row[Nhom] + " - " + row[HP];
      var des = "Mã lớp học phần: " + row[MLHP] + "\nTín chỉ: " + row[TC] + "\nGiảng viên/ Trợ giảng: " + row[GVTG];
      if (row[GC] != "") des += "\nGhi chú: " + row[GC];
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
      if (valid) {
        if(weeks.length == 0) {
          var eventSeries = calendar.createEventSeries( title, 
                                      startTime, 
                                      endTime, 
                                      CalendarApp.newRecurrence().addWeeklyRule().times(15), 
                                      {description: des, location: loca});

          Logger.log(title + " " + startTime + " " + endTime);
          sheet.getRange(rowNumb, ID+1).setValue(eventSeries.getId());
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
          sheet.getRange(rowNumb, ID+1).setValue(eventSeries.getId());
        }
      }
    }
  };
}

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
    name: "Clear Sheet",
    functionName: "clearSheet"
  }], activeSheet;

  activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  activeSheet.addMenu('Code', menuEntries);
}

function make() {
  (new gCalendar).exportToCalendar();
}

function addWeekNumb() {
  (new gCalendar).addWeekNumb();
}

function deleteCalendar() {
  (new gCalendar).deleteCal();
}

function clearSheet() {
  (new gSheet).clearSheet();
}

