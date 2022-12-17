const STARTROW = 2;
const DK = 0, MHP = 1, HP = 2, TC = 3, MLHP = 4, SSV = 5, GVTG = 6, Thu = 7, Tiet = 8, GD = 9, Nhom = 10, GC = 11, ID = 12;

function gSheet(name) {

  let sheet = SpreadsheetApp.getActive().getSheetByName(name);

  this.getSheet = function() {
    return sheet;
  }

  // Select all classes in registerList
  this.selectClasses = function(registerList) {
    
    for (let i = 0; i < registerList.length; i+=2) {
      let check = false;

      let object = registerList[i];
      let group = registerList[i+1];

      let foundCells = sheet.createTextFinder(object).findAll();
    
      // Loop through table to select all elements of this subject
      for (let j = 0; j < foundCells.length; j++) {
        if (foundCells[j].getColumn() != 5) continue;
        if (foundCells[j].getValue().split('(')[0] != object) continue;

        let thisGroup = sheet.getRange("K" + foundCells[j].getRow()).getValue();
        if (thisGroup == group || (group != 'CL' && thisGroup.toUpperCase() == 'CL')) {
          sheet.getRange("A" + foundCells[j].getRow()).setValue('1');
          check = true;
        }
      }

      if (check == false) {
        SpreadsheetApp.getUi().alert("Data missing " + object + ", group " + group);
      }
    }
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
    sheet.getRange('B1').clearContent();
  };
}

function gCalendar(infoSh, scheduleSh) {
  
  let infoSheet = infoSh.getSheet();
  let calendarId = infoSheet.getRange('B1').getValue().toString(); 
  let calendar;
  let startDate = infoSheet.getRange('B2').getValue();
  let endWeek = infoSheet.getRange('B3').getValue();

  let scheduleSheet = scheduleSh.getSheet();
  let table;

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

    // validate start date = Monday
    startDate.setDate(startDate.getDate() + (1 + 7 - startDate.getDay()) % 7);
    infoSheet.getRange('B2').setValue(startDate);

    // init endWeek
    if (endWeek == "") {
      endWeek = 15;
      infoSheet.getRange('B3').setValue(endDate);
    }

    table = scheduleSheet.getRange("A" + STARTROW.toString() + ":M").getValues();
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
    let firstDay = new Date(startDate);
    let lastDay = new Date(startDate);
    lastDay.setDate(firstDay.getDate() + 7);

    // Add endWeek times
    for (i=0; i<endWeek; i++) {
      calendar.createEvent("Tuần " + (i+1).toString(), firstDay, lastDay);

      firstDay.setDate(firstDay.getDate() + 7);
      lastDay.setDate(lastDay.getDate() + 7);
    }
  };

  // Find study weeks 
  this.findWeeks = function(row) {
    let weeks = [];

    let afterBracket = row[HP].split('(')[1];

    // Every weeks
    if (afterBracket == null) return weeks;
    if (afterBracket.match(/\d+/g) == null) return weeks;

    // Weeks between "đến tuần"
    if (afterBracket.search("đến tuần") != -1) {
      let numbs = afterBracket.match(/\d+/g);
      if (numbs.length != 2) {
        SpreadsheetApp.getUi().alert("Error: Weird week format");
      }
      else {
        for (let j=parseInt(numbs[0]); j<=parseInt(numbs[1]); j++) weeks.push(j);
      }
      return weeks;
    }
    
    // Week between '-' or ','
    let arr = afterBracket.split(' ');
    for (let i=0; i<arr.length;i++) {
      if (!/\d/.test(arr[i])) continue;

      let numbs = arr[i].match(/\d+/g);

      if (numbs.length > 2) {
        for (let j=0; j<numbs.length; j++) weeks.push(numbs[j]);
      }
      else if (arr[i].search('-') != -1) {
        for (let j=parseInt(numbs[0]); j<=parseInt(numbs[1]); j++) weeks.push(j);
      }
      else {
        weeks.push(parseInt(numbs[0]));
      }
    }

    return weeks;
  };

  this.deleteEventSeries = function(row, rowNumb) {
    let events = calendar.getEventSeriesById(row[ID].toString());
    events.deleteEventSeries();

    scheduleSheet.getRange(rowNumb, ID+1).clearContent();
  };

  this.setStartTime = function(row) {
    let startTime = new Date(startDate);
    let dayOfWeek = parseInt(row[Thu])-2;
    let startHour = parseInt(row[Tiet].toString().split("-")[0]) + 6;
    startTime.setDate(startTime.getDate()+dayOfWeek);
    startTime.setHours(startHour);
    return startTime;
  };

  this.setEndTime = function(row) {
    let endTime = new Date(startDate);
    let dayOfWeek = parseInt(row[Thu])-2;
    let endHour = parseInt(row[Tiet].toString().split("-")[1]) + 7;
    endTime.setDate(endTime.getDate()+dayOfWeek);
    endTime.setHours(endHour);
    return endTime;
  };

  // Export from sheet to calendar
  this.exportToCalendar = function() {

    // Add subject if Dang ky and ID == ''
    for (let i=0; i<table.length; i++) {
      let rowNumb = i + STARTROW;
      let row = table[i];

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

      let title = row[Nhom] + " - " + row[HP];
      let des = "Mã lớp học phần: "+row[MLHP] + "\nTín chỉ: "+row[TC] + "\nGiảng viên/ Trợ giảng: "+row[GVTG] + "\nGhi chú: "+row[GC];
      let loca = row[GD];

      let weeks = this.findWeeks(row);
      let startTime = this.setStartTime(row);
      let endTime = this.setEndTime(row);

      let times = 0;
      // Todo: remove first practical classes if not suitable
      // if(weeks.length == 0 && row[Nhom] == "CL") {
      //   times = endWeek;
      // }
      // else if (weeks.length == 0 && row[Nhom] != "CL") {
      //   times = endWeek-1;
      //   startTime.setDate(startTime.getDate()+7);
      //   endTime.setDate(endTime.getDate()+7);
      // }
      
      if(weeks.length == 0) {
        times = endWeek;
      }
      else {
        times = weeks[weeks.length-1]-weeks[0]+1;
        startTime.setDate(startTime.getDate()+(weeks[0]-1)*7);
        endTime.setDate(endTime.getDate()+(weeks[0]-1)*7);
      }

      let eventSeries = calendar.createEventSeries( title, 
                                                    startTime, 
                                                    endTime, 
                                                    CalendarApp.newRecurrence().addWeeklyRule().times(times),  //.onlyOnWeeks(weeks)
                                                    {description: des, location: loca});
      scheduleSheet.getRange(rowNumb, ID+1).setValue(eventSeries.getId());
      Logger.log(title + " " + startTime + " " + endTime + " " + weeks);

      // Remove redundant events
      let weekNumb = weeks[0];
      let removeStartTime = startTime;
      let removeEndTime = endTime;

      for (let i = 0; i < weeks.length; i++) {
        if (weekNumb != weeks[i]) {
          i--;
          let events = calendar.getEvents(removeStartTime,removeEndTime);
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

function fetchData(msv) {
  // Todo: Add auto inc semester id
  let semesterID = '036';
  let url = `http://112.137.129.87/qldt/?SinhvienLmh%5BmasvTitle%5D=${msv}&SinhvienLmh%5BhotenTitle%5D=&SinhvienLmh%5BngaysinhTitle%5D=&SinhvienLmh%5BlopkhoahocTitle%5D=&SinhvienLmh%5BtenlopmonhocTitle%5D=&SinhvienLmh%5BtenmonhocTitle%5D=&SinhvienLmh%5Bnhom%5D=&SinhvienLmh%5BsotinchiTitle%5D=&SinhvienLmh%5Bghichu%5D=&SinhvienLmh%5Bterm_id%5D=${semesterID}&SinhvienLmh_page=1&ajax=sinhvien-lmh-grid`;

  let response = UrlFetchApp.fetch(url);
  let $ = Cheerio.load(response.getContentText());
  let registerList = $('tbody tr :nth-child(6), tbody tr :nth-child(8)').map(function() {
    return $(this).text();
  }).get();

  Logger.log(registerList);
  return registerList;
}

/* --------------------------------------------------------------------------------- */

function onOpen() {
  "use strict";
  let menuEntries = [{
    name: "Auto Select Classes",
    functionName: "autoSelectClasses"
  }, {
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

let infoSh = new gSheet("info");
let scheduleSh = new gSheet(infoSh.getSheet().getRange('B4').getValue().toString());
let gCal = new gCalendar(infoSh, scheduleSh);

function autoSelectClasses() {
  let msv = SpreadsheetApp.getUi().prompt("Nhập mã sinh viên (VD: 21020537)").getResponseText();
  let registerList = fetchData(msv);

  scheduleSh.selectClasses(registerList);
  scheduleSh.getSheet().getRange("A2:M").sort(1);
}

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

function removeDuplicates() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var newData = [];
  for (var i in data) {
    var row = data[i];
    var duplicate = false;
    for (var j in newData) {
      if (row.join() == newData[j].join()) {
        duplicate = true;
        break;
      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}
