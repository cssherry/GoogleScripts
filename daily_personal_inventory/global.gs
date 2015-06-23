// GLOBAL VARIABLES

// Header/Footer used in all emails
var header = "Don't forget! !1 for Urgent&Important (DO NOW!), !2 for Urgent (Try to Delegate), !3 for Important (Decide when to do), !4 for do later\nUse #e0 for things to do as a break\n\n" +
             "Start to-do list with three entries: *must* do (an immediate important task), *should* do (something for long-term goals), and something genuinely *want* to do\n\n" +
             "Keep to-do short with 1 big thing, 3 medium things, and 5 little things\n\n",
    currentDate = createPrettyDate(new Date(), 'short'),
    footer = "\n\nAs reported on { timestamp }\n\n" +
             "Fill out the form: https://docs.google.com/forms/d/1FUw_hkDrKN_PVS3oJLHGpM13il-Ugyvfhc_Tg5E_JKc/viewform\n\n" +
             "See the spreadsheet: https://docs.google.com/spreadsheet/ccc?key=0AggnWnxIWH43dFEtdU5jZmwxM2kyU2ZaNk5KOVl1SXc#gid=0";

// To convert column index to letter for cells
var NumberToLetters = function(n) {
  var ordA = 'a'.charCodeAt(0);
  var ordZ = 'z'.charCodeAt(0);
  var len = ordZ - ordA + 1;

  var s = "";
  while(n >= 0) {
      s = String.fromCharCode(n % len + ordA) + s;
      n = Math.floor(n / len) - 1;
  }
  return s.toUpperCase();
};

// HELPER FUNCTIONS

// Create hash with column name keys pointing to column index
// For greater flexibility (columns can be moved around)
var indexSheet = function(sheetData) {
  var result = {},
      length = sheetData[0].length;

  for (var i = 0; i < length; i++) {
    result[sheetData[0][i]] = i;
  }

  return result;
};

// Find string in an array
var findInArray = function(array, string) {
  for (var j=0; j < array.length; j++) {
      if (array[j].match(string)) {
        return j;
      }
  }
  return -1;
};

// Find last row of column
var numberOfRows = function (sheetData, _columnIndex) {
  var columnIndex = _columnIndex ? _columnIndex : 0;

  for (var i = 0; i < sheetData.length; i++) {
    if (!sheetData[i][columnIndex]) {
      return i;
    }
  }

  return i;
};

var sameDay = function (date1, date2) {
  date1 = new Date(date1);
  date2 = new Date(date2);

  if (date1.getDate() === date2.getDate() && date1.getMonth() === date2.getMonth() && date1.getYear() === date2.getYear()) {
    return true;
  } else {
    return false;
  }
};

// Create pretty date in either short (yyyy/mm/dd) or long (ww, mm dd) format
var createPrettyDate = function(date, format) {
  var dateObject = new Date(date), dd, mm, yyyy, dayOfWeek;
  if (format === 'short') {
    mm = dateObject.getMonth() + 1;
    dd = dateObject.getDate();
    yyyy = dateObject.getFullYear();
    return yyyy + "/" + mm + "/" + dd;
  } else {
    var daysOfWeekIndex = { 0: 'Sunday',
                            1: 'Monday',
                            2: 'Tuesday',
                            3: 'Wednesday',
                            4: 'Thursday',
                            5: 'Friday',
                            6: 'Saturday',
                          };

    var monthIndex = { 0: 'January',
                       1: 'February',
                       2: 'March',
                       3: 'April',
                       4: 'May',
                       5: 'June',
                       6: 'July',
                       7: 'August',
                       8: 'September',
                       9: 'October',
                       10: 'November',
                       11: 'December',
                     };

    dd = dateObject.getDate();
    mm = dateObject.getMonth();
    dayOfWeek = dateObject.getDay(); // starts at Sunday

    return daysOfWeekIndex[dayOfWeek] + ', ' + monthIndex[mm] + ' ' + dd;
  }
};
