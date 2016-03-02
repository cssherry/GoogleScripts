// GLOBAL VARIABLES

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

// Header/Footer used in all emails
var header = "Don't forget! !1 for Urgent&Important (DO NOW!), !2 for Urgent (Try to Delegate), !3 for Important (Decide when to do), !4 for do later\nUse #e0 for things to do as a break\n" +
             "Start to-do list with three entries: *must* do (an immediate important task), *should* do (something for long-term goals), and something genuinely *want* to do\n" +
             "Keep to-do short with 1 big thing, 3 medium things, and 5 little things\n\n",
    currentDate = createPrettyDate(new Date(), 'short'),
    asReported = '\n\nAs reported on { timestamp }',
    footer = "\n\nFill out the form: https://docs.google.com/forms/d/1FUw_hkDrKN_PVS3oJLHGpM13il-Ugyvfhc_Tg5E_JKc/viewform\n\n" +
             "See the spreadsheet: https://docs.google.com/spreadsheet/ccc?key=0AggnWnxIWH43dFEtdU5jZmwxM2kyU2ZaNk5KOVl1SXc#gid=0\n\n";

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

function getElementByVal( element, elementType, attr, val ) {
  var value = element[attr];
  // If the current element matches, return it.
  if (element[attr] &&
      value == val &&
      element.getName().getLocalName() == elementType) {
    return element;
  }

  // Current element didn't match, check its children
  var elList = element.getElement();
  var i = elList.length;
  while (i--) {
    // (Recursive) Check each child, in document order.
    var found = getElementByVal( elList[i], elementType, attr, val );
    if (found !== null) return found;
  }
  // No matches at this element OR its children
  return null;
}

function getHTML(url, el, attr, attrVal) {
  var response;
  try {
    response = UrlFetchApp.fetch(url);
  } catch (e) {
    return "Sorry but Google couldn't fetch the requested web page. " +
           "Please try another URL!<br />" +
           "<small>" + e.toString() + "</small>";
  }

  var xml = response.getContentText();
  var document = Xml.parse(xml, true);


  var element = getElementByVal(document, el, attr, attrVal);
  return element;
}

function runGetHtml(){
  return getHTML('http://www.merriam-webster.com/word-of-the-day/', 'div', 'class', 'wod-definition-container' );
}
// Day2 is newer date, day1 is older date
function getDayDiff (day1, day2) {
  if (!isDate(day1)) {
    day1 = parseDateString(day1);
  }
  if (day2) {
    if (!isDate(day2)) {
      day2 = parseDateString(day2);
    }
  } else {
    day2 = new Date();
  }
  var milliPerSecond = 1000,
      secondPerMin = 60,
      minToHour = 60,
      hourToDay = 24,
      conversionFactor = milliPerSecond * secondPerMin * minToHour * hourToDay;
  return Math.floor((day2 - day1) / conversionFactor);

  function isDate (date) {
    return Object.prototype.toString.call(date) === '[object Date]';
  }
  function parseDateString (date) {
    var mdy = date.split('/');
    return new Date(mdy[2], mdy[0]-1, mdy[1]);
  }
}
