// GLOBAL VARIABLES

// Footer used in all emails
var rules = '\n\nBookclub Rules:\n- Choose a book you have read\n- Each person has 2 months to finish their book, write some reflection / thoughts in the back\n- Annotations welcomed; use a different color pen than what you\'ve found in the book. Put your initials next to your comments!\n- Once a book is finished, log it by filling out the Google Form (link below). \n- An email will arrive with the address you should send the book to -- please send the book 7 days after receiving the assignment. A separate email will arrive once a new book is about to be sent to you. ',
    urls = '\n\nGoogle Form: https://docs.google.com/forms/d/1j6oYWu4QcadddV2VD0hBQ7XUVbYnwUrAkgowP_jXSaQ/viewform\nSchedule: https://docs.google.com/spreadsheets/d/1wv54jAwqRxPyWAd8a-m_yLNJo2vHYmjEkfp8TCKRWWY/edit?usp=sharing\nGoodreads:https://www.goodreads.com/group/show/160644-ramikip-2-0.html';

// To convert column index to letter for cells
var NumberToLetters = {
  0: 'A',
  1: 'B',
  2: 'C',
  3: 'D',
  4: 'E',
  5: 'F',
  6: 'G',
  7: 'H',
  8: 'I',
  9: 'J',
  10: 'K',
  11: 'L',
  12: 'M',
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

// Create pretty date in either short (yyyy/mm/dd) or long (ww, mm dd) format
createPrettyDate = function(date, format) {
  var dateObject = new Date(date), dd, mm, yyyy, dayOfWeek;
  if (format === 'short') {
    mm = dateObject.getMonth() + 1;
    dd = dateObject.getDate();
    yyyy = dateObject.getFullYear();
    return yyyy + "/" + mm + "/" + dd
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
