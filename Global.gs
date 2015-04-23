// GLOBAL VARIABLES

// Footer used in all emails
var rules = '\n\nBookclub Rules:\n- Choose a book you have read\n- The person who receives has 2 months to finish it, write some reflection / thoughts in the back\n- Annotations welcomed; use a different color pen than what you\'ve found in the book. Put your initials next to your comments!\n- Once a book is finished, log it by filling out a Google Form. An email will arrive with the next person who should read this book. A separate email will arrive once a new book is about to be sent to you. ',
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

// Shuffle Arrays
var shuffle = function (array) {
  var l = array.length + 1;
  while (l--) {
    var r = ~~(Math.random() * l),
        o = array[r];

    array[r] = array[0];
    array[0] = o;
  }

  return array;
};

// Find first row that is not before today's date -- remember date
var findNextCycle = function(scheduleSheetData, scheduleSheetIndex) {
  var newCycleColumnIdx = scheduleSheetIndex.NewCycle,
      today = new Date();

  for (i = 1; i < scheduleSheetData.length; i++) {
    var newCycle = scheduleSheetData[i][newCycleColumnIdx];

    if (newCycle > today) {
      return [i, newCycle];
    }
  }
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