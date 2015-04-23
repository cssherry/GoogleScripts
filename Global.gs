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

// Global variable
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