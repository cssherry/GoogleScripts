// Instantiate and run constructor
function runDeadlineReminder() {
  // Change this template to change text in automated email
  var reminderEmail = "Hi { firstName },\n\nPlease remember to complete  { bookName } by { NewCycle }.\n\nHappy reading!\n\nSchedule here: https://docs.google.com/spreadsheets/d/1wv54jAwqRxPyWAd8a-m_yLNJo2vHYmjEkfp8TCKRWWY/edit?usp=sharing",
      subject = '[BOOKCLUB] Reminder For Upcoming Cycle';

  new DeadlineReminder(reminderEmail, subject).run();
}

// Constructor for assigning book
function DeadlineReminder(reminderEmail, subject) {
  var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet()
                                     .getSheetByName("Schedule");
  this.scheduleSheetData = scheduleSheet.getDataRange().getValues();
  this.scheduleSheetIndex = this.indexSheet(this.scheduleSheetData);

  var addressesSheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName("Addresses");
  this.addressesSheetData = addressesSheet.getDataRange().getValues();
  this.addressesSheetIndex = this.indexSheet(this.addressesSheetData);

  this.reminderEmail = reminderEmail;
  this.subject = subject;

  this.numberToLetters = {
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

  this.today = new Date();
}

DeadlineReminder.prototype.indexSheet = function(sheetData) {
  var result = {},
      length = sheetData[0].length;

  for (var i = 0; i < length; i++) {
    result[sheetData[0][i]] = i;
  }

  return result;
};

// Main script for running function
DeadlineReminder.prototype.run = function() {
  // Get today's date
  var newCycle = findNextCycle(this.scheduleSheetData, this.scheduleSheetIndex),
      newCycleDate = newCycle[1],
      newCycleRowIdx = newCycle[0];

  // Go through every column in Schedule tab, send email if the person has not finished book yet -- add note when successfully sent email
  for (var i = 1; i < this.scheduleSheetData[0].length; i++) {
    if (this.scheduleSheetData[i][newCycleRowIdx] === "") {
      var emailIdx = this.addressesSheetIndex.Email,
          contactEmail = this.addressesSheetData[i][emailIdx],
          sheetName = 'Schedule',
          cellCode = this.numberToLetters[i] + newCycleRowIdx,
          nameIdx = this.addressesSheetIndex.Name,
          options = {note: "Reminder sent: " + this.today,
                     NewCycle: newCycleDate,
                     bookName: this.scheduleSheetData[newCycleRowIdx - 1][i],
                     firstName: this.addressesSheetData[i][nameIdx]};

      new Email(contactEmail, this.subject, this.reminderEmail, sheetName, cellCode, options);
    }
  }
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
