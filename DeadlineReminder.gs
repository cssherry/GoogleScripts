// Instantiate and run constructor
function runDeadlineReminder() {
  // Change this template to change text in automated email
  var reminderEmail = "Hi { firstName },\n\nPlease remember to complete  { bookName } by { NewCycle }.\n\nHappy reading!",
      subject = '[BOOKCLUB] Reminder For Upcoming Cycle';

  new DeadlineReminder(reminderEmail, subject).run();
}

// Constructor for assigning book
function DeadlineReminder(reminderEmail, subject) {
  var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet()
                                     .getSheetByName("Schedule");
  this.scheduleSheetData = scheduleSheet.getDataRange().getValues();
  this.scheduleSheetIndex = indexSheet(this.scheduleSheetData);

  var addressesSheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName("Addresses");
  this.addressesSheetData = addressesSheet.getDataRange().getValues();
  this.addressesSheetIndex = indexSheet(this.addressesSheetData);

  this.reminderEmail = reminderEmail;
  this.subject = subject;

  this.today = new Date();
}

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
          cellCode = NumberToLetters[i] + newCycleRowIdx,
          nameIdx = this.addressesSheetIndex.Name,
          options = {note: "Reminder sent: " + this.today,
                     NewCycle: newCycleDate,
                     bookName: this.scheduleSheetData[newCycleRowIdx - 1][i],
                     firstName: this.addressesSheetData[i][nameIdx]};

      new Email(contactEmail, this.subject, this.reminderEmail, sheetName, cellCode, options);
    }
  }
};