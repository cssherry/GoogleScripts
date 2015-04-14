// Instantiate and run constructor
function runAssignBook() {
  // Change this template to change text in automated email
  var mailInfo = "Hi {{ firstName }},\n\nPlease send your book to {{ sendToPerson }}. Their address is below:\n{{ sendAddress }}\n\nHappy reading!",
      nextBookInfo = "Hi {{ sendToPerson }},\n\nExpect to get {{ newBook }} soon from {{ firstName }}\n\nHappy reading!",

      assign = AssignBook(mailInfo, nextBookInfo);

  assign.run();
}

// Constructor for assigning book
function AssignBook(mailInfo, nextBookInfo) {
  var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet()
                                     .getSheetByName("Schedule");
  this.scheduleSheetData = scheduleSheet.getDataRange().getValues();
  this.scheduleSheetIndex = this.indexSheet(this.scheduleSheet);

  var formSheet = SpreadsheetApp.getActiveSpreadsheet()
                                 .getSheetByName("Form Responses 1");
  this.formSheetData = formSheet.getDataRange().getValues();
  this.formSheetIndex = this.indexSheet(this.formSheet);
  this.formLastEntryIdx = this.formSheetIndex.length - 1;

  var addressesSheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName("Addresses");
  this.addressesSheetData = addressesSheet.getDataRange().getValues();
  this.addressesSheetIndex = this.indexSheet(this.addressesSheetData);

  this.mailInfo = mailInfo;
  this.mailInfo = mailInfo;
}

AssignBook.prototype.indexSheet = function(sheetData) {
  var result = {},
      length = sheetData[0].length;

  for (var i = 0; i < length; i++) {
    result[i] = sheetData[0][i];
  }

  return result;
};

AssignBook.prototype.run = function() {
  // PSEUDOCODE

  // Get last entry in form response
  // find what book the person last read
  // go to schedule and create a hash for each person of books read
  // Find "pending" in schedule
  // check to make sure pending has not already read book from previous person
  // if so, check all pendings
  // if no match, then add note saying 'waitingForMatch', with note that has id of form submission (https://developers.google.com/apps-script/reference/spreadsheet/)
  // if match, find entry and note new book on the way, and send another email letting people know
  // Check to see if people with notes saying "waitingForMatch" can send to person. If not, add Pending with note on formid
  // send to original sender if length of schedule column === 7

  for (i = 1; i < length; i++) {

    var emailWasSent = data[i][COLUMNMAP['emailWasSent']];
    if (!emailWasSent) {

      var jobApplication = new JobApplication(this, data[i]);

      if (jobApplication['applyByEmail']) {
        jobApplication.createEmail();
        jobApplication.fire();
        this.recordEmailSent(i);
      }

      break; // Only send one email when the script is run.
    }

  }
};

AssignBook.prototype.findWhoWillReadNext = function() {

};

AssignBook.prototype.checkForMatches = function() {

};

AssignBook.prototype.recordEmailSent = function(companyIndex) {
  var rowIdx = companyIndex + 1,
    emailWasSentColumn = COLUMNLETTERS[COLUMNMAP['emailWasSent']],
    cell = SpreadsheetApp.getActiveSheet().getRange(emailWasSentColumn + rowIdx);

  cell.setValue(true);
};
