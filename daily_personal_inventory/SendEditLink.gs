// Instantiate and run constructor
function runSendEditLink() {
  // Change this template to change text in automated email
  var reminderEmail = "Edit link: { link }\n" + asReported,
      subject = "Edit Link for Daily Personal Inventory (" + currentDate + ")",
      sendTo = myEmail;

  Utilities.sleep(4 * 1000);
  new getEditLink(reminderEmail, subject, sendTo).run();
}

// Store email template, subject, and sendto
function getEditLink(emailTemplate, subject, sendTo) {
  var form = FormApp.openById('1kL9sSIQbbBnb3Botbf0RuJepG6ird_GXqUwkSZ1oTg4'); //form ID
  this.responses = form.getResponses(); //get email responses

  this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
                                   .getSheetByName("Daily Inventory Data");
  this.responseSheetData = this.spreadsheet.getDataRange().getValues();
  this.responseSheetIndex = indexSheet(this.responseSheetData);

  this.emailTemplate = emailTemplate;
  this.subject = subject;
  this.sendTo = sendTo;

  this.today = new Date();
}

// gets editLink for form and updates spreadsheet/sends link if it's for current day
getEditLink.prototype.run = function () {
  var startRow = 3,  // First row of data to process
      numberEntries = this.spreadsheet.getLastRow(),// figure out what the last row is (the first row has 2 entries before first real entry)
      editLinkIdx = this.responseSheetIndex.EditLink,
      timestampIdx = this.responseSheetIndex.Timestamp,
      dateIdx = this.responseSheetIndex.Date,
      hoursSleepIdx = this.responseSheetIndex['How many hours did you sleep?'],
      rowIdx, editLink, entryDate, sleepTime;

  // Go through each line and check to make sure it has an editLink
  for (var i = 0; i < numberEntries - startRow; i++) {
    rowIdx = startRow + i;
    editLink = this.responseSheetData[rowIdx][editLinkIdx];
    timestamp = this.responseSheetData[rowIdx][timestampIdx];
    entryDate = this.responseSheetData[rowIdx][dateIdx];
    sleepTime = this.responseSheetData[rowIdx][hoursSleepIdx];

    // If there is not an editLink, put it in, so long as form timestamp and spreadsheet timestamp match
    if (!editLink && timestamp /*check for timestamp and make sure row isn't empty*/){
      var response = this.responses.filter(checkTimestamp)[0],
          formUrl = response.getEditResponseUrl(); //grabs the url from the form

          var cellcode = NumberToLetters(editLinkIdx) + (rowIdx + 1),
              emailOptions = {
                  link: formUrl,
                  timestamp: timestamp
                },
              updateCellOptions = {
                  sheetName: 'Daily Inventory Data',
                  cellCode: cellcode,
                  message: formUrl,
                },
              email;


        // Only send edit link if today is the day that the entry is about
        if (sameDay(this.today, entryDate)) {
          updateCellOptions.note = "Reminder sent: " + this.today;
          email = new Email(this.sendTo, this.subject, this.emailTemplate, emailOptions, [updateCellOptions]);
          email.send();
        } else {
          updateCellOptions.note = "Script ran: " + this.today;
          email = new Email(this.sendTo, this.subject, this.emailTemplate, emailOptions, [updateCellOptions]);
          email.updateCell();
        }
    }
    // recalculate sleeptime if less than 4 hours -- probably calendar event wasn't uploaded yet
    // Only do this for last 30 dates (otherwise, overuse)
    if (rowIdx > numberEntries - 31 && (!sleepTime || sleepTime < 4)) {
      // If user hasn't put in sleep time, insert sleep like an android sleep time and info
      this.getSleep(entryDate, rowIdx, hoursSleepIdx);
    }
  }

  function checkTimestamp(response){
    var rTimestamp = response.getTimestamp();
    if (timestamp.getTime() === rTimestamp.getTime()) {
      return response;
    }
  }
};

// Gets sleep info from calendar inserted by Sleep like an Android
getEditLink.prototype.getSleep = function(currDate, row, sleepIdx) {
  if (!this.sleepCalendar) {
    this.sleepCalendar = CalendarApp.getCalendarsByName("Sleep")[0];
  }
  var startTime = new Date(currDate),
      endTIme = new Date(currDate),
      cellcode = NumberToLetters(sleepIdx) + (row + 1),
      eventLength = 0,
      eventDescription = "",
      updateCellOptions;

  startTime.setDate(currDate.getDate() - 1);
  startTime.setHours(22);
  endTIme.setHours(22);

  // Lag between calls prevents 'Service invoked too many times in a short time'
  Utilities.sleep(1000);
  var sleepEvents = this.sleepCalendar.getEvents(startTime, endTIme);

  for (var i = 0; i < sleepEvents.length; i++) {
    eventLength += (sleepEvents[i].getEndTime() - sleepEvents[i].getStartTime());
    eventDescription += ("\n\n" + sleepEvents[i].getDescription());
  }

  updateCellOptions = {
    sheetName: 'Daily Inventory Data',
    cellCode: cellcode,
    message: eventLength / 1000 / 60 / 60,
    note: eventDescription,
    overwrite: true,
  };

  new Email(null, null, null, null, [updateCellOptions]).updateCell();
};
