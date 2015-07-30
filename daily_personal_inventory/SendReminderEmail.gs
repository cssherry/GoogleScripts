function runSendReminderEmail() {
  // Change this template to change text in automated email
  var reminder = "Today's To-Dos:\n{ goal } " +
                 "\n\nLife goals: \n{ life } \n\n " +
                 "Remember to be thankful for: \n" +
                 "     1) { grateful1 } \n" +
                 "     2) { grateful2 } \n" +
                 "     3) { grateful3 } \n\n" +
                 "Edit at: { editLink } \n\n" +
                 "-----------------------------\n\n",
      reminderFooter = "{ missed }\n\n" +
                       "-----------------------------\n\n" +
                       "{ throwbacks }" +
                       asReported,
      subject = "To-Do's for Today (" + currentDate + ")";
      missingOnlySubject = "Missing " + this.missingDays + " Days (" + currentDate + ")";
      sendTo = '7a828627@opayq.com';

  new sendReminderEmail(reminder + reminderFooter, subject,
                        reminderFooter, missingOnlySubject,
                        sendTo).run();
}

function sendReminderEmail(emailTemplate, subject, missingOnlyEmail, missingOnlySubject, sendTo) {
  this.responseSheetSorted = SpreadsheetApp.getActiveSpreadsheet()
                                     .getSheetByName("Daily Inventory Data");
  this.responseSheetData = this.responseSheetSorted.getDataRange().getValues();
  this.responseSheetIndex = indexSheet(this.responseSheetData);

  // Bunch of numbers from responseSheetIndex
  this.goalIdx = this.responseSheetIndex["What are tomorrow's goals?"];
  this.lifeIdx = this.responseSheetIndex['What are your life goals?'];
  this.grateful1Idx = this.responseSheetIndex['What are you grateful for?'];
  this.grateful2Idx = this.responseSheetIndex['What are you grateful for?'] - 1;
  this.grateful3Idx = this.responseSheetIndex['What are you grateful for?'] - 2;
  this.editLinkIdx = this.responseSheetIndex.EditLink;
  this.creativeWritingIdx = this.responseSheetIndex['Creative Writing'];
  this.timestampIdx = this.responseSheetIndex.Timestamp;
  this.dateIdx = this.responseSheetIndex.Date;
  this.daysFromIdx = this.responseSheetIndex.DaysFromToday;
  this.emailSentIdx = this.responseSheetIndex.EmailSent;

  this.scoreCardSorted = SpreadsheetApp.getActiveSpreadsheet()
                                       .getSheetByName("Score Card");
  this.scoreCardData = this.scoreCardSorted.getDataRange().getValues();
  this.scoreCardIndex = indexSheet(this.scoreCardData);

  this.emailTemplate = emailTemplate;
  this.subject = subject;
  this.missingOnlyEmailTemplate = missingOnlyEmail;
  this.missingOnlySubject = missingOnlySubject;
  this.sendTo = sendTo;

  this.today = new Date();
}

sendReminderEmail.prototype.run = function () {
  var startRow = 3,  // First row of data to process
      endRow = this.scoreCardData.length - 1;

  // Gets all missing dates
  var missed = this.getMissingDates(startRow, endRow);

  // Get all the throwback dates
  startRow = 4;  // First row of data to process
  endRow = this.responseSheetData.length - 1;// figure out what the last row is
  var throwbacks = this.getThrowbacks(startRow, endRow);

  //Creates email message for yesterday's response if possible
  this.trySendingYesterdayEmail(startRow, endRow,
                               {missed: missed, throwbacks: throwbacks});

  // Send a missing reminder if the form wasn't filled out yesterday
  if (this.missingYesterday && this.missingDays.highestStreak < 21){
    var emailOptions = {
            missed: missed,
            throwbacks: throwbacks,
            timestamp: this.today,
          };
    new Email(this.sendTo, this.missingOnlySubject, this.missingOnlyEmailTemplate, emailOptions)
      .send();
  }
};

// Return the "missed" text
// Saves a this.missingDays object (with currentStreak + highestStreak attributes) if there are missing days
sendReminderEmail.prototype.getMissingDates = function (startRow, endRow) {
  // Get all the dates that were missed
  var numRows = endRow - startRow,
      numberMissed = {currentStreak: 0, highestStreak: 0},
      daysFromIdx = this.scoreCardIndex['Drinking Score'],
      n = 0,
      missed = '';

  for (var i = endRow; i >= startRow; i--) {
    var daysFrom = this.scoreCardData[i][daysFromIdx];
    if(daysFrom !== '1') {
      missed += "     " + daysFrom + "\n";
      n++;
      numberMissed.currentStreak++;
    } else {
      // reset streak counter
      numberMissed.currentStreak = 0;
    }
    if (numberMissed.highestStreak < numberMissed.currentStreak) {
      numberMissed.highestStreak = numberMissed.currentStreak;
    }
  }

  if (n > 0){
    this.missingDays = numberMissed;
    return "You missed " + n + " days this month in the Daily Personal Inventory " +
            "(https://docs.google.com/forms/d/1FUw_hkDrKN_PVS3oJLHGpM13il-Ugyvfhc_Tg5E_JKc/viewform)\n" +
            missed + "\n";
  } else {
    return "";
  }
};

// Sends email if yesterday's was filled out
// saves object missingYesterday if yesterday's form was not filled out
sendReminderEmail.prototype.trySendingYesterdayEmail = function (startRow, endRow, emailOptions) {
  var missingYesterday = true;
  for (var i = startRow; i <= endRow; i++) {
    var daysFrom = this.responseSheetData[i][this.daysFromIdx];
    if(daysFrom === 1) {
      missingYesterday = false;
        emailOptions.goal = this.responseSheetData[i][this.goalIdx];
        emailOptions.life = this.responseSheetData[i][this.lifeIdx];
        emailOptions.grateful1 = this.responseSheetData[i][this.grateful1Idx];
        emailOptions.grateful2 = this.responseSheetData[i][this.grateful2Idx];
        emailOptions.grateful3 = this.responseSheetData[i][this.grateful3Idx];
        emailOptions.editLink = this.responseSheetData[i][this.editLinkIdx];
        emailOptions.timestamp = this.responseSheetData[i][this.timestampIdx];
        var emailSent = this.responseSheetData[i][this.emailSentIdx];     // 36 column
          if (emailSent.indexOf('EMAIL_SENT') === -1) {  // Prevents sending duplicates
            var cellcode = NumberToLetters(this.emailSentIdx) + (i + 1),
                updateCellOptions = {
                    sheetName: 'Daily Inventory Data',
                    cellCode: cellcode,
                    message: 'EMAIL_SENT',
                    note: "Email sent: " + this.today,
                  };

            new Email(this.sendTo, this.subject, this.emailTemplate, emailOptions, [updateCellOptions])
              .send();
          }
    }
  }

  this.missingYesterday = missingYesterday;
};

// returns back message with 4 days of thankfulness from 20 days ago
sendReminderEmail.prototype.getThrowbacks = function (startRow, endRow) {
  var message = "",
      m = 0;

  //Creates message with random grateful thing
  for (var j = endRow - 20; j >= startRow ; j--) {
    if(this.responseSheetData[j][this.grateful1Idx] !== "" && m < 4) {
      var days_from = this.responseSheetData[j][this.daysFromIdx],
          grateful1 = this.responseSheetData[j][this.grateful1Idx],
          grateful2 = this.responseSheetData[j][this.grateful2Idx],
          grateful3 = this.responseSheetData[j][this.grateful3Idx],
          editLink = this.responseSheetData[j][this.editLinkIdx],
          creativeWriting = this.responseSheetData[j][this.creativeWritingIdx];

      message += "\n\nRemember " + days_from + " ago you were thankful for: \n" +
                 "Edit entry here: " + editLink + " \n" +
                 "     1) " + grateful1 + " \n";

      if (grateful2 !== "") {
        message += "     2) " + grateful2 + " \n";
        if (grateful3 !== "") {
          message += "     3) " + grateful3 + " \n\n";
        } else {
          message += "\n";
        }
      } else {
        message += "\n";
      }
      if (creativeWriting !== "") {
        message += "\n Creative Writing\n" + creativeWriting + "\n\n";
      }
      m++;
    }
  }

  return message;
};
