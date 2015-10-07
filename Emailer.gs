// Define the Email constructor
// contactEmail: who to email to
// subject: email subject
// template: email body
// emailOptions: Any regex keys and their values
// updateCellsOptions: an array of objects, each with sheetName, cellCode, note, and message properties
function Email(contactEmail, subject, template, emailOptions, updateCellsOptions) {
  this.contactEmail = contactEmail;
  this.subject = subject;
  this.template = template;
  this.options = emailOptions;
  this.updateCellsOptions = updateCellsOptions;

  this.send();
}

// Replaces all keywords in email template with their actual values
Email.prototype.populateEmail = function() {
  // dateColumn should be edited to include titles of all columns that contain dates
  var dateColumns = ['Timestamp', 'NewCycle'];

  for (var keyword in this.options) {
    if (findInArray(dateColumns, keyword) > -1) {
      this.template = this.template.replace('{ ' + keyword + ' }', this.createPrettyDate(this.options[keyword]));
    } else {
      this.template = this.template.replace('{ ' + keyword + ' }', this.options[keyword]);
    }
  }

  return this.template;
};

// Calls MailApp to send email
Email.prototype.send = function () {
    MailApp.sendEmail({
      to: this.contactEmail,
      subject: this.subject,
      body: this.populateEmail() +
            "\n\n---------------------------------------" +
            rules +
            urls,
    });

    this.updateCell();
};

Email.prototype.createPrettyDate = function(date) {
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

  var dateObject = new Date(date);
  var dd = dateObject.getDate();
  var mm = dateObject.getMonth();
  var dayOfWeek = dateObject.getDay(); // starts at Sunday

  var prettyDate = daysOfWeekIndex[dayOfWeek] + ', ' + monthIndex[mm] + ' ' + dd;
  return '*' + prettyDate + '*';
};

// Function that records when an email is successfully sent
Email.prototype.updateCell = function() {
  var options, sheetName, cellCode, note, message, cell, currentNote, currentMessage;
  if (this.updateCellsOptions) {
    for (var i = 0; i < this.updateCellsOptions.length; i++) {
      options = this.updateCellsOptions[i];
      sheetName = options.sheetName;
      cellCode = options.cellCode;
      note = options.note;
      message = options.message;

      cell = SpreadsheetApp.getActiveSpreadsheet()
                               .getSheetByName(sheetName)
                               .getRange(cellCode);

      if (note) {
        currentNote = cell.getNote();
        cell.setNote(currentNote + "\n" + note);
      }

      if (message) {
        currentMessage = cell.getValue();
        cell.setValue(currentMessage + "\n" + message);
      }
    }
  }
};
