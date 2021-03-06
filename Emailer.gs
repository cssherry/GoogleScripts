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
  var dateColumns = ['Timestamp'];

  for (var keyword in this.options) {
    if (findInArray(dateColumns, keyword) > -1) {
      this.template = this.template.replace('{ ' + keyword + ' }', createPrettyDate(this.options[keyword]));
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
      htmlBody: this.populateEmail() +
                "\n\n---------------------------------------" +
                rules +
                urls,
    });

    this.updateCell();
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
        if (currentNote !== "") {
          currentNote += "\n";
        }
        cell.setNote(currentNote + note);
      }

      if (message) {
        currentMessage = cell.getValue();
        if (currentMessage !== "") {
          currentMessage += "\n";
        }
        cell.setValue(currentMessage + message);
      }
    }
  }
};
