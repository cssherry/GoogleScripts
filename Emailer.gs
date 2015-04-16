// Define the Email constructor
function Email(contactEmail, subject, template, options) {
  this.contactEmail = contactEmail;
  this.template = template;
  this.subject = subject;
  this.options = options;
}

// Replaces all keywords in email template with their actual values
Email.prototype.populateEmail = function() {
  // dateColumn should be edited to include titles of all columns that contain dates
  var dateColumns = ['Timestamp', 'NewCycle'];

  for (var keyword in this.options) {
    if (this.findInArray(dateColumns, keyword) > -1) {
      this.template.replaceText('{ ' + keyword + ' }', createPrettyDate(this.options[keyword]));
    } else {
      this.template.replaceText('{ ' + keyword + ' }', this.options[keyword]);
    }
  }
};

// Calls MailApp to send email
Email.prototype.send = function (sheetName, cellCode, options) {
    MailApp.sendEmail({
      to: this.contactEmail,
      subject: this.subject,
      body: this.populateEmail(),
    });

    this.updateCell(sheetName, cellCode, options);
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

  prettyDate = daysOfWeekIndex.dayOfWeek + ', ' + monthIndex.mm + ' ' + dd;
  return '*' + prettyDate + '*';
};

// Helper function to find string in an array
Email.prototype.findInArray = function(array, string) {
  for (var j=0; j < array.length; j++) {
      if (array[j].match(string)) return j;
  }
  return -1;
};

// Function that records when an email is successfully sent
Email.prototype.updateCell = function(sheetName, cellCode, options) {
  SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName(sheetName)
                .getRange(cellCode)
                .setNote(options.note)
                .setValue(options.message);
};
