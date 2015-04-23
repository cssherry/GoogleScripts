// Footer used in all emails
var rules = '\n\nBookclub Rules:\n- Choose a book you have read\n- The person who receives has 2 months to finish it, write some reflection / thoughts in the back\n- Annotations welcomed; use a different color pen than what you\'ve found in the book. Put your initials next to your comments!\n- Once a book is finished, log it by filling out a Google Form. An email will arrive with the next person who should read this book. A separate email will arrive once a new book is about to be sent to you. ',
    urls = '\n\nGoogle Form: https://docs.google.com/forms/d/1j6oYWu4QcadddV2VD0hBQ7XUVbYnwUrAkgowP_jXSaQ/viewform\nSchedule: https://docs.google.com/spreadsheets/d/1wv54jAwqRxPyWAd8a-m_yLNJo2vHYmjEkfp8TCKRWWY/edit?usp=sharing\nGoodreads:https://www.goodreads.com/group/show/160644-ramikip-2-0.html';

// Define the Email constructor
function Email(contactEmail, subject, template, sheetName, cellCode, options) {
  this.contactEmail = contactEmail;
  this.template = template;
  this.subject = subject;
  this.sheetName = sheetName;
  this.cellCode = cellCode;
  this.options = options;

  this.send();
}

// Replaces all keywords in email template with their actual values
Email.prototype.populateEmail = function() {
  // dateColumn should be edited to include titles of all columns that contain dates
  var dateColumns = ['Timestamp', 'NewCycle'];

  for (var keyword in this.options) {
    if (this.findInArray(dateColumns, keyword) > -1) {
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

// Helper function to find string in an array
Email.prototype.findInArray = function(array, string) {
  for (var j=0; j < array.length; j++) {
      if (array[j].match(string)) {
        return j;
      }
  }
  return -1;
};

// Function that records when an email is successfully sent
Email.prototype.updateCell = function(_) {
  var sheetName,
      cellCode,
      note,
      message;

  if (_) {
    sheetName = _.sheetName;
    cellCode = _.cellCode;
    note = _.note;
    message = _.message;
  } else {
    sheetName = this.sheetName;
    cellCode = this.cellCode;
    note = this.options.note;
    message = this.options.message;
  }

  var cell = SpreadsheetApp.getActiveSpreadsheet()
                           .getSheetByName(sheetName)
                           .getRange(cellCode);

  if (note) {
    cell.setNote(note);
  }

  if (message) {
    cell.setValue(message);
  }
};
