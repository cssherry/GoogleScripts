function Email(options) {
  this.contactEmail = options.contactEmail;
  this.subject = options.subject;
  this.sendDate = options.sendDate;
  this.sendToPerson = options.sendToPerson;
  this.sendAddress = options.sendAddress;
  this.deadline = options.deadline;
  this.firstName = options.firstName;
  this.template = options.template;
}

Email.prototype.populateEmail = function() {
  this.coverLetter.getBody()
                  .replaceText('{ sendDate }', createPrettyDate(this.sendDate))
                  .replaceText('{ sendToPerson }', this.sendToPerson)
                  .replaceText('{ sendAddress }', this.sendAddress)
                  .replaceText('{ deadline }', createPrettyDate(this.deadline))
                  .replaceText('{ firstName }', this.firstName);
};

Email.prototype.send = function (jobberator) {
    MailApp.sendEmail({
      to: this.contactEmail,
      subject: this.subject,
      body: this.populateEmail(),
    });
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
