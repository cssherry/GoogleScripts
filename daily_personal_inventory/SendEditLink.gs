// Instantiate and run constructor
function runSendEditLink() {
  // Change this template to change text in automated email
  var reminderEmail = "If you would like to edit today's daily inventory, please visit the following link: { link }\n",
      currentDate = createPrettyDate(new Date(), 'short'),
      subject = "Edit Link for Daily Personal Inventory (" + currentDate + ")",
      sendTo = 'xiao.qiao.zhou+dpiedit@gmail.com';

  new SendEditLink(emailTemplate, subject, sendTo).run();
}

// Store email template, subject, and sendto
function SendEditLink(emailTemplate, subject, sendTo) {
  var form = FormApp.openById('1FUw_hkDrKN_PVS3oJLHGpM13il-Ugyvfhc_Tg5E_JKc'); //form ID
  this.responses = form.getResponses(); //get email responses

  this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
                                     .getSheetByName("Daily Inventory Data");
  this.scheduleSheetData = this.spreadsheet.getDataRange().getValues();
  this.scheduleSheetIndex = indexSheet(this.scheduleSheetData);

  this.emailTemplate = emailTemplate;
  this.subject = subject;
  this.sendTo = sendTo;

  this.today = new Date();
}

// gets editLink for form and updates spreadsheet/sends link if it's for current day
SendEditLink.prototype.run = function () {
  var startRow = 4,  // First row of data to process
      lastRow = numberOfRows(this.scheduleSheetData),// figure out what the last row is
      editLinkIdx = this.scheduleSheetIndex.EditLink,
      timestampIdx = this.scheduleSheetIndex.Timestamp,
      dateIdx = this.scheduleSheetIndex.Date;

  // Go through each line and check to make sure it has an editLink
  for (var i = 0; i < lastRow ; i++) {
    var editLink = this.addressesSheetData[i][editLinkIdx],
        timestamp = this.addressesSheetData[i][timestampIdx],
        entryDate = this.addressesSheetData[i][dateIdx];

    // If there is not an editLink, put it in, so long as form timestamp and spreadsheet timestamp match
    if (!editLink){
      var response = this.responses[i],
          formResponse = response.getEditResponseUrl(), //grabs the url from the form
          formTimestamp = response.getTimestamp(); //grabs the timestamp from the form

        if (formTimestamp === timestamp) {
          var cellcode = NumberToLetters[editLinkIdx] + (i + 1),
              emailOptions = {
                  link: editLink,
                },
              updateCellOptions = {
                  note: "Reminder sent: " + this.today,
                  sheetName: 'Daily Inventory Data',
                  cellCode: cellCode,
                  message: editLink,
                };

          var email = new Email(this.sendTo, this.subject, this.emailTemplate, emailOptions, [updateCellOptions]);

          // Only send edit link if today is the day that the entry is about
          if (this.today === new Date(entryDate)) {
            email.send();
          } else {
            email.updateCell();
          }
      }
    }
  }
};
