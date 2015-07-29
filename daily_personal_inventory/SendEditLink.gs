// Instantiate and run constructor
function runSendEditLink() {
  // Change this template to change text in automated email
  var reminderEmail = "Edit link: { link }\n" + asReported,
      subject = "Edit Link for Daily Personal Inventory (" + currentDate + ")",
      sendTo = '7a828627@opayq.com';

  new getEditLink(reminderEmail, subject, sendTo).run();
}

// Store email template, subject, and sendto
function getEditLink(emailTemplate, subject, sendTo) {
  var form = FormApp.openById('1FUw_hkDrKN_PVS3oJLHGpM13il-Ugyvfhc_Tg5E_JKc'); //form ID
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
      numberEntries = this.responseSheetData.length - startRow,// figure out what the last row is (the first row has 2 entries before first real entry)
      editLinkIdx = this.responseSheetIndex.EditLink,
      timestampIdx = this.responseSheetIndex.Timestamp,
      dateIdx = this.responseSheetIndex.Date;

  // Go through each line and check to make sure it has an editLink
  for (var i = 0; i <= numberEntries ; i++) {
    var rowIdx = startRow + i,
        editLink = this.responseSheetData[rowIdx][editLinkIdx],
        timestamp = this.responseSheetData[rowIdx][timestampIdx],
        entryDate = this.responseSheetData[rowIdx][dateIdx];

    // If there is not an editLink, put it in, so long as form timestamp and spreadsheet timestamp match
    if (!editLink){
      var response = this.responses.filter(function(r){
                                            var rTimestamp = r.getTimestamp();
                                            if (timestamp.getTime() === rTimestamp.getTime()) {
                                              return r;
                                            }
                                          })[0],
          formUrl = response.getEditResponseUrl(); //grabs the url from the form

        // Use + to call valueOf() behind the scenes. Another option would be to call getTime()
        // a =  function () {
        //   var d1 = new Date(2013, 0, 1);
        //   var d2 = new Date(2013, 0, 1);
        //   console.time('valueOf');
        //     console.log((+d1 === +d2));
        //     console.log((+d1 === +d2));
        //     console.log((+d1 === +d2));
        //     console.log((+d1 === +d2));
        //     console.log((+d1 === +d2));
        //     console.log((+d1 === +d2));
        //   console.timeEnd('valueOf'); // 1.09ms
        //   console.time('valueOf explicit');
        //     console.log((d1.valueOf() === d2.valueOf()));
        //     console.log((d1.valueOf() === d2.valueOf()));
        //     console.log((d1.valueOf() === d2.valueOf()));
        //     console.log((d1.valueOf() === d2.valueOf()));
        //     console.log((d1.valueOf() === d2.valueOf()));
        //     console.log((d1.valueOf() === d2.valueOf()));
        //   console.timeEnd('valueOf explicit'); // 0.08ms
        //   console.time('getTIme explicit');
        //     console.log(d1.getTime() === d2.getTime());
        //     console.log(d1.getTime() === d2.getTime());
        //     console.log(d1.getTime() === d2.getTime());
        //     console.log(d1.getTime() === d2.getTime());
        //     console.log(d1.getTime() === d2.getTime());
        //     console.log(d1.getTime() === d2.getTime());
        //   console.timeEnd('getTIme explicit'); // 0.06ms
        // };
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
    }
};
