function SendEditLink() {

  this.form = FormApp.openById('1FUw_hkDrKN_PVS3oJLHGpM13il-Ugyvfhc_Tg5E_JKc'); //form ID

  this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
                                     .getSheetByName("Daily Inventory Data");
  this.scheduleSheetData = this.spreadsheet.getDataRange().getValues();
  this.scheduleSheetIndex = indexSheet(this.scheduleSheetData);

  var startRow = 4;  // First row of data to process
  var lastRow = numberOfRows(this.scheduleSheetData);// figure out what the last row is

  for (var i = 0; i < numRows ; i++) {
    if (data[i][0] != ""){
      var responses = response[i].getEditResponseUrl(); //grabs the url from the form
      sheet.getRange(startRow + i, 10).setValue(responses);

      var EMAIL_SENT = "EMAIL_SENT";// This constant is written in column S for rows for which an email has been sent successfully.

      var currentTime = createPrettyDate(new Date(), 'short');

      var message = "";

      // grab column 3 (the 'First Name' column)
      var dataRange1 = sheet.getRange(3,3,lastRow-startRow+1,1);

      // grab column 4 (the 'Last name' column)
      var dataRange2 = sheet.getRange(3,4,lastRow-startRow+1,1);

      // grab column 4 (the 'email address' column)
      var dataRange3 = sheet.getRange(3,4,lastRow-startRow+1,1);

      // grab column 11 (the Email_Sent column)
      var dataRange4 = sheet.getRange(3, 11,lastRow-startRow+1,1);

      // grab column 10 (the editable url column)
      var dataRange5 = sheet.getRange(3, 10,lastRow-startRow+1,1);

      // Fetch values for each row in the Range.
      var data1 = dataRange1.getValues();
      var data2 = dataRange2.getValues();
      var data3 = dataRange3.getValues();
      var data4 = dataRange4.getValues();
      var data5 = dataRange5.getValues();

      var First = data1[i][0];       // 3 column
      var Last = data2[i][0];       // 4 column
      var Email = data3[i][0];       // 8 column
      var emailSent = data4[i][0];     // 19 column
      var link = data5[i][0];     // 18 column
      if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
        var subject = "Edit Link for Daily Personal Inventory (" + currentTime + ")";
        message = "If you would like to edit today's daily inventory, please visit the following link: "+link+"\n";
        MailApp.sendEmail("xiao.qiao.zhou@gmail.com", subject, message);
        sheet.getRange(startRow + i, 11).setValue(EMAIL_SENT);// Make sure the cell is updated right away in case the script is interrupted
        SpreadsheetApp.flush();
      }
    }
  }
}
