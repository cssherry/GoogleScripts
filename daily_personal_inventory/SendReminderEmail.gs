var EMAIL_SENT = "EMAIL_SENT";// This constant is written in column AI for rows for which an email has been sent successfully.

function sendReminderEmail() {

  var message1 = "";
  var message2 = "";
  var n=0;

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();// get the spreadsheet object
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[2]);// set the third sheet as active
  var sheet = spreadsheet.getActiveSheet(); // fetch this sheet

  var startRow = 4;  // First row of data to process
  var lastRow = 33;// go for a month

  var dataRange0 = sheet.getRange(4,2,lastRow-startRow+1,1);// grab column 2 (the 'missing days' column)
  var numRows = dataRange0.getNumRows(); // Number of rows to process

  var data0 = dataRange0.getValues();// Fetch values for each row in the Range.

    for (var i = 0; i <= numRows - 1; i++) {
    var days_from = data0[i][0];
    if(days_from != 1) {
      var row = data0[i];
      var message1 = message1 +"     "+ row + "\n";
      var n = n+1;
      }
    }
  if (n > 0){
    var message1 = "You missed "+n+" days this month in the Daily Personal Inventory (https://docs.google.com/spreadsheet/viewform?usp=drivesdk&formkey=dFhPMklFeUtJY0RCUWhaVUF1QW52MVE6MA#gid=4)\n"+message1+ "\n";
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();// get the spreadsheet object
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[0]);// set the first sheet as active
  var sheet = spreadsheet.getActiveSheet(); // fetch this sheet

  var startRow = 4;  // First row of data to process
  var lastRow = sheet.getLastRow();// figure out what the last row is

  //Grab info from relevant columns
  var dataRange1 = sheet.getRange(4,2,lastRow-startRow+1,1);  // grab column 2 (the 'days from today' column)
  var dataRange2 = sheet.getRange(4,7,lastRow-startRow+1,1);  // grab column 7 (the 'What are tomorrow's goals?' column)
  var dataRange3 = sheet.getRange(4,8,lastRow-startRow+1,1);  // grab column 8 (the 'What are your life goals?' column)
  var dataRange4 = sheet.getRange(4,4,lastRow-startRow+1,1);  // grab column 4 (the 'What are you grateful for?' column)
  var dataRange5 = sheet.getRange(4,5,lastRow-startRow+1,1);  // grab column 5 (the 'What are you grateful for?' column)
  var dataRange6 = sheet.getRange(4,6,lastRow-startRow+1,1);  // grab column 6 (the 'What are you grateful for?' column)
  var dataRange7 = sheet.getRange(4,1,lastRow-startRow+1,1);  // grab column 1 (the 'Timestamp' column)
  var dataRange8 = sheet.getRange(4,39,lastRow-startRow+1,1);  // grab column 36 (the Email_Sent column)

  var numRows = dataRange1.getNumRows(); // Number of rows to process

  // Fetch values for each row in the Range.
  var data1 = dataRange1.getValues();
  var data2 = dataRange2.getValues();
  var data3 = dataRange3.getValues();
  var data4 = dataRange4.getValues();
  var data5 = dataRange5.getValues();
  var data6 = dataRange6.getValues();
  var data7 = dataRange7.getValues();
  var data8 = dataRange8.getValues();

  var message3="";

  var m = 0
  //Creates message with random grateful thing
  for (var j = numRows-30; j <= numRows-1 ; j++) {
    if(data4[j][0] != "" && m < 4) {
      var days_fromb = data1[j][0];          // 2 column
      var grateful1b = data4[j][0];       // 4 column
      var grateful2b = data5[j][0];      // 5 column
      var grateful3b = data6[j][0];       // 6 column
      var timeb = data7[j][0];       // 1 column
      var message3= message3+"\n\nRemember "+days_fromb+" ago you were thankful for: \n     1) "+grateful1b+" \n     2) "+grateful2b+" \n     3) "+grateful3b+" \n\nAs reported on "+timeb;
      m++
    }
  }

  var yesterday=1;

  //Creates email message for today
  for (var i = 0; i <= numRows - 1; i++) {
    var days_from = data1[i][0];
    if(days_from == 1) {
        var goal = data2[i][0];       // 7 column
        var life = data3[i][0];       // 8 column
        var grateful1 = data4[i][0];       // 4 column
        var grateful2 = data5[i][0];      // 5 column
        var grateful3 = data6[i][0];       // 6 column
        var time = data7[i][0];       // 1 column
        var emailSent = data8[i][0];     // 36 column
        yesterday++;
          if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
            var subject = "To-Do's for Today ("+currentTime+")";
            var message2 = "Don't forget! !1 for Urgent&Important (DO NOW!), !2 for Urgent (Try to Delegate), !3 for Important (Decide when to do), !4 for do later\nUse #e0 for things to do as a break\n\nToday's To-Dos:\n"+goal+" \n\nLife goals: \n"+life+" \n\nRemember to be thankful for: \n     1) "+grateful1+" \n     2) "+grateful2+" \n     3) "+grateful3+" \n\nAs reported on "+time;
            MailApp.sendEmail("xiao.qiao.zhou@gmail.com", subject, message1+message2+message3);
            sheet.getRange(startRow + i, 39).setValue(EMAIL_SENT);// Make sure the cell is updated right away in case the script is interrupted
            SpreadsheetApp.flush();
          }
    }
  }
  if (yesterday==1 && n > 0){
    var subject = "You are Missing "+n+" days as of "+currentTime;
    MailApp.sendEmail("xiao.qiao.zhou@gmail.com", subject, "Don't forget! !1 for Urgent&Important (DO NOW!), !2 for Urgent (Try to Delegate), !3 for Important (Decide when to do), !4 for do later\nUse #e0 for things to do as a break\n\n"+message1+"\n\n"+message3);
    SpreadsheetApp.flush();
    }
}
