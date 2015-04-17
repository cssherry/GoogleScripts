// Instantiate and run constructor
function runAssignBook() {
  // Change this template to change text in automated email
  var mailInfo = "Hi { firstName },\n\nPlease send your book to { sendToPerson }. Their address is below:\n{ sendAddress }\n\nHappy reading!\n\nSchedule here: https://docs.google.com/spreadsheets/d/1wv54jAwqRxPyWAd8a-m_yLNJo2vHYmjEkfp8TCKRWWY/edit?usp=sharing",
      mailInfoSubject = "[BOOKCLUB] Mailing Instructions (Due in 7 days)",
      nextBookInfo = "Hi { sendToPerson },\n\nExpect to get { newBook } soon from { firstName }\n\nHappy reading!\n\n\n\nSchedule here: https://docs.google.com/spreadsheets/d/1wv54jAwqRxPyWAd8a-m_yLNJo2vHYmjEkfp8TCKRWWY/edit?usp=sharing",
      nextBookInfoSubject = "[BOOKCLUB] You're Next Book's in the Mail",

      assign = new AssignBook(mailInfo, mailInfoSubject, nextBookInfo, nextBookInfoSubject);

  assign.run();
}

// Constructor for assigning book
function AssignBook(mailInfo, mailInfoSubject, nextBookInfo, nextBookInfoSubject) {
  var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet()
                                     .getSheetByName("Schedule");
  this.scheduleSheetData = scheduleSheet.getDataRange().getValues();
  this.scheduleSheetIndex = indexSheet(this.scheduleSheetData);

  var formSheet = SpreadsheetApp.getActiveSpreadsheet()
                                 .getSheetByName("Form Responses 1");
  this.formSheetData = formSheet.getDataRange().getValues();
  this.formSheetIndex = indexSheet(this.formSheetData);
  this.formLastEntryIdx = this.formSheetIndex.length - 1;

  var addressesSheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName("Addresses");
  this.addressesSheetData = addressesSheet.getDataRange().getValues();
  this.addressesSheetIndex = indexSheet(this.addressesSheetData);

  this.mailInfo = mailInfo;
  this.mailInfoSubject = mailInfoSubject;
  this.nextBookInfo = nextBookInfo;
  this.nextBookInfoSubject = nextBookInfoSubject;

  this.nextCycle = findNextCycle(this.scheduleSheetData, this.scheduleSheetIndex);

  this.numberToLetters = {
    0: 'A',
    1: 'B',
    2: 'C',
    3: 'D',
    4: 'E',
    5: 'F',
    6: 'G',
    7: 'H',
    8: 'I',
    9: 'J',
    10: 'K',
    11: 'L',
    12: 'M',
  };
}

AssignBook.prototype.run = function() {
  // If this is the first week, randomly assign books and send out emails
  if (this.nextCycle[0] === 1) {
    this.firstWeekRandomAssignment();
  } else {
    // find what book the form submitter last read
    var bookToMatch = this.oldBook();

    // Create a hash with people's name linked the books they've read
    var peopleAndBooks = this.hashifyPeopleAndBooks();

    // Find "pending" in schedule
    var peopleWhoNeedBooks = this.waitingForBooks();

    // check to make sure pending has not already read book from previous person
    // if so, check all pendings
    var sendToPerson = this.findWhoWillReadNext();

    // if no match, then add note saying 'waitingForMatch', with note that has id of form submission (https://developers.google.com/apps-script/reference/spreadsheet/)


    // if match, find entry and note new book on the way, and send another email letting people know
    // Check to see if people with notes saying "waitingForMatch" can send to person. If not, add Pending with note on formid
    // send to original sender if length of schedule column === 7
  }
};

AssignBook.prototype.firstWeekRandomAssignment = function() {
  var people = this.peopleInfo(),
      peopleArray = shuffle(this.peopleArray());

  for (var i = 0; i < peopleArray.length; i++) {
    var sender = peopleArray[i];
    if (i === peopleArray.length - 1) {
      receivingPerson = peopleArray[0];
    } else {
      receivingPerson = peopleArray[i + 1];
    }

    var contactEmail = people[sender].email,
        subject = this.mailInfoSubject,
        emailTemplate = this.mailInfo,
        sheetName = 'Schedule',
        cellCode = this.numberToLetters[people[receivingPerson].idx] + 2,
        options = {
                    firstName: sender,
                    sendToPerson: receivingPerson,
                    sendAddress: people[receivingPerson].address,
                    note: "Assigned: " + new Date(),
                    message: book,
                  };

    new Email(contactEmail, subject, emailTemplate, sheetName, cellCode, options);
  }
};

AssignBook.prototype.peopleInfo = function() {
  var result = {},
      nameIdx = this.addressesSheetIndex.Name,
      emailIdx = this.addressesSheetIndex.Email,
      addressIdx = this.addressesSheetIndex.Address,
      bookIdx = this.addressesSheetIndex.BookChoices;

  for (var i = 1; i < this.addressesSheetData.length; i++) {
    var book = this.addressesSheetData[i][bookIdx],
        email = this.addressesSheetData[i][emailIdx],
        address = this.addressesSheetData[i][addressIdx],
        name = this.addressesSheetData[i][nameIdx],
        options = {book: book,
                   email: email,
                   address: address,
                   idx: i,
                  };
    result[name] = options;
  }

  return result;
};

AssignBook.prototype.peopleArray = function() {
  var result = [],
      nameIdx = this.addressesSheetIndex.Name;

  for (var i = 1; i < this.addressesSheetData.length; i++) {
    var name = this.addressesSheetData[i][nameIdx];
    result.push(name);
  }

  return result;
};

AssignBook.prototype.oldBook = function() {
  var lastBook;

  if (didNotReadAssignedBook) {

  } else {
    this.formSheetData;
    this.formSheetIndex;
    this.formLastEntryIdx;
  }

  return lastBook;
};

AssignBook.prototype.hashifyPeopleAndBooks = function() {
  this.scheduleSheetData = scheduleSheet.getDataRange().getValues();
  this.scheduleSheetIndex = indexSheet(this.scheduleSheetData);
};

AssignBook.prototype.waitingForBooks = function() {
  this.scheduleSheetData = scheduleSheet.getDataRange().getValues();
  this.scheduleSheetIndex = indexSheet(this.scheduleSheetData);
};

AssignBook.prototype.findWhoWillReadNext = function() {

};

AssignBook.prototype.checkForNewMatches = function() {

};


var indexSheet = function(sheetData) {
  var result = {},
      length = sheetData[0].length;

  for (var i = 0; i < length; i++) {
    result[sheetData[0][i]] = i;
  }

  return result;
};

var shuffle = function (array) {
  var l = this.length + 1;
  while (l--) {
    var r = ~~(Math.random() * l),
        o = this[r];

    this[r] = this[0];
    this[0] = o;
  }

  return this;
};