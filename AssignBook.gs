// Instantiate and run constructor
function runAssignBook() {
  // Change this template to change text in automated email
  var mailInfo = "Hi { firstName },\n\nPlease send your book to { sendToPerson }; address below:\n{ sendAddress }\n\nHappy reading!",
      mailInfoSubject = "[BOOKCLUB] Due Soon: Mailing Instructions",
      nextBookInfo = "Hi { sendToPerson },\n\nExpect to get { newBook } soon from { firstName }. Don't forget to fill out the Google Form once you're done reading the book (https://docs.google.com/forms/d/1j6oYWu4QcadddV2VD0hBQ7XUVbYnwUrAkgowP_jXSaQ/viewform).\n\nHappy reading!\n\n",
      nextBookInfoSubject = "[BOOKCLUB] Your Next Book",

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

  var addressesSheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName("Addresses");
  this.addressesSheetData = addressesSheet.getDataRange().getValues();
  this.addressesSheetIndex = indexSheet(this.addressesSheetData);

  this.mailInfo = mailInfo;
  this.mailInfoSubject = mailInfoSubject;
  this.nextBookInfo = nextBookInfo;
  this.nextBookInfoSubject = nextBookInfoSubject;
}

AssignBook.prototype.run = function() {
  // If this is the first month, randomly assign books and send out emails
  var cycleNumber = numberOfRows(this.scheduleSheetData, this.scheduleSheetIndex.Sherry);
  if (cycleNumber === 1) {
    this.firstWeekRandomAssignment();
  } else {
    // Create an object with needNewBook, needSendBook, and readingHistory properties
    var formResponseInfo = this.reviewFormResponseSheet();

    // Check to make sure first person waiting for book
    // Has not already read book from form submitter
    // If so, check next person waiting for book
    this.assignReaders(formResponseInfo);
  }
};

AssignBook.prototype.firstWeekRandomAssignment = function() {
  // Randomize people
  var people = shuffle(this.peopleInfo());

  // Loop through everyone
  for (var i = 0; i < people.length; i++) {
    // assign sender/receiver
    var sender = people[i];

    if (i === people.length - 1) {
      receiver = people[0];
    } else {
      receiver = people[i + 1];
    }

    // Create info for email
    var contactEmail = sender.email,
        subject = this.mailInfoSubject,
        emailTemplate = this.mailInfo,
        cellCode = NumberToLetters[receiver.idx] + 2,
        emailOptions = {
                    firstName: sender.name,
                    sendToPerson: receiver.name,
                    sendAddress: receiver.address,
                  },
        updateCellOptions = {
                              note: "Assigned: " + new Date(),
                              message: sender.book,
                              cellCode: cellCode,
                              sheetName: 'Schedule'
                            };
    // Send email
    new Email(contactEmail, subject, emailTemplate, emailOptions, [updateCellOptions]);
  }
};

// Array of people with their email, address, and name
AssignBook.prototype.peopleInfo = function() {
  var result = [],
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
                   name: name
                  };
    result.push(options);
  }

  return result;
};

// needNewBook: array of form submitter + everyone who has nothing in the HasNewBook column
// needSendBook: array of people with nothing in the WhoWillReadNext column
// readingHistory: hash of book names with people who have read it, and how many people have read it
// hash to book & cell info for needNewBook and needSendBook objects
AssignBook.prototype.reviewFormResponseSheet = function() {
    var whoWillReadNextIndex = this.formSheetIndex.WhoWillReadNext,
        hasNewBookIndex = this.formSheetIndex.HasNewBook,
        bookIndex = this.formSheetIndex['Which book did you just finish reading?'],
        nameIndex = this.formSheetIndex.Name,
        result = {needNewBook: [], needSendBook: [], readingHistory: {}, needNewBookHash: {}, needSendBookHash: {}},
        cell, userBookObject,
        numberEntries = numberOfRows(this.formSheetData);

    for (var i = 1; i < numberEntries; i++) {
      var book = this.formSheetData[i][bookIndex],
          name = this.formSheetData[i][nameIndex];

      // add name/book/cell to needSendBook object if WhoWillReadNext empty
      // add sanity check so book doesn't get added twice
      if (!this.formSheetData[i][whoWillReadNextIndex]) {
        cell = NumberToLetters[whoWillReadNextIndex] + (i + 1);
        userBookObject = {name: name, book: book, cell: cell};
        // Only add book once. If book appears more than once, send me troubleshooting email
        if (result.needSendBookHash[book]) {
          this.sendErrorMessage(result.needSendBookHash[book], userBookObject, "Duplicate book being sent");
        } else {
          result.needSendBookHash[book] = userBookObject;
          result.needSendBook.push(userBookObject);
        }
      }

      // add name/book/cell to needNewBook object if HasNewBook empty
      // add sanity check so user isn't added twice
      if (!this.formSheetData[i][hasNewBookIndex]) {
        cell = NumberToLetters[hasNewBookIndex] + (i + 1);
        userBookObject = {name: name, book: book, cell: cell};
        // Only add name once. If name appears more than once, send me troubleshooting email
        if (result.needSendBookHash[name]) {
          this.sendErrorMessage(result.needNewBookHash[name], userBookObject, "Duplicate person recieving books");
        } else {
          result.needNewBookHash[name] = userBookObject;
          result.needNewBook.push(userBookObject);
        }
      }

      // Add name and book to reading history
      if (!result.readingHistory[book]) {
        result.readingHistory[book] = {numberTimesRead: 0};
      }

      result.readingHistory[book][name] = true;
      result.readingHistory[book].numberTimesRead++;
    }

    return result;
};

AssignBook.prototype.assignReaders = function(formSheetData) {
  // Go through everyone who needs to send a book (first to last)
  // Try to pair with people who need a book (last to first)
  var sender,
      receiver,
      numberNeedSendBook = formSheetData.needSendBook.length;

  for (var k = 0; k < numberNeedSendBook; k++) {
    sender = formSheetData.needSendBook[k];
    numberNeedSendBook = formSheetData.needSendBook.length;
    var book = sender.book,
        maxTimeBooksRead = numberOfRows(this.addressesSheetData) - 2,
        numberTimesBookRead = formSheetData.readingHistory[book].numberTimesRead;

    if (numberTimesBookRead >= maxTimeBooksRead) {
      // If everyone has already read this book, send it back
      var bookIndex = this.addressesSheetIndex.BookChoices,
          nameIndex = this.addressesSheetIndex.Name;

      for (var j = 1; j < this.addressesSheetData.length; j++) {
        if (book === this.addressesSheetData[j][bookIndex]) {
          var name = this.addressesSheetData[j][nameIndex];
          receiver = {name: name};
          break;
        }
      }

      formSheetData.needSendBook.splice(k, 1);
      this.bookAssigned(sender, receiver);
    } else {

      var needNewBookMaxIdx = formSheetData.needNewBook.length - 1;

      // Else try to find a match
      for (var i = needNewBookMaxIdx; i > -1 ; i--) {
        receiver = formSheetData.needNewBook[i];

        var receiverIdx = this.scheduleSheetIndex[receiver.name],
            bookIdx = this.addressesSheetIndex.BookChoices,
            receiversOriginalBook = this.addressesSheetData[receiverIdx][bookIdx],
            receiverReadBook = formSheetData.readingHistory[book][receiver.name];

        // Skip person if it's their book or if they've already read it
        if (receiversOriginalBook !== book && !receiverReadBook) {
          formSheetData.needNewBook.splice(i, 1);
          this.bookAssigned(sender, receiver);
          break;
        }
      }
    }
  }
};

AssignBook.prototype.bookAssigned = function(sender, receiver) {
  var peopleInfo = this.peopleInfo(),
      receiverInfo,
      senderInfo;

  // get receiverInfo and senderInfo
  while (!senderInfo || !receiverInfo) {
    for (var j = 0; j < peopleInfo.length; j++) {
      if (peopleInfo[j].name === sender.name) {
        senderInfo = peopleInfo[j];
      }
      if (peopleInfo[j].name === receiver.name) {
        receiverInfo = peopleInfo[j];
      }
    }
  }

  // Send out receiver's address to sender
  // firstName: sender
  // sendToPerson: receiver
  // sendAddress: receiver's address
  var contactEmail1 = senderInfo.email,
      emailOptions1 = {
                  sendAddress: receiverInfo.address,
                  sendToPerson: receiver.name,
                  firstName: sender.name
                },
      updateCellOptions1 = {
                            note: "Assigned: " + new Date(),
                            message: receiver.name,
                            cellCode: sender.cell,
                            sheetName: 'Form Responses 1'
                          };

  new Email(contactEmail1, this.mailInfoSubject, this.mailInfo, emailOptions1, [updateCellOptions1]);

  // Send heads up email
  // sendToPerson: receiver's name
  // newBook: sender's books name
  // firstName: sender's name
  var contactEmail2 = receiverInfo.email,
      lastIdx = numberOfRows(this.scheduleSheetData, receiverInfo.idx),
      cell = NumberToLetters[receiverInfo.idx] + (lastIdx + 1),
      emailOptions2 = {note: "Assigned: " + new Date(),
                        message: sender.book,
                        sendToPerson: receiver.name,
                        newBook: sender.book,
                        firstName: sender.name
                      },
      updateCellsOptions2 = [{
                            note: "Assigned: " + new Date(),
                            message: sender.book,
                            cellCode: cell,
                            sheetName: 'Schedule'
                          }];

  if (receiver.cell) {
    updateCellsOptions2.push({
      note: "Reminder sent: " + new Date(),
      message: "TRUE",
      cellCode: receiver.cell,
      sheetName: 'Form Responses 1'
    });
  }

  new Email(contactEmail2, this.nextBookInfoSubject, this.nextBookInfo, emailOptions2, updateCellsOptions2);
};

AssignBook.prototype.sendErrorMessage = function (duplicate1, duplicate2, errortype) {
  // Send out receiver's address to sender
  // firstName: sender
  // sendToPerson: receiver
  // sendAddress: receiver's address
  var contactEmail1 = myEmail,
      mailSubject = "[BOOKCLUB] ERROR: " + errortype,
      mailBody = "There are duplicate form entries:\n" +
                 "First entry: name ( { name1 } ), book ( { book1 } ), cell ( { cell1 } )\n" +
                 "Second entry: name ( { name2 } ), book ( { book2 } ), cell ( { cell2 } )\n",
      emailOptions1 = {
                  name1: duplicate1.name,
                  book1: duplicate1.book,
                  cell1: duplicate1.cell,
                  name2: duplicate2.name,
                  book2: duplicate2.book,
                  cell2: duplicate2.cell
                };

  new Email(contactEmail1, this.mailInfoSubject, this.mailInfo, emailOptions1);

};