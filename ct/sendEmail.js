var contextValues = {};
var fetchPayload = {
};
var urls = {
};

// Get login cookie
function login() {
  var postOptions = {
    method: 'post',
    payload: {
      'email': fetchPayload.email,
      'password': sjcl.decrypt(fetchPayload.salt, fetchPayload.password),
    },
    followRedirects: false,
  };
  var loginPage = UrlFetchApp.fetch(urls.login, postOptions);
  var loginCode = loginPage.getResponseCode();
  if (loginCode === 200) { //could not log in.
    return "Couldn't login. Please make sure your username/password is correct.";
  } else if (loginCode == 303 || loginCode == 302) {
    return loginPage.getAllHeaders()['Set-Cookie'];
  }
}

// Get the main page
function getMainPage() {
  if (!fetchPayload.cookie) {
    fetchPayload.cookie = login();
  }

  var mainPage = UrlFetchApp.fetch(urls.main,
                                  {
                                    headers : {
                                      Cookie: fetchPayload.cookie,
                                    },
                                  });
  return mainPage.getContentText();
}

// Main function for each sheet
// Add to arrays for emailing out later
var updatedItems = [];
var newItemsForUpdate = [];
function updateSheet() {
  contextValues.sheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName('Current');
  contextValues.sheetData = contextValues.sheet.getDataRange().getValues();
  contextValues.sheetIndex = indexSheet(contextValues.sheetData);
  processPreviousListings();
  var mainPage = getMainPage().match(/<body[\s\S]*?<\/body>/)[0]
                              .replace(/<(no)?script[\s\S]*?<\/(no)?script>/g, '')
                              .replace(/<!--|-->/g, '');
  var doc = Xml.parse(mainPage, true).getElement();
  var mainList = getElementsByTagName(doc, 'ul');
  var items = getElementsByTagName(mainList[2], 'li');
  items.forEach(addOrUpdate);
  updateCellRow();
  sendEmail();
  archiveExpiredItems();
}

// Process previous data, including title and fee in case those change
function processPreviousListings() {
  var titleIdx = contextValues.sheetIndex.Title;
  var idIdx = contextValues.sheetIndex.Url;
  var feeIdx = contextValues.sheetIndex.AdminFee;
  var imageIdx = contextValues.sheetIndex.Image;
  var dateIdx = contextValues.sheetIndex.Date;
  contextValues.lastRow = numberOfRows(contextValues.sheetData, titleIdx);
  contextValues.previousListings = {};

  // Also, get formula for image
  // Get range by row, column, row length, column length
  var cells = contextValues.sheet.getRange(1, imageIdx + 1, contextValues.lastRow);
  var imageFormulas = cells.getFormulas();
  var previousListingObject = {};
  var urlValue, titleValue, feeValue, dateValue;
  for (var i = 1; i < contextValues.lastRow; i++) {
    urlValue = contextValues.sheetData[i][idIdx];
    titleValue = contextValues.sheetData[i][titleIdx];
    feeValue = contextValues.sheetData[i][feeIdx];
    dateValue = contextValues.sheetData[i][dateIdx];
    contextValues.sheetData[i][imageIdx] = imageFormulas[i][0];
    previousListingObject = {
      row: i,
      title: titleValue,
      fee: feeValue,
      date: dateValue,
    };
    previousListingObject[titleIdx] = titleValue;
    previousListingObject[feeIdx] = feeValue;
    previousListingObject[idIdx] = urlValue;
    previousListingObject[dateIdx] = dateValue;
    contextValues.previousListings[urlValue] = previousListingObject;
  }
}

// Figure out of the page which listings are new
function addOrUpdate(item) {
  // Get href
  var aElement = getElementsByTagName(item, 'a')[0];
  var url = aElement.getAttribute('href').getValue();
  var itemInfo = contextValues.previousListings[url];
  var htmlText = item.toXmlString();
  if (itemInfo) {
    // see if there's anything to update, if not, then just delete
    var title = getTitle(item),
        fee = getFee(htmlText),
        date = getDate(htmlText),
        currentItem = [];
    if (fee !== itemInfo.fee) {
      updateCell(itemInfo.row + 1, 'AdminFee', fee);
      currentItem[contextValues.sheetIndex.AdminFee] = fee + '<br>(Previously ' + itemInfo.fee + ')';
    }

    if (date !== itemInfo.date) {
      updateCell(itemInfo.row + 1, 'Date', date);
      currentItem[contextValues.sheetIndex.Date] = date + '<br>(Previously ' + itemInfo.date + ')';
    }

    if (title !== itemInfo.title) {
      updateCell(itemInfo.row + 1, 'Title', title);
      currentItem[contextValues.sheetIndex.Title] = title + '<br>(Previously ' + itemInfo.title + ')';
    }

    if (currentItem.length) {
      if (!currentItem[contextValues.sheetIndex.Title]) {
        currentItem[contextValues.sheetIndex.Title] = title;
      }

      currentItem[contextValues.sheetIndex.Url] = url;
      updatedItems.push(currentItem);
    }

    delete contextValues.previousListings[url];
  } else {
    addNewListing(item, htmlText, url);
  }
}

// Get listing full page
function addNewListing(item, htmlText, url) {
  var ImageUrl = getElementsByTagName(item, 'img')[0].getAttribute('src').getValue();
  var listingInfo = [];
  listingInfo[contextValues.sheetIndex.Image] = '=Image("' + ImageUrl + '")';
  listingInfo[contextValues.sheetIndex.Title] = getTitle(item);
  listingInfo[contextValues.sheetIndex.AdminFee] = getFee(htmlText);
  listingInfo[contextValues.sheetIndex.Date] = getDate(htmlText);
  listingInfo[contextValues.sheetIndex.Category] = getColonSeparatedText(htmlText, 'Category');
  listingInfo[contextValues.sheetIndex.Location] = getColonSeparatedText(htmlText, 'Location');
  listingInfo[contextValues.sheetIndex.Url] = url;
  listingInfo[contextValues.sheetIndex.EventManager] = getColonSeparatedText(htmlText, 'Event Manager');
  listingInfo[contextValues.sheetIndex.UploadDate] = new Date();
  newItemsForUpdate.push(listingInfo);
}

// Parse with text
function getTitle(item) {
  return getElementsByTagName(getElementsByTagName(item, 'h4')[0], 'a')[0].getText().trim();
}

function getFee(htmlText) {
  return getColonSeparatedText(htmlText, 'Admin Fee');
}

function getDate(htmlText) {
  return getColonSeparatedText(htmlText, 'Event Date');
}

function getColonSeparatedText(text, expression) {
  var regexExpr = new RegExp(expression + '\\s*:[\\s\\S]*?</p>', 'im');
  var match = text.match(regexExpr);
  if (match) {
    return trimHeader(trimHtml(match[0]));
  }

  return 'None';
}

function trimHtml(text) {
  return text.replace(/<.*?>|&amp/g, '');
}

function trimHeader(text) {
  return text.replace(/[\s\S]*?:/, '').trim();
}

function getImageUrl(imageFormula) {
  return imageFormula.slice(0, imageFormula.length - 2).replace(/=image\("?'?/i, '')
}

// Work with HTML
function getElementsByTagName(element, tagName) {
  var data = element.getElements(tagName);
  var elList = element.getElements();
  var i = elList.length;
  while (i--) {
    // (Recursive) Check each child, in document order.
    var found = getElementsByTagName(elList[i], tagName);
    if (found) {
      data = data.concat(found);
    }
  }
  return data;
}

// Send email with new listing information
function sendEmail() {
  // Only send if there's new items
  if (!updatedItems.length && !newItemsForUpdate.length) return;

  var footer = '<hr>' +
  var newItemsText = newItemsForUpdate.length ? '<hr><h2>New:</h2><br>' + newItemsForUpdate.map(getElementSection).join('') : '';
  var updatedItemsText = updatedItems.length ? '<hr><h2>Updated:</h2><br>' + updatedItems.map(getElementSection).join('') : '';
  var archivedItemsText = '';
  if (Object.keys(contextValues.previousListings).length) {
    archivedItemsText = '<hr><h2>Archived:</h2><br>';
    for (var showUrl in contextValues.previousListings) {
      if (contextValues.previousListings.hasOwnProperty(showUrl)) {
        archivedItemsText += getElementSection(contextValues.previousListings[showUrl]);
      }
    }
  }

  var emailTemplate = newItemsText +
                      updatedItemsText +
                      archivedItemsText +
                      footer;
  var subject = '[CT] *' + updatedItems.length + '* Updated || *' + newItemsForUpdate.length + '* New ' + new Date().toLocaleString();


  // Get information from TotalSavings tab
  var email = MailApp.sendEmail({
    to: myEmail,
    subject: subject,
    htmlBody: emailTemplate,
  });
}

function getElementSection(listingInfo) {
  var imageIdx = contextValues.sheetIndex.Image;
  var titleIdx = contextValues.sheetIndex.Title;
  var locationIdx = contextValues.sheetIndex.Location;
  var dateIdx = contextValues.sheetIndex.Date;
  var categoryIdx = contextValues.sheetIndex.Category;
  var urlIdx = contextValues.sheetIndex.Url;
  var feeIdx = contextValues.sheetIndex.AdminFee;
  var imageUrl = listingInfo[imageIdx] ? getImageUrl(listingInfo[imageIdx])  : '';
  var imageDiv = imageUrl ? '<img src="' + imageUrl + '" alt="' + listingInfo[titleIdx] + '" width="128">' :
                 '';
  var feeDiv = listingInfo[feeIdx] ? (listingInfo[feeIdx] + '<br>') : '';
  var locationDiv = listingInfo[locationIdx] ? (listingInfo[locationIdx] + '<br>') : '';
  var dateDiv = listingInfo[dateIdx] ? (listingInfo[dateIdx] + '<br>') : '';
  var categoryDiv = listingInfo[categoryIdx] ? (listingInfo[categoryIdx] + '<br>') : '';
  return '<h3>' + listingInfo[titleIdx] + '</h3><br>' +
         feeDiv +
         locationDiv +
         dateDiv +
         categoryDiv +
         '<br>' +
         imageDiv +
         '<br><br>' +
         'Url: <a href="' + listingInfo[urlIdx] + '" target="_blank">' + listingInfo[urlIdx] + '</a>' +
         '<hr>';
}

// Function that updates sheet
function updateCellRow() {
  if (!newItemsForUpdate.length) return;

  // Get range by row, column, row length, column length
  var cells = contextValues.sheet.getRange((contextValues.lastRow + 1), 1, newItemsForUpdate.length, newItemsForUpdate[0].length);
  cells.setValues(newItemsForUpdate);
}

// Move expired items to "Archive" sheet
function archiveExpiredItems() {
  // Now archive events that passed
  var cutRange, newRange, currentItem, row, oldValues, oldNotes;
  var toDelete = [];
  var archive = SpreadsheetApp.getActiveSpreadsheet()
                                        .getSheetByName('Archive');
  var archiveData = archive.getDataRange().getValues();
  var lastArchiveRow = numberOfRows(archiveData);
  var imageIdx = contextValues.sheetIndex.Image;
  var currentTime = new Date();
  for (var expiredItem in contextValues.previousListings) {
    if (contextValues.previousListings.hasOwnProperty(expiredItem)) {
      lastArchiveRow++;
      currentItem = contextValues.previousListings[expiredItem];
      row = currentItem.row + 1;
      cutRange = contextValues.sheet.getRange('A' + row + ':I' + row);
      newRange = archive.getRange('A' + lastArchiveRow + ':J' + lastArchiveRow)
      oldValues = cutRange.getValues();
      oldValues[0][imageIdx] = getImageUrl(contextValues.sheetData[currentItem.row][imageIdx]);
      oldValues[0].push(currentTime);
      newRange.setValues(oldValues);
      oldNotes = cutRange.getNotes();
      oldNotes[0].push('');
      newRange.setNotes(oldNotes);
      toDelete.push({
        range: cutRange,
        row: row,
      });
    }
  }

  toDelete.sort(function sortByRow(a, b){
    return b.row - a.row;
  }).forEach(function deleteItem(rangeToDelete) {
    rangeToDelete.range.deleteCells(SpreadsheetApp.Dimension.ROWS);
    Utilities.sleep(200);
  });
}

// Add item information to specific cell, archiving previous value as note
function updateCell(row, key, value) {
  var cellColumn = contextValues.sheetIndex[key];
  if (cellColumn !== undefined) {
    var cellCode = NumberToLetters(cellColumn) + row;
    var cell = contextValues.sheet.getRange(cellCode);
    var previousMessage = cell.getValue();
    if (previousMessage) {
      var oldNote = cell.getNote();
      var previousMessage = new Date().toISOString() + ' overwrote: ' + previousMessage + '\n';
      var currentNote = (oldNote ? oldNote + previousMessage : previousMessage );
      cell.setNote(currentNote);
    }

    cell.setValue(value);
  }
}
