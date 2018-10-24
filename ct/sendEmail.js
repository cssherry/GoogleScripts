var contextValues = {
  alreadyDeleted: {},
};
var fetchPayload = {
};
var urls = {
};

// Get login cookie
function login() {
  var postOptions = {
    method: 'post',
    payload: {
      email: fetchPayload.email,
      password: sjcl.decrypt(fetchPayload.salt, fetchPayload.password),
    },
    followRedirects: false,
  };
  var loginPage = UrlFetchApp.fetch(urls.login, postOptions);
  var loginCode = loginPage.getResponseCode();
  if (loginCode === 200) { //could not log in.
    return "Couldn't login. Please make sure your username/password is correct.";
  } else if (loginCode === 303 || loginCode === 302) {
    return loginPage.getAllHeaders()['Set-Cookie'];
  }
}

// Get the main page
function getMainPageCT() {
  if (!fetchPayload.cookie) {
    fetchPayload.cookie = login();
  }

  var mainPage = UrlFetchApp.fetch(urls.main,
                                  {
                                    headers: {
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
  // Only run after 7 AM or before 11 PM
  var currentDate = new Date();
  var currentHour = currentDate.getHours();
  if (currentHour <= 7 || currentHour > 23) {
    return;
  }

  contextValues.sheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName('Current');
  contextValues.sheetRange = contextValues.sheet.getDataRange();
  contextValues.sheetData = contextValues.sheetRange.getValues();
  contextValues.sheetNotes = contextValues.sheetRange.getNotes();
  contextValues.sheetIndex = indexSheet(contextValues.sheetData);
  processPreviousListings();

  contextValues.ratings = {};
  contextValues.ratingSheet = SpreadsheetApp.getActiveSpreadsheet()
                                            .getSheetByName('Ratings');
  contextValues.ratingRange = contextValues.ratingSheet.getDataRange();
  contextValues.ratingData = contextValues.ratingRange.getValues();
  contextValues.ratingNotes = contextValues.ratingRange.getNotes();
  contextValues.ratingIndex = indexSheet(contextValues.ratingData);
  contextValues.ratingData.forEach(processOldRatings);

/** NOT WORKING
  // Process FM Listings
  var fmHTML = UrlFetchApp.fetch(urls.fmMain,
                                  {
                                    headers : {
                                      Cookie: fetchPayload.fmCookie,
                                    },
                                  });
  var fmPage = cleanupHTML(fmHTML.getContentText());
  var errorMessage = 'You do not have the required permissions to read topics within this forum';
  if (fmPage.indexOf(errorMessage) === -1) {
    var fmDoc = Xml.parse(fmPage, true).getElement();
    var fmList = getElementsByTagName(fmDoc, 'table');
    var fmItems = getElementsByTagName(fmList[5], 'tr');
    fmItems.forEach(addOrUpdateFm);
  } else {
    removeAndEmail(urls.fmDomain);
    return;
  }
**/

  // Process AC Listings
  // First, figure out which ones are free
  var acToken = UrlFetchApp.fetch(urls.acLogin);
  var headers1 = acToken.getAllHeaders();
  fetchPayload.acCookie = headers1['Set-Cookie'].join('');
  var acToken = acToken.getContentText().match(/type="hidden" value="(.*?)"/)[1];
  var loginAcPage = UrlFetchApp.fetch(urls.acLoginAPI, {
      method: 'post',
      followRedirects: false,
      headers: {
        Cookie: fetchPayload.acCookie,
      },
      payload: {
        userName: fetchPayload.email,
        password: sjcl.decrypt(fetchPayload.salt, fetchPayload.acPassword),
        token: acToken,
        'return': '',
      },
    });
  var loginAcCode = loginAcPage.getResponseCode();
  if (loginAcCode === 200) { //could not log in.
    removeAndEmail(urls.acDomain);
  } else if (loginAcCode === 303 || loginAcCode === 302) {
    // Only get ratings once a day
    var currentHour = currentDate.getHours();
    if (currentHour >= 23) {
      contextValues.ratingMin = contextValues.ratingData.length + 1;
      var acReviewString = UrlFetchApp.fetch(urls.acReviews,
                                      {
                                        headers: {
                                          Cookie: fetchPayload.acCookie,
                                        },
                                      });
      var ratingItems = acReviewString.getContentText().match(/<table.*?>[\s\S]*?<\/table>/g);
      ratingItems.forEach(processRatingItem);
      var firstRow = contextValues.ratingMin;
      var arrayIdx = firstRow - 1;
      var rowLength = contextValues.ratingData.length - arrayIdx;
      var columnLength = contextValues.ratingData[0].length;
      var updateRange = contextValues.ratingSheet.getRange(firstRow, 1, rowLength, columnLength);
      updateRange.setValues(contextValues.ratingData.slice(arrayIdx));
      updateRange.setNotes(contextValues.ratingNotes.slice(arrayIdx));
      return;
    }

    // Process free items
    var acFreeHTML = UrlFetchApp.fetch(urls.acFree,
                                    {
                                      headers: {
                                        Cookie: fetchPayload.acCookie,
                                      },
                                    });
    var acFreePage = cleanupHTML(acFreeHTML.getContentText());
    contextValues.freeAC = {};

    var acFreeDoc = Xml.parse(acFreePage, true).getElement();
    var acFreeItem = getElementByClassName(acFreeDoc, 'newShowPane');
    if (acFreeItem && acFreeItem.length) {
      acFreeItem.forEach(processFreeItems);
    }

    var acHTML = UrlFetchApp.fetch(urls.acMain,
                                    {
                                      headers: {
                                        Cookie: fetchPayload.acCookie,
                                      },
                                    });
    var acPage = cleanupHTML(acHTML.getContentText());

    var acDoc = Xml.parse(acPage, true).getElement();
    var acTable = getElementByClassName(acDoc, 'page_content')[0];
    var acItems = getElementByClassName(acTable, 'ladderrung');
    acItems.forEach(addOrUpdateAc);
  } else {
    removeAndEmail(urls.acDomain);
  }

  // Log in to SF
  var loginSfPage = UrlFetchApp.fetch(urls.sfLogin, {
      method: 'post',
      followRedirects: false,
      payload: {
        email_address: fetchPayload.email_address,
        password: sjcl.decrypt(fetchPayload.salt, fetchPayload.sfPassword),
      },
    });
  var loginSfCode = loginSfPage.getResponseCode();
  var sfHeaders = loginSfPage.getAllHeaders();
  if (loginSfCode === 200) { //could not log in.
    removeAndEmail(urls.sf);
  } else if (loginSfCode === 303 || loginSfCode === 302) {
    fetchPayload.sfCookie = sfHeaders['Set-Cookie'];
    // Process AC Listings
    var sfHTML = UrlFetchApp.fetch(urls.sfMain,
                                    {
                                      headers: {
                                        Cookie: fetchPayload.sfCookie,
                                      },
                                    });
    var sfPage = cleanupHTML(sfHTML.getContentText());
    var sfDoc = Xml.parse(sfPage, true).getElement();
    getElementByClassName(sfDoc, 'showcasebox').forEach(addOrUpdateSf);
  }

  // Process CT Listings
  var mainPage = cleanupHTML(getMainPageCT());
  var doc = Xml.parse(mainPage, true).getElement();
  var mainList = getElementsByTagName(doc, 'ul');
  var items = getElementsByTagName(mainList[3], 'li');
  items.forEach(addOrUpdate);

  try {
    // Process OTL listings
    var otlHTML = UrlFetchApp.fetch(urls.otlMain,
                                    {
                                      headers : {
                                        Cookie: fetchPayload.otlCookie,
                                      },
                                    });
    var otlPage = cleanupHTML(otlHTML.getContentText());

    var otlError = 'On the List (OTL) Seat Filler Memberships';
    if (otlPage.indexOf(otlError) === -1) {
      var otlDoc = Xml.parse(otlPage, true).getElement();
      var otlTable = getElementsByTagName(otlDoc, 'table')[0];
      var otlItems = getElementsByTagName(otlTable, 'tr');
      otlItems.forEach(addOrUpdateOtl);
    } else {
      removeAndEmail(urls.otlDomain);
    }
  } catch (e) {
    removeAndEmail(urls.otlDomain, 'errorLoadingPage');
  }

  addNewCellItemsRow();
  sendEmail();
  updateAllCells();
  archiveExpiredItems();
}

function cleanupHTML(htmlText) {
  return htmlText.match(/<body[\s\S]*?<\/body>/)[0]
                 .replace(/<(no)?script[\s\S]*?<\/(no)?script>/g, '')
                 .replace(/<!--|-->/g, '');
}

// Process previous data, including title and fee in case those change
function processPreviousListings() {
  var titleIdx = contextValues.sheetIndex.Title;
  var idIdx = contextValues.sheetIndex.Url;
  var feeIdx = contextValues.sheetIndex.AdminFee;
  var imageIdx = contextValues.sheetIndex.Image;
  var dateIdx = contextValues.sheetIndex.Date;
  var categoryIdx = contextValues.sheetIndex.Category;
  var locIdx = contextValues.sheetIndex.Location;
  var ratingIdx = contextValues.sheetIndex.Rating;
  contextValues.lastRow = contextValues.sheetData.length;
  contextValues.previousListings = {};

  // Also, get formula for image
  // Get range by row, column, row length, column length
  var imageFormulas = contextValues.sheetRange.getFormulas();
  var previousListingObject = {};
  var urlValue, titleValue, feeValue, dateValue, categoryValue, locValue, currItem;
  for (var i = 1; i < contextValues.lastRow; i++) {
    currItem = contextValues.sheetData[i];
    urlValue = currItem[idIdx].trim();
    if (urlValue) {
      titleValue = currItem[titleIdx];
      feeValue = currItem[feeIdx];
      dateValue = currItem[dateIdx];
      categoryValue = currItem[categoryIdx];
      locValue = currItem[locIdx];
      currItem[imageIdx] = imageFormulas[i][0];
      previousListingObject = {
        row: i,
        title: titleValue,
        fee: feeValue,
        date: dateValue,
        category: categoryValue,
        location: locValue,
        rating: currItem[ratingIdx],
      };
      previousListingObject[titleIdx] = titleValue;
      previousListingObject[feeIdx] = feeValue;
      previousListingObject[idIdx] = urlValue;
      previousListingObject[dateIdx] = dateValue;
      contextValues.previousListings[urlValue] = previousListingObject;
    }
  }
}

function removeAndEmail(domain, errorLoadingPage) {
  for (var oldUrl in contextValues.previousListings) {
    if (contextValues.previousListings.hasOwnProperty(oldUrl) && oldUrl.indexOf(domain) !== -1) {
      delete contextValues.previousListings[oldUrl];
    }
  }

  // Only send it once by storing on "Errors" sheet
  if (!contextValues.errorSheet) {
    contextValues.errorSheet = SpreadsheetApp.getActiveSpreadsheet()
                                             .getSheetByName('Errors');
    contextValues.errorData = contextValues.errorSheet.getDataRange().getValues();
    contextValues.errorIndex = indexSheet(contextValues.errorData);
    contextValues.lastErrorRow = contextValues.errorData.length;
    contextValues.errorDateIdx = contextValues.errorIndex.Date;
    contextValues.errorSitesIdx = contextValues.errorIndex.Sites;
  }

  // If it hasn't been emailed today
  var lastEmailedDate = contextValues.errorData[contextValues.lastErrorRow - 1][contextValues.errorDateIdx];
  if (!lastEmailedDate.toDateString || lastEmailedDate.toDateString() !== new Date().toDateString()) {
    var updateMessage = 'Update ' + domain + ' Token';

    if (errorLoadingPage) {
      updateMessage = 'Error loading page for :' + domain;
    }

    var email = MailApp.sendEmail({
      to: myEmail,
      subject: '[CT] ' + updateMessage,
      htmlBody: updateMessage
                 + ': https://docs.google.com/spreadsheets/d/1AC4XDCtUaCaG7O21w1GpJS59Vxt4QmTBypIjKhBR3TU/edit#gid=0',
    });

    contextValues.lastErrorRow++;
  }

  // Add current page to list of pages needing update
  if (!contextValues.errorData[contextValues.lastErrorRow - 1]) {
    contextValues.errorData[contextValues.lastErrorRow - 1] = [];
  }

  var currentData = contextValues.errorData[contextValues.lastErrorRow - 1][contextValues.errorSitesIdx];
  if (!currentData || currentData.indexOf(domain) === -1) {
    var cells = contextValues.errorSheet.getRange(contextValues.lastErrorRow, 1, 1, 2);
    currentData = currentData ? currentData + ', ' + domain : domain;
    cells.setValues([[new Date(), currentData]]);
  }
}

// Figure out of the page which listings are new
var hasPassedTopic;
function addOrUpdateFm(item) {
  var htmlText = item.toXmlString();
  if (!hasPassedTopic) {
    hasPassedTopic = htmlText.indexOf('Topics') !== -1;
    return;
  }

  var links = getElementsByTagName(item, 'a')
  var aElement = links[0];
  var url = aElement.getAttribute('href').getValue().replace('.', urls.fmDomain).trim();
  var itemInfo = contextValues.previousListings[url];
  if (itemInfo) {
    delete contextValues.previousListings[url];
    contextValues.alreadyDeleted[url] = true;
  } else if (!contextValues.alreadyDeleted[url]) {
    var title = aElement.getText().trim();
    if (!title) {
      return;
    }

    var listingInfo = [];
    listingInfo[contextValues.sheetIndex.Image] = '=Image("' + urls.fmImage + '")';
    listingInfo[contextValues.sheetIndex.Title] = title;
    listingInfo[contextValues.sheetIndex.Rating] = getRating(title);
    listingInfo[contextValues.sheetIndex.AdminFee] = 0;
    listingInfo[contextValues.sheetIndex.Date] = trimHeader(mainInfo[0].getText());
    listingInfo[contextValues.sheetIndex.Category] = 'Movie';
    listingInfo[contextValues.sheetIndex.Location] = aElement.getAttribute('title').getValue();
    listingInfo[contextValues.sheetIndex.Url] = url;
    listingInfo[contextValues.sheetIndex.EventManager] = links[1].getText().trim();
    listingInfo[contextValues.sheetIndex.UploadDate] = new Date();
    newItemsForUpdate.push(listingInfo);
  }
}


// Figure out of the page which listings are new
function addOrUpdateSf(item) {
  var aElement = getElementsByTagName(item, 'a')[0];
  var url = aElement.getAttribute('href').getValue().trim();
  var itemInfo = contextValues.previousListings[url];
  if (itemInfo) {
    delete contextValues.previousListings[url];
    contextValues.alreadyDeleted[url] = true;
  } else if (!contextValues.alreadyDeleted[url]) {
    var itemHtml = item.toXmlString();
    var ImageUrl = itemHtml.match(/background-image:url\(.*?(http:\/\/.*?\.jpg)/i);
    var title = getElementsByTagName(item, 'h2');
    title = title[0].getText().trim();
    if (!title) {
      return;
    }

    var date = getElementByClassName(item, 'date-event');
    var description = getElementByClassName(item, 'internal_content');

    var detailPage = cleanupHTML(UrlFetchApp.fetch(url).getContentText());
    var detailError = 'Sorry, this offer has now ended';
    var price = '', location = '', time = '';
    if (detailPage.indexOf(detailError) === -1) {
      price = detailPage.match(/price_info_box.*?>([\s\S]*?)<\/span>/);
      location = detailPage.match(/location_td.*?>([\s\S]*?)<\/td>/);
      price = price ? trimHtml(price[1]) : '';
      location = location ? trimHtml(location[1]) : '';
      time = ' @ ' + detailPage.match(/<td>(.*?:.*?)<\/td>/)[1];
    }

    var listingInfo = [];
    listingInfo[contextValues.sheetIndex.Image] = '=Image("' + ((ImageUrl && ImageUrl[1]) || '') + '")';
    listingInfo[contextValues.sheetIndex.Title] = title;
    listingInfo[contextValues.sheetIndex.Rating] = getRating(title);
    listingInfo[contextValues.sheetIndex.AdminFee] = price;
    listingInfo[contextValues.sheetIndex.Date] = date[0].getText().trim() + time;
    listingInfo[contextValues.sheetIndex.Category] = trimHtml(description[0].toXmlString());
    listingInfo[contextValues.sheetIndex.Location] = location;
    listingInfo[contextValues.sheetIndex.Url] = url;
    listingInfo[contextValues.sheetIndex.EventManager] = '';
    listingInfo[contextValues.sheetIndex.UploadDate] = new Date();
    newItemsForUpdate.push(listingInfo);
  }
}

function getACUrl(urlEnd) {
  return urls.acDomain + 'member/' + urlEnd.replace(/&return=.*$/, '').trim();
}

function processFreeItems(item) {
  var containsFreeBooking = item.toXmlString().match(/no booking fee/i);
  if (containsFreeBooking) {
    var freeUrl = getElementsByTagName(item, 'a')[0].getAttribute('href').getValue();
    contextValues.freeAC[getACUrl(freeUrl)] = true;
  }
}

function processOldRatings(ratingData, idx) {
  contextValues.ratings[ratingData[contextValues.ratingIndex.URL]] = idx;
  contextValues.ratings[ratingData[contextValues.ratingIndex.Title]] = ratingData;
}

function processRatingItem(item) {
  var fullTitleIdx = contextValues.ratingIndex.FullTitle;
  var titleIdx = contextValues.ratingIndex.Title;
  var locationIdx = contextValues.ratingIndex.Location;
  var ratingIdx = contextValues.ratingIndex.Rating;
  var numberReviewsIdx = contextValues.ratingIndex.NumberReviews;
  var urlIdx = contextValues.ratingIndex.URL;
  var addedIdx = contextValues.ratingIndex.AddedDate;

  var data = [];
  var noteArray;
  var url = item.match(/<a href="(.*?)"/);
  if (url) {
    url = getACUrl(url[1]);
    var rowIdx = contextValues.ratings[url];
    var rating = item.match(/<img /g);
    rating = rating ? rating.length : 0;
    var numberReviews = item.match(/see\s+(\d*)\s+review/i);
    if (numberReviews) {
      numberReviews = +numberReviews[1];
    }

    var fullTitle = item.match(/<h3>(.*?)<\/h3>/);
    if (fullTitle) {
      fullTitle = fullTitle[1];
    }

    if (rowIdx !== undefined) {
      data = contextValues.ratingData[rowIdx];
      noteArray = contextValues.ratingNotes[rowIdx];
      var updated = false;

      if (rating !== data[ratingIdx]) {
        noteArray[ratingIdx] = new Date().toISOString() + ' overwrote: ' + data[ratingIdx] + '\n' + noteArray[ratingIdx];
        data[ratingIdx] = rating;
        updated = true;
      }

      if (numberReviews !== data[numberReviewsIdx]) {
        noteArray[numberReviewsIdx] = new Date().toISOString() + ' overwrote: ' + data[numberReviewsIdx] + '\n' + noteArray[numberReviewsIdx];
        data[numberReviewsIdx] = numberReviews;
        updated = true;
      }

      if (updated === true) {
        contextValues.ratingMin = Math.min(rowIdx + 1, contextValues.ratingMin);
      }
    } else {
      var ratingTitle = fullTitle.split(/\s+at\s+/);
      var title = ratingTitle[0];
      var location = ratingTitle[1];

      data[fullTitleIdx] = fullTitle;
      data[titleIdx] = cleanupTitle(title);
      data[locationIdx] = location;
      data[ratingIdx] = rating;
      data[numberReviewsIdx] = numberReviews;
      data[urlIdx] = url;
      data[addedIdx] = new Date();
      noteArray = [];
      for (var i = 0; i < data.length;) noteArray[i++] = '';

      var currIdx = contextValues.ratingNotes.length;
      contextValues.ratings[url] = currIdx;
      contextValues.ratings[data[titleIdx]] = currIdx;
    }

    contextValues.ratingNotes.push(noteArray);
    contextValues.ratingData.push(data);
  }
}

function cleanupTitle(title) {
  return title.replace(/\s+\(.*?\)\s*$/i, '');
}

function getRating(title) {
  title = cleanupTitle(title);
  return contextValues.ratings[title] ?
         contextValues.ratings[title][contextValues.ratingIndex.Rating] :
               '';
}

// Figure out of the page which listings are new
function addOrUpdateAc(item) {
  var header = getElementByClassName(item, 'showtitle')
  var aElement = getElementsByTagName(header[0], 'a')[0];
  var title = aElement.getText().trim();
  var soldOut = getElementByClassName(item, 'soldOut');

  if (!title) {
    return;
  }

  var rating = getRating(title);
  if (soldOut.length) {
    title += ' (SOLD OUT)';
  }

  var date = getElementByClassName(item, 'dateTime')[0]
              .getText()
              .replace('Check dates and availability...', '')
              .trim();
  var description = getElementByClassName(item, 'showdescription')[0].getText().trim();
  var url = getACUrl(aElement.getAttribute('href').getValue());
  var itemInfo = contextValues.previousListings[url];
  var currentItem = [];
  if (itemInfo) {
    var isNowFree = contextValues.freeAC[url] && itemInfo[contextValues.sheetIndex.AdminFee] !== 'FREE';
    var isNowPaid = !contextValues.freeAC[url] && itemInfo[contextValues.sheetIndex.AdminFee] === 'FREE';
    if (isNowFree || isNowPaid) {
      var newFee = isNowFree ? 'FREE' : '~£3.60';
      var oldFee = isNowFree ? '~£3.60' : 'FREE';
      markCellForUpdate(itemInfo.row, 'AdminFee', newFee);
      currentItem[contextValues.sheetIndex.AdminFee] = newFee + ' <br><em>(Previously ' + oldFee + ')</em>';
    }

    if (date !== itemInfo.date) {
      markCellForUpdate(itemInfo.row, 'Date', date);
      currentItem[contextValues.sheetIndex.Date] = date + '<br><em>(Previously ' + itemInfo.date + ')</em>';
    }

    if (title !== itemInfo.title &&
        title !== "'" + itemInfo.title &&
        title !== "'" + itemInfo.title + "'" &&
        title !== itemInfo.title + "'" &&
        title !== '"' + itemInfo.title &&
        title !== '"' + itemInfo.title + '"' &&
        title !== itemInfo.title + '"') {
      markCellForUpdate(itemInfo.row, 'Title', title);
      currentItem[contextValues.sheetIndex.Title] = title + '<br><em>(Previously ' + itemInfo.title + ')</em>';
    }

    if (rating !== itemInfo.rating) {
      markCellForUpdate(itemInfo.row, 'Rating', rating);
      currentItem[contextValues.sheetIndex.Rating] = rating + '<br><em>(Previously ' + itemInfo.rating + ')</em>';
    }

    if (description !== itemInfo.category &&
        description !== "'" + itemInfo.category &&
        description !== "'" + itemInfo.category + "'" &&
        description !== itemInfo.category + "'" &&
        description !== '"' + itemInfo.category &&
        description !== '"' + itemInfo.category + '"' &&
        description !== itemInfo.category + '"') {
      markCellForUpdate(itemInfo.row, 'Category', description);
      currentItem[contextValues.sheetIndex.Category] = description + '<br><em>(Previously ' + itemInfo.category + ')</em>';
    }

    if (currentItem.length) {
      if (!currentItem[contextValues.sheetIndex.Title]) {
        currentItem[contextValues.sheetIndex.Title] = title;
      }

      if (!currentItem[contextValues.sheetIndex.Date]) {
        currentItem[contextValues.sheetIndex.Date] = date;
      }

      if (!currentItem[contextValues.sheetIndex.Category]) {
        currentItem[contextValues.sheetIndex.Category] = description;
      }

      if (!currentItem[contextValues.sheetIndex.Rating]) {
        currentItem[contextValues.sheetIndex.Rating] = rating;
      }

      currentItem[contextValues.sheetIndex.Location] = itemInfo.location;
      currentItem[contextValues.sheetIndex.Url] = url;
      updatedItems.push(currentItem);
    }

    delete contextValues.previousListings[url];
    contextValues.alreadyDeleted[url] = true;
  } else if (!contextValues.alreadyDeleted[url]) {
    var ImageElements = getElementByClassName(item, 'pic');
    var ImageUrl = ImageElements[1] ? urls.acDomain + ImageElements[1].getAttribute('src').getValue() : '';

    var venue = getElementByClassName(item, 'venue');
    venue = trimHtml(venue[0].toXmlString()).trim();

    var listingInfo = [];
    listingInfo[contextValues.sheetIndex.Image] = '=Image("' + ImageUrl + '")';
    listingInfo[contextValues.sheetIndex.Title] = title;
    listingInfo[contextValues.sheetIndex.Rating] = getRating(title);
    listingInfo[contextValues.sheetIndex.AdminFee] = contextValues.freeAC[url] ? 'FREE' : '~£3.60';
    listingInfo[contextValues.sheetIndex.Date] = date;
    listingInfo[contextValues.sheetIndex.Category] = description;
    listingInfo[contextValues.sheetIndex.Location] = venue;
    listingInfo[contextValues.sheetIndex.Url] = url;
    listingInfo[contextValues.sheetIndex.EventManager] = '';
    listingInfo[contextValues.sheetIndex.UploadDate] = new Date();
    newItemsForUpdate.push(listingInfo);
  }
}

// Figure out of the page which listings are new
function addOrUpdateOtl(item) {
  var className = item.getAttribute('class');
  if (className && className.getValue().indexOf('event-row') === -1) {
    return;
  }

  var links = getElementsByTagName(item, 'a')
  if (!links.length) {
    return;
  }

  var aElement = links[1] || links[0];
  var url = aElement.getAttribute('href').getValue().trim();
  var itemInfo = contextValues.previousListings[url];
  if (itemInfo) {
    delete contextValues.previousListings[url];
    contextValues.alreadyDeleted[url] = true;
  } else if (!contextValues.alreadyDeleted[url]) {
    var htmlText = item.toXmlString();
    var ImageElements = getElementsByTagName(item, 'img');
    var ImageUrl = ImageElements[1].getAttribute('src').getValue();
    if (!ImageUrl && ImageElements[0]) {
      ImageUrl = ImageElements[0].getAttribute('src').getValue();
    }
    var paragraphs = getElementsByTagName(item, 'p');
    var mainInfo = getElementsByTagName(paragraphs[0], 'span');

    var listingInfo = [];
    var title = aElement.getText().trim();

    if (!title) {
      return;
    }

    listingInfo[contextValues.sheetIndex.Image] = '=Image("' + ImageUrl + '")';
    listingInfo[contextValues.sheetIndex.Title] = title;
    listingInfo[contextValues.sheetIndex.Rating] = getRating(title);
    listingInfo[contextValues.sheetIndex.AdminFee] = trimHeader(mainInfo[2].getText());
    listingInfo[contextValues.sheetIndex.Date] = trimHeader(mainInfo[0].getText());
    listingInfo[contextValues.sheetIndex.Category] = paragraphs[1].getText().trim() || '';
    listingInfo[contextValues.sheetIndex.Location] = trimHeader(mainInfo[1].getText());
    listingInfo[contextValues.sheetIndex.Url] = url;
    listingInfo[contextValues.sheetIndex.EventManager] = '';
    listingInfo[contextValues.sheetIndex.UploadDate] = new Date();
    newItemsForUpdate.push(listingInfo);
  }
}


function addOrUpdate(item) {
  // Get href
  var aElement = getElementsByTagName(item, 'a')[0];
  var url = aElement.getAttribute('href').getValue().trim();
  var itemInfo = contextValues.previousListings[url];
  var htmlText = item.toXmlString();
  if (itemInfo) {
    // see if there's anything to update, if not, then just delete
    var title = getTitle(item),
        rating = getRating(title),
        fee = getFee(htmlText),
        date = getDate(htmlText),
        currentItem = [];
    if (fee !== itemInfo.fee) {
      markCellForUpdate(itemInfo.row, 'AdminFee', fee);
      currentItem[contextValues.sheetIndex.AdminFee] = fee + '<br><em>(Previously ' + itemInfo.fee + ')</em>';
    }

    if (date !== itemInfo.date) {
      markCellForUpdate(itemInfo.row, 'Date', date);
      currentItem[contextValues.sheetIndex.Date] = date + '<br><em>(Previously ' + itemInfo.date + ')</em>';
    }

    if (rating !== itemInfo.rating) {
      markCellForUpdate(itemInfo.row, 'Rating', rating);
      currentItem[contextValues.sheetIndex.Rating] = rating + '<br><em>(Previously ' + itemInfo.rating + ')</em>';
    }

    if (title !== itemInfo.title &&
        title !== "'" + itemInfo.title &&
        title !== "'" + itemInfo.title + "'" &&
        title !== itemInfo.title + "'" &&
        title !== '"' + itemInfo.title &&
        title !== '"' + itemInfo.title + '"' &&
        title !== itemInfo.title + '"') {
      markCellForUpdate(itemInfo.row, 'Title', title);
      currentItem[contextValues.sheetIndex.Title] = title + '<br><em>(Previously ' + itemInfo.title + ')</em>';
    }

    if (currentItem.length) {
      if (!currentItem[contextValues.sheetIndex.Title]) {
        currentItem[contextValues.sheetIndex.Title] = title;
      }

      if (!currentItem[contextValues.sheetIndex.Date]) {
        currentItem[contextValues.sheetIndex.Date] = date;
      }

      if (!currentItem[contextValues.sheetIndex.Rating]) {
        currentItem[contextValues.sheetIndex.Rating] = rating;
      }

      currentItem[contextValues.sheetIndex.Category] = itemInfo.category;
      currentItem[contextValues.sheetIndex.Location] = itemInfo.location;
      currentItem[contextValues.sheetIndex.Url] = url;
      updatedItems.push(currentItem);
    }

    delete contextValues.previousListings[url];
    contextValues.alreadyDeleted[url] = true;
  } else if (!contextValues.alreadyDeleted[url]) {
    addNewListing(item, htmlText, url);
  }
}

// Get listing full page
function addNewListing(item, htmlText, url) {
  var ImageUrl = getElementsByTagName(item, 'img')[0].getAttribute('src').getValue();
  var listingInfo = [];
  var title = getTitle(item);

  if (!title) {
    return;
  }

  listingInfo[contextValues.sheetIndex.Image] = '=Image("' + ImageUrl + '")';
  listingInfo[contextValues.sheetIndex.Title] = title;
  listingInfo[contextValues.sheetIndex.Rating] = getRating(title);
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
  return text.replace(/<.*?>|&amp|\n|\r/g, '');
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

function getElementByClassName(element, className) {
  function containsClass(element) {
    var currClass = element.getAttribute('class');
    if (!currClass) {
      return false;
    }

    var currClass = currClass.getValue();
    return currClass === className ||
           currClass.indexOf(' ' + className) !== -1 ||
           currClass.indexOf(className + ' ') !== -1;
  }

  var elList = element.getElements();
  var data = elList.filter(containsClass);

  var i = elList.length;
  while (i--) {
    // (Recursive) Check each child, in document order.
    var found = getElementByClassName(elList[i], className);
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
  var subject = '[CT] *' + newItemsForUpdate.length + '* New || *' + updatedItems.length + '* Updated ' + new Date().toLocaleString();


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
  var ratingIdx = contextValues.sheetIndex.Rating;
  var locationIdx = contextValues.sheetIndex.Location;
  var dateIdx = contextValues.sheetIndex.Date;
  var categoryIdx = contextValues.sheetIndex.Category;
  var urlIdx = contextValues.sheetIndex.Url;
  var feeIdx = contextValues.sheetIndex.AdminFee;
  var imageUrl = listingInfo[imageIdx] ? getImageUrl(listingInfo[imageIdx])  : '';
  var imageDiv = imageUrl ? '<img src="' + imageUrl + '" alt="' + listingInfo[titleIdx] + '" width="128">' :
                 '';
  var rating = listingInfo[ratingIdx] ? ' <small>' + listingInfo[ratingIdx] + '/5</small>' : '';
  var feeDiv = listingInfo[feeIdx] ? (listingInfo[feeIdx] + '<br>') : '';
  var locationDiv = listingInfo[locationIdx] ? (listingInfo[locationIdx] + '<br>') : '';
  var dateDiv = listingInfo[dateIdx] ? (listingInfo[dateIdx] + '<br>') : '';
  var categoryDiv = listingInfo[categoryIdx] ? (listingInfo[categoryIdx] + '<br>') : '';
  return '<h3>' + listingInfo[titleIdx] + rating + '</h3><br>' +
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
function addNewCellItemsRow() {
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
  var lastArchiveRow = archiveData.length;
  var imageIdx = contextValues.sheetIndex.Image;
  var currentTime = new Date();
  for (var expiredItem in contextValues.previousListings) {
    if (contextValues.previousListings.hasOwnProperty(expiredItem) && expiredItem) {
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
    Utilities.sleep(50);
  });
}

// Add item information to specific cell, archiving previous value as note
function markCellForUpdate(row, key, value) {
  var cellColumn = contextValues.sheetIndex[key];
  var previousMessage = contextValues.sheetData[row][cellColumn];
  var oldNote = contextValues.sheetNotes[row][cellColumn];
  var newNote = new Date().toISOString() + ' overwrote: ' + previousMessage + '\n';
  var currentNote = oldNote ? (oldNote + newNote) : newNote;
  contextValues.sheetData[row][cellColumn] = value;
  contextValues.sheetNotes[row][cellColumn] = currentNote;
}

function updateAllCells() {
  if (updatedItems.length) {
    contextValues.sheetRange.setValues(contextValues.sheetData);
    contextValues.sheetRange.setNotes(contextValues.sheetNotes);
  }
}
