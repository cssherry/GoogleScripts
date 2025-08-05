var startHour = 6; // This is the hour ratings will be calculated
var endHour = 24;
var contextValues = {
  alreadyDeleted: {},
  currAcItems: {},
};

const date = new Date();
const currYear = date.getFullYear();

var fetchPayload = {





// =====================================
// Main function for each sheet
// =====================================

// Add to arrays for emailing out later
var updatedItems = [];
var newItemsForUpdate = [];
function updateSheet() {
  // Only run after 7 AM or before 11 PM
  var currentDate = new Date();
  var currentHour = currentDate.getHours();
  if (currentHour < startHour || currentHour >= endHour) {
    return;
  }

  // Get ratings for location
  contextValues.locationRatings = {};
  contextValues.locationRatingData = SpreadsheetApp.getActiveSpreadsheet()
                                                   .getSheetByName('ratingAnalysis')
                                                   .getDataRange()
                                                   .getValues();
  contextValues.locationRatingIndex = indexSheet(contextValues.locationRatingData);
  contextValues.locationRatingData.forEach(processOldLocationRatings);

  // Process previous values
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
  // // Remove duplicates
  // contextValues.ratingRange.setValues(contextValues.ratingData);

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
    var fmDoc = XmlService.parse(fmPage).getRootElement();
    var fmList = getElementsByTagName(fmDoc, 'table');
    var fmItems = getElementsByTagName(fmList[5], 'tr', true);
    fmItems.forEach(addOrUpdateFm);
  } else {
    removeAndEmail(urls.fmDomain);
    return;
  }

  // Figure out of the page which listings are new
  var hasPassedTopic;
  function addOrUpdateFm(item) {
    var htmlText = item.getText();
    if (!hasPassedTopic) {
      hasPassedTopic = htmlText.indexOf('Topics') !== -1;
      return;
    }

    var links = getElementsByTagName(item, 'a', true)
    var aElement = links[0];
    var title = aElement.getText().trim();
    if (!title) {
      return;
    }

    var url = aElement.getAttribute('href').getValue().replace('.', urls.fmDomain).trim();
    var itemInfo = contextValues.previousListings[url];
    if (itemInfo) {
      checkRatingAndDeletePreviousListing(itemInfo, url, [], title);
    } else if (!contextValues.alreadyDeleted[url]) {
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
**/

  // --------------------------------------------
  // RATING MANAGEMENT LISTINGS
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
    // Only get ratings once a day or whenever we're not getting new events
    var isFirstRun = currentHour <= startHour;
    var lastIdx = contextValues.sheetData[0].length - 1;
    var customTimeToRun = currentHour % contextValues.sheetData[0][lastIdx] === 0;
    if (isFirstRun || customTimeToRun) {
      contextValues.ratingMin = contextValues.ratingData.length + 1;
      var acReviewHTML = UrlFetchApp.fetch(urls.acReviews,
                                      {
                                        headers: {
                                          Cookie: fetchPayload.acCookie,
                                        },
                                      });
      var acReviewPage = cleanupHTMLElement(acReviewHTML);
      var ratingItems = acReviewPage.match(/<div.*?card h-100([\s\S]+?)(<\/div>\s+){3}/g);
      contextValues.newRatings = [];
      contextValues.updatedRatings = [];
      ratingItems.forEach(processRatingItem);
      var firstRow = contextValues.ratingMin;
      var arrayIdx = firstRow - 1;
      var rowLength = contextValues.ratingData.length - arrayIdx;
      var columnLength = contextValues.ratingData[0].length;

      if (rowLength > 0) {
        var updateRange = contextValues.ratingSheet.getRange(firstRow, 1, rowLength, columnLength);
        updateRange.setValues(contextValues.ratingData.slice(arrayIdx));
        updateRange.setNotes(contextValues.ratingNotes.slice(arrayIdx));
      }

      // It's nice to have email of updates
      var ratingMessage = '';
      if (contextValues.newRatings.length) {
        ratingMessage = '<h2>New Ratings</h2>' +
                        contextValues.newRatings.join('<hr><br>') + '<hr>';
      }

      if (contextValues.updatedRatings.length) {
        ratingMessage += '<h2>Updated Ratings</h2>' +
                        contextValues.updatedRatings.join('<hr><br>') + '<hr>';
      }

      if (isFirstRun) {
        MailApp.sendEmail({
          to: myEmail,
          subject: '[CT] New Ratings (' + contextValues.newRatings.length +
                   ') | Updated Ratings (' + contextValues.updatedRatings.length + ')',
          htmlBody: ratingMessage
                      + 'Link: ' + spreadsheetURL,
        });
      }

      return;
    }

    try {
      // --------------------------------------------
      // AC LISTINGS
      // Process new items
      var acHTML = UrlFetchApp.fetch(urls.acMain,
                                     {
                                       headers: {
                                         Cookie: fetchPayload.acCookie,
                                       },
                                     });
      var acPage = cleanupHTMLElement(acHTML);
      var acDoc = XmlService.parse(acPage).getRootElement();
      var acItems = getElementByClassName(acDoc, 'showlisting');

      if (!acItems.length) {
        throw new Error('Nothing found for AC -- something is probably wrong');
      }

      acItems.forEach(parseAcItems);
      Object.keys(contextValues.currAcItems).forEach(addOrUpdateAc)
    } catch (e) {
      printError(e)
      removeAndEmail(urls.acDomain, `errorParsingPage: ${e}`);
    }
  } else {
    removeAndEmail(urls.acDomain);
  }

  // --------------------------------------------
  // PBP LISTINGS
  /*
  try {
    // Process PBP listings
    var pbpHTML = UrlFetchApp.fetch(urls.pbpShows,
                                    {
                                      headers : {
                                        Cookie: fetchPayload.pbpCookie,
                                      },
                                    });
    var pbpPage = cleanupHTMLElement(pbpHTML);
    if (pbpPage.indexOf('Login') === -1) {
      var pbpDoc = XmlService.parse(pbpPage.match(/<main[\s\S]*<\/main>/)[0]).getRootElement();
      var pbpItems = getElementByClassName(getElementByClassName(pbpDoc, 'col-lg-12', true)[0], 'row', true);

      if (!pbpItems.length) {
        throw new Error('Nothing found for PbP -- something is probably wrong');
      }

      contextValues.pbpByUrl = {};
      pbpItems.forEach(processPBP);

      for (var currPbpItem in contextValues.pbpByUrl) {
        if (contextValues.pbpByUrl.hasOwnProperty(currPbpItem)) {
          addOrUpdatePbPItem(contextValues.pbpByUrl[currPbpItem]);
        }
      }
    } else {
      removeAndEmail(urls.pbpDomain);
    }
  } catch (e) {
    printError(e)
    removeAndEmail(urls.pbpDomain, `errorParsingPage: ${e}`);
  }
*/
  // --------------------------------------------
  // SF LISTINGS
  // Log in to SF
  /*
  try {
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
      var sfPage = cleanupHTMLElement(sfHTML);
      var sfDoc = XmlService.parse(sfPage).getRootElement();
      getElementByClassName(sfDoc, 'showcasebox', true).forEach(addOrUpdateSf);
    }
  } catch (e) {
    printError(e)
    removeAndEmail(urls.sf, `errorParsingPage: ${e}`);
  }
  */

  // --------------------------------------------
  // CT LISTINGS
  try {
    var mainPage = getMainPageCT();
    if (mainPage) {
      var items = mainPage.data;
      items.forEach(addOrUpdateCT);
    } else {
      removeAndEmail(urls.ctDomain);
    }
  } catch (e) {
    printError(e)
    removeAndEmail(urls.ctDomain, `errorParsingPage: ${e}`);
  }



  addNewCellItemsRow();
  updateAllCells();

  // Only email if user wants email. This handles vacations
  var secondLastIdx = contextValues.sheetData[0].length - 2;
  var timeEmail = contextValues.sheetData[0][secondLastIdx];
  if (timeEmail) {
    sendEmail(timeEmail);
  }

  archiveExpiredItems();
}

// =====================================
// PROCESSING HELPER FUNCTIONS
// =====================================

// --------------------------------------------
// Process previous data, including title and fee in case those change
function processPreviousListings() {
  var titleIdx = contextValues.sheetIndex.Title;
  var idIdx = contextValues.sheetIndex.Url;
  var feeIdx = contextValues.sheetIndex.AdminFee;
  var imageIdx = contextValues.sheetIndex.Image;
  var dateIdx = contextValues.sheetIndex.Date;
  var categoryIdx = contextValues.sheetIndex.Category;
  var locationRatingIdx = contextValues.sheetIndex.LocationRating;
  var locIdx = contextValues.sheetIndex.Location;
  var ratingIdx = contextValues.sheetIndex.Rating;
  contextValues.lastRow = contextValues.sheetData.length;
  contextValues.previousListings = {};

  // Also, get formula for image
  // Get range by row, column, row length, column length
  var imageFormulas = contextValues.sheetRange.getFormulas();
  var previousListingObject = {};
  var urlValue, titleValue, feeValue, dateValue, categoryValue, locValue, locRatingValue, currItem;
  for (var i = 1; i < contextValues.lastRow; i++) {
    currItem = contextValues.sheetData[i];
    urlValue = currItem[idIdx].trim();
    if (urlValue) {
      titleValue = currItem[titleIdx];
      feeValue = currItem[feeIdx];
      dateValue = currItem[dateIdx];
      categoryValue = currItem[categoryIdx];
      locValue = currItem[locIdx];
      locRatingValue = currItem[locationRatingIdx];

      // if (!locRatingValue) {
      //   locRatingValue = getLocationRating(locValue);
      //   currItem[locationRatingIdx] = locRatingValue;
      // }

      currItem[imageIdx] = imageFormulas[i][0];
      previousListingObject = {
        row: i,
        title: titleValue,
        fee: feeValue,
        date: dateValue,
        category: categoryValue,
        location: locValue,
        locationRating: locRatingValue,
        rating: currItem[ratingIdx],
        url: urlValue,
      };
      previousListingObject[titleIdx] = titleValue;
      previousListingObject[feeIdx] = feeValue;
      previousListingObject[idIdx] = urlValue;
      previousListingObject[dateIdx] = dateValue;
      contextValues.previousListings[urlValue] = previousListingObject;
    }
  }
}

// --------------------------------------------
// Rating Management Helpers
function arraysEqual(arr1, arr2) {
    if(arr1.length !== arr2.length)
        return false;
    for(var i = arr1.length; i--;) {
        if(arr1[i] !== arr2[i])
            return false;
    }

    return true;
}

function processOldRatings(ratingData, idx) {
  // // Remove duplicates
  // var oldData = contextValues.ratings[ratingData[contextValues.ratingIndex.URL]];
  // // if (oldData && arraysEqual(ratingData, contextValues.ratingData[oldData])) {
  // if (oldData) {
  //   for (var i = 0; i < ratingData.length;) ratingData[i++] = '';
  //   return;
  // }

  contextValues.ratings[ratingData[contextValues.ratingIndex.URL]] = idx;
  contextValues.ratings[cleanupTitle(ratingData[contextValues.ratingIndex.Title])] = ratingData;
}

function processOldLocationRatings(ratingData, idx) {
  var cleanedLocation = cleanupLocation(ratingData[contextValues.locationRatingIndex.SortedName]);
  contextValues.locationRatings[cleanedLocation] = ratingData;
}

function processRatingItem(itemText) {
  var fullTitleIdx = contextValues.ratingIndex.FullTitle;
  var titleIdx = contextValues.ratingIndex.Title;
  var locationIdx = contextValues.ratingIndex.Location;
  var ratingIdx = contextValues.ratingIndex.Rating;
  var numberReviewsIdx = contextValues.ratingIndex.NumberReviews;
  var urlIdx = contextValues.ratingIndex.URL;
  var addedIdx = contextValues.ratingIndex.AddedDate;
  var yearIdx = contextValues.ratingIndex.YearAdded;
  var countColIdx = contextValues.ratingIndex.CountCol;

  var data = [];
  var noteArray;
  var url = itemText.match(/<a href="(.*?)"/);
  if (url) {
    url = getACUrl(url[1]);
    var rowIdx = contextValues.ratings[url];

    if (rowIdx !== undefined) {
      const currData = contextValues.ratingData[rowIdx];
      if (currData[yearIdx] >= currYear) {
        data = currData;
      }
    }

    var rating = itemText.match(/<img .*?icon-star\s/g);
    rating = rating ? rating.length : 0;
    var numberReviews = itemText.match(/read\s+(\d*)\s+review/i);
    if (numberReviews) {
      numberReviews = +numberReviews[1];
    }

    var color = 'darkgrey';
    if (rating >= 5) {
      if (numberReviews >= 40) {
        color = 'forestgreen';
      } else if (numberReviews >= 20) {
        color = 'limegreen';
      } else if (numberReviews >= 10) {
        color = 'greenyellow';
      } else {
        color = 'black';
      }
    }

    var headerOpening = '<h4 style="color:' + color + ';">';

    var fullTitle = itemText.match(/text-tac-gold.*?>(.*?)</)[1];

    if (data.length) {
      noteArray = contextValues.ratingNotes[rowIdx];
      var updated = [];

      if (rating !== data[ratingIdx]) {
        updated.push('Rating: ' + boldWord(rating) + '<em>(Previously ' + boldWord(data[ratingIdx]) + ')</em>');
        noteArray[ratingIdx] = new Date().toISOString() + ' overwrote: ' + data[ratingIdx] + '\n' + noteArray[ratingIdx];
        data[ratingIdx] = rating;
      }

      if (numberReviews !== data[numberReviewsIdx]) {
        updated.push('NumberReviews: ' + boldWord(numberReviews) + '<em>(Previously ' + boldWord(data[numberReviewsIdx]) + ')</em>');
        noteArray[numberReviewsIdx] = new Date().toISOString() + ' overwrote: ' + data[numberReviewsIdx] + '\n' + noteArray[numberReviewsIdx];
        data[numberReviewsIdx] = numberReviews;
      }

      // TODO: Remove after first run
      // cleanup old location
      // var lastIndex = fullTitle.lastIndexOf(' at ');
      // var cleanedLocation = cleanupLocation(fullTitle.substr(lastIndex + 4).replace(/\s*\(the\)\s*$/i, ''));
      // if (cleanedLocation !== data[locationIdx]) {
      //   updated.push('Location: ' + boldWord(cleanedLocation) + '<em>(previously ' + boldWord(data[locationIdx]) + ')</em>');
      //   noteArray[locationIdx] = new Date().toISOString() + ' overwrote: ' + data[locationIdx] + '\n' + noteArray[locationIdx];
      //   data[locationIdx] = cleanedLocation;
      // }

      if (updated.length) {
        updated.push('url: ' + url);
        contextValues.updatedRatings.push(headerOpening + fullTitle + '  <small>' + rating + '</small></h4>' + updated.join('<br>'));
        contextValues.ratingMin = Math.min(rowIdx + 1, contextValues.ratingMin);
      }
    } else {
      var lastIndex = fullTitle.lastIndexOf(' at ');
      var title = fullTitle.substr(0, lastIndex).trim();
      var location = cleanupLocation(fullTitle.substr(lastIndex + 4).replace(/\s*\(the\)\s*$/i, ''));
      var newItems = [];

      data[fullTitleIdx] = fullTitle;
      data[titleIdx] = cleanupTitle(title);
      data[locationIdx] = location;
      data[ratingIdx] = rating;
      data[numberReviewsIdx] = numberReviews;
      data[urlIdx] = url;
      data[addedIdx] = new Date();
      data[yearIdx] = currYear;
      data[countColIdx] = 1;
      noteArray = [];
      for (var i = 0; i < data.length;) noteArray[i++] = '';

      var currIdx = contextValues.ratingNotes.length;
      contextValues.ratings[url] = currIdx;
      contextValues.ratings[data[titleIdx]] = currIdx;
      contextValues.ratingNotes.push(noteArray);
      contextValues.ratingData.push(data);

      newItems.push('NumberReviews: ' + numberReviews)
      newItems.push('url: ' + url)
      contextValues.newRatings.push(headerOpening + fullTitle + '  <small>' + rating + '</small></h4>' + newItems.join('<br>'));
    }
  }
}

// --------------------------------------------
// AC Helpers
// Figure out of the page which listings are new
function parseAcItems(item) {
  var header = getElementsByTagName(item, 'h5', true)[0];
  var aElement = getElementsByTagName(header, 'a', true)[0];
  var title = aElement.getText().trim();

  if (!title) {
    return;
  }
  
  var bookElement = getElementByClassName(item, 'ladder-book-now', true)[0];
  var dateElement = getElementByClassName(bookElement, 'text-center', true)[0];
  var date = dateElement
              .getValue()
              .replace(/\s*shows from\s*/i, '')
              .replace(/\s+/g, ' ')
              .trim();

  var bookText = bookElement.getValue();
  if (isSold(bookText)) {
    date += " SOLD OUT";
  }

  var url = getACUrl(aElement.getAttribute('href').getValue());
  var listingInfo = contextValues.currAcItems[url];
  var isFree = item.getAttribute('class').getValue().includes('rush-ticket');
  if (listingInfo) {
    listingInfo.isFree = listingInfo.isFree || isFree;
    // Updated the date only with new date
    var previousDate = listingInfo[contextValues.sheetIndex.Date];
    listingInfo[contextValues.sheetIndex.Date] = previousDate + '; ' + date;
  } else if (!contextValues.alreadyDeleted[url]) {
    var ImageElements = getElementByClassName(item, 'ladder-rung-title', true)[0];
    var ImageUrl = ImageElements ? ImageElements.getAttribute('style').getValue().match(/url\(\/?(.*)\)/) : null;
    ImageUrl = (ImageUrl && `${urls.acDomain}${ImageUrl[1]}`) || ''

    var contentElement = getElementByClassName(item, 'ladder-rung-content', true)[0];

    var description = contentElement.getValue().replace(/\bread\s+more\b/i, '').replace(/\s+/g, ' ').trim();

    var venue = getElementByClassName(item, 'ladder-rung-venue', true)[0].getValue().trim();
    var rating = getRating(title, url);
    listingInfo = [];
    listingInfo[contextValues.sheetIndex.Image] = ImageUrl ? '=Image("' + ImageUrl + '")' : '';
    listingInfo[contextValues.sheetIndex.Title] = title;
    listingInfo[contextValues.sheetIndex.Rating] = rating;
    listingInfo[contextValues.sheetIndex.LocationRating] = getLocationRating(venue);
    listingInfo[contextValues.sheetIndex.AdminFee] = isFree ? 'FREE' : '~£4.50';
    listingInfo[contextValues.sheetIndex.Date] = date;
    listingInfo[contextValues.sheetIndex.Category] = description;
    listingInfo[contextValues.sheetIndex.Location] = venue;
    listingInfo[contextValues.sheetIndex.Url] = url;
    listingInfo[contextValues.sheetIndex.EventManager] = '';
    listingInfo[contextValues.sheetIndex.UploadDate] = new Date();
    listingInfo.isFree = isFree;
    contextValues.currAcItems[url] = listingInfo;
  }
}

function isSold(singleDate) {
  return singleDate.match(/\bsold\s+out\b/i);
}

function addOrUpdateAc(acUrl) {
  var currItem = contextValues.currAcItems[acUrl];
  var itemInfo = contextValues.previousListings[acUrl];

  // Update title depending on whether it's sold out
  var title = currItem[contextValues.sheetIndex.Title];
  var date = currItem[contextValues.sheetIndex.Date];
  var isSoldOut = date.split(/\s*;\s+/g).every(isSold);
  if (isSoldOut) {
    title += ' (SOLD OUT)';
  }

  if (itemInfo) {
    var currentItem = [];
    var isNowFree = currItem.isFree && itemInfo[contextValues.sheetIndex.AdminFee] !== 'FREE';
    var isNowPaid = !currItem.isFree && itemInfo[contextValues.sheetIndex.AdminFee] === 'FREE';
    if (isNowFree || isNowPaid) {
      var newFee = isNowFree ? 'FREE' : '~£4.50';
      var oldFee = isNowFree ? '~£4.50' : 'FREE';
      markCellForUpdate(itemInfo.row, 'AdminFee', newFee);
      currentItem[contextValues.sheetIndex.AdminFee] = boldWord(newFee) + ' <br><em>(Previously ' + boldWord(oldFee) + ')</em>';
    }

    if (date !== itemInfo.date) {
      markCellForUpdate(itemInfo.row, 'Date', date);
      currentItem[contextValues.sheetIndex.Date] = boldWord(date) + '<br><em>(Previously ' + boldWord(itemInfo.date) + ')</em>';
    }

    if (title !== itemInfo.title) {
      markCellForUpdate(itemInfo.row, 'Title', title);
      currentItem[contextValues.sheetIndex.title] = boldWord(title) + '<br><em>(Previously ' + boldWord(itemInfo.title) + ')</em>';
    }

    // Fill out incorrectly empty locations
    var location = currItem[contextValues.sheetIndex.Location];
    if (location && location !== itemInfo.location) {
      markCellForUpdate(itemInfo.row, 'Location', location);
      currentItem[contextValues.sheetIndex.Date] = boldWord(location) + '<br><em>(Previously ' + boldWord(itemInfo.location) + ')</em>';
    }

    var description = currItem[contextValues.sheetIndex.Category];
    if (description !== itemInfo.category &&
        description !== "'" + itemInfo.category &&
        description !== "'" + itemInfo.category + "'" &&
        description !== itemInfo.category + "'" &&
        description !== '"' + itemInfo.category &&
        description !== '"' + itemInfo.category + '"' &&
        description !== itemInfo.category + '"') {
      markCellForUpdate(itemInfo.row, 'Category', description);
      currentItem[contextValues.sheetIndex.Category] = boldWord(description) + '<br><em>(Previously ' + boldWord(itemInfo.category) + ')</em>';
    }

    checkRatingAndDeletePreviousListing(itemInfo, acUrl, currentItem, title);
  } else if (!contextValues.alreadyDeleted[acUrl]) {
    newItemsForUpdate.push(currItem);
  }
}

function getACUrl(urlEnd) {
  return urls.acDomain + 'member/' + urlEnd.replace(/&return=.*$/, '').trim();
}

// --------------------------------------------
// PbP Helpers
// Process PBP rows -- pretty complicated because everything is nested tables
function processPBP(rowEl) {
  var header = getElementsByTagName(getElementByClassName(rowEl, 'card-header', true)[0], 'a', true)[0];
  var pbpId = header.getAttribute('href').getValue();
  if (!pbpId) {
    return;
  }

  var url = getFullPbpUrl(pbpId);

  // If this already exists, get location, time, and picture
  if (contextValues.pbpByUrl[url]) {
    getLocationTimePicture(rowEl, contextValues.pbpByUrl[url]);
  } else {
    contextValues.pbpByUrl[url] = [];
    getLocationTimePicture(rowEl, contextValues.pbpByUrl[url]);
    contextValues.pbpByUrl[url][contextValues.sheetIndex.Url] = url;
    contextValues.pbpByUrl[url][contextValues.sheetIndex.Rating] = getRating(contextValues.pbpByUrl[url][contextValues.sheetIndex.Title]);
  }
}

function addOrUpdatePbPItem(pbpItem) {
  var url = pbpItem[contextValues.sheetIndex.Url];
  var itemInfo = contextValues.previousListings[url];
  if (itemInfo) {
    var currentItem = [];
    var date = pbpItem[contextValues.sheetIndex.Date];
    var title = pbpItem[contextValues.sheetIndex.Title];

    if (title !== itemInfo.title) {
      markCellForUpdate(itemInfo.row, 'Title', title);
      currentItem[contextValues.sheetIndex.title] = boldWord(title) + '<br><em>(Previously ' + boldWord(itemInfo.title) + ')</em>';
    }

    if (date !== itemInfo.date) {
      markCellForUpdate(itemInfo.row, 'Date', date);
      currentItem[contextValues.sheetIndex.Date] = boldWord(date) + '<br><em>(Previously ' + boldWord(itemInfo.date) + ')</em>';
    }

    checkRatingAndDeletePreviousListing(itemInfo, url, currentItem, title);
  } else if (!contextValues.alreadyDeleted[url]) {
    pbpItem[contextValues.sheetIndex.AdminFee] = '~£2/4';
    pbpItem[contextValues.sheetIndex.Category] = '';
    pbpItem[contextValues.sheetIndex.EventManager] = '';
    pbpItem[contextValues.sheetIndex.UploadDate] = new Date();
    newItemsForUpdate.push(pbpItem);
  }
}

function getLocationTimePicture(rowEl, elData) {
  var image = getElementsByTagName(rowEl, 'img', true)[0];
  var imageUrl = image.getAttribute('src').getValue();
  var title = image.getAttribute('alt').getValue();
  var locationEl = getElementByClassName(rowEl, 'col-md-3', true)[0];
  var locationName = locationEl.getChildren('h4')[0].getValue();
  var location = `${locationName}; ${locationEl.getText().replace(/([a-z])([A-Z])/g, '$1; $2')}`
  var dateAll = getElementByClassName(rowEl, 'col-md-7', true)[0];
  var isFull = true;
  var dateArray = getElementByClassName(dateAll, 'row', true).map(function(row) {
    const fullText = getFullText(row, ' - ');
    const isCurrFull = !fullText.match(/reserve/i);
    isFull = isFull && !!isCurrFull;
    return fullText;
  });
  elData[contextValues.sheetIndex.Image] = imageUrl ? '=Image("' + getFullPbpUrl(imageUrl) + '")' : '';
  elData[contextValues.sheetIndex.Date] = dateArray.join('; ');
  elData[contextValues.sheetIndex.Title] = isFull ? `${title} (FULLY BOOKED)` : title;
  elData[contextValues.sheetIndex.Location] = location;
  elData[contextValues.sheetIndex.LocationRating] = getLocationRating(location);
}

function getFullPbpUrl(locationString) {
  return urls.pbpDomain + (locationString.indexOf('/') === 0 ? '' : '/') + locationString;
}

// --------------------------------------------
// SF Helpers
// Figure out if the page which listings are new
function addOrUpdateSf(item) {
  var title = getElementsByTagName(item, 'h2', true);
  title = title[0].getText().trim();
  if (!title) {
    return;
  }

  var aElement = getElementsByTagName(item, 'a', true)[0];
  var url = aElement.getAttribute('href').getValue().trim();
  var itemInfo = contextValues.previousListings[url];
  if (itemInfo) {
    checkRatingAndDeletePreviousListing(itemInfo, url, [], title);
  } else if (!contextValues.alreadyDeleted[url]) {
    var itemHtml = item.getText();
    var ImageUrl = itemHtml.match(/background-image:url\(.*?(http:\/\/.*?\.jpg)/i);

    var date = getElementByClassName(item, 'date-event', true);
    var description = getElementByClassName(item, 'internal_content', true);

    var detailPage = cleanupHTMLElement(UrlFetchApp.fetch(url));
    var detailError = 'Sorry, this offer has now ended';
    var price = '', location = '', time = '';
    if (detailPage.indexOf(detailError) === -1) {
      price = detailPage.match(/price_info_box.*?>([\s\S]*?)<\/span>/);
      location = detailPage.match(/location_td.*?>([\s\S]*?)<\/td>/);
      price = price ? trimHtml(price[1]).trim() : '';
      location = location ? trimHtml(location[1]).trim() : '';
      time = ' @ ' + detailPage.match(/<td>(.*?:.*?)<\/td>/)[1];
    }

    var listingInfo = [];
    listingInfo[contextValues.sheetIndex.Image] = '=Image("' + ((ImageUrl && ImageUrl[1]) || '') + '")';
    listingInfo[contextValues.sheetIndex.Title] = title;
    listingInfo[contextValues.sheetIndex.Rating] = getRating(title);
    listingInfo[contextValues.sheetIndex.LocationRating] = getLocationRating(location);
    listingInfo[contextValues.sheetIndex.AdminFee] = price;
    listingInfo[contextValues.sheetIndex.Date] = date[0].getText().trim() + time;
    listingInfo[contextValues.sheetIndex.Category] = trimHtml(description[0].getText()).trim();
    listingInfo[contextValues.sheetIndex.Location] = location;
    listingInfo[contextValues.sheetIndex.Url] = url;
    listingInfo[contextValues.sheetIndex.EventManager] = '';
    listingInfo[contextValues.sheetIndex.UploadDate] = new Date();
    newItemsForUpdate.push(listingInfo);
  }
}

// --------------------------------------------
// CT Helpers
// Get the main CT page
function getMainPageCT() {
  /*
  var ctToken = UrlFetchApp.fetch(urls.login);
  var headers1 = ctToken.getAllHeaders();
  fetchPayload.ctPasswordCookie = headers1['Set-Cookie'].join('');
  var keyInfo = ctToken.getContentText().match(/type="hidden" value="(.*?)"/)[1];
  */

  var loginCtPage = UrlFetchApp.fetch(urls.ctLoginApi, {
      method: 'POST',
      followRedirects: false,
      payload: {
        email: fetchPayload.email,
        password: sjcl.decrypt(fetchPayload.salt, fetchPayload.password),
      },
    });

  // var ctHeaders = loginCtPage.getAllHeaders();
  // fetchPayload.ctPasswordCookie = ctHeaders['Set-Cookie'];

  var loginContent = JSON.parse(loginCtPage.getContentText());

  var mainPage = UrlFetchApp.fetch(urls.main,
                                  {
                                    method: 'post',
                                    headers: {
                                      authorization: `Bearer ${loginContent.data.token}`,
                                    },
                                  });
  
  var mainPageText = JSON.parse(mainPage.getContentText());

  if (!mainPageText || mainPageText.message.includes('login')) {
    return null;
  }

  return mainPageText;
}

// Updates or calls add new listing
function addOrUpdateCT(item) {
  // Get href
  var url = urls.ctEvent + item.id;
  var oldUrl = urls.ctEventOld + item.id;
  var itemInfo = contextValues.previousListings[url] || contextValues.previousListings[oldUrl];
  var ticketsLeft = item.leftTickets;

  // Better show actual time/date
  var date = item.event_date_time + ' (' + item.filter_arrays['3'].join('; ') + ')';

  var fee = '£' + item.admin_fees;
  var title = item.event_name;

  if (!ticketsLeft) {
    title += ' (FULLY BOOKED)'
  } else if (ticketsLeft < 20) {
    title += ' (' + ticketsLeft + ' Tickets Left)'
  }

  if (itemInfo) {
    // see if there's anything to update, if not, then just delete
    var currentItem = [];
    if (fee !== itemInfo.fee) {
      markCellForUpdate(itemInfo.row, 'AdminFee', fee);
      currentItem[contextValues.sheetIndex.AdminFee] = boldWord(fee) + '<br><em>(Previously ' + boldWord(itemInfo.fee) + ')</em>';
    }

    if (date !== itemInfo.date) {
      markCellForUpdate(itemInfo.row, 'Date', date);
      // Only flag date change in email if date changed or booked status changed
      var oldItemHasTickets = itemInfo.title.indexOf('(FULLY BOOKED)') === -1;
      if (!!ticketsLeft !== oldItemHasTickets) {
        currentItem[contextValues.sheetIndex.Date] = boldWord(date) + '<br><em>(Previously ' + boldWord(itemInfo.date) + ')</em>';
      }
    }

    if (url !== itemInfo.url) {
      markCellForUpdate(itemInfo.row, 'Url', url);
    }

    checkRatingAndDeletePreviousListing(itemInfo, url, currentItem, title);
  } else if (!contextValues.alreadyDeleted[url]) {
    addNewListingCT(item, url, title, date);
  }
}

// Process CT
function addNewListingCT(item, url, title, date) {
  var ImageUrl = item.event_image;
  var listingInfo = [];
  var clean_title = item.event_name;

  if (!clean_title) {
    return;
  }

  var location = item.address;
  listingInfo[contextValues.sheetIndex.Image] = '=Image("' + ImageUrl + '")';
  listingInfo[contextValues.sheetIndex.Title] = title;
  listingInfo[contextValues.sheetIndex.Rating] = getRating(clean_title);
  listingInfo[contextValues.sheetIndex.LocationRating] = getLocationRating(location);
  listingInfo[contextValues.sheetIndex.AdminFee] = '£' + item.admin_fees;
  listingInfo[contextValues.sheetIndex.Date] = date;
  listingInfo[contextValues.sheetIndex.Category] = item.categories + ' - ' + item.tag + ' - ' + item.short_description;
  listingInfo[contextValues.sheetIndex.Location] = location;
  listingInfo[contextValues.sheetIndex.Url] = url;
  listingInfo[contextValues.sheetIndex.EventManager] = 'Type: ' + item.tickets_type + ': Private: ' + item.privateevent + ': Top Event: ' + item.is_top_member_event;
  listingInfo[contextValues.sheetIndex.UploadDate] = new Date();
  newItemsForUpdate.push(listingInfo);
}

// --------------------------------------------
// Function that updates sheet
function addNewCellItemsRow() {
  if (!newItemsForUpdate.length) return;

  // Get range by row, column, row length, column length
  var cells = contextValues.sheet.getRange((contextValues.lastRow + 1), 1, newItemsForUpdate.length, newItemsForUpdate[0].length);
  cells.setValues(newItemsForUpdate);
}

// --------------------------------------------
// returns back regexp string and weight for that string
function getRegexpAndWeight(weightRegexpNote) {
  var weightAndRegex = weightRegexpNote.match(/(\d+):(.*)/);
  var weight = parseInt(weightAndRegex[1]);
  var regexp = weightAndRegex[2].replace(/\s+/g, '\\s+');

  return {
    regexpStr: regexp,
    weight: weight,
  };
}

// --------------------------------------------
// Weight sort function
function sortByWeight(a, b) {
  return b.weight - a.weight;
}

// --------------------------------------------
// Sorts items in array based on location or number of matching category keywords
function sortEmailItems(items) {
  var locationIdx = contextValues.sheetIndex.Location;
  var categoryIdx = contextValues.sheetIndex.Category;
  var feeIdx = contextValues.sheetIndex.AdminFee;
  var locationSearch = getRegexpAndWeight(contextValues.sheetNotes[0][locationIdx]);
  var locationWords = locationSearch.regexpStr.split('|');
  var locationRegexp = new RegExp('\\b' + locationWords.join('\\b|\\b') + '\\b', 'i');
  var categorySearch = getRegexpAndWeight(contextValues.sheetNotes[0][categoryIdx]);
  var categoryRegexp = new RegExp(categorySearch.regexpStr, 'gi');
  var pattern = '<em style="color: darkred;">$&</em>';
  var numberFound = 0;

  var currloc, currcat, locationFound, catFound, currFee, currFeeMatch, currFeeWeight;
  items.map(function addRating(itm) {
    currloc = itm[locationIdx];
    currcat = itm[categoryIdx];
    currFee = itm[feeIdx];
    locationFound = currloc.match(locationRegexp) ? locationSearch.weight : 0;
    catFound = currcat.match(categoryRegexp)
    catFound = catFound ? catFound.length : 0;

    // Put events that cost a lot at the bottom
    currFeeMatch = currFee.match(/\d+/);
    currFee = currFeeMatch ? +currFeeMatch[0] : 0;
    currFeeWeight = currFee ? 0 : 1; // Add weight for free events
    weight = currFee < 10 ? locationFound + catFound * categorySearch.weight + currFeeWeight: -1;

    // Update email values
    itm[locationIdx] = currloc.replace(locationRegexp, pattern);
    itm[categoryIdx] = currcat.replace(categoryRegexp, pattern);
    if (currFeeWeight) {
      itm[feeIdx] = pattern.replace('$&', itm[feeIdx]);
    }

    // add weight
    itm.weight = weight;

    if (weight) {
      numberFound += 1;
    }
  });

  return {
    sortedItems: items.sort(sortByWeight),
    numberFound: numberFound,
  };
}

// --------------------------------------------
// Send email with new listing information
function sendEmail(numberItemsToSend) {
  // Only send if there's new items
  var noUpdatedItem = !updatedItems.length && numberItemsToSend === 2;
  if (!newItemsForUpdate.length &&
      (noUpdatedItem || numberItemsToSend === 1)) {
    return;
  }

  var footer = '<hr>' +
               'Sheet: ' + spreadsheetURL;
  var sortNewItems = sortEmailItems(newItemsForUpdate);
  var sortUpdatedItems = sortEmailItems(updatedItems);
  var newItemsText = newItemsForUpdate.length ? '<hr><h2>New:</h2><br>' +
                     sortNewItems.sortedItems.map(getElementSection).join('') : '';
  var updatedItemsText = updatedItems.length ? '<hr><h2>Updated:</h2><br>' +
                         sortUpdatedItems.sortedItems.map(getElementSection).join('') : '';
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
  var newFound = sortNewItems.numberFound ? sortNewItems.numberFound + '/' : '';
  var updatedFound = sortUpdatedItems.numberFound ? sortUpdatedItems.numberFound + '/' : '';
  var subject = '[CT] *' + newFound + newItemsForUpdate.length + '* New || *' +
                updatedFound + updatedItems.length + '* Updated ' + new Date().toLocaleString();

  var email = MailApp.sendEmail({
    to: myEmail,
    subject: subject,
    htmlBody: emailTemplate,
  });
}

// --------------------------------------------
// Updates spreadsheet
function updateAllCells() {
  if (updatedItems.length) {
    contextValues.sheetRange.setValues(contextValues.sheetData);
    contextValues.sheetRange.setNotes(contextValues.sheetNotes);
  }
}

// --------------------------------------------
// Archives deleted items
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

      // UPDATE IF ADD NEW ROWS
      cutRange = contextValues.sheet.getRange('A' + row + ':K' + row);
      newRange = archive.getRange('A' + lastArchiveRow + ':L' + lastArchiveRow)
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

// =====================================
// GENERAL HELPER FUNCTIONS
// =====================================

function getRating(title, acShowUrl) {
  var currItem;
  if (acShowUrl) {
    let id = acShowUrl.match(/recid=(\d+)/);
    if (id) {
      const reviewUrl = urls.acReviews.replace(/\?year=.*/, `?sid=${id[1]}`).replace('viewReview', 'viewShowReview');
      const ratingIdx = contextValues.ratings[reviewUrl];
      currItem = contextValues.ratingData[ratingIdx];
    }
  }

  if (!currItem) {
    title = cleanupTitle(title);
    currItem = contextValues.ratings[title];
  }

  return currItem ?
         currItem[contextValues.ratingIndex.Rating] + '/5 (' +
            currItem[contextValues.ratingIndex.NumberReviews] + ' reviews - ' +
            currItem[contextValues.ratingIndex.URL] + ')' :
         '';
}

function emailRatingIfRatingChanged(newRating, oldRating, emailInfo) {
  var replaceRegex = /\/.*/;
  if (newRating.replace(replaceRegex, '') !== oldRating.replace(replaceRegex, '')) {
    var noUrlOld = oldRating.replace(/-.*\)/, ')');
    emailInfo[contextValues.sheetIndex.Rating] = boldWord(newRating) + '<em>(Previously ' + boldWord(noUrlOld) + ')</em>';
  }
}

function checkRatingAndDeletePreviousListing(itemInfo, url, currentItem, title) {
  var rating = getRating(title);
  var location = itemInfo.location;
  var locationRating = getLocationRating(location);

  if (title !== itemInfo.title &&
      title !== "'" + itemInfo.title &&
      title !== "'" + itemInfo.title + "'" &&
      title !== itemInfo.title + "'" &&
      title !== '"' + itemInfo.title &&
      title !== '"' + itemInfo.title + '"' &&
      title !== itemInfo.title + '"') {
    markCellForUpdate(itemInfo.row, 'Title', title);
    currentItem[contextValues.sheetIndex.Title] = boldWord(title) + '<br><em>(Previously ' + boldWord(itemInfo.title) + ')</em>';
  }

  if (rating !== itemInfo.rating) {
    markCellForUpdate(itemInfo.row, 'Rating', rating);
    emailRatingIfRatingChanged(rating, itemInfo.rating, currentItem);
  }

  if (locationRating !== itemInfo.locationRating) {
    markCellForUpdate(itemInfo.row, 'LocationRating', locationRating);
  }

  if (currentItem.length) {
    if (!currentItem[contextValues.sheetIndex.Title]) {
      currentItem[contextValues.sheetIndex.Title] = title;
    }

    if (!currentItem[contextValues.sheetIndex.AdminFee]) {
      currentItem[contextValues.sheetIndex.AdminFee] = itemInfo.fee;
    }

    if (!currentItem[contextValues.sheetIndex.Date]) {
      currentItem[contextValues.sheetIndex.Date] = itemInfo.date;
    }

    if (!currentItem[contextValues.sheetIndex.Rating]) {
      currentItem[contextValues.sheetIndex.Rating] = rating;
    }

    if (!currentItem[contextValues.sheetIndex.Category]) {
      currentItem[contextValues.sheetIndex.Category] = itemInfo.category;
    }


    currentItem[contextValues.sheetIndex.Location] = itemInfo.location;
    currentItem[contextValues.sheetIndex.LocationRating] = locationRating;
    currentItem[contextValues.sheetIndex.Url] = url;
    updatedItems.push(currentItem);
  }

  delete contextValues.previousListings[url];
  contextValues.alreadyDeleted[url] = true;
}

function removeAndEmail(domain, specificErrorMessage) {
  specificErrorMessage = specificErrorMessage || '';

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
  var lastRow = contextValues.errorData[contextValues.lastErrorRow - 1];
  var lastEmailedDate = lastRow[contextValues.errorDateIdx];
  var sameErrorMessage = lastRow[contextValues.errorSitesIdx].includes(domain) && lastRow[contextValues.errorIndex.errorMessage] === specificErrorMessage;
  if ((specificErrorMessage && !sameErrorMessage) || !lastEmailedDate.toDateString || lastEmailedDate.toDateString() !== new Date().toDateString()) {
    var updateMessage = 'Update ' + domain + ' Token';

    if (specificErrorMessage) {
      updateMessage = 'Error loading page for : ' + domain + ' (' + specificErrorMessage + ')';
    }

    var email = MailApp.sendEmail({
      to: myEmail,
      subject: '[CT] ' + updateMessage.substring(0, 150),
      htmlBody: updateMessage
                 + ': ' + spreadsheetURL,
    });

    contextValues.lastErrorRow++;
    contextValues.errorData.push([new Date(), domain, specificErrorMessage || 'updateToken'])
  }

  // Add current page to list of pages needing update
  if (!contextValues.errorData[contextValues.lastErrorRow - 1]) {
    contextValues.errorData[contextValues.lastErrorRow - 1] = [];
  }

  var currentData = contextValues.errorData[contextValues.lastErrorRow - 1][contextValues.errorSitesIdx];
  if (specificErrorMessage || !currentData || currentData.indexOf(domain) === -1) {
    var cells = contextValues.errorSheet.getRange(contextValues.lastErrorRow, 1, 1, 3);
    currentData = currentData && !currentData.includes(domain) ? currentData + ', ' + domain : domain;
    contextValues.errorData[contextValues.lastErrorRow - 1][contextValues.errorSitesIdx] = currentData;
    cells.setValues([[new Date(), currentData, specificErrorMessage || 'updateToken']]);
  }
}

// =====================================
// HTML HELPER FUNCTIONS
// =====================================
function getFullText(element, separator='; ') {
  const result = element.getText().trim();
  const remainingText = element.getChildren().map((child) => child.getValue().trim()).filter(val => !!val);
  return `${result}${result && remainingText.length ? separator : ''}${remainingText.join(separator)}`
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
function getElementsByTagName(baseElement, tagName, onlyFirstLevel) {
  var data = baseElement.getChildren(tagName);

  var elList = baseElement.getChildren();
  var i = elList.length;
  while (i-- && (!onlyFirstLevel || !data.length)) {
    // (Recursive) Check each child, in document order.
    var found = getElementsByTagName(elList[i], tagName, onlyFirstLevel);
    if (found.length) {
      data = data.concat(found);

      if (onlyFirstLevel) {
        return data;
      }
    }
  }

  return data;
}

function getElementByClassName(baseElement, className, stopAfterFind) {
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

  var elList = baseElement.getChildren();
  var data = elList.filter(containsClass);

  if (stopAfterFind && data.length) {
    return data;
  }

  var i = elList.length;
  while (i--) {
    // (Recursive) Check each child, in document order.
    var found = getElementByClassName(elList[i], className, stopAfterFind);
    if (found.length) {
      data = data.concat(found);

      if (stopAfterFind) {
        return data;
      }
    }
  }

  return data;
}

function getElementSection(listingInfo) {
  var imageIdx = contextValues.sheetIndex.Image;
  var titleIdx = contextValues.sheetIndex.Title;
  var ratingIdx = contextValues.sheetIndex.Rating;
  var locationRatingIdx = contextValues.sheetIndex.LocationRating;
  var locationIdx = contextValues.sheetIndex.Location;
  var dateIdx = contextValues.sheetIndex.Date;
  var categoryIdx = contextValues.sheetIndex.Category;
  var urlIdx = contextValues.sheetIndex.Url;
  var feeIdx = contextValues.sheetIndex.AdminFee;
  var imageUrl = listingInfo[imageIdx] ? getImageUrl(listingInfo[imageIdx]) : '';
  var imageDiv = imageUrl ? '<img src="' + imageUrl + '" alt="' + listingInfo[titleIdx] + '" width="128">' :
                 '';
  var rating = listingInfo[ratingIdx] ? ' <small>' + listingInfo[ratingIdx] + '</small>' : '';
  var feeDiv = listingInfo[feeIdx] ? (listingInfo[feeIdx] + '<br>') : '';
  var locationRatingDiv = listingInfo[locationRatingIdx] ? ' (<small>' + (listingInfo[locationRatingIdx] + '</small>)') : '';
  var locationDiv = listingInfo[locationIdx] ? (listingInfo[locationIdx] + locationRatingDiv + '<br>') : '';
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

function cleanupHTMLElement(html) {
  return cleanupHTML(html.getContentText());
}

function cleanupHTML(htmlText) {
  return htmlText.match(/<body[\s\S]*?<\/body>/)[0]
                 .replace(/<(no)?script[\s\S]*?<\/(no)?script>|<link[\s\S]*?<\/link>|<footer[\s\S]*?<\/footer>|<button[\s\S]*?<\/button>|&copy;/g, '')
                 .replace(/&nbsp;|<\/?span[\s\S]*?>|<table[\s\S]+?<\/table>/g, ' ') // ugh sf
                 .replace(/<select[\s\S]*?<\/select>/g, ' ') // Audience club -- some select messed up and gives "Attribute name "multiple" associated with an element type "select" must be followed by the ' = ' character." error
                 .replace(/<a.*?genreModal.*?<\/a><\/p>/g, ' ')//  Audience club -- The Element Type "div" Must Be Terminated By The Matching End-tag -- <a class="btn btn-tac-gold genreModal" data-toggle="modal" data-id="recid=378655" href="#genreModal" />Select Genres</a></p>		
                 .replace(/<img([\s\S]*?)\/?>/ig, '<img$1 />') // ugh sf
                 .replace(/<\/?input.*>/ig, '') // AC <input needs </input> 
                 .replace(/ & /g, ' and ') // ugh sf
                 .replace(/&/g, '&amp;') // ugh sf
                 .replace(/�/g, "'") // ugh sf
                 .replace(/<br>/ig, "<br/>") // ugh sf
                 .replace(/<!--\s*?(?!<)[\s\S]*?-->/g, '')
                 .replace(/<!--|-->/g, '');
}

// Add item information to specific cell, archiving previous value as note
function markCellForUpdate(row, key, value) {
  var cellColumn = contextValues.sheetIndex[key];
  var previousMessage = contextValues.sheetData[row][cellColumn];
  if (previousMessage) {
    var oldNote = contextValues.sheetNotes[row][cellColumn];
    var newNote = new Date().toISOString() + ' overwrote: ' + previousMessage + '\n';
    var currentNote = oldNote ? (oldNote + newNote) : newNote;
    contextValues.sheetNotes[row][cellColumn] = currentNote;
  }

  contextValues.sheetData[row][cellColumn] = value;
}

function cleanupTitle(title) {
  if (title.replace) {
    return title.replace(/\s*\(.*\)\s*$/i, '').replace(/\s\s*/, ' ').toLowerCase();
  } else {
    return title;
  }
}

function getLocationRating(location) {
  location = cleanupLocation(location, 'thoroughReplace');
  var currLocation = contextValues.locationRatings[location];
  var ratingIdx = contextValues.locationRatingIndex.SortedRating;
  var countIdx = contextValues.locationRatingIndex.SortedCount;
  var reviewIdx = contextValues.locationRatingIndex.SortedReviews;
  var currRating = currLocation ? currLocation[ratingIdx] : '';
  return (currRating !== '' && !isNaN(currRating)) ?
    Math.round(currRating * 100) / 100 +
            ' (' + currLocation[countIdx] +
            ' shows - ' + currLocation[reviewIdx] + ' reviews - ' + location + ')' :
         '';
}

function cleanupLocation(location, thoroughReplace) {
  if (location.replace) {
    location = location.trim().replace(/\([\s\S]*|,[\s\S]*|\s-\s[\s\S]*|;[\s\S]*|^the\s*|/ig, '');

    if (thoroughReplace) {
      location = location.trim().replace(/\s\s[\s\S]*/, '');
    }

    return location.trim()
                   .toLowerCase();
  } else {
    return location;
  }
}

function boldWord(word) {
  return '<b>' + word + '</b>';
}

function printError(error) {
  console.log("Error", error.stack);
  console.log("Error", error.name);
  console.log("Error", error.message);
}
