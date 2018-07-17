// sheetsToUpdate is the name of the tab and freecycle group
// sheetsToEmail are sheet names that will be emailed if new items added
var sheetsToUpdate = [
];
var sheetsToEmail = {
};
var searchTerms = new RegExp(' cat | pet | keyboard | yoga | plant ');
var maxPrevious = 60;
var contextValues = {};
var fetchPayload = {
  headers: {
  },
};

/** MAIN FUNCTION
    Runs upon change event, ideally new row being added to ItemizedBudget Sheet
*/
function updateAllSheets() {
  sheetsToUpdate.forEach(updateSheet);
}

// Main function for each sheet
function updateSheet(sheetName) {
  Utilities.sleep(1000);
  contextValues.sheetName = sheetName;
  contextValues.sheet = SpreadsheetApp.getActiveSpreadsheet()
                            .getSheetByName(sheetName);
  contextValues.sheetData = contextValues.sheet.getDataRange().getValues();
  contextValues.sheetIndex = indexSheet(contextValues.sheetData);
  processPreviousListings();

  var freecycleUrl = 'https://groups.freecycle.org/group/' + sheetName + '/posts/all?resultsperpage=50&include_offers=on&include_wanteds=off&include_receiveds=off&include_takens=off';
  var freecycleHTML = UrlFetchApp.fetch(freecycleUrl, fetchPayload)
  freecycleHTML = freecycleHTML.getContentText().match(/<table[\s\S]*<\/table>/)[0];
  var items = freecycleHTML.match(/<tr[\s\S]*?<\/tr>/g);
  contextValues.missingItems = [];
  items.forEach(trackIfMissing);
  contextValues.missingItems.forEach(getListing);
}

// Figure out last 30 listings so there are no repeats
function processPreviousListings() {
  contextValues.lastRow = numberOfRows(contextValues.sheetData);
  var firstRow = Math.max(contextValues.lastRow - maxPrevious, 1);
  var idIdx = contextValues.sheetIndex.ID;
  contextValues.previousListings = {};
  for (var i = firstRow; i < contextValues.lastRow; i++) {
    contextValues.previousListings[parseInt(contextValues.sheetData[i][idIdx])] = i;
  }
}

// Figure out of the page which listings are new
function trackIfMissing(item) {
  if (item.match('icon_wanted')) return;

  // Get first column's text
  var currId = parseInt(item.match(/\(#(\d+)\)/)[1]);
  var row = contextValues.previousListings[currId];
  if (row) {
    if (row > contextValues.previousRow + 1) {
      for (var i = contextValues.previousRow + 1; i < row; i++) {
        addListingCell(i + 1, 'TakenDate', new Date());
      }
    }

    contextValues.previousRow = row;
  } else {
    contextValues.missingItems.push(currId);
  }
}

// Get listing full page
function getListing(listingId) {
  // Max 1 query per second
  Utilities.sleep(1000);
  contextValues.lastRow++;
  var freecycleItemUrl = 'https://groups.freecycle.org/group/' + contextValues.sheetName + '/posts/' + listingId;
  var freecycleItemHTML = UrlFetchApp.fetch(freecycleItemUrl, fetchPayload).getContentText().match(/<section[\s\S]*<\/section>/)[0];
  var Title = trimHeader(trimHtml(getElementByTag(freecycleItemHTML, 'h2')[2]));
  var ImageUrl = freecycleItemHTML.match('<img.+?src=".+?"');
  ImageUrl = ImageUrl ? ImageUrl[0].replace(/.*src=/,'') : '';
  ImageUrl = ImageUrl.slice(1, ImageUrl.length - 1);
  var Description = trimHtml(getElementByTag(freecycleItemHTML, 'p')[0]);
  var Location = getColonSeparatedText(freecycleItemHTML, 'Location');
  var currDate = getColonSeparatedText(freecycleItemHTML, 'Date');
  // Posted by not working for some reason
  var PostedBy = getColonSeparatedText(freecycleItemHTML, 'Posted by');
  var listingInfo = {
    PostUrl: freecycleItemUrl,
    RetrievedDate: new Date(),
    row: contextValues.lastRow,
    ID: listingId,
    ImageUrl: ImageUrl,
    Title: Title,
    Description: Description,
    Location: Location,
    Date: currDate,
    PostedBy: PostedBy,
  };

  updateCellRow(listingInfo);

  // Only email if nearby
  if (sheetsToEmail[contextValues.sheetName] ||
      Description.match(searchTerms) ||
      Title.match(searchTerms)) {
    Utilities.sleep(1000);
    sendEmail(listingInfo);
  }
}

function getColonSeparatedText(text, expression) {
  var regexExpr = new RegExp(expression + '\\s:.*?</div>');
  var match = text.match(regexExpr);
  if (match) {
    return trimHeader(trimHtml(match[0]));
  }

  return 'None';
}

function getElementByTag(text, tag) {
  var regexExpr = new RegExp('<' + tag + '[\\s\\S]*?<\\/' + tag + '>', 'g');
  return text.match(regexExpr);
}

function trimHtml(text) {
  return text.replace(/<.*?>/g, '');
}

function trimHeader(text) {
  return text.replace(/.*?:/, '').trim();
}

// Add item information to code
function addListingCell(row, key, value) {
  var cellColumn = contextValues.sheetIndex[key];
  if (cellColumn !== undefined) {
    var cellCode = NumberToLetters(cellColumn) + row;
    updateCell(cellCode, new Date(), value, true);
  }
}

// Send email with new listing information
function sendEmail(listingInfo) {
  var footer = '<hr>' +
  var imageDiv = listingInfo.ImageUrl ? '<img src="' + listingInfo.ImageUrl + '" alt="' + listingInfo.Title + '" width="128">' :
                                          '';
  var emailTemplate = listingInfo.Description + '<br>' +
                        listingInfo.Location + '<br>' +
                        listingInfo.Date + '<br>' +
                        '<br>' +
                        imageDiv +
                        '<br><hr><br>' +
                        'Email: <a href="mailto:' + listingInfo.id + '@posts.freecycle.org?subject=' + listingInfo.Title + ' posted to Freecycle" target="_blank">Reply by email</a><br>' +
                        'Url: <a href="' + listingInfo.PostUrl + '" target="_blank">Link</a>' +
                        '<hr>' +
                        footer;
  var subject = '[' + contextValues.sheetName + '] ' + listingInfo.Title + ' (' + listingInfo.Location + ')';


  // Get information from TotalSavings tab
  var email = MailApp.sendEmail({
    to: myEmail,
    subject: subject,
    htmlBody: emailTemplate,
  });
}

// Function that updates sheet
function updateCellRow(listingInfo) {
  var cellInfo = contextValues.sheetData[0].map(function getRowValues(header){
    return listingInfo[header];
  }).filter(function filterOutUndefined(clInfo){
    return clInfo !== undefined;
  });
  var cells = contextValues.sheet.getRange(listingInfo.row, 1, 1, cellInfo.length);
  cells.setValues([cellInfo]);
}


// Function that updates sheet
function updateCell(cellCode, _note, _message, _overwrite) {
  var cell = contextValues.sheet.getRange(cellCode);

  if (_note) {
    var currentNote = _overwrite ? '' : cell.getNote() + '\n';

    if (currentNote) {
      cell.setNote(currentNote + _note);
    } else {
      cell.setNote(_note);
    }
  }

  if (_message) {
    var currentMessage = _overwrite ? '' : cell.getValue() + '\n';

    if (currentMessage) {
      cell.setValue(currentMessage + _message);
    } else {
      cell.setValue(_message);
    }
  }
}
