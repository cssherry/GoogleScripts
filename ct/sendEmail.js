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
function updateSheet() {
  contextValues.sheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName('Current');
  contextValues.sheetData = contextValues.sheet.getDataRange().getValues();
  contextValues.sheetIndex = indexSheet(contextValues.sheetData);
  processPreviousListings();
  var mainPage = getMainPage().match(/<body[\s\S]*?<\/body>/)[0].replace(/<(no)?script[\s\S]*?<\/(no)?script>/g, '');
  var doc = Xml.parse(mainPage, true).getElement();
  var mainList = getElementsByTagName(doc, 'ul');
  var items = getElementsByTagName(mainList[2], 'li');
  archiveExpiredItems();
}

// Process previous data, including title and fee in case those change
function processPreviousListings() {
  contextValues.lastRow = numberOfRows(contextValues.sheetData);
  var idIdx = contextValues.sheetIndex.Url;
  var titleIdx = contextValues.sheetIndex.Title;
  var feeIdx = contextValues.sheetIndex.AdminFee;
  contextValues.previousListings = {};
  for (var i = 1; i < contextValues.lastRow; i++) {
    contextValues.previousListings[contextValues.sheetData[i][idIdx]] = {
      row: i,
      title: contextValues.sheetData[i][titleIdx],
      fee: contextValues.sheetData[i][feeIdx],
    };
  }
}

    }
  }

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
  if (!updated.length && !newItems.length) return;

  var footer = '<hr>' +
  var emailTemplate = newItems.length ? '<hr><h2>New:<h2><br>' + newItems.forEach(getElementSection) : '';
                      updated.length ? '<hr><h2>Updated:<h2><br>' + updated.forEach(getElementSection) : '';
                      footer;
  var subject = '[CT] ' + updated.length + ' Updated, ' + newItems.length + ' New ' + new Date();


  // Get information from TotalSavings tab
  var email = MailApp.sendEmail({
    to: myEmail,
    subject: subject,
    htmlBody: emailTemplate,
  });

  // HELPER
  var imageIdx = contextValues.sheetIndex.Image;
  var titleIdx = contextValues.sheetIndex.Title;
  var locationIdx = contextValues.sheetIndex.Location;
  var dateIdx = contextValues.sheetIndex.Date;
  var categoryIdx = contextValues.sheetIndex.Category;
  var urlIdx = contextValues.sheetIndex.Url;
  var feeIdx = contextValues.sheetIndex.AdminFee;
  function getElementSection(listingInfo) {
    var imageDiv = listingInfo[imageIdx] ? '<img src="' + listingInfo[imageIdx] + '" alt="' + listingInfo[titleIdx] + '" width="128">' :
    return '<h3>' + listingInfo[titleIdx] + '</h3><br>' +
           listingInfo[feeIdx] + '<br>' +
           listingInfo[locationIdx] + '<br>' +
           listingInfo[dateIdx] + '<br>' +
           listingInfo[categoryIdx] + '<br>' +
           '<br>' +
           imageDiv +
           '<br><br>' +
           'Url: <a href="' + listingInfo[urlIdx] + '" target="_blank">' + listingInfo[urlIdx] + '</a>' +
           '<hr>' +
  }
}

}

// Move expired items to "Archive" sheet
function archiveExpiredItems() {
  // Now archive events that passed
  var cutRange, newRange, currentItem, row;
  var archive = SpreadsheetApp.getActiveSpreadsheet()
                                        .getSheetByName('Archive');
  var archiveData = archive.getDataRange().getValues();
  var archiveIndex = indexSheet(archiveData);
  var lastArchiveRow = numberOfRows(archiveData);
  for (var expiredItem in contextValues.previousListings) {
    if (contextValues.previousListings.hasOwnProperty(expiredItem)) {
      lastArchiveRow++;
      currentItem = contextValues.previousListings[expiredItem];
      row = currentItem.row + 1;
      cutRange = contextValues.sheet.getRange('A' + row + 'I' + row);
      newRange = archive.getRange('A' + lastArchiveRow + 'I' + lastArchiveRow)
      newRange.setValues(cutRange.getValues());
      cutRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
  }
}
