var sheetsToUpdate = ['ItemizedBudget', 'MonthlyTheater', 'Charities-notax'];

/** MAIN FUNCTION
    Runs upon change event, ideally new row being added to ItemizedBudget Sheet
*/
function createConversions() {
  var convertInstance = new convertUponNewRow();

  // Send email on Sunday
  var today = new Date();
  var emailTemplate = 'THIS WEEK:\n' +
                      '{ itemsWeek }\n\n' +
                      'THIS MONTH:\n' +
                      '{ itemsMonth }\n\n' +
                      'THIS YEAR:\n' +
                      '{ itemsTotal }\n\n' +
                      '-----------\n' +
                      'CONVERSIONS (to USD):\n' +
                      '{ conversionUSD }\n\n' +
                      '-----------\n' +
                      'ITEMS:' +
                      '{ itemList }';

  if (today.getDay() === 0) {
    var subject = 'Weekly Budget Report (' + today.toDateString() + ')';
    var overviewSheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName('Overview');
    var data = overviewSheet.getDataRange().getValues();
    var index = indexSheet(data);
    var itemIdx = index['Item'];
    var weekIdx = index['Items (Week)'];
    var monthIdx = index['Items (Month)'];
    var totalIdx = index['Items (Total)'];
    var actualIdx = index.Actual;
    var itemsWeek = '';
    var itemsMonth = '';
    var itemsTotal = '';
    var itemList = '';
    for (var i = 2; i < data.length; i++) {
      var itemName = data[i][itemIdx];

      if (!itemName) {
        break;
      }

      itemsWeek += ('\n' + itemName + ':\n' + data[i][weekIdx]);
      itemsMonth += ('\n' + itemName + ':\n' + data[i][monthIdx]);
      itemsTotal += ('\n' + itemName + ' (£' + data[i][actualIdx] + '):\n' + data[i][totalIdx]);
      itemList += ('\n' + (i - 1) + ': ' + itemName);
    }

    var allCodes = ['USD'];
    var RateTypeIdx = convertInstance['ItemizedBudgetIndex'].RateType;
    for (var i = 1; i < convertInstance['ItemizedBudgetData'].length; i++) {
      var rateValue = convertInstance['ItemizedBudgetData'][i][RateTypeIdx];
      if (!rateValue) {
        break;
      }

      allCodes.push(rateValue);
    }

    var rateUrl = 'https://api.fixer.io/latest?base=GBP&symbols=' + allCodes.join(',');
    var response = UrlFetchApp.fetch(rateUrl);
    var conversionUSD = response.getContentText();

    var emailOptions = {
      itemsWeek: itemsWeek,
      itemsMonth: itemsMonth,
      itemsTotal: itemsTotal,
      conversionUSD: conversionUSD,
      itemList: itemList,
    };
    email = new Email(myEmail, subject, emailTemplate, emailOptions);
    email.send();
  }
}

function convertUponNewRow() {
  // Index all the sheets
  for (var i = 0; i < sheetsToUpdate.length; i++) {
    this.getDataAndIndex(sheetsToUpdate[i]);
  }

  // Figure out if there's new ConversionRate rows
  var hasNewRow = this.hasNewRow();
  if (hasNewRow.hasNew.length) {
    for (var j = 0; j < hasNewRow.hasNew.length; j++) {
      var sheetName = hasNewRow.hasNew[j];
      this.updateCell(sheetName, hasNewRow[sheetName]);
    }
  }
}

/** HELPER FUNCTION */
/** Gets data and indexs relevant sheets */
convertUponNewRow.prototype.getDataAndIndex = function (sheetName) {
  var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet()
                                    .getSheetByName(sheetName);
  this[sheetName + 'Data'] = scheduleSheet.getDataRange().getValues();
  this[sheetName + 'Index'] = indexSheet(this[sheetName + 'Data']);
};

/** Returns back array of new rows without anything in ConversionRate column,
    or false if there's no new rows */
convertUponNewRow.prototype.hasNewRow = function () {
  var result = {
    hasNew: [],
  };

  for (var i = 0; i < sheetsToUpdate.length; i++) {
    var currentSheet = sheetsToUpdate[i];
    var currentSheetIndex = this[currentSheet + 'Index'];
    var currentSheetData = this[currentSheet + 'Data'];
    var idxConversionRate = currentSheetIndex['ConversionRate'];
    var idxDate = currentSheetIndex['Date'];
    var numberConversionRate = numberOfRows(currentSheetData, idxConversionRate);
    var numberDate = numberOfRows(currentSheetData, idxDate);
    if (numberConversionRate !== numberDate) {
      result[currentSheet] = {
        start: numberConversionRate,
        end: numberDate,
      };

      result.hasNew.push(currentSheet);
    }
  }

  return result;
};

/** Updates ConversionRate for start row to end row of specified sheet */
convertUponNewRow.prototype.updateCell = function (sheetName, startAndEnd) {
  var index = this[sheetName + 'Index'];
  var data = this[sheetName + 'Data'];
  var conversionIdx = index.Conversion;
  for (var i = startAndEnd.start; i < startAndEnd.end; i++) {
    var conversion = data[i][conversionIdx];
    var cellCode = NumberToLetters[index.ConversionRate] + (i + 1);
    if (conversion) {
      var conversion = this.getConversion(conversion);
      updateCell(sheetName, cellCode, new Date(), conversion, true);
    } else {
      updateCell(sheetName, cellCode, new Date(), 1, true);
    }
  }
};

/** Returns back array of new rows without anything in ConversionRate column,
    or false if there's no new rows */
convertUponNewRow.prototype.getConversion = function (convertTo, _convertFrom) {
  convertTo = convertTo.trim().toUpperCase();
  var conversionRow = this.getConversionRow(convertTo);
  var today = new Date();
  var convertedDate = new Date(conversionRow.CacheDay);
  if (today.toDateString() === convertedDate.toDateString()) {
    return conversionRow.Rate;
  }

  return this.getOnlineRate(convertTo, _convertFrom, conversionRow.row);
};

/** Calculates if date is for current day */
convertUponNewRow.prototype.getConversionRow = function (convertTo) {
  var today = new Date();
  var rateIdx = this['ItemizedBudgetIndex'].Rate;
  var cacheDayIdx = this['ItemizedBudgetIndex'].CacheDay;
  var rateTypeIdx = this['ItemizedBudgetIndex'].RateType;
  var rateTypeData = this['ItemizedBudgetData'];

  for (var i = 1; i < rateTypeData.length; i++) {
    if (!rateTypeData[i][rateTypeIdx] || rateTypeData[i][rateTypeIdx] === convertTo) {
      break;
    }
  }

  return {
    Rate: this['ItemizedBudgetData'][i][rateIdx],
    CacheDay: this['ItemizedBudgetData'][i][cacheDayIdx],
    row: i,
  };
};

/** Gets newest rate from API and adds it to row */
convertUponNewRow.prototype.getOnlineRate = function (convertTo, _convertFrom, rowIdx) {
  _convertFrom = _convertFrom || 'GBP';
  _convertFrom = _convertFrom.trim().toUpperCase();
  var url = 'https://api.fixer.io/latest?base=' + convertTo + '&symbols=﻿' + _convertFrom;
  var response = UrlFetchApp.fetch(url);
  var conversionData = JSON.parse(response.getContentText());
  var rate = conversionData.rates[_convertFrom];
  var dateUpdated = 'Rate from: ' + conversionData.date;
  var row = rowIdx + 1;
  var RateIdx = this['ItemizedBudgetIndex'].Rate;
  var CacheDayIdx = this['ItemizedBudgetIndex'].CacheDay;
  var RateTypeIdx = this['ItemizedBudgetIndex'].RateType;
  var RateCell = NumberToLetters[RateIdx] + row;
  var CacheDayCell = NumberToLetters[CacheDayIdx] + row;
  var RateTypeCell = NumberToLetters[RateTypeIdx] + row;
  var today = new Date();

  updateCell('ItemizedBudget', RateCell, dateUpdated, rate, true);
  updateCell('ItemizedBudget', CacheDayCell, dateUpdated, today, true);
  updateCell('ItemizedBudget', RateTypeCell, dateUpdated, convertTo, true);
  this['ItemizedBudgetData'][rowIdx][RateIdx] = rate;
  this['ItemizedBudgetData'][rowIdx][CacheDayIdx] = today;
  this['ItemizedBudgetData'][rowIdx][RateTypeIdx] = convertTo;

  return rate;
};

// Function that records when an email is successfully sent
function updateCell(sheetName, cellCode, _note, _message, _overwrite) {
  var cell = SpreadsheetApp.getActiveSpreadsheet()
                           .getSheetByName(sheetName)
                           .getRange(cellCode);

  if (_note) {
    var currentNote = _overwrite ? "" : cell.getNote() + "\n";
    cell.setNote(currentNote + _note);
  }

  if (_message) {
    var currentMessage = _overwrite ? "" : cell.getValue() + "\n";
    cell.setValue(currentMessage + _message);
  }
}

// Evaluate
function evaluate(cellValue) {
  if (cellValue instanceof Array) {
    for (i in cellValue){
      for (j in cellValue[i]){
        if (cellValue[i][j].match && cellValue[i][j].match('[/*/+-]')) {
          cellValue[i][j] = eval(cellValue[i][j]);
        }
      }
    }
    return cellValue;
  } else {
    return eval(cellValue);
  }
}