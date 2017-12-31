var sheetsToUpdate = ['Itemized Budget', 'MonthlyTheater', 'Charities-notax'];

/** MAIN FUNCTION
    Runs upon change event, ideally new row being added to Itemized Budget Sheet
*/
function createConversions() {
  new convertUponNewRow();
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
  var rateIdx = this['Itemized BudgetIndex'].Rate;
  var cacheDayIdx = this['Itemized BudgetIndex'].CacheDay;
  var rateTypeIdx = this['Itemized BudgetIndex'].RateType;
  var rateTypeData = this['Itemized BudgetData'];

  for (var i = 0; i < rateTypeData.length; i++) {
    if (!rateTypeData[i][rateTypeIdx] || rateTypeData[i][rateTypeIdx] === convertTo) {
      break;
    }
  }

  return {
    Rate: this['Itemized BudgetData'][i][rateIdx],
    CacheDay: this['Itemized BudgetData'][i][cacheDayIdx],
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
  var RateIdx = this['Itemized BudgetIndex'].Rate;
  var CacheDayIdx = this['Itemized BudgetIndex'].CacheDay;
  var RateTypeIdx = this['Itemized BudgetIndex'].RateType;
  var RateCell = NumberToLetters[RateIdx] + row;
  var CacheDayCell = NumberToLetters[CacheDayIdx] + row;
  var RateTypeCell = NumberToLetters[RateTypeIdx] + row;
  var today = new Date();

  updateCell('Itemized Budget', RateCell, dateUpdated, rate, true);
  updateCell('Itemized Budget', CacheDayCell, dateUpdated, today, true);
  updateCell('Itemized Budget', RateTypeCell, dateUpdated, convertTo, true);
  this['Itemized BudgetData'][rowIdx][RateIdx] = rate;
  this['Itemized BudgetData'][rowIdx][CacheDayIdx] = today;
  this['Itemized BudgetData'][rowIdx][RateTypeIdx] = convertTo;

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
