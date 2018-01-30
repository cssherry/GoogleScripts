var sheetsToUpdate = ['ItemizedBudget', 'MonthlyTheater', 'Charities-notax'];

/** MAIN FUNCTION
    Runs upon change event, ideally new row being added to ItemizedBudget Sheet
*/
function createConversions() {
  var convertInstance = new convertUponNewRow();

  // Send email on Sunday
  var today = new Date();

  var alertTemplate = '<h1>SAVINGS (in USD):</h1>' +
                      '<b> Total (liquid): </b> { totalSavingsLiquid } <br/>' +
                      '<b> Total + Savings: </b> { totalSavingsAll } <br/>' +
                      '{ totalSavingsTable }' +
                      '<hr>';

  var emailTemplate = '<h1>THIS MONTH <span { monthStyle }>({ monthSpent }/{ monthBudget })</span>:</h1>' +
                      '{ itemsMonth }' +
                      '<hr>' +

                      '<h1>THIS YEAR <span { totalStyle }>({ totalSpent }/{ totalBudget })</span>:</h1>' +
                      '{ itemsTotal }' +
                      '<hr>' +

                      '<h1>CONVERSIONS (to USD):</h1>' +
                      '{ conversionUSD }' +
                      '<hr>' +

                      alertTemplate +

                      '<h1>ITEMS:</h1>' +
                      '{ itemList }' +
                      '<hr>' +


  var lastDayOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  var isLastDay = today.getDate() === lastDayOfMonth.getDate();
  var isSunday = today.getDay() === 0;
  var subject;


  // Get information from TotalSavings tab
  var stockAlerts = [];
  var alertInfo = {};
  var totalSavingSheet = SpreadsheetApp.getActiveSpreadsheet()
                                       .getSheetByName('TotalSavings');
  var totalSavingData = totalSavingSheet.getDataRange().getValues();
  var totalSavingIndex = indexSheet(totalSavingData);
  var currentDollarIdx = totalSavingIndex['CURRENT DOLLARS'];
  var currentPriceIdx = totalSavingIndex.Current;
  var highIdx = totalSavingIndex['Year High'];
  var lowIdx = totalSavingIndex['Year Low'];
  var nameIdx = totalSavingIndex.Ticker;
  var alertPercent = 0.01;
  var columnsToAdd = [
                        // Doesn't show up properly for some reason
                        // '52W growth', '1W growth',
                        'Current',
                        'Year High', 'Year Low',
                        'PE Ratio', 'Expense Ratio',
                        'NAME', 'CURRENT DOLLARS',
                      ];
  var columnsToColorCode = {
                            '52W growth': 0,
                            '1W growth': 0,
                            'Current': ['Year High', 'Year Low'],
                            'Expense Ratio': .5,
                          };
  var totalSavingsTable = '<table> <tr> <th>' + columnsToAdd.join('</th><th>') + '</th> </tr>';
  var row, currentColumn, currentColumnIdx, currentRow, currentValue, colorStyle = '', colorCompare;
  var currentPrice, highValue, lowValue, currentName;
  for (var i = 3; i < totalSavingData.length; i++) {
    var firstValue = totalSavingData[i][0];
    if (!firstValue) {
      break;
    }

    currentRow = totalSavingData[i];

    // Determine if the price is extreme enough to send alert
    currentPrice = parseFloat(currentRow[currentPriceIdx]);
    highValue = parseFloat(currentRow[highIdx]);
    lowValue = parseFloat(currentRow[lowIdx]);
    currentName = currentRow[nameIdx];
    if (currentName && currentPrice > highValue * (1 - alertPercent)) {
      stockAlerts.push(currentName);
      alertInfo[currentName] = '$' + currentPrice + ' within ' + alertPercent * 100 + '% of High ($' + highValue + ')';
    } else if (currentName && currentPrice < lowValue * (1 + alertPercent)) {
      stockAlerts.push(currentName);
      alertInfo[currentName] = '$' + currentPrice + ' within ' + alertPercent * 100 + '% of Low ($' + lowValue + ')';
    }

    row = '<tr>';
    for (var j = 0; j < columnsToAdd.length; j++) {
      currentColumn = columnsToAdd[j];
      currentColumnIdx = totalSavingIndex[currentColumn];
      currentValue = currentRow[currentColumnIdx];

      // Don't print 'N/A' or 'REF!' cells
      if (currentValue === '#N/A' || currentValue === '#REF!') {
        currentValue = '';
      }

      if (columnsToColorCode[currentColumn]) {
        colorCompare = columnsToColorCode[currentColumn];
        if (parseFloat(colorCompare) == colorCompare) {
          colorCompare = colorCompare;
        } else {
          colorCompare = currentRow[totalSavingIndex[colorCompare]];
        }

        colorStyle = getColorStyle(currentValue, colorCompare);
      } else {
        colorStyle = '';
      }

      row += ('<td ' + colorStyle + '>' + currentValue + '</td>');
    }

    totalSavingsTable += (row +'</tr>');
  }

  var emailOptions = {
    totalSavingsLiquid: '$' + roundMoney(totalSavingData[1][currentDollarIdx]),
    totalSavingsAll: '$' + roundMoney(totalSavingData[2][currentDollarIdx]),
    totalSavingsTable: totalSavingsTable,
  };

  // Send full email if Sunday or last day of the month
  if (isSunday || isLastDay) {
    subject = 'Weekly ';
    if (isSunday) {
      emailTemplate = '<h1>THIS WEEK <span { weekStyle }>({ weekSpent }/{ weekBudget })</span>:</h1>' +
                      '{ itemsWeek }' +
                      '<hr>' +
                      emailTemplate;
    }

    if (isLastDay) {
      subject = 'Monthly ';
      emailTemplate = '<h2>ANEESH TASKS:</h2>' +
                      '<ul>' +
                        '<li> Add tube/bus charges (4) </li>' +
                        '<li> Add home bills (7) </li>' +
                        '<li> Update TotalSavings/Make Investments (https://docs.google.com/spreadsheets/d/1wRSKZh7nMRJI2zICg7uKy2YXeNE1NYEfY2Ru8ZGkeS8/edit#gid=989997624) </li>' +
                      '</ul>' +
                      '<h2>SHERRY TASKS:</h2>' +
                      '<ul>' +
                        '<li> Update TotalSavings (https://docs.google.com/spreadsheets/d/1wRSKZh7nMRJI2zICg7uKy2YXeNE1NYEfY2Ru8ZGkeS8/edit#gid=989997624) </li>' +
                        '<li> Make donations (2) </li>' +
                      '</ul>' +
                      '<hr>' +
                      emailTemplate;
    }

    if (stockAlerts.length) {
      subject = '**' + subject;
      emailTemplate = processStocks(stockAlerts, alertInfo) + emailTemplate;
    }

    subject += 'Budget Report (' + today.toDateString() + ')';
    var overviewSheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName('Overview');
    var data = overviewSheet.getDataRange().getValues();
    var index = indexSheet(data);
    var itemIdx = index['Item'];
    var weekIdx = index['Items (Week)'];
    var monthIdx = index['Items (Month)'];
    var totalIdx = index['Items (Total)'];
    var budgetIdx = index.Budget;
    var actualIdx = index.Actual;
    var weekbudgetIdx = index['Week Budget'];
    var monthbudgetIdx = index['Month Budget'];
    var weekBudget = roundMoney(data[1][weekbudgetIdx]);
    var monthBudget = roundMoney(data[1][monthbudgetIdx]);
    var totalBudget = roundMoney(data[1][budgetIdx]);
    var itemsWeek = '';
    var itemsMonth = '';
    var itemsTotal = '';
    var itemList = '';
    var weekSpent = 0;
    var monthSpent = 0;
    var totalSpent = roundMoney(data[1][actualIdx]);
    var skipItems = ['Home', 'Retirement', 'Taxes', 'Savings'];

    // Get week/month/total and list of budget categories
    var itemName, weekData, monthData, totalData, currentValue, currentBudget;
    for (var i = 2; i < data.length; i++) {
      itemName = data[i][itemIdx];

      if (!itemName) {
        break;
      }

      weekData = data[i][weekIdx];
      monthData = data[i][monthIdx];
      totalData = data[i][totalIdx];
      if (weekData && weekData !== '#N/A') {
        currentValue = roundMoney(parseFloat(weekData.replace('£', '')));
        currentBudget = roundMoney(data[i][weekbudgetIdx]);

        itemsWeek += ('<h3 ' + getColorStyle(currentValue, currentBudget) + '>' + itemName + ' (£' + currentValue + '/£' + currentBudget + ')</h3>' + formatWeekMonth(weekData));

        if (skipItems.indexOf(itemName) === -1) {
          weekSpent += currentValue;
        }
      }

      if (monthData && monthData !== '#N/A') {
        currentValue = roundMoney(parseFloat(monthData.replace('£', '')));
        currentBudget = roundMoney(data[i][monthbudgetIdx]);
        itemsMonth += ('<h3 ' + getColorStyle(currentValue, currentBudget) + '>' + itemName + ' (£' + currentValue + '/£' + currentBudget + ')</h3>' + formatWeekMonth(monthData));

        if (skipItems.indexOf(itemName) === -1) {
          monthSpent += currentValue;
        }
      }

      if (totalData && totalData !== '#N/A') {
        currentValue = roundMoney(data[i][actualIdx]);
        currentBudget = roundMoney(data[i][budgetIdx]);
        totalData = '<ul><li>' + totalData.replace(' | ', '</li><li>') + '</li></ul>';
        itemsTotal += ('<h3 ' + getColorStyle(currentValue, currentBudget) + '>' + itemName + ' (£' + currentValue + '/£' + currentBudget + ')</h3>' + totalData);
      }

      itemList += ('<br>' + (i - 1) + ': ' + itemName);
    }

    // Get conversion information
    var allCodes = ['USD'];
    var RateTypeIdx = convertInstance.ItemizedBudgetIndex.RateType;
    for (var i = 1; i < convertInstance.ItemizedBudgetData.length; i++) {
      var rateValue = convertInstance.ItemizedBudgetData[i][RateTypeIdx];
      if (!rateValue) {
        break;
      }

      allCodes.push(rateValue);
    }

    var rateUrl = 'https://api.fixer.io/latest?base=GBP&symbols=' + allCodes.join(',');
    var response = UrlFetchApp.fetch(rateUrl);
    var conversionUSD = response.getContentText();

    emailOptions.itemsWeek = itemsWeek,
    emailOptions.itemsMonth = itemsMonth;
    emailOptions.itemsTotal = itemsTotal;
    emailOptions.conversionUSD = conversionUSD;
    emailOptions.itemList = itemList;
    emailOptions.weekBudget = '£' + weekBudget;
    emailOptions.monthBudget = '£' + monthBudget;
    emailOptions.totalBudget = '£' + totalBudget;
    emailOptions.weekSpent = '£' + roundMoney(weekSpent);
    emailOptions.monthSpent = '£' + roundMoney(monthSpent);
    emailOptions.totalSpent = '£' + roundMoney(totalSpent);
    emailOptions.weekStyle = getColorStyle(weekSpent, weekBudget);
    emailOptions.monthStyle = getColorStyle(monthSpent, monthBudget);
    emailOptions.totalStyle = getColorStyle(totalSpent, totalBudget);

    var email = new Email(myEmail, subject, emailTemplate, emailOptions);
    email.send();
  } else if (stockAlerts.length) {
    alertTemplate = processStocks(stockAlerts, alertInfo) + alertTemplate;
    var email = new Email(myEmail, 'Stock Alert (' + today.toDateString() + ': ' + stockAlerts.join(', ') + ')', alertTemplate, emailOptions);
    email.send();
  }
}

function processStocks(stockAlerts, alertInfo) {
  return '<h1>Stock Alerts</h1> <ul>' +
        stockAlerts.map(function concatAlerts(alert) {
          return '<li>' + alert + ': ' + alertInfo[alert] + '</li>';
        }).join('') +
        '</ul>';
}

// Format month and week data by removing total and making into list
function formatWeekMonth(fullData) {
  var resultString = '';
  var list = '';
  fullData.split('\n').forEach(function addToResult(line) {
    if (line) {
      // italicize total
      if (line[0] === '£') {
        resultString += ('<em>' + line + '</em>');
      } else if (line[0] === '-') {
        // Add items to list
        if (!list) {
          list = '<ul>';
        }

        list += ('<li>' + line.slice(1) + '</li>');
      } else {
        // Complete out list if there are any
        if (list) {
          resultString += (list + '</ul>');
          list = '';
        }

        // Add header
        resultString += ('<br/><b>' + line + '</b>');
      }
    }
  });

  if (list) {
    resultString += (list + '</ul>');
  }

  return resultString;
}

// Returns back correct styling given actual versus expected number
function getColorStyle (actual, expected) {
  var styleColor;
  actual = parseFloat(actual);
  expected = parseFloat(expected);
  if (actual < expected * .85) {
    styleColor = 'darkgreen';
  } else if (actual < expected * .95) {
    styleColor = 'green';
  } else if (actual <= expected) {
    styleColor = 'darkseagreen';
  } else if (actual > expected * 1.15) {
    styleColor = 'darkred';
  } else if (actual > expected * 1.05) {
    styleColor = 'mediumvioletred';
  } else {
    styleColor = 'indianred';
  }

  return 'style="color:' + styleColor + ';"';
}

function roundMoney(value) {
  return Math.round(parseFloat(value) * 100) / 100;
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
  var today = new Date();
  for (var i = startAndEnd.start; i < startAndEnd.end; i++) {
    var conversion = data[i][conversionIdx];
    var cellCode = NumberToLetters[index.ConversionRate] + (i + 1);
    if (conversion) {
      var conversion = this.getConversion(conversion);
      var note = 'Updated: ' + (conversion.cacheDay ? conversion.cacheDay : today);

      if (conversion.date) {
        note += '\nRate from: ' + conversion.date;
      }

      updateCell(sheetName, cellCode, note, conversion.rate, true);
    } else {
      updateCell(sheetName, cellCode, today, 1, true);
    }
  }
};

/** Returns back array of new rows without anything in ConversionRate column,
    or false if there's no new rows */
convertUponNewRow.prototype.getConversion = function (convertTo, _convertFrom) {
  convertTo = convertTo.trim().toUpperCase();
  _convertFrom = _convertFrom && _convertFrom.trim().toUpperCase();

  if (convertTo === _convertFrom) {
    return {
      rate: 1,
    };
  }

  var conversionRow = this.getConversionRow(convertTo);
  var today = new Date();
  var convertedDate = new Date(conversionRow.CacheDay);
  if (today.toDateString() === convertedDate.toDateString()) {
    return {
      rate: conversionRow.Rate,
      date: conversionRow.RateDate,
      cacheDay: conversionRow.CacheDay,
    };
  }

  return this.getOnlineRate(convertTo, _convertFrom, conversionRow.row);
};

/** Calculates if date is for current day */
convertUponNewRow.prototype.getConversionRow = function (convertTo) {
  var today = new Date();
  var rateIdx = this.ItemizedBudgetIndex.Rate;
  var cacheDayIdx = this.ItemizedBudgetIndex.CacheDay;
  var rateDateIdx = this.ItemizedBudgetIndex.RateDate;
  var rateTypeIdx = this.ItemizedBudgetIndex.RateType;
  var rateTypeData = this.ItemizedBudgetData;

  for (var i = 1; i < rateTypeData.length; i++) {
    if (!rateTypeData[i][rateTypeIdx] || rateTypeData[i][rateTypeIdx] === convertTo) {
      break;
    }
  }

  return {
    Rate: this.ItemizedBudgetData[i][rateIdx],
    CacheDay: this.ItemizedBudgetData[i][cacheDayIdx],
    RateDate: this.ItemizedBudgetData[i][rateDateIdx],
    row: i,
  };
};

/** Gets newest rate from API and adds it to row */
convertUponNewRow.prototype.getOnlineRate = function (convertTo, _convertFrom, rowIdx) {
  _convertFrom = _convertFrom || 'GBP';
  _convertFrom = _convertFrom.trim().toUpperCase();

  if (_convertFrom === convertTo) {
    return {
      rate: 1,
    };
  }

  var url = 'https://api.fixer.io/latest?base=' + convertTo + '&symbols=﻿' + _convertFrom;
  var response = UrlFetchApp.fetch(url);
  var conversionData = JSON.parse(response.getContentText());
  var rate = conversionData.rates[_convertFrom];
  var dateUpdated = 'Rate from: ' + conversionData.date;
  var row = rowIdx + 1;
  var RateIdx = this.ItemizedBudgetIndex.Rate;
  var CacheDayIdx = this.ItemizedBudgetIndex.CacheDay;
  var RateDateIdx = this.ItemizedBudgetIndex.RateDate;
  var RateTypeIdx = this.ItemizedBudgetIndex.RateType;
  var RateCell = NumberToLetters[RateIdx] + row;
  var CacheDayCell = NumberToLetters[CacheDayIdx] + row;
  var RateDateCell = NumberToLetters[RateDateIdx] + row;
  var RateTypeCell = NumberToLetters[RateTypeIdx] + row;
  var today = new Date();

  updateCell('ItemizedBudget', RateCell, dateUpdated, rate, true);
  updateCell('ItemizedBudget', CacheDayCell, dateUpdated, today, true);
  updateCell('ItemizedBudget', RateDateCell, dateUpdated, conversionData.date, true);
  updateCell('ItemizedBudget', RateTypeCell, dateUpdated, convertTo, true);
  this.ItemizedBudgetData[rowIdx][RateIdx] = rate;
  this.ItemizedBudgetData[rowIdx][CacheDayIdx] = today;
  this.ItemizedBudgetData[rowIdx][RateDateIdx] = conversionData.date;
  this.ItemizedBudgetData[rowIdx][RateTypeIdx] = convertTo;

  return {
    rate: rate,
    date: conversionData.date,
    cacheDay: today,
  };
};

// Function that records when an email is successfully sent
function updateCell(sheetName, cellCode, _note, _message, _overwrite) {
  var cell = SpreadsheetApp.getActiveSpreadsheet()
                           .getSheetByName(sheetName)
                           .getRange(cellCode);

  if (_note) {
    var currentNote = _overwrite ? '' : cell.getNote() + '\n';
    cell.setNote(currentNote + _note);
  }

  if (_message) {
    var currentMessage = _overwrite ? '' : cell.getValue() + '\n';
    cell.setValue(currentMessage + _message);
  }
}

// Evaluate
function evaluate(cellValue) {
  if (cellValue instanceof Array) {
    var i, j;
    for (i in cellValue) {
      for (j in cellValue[i]) {
        if (cellValue[i][j] && cellValue[i][j].match && cellValue[i][j].match('[/*/+-]')) {
          cellValue[i][j] = eval(cellValue[i][j]);
        }
      }
    }
    return cellValue;
  }

  if (cellValue && cellValue.match && cellValue.match('[/*/+-]')) {
    return eval(cellValue);
  }

  return cellValue;
}
