// Needs GLOBALS_VARIABLES and indexSheet function
// GLOBALS_VARIABLES = {
//   headers: {
//     'x-famly-accesstoken': '',
//   },
//   urls: [''],

//   myEmails: [
//     ''
//   ],
// };

const lineSeparators = '\n\n-----------------------\n\n';
function pullAndUpdateEvents() {
    // Only run after 7 AM or before 7 PM on weekdays
    const currDate = new Date();
    const currentHour = currDate.getHours();
    const weekNum = currDate.getDay();
    if (currentHour < 7 || currentHour >= 19 || weekNum in [0, 6]) {
        return;
    }

    const allSheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = allSheet.getSheetByName('Tracker');
    const currRange = sheet.getDataRange();
    GLOBALS_VARIABLES.data = currRange.getValues();
    GLOBALS_VARIABLES.newData = [];
    GLOBALS_VARIABLES.index = indexSheet(GLOBALS_VARIABLES.data);
    GLOBALS_VARIABLES.loggedEvents = new Set();

    // Add previous data so we can not repeat
    GLOBALS_VARIABLES.data.forEach((row) => {
      const eventIdentifier = getIdentifier(row);
      GLOBALS_VARIABLES.loggedEvents.add(eventIdentifier)
    });

    // Save dates
    const startDate = getStartDate()
    GLOBALS_VARIABLES.startDate = startDate;
    GLOBALS_VARIABLES.endDate = parseDate(currDate);
    const dateIdx = GLOBALS_VARIABLES.index.LastUpdated + 1;
    currDate.setDate(currDate.getDate() - 1);

    // Go through each url set and parse dates
    GLOBALS_VARIABLES.urls.forEach(getAndParseEvents);

    // Save changes if there are any
    if (!GLOBALS_VARIABLES.newData.length) return;

    const startRow = sheet.getLastRow() + 1;
    const startCol = 1;
    const numRows = GLOBALS_VARIABLES.newData.length;
    const numCols = GLOBALS_VARIABLES.newData[0].length;
    const newRange = sheet.getRange(startRow, startCol, numRows, numCols);
    newRange.setValues(GLOBALS_VARIABLES.newData);

    const newDate = sheet.getRange(1, dateIdx + 1, 1, 1);
    newDate.setValues([[currDate]]);

    let famlySummary = '';
    const data = GLOBALS_VARIABLES.newData.map((row) => {
      famlySummary += `${row[GLOBALS_VARIABLES.index.Note]}${lineSeparators}`;
      return row.filter(item => !!item)
                .map(item => {
                  if (item.toString().startsWith('{') & item.toString().endsWith('}')) {
                    return JSON.stringify(JSON.parse(item), null, '    ');
                  }

                  return item;
                }).join('\n');
      }).join(lineSeparators);
    MailApp.sendEmail({
        to: GLOBALS_VARIABLES.myEmails.join(','),
        subject: `[Famly] New Events Logged ${GLOBALS_VARIABLES.endDate}`,
        body: famlySummary + data,
      });
}

function getAndParseEvents(baseUrl) {
    const fullUrl = getFullUrl(baseUrl);
    const returnedValue = JSON.parse(UrlFetchApp.fetch(fullUrl, {
        method: 'get',
        followRedirects: false,
        headers: GLOBALS_VARIABLES.headers,
    }).getContentText());

    returnedValue[0].days.forEach((dayObj) => {
        const { events } = dayObj;
        events.forEach(processEvent);
    });
}

const amountToDescription = {
   1: "Little",
   2: "Half",
   3: "Most",
   4: "All",
   5: "All+",
}

function processEvent(event) {
    if (!event.from) return;

    const dateIdx = GLOBALS_VARIABLES.index.Date;
    const actionIdx = GLOBALS_VARIABLES.index.Action;
    const noteIdx = GLOBALS_VARIABLES.index.Note;
    const infoIdx = GLOBALS_VARIABLES.index.FamilyInfo;
    const totalTime = GLOBALS_VARIABLES.index.TotalTime;
    const newDataRow = [];
    newDataRow[dateIdx] = new Date(event.from);
    newDataRow[actionIdx] = event.originator.type;
    newDataRow[noteIdx] = event.title;
    newDataRow[infoIdx] = JSON.stringify(event);

    if (event.to) {
        newDataRow[totalTime] = (new Date(event.to) - new Date(event.from)) /60/1000;
    } else {
        newDataRow[totalTime] = undefined;
    }

    if (event.embed && event.embed.mealItems && event.embed.mealItems.length) {
        newDataRow[noteIdx] += ` (${event.embed.mealItems.map(item => `${item.foodItem.title} - ${amountToDescription[item.amount] || item.amount}`).join(', ')})`;
    }

    const eventIdentifier = getIdentifier(newDataRow);
    if (!GLOBALS_VARIABLES.loggedEvents.has(eventIdentifier)) {
        GLOBALS_VARIABLES.newData.push(newDataRow);
        GLOBALS_VARIABLES.loggedEvents.add(eventIdentifier)
    }
}

// HELPER FUNCTIONS
function getIdentifier(row) {
    const fromIdx = GLOBALS_VARIABLES.index.Date;
    const actionIdx = GLOBALS_VARIABLES.index.Action;
    const noteIdx = GLOBALS_VARIABLES.index.Note;
    return `${row[fromIdx]}-${row[actionIdx]}-${row[noteIdx]}`;
}

function getStartDate() {
    const dateIdx = GLOBALS_VARIABLES.index.LastUpdated + 1;
    const startDate = parseDate(GLOBALS_VARIABLES.data[0][dateIdx]);
    return startDate;
}

function getFullUrl(url) {
    return `${url}&day=${GLOBALS_VARIABLES.startDate}&to=${GLOBALS_VARIABLES.endDate}`;
}

function parseDate(date) {
    const year = date.getFullYear();
    const month = padNumber(date.getMonth() + 1);
    const day = padNumber(date.getDate());
    return `${year}-${month}-${day}`;
}

function padNumber(num) {
    return num < 10 ? `0${num}` : num.toString();
}

// SHEET CUSTOM FUNCTION
// Based on https://stackoverflow.com/a/16086964
// https://github.com/darkskyapp/tz-lookup-oss/ -> https://github.com/darkskyapp/tz-lookup-oss/blob/master/tz.js from 8f09dc19104a006fa386ad86a69d26781ce31637
function getTimezoneTime(sheetDate, lat, long) {
    if (sheetDate instanceof Array) {
      const result = [];
      for (let rowIdx in sheetDate) {
        const currDate = sheetDate[rowIdx][0];
        const currLat = lat[rowIdx][0];
        const currLong = long[rowIdx][0];
        result[rowIdx] = convertFromPacific(currDate, currLat, currLong);
      }

      return result;
    }

    return convertFromPacific(sheetDate, lat, long);
}

// Array filter

// Array filter

// Array filter
function customArrayFilterJoin(joinText, prependText, range, ...restArguments) {
    const result = []
    let arrayLength = 1;
    restArguments.forEach((arg, idx) => {
        if (idx % 2 == 0 & arg instanceof Array) {
            arrayLength = arg.length;
        }
    });

    for (let arrayIdx = 0; arrayIdx < arrayLength; arrayIdx++) {
        const filteredItems = range.filter((_, idx) => {
            for (let argPairIndex = 0; argPairIndex < restArguments.length / 2; argPairIndex++) {
                const argIndex = argPairIndex * 2;
                const compareItem = restArguments[argIndex][idx];
                let staticItem = restArguments[argIndex + 1];
                staticItem = staticItem instanceof Array ? staticItem[arrayIdx][0] : staticItem;
                if (compareItem.toString() !== staticItem.toString()) return false;
            }

            return true;
        });
        const filteredText = filteredItems.length
            ? `${prependText}${filteredItems.join(joinText)}`
            : '';
        result.push(filteredText)
    }

    return result;
}

// CUSTOM FUNCTION HELPERS

// Faster isNan
function myIsNaN(val) {
	return !(val <= 0) && !(val > 0);
}

function convertFromPacific(date, latitude, longitude) {
    return changeTimezone(date, 'America/Los_Angeles', tzlookup(latitude, longitude));
}

// converting from time zone:https://stackoverflow.com/a/53652131
function changeTimezone(date, oldTimezone, newTimezone) {
    const oldDate = new Date(date.toLocaleString('en-US', {
        timeZone: oldTimezone,
    }));
    const newDate = new Date(date.toLocaleString('en-US', {
        timeZone: newTimezone,
    }));

    const diff = newDate.getTime() - oldDate.getTime();
    return new Date(date.getTime() + diff);
}
