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

    MailApp.sendEmail({
        to: GLOBALS_VARIABLES.myEmails.join(','),
        subject: '[Famly] New Events Logged',
        body: GLOBALS_VARIABLES.newData.map((row) => row.filter(item => !!item).join('\n')).join('\n\n\n'),
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
    newDataRow[infoIdx] = JSON.stringify(event, null, '  ');

    if (event.to) {
        newDataRow[totalTime] = (new Date(event.to) - new Date(event.from)) /60/1000;
    } else {
        newDataRow[totalTime] = undefined;
    }

    if (event.embed && event.embed.mealItems && event.embed.mealItems.length) {
        newDataRow[noteIdx] += ` (${event.embed.mealItems.map(item => `${item.foodItem.title} - ${item.amount}`).join(', ')})`;
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
    const actionIdx = GLOBALS_VARIABLES.index.Action
    return `${row[fromIdx]}-${row[actionIdx]}`;
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
    const month = padNumber(date.getMonth());
    const day = padNumber(date.getDate());
    return `${year}-${month}-${day}`;
}

function padNumber(num) {
    return num < 10 ? `0${num}` : num.toString();
}