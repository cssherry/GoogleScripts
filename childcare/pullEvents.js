// Needs GLOBALS_VARIABLES and indexSheet function
// GLOBALS_VARIABLES = {
//   headers: {
//     'x-famly-accesstoken': '',
//   },
//   urls: [''],
//
//   myEmails: [ // Daycare event emails
//     ''
//   ],
//
//   summaryEmails: [ // Daily summary emails
//     ''
//   ],
//
//   folderId: '',
//
//   messageUrl: '',
//   postUrl: '',
//   feedItemUrl: '',
// };

const lineSeparators = '\n\n-----------------------\n\n';
const maxLimit = 100;
const idDelimiter = ',';
const attachDelimiter = ' ; ';
const messageType = 'Message';
const postType = 'Post';
function pullAndUpdateEvents() {
    // Only run after 7 AM or before 7 PM on weekdays
    const currDate = new Date();
    const currentHour = currDate.getHours();
    const weekNum = currDate.getDay();
    if (currentHour < 8 || currentHour >= 20 || weekNum in [0, 6]) {
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

    // Now parse Family events
    const familySheet = allSheet.getSheetByName('Famly');
    const familyRange = familySheet.getDataRange();
    GLOBALS_VARIABLES.familyData = familyRange.getValues();
    GLOBALS_VARIABLES.newFamilyData = [];
    GLOBALS_VARIABLES.familyIndex = indexSheet(GLOBALS_VARIABLES.familyData);
    GLOBALS_VARIABLES.familyLoggedEvents = {};
    GLOBALS_VARIABLES.familyLoggedEvents[messageType] = {};
    GLOBALS_VARIABLES.familyLoggedEvents[postType] = {};

    // Process existing posts/messages
    const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
    const idIdx = GLOBALS_VARIABLES.familyIndex.SelfId;
    GLOBALS_VARIABLES.familyData.forEach((row, idx) => {
        if (idx === 0) return;
        if (!row[0]) return;
        const type = row[typeIdx];
        const ids = row[idIdx].split(idDelimiter);
        ids.forEach((id) => {
            GLOBALS_VARIABLES.familyLoggedEvents[type][id] = idx;
            GLOBALS_VARIABLES.familyLoggedEvents[type][id] = idx;
        });
    });

    // Check for new messages
    getAndParseMessages();

    // Check for new posts
    getAndParsePosts();

    // Save changes if there are any
    if (!GLOBALS_VARIABLES.newData.length && !GLOBALS_VARIABLES.newFamilyData.length) return;

    if (GLOBALS_VARIABLES.newData.length) {
      appendRows(sheet, GLOBALS_VARIABLES.newData);
    }

    if (GLOBALS_VARIABLES.newFamilyData.length) {
      appendRows(familySheet, GLOBALS_VARIABLES.newFamilyData);
    }

    const newDate = sheet.getRange(1, dateIdx + 1, 1, 1);
    newDate.setValues([[currDate]]);

    let famlySummary = '';
    const loggedData = GLOBALS_VARIABLES.newData.map((row) => {
      famlySummary += `${row[GLOBALS_VARIABLES.index.Note]}${lineSeparators}`;
      return row.filter(item => !!item)
                .map(item => {
                  if (item.toString().startsWith('{') & item.toString().endsWith('}')) {
                    return JSON.stringify(JSON.parse(item), null, '    ');
                  }

                  return item;
                }).join('\n');
    }).join(lineSeparators);

    // Generate text for new messages/posts
    const separator = GLOBALS_VARIABLES.newFamilyData.length ? lineSeparators : '';
    const daycareGeneral = GLOBALS_VARIABLES.newFamilyData.map((row) => {
        const type = row[GLOBALS_VARIABLES.familyIndex.Type];
        const content = row[GLOBALS_VARIABLES.familyIndex.Content];
        const attachments = row[GLOBALS_VARIABLES.familyIndex.Attachments];
        const jsonData = JSON.stringify(JSON.parse(row[GLOBALS_VARIABLES.familyIndex.FamilyInfo]), null, '    ');
        const header = `${type}\n${content}`
        famlySummary += header + lineSeparators;

        return `${header}\n\n${attachments}`;
    }).join(lineSeparators);

    // SEND EMAIL
    MailApp.sendEmail({
        to: GLOBALS_VARIABLES.myEmails.join(','),
        subject: `[Famly] New Events Logged ${GLOBALS_VARIABLES.endDate}`,
        body: famlySummary + loggedData + separator + daycareGeneral,
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

function getAndParseMessages() {
    const hasChanged = [];
    const fullUrl = `${GLOBALS_VARIABLES.messagesUrl}?limit=${maxLimit}`;
    const conversationList = JSON.parse(UrlFetchApp.fetch(fullUrl, {
        method: 'get',
        followRedirects: false,
        headers: GLOBALS_VARIABLES.headers,
    }).getContentText());

    conversationList.forEach((convo) => {
        const messageId = convo.lastMessage.messageId;
        if (!isLogged(messageId, messageType)) {
            hasChanged.push(convo.conversationId);
        }
    });

    const dateIdx = GLOBALS_VARIABLES.familyIndex.Date;
    const chainIdx = GLOBALS_VARIABLES.familyIndex.ChainId;
    const selfId = GLOBALS_VARIABLES.familyIndex.SelfId;
    const lastUpdateIdx = GLOBALS_VARIABLES.familyIndex.LastDate;
    const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
    const contentIdx = GLOBALS_VARIABLES.familyIndex.Content;
    const attachmentIdx = GLOBALS_VARIABLES.familyIndex.Attachments;
    const infoIdx = GLOBALS_VARIABLES.familyIndex.FamilyInfo;
    hasChanged.forEach((newConvoId) => {
        const conversationUrl = `${GLOBALS_VARIABLES.messagesUrl}/${newConvoId}`;
        const conversationData = JSON.parse(UrlFetchApp.fetch(conversationUrl, {
            method: 'get',
            followRedirects: false,
            headers: GLOBALS_VARIABLES.headers,
        }).getContentText());
        const newMessages = [];
        newMessages[dateIdx] = new Date();
        newMessages[chainIdx] = conversationData.conversationId;
        newMessages[lastUpdateIdx] = conversationData.lastActivityAt;
        newMessages[typeIdx] = messageType;
        newMessages[infoIdx] = JSON.stringify(conversationData);

        const {
            newMessageIds,
            newContent,
            attachmentUrls,
        } = processMessage(conversationData.messages);

        newMessages[selfId] = newMessageIds.join(idDelimiter);
        newMessages[contentIdx] = newContent;
        newMessages[attachmentIdx] = attachmentUrls.join(attachDelimiter);
        GLOBALS_VARIABLES.newFamilyData.push(newMessages);
    });
}

function processMessage(messages) {
    const newMessageIds = [];
    let newContent = '';
    const attachmentUrls = [];

    messages.forEach((message) => {
        if (!isLogged(message.messageId, messageType)) {
            newMessageIds.push(message.messageId);
            newContent += `- ${message.body}\n`
            const newFiles = downloadFiles(message);
            if (newFiles.length) {
                attachmentUrls.push(...newFiles);
            }
        }
    });

    return {
        newMessageIds,
        newContent,
        attachmentUrls,
    }
}

function getAndParsePosts() {
    const hasChanged = [];
    const fullUrl = `${GLOBALS_VARIABLES.postUrl}?limit=${maxLimit}`;
    const postList = JSON.parse(UrlFetchApp.fetch(fullUrl, {
        method: 'get',
        followRedirects: false,
        headers: GLOBALS_VARIABLES.headers,
    }).getContentText());

    const dateIdx = GLOBALS_VARIABLES.familyIndex.Date;
    const chainIdx = GLOBALS_VARIABLES.familyIndex.ChainId;
    const selfId = GLOBALS_VARIABLES.familyIndex.SelfId;
    const lastUpdateIdx = GLOBALS_VARIABLES.familyIndex.LastDate;
    const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
    const contentIdx = GLOBALS_VARIABLES.familyIndex.Content;
    const attachmentIdx = GLOBALS_VARIABLES.familyIndex.Attachments;
    const infoIdx = GLOBALS_VARIABLES.familyIndex.FamilyInfo;

    postList.forEach((post) => {
        const messageId = post.target.feedItemId || post.notificationId;
        if (!isLogged(messageId, postType)) {
            if (post.target.feedItemId) {
                hasChanged.push(messageId);
            } else {
                const newMessage = [];
                newMessage[dateIdx] = new Date();
                newMessage[chainIdx] = post.notificationId;
                newMessage[selfId] = post.notificationId;
                newMessage[lastUpdateIdx] = post.createdDate;
                newMessage[typeIdx] = postType;
                newMessage[infoIdx] = JSON.stringify(post);
                newMessage[contentIdx] = post.body;
                GLOBALS_VARIABLES.newFamilyData.push(newMessage);
            }
        }
    });

    hasChanged.forEach((newPostId) => {
        const postUrl = `${GLOBALS_VARIABLES.feedItemUrl}?feedItemId=${newPostId}`;
        const postData = JSON.parse(UrlFetchApp.fetch(postUrl, {
            method: 'get',
            followRedirects: false,
            headers: GLOBALS_VARIABLES.headers,
        }).getContentText());
        const newMessages = [];
        newMessages[dateIdx] = new Date();
        newMessages[chainIdx] = postData.feedItem.originatorId;
        newMessages[selfId] = postData.feedItem.feedItemId;
        newMessages[lastUpdateIdx] = postData.feedItem.createdDate;
        newMessages[typeIdx] = postType;
        newMessages[infoIdx] = JSON.stringify(postData);

        newMessages[contentIdx] = postData.feedItem.body;
        newMessages[attachmentIdx] = downloadFiles(postData.feedItem).join('\n');
        GLOBALS_VARIABLES.newFamilyData.push(newMessages);
    });

}

function downloadFiles(containerObj) {
    const attachments = [];
    const createDate = containerObj.createdDate || containerObj.createdAt || 'Unknown Date';
    const body = containerObj.body || '';
    const description = `${body}\n\nShared on: ${createDate}`;
    if (containerObj.files.length) {
        containerObj.files.forEach((fileObj) => {
            const fileUrl = uploadFile(fileObj.url, fileObj.filename, description);
            attachments.push(fileUrl);
        });
    }

    if (containerObj.images.length) {
        containerObj.images.forEach((imgObj) => {
            const url = `${imgObj.prefix}/${imgObj.width}x${imgObj.height}/${imgObj.key}`;
            const fileName = imgObj.key.replace(/\//g, '_').replace(/\?.*/, '');
            const fileUrl = uploadFile(url, fileName, description);
            attachments.push(fileUrl);
        });
    }

    if (containerObj.videos && containerObj.videos.length) {
        containerObj.videos.forEach((videoObj) => {
            const videoName = `${parseDate(new Date(createDate))}_video_${videoObj.videoId}`;
            const fileUrl = uploadFile(videoObj.videoUrl, videoName, description, true);
            attachments.push(fileUrl);
        });
    }

    return attachments;
}

function uploadFile(fileUrl, fileName, additionalDescription, keepExtension = false) {
    const response = UrlFetchApp.fetch(fileUrl);
    if (!GLOBALS_VARIABLES.googleDrive) {
        GLOBALS_VARIABLES.googleDriveExistingFiles = {};
        GLOBALS_VARIABLES.googleDrive = DriveApp.getFolderById(GLOBALS_VARIABLES.folderId);
        const existingFiles = GLOBALS_VARIABLES.googleDrive.getFiles();
        while (existingFiles.hasNext()) {
            const file = existingFiles.next();
            const fileName = file.getName();
            const fileUrl = file.getUrl();
            GLOBALS_VARIABLES.googleDriveExistingFiles[fileName] = fileUrl;
        }
    }

    if (GLOBALS_VARIABLES.googleDriveExistingFiles[fileName]) {
        return GLOBALS_VARIABLES.googleDriveExistingFiles[fileName];
    }

    const blob = response.getBlob();
    const file = GLOBALS_VARIABLES.googleDrive.createFile(blob);

    if (keepExtension) {
        const extension = file.getName().match(/\..*$/)[0];
        if (extension) {
            fileName += extension;
        }
    }

    file.setName(fileName);
    file.setDescription(`Download from ${fileUrl} on ${new Date()}\n\n${additionalDescription}`);
    return file.getUrl();
}

const amountToDescription = {
   1: "Little",
   2: "Half",
   3: "Most",
   4: "All",
   5: "All+",
}

function isLogged(messageId, type) {
    return messageId in GLOBALS_VARIABLES.familyLoggedEvents[type];
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

function appendRows(sheet, newData) {
    const startRow = sheet.getLastRow() + 1;
    const startCol = 1;
    const numRows = newData.length;
    const numCols = newData[0].length;
    const newRange = sheet.getRange(startRow, startCol, numRows, numCols);
    newRange.setValues(newData);
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
