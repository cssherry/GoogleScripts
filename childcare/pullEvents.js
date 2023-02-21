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
//   missedPostUrl: '',
//   feedItemUrl: '',
//   incidentUrl: '',
//   graphqlUrl: '',
//   graphqlQuery: { query, operationName, variables: {observationIds:[]}, },
// };

const lineSeparators = '\n\n-----------------------\n\n';
const maxLimit = 100;
const idDelimiter = ',';
const attachDelimiter = ' ; ';
const messageType = 'Message';
const postType = 'Post';
const incidentType = 'Incident';
function pullAndUpdateEvents() {
  // Only run after 7 AM or before 7 PM on weekdays
  const currDate = new Date();
  const currentHour = currDate.getHours();
  const weekNum = currDate.getDay();
  if (currentHour < 8 || currentHour >= 22 || [0, 6].includes(weekNum)) {
    return;
  }

  GLOBALS_VARIABLES.startTime = new Date();
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
    GLOBALS_VARIABLES.loggedEvents.add(eventIdentifier);
  });

  // Save dates
  const startDate = getStartDate();
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
  GLOBALS_VARIABLES.familyLoggedEvents[incidentType] = {};

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

  // Add bookmarks
  getAndParseBookmarks();

  // Add any incident reports
  getAndParseIncidents();

  // Add any incident reports
  getAndParseObservations();

  // Save changes if there are any
  if (
    !GLOBALS_VARIABLES.newData.length &&
    !GLOBALS_VARIABLES.newFamilyData.length
  )
    return;

  // SEND EMAIL
  let famlySummary = '';
  const loggedData = GLOBALS_VARIABLES.newData
    .map((row) => {
      famlySummary += `${row[GLOBALS_VARIABLES.index.Note]} (${
        row[GLOBALS_VARIABLES.index.Date]
      })${lineSeparators}`;
      return row
        .filter((item) => !!item)
        .map((item) => {
          if (item.toString().startsWith('{') & item.toString().endsWith('}')) {
            return JSON.stringify(JSON.parse(item), null, '    ');
          }

          return item;
        })
        .join('\n');
    })
    .join(lineSeparators);

  // Generate text for new messages/posts
  const separator = GLOBALS_VARIABLES.newFamilyData.length
    ? lineSeparators
    : '';
  const daycareGeneral = GLOBALS_VARIABLES.newFamilyData
    .map((row) => {
      const type = row[GLOBALS_VARIABLES.familyIndex.Type];
      const date = row[GLOBALS_VARIABLES.familyIndex.LastDate];
      const content = row[GLOBALS_VARIABLES.familyIndex.Content];
      const attachments = row[GLOBALS_VARIABLES.familyIndex.Attachments] || '';
      const attachmentText = attachments.split(attachDelimiter).map((link, idx) => link ? `#${idx + 1}: ${link} ` : '').join(' \n');
      let fromInfo = row[GLOBALS_VARIABLES.familyIndex.From];
      fromInfo = fromInfo ? `from ${fromInfo}` : '';
      const header = `${type} ${fromInfo} (${date})\n\n${content}`;
      famlySummary += header + lineSeparators;

      const chainId = row[GLOBALS_VARIABLES.familyIndex.ChainId];
      const selfId = row[GLOBALS_VARIABLES.familyIndex.SelfId];
      return `${header}\nChain: ${chainId}\nSelf: ${selfId}\n\n${attachmentText}`;
    })
    .join(lineSeparators);

  MailApp.sendEmail({
    to: GLOBALS_VARIABLES.myEmails.join(','),
    subject: `[Famly] New Events Logged ${GLOBALS_VARIABLES.endDate}`,
    body: famlySummary + loggedData + separator + daycareGeneral,
  });

  // UPDATE ROWS
  if (GLOBALS_VARIABLES.newData.length) {
    appendRows(sheet, GLOBALS_VARIABLES.newData);
  }

  if (GLOBALS_VARIABLES.newFamilyData.length) {
    appendRows(
      familySheet,
      GLOBALS_VARIABLES.newFamilyData,
      GLOBALS_VARIABLES.familyIndex.Attachments
    );
  }

  const newDate = sheet.getRange(1, dateIdx + 1, 1, 1);
  newDate.setValues([[currDate]]);
}

function getAndParseEvents(baseUrl) {
  const fullUrl = getFullUrl(baseUrl);
  const returnedValue = JSON.parse(
    UrlFetchApp.fetch(fullUrl, {
      method: 'get',
      followRedirects: false,
      headers: GLOBALS_VARIABLES.headers,
    }).getContentText()
  );

  returnedValue[0].days.forEach((dayObj) => {
    const { events } = dayObj;
    events.forEach(processEvent);
  });
}

function getAndParseMessages() {
  const hasChanged = [];
  const fullUrl = `${GLOBALS_VARIABLES.messagesUrl}?limit=${maxLimit}`;
  const conversationList = JSON.parse(
    UrlFetchApp.fetch(fullUrl, {
      method: 'get',
      followRedirects: false,
      headers: GLOBALS_VARIABLES.headers,
    }).getContentText()
  );

  conversationList.forEach((convo) => {
    const messageId = convo.lastMessage.messageId;
    if (!isLogged(messageId, messageType)) {
      hasChanged.push(convo.conversationId);
    }
  });

  const dateIdx = GLOBALS_VARIABLES.familyIndex.Date;
  const fromId = GLOBALS_VARIABLES.familyIndex.From;
  const chainIdx = GLOBALS_VARIABLES.familyIndex.ChainId;
  const selfId = GLOBALS_VARIABLES.familyIndex.SelfId;
  const lastUpdateIdx = GLOBALS_VARIABLES.familyIndex.LastDate;
  const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
  const contentIdx = GLOBALS_VARIABLES.familyIndex.Content;
  const attachmentIdx = GLOBALS_VARIABLES.familyIndex.Attachments;
  hasChanged.forEach((newConvoId) => {
    if (exceedingTimeLimit()) return;
    const conversationUrl = `${GLOBALS_VARIABLES.messagesUrl}/${newConvoId}`;
    const conversationData = JSON.parse(
      UrlFetchApp.fetch(conversationUrl, {
        method: 'get',
        followRedirects: false,
        headers: GLOBALS_VARIABLES.headers,
      }).getContentText()
    );
    const newMessages = [];
    newMessages[dateIdx] = new Date();
    newMessages[fromId] = conversationData.participants
      .filter((participant) => {
        return (
          !participant.title.includes('Aneesh') &&
          !participant.title.includes('Sherry')
        );
      })
      .map((participant) => {
        const additionalInfo = participant.subtitle
          ? `(${participant.subtitle})`
          : '';
        return `${participant.title} ${additionalInfo}`;
      })
      .join(', ');
    newMessages[chainIdx] = conversationData.conversationId;
    newMessages[lastUpdateIdx] = conversationData.lastActivityAt;
    newMessages[typeIdx] = messageType;
    addInfo(newMessages, conversationData);

    tryCatchTimeout(() => {
      const { newMessageIds, newContent, attachmentUrls } = processMessage(
        conversationData.messages
      );

      newMessages[selfId] = newMessageIds.join(idDelimiter);
      newMessages[contentIdx] = newContent;
      newMessages[attachmentIdx] = attachmentUrls.join(attachDelimiter);
      GLOBALS_VARIABLES.newFamilyData.push(newMessages);
    });
  });
}

function processMessage(messages) {
  const newMessageIds = [];
  let newContent = '';
  const attachmentUrls = [];

  messages.forEach((message) => {
    if (!isLogged(message.messageId, messageType)) {
      newMessageIds.push(message.messageId);
      newContent += `- ${message.body}\n`;
      const newFiles = downloadFiles(message);
      if (newFiles.length) {
        attachmentUrls.push(...newFiles);
      }

      GLOBALS_VARIABLES.familyLoggedEvents[messageType][
        message.messageId
      ] = true;
    }
  });

  return {
    newMessageIds,
    newContent,
    attachmentUrls,
  };
}

// Check if it's an observation
function isObservation(postData) {
  const observationId = postData.embed?.observationId;
  if (observationId) {
    if (
      !GLOBALS_VARIABLES.graphqlQuery.variables.observationIds.includes(
        observationId
      )
    ) {
      GLOBALS_VARIABLES.graphqlQuery.variables.observationIds.push(
        observationId
      );
    }

    return true;
  }

  return false;
}

function getAndParsePosts() {
  const hasChanged = [];
  const fullUrl = `${GLOBALS_VARIABLES.postUrl}?limit=${maxLimit}`;
  const postList = JSON.parse(
    UrlFetchApp.fetch(fullUrl, {
      method: 'get',
      followRedirects: false,
      headers: GLOBALS_VARIABLES.headers,
    }).getContentText()
  );

  const dateIdx = GLOBALS_VARIABLES.familyIndex.Date;
  const fromId = GLOBALS_VARIABLES.familyIndex.From;
  const chainIdx = GLOBALS_VARIABLES.familyIndex.ChainId;
  const selfId = GLOBALS_VARIABLES.familyIndex.SelfId;
  const lastUpdateIdx = GLOBALS_VARIABLES.familyIndex.LastDate;
  const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
  const contentIdx = GLOBALS_VARIABLES.familyIndex.Content;
  const attachmentIdx = GLOBALS_VARIABLES.familyIndex.Attachments;
  postList.forEach((post) => {
    if (exceedingTimeLimit()) return;
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
        newMessage[contentIdx] = post.body;
        addInfo(newMessage, post);
        GLOBALS_VARIABLES.newFamilyData.push(newMessage);
        GLOBALS_VARIABLES.familyLoggedEvents[messageType][
          post.notificationId
        ] = true;
      }
    }
  });

  hasChanged.forEach((newPostId) => {
    if (exceedingTimeLimit()) return;
    const postUrl = `${GLOBALS_VARIABLES.feedItemUrl}?feedItemId=${newPostId}`;
    const postData = JSON.parse(
      UrlFetchApp.fetch(postUrl, {
        method: 'get',
        followRedirects: false,
        headers: GLOBALS_VARIABLES.headers,
      }).getContentText()
    );
    const newMessages = [];

    if (isObservation(postData.feedItem)) {
      return;
    }

    newMessages[dateIdx] = new Date();
    newMessages[fromId] = getFrom(postData.feedItem);
    newMessages[chainIdx] = postData.feedItem.originatorId;
    newMessages[selfId] = postData.feedItem.feedItemId;
    newMessages[lastUpdateIdx] = postData.feedItem.createdDate;
    newMessages[typeIdx] = postType;
    addInfo(newMessages, postData);

    newMessages[contentIdx] = postData.feedItem.body;

    // Handle vacations
    if (postData.feedItem.embed && postData.feedItem.embed.vacationId) {
      const toPeriod = postData.feedItem.embed.period.from_localdate === postData.feedItem.embed.period.to_localdate ? '' : ` - ${postData.feedItem.embed.period.to_localdate}`;
      const period = `${postData.feedItem.embed.period.from_localdate}${toPeriod}`;
      newMessages[contentIdx] += `\n${postData.feedItem.embed.title} (${period})\nResponse needed by: ${postData.feedItem.embed.deadline_localdate}\n${postData.feedItem.embed.vacationId}: ${postData.feedItem.embed.type}`;
    }

    tryCatchTimeout(() => {
      newMessages[attachmentIdx] = downloadFiles(postData.feedItem).join(
        attachDelimiter
      );
      GLOBALS_VARIABLES.newFamilyData.push(newMessages);
      GLOBALS_VARIABLES.familyLoggedEvents[messageType][
        postData.feedItem.feedItemId
      ] = true;
    });
  });
}

// Grab info from posts that are not in the notification center for some reason
function getAndParseBookmarks() {
  const fullUrl = `${GLOBALS_VARIABLES.missedPostUrl}?limit=${maxLimit}`;
  const postList = JSON.parse(
    UrlFetchApp.fetch(fullUrl, {
      method: 'get',
      followRedirects: false,
      headers: GLOBALS_VARIABLES.headers,
    }).getContentText()
  ).feedItems;

  const dateIdx = GLOBALS_VARIABLES.familyIndex.Date;
  const fromId = GLOBALS_VARIABLES.familyIndex.From;
  const chainIdx = GLOBALS_VARIABLES.familyIndex.ChainId;
  const selfId = GLOBALS_VARIABLES.familyIndex.SelfId;
  const lastUpdateIdx = GLOBALS_VARIABLES.familyIndex.LastDate;
  const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
  const contentIdx = GLOBALS_VARIABLES.familyIndex.Content;
  const attachmentIdx = GLOBALS_VARIABLES.familyIndex.Attachments;
  postList.forEach((post) => {
    if (exceedingTimeLimit()) return;
    const messageId = post.feedItemId || post.originatorId;
    if (!isLogged(messageId, postType)) {
      if (isObservation(post)) {
        return;
      }

      const newMessage = [];
      newMessage[dateIdx] = new Date();
      newMessage[fromId] = getFrom(post);
      newMessage[chainIdx] = post.originatorId || post.feedItemId;
      newMessage[selfId] = messageId;
      newMessage[lastUpdateIdx] = post.createdDate;
      newMessage[typeIdx] = postType;
      newMessage[contentIdx] = post.body;
      addInfo(newMessage, post);

      tryCatchTimeout(() => {
        newMessage[attachmentIdx] = downloadFiles(post).join(attachDelimiter);
        GLOBALS_VARIABLES.newFamilyData.push(newMessage);
        GLOBALS_VARIABLES.familyLoggedEvents[messageType][messageId] = true;
      });
    }
  });
}

// Grab info on incident reports
function getAndParseIncidents() {
  const fullUrl = GLOBALS_VARIABLES.incidentUrl;
  const reportList = JSON.parse(
    UrlFetchApp.fetch(fullUrl, {
      method: 'get',
      followRedirects: false,
      headers: GLOBALS_VARIABLES.headers,
    }).getContentText()
  ).reports;

  const dateIdx = GLOBALS_VARIABLES.familyIndex.Date;
  const fromId = GLOBALS_VARIABLES.familyIndex.From;
  const chainIdx = GLOBALS_VARIABLES.familyIndex.ChainId;
  const selfId = GLOBALS_VARIABLES.familyIndex.SelfId;
  const lastUpdateIdx = GLOBALS_VARIABLES.familyIndex.LastDate;
  const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
  const contentIdx = GLOBALS_VARIABLES.familyIndex.Content;
  const attachmentIdx = GLOBALS_VARIABLES.familyIndex.Attachments;
  reportList.forEach((report) => {
    if (exceedingTimeLimit()) return;
    // Add unacknowledged and acknowledged version
    const messageId = `${report.reportId}-${report.acknowledged ? report.acknowledged.name.replace(
      /\s+/g,
      '_'
    ) : 'None'}`;
    if (!isLogged(messageId, incidentType)) {
      const newMessage = [];
      const reporter = getFrom(report);
      newMessage[dateIdx] = new Date();
      newMessage[fromId] = reporter;
      newMessage[chainIdx] = messageId;
      newMessage[selfId] = messageId;
      newMessage[lastUpdateIdx] = report.createdAt;
      newMessage[typeIdx] = incidentType;
      const witnessText = report.witnesses.map(
        (witness) => `${witness.name.fullName} (${witness.employeeId})`
      );
      newMessage[contentIdx] = `${
        report.description
      }\nReported by: ${reporter}\nFrom Arrival? ${
        report.onArrival ? 'Yes' : 'No'
      }\nWitnesses: ${witnessText.join(', ')}`;
      addInfo(newMessage, report);

      tryCatchTimeout(() => {
        newMessage[attachmentIdx] = downloadFiles(report).join(attachDelimiter);
        GLOBALS_VARIABLES.newFamilyData.push(newMessage);
        GLOBALS_VARIABLES.familyLoggedEvents[messageType][messageId] = true;
      });
    }
  });
}

function getAndParseObservations() {
  const observationsData = JSON.parse(
    UrlFetchApp.fetch(GLOBALS_VARIABLES.graphqlUrl, {
      method: 'POST',
      payload: JSON.stringify(GLOBALS_VARIABLES.graphqlQuery),
      headers: {
        ...GLOBALS_VARIABLES.headers,
        'Content-Type': 'application/json',
      },
    }).getContentText()
  );

  const observations =
    observationsData.data.childDevelopment.observations.results;

  const dateIdx = GLOBALS_VARIABLES.familyIndex.Date;
  const fromId = GLOBALS_VARIABLES.familyIndex.From;
  const chainIdx = GLOBALS_VARIABLES.familyIndex.ChainId;
  const selfId = GLOBALS_VARIABLES.familyIndex.SelfId;
  const lastUpdateIdx = GLOBALS_VARIABLES.familyIndex.LastDate;
  const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
  const contentIdx = GLOBALS_VARIABLES.familyIndex.Content;
  const attachmentIdx = GLOBALS_VARIABLES.familyIndex.Attachments;

  observations.forEach((observation) => {
    const messageId = `Observation:${observation.id}`;
    const newMessage = [];
    newMessage[dateIdx] = new Date();
    newMessage[fromId] = observation.createdBy.name.fullName;
    newMessage[chainIdx] = messageId;
    newMessage[selfId] = observation.feedItem.id;
    newMessage[lastUpdateIdx] = observation.status.createdAt;
    newMessage[typeIdx] = postType;

    const areas = observation.remark.areas
      ? `\nAreas: ${observation.remark.areas
          .map((area) => area.area.title)
          .join(', ')}`
      : '';
    const nextSteps = observation.nextStep
      ? `\nNext Steps: ${observation.nextStep.body}`
      : '';

    newMessage[
      contentIdx
    ] = `New Observation:\n${observation.remark.body}${areas}${nextSteps}`;
    addInfo(newMessage, observation);

    tryCatchTimeout(() => {
      newMessage[attachmentIdx] =
        downloadFiles(observation).join(attachDelimiter);
      GLOBALS_VARIABLES.newFamilyData.push(newMessage);
      GLOBALS_VARIABLES.familyLoggedEvents[messageType][messageId] = true;
    });
  });
}

function getFrom(post) {
  if (post.receivers && post.receivers.length) {
    return post.receivers.join(', ');
  }

  if (post.sender) {
    return `${post.sender.name} (${post.sender.id})`;
  }

  if (post.createdBy) {
    if (post.createdBy.name.fullName) {
      return post.createdBy.name.fullName;
    }

    return `${post.createdBy.name} (${post.createdBy.id})`;
  }

  if (post.author) {
    return post.author.subtitle || post.author.title;
  }

  return 'Unknown';
}

function downloadFiles(containerObj) {
  const attachments = [];
  const createDate =
    containerObj.createdDate || containerObj.createdAt || containerObj.status?.createdAt || 'Unknown Date';
  const body = containerObj.body || '';
  const description = `${body}\n\nShared on: ${createDate}\n\nUploaded by ${getFrom(
    containerObj
  )}`;
  if (containerObj.files.length) {
    containerObj.files.forEach((fileObj) => {
      const fileUrl = uploadFile(fileObj.url, fileObj.filename, description);
      attachments.push(fileUrl);
    });
  }

  if (containerObj.images.length) {
    containerObj.images.forEach((imgObj) => {
      let url;
      let fileNameSource;
      if (imgObj.prefix) {
        url = `${imgObj.prefix}/${imgObj.width}x${imgObj.height}/${imgObj.key}`;
        fileNameSource = imgObj.key;
      } else {
        // This is observation url, needs custom url
        fileNameSource = imgObj.secret.path;
        url = `${imgObj.secret.prefix}/${imgObj.secret.key}/${imgObj.width}x${imgObj.height}/${fileNameSource}?expires=${imgObj.secret.expires}`;
      }

      const fileName = fileNameSource.replace(/\//g, '_').replace(/\?.*/, '');
      const fileUrl = uploadFile(url, fileName, description);
      attachments.push(fileUrl);
    });
  }

  // Handle weird naming in observations
  if (containerObj.video && !containerObj.videos) {
    containerObj.videos = [containerObj.video];
  }

  if (containerObj.videos && containerObj.videos.length) {
    containerObj.videos.forEach((videoObj) => {
      const currDate = createDate === 'Unknown Date' ? new Date() : new Date(createDate);
      const videoName = `${parseDate(currDate)}_video_${
        videoObj.videoId || videoObj.id
      }`;
      const fileUrl = uploadFile(
        videoObj.videoUrl,
        videoName,
        description,
        true
      );
      attachments.push(fileUrl);
    });
  }

  if (containerObj.embed && containerObj.embed.invoice) {
    const invoiceObj = containerObj.embed.invoice;
    const invoiceName = `Invoice_${
      invoiceObj.coveringMonths.join('_') || invoiceObj.invoiceNo
    }_amount-${invoiceObj.amount.toString().replace(/\./g, '-')}`;
    const linesText = invoiceObj.lines.map(
      (obj) => `${obj.title}: £${obj.amount}`
    );
    const additionalDescription = `Total: £${
      invoiceObj.amount
    }\n${linesText.join('\n')}\n\nDue: ${invoiceObj.due}\nDate: ${
      invoiceObj.date
    }\n\n`;
    const fileUrl = uploadFile(
      invoiceObj.pdf,
      invoiceName,
      additionalDescription + description,
      true
    );
    attachments.push(fileUrl);
  }

  return attachments;
}

function addInfo(dataArray, info) {
  const infoIdx = GLOBALS_VARIABLES.familyIndex.FamilyInfo;
  const infoJson = JSON.stringify(info);
  dataArray[infoIdx] = infoJson.substr(0, 45000);
  dataArray[infoIdx + 1] = infoJson.substr(45000, 45000);
  dataArray[infoIdx + 2] = infoJson.substr(45000 * 2, 45000);
  dataArray[infoIdx + 3] = infoJson.substr(45000 * 3);
}

function getExistingFile(fileName) {
  return GLOBALS_VARIABLES.googleDriveExistingFiles[fileName];
}

function uploadFile(
  fileUrl,
  fileName,
  additionalDescription,
  keepExtension = false
) {
  if (exceedingTimeLimit()) {
    throw new TimeoutError('Exceeded 5 minutes');
  }

  if (!GLOBALS_VARIABLES.googleDrive) {
    GLOBALS_VARIABLES.googleDriveExistingFiles = {};
    GLOBALS_VARIABLES.googleDriveExistingFilesByUrl = {};
    GLOBALS_VARIABLES.googleDrive = DriveApp.getFolderById(
      GLOBALS_VARIABLES.folderId
    );
    const existingFiles = GLOBALS_VARIABLES.googleDrive.getFiles();
    while (existingFiles.hasNext()) {
      const file = existingFiles.next();
      const curFileName = file.getName();
      const fileUrl = file.getUrl();
      GLOBALS_VARIABLES.googleDriveExistingFiles[curFileName] = fileUrl;
      GLOBALS_VARIABLES.googleDriveExistingFilesByUrl[fileUrl] = curFileName;
    }
  }

  let existingFileUrl = getExistingFile(fileName);
  if (existingFileUrl) {
    return existingFileUrl;
  }

  const response = UrlFetchApp.fetch(fileUrl);
  const blob = response.getBlob();
  const file = GLOBALS_VARIABLES.googleDrive.createFile(blob);

  if (keepExtension) {
    const extension = file.getName().match(/\..*?$/)[0];
    if (extension) {
      fileName += extension;
    }
  }

  file.setName(fileName);
  file.setDescription(
    `Download from ${fileUrl} on ${new Date()}\n\n${additionalDescription}`
  );

  // If the file is duplicate of one with extension, delete it
  existingFileUrl = getExistingFile(fileName);
  if (existingFileUrl) {
    file.setTrashed(true);
    return existingFileUrl;
  }

  const driveLink = file.getUrl();

  GLOBALS_VARIABLES.googleDriveExistingFiles[fileName] = driveLink;
  GLOBALS_VARIABLES.googleDriveExistingFilesByUrl[driveLink] = fileName;
  return driveLink;
}

const amountToDescription = {
  1: 'Little',
  2: 'Half',
  3: 'Most',
  4: 'All',
  5: 'All+',
};

function isLogged(messageId, type) {
  return messageId in GLOBALS_VARIABLES.familyLoggedEvents[type];
}

// Times out after 6 minutes (360s)
// Mark as true after 5 minutes
function exceedingTimeLimit() {
  return (new Date() - GLOBALS_VARIABLES.startTime) / 1000 > 300;
}

class TimeoutError extends Error {
  constructor(message) {
    super(message);
    this.name = 'TimeOutError';
  }
}

function tryCatchTimeout(cb) {
  try {
    cb();
  } catch (error) {
    if (error instanceof TimeoutError) {
      console.log('Timed Out While Processing');
    } else {
      throw error;
    }
  }
}

// Process all logged events
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
    newDataRow[totalTime] =
      (new Date(event.to) - new Date(event.from)) / 60 / 1000;
  } else {
    newDataRow[totalTime] = undefined;
  }

  if (event.embed && event.embed.mealItems && event.embed.mealItems.length) {
    newDataRow[noteIdx] += ` (${event.embed.mealItems
      .map(
        (item) =>
          `${item.foodItem.title} - ${
            amountToDescription[item.amount] || item.amount
          }`
      )
      .join(', ')})`;
  }

  const eventIdentifier = getIdentifier(newDataRow);
  if (!GLOBALS_VARIABLES.loggedEvents.has(eventIdentifier)) {
    GLOBALS_VARIABLES.newData.push(newDataRow);
    GLOBALS_VARIABLES.loggedEvents.add(eventIdentifier);
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

function appendRows(sheet, newData, attachmentIdx) {
  const startRow = sheet.getLastRow() + 1;
  const startCol = 1;
  const numRows = newData.length;
  const numCols = newData[0].length;

  let attachments = [];
  if (attachmentIdx !== undefined) {
    newData.forEach((data) => {
      const attachmentLink = data[attachmentIdx];
      attachments.push(attachmentLink);
      data[attachmentIdx] = '';
    });
  }

  const newRange = sheet.getRange(startRow, startCol, numRows, numCols);
  newRange.setValues(newData);

  if (attachments.length) {
    const newRange = sheet.getRange(startRow, attachmentIdx + 1, numRows, 3);
    const richTextAttachments = attachments.map((attachment) => {
      const newRichText = {
        0: SpreadsheetApp.newRichTextValue(),
        1: SpreadsheetApp.newRichTextValue(),
        2: SpreadsheetApp.newRichTextValue(),
      };

      const newText = {
        0: '',
        1: '',
        2: '',
      };
      const links = [];

      if (attachment) {
        const allAttachments = attachment.split(attachDelimiter);
        let numImg = 0;
        allAttachments.forEach((url) => {
          const version = parseInt(numImg / 40);
          const start = newText[version].length;

          let fileType = (
            GLOBALS_VARIABLES.googleDriveExistingFilesByUrl[url] || ''
          ).match(/\.(.*?)$/);
          fileType = fileType ? fileType[1] : 'image';
          newText[version] += `${fileType}${numImg}, `;
          numImg += 1;

          const end = newText[version].length - 4;

          links.push({
            start,
            end,
            url,
            version,
          });
        });
      }

      newRichText[0].setText(newText[0]);
      newRichText[1].setText(newText[1]);
      newRichText[2].setText(newText[2]);
      links.forEach(({ start, end, url, version }) => {
        newRichText[version].setLinkUrl(start, end, url);
      });

      return [
        newRichText[0].build(),
        newRichText[1].build(),
        newRichText[2].build(),
      ];
    });

    newRange.setRichTextValues(richTextAttachments);
  }
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
  const result = [];
  let arrayLength = 1;
  restArguments.forEach((arg, idx) => {
    if ((idx % 2 === 1) & (arg instanceof Array)) {
      arrayLength = arg.length;
    }
  });

  for (let arrayIdx = 0; arrayIdx < arrayLength; arrayIdx++) {
    const filteredItems = range.filter((_, idx) => {
      for (
        let argPairIndex = 0;
        argPairIndex < restArguments.length / 2;
        argPairIndex++
      ) {
        const argIndex = argPairIndex * 2;
        const compareItem = restArguments[argIndex][idx];
        let staticItem = restArguments[argIndex + 1];
        staticItem =
          staticItem instanceof Array ? staticItem[arrayIdx][0] : staticItem;
        if (compareItem.toString() !== staticItem.toString()) return false;
      }

      return true;
    });
    const filteredText = filteredItems.length
      ? `${prependText}${filteredItems.join(joinText)}`
      : '';
    result.push(filteredText);
  }

  return result;
}

// CUSTOM FUNCTION HELPERS

// Faster isNan
function myIsNaN(val) {
  return !(val <= 0) && !(val > 0);
}

function convertFromPacific(date, latitude, longitude) {
  return changeTimezone(
    date,
    'America/Los_Angeles',
    tzlookup(latitude, longitude)
  );
}

function changeTimezone(timeZone) {
  if (timeZone == 'Europe/London') {
    return 'Etc/GMT'
  } else if (timeZone.startsWith('Europe/')) {
    return 'Etc/GMT+1'
  } else {
    return 'Etc/GMT';
  }
}

// converting from time zone:https://stackoverflow.com/a/53652131
function changeTimezone(date, oldTimezone, newTimezone) {
  if (!date.getTime) return date;

  let oldDate = new Date(
    new Date(Utilities.formatDate(date, oldTimezone, 'YYYY-MM-dd hh:mm a'))
  );

  let newDate = new Date(
    new Date(Utilities.formatDate(date, newTimezone, 'YYYY-MM-dd hh:mm a'))
  );

  if (isNaN(newDate.valueOf()) || isNaN(oldDate.valueOf())) {
    console.error(`ERROR: newDate ${newDate} (${newTimezone}); oldDate ${oldDate} (${oldTimezone})`);
    // Hopefully not necessary now that I've changed to Utlities.formatDate
    // if (isNaN(newDate.valueOf())) {
    //   newDate = new Date(Utilities.formatDate(date, changeTimezone(newTimezone), 'YYYY-MM-dd hh:mm a'));
    // }

    // if (isNaN(oldDate.valueOf())) {
    //   oldDate = new Date(Utilities.formatDate(date, changeTimezone(oldTimezone), 'YYYY-MM-dd hh:mm a'));
    // }
  }

  const diff = newDate.getTime() - oldDate.getTime();
  return new Date(date.getTime() + diff);
}
