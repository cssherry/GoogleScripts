// Needs GLOBALS_VARIABLES and indexSheet function
// GLOBALS_VARIABLES = {
//   headersNest: { cookies: '_brightwheel_v2=XXX' },
//   nestStudents: [''],
//   nestGuardian: '',
//   nestBaseUrl: '',
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
//   incidentUrl: [''],
//   graphqlUrl: '',
//   graphqlQuery: { query, operationName, variables: {observationIds:[]}, },
// };

const lineSeparators = '\n\n-----------------------\n\n';
const maxLimit = 100;
const idDelimiter = ',';
const attachDelimiter = ' ; ';
const cannotUploadText = 'FILE_IS_TOO_LARGE';
const cannotUploadSeparator = ': ';
const messageType = 'Message';
const postType = 'Post';
const incidentType = 'Incident';
const emailType = 'Email';
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
  GLOBALS_VARIABLES.data.forEach(row => {
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
  GLOBALS_VARIABLES.familyLoggedEvents[emailType] = {};

  // Process existing posts/messages
  const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
  const idIdx = GLOBALS_VARIABLES.familyIndex.SelfId;
  GLOBALS_VARIABLES.familyData.forEach((row, idx) => {
    if (idx === 0) return;
    if (!row[0]) return;
    const type = row[typeIdx];
    const ids = row[idIdx].split(idDelimiter);
    ids.forEach(id => {
      GLOBALS_VARIABLES.familyLoggedEvents[type][id] = idx;
      GLOBALS_VARIABLES.familyLoggedEvents[type][id] = idx;
    });
  });

  GLOBALS_VARIABLES.napStart = [];
  GLOBALS_VARIABLES.napEnd = [];

  // ===========================
  // Get any Nest messages
  console.log('Getting Nest Messages')
  getAndParseNestMessages();

  // Get nest events
  console.log('Getting Nest Activities')
  getAndParseActivities();

  // Add nest naps separately
  console.log('Adding nap total times')
  parseNestNaps();

  // ===========================
  // Get Goddard and Nest emails
  console.log('Adding emails')
  getAndParseEmails();

  // ===========================
  // Check for new messages
  console.log('Adding Famly Messages')
  getAndParseMessages();

  // Check for new posts
  console.log('Adding Famly Posts')
  getAndParsePosts();

  // Add bookmarks
  console.log('Adding Famly Bookmarks (hack for getting all photos)')
  getAndParseBookmarks();

  // Add any incident reports
  console.log('Adding Famly Incidents')
  getAndParseIncidents();

  // Add any incident reports
  console.log('Adding Famly Observations')
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
    .map(row => {
      famlySummary += `${row[GLOBALS_VARIABLES.index.Note]} (${
        row[GLOBALS_VARIABLES.index.Date]
      })${lineSeparators}`;
      return row
        .filter(item => !!item)
        .map(item => {
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
    .map(row => {
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
    body: famlySummary + loggedData + separator + daycareGeneral + separator + GLOBALS_VARIABLES.famlyUrls,
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

function getAndParseNestMessages() {
  // Step 1: Get threads
  const fullUrl = `${GLOBALS_VARIABLES.nestBaseUrl}/v2/guardians/${GLOBALS_VARIABLES.nestGuardian}/message_threads`;
    const returnedValue = JSON.parse(UrlFetchApp.fetch(fullUrl, {
        method: 'get',
        followRedirects: false,
        headers: GLOBALS_VARIABLES.headersNest,
      }).getContentText());
    console.log(`Processing ${returnedValue.results.length} threads`);
    returnedValue.results.forEach(getMessagesInThread);
}


function getMessagesInThread(messageThread) {
  // Step 2: Go through each message on thread and create new event
  const fullUrl = `${GLOBALS_VARIABLES.nestBaseUrl}/v2/guardians/${GLOBALS_VARIABLES.nestGuardian}/message_threads/${messageThread.object_id}/messages?page_limit=50`;
  const returnedValue = JSON.parse(UrlFetchApp.fetch(fullUrl, {
      method: 'get',
      followRedirects: false,
      headers: GLOBALS_VARIABLES.headersNest,
    }).getContentText());
  console.log(`Looking in ${returnedValue.results.length} messages`);
  returnedValue.results.forEach(message => parseMessage(message, messageThread.student.first_name));
}

function parseMessage({message}, student) {
  // We assume there's only new messages -- no continued threads
  const objectId = message.message_content_id
  if (isLogged(objectId, messageType)) return;

  const dateIdx = GLOBALS_VARIABLES.familyIndex.Date;
  const fromIdx = GLOBALS_VARIABLES.familyIndex.From;
  const chainIdx = GLOBALS_VARIABLES.familyIndex.ChainId;
  const selfId = GLOBALS_VARIABLES.familyIndex.SelfId;
  const lastUpdateIdx = GLOBALS_VARIABLES.familyIndex.LastDate;
  const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
  const contentIdx = GLOBALS_VARIABLES.familyIndex.Content;
  const attachmentIdx = GLOBALS_VARIABLES.familyIndex.Attachments;
  const currMessage = [];
  tryCatchTimeout(() => {
    currMessage[dateIdx] = new Date(message.created_at);
    currMessage[lastUpdateIdx] = message.created_at;
    currMessage[chainIdx] = message.object_id; // should be the same
    currMessage[selfId] = objectId;
    currMessage[typeIdx] = messageType;
    const sender = `${message.sender.first_name} ${message.sender.last_name} - ${message.sender.user_type}`;
    currMessage[contentIdx] = `${student}:\n${message.body}\nSent by: ${sender}`;


    const attachments = message.attachments.map(currAttach => {
      return uploadFile(
        currAttach.url,
        currAttach.name,
        `${currMessage[contentIdx]}\nOn: ${currAttach.created_at.toString()}`,
        true,
      );
    });

    currMessage[attachmentIdx] = attachments.join(attachDelimiter);
    currMessage[fromIdx] = sender;
    addInfo(currMessage, message);
    GLOBALS_VARIABLES.familyLoggedEvents[messageType][objectId] = true;
    GLOBALS_VARIABLES.newFamilyData.push(currMessage);
  });
}

function getAndParseActivities() {
  GLOBALS_VARIABLES.nestStudents.forEach(studentId => {
    const endDate = new Date();
    endDate.setDate(endDate.getDate() + 1);
    const fullUrl = `${GLOBALS_VARIABLES.nestBaseUrl}/v1/students/${studentId}/activities?page=0&page_size=100&start_date=${GLOBALS_VARIABLES.startDate}&end_date=${endDate.toISOString().split('T')[0]}&include_parent_actions=true`;
    const returnedValue = JSON.parse(UrlFetchApp.fetch(fullUrl, {
        method: 'get',
        followRedirects: false,
        headers: GLOBALS_VARIABLES.headersNest,
      }).getContentText());
    console.log(`Parsing ${returnedValue.activities.length} activities`);
    returnedValue.activities.forEach(processNestActivity);
  });
}

// Process all logged events -- this is a mix of logged events (pee / poo / food / nap) + uploaded attachments with messages
function processNestActivity(activity) {
  const objectId = activity.object_id;
  const needsDetailedLog = !!activity.media || !!activity.video_info;

  if (needsDetailedLog && !isLogged(objectId, postType)) {
    const newDataRow = parseAsNestMedia(activity, objectId);
    GLOBALS_VARIABLES.familyLoggedEvents[postType][objectId] = GLOBALS_VARIABLES.newFamilyData.length;
    GLOBALS_VARIABLES.newFamilyData.push(newDataRow);
  } else if (!needsDetailedLog) {
    parseAsNestEvent(activity);
  }
}

function parseAsNestMedia(event, objectId) {
  const dateIdx = GLOBALS_VARIABLES.familyIndex.Date;
  const fromIdx = GLOBALS_VARIABLES.familyIndex.From;
  const chainIdx = GLOBALS_VARIABLES.familyIndex.ChainId;
  const selfId = GLOBALS_VARIABLES.familyIndex.SelfId;
  const lastUpdateIdx = GLOBALS_VARIABLES.familyIndex.LastDate;
  const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
  const contentIdx = GLOBALS_VARIABLES.familyIndex.Content;
  const attachmentIdx = GLOBALS_VARIABLES.familyIndex.Attachments;
  const message = [];
  tryCatchTimeout(() => {
    message[dateIdx] = new Date(event.event_date);
    message[lastUpdateIdx] = event.event_date;
    message[chainIdx] = objectId;
    message[selfId] = objectId;
    message[typeIdx] = postType;
    message[contentIdx] = `${getNestChild(event)} ${event.note || ''} (${getNestActor(event)})`;

    let attachment
    let fileName = `${message[dateIdx].getYear()}_${message[dateIdx].getMonth()}_${message[dateIdx].getDate()}_${getNestChild(event)}_`;
    if (event.media) {
      fileName += `${event.media.object_id}_${objectId}`;
      attachment = uploadFile(
        event.media.image_url,
        fileName,
        `${message[contentIdx]}\nOn: ${message[dateIdx].toString()}`,
        true,
      );
    } else {
      fileName += `${event.video_info.object_id}_${objectId}`;
      attachment = uploadFile(
        event.video_info.downloadable_url,
        fileName,
        `${message[contentIdx]}\nOn: ${message[dateIdx].toString()}`,
        true,
      );
    }


    message[attachmentIdx] = attachment;
    message[fromIdx] = `${getNestChild(event)}: ${event.room.name}`;
    addInfo(message, event);
  });

  return message;
}

function parseAsNestEvent(event) {
  const dateIdx = GLOBALS_VARIABLES.index.Date;
  const actionIdx = GLOBALS_VARIABLES.index.Action;
  const noteIdx = GLOBALS_VARIABLES.index.Note;
  const infoIdx = GLOBALS_VARIABLES.index.FamilyInfo;

  const newDataRow = [];
  newDataRow[dateIdx] = new Date(event.event_date);
  newDataRow[actionIdx] = event.action_type;
  let eventTitle;
  const eventDetails = event.details_blob
  if (event.action_type === 'ac_potty') {
    eventTitle = `Potty ${eventDetails.potty} (${eventDetails.potty_type} - ${eventDetails.potty_extras.join(' , ')})`;
  } else if (event.action_type === 'ac_food') {
    const amount = event.details_blob.amount;
    eventTitle = `Meal type: ${event.details_blob.food_meal_type} (Amount - ${amount ? 'Most' : 'All'} - ${amount})`
    const foodTag = event.menu_item_tags.map(item => item.name).join(', ');
    if (foodTag) {
      eventTitle += `\n${foodTag}`
    }
  } else if (event.action_type === 'ac_nap') {
    eventTitle = event.state;

    if (parseInt(event.state)) {
      GLOBALS_VARIABLES.napStart.push(newDataRow);
    } else {
      GLOBALS_VARIABLES.napEnd.push(newDataRow);
    }

  } else {
    eventTitle = Object.entries(eventDetails).filter((_, value) => typeof value !== 'string').map((key, value) => `${key} - ${value}`).join('; ');
  }

  if (event.note) {
    eventTitle += `\n${event.note}`;
  }

  newDataRow[noteIdx] = `${getNestChild(event)}: ${eventTitle} (${getNestActor(event)})`;
  newDataRow[infoIdx] = JSON.stringify(event);

  const identifier = getIdentifier(newDataRow);
  if (GLOBALS_VARIABLES.loggedEvents.has(identifier)) return;
  GLOBALS_VARIABLES.newData.push(newDataRow);
  GLOBALS_VARIABLES.loggedEvents.add(identifier);
}

function getNestActor(event) {
  return `${event.actor.first_name} ${event.actor.last_name} - ${event.actor.email}`
}

function getNestChild(event) {
  return event.target.first_name;
}

// Really annoying, but Nest naps are all randomly dumped into activity log. I need to parse start / end dates and add those as actual full nap times :(
function parseNestNaps() {
  const dateIdx = GLOBALS_VARIABLES.index.Date;
  const noteIdx = GLOBALS_VARIABLES.index.Note;
  const totalTime = GLOBALS_VARIABLES.index.TotalTime;
  function sortNaps(a, b) {
    return a[dateIdx] - b[dateIdx];
  }

  // Sort so start nap and end nap are theoretically together / same idx
  GLOBALS_VARIABLES.napStart.sort(sortNaps);
  GLOBALS_VARIABLES.napEnd.sort(sortNaps);

  GLOBALS_VARIABLES.napStart.forEach((startNap, idx) => {
    const newDataRow = [...startNap];

    newDataRow[noteIdx] = `Total Nap: ${newDataRow[noteIdx]} (${startNap[dateIdx].toISOString()} - ${GLOBALS_VARIABLES.napEnd[idx][dateIdx].toISOString()})`;
    newDataRow[totalTime] = ((GLOBALS_VARIABLES.napEnd[idx][dateIdx] - startNap[dateIdx]) / 1000 / 60).toFixed(2);
    const identifier = getIdentifier(newDataRow);
    if (!GLOBALS_VARIABLES.loggedEvents.has(identifier)) {
      GLOBALS_VARIABLES.newData.push(newDataRow);
      GLOBALS_VARIABLES.loggedEvents.add(identifier);
    };
  });
}

function getAndParseEmails() {
  const startDate = getStartDate();
  tryCatchTimeout(() => {
    // Nest
    GmailApp.search(`from:nestgreenfieldhill@gmail.com after:${startDate}`).forEach(parseEmail);
    // Goddard
    GmailApp.search(`from:kaymbu.com after:${startDate}`).forEach(parseEmail);
    GmailApp.search(`from:goddard after:${startDate}`).forEach(parseEmail);
    GmailApp.search(`from:goddardschools after:${startDate}`).forEach(parseEmail);
  });
}

function parseEmail(gmailThread) {
  const messages = gmailThread.getMessages();
  const messageId = gmailThread.getId();
  const lastDate = gmailThread.getLastMessageDate();
  let message;

  const dateIdx = GLOBALS_VARIABLES.familyIndex.Date;
  const fromIdx = GLOBALS_VARIABLES.familyIndex.From;
  const chainIdx = GLOBALS_VARIABLES.familyIndex.ChainId;
  const selfId = GLOBALS_VARIABLES.familyIndex.SelfId;
  const lastUpdateIdx = GLOBALS_VARIABLES.familyIndex.LastDate;
  const typeIdx = GLOBALS_VARIABLES.familyIndex.Type;
  const contentIdx = GLOBALS_VARIABLES.familyIndex.Content;
  const attachmentIdx = GLOBALS_VARIABLES.familyIndex.Attachments;
  const subject = gmailThread.getFirstMessageSubject();

  if (isLogged(messageId, emailType)) {
    const dataIdx = GLOBALS_VARIABLES.familyLoggedEvents[emailType][messageId];
    message = GLOBALS_VARIABLES.familyData[dataIdx];
    if (message[lastUpdateIdx].toISOString() === lastDate.toISOString()) return;

    // Remove existing date
    message.pop();
  } else {
    message = [];
    message[dateIdx] = new Date();
    message[selfId] = messageId;
    message[contentIdx] = `${subject}\n`;
    message[typeIdx] = emailType;
  };

  const attachments = message[attachmentIdx] ? message[attachmentIdx].split(attachDelimiter) : [];
  const messageIds = message[chainIdx] ? message[chainIdx].split(attachDelimiter): [];
  const fromIds = message[fromIdx] ? message[fromIdx].split(attachDelimiter): [];
  let fullMessage = '';
  let fullContent = '';
  const subjectForImages = subject.matchAll(/\b[a-zA-Z0-9]+\b/g).reduce((all, curr) => all ? `${all}_${curr}` : curr, '');
  messages.forEach((message, idx) => {
    const messageId = message.getId();
    if (!messageIds.includes(messageId)) {
      const body = message.getPlainBody();
      fullMessage += `\n\nRAW CONTENTS ${idx}\n`;
      fullMessage += message.getRawContent();
      fullMessage += `\n\nPLAIN CONTENTS ${idx}\n`;
      fullMessage += body;

      // Trims out newlines in Nest emails
      // console.log(0.1, body);
      // const cleanedBody = body.replace(/\n{2,}/g, '\n').replace(/\s{2,}/g, ' ');
      // console.log(0.2, cleanedBody);
      // fullContent += cleanedBody;
      fullContent += body
      const createdDate = message.getDate();
      const dateForFiles = createdDate.toISOString().replaceAll(/[:\.]/g, '_');
      const author = message.getFrom();

      // const downloadAll = body.matchAll(/Download All[\s\n]*<(http.+?)>/g);
      // downloadAll.forEach((captureGroup, idx) => {
      //   attachments.push(uploadFile(
      //     captureGroup[1],
      //     `${dateForFiles}_${subjectForImages}_AllAssets_${idx + 1}`,
      //     `${body}]nFrom: ${author}\nOn: ${createdDate}`,
      //     true,
      //   ));
      // });

      let currText = '';
      let skipNextUrl = false;
      let nextIsLink = false;
      let passedLinks = false;
      let imageIdx = 1;
      // Go through each newline
      body.trim().split(/\s+Do more with the Goddard Family Hub app\s+/i)[0].split(/\s*\n+\s*/).forEach(text => {
        // If it is the download all option -- we don't need to do this anymore
        // Then handle when a link is present -- use the link after "Download this moment" line
        // Otherwise, append text to currText for description
        if (text.match(/Download\s+All/)) {
          skipNextUrl = true;
        } else if (text.includes('Download this moment')) {
          nextIsLink = true;
        } else if (text.startsWith('<http')) {
          // If previous was download all, let's ignore
          // Otherwise, let's add this to the files for download if there's a "Download" option
          passedLinks = true;
          if (skipNextUrl) {
            currText = '';
            skipNextUrl = false;
          } else if (nextIsLink) {
            attachments.push(
              uploadFile(
                text.match(/<(http.+)>/)[1],
                `${dateForFiles}_${subjectForImages}_${imageIdx}`,
                `From: ${author}\nOn: ${createdDate}\n${currText}`,
                true,
              )
            );
            imageIdx += 1;
            currText = '';
            nextIsLink = false;
          }
        } else if (passedLinks) {
          passedLinks = false;
          if (currText) {
            currText += `\n${text}`;
          } else {
            currText = text;
          }
        } else {
          currText += `${text} `;
        }
      });

      message.getAttachments().forEach(currAttach => {
        const currName = currAttach.getName();
        attachments.push(uploadFile(
          null,
          currName,
          `${body}]nFrom: ${author}\nOn: ${createdDate}`,
          true,
          currAttach.copyBlob(),
        ));
      });

      messageIds.push(messageId);
      fromIds.push(author);
    }
  });

  tryCatchTimeout(() => {
    message[lastUpdateIdx] = lastDate;
    message[chainIdx] = messageIds.join(attachDelimiter);
    message[attachmentIdx] = attachments.join(attachDelimiter);
    message[fromIdx] = fromIds.join(attachDelimiter);
    message[contentIdx] += fullContent;

    const currContent = message[contentIdx];
    if (currContent.length > 50000) {
      message[contentIdx] = currContent.substr(0, 49000);
    }
    appendInfoString(message, fullMessage);
    GLOBALS_VARIABLES.familyLoggedEvents[messageType][messageId] = GLOBALS_VARIABLES.newFamilyData.length;
    GLOBALS_VARIABLES.newFamilyData.push(message);
  });
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

  conversationList.forEach(convo => {
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
      .filter(participant => {
        return (
          !participant.title.includes('Aneesh') &&
          !participant.title.includes('Sherry')
        );
      })
      .map(participant => {
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

  messages.forEach(message => {
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
  postList.forEach(post => {
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

  hasChanged.forEach(newPostId => {
    if (exceedingTimeLimit()) return;
    const postUrl = `${GLOBALS_VARIABLES.feedItemUrl}?feedItemId=${newPostId}`;
    let postData;
    try {
      const result = UrlFetchApp.fetch(postUrl, {
        method: 'get',
        followRedirects: false,
        headers: GLOBALS_VARIABLES.headers,
      }).getContentText();
      postData = JSON.parse(result);
    } catch (error) {
      postData = {
        feedItem: {
          sender: {
            name: 'ERROR',
            id: '',
          },
          files: [],
          images: [],
          originatorId: newPostId,
          feedItemId: newPostId,
          createdDate: new Date(),
          body: `Error: ${error.message}\nURL: ${error.url || postUrl}`,
        },
      }
    }

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
  postList.forEach(post => {
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
  GLOBALS_VARIABLES.incidentUrl.forEach(getAndParseIncidentForChild);
}

function getAndParseIncidentForChild(fullUrl, _idx) {
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
  reportList.forEach(report => {
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
        witness => `${witness.name.fullName} (${witness.employeeId})`
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

  observations.forEach(observation => {
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
          .map(area => area.area.title)
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
    containerObj.files.forEach(fileObj => {
      const fileUrl = uploadFile(fileObj.url, fileObj.filename, description);
      attachments.push(fileUrl);
    });
  }

  if (containerObj.images.length) {
    containerObj.images.forEach(imgObj => {
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
    containerObj.videos.forEach(videoObj => {
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
      obj => `${obj.title}: £${obj.amount}`
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
  const infoJson = JSON.stringify(info);
  addInfoString(dataArray, infoJson);
}

function appendInfoString(dataArray, infoString) {
  const infoIdx = GLOBALS_VARIABLES.familyIndex.FamilyInfo;
  const previousInfo = (dataArray[infoIdx] || '') + (dataArray[infoIdx + 1] || '') + (dataArray[infoIdx + 2] || '') + (dataArray[infoIdx + 3] || '') + (dataArray[infoIdx + 4] || '') + (dataArray[infoIdx + 5] || '');
  addInfoString(dataArray, previousInfo + infoString)
}

function addInfoString(dataArray, infoJson) {
  const infoIdx = GLOBALS_VARIABLES.familyIndex.FamilyInfo;
  dataArray[infoIdx] = infoJson.substr(0, 49000);
  dataArray[infoIdx + 1] = infoJson.substr(49000, 49000);
  dataArray[infoIdx + 2] = infoJson.substr(49000 * 2, 49000);
  dataArray[infoIdx + 3] = infoJson.substr(49000 * 3, 49000);
  dataArray[infoIdx + 4] = infoJson.substr(49000 * 4, 49000);
  dataArray[infoIdx + 5] = infoJson.substr(49000 * 5, 50000);
}

function getExistingFile(fileName) {
  return GLOBALS_VARIABLES.googleDriveExistingFiles[fileName];
}

function uploadFile(
  fileUrl,
  fileName,
  additionalDescription,
  keepExtension = false,
  blob = null, // if include blob, no longer need fileUrl
) {
  console.log('Starting to upload file');
  if (exceedingTimeLimit()) {
    throw new TimeoutError('Exceeded 5 minutes');
  }

  if (!GLOBALS_VARIABLES.googleDrive) {
    console.log('Fetching Google Drive');
    GLOBALS_VARIABLES.googleDriveExistingFiles = {};
    GLOBALS_VARIABLES.googleDriveExistingFilesByUrl = {};
    GLOBALS_VARIABLES.googleDrive = DriveApp.getFolderById(
      GLOBALS_VARIABLES.folderId
    );
    const existingFiles = GLOBALS_VARIABLES.googleDrive.searchFiles('modifiedDate > "2025-08-01"');
    while (existingFiles.hasNext()) {
      const file = existingFiles.next();
      const curFileName = file.getName();
      const fileUrl = file.getUrl();
      const curFileNameWithoutExt = curFileName.replace(/\.[a-zA-Z]*?$/, '')
      GLOBALS_VARIABLES.googleDriveExistingFiles[curFileNameWithoutExt] = fileUrl;
      GLOBALS_VARIABLES.googleDriveExistingFilesByUrl[fileUrl] = curFileNameWithoutExt;
      GLOBALS_VARIABLES.googleDriveExistingFiles[curFileName] = fileUrl;
      GLOBALS_VARIABLES.googleDriveExistingFilesByUrl[fileUrl] = curFileName;
    }
  }

  if (!blob) {
    console.log('Checking Filename');
    let existingFileUrl = getExistingFile(fileName);
    if (existingFileUrl) {
      return existingFileUrl;
    }

    try {
      console.log("File doesn't exist, fetching");
      const response = UrlFetchApp.fetch(fileUrl);
      blob = response.getBlob();
    } catch (e) {
      if (e.message.includes("URLFetch URL Length") || e.message.includes("Request-URI Too Large")) {
        return `URLFetchLong_${cannotUploadText}${cannotUploadSeparator}${fileName}${cannotUploadSeparator}${fileUrl || fileName}`;
      } else {
        // Handle other types of errors
        console.error(`An unexpected error occurred: ${e.message}`);
      }
    }

  }

  const blobSize = blob.getBytes().length / 1000000;
  console.log(`Blob size ${blobSize} (needs to be under 50 to get uploaded)`);
  if (blobSize >= 50) return `${cannotUploadText}${cannotUploadSeparator}${fileName}${cannotUploadSeparator}${fileUrl || fileName}`;
  console.log('Creating blob');
  const file = GLOBALS_VARIABLES.googleDrive.createFile(blob);

  if (keepExtension) {
    const extension = file.getName().match(/\..*?$/)[0];
    console.log(`Adding ext ${extension}`);
    if (extension) {
      fileName += extension;
    }
  }

  console.log(`Setting filename ${fileName}`);
  file.setName(fileName);
  file.setDescription(
    `Download from ${fileUrl} on ${new Date()}\n\n${additionalDescription}`
  );

  // If the file is duplicate of one with extension, delete it
  existingFileUrl = getExistingFile(fileName);
  if (existingFileUrl) {
    console.log(`File already exists with extension, delete ${fileName}`);
    file.setTrashed(true);
    return existingFileUrl;
  }

  const driveLink = file.getUrl();
  console.log(`New file created ${driveLink}`);
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

  if (event.embed) {
    let additionalContent = '';
    if (event.embed.mealItems && event.embed.mealItems.length) {
      additionalContent = ` (${event.embed.mealItems
        .map(
          (item) =>
            `${item.foodItem.title} - ${
              amountToDescription[item.amount] || item.amount
            }`
        )
        .join(', ')})`;
    } else if (event.embed.actionType === 'TOILETVISIT' && event.embed.diaperingType) {
      additionalContent = event.embed.diaperingType;
    }

    newDataRow[noteIdx] += additionalContent;
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
    newData.forEach(data => {
      const attachmentLink = data[attachmentIdx];
      attachments.push(attachmentLink);
      data[attachmentIdx] = '';
    });
  }

  const newRange = sheet.getRange(startRow, startCol, numRows, numCols);
  newRange.setValues(newData);

  if (attachments.length) {
    const newRange = sheet.getRange(startRow, attachmentIdx + 1, numRows, 3);
    const richTextAttachments = attachments.map(attachment => {
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
        allAttachments.forEach(url => {
          const version = parseInt(numImg / 40);
          const start = newText[version].length;
          let fileType;
          if (url.includes(cannotUploadText)) {
            const index = url.lastIndexOf(cannotUploadSeparator);
            fileType = url.slice(0, index);
            url = url.slice(index + cannotUploadSeparator.length);
          } else {
            fileType = (
              GLOBALS_VARIABLES.googleDriveExistingFilesByUrl[url] || ''
            ).match(/\.(.*?)$/);
            fileType = fileType ? fileType[1] : 'image';
          }

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
/**
 * Converts local date/time to ET
 * @param {date} sheetDate The local date
 * @param {number} latitude
 * @param {number} longitude
 * @return {date} The value in millimeters.
 * @customfunction
 */
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
  try {
    return changeTimezone(
      date,
      'America/Los_Angeles',
      tzlookup(latitude, longitude)
    );
  } catch (error) {
    console.log(date, latitude, longitude, error);
    return date
  }
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
