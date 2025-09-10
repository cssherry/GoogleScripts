let writingCalendar;
const lengthOfResponse = 300;
const emailPrefix = '[CreativeWriting] '

// https://support.google.com/calendar/forum/AAAAd3GaXpEoHbbc3DuDt0/?hl=en
var charLimit = 8148;
var graceLimit = 500; // we allow for average length + gracelimit before dropping 1 day's writing

var warningDay = 3;
var moveDay = 7;

var lengthEvent = 3; // Hours writing event should last

const prefixCharLength = 20;
var divider = '===================';
var noteDivider = '\n' + divider + '\n';
var summaryHeader = 'Final: ';
var summaryPartRegex = new RegExp('^' + summaryHeader + '\\d+(\\.(\\d+)):');

// ==========================================
// Runs at 5 AM Time
// 1. If there isn't an all day event for correct prompt today, move today's event to next day
// 2. If it has been "moveDay" number days since prompt was updated, switch to different person
// 2.1. Send email to previous person
// 2.2. Remove previous person from event, add new person to event
// 3. If it has been "warningDay" number days since prompt was updated, send warning
// ==========================================
function checkDaysProgress() {
  var searchDate = new Date();
  changeDate(searchDate, -1);
  var scriptInfo = getSheetInformation('ScriptInfo');
  var scriptLength = scriptInfo.data.length - 1;
  var currParticipantIdx = scriptInfo.index.CurrentParticipantEmail;
  var currEmail = scriptInfo.data[scriptLength][currParticipantIdx];

  if (!writingCalendar) {
    writingCalendar = CalendarApp.getCalendarById(calendarId);
  }

  var events = writingCalendar.getEventsForDay(searchDate);
  var promptId = scriptInfo.data[scriptLength][scriptInfo.index.PromptID]
  var lastDateIdx = scriptInfo.index.LastDate;
  var lastDate = scriptInfo.data[scriptLength][lastDateIdx]
  var currentNumber = scriptInfo.data[scriptLength][scriptInfo.index.CurrentNumber]
  var promptPrefix = getTitlePrefix(promptId, currentNumber);
  var promptEvent, currEvent, currEventTitle;
  var participantInfo, partEmailIdx, submissionInfo, finaleSections = [], summaryTitle;

  for (var i = 0; i < events.length; i++) {
    currEvent = events[i];
    currEventTitle = currEvent.getTitle();
    if (currEventTitle.indexOf(promptPrefix) === 0) {
      // Move out yesterday's event if it's not all-day
      // If 8 - 9 days late, send out email that it will be switched to next person
      // If 10 days late, switch to next person
      promptEvent = currEvent;
      var startTime = promptEvent.getStartTime();
      var endTime = promptEvent.getEndTime();

      if (!promptEvent.isAllDayEvent()) {
        console.log('Moving out event');
        var daysSince = getDayDifference(lastDate, new Date());
        // const isReadyForLLM = true;
        const isReadyForLLM = isLLM(currEmail) && daysSince >= warningDay;
        if (isReadyForLLM) {
          if (!participantInfo) {
            participantInfo = getSheetInformation('Participants');
            partEmailIdx = participantInfo.index.Email;
          }

          var allParts = [];
          for (var i = 1; i < participantInfo.data.length; i++) {
            const email = participantInfo.data[i][partEmailIdx];
            if (!isLLM(email)) {
              allParts.push(email);
            }
          }

          const currDescription = promptEvent.getDescription().trim();
          var currRoundIdx = scriptInfo.index.currentRounds;
          var currentRound = scriptInfo.data[scriptLength][currRoundIdx];
          const numParticipants = allParts.length;
          const llmResult = runLLM(currEventTitle, currDescription, currentNumber, currentRound * numParticipants);
          const results = getResultFromLLM(llmResult);
          const newContent = results.newWriting.trim()
          const newDescription = currDescription ?
            `${currDescription}\n${newContent}` :
            `${currDescription}${newContent}`;

          const emailForNewContent = '\nNEW CONTENT\n' +
                newDescription +
                '\n' + noteDivider +
                JSON.stringify(llmResult, null, 2) +
                '\n' + noteDivider +
                'Link: ' + writingSpreadsheetUrl;

          if (currDescription) {
            MailApp.sendEmail({
              to: allParts.join(','),
              subject: emailPrefix + 'LLM has given feedback',
              body: 'FEEDBACK:' +
                `\n\nGRAMMAR FEEDBACK\n` +
                results.grammarFeedback +
                '\n\nPLOT FEEDBACK\n' +
                results.plotFeedback +
                '\n' + noteDivider +
                emailForNewContent,
            });
          } else {
            MailApp.sendEmail({
              to: myEmail,
              subject: emailPrefix + 'LLM has written',
              body: emailForNewContent,
            });
          }

          promptEvent.setDescription(newDescription);
        } else if (daysSince >= moveDay) {
          if (!participantInfo) {
            participantInfo = getSheetInformation('Participants');
            partEmailIdx = participantInfo.index.Email;
            submissionInfo = getSheetInformation('Submission');
          }

          var partEmailIdx = participantInfo.index.Email;
          var currNumberTotalIdx = scriptInfo.index.CurrentNumberTotal;
          var currentNumberTotal = scriptInfo.data[scriptLength][currNumberTotalIdx] + 1;
          var nextParticipantRow = calculateNextParticipant(participantInfo, currentNumberTotal, startTime);
          endTime = getEndTime(startTime);
          var nextGuest = nextParticipantRow[partEmailIdx];
          promptEvent.removeGuest(getEmailOnly(currEmail));
          promptEvent.addGuest(getEmailOnly(nextGuest));

          // Update ScriptInfo sheet
          scriptInfo.data[scriptLength][lastDateIdx] = new Date();
          scriptInfo.data[scriptLength][currNumberTotalIdx] = currentNumberTotal;
          scriptInfo.data[scriptLength][currParticipantIdx] = nextGuest;
          scriptInfo.range.setValues(scriptInfo.data);

          console.log(`Giving event to ${nextGuest} and updated total to ${currentNumberTotal}`);

          // Now update the last ParticipantEmail on Submission page
          var lastSubmissionIdx = submissionInfo.data.length;
          var participantEmailIdx = submissionInfo.index.ParticipantEmail + 1;
          var newParticipantRange = submissionInfo.sheet
            .getRange(lastSubmissionIdx, participantEmailIdx, 1, 1);

          var newNote = newParticipantRange.getNotes();
          newNote[0][0] += (noteDivider + new Date().toLocaleString() + ' overwrote:\n' + currEmail + '\n');
          newParticipantRange.setNotes(newNote);
          newParticipantRange.setValues([[nextGuest]]);
        } else if (daysSince >= warningDay) {
          var sendEmail = myEmail;
          if (currEmail !== sendEmail) {
            sendEmail += (',' + currEmail)
          }

          console.log(`Sending warning to ${currEmail}`);

          MailApp.sendEmail({
            to: sendEmail,
            subject: emailPrefix + getEmailUser(currEmail) + ' Update event!',
            body: 'It has been ' + daysSince + ' days. Update within the next ' +
              (moveDay - daysSince) + ' day(s) or #' + promptPrefix +
              ' (' + currEventTitle + ') will be reassigned.' +
              '\n' + noteDivider +
              'Link: ' + writingSpreadsheetUrl,
          });
        }

        changeDate(startTime, 1);
        changeDate(endTime, 1);
        promptEvent.setTime(startTime, endTime);

        console.log(`Moved event to ${startTime}`);

        MailApp.sendEmail({
          to: myEmail,
          subject: `${emailPrefix}Event moved to next day`,
          body: 'Event moved to next day for #' + promptPrefix +
            ' (' + currEventTitle + ') originally on ' +
            searchDate.toLocaleString() +
            ', moved to ' + startTime.toLocaleString() +
            '\n' + noteDivider +
            'Link: ' + writingSpreadsheetUrl,
        });
      }
    } else if (currEventTitle.indexOf(summaryHeader) === 0) {
      var matches = currEventTitle.match(summaryPartRegex);
      if (matches) {
        summaryTitle = summaryTitle || currEventTitle.replace(matches[1], '');
        console.log(`Found matches for summary ${summaryTitle}`);
        finaleSections[parseInt(matches[2], 10)] = currEvent.getDescription();
      } else {
        console.log('No matches for summary: %s', currEventTitle);
      }
    }
  }

  // Send finale email if needed
  if (finaleSections.length) {
    // If it's the finale -- give it 1 day to be updated, then send it out to everyone!
    if (!participantInfo) {
      participantInfo = getSheetInformation('Participants');
      partEmailIdx = participantInfo.index.Email;
    }

    var allParts = [];
    for (var i = 1; i < participantInfo.data.length; i++) {
      allParts.push(participantInfo.data[i][partEmailIdx]);
    }

    var shortenedString = summaryTitle.substring(0, 200);
    var lastIndex = summaryTitle.lastIndexOf(' ');
    shortenedString = shortenedString.substring(0, lastIndex - 2);
    var allSections = finaleSections.join(noteDivider);
    const additionalEmails = scriptInfo.data[scriptLength][scriptInfo.index.AdditionalEmails];
    const to = allParts.join(',') + (additionalEmails ? ',' + additionalEmails : '');
    console.log(`Sending email of overview to ${to}`);
    MailApp.sendEmail({
      to,
      subject: emailPrefix + shortenedString,
      body: 'Prompt:\n\n' + summaryTitle + '\n' +
        cleanHtmlFromDescription(allSections) +
        noteDivider +
        'Google Doc Link: ' + totalDoc +
        '\nSpreadsheet Link: ' + writingSpreadsheetUrl,
    });

    // Now add it to google docs
    console.log(`Adding to google docs`);
    var doc = DocumentApp.openById(docID);
    var body = doc.getBody();
    var header = body.appendParagraph(summaryTitle);
    header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(new Date().toString());
    body.appendParagraph('\n' + cleanHtmlFromDescription(allSections));
  }
}

// ==========================================
// Run on change (or on initialization of creative writing) --
// 1. Check text -- update "Submission" tab
//      Figure out which parts are changed and update text in spreadsheet (as identified by CalendarEventId) on Submission tab
//      Update "EditedDate"
// 2.1. update nextSyncToken
// 2.2. If is last event: increment currentNumber/CurrentNumberTotal
// 3.1. If currentNumber is 2x rounds -- reset currentNumber and randomly identify new PromptID, create new all-day calendar event, invite both participants, create new row for information
// 3.2. Identify new participant, create new calendar event with updated text in note, create new row with "---------" between each day's note
// ==========================================
function runOnChange() {
  // Gets a script lock before modifying a shared resource.
  const lock = LockService.getScriptLock();

  // Waits for up to 0.01 seconds for other processes to finish.
  lock.waitLock(10);

  if (!writingCalendar) {
    writingCalendar = CalendarApp.getCalendarById(calendarId);
  }

  // Calculate latest event
  var scriptInfo = getSheetInformation('ScriptInfo');
  var promptIdIdx = scriptInfo.index.PromptID;
  var currNumberIdx = scriptInfo.index.CurrentNumber;
  var defaultRoundIdx = scriptInfo.index.defaultRounds;
  var currParticipantIdx = scriptInfo.index.CurrentParticipantEmail;
  var scriptLength = scriptInfo.data.length - 1;
  var promptId = scriptInfo.data[scriptLength][promptIdIdx];
  var currentNumber = (scriptInfo.data[scriptLength][currNumberIdx] || 0);
  var latestEventPrefix = getTitlePrefix(promptId, currentNumber);
  var lastEvent = false;
  var newToken = false;

  // Get submission information
  var submissionInfo = getSheetInformation('Submission', true);
  var emailIdx = submissionInfo.index.ParticipantEmail;
  var subPromptId = submissionInfo.index.PromptID;
  var subCurrNumIdx = submissionInfo.index.CurrentNumber;
  var eventIdIdx = submissionInfo.index.EventName;
  var calendarEventIdx = submissionInfo.index.CalendarEventId;
  var inNumberIdx = submissionInfo.index.InNumbers;
  var textIdx = submissionInfo.index.Text;
  var wordsIdx = submissionInfo.index.Words;
  var charIdx = submissionInfo.index.Characters;
  var createdDateIdx = submissionInfo.index.CreatedDate;
  var editedDateIdx = submissionInfo.index.EditedDate;
  var submissionInfoNeedsUpdating = false;
  var lastSubmissionIdx = submissionInfo.data.length;
  var totalCharacters = 0;
  var totalSubmissions = 0;

  var eventToFullTextArray = {}; // links ID with array of text
  processSubmissions();

  var updatedEventIdToTextArray = {}; // links ID with array of text, but only for updated ones
  getEvents(calendarId);

  var currArray;
  for (var calId in updatedEventIdToTextArray) {
    if (updatedEventIdToTextArray.hasOwnProperty(calId)) {
      currArray = updatedEventIdToTextArray[calId];
      writingCalendar.getEventById(calId)
        .setDescription(currArray.join(noteDivider));
    }
  }

  if (submissionInfoNeedsUpdating) {
    console.log('Submission Updating');
    submissionInfo.range.setValues(submissionInfo.data);
    submissionInfo.range.setNotes(submissionInfo.note);
  }

  // Handle cases when new section has been added
  // Or if new creative writing is being created
  const isNewCreativeWriting = submissionInfo.data.length <= 1;
  if (lastEvent || isNewCreativeWriting) {
    const message = isNewCreativeWriting ? 'Starting new creative writing thread' : 'Last Event Updating';
    console.log(message);
    var currNumberTotalIdx = scriptInfo.index.CurrentNumberTotal;
    var newCurrNumberTotal = (scriptInfo.data[scriptLength][currNumberTotalIdx] || 0) + 1;
    scriptInfo.data[scriptLength][currNumberTotalIdx] = newCurrNumberTotal;
    scriptInfo.data[scriptLength][scriptInfo.index.LastDate] = new Date();

    // get next day
    let nextStartTime = new Date();

    if (nextStartTime.getDate() < lateEventTime.getDate()) {
      nextStartTime = lastEventDate;
    }

    changeDate(nextStartTime, 1);

    // Add new row to submissionsheet
    var newNumber = currentNumber + 1;
    scriptInfo.data[scriptLength][currNumberIdx] = newNumber;

    // Figure out next guest
    var participantInfo = getSheetInformation('Participants');
    var partEmailIdx = participantInfo.index.Email;
    var numberParticipants = participantInfo.data.length - 1;
    var nextParticipantRow = calculateNextParticipant(participantInfo, newCurrNumberTotal, nextStartTime);
    var guest = nextParticipantRow[partEmailIdx];

    scriptInfo.data[scriptLength][currParticipantIdx] = guest;

    var title, text, currentText;
    var currRoundIdx = scriptInfo.index.currentRounds;
    var currentRound = scriptInfo.data[scriptLength][currRoundIdx];
    if (isNewCreativeWriting || (newNumber > (numberParticipants * currentRound + 1))) {
      // Set new currenRound and promptId
      var oldPromptId = scriptInfo.data[scriptLength][promptIdIdx];
      var promptInfo = getSheetInformation('Prompts');
      var numberPrompts = promptInfo.data.length - 1;
      var dateIdx = promptInfo.index.Date;
      var promptToUse = 1;
      var newPrompt = promptInfo.data[promptToUse];
      while (newPrompt[dateIdx]) {
        promptToUse = Math.ceil(Math.random() * numberPrompts);
        newPrompt = promptInfo.data[promptToUse];
      }

      // Update the Date and Order of the prompt row
      promptInfo.sheet.getRange(promptToUse + 1, dateIdx + 1, 1, 2)
        .setValues([[new Date(), newCurrNumberTotal]]);

      // Define the title/text for new prompt
      promptId = newPrompt[promptInfo.index.PromptID];
      scriptInfo.data[scriptLength][currNumberIdx] = 1;
      scriptInfo.data[scriptLength][promptIdIdx] = promptId;
      scriptInfo.data[scriptLength][currRoundIdx] = scriptInfo.data[scriptLength][defaultRoundIdx];
      title = getTitlePrefix(promptId, 1) + ' ' + newPrompt[promptInfo.index.Prompt];
      text = `${getEmailUser(guest)}:\n`;
      console.log('New Prompt %s: %s', promptId, title);

      // Create overview events for the last writing prompt, keeping within calendar description limit
      if (lastEvent) {
        createOverviewEvent(lastEvent, participantInfo);
      }
    } else {
      var newPrefix = getTitlePrefix(promptId, newNumber);
      scriptInfo.data[scriptLength][currNumberIdx] = newNumber;
      title = lastEvent.getTitle().replace(new RegExp('^' + latestEventPrefix), newPrefix);
      currentText = getEmailUser(guest)
      text = lastEvent.getDescription() + noteDivider + currentText + ':\n';
      console.log('New Section: %s', title);
    }

    const avgCharIdx = scriptInfo.index.AverageCharacters;
    const averageChar = Math.round(totalCharacters / totalSubmissions);
    console.log(`Update ScriptInfo:\nNew Token ${newToken}\nNew Average Characters: ${averageChar}`);
    scriptInfo.data[scriptLength][avgCharIdx] = averageChar;

    // If text is longer than (charLimit - maximum length - graceLimit), then remove one section
    // Only need to calculate for events that have text (ie: not new prompts)
    var inNumbers = '';
    var textLength = text.length;
    if (textLength) {
      var avgChars = scriptInfo.data[scriptLength][avgCharIdx];
      var charLimitWithGrace = charLimit - avgChars - graceLimit;
      console.log('Current text length: %s', textLength);
      var firstSectionRegexp = new RegExp('^[\\s\\S]*?' + divider + '+?\\s*')
      inNumbers = submissionInfo.titlePrefixToRow[latestEventPrefix][inNumberIdx] +
        latestEventPrefix.replace(':', '') + ', ';

      while (textLength >= charLimitWithGrace) {
        text = text.replace(firstSectionRegexp, '');
        inNumbers = inNumbers.replace(/[0-9\.]+,\s*/, '');
        textLength = text.length;
        console.log('Trim Description: %s', text);
        console.log('Trim InNumbers: %s', inNumbers);
        console.log('New text length: %s', textLength);
      }

      console.log('InNumbers: %s', inNumbers);
    }

    createEventAndNewRow({
      title,
      text,
      startDate: nextStartTime,
      guests: guest,
      inNumbers: inNumbers,
      addNewRow: true,
      currentText,
    });
  } // end of "if (lastEvent || isNewCreativeWriting) {"

  if (lastEvent || newToken) {
    // Update scriptInfo
    console.log('Update ScriptInfo');
    scriptInfo.range.setValues(scriptInfo.data);
  }

  // Sleep for 1000 seconds so other triggered events can't run
  Utilities.sleep(1000);
  lock.releaseLock();

  // Helper functions
  /**
   * Add overview event
   */
  function createOverviewEvent(lastEvent, participantInfo) {
    var currIndex = 1;
    var totalCharCount = 0;
    var originalTitle = lastEvent.getTitle();

    function getNewTitle() {
      return originalTitle.replace(new RegExp('^' + latestEventPrefix), summaryHeader + oldPromptId + '.' + currIndex + ': ');
    }

    // Get all participants
    // If it's the finale -- give it 1 day to be updated, then send it out to everyone!
    if (!participantInfo) {
      participantInfo = getSheetInformation('Participants');
      partEmailIdx = participantInfo.index.Email;
    }

    var allParticipants = [];
    for (var i = 1; i < participantInfo.data.length; i++) {
      allParticipants.push(participantInfo.data[i][partEmailIdx]);
    }

    // Get all parts
    var overviewTitle = getNewTitle();
    var allParts = [];
    var currText;
    var sectionDates = [];
    for (var j = 1; j < submissionInfo.data.length; j++) {
      if (oldPromptId !== submissionInfo.data[j][subPromptId]) {
        continue;
      }

      currText = submissionInfo.data[j][textIdx];
      totalCharCount += currText.length;
      sectionDates.push(submissionInfo.data[j][createdDateIdx], submissionInfo.data[j][editedDateIdx]);
      if (totalCharCount > (charLimit - graceLimit)) {
        console.log('Adding Overview Part: %s (%s)', overviewTitle, currIndex);
        var startDateText = currIndex === 1 ? 'Started on: ' + sectionDates[0].toDateString() + noteDivider : '';
        var endDateText = j >= submissionInfo.data.length - 2 ? 'Finished on: ' + sectionDates[sectionDates.length - 1].toDateString() + noteDivider : '';
        createEventAndNewRow({
          title: overviewTitle,
          text: startDateText + allParts.join(noteDivider) + endDateText,
          startDate: nextStartTime,
          guests: allParticipants.join(','),
          isAllDay: true,
        });
        allParts = [];
        currIndex += 1;
        totalCharCount = currText.length;
        overviewTitle = getNewTitle();
      }

      allParts.push(currText);
    }

    if (allParts.length) {
      console.log('Adding Overview Final: %s (%s)', overviewTitle, currIndex);
      var startDateText = currIndex === 1 ? 'Started on: ' + sectionDates[0].toDateString() + noteDivider : '';
      var endDateText = 'Finished on: ' + sectionDates[sectionDates.length - 1].toDateString() + noteDivider;
      createEventAndNewRow({
        title: overviewTitle,
        text: allParts.join(noteDivider) + endDateText,
        startDate: nextStartTime,
        guests: allParticipants.join(','),
        isAllDay: true,
      });
    }
  }

  /**
   * Process previous submissions
   */
  function processSubmissions() {
    submissionInfo.eventNameToRow = {};
    submissionInfo.titlePrefixToRow = {};
    var calendarId, currSub, inNumbers, iterPrefix, iterDescription, iterCalId, iterRow;
    for (var i = 1; i < submissionInfo.data.length; i++) {
      currSub = submissionInfo.data[i];

      submissionInfo.eventNameToRow[currSub[eventIdIdx]] = i;

      var titlePrefix = getTitlePrefix(currSub[subPromptId], currSub[subCurrNumIdx]);
      submissionInfo.titlePrefixToRow[titlePrefix] = currSub;

      // Add to characters and count of non-empty boxes so we can calculate average character count
      const currContent = currSub[textIdx];
      if (currContent >= prefixCharLength) {
        totalSubmissions += 1;
        totalCharacters += currContent.length;
      }

      // Now compile all sections that are part of this row's calendar event
      calendarId = currSub[calendarEventIdx];
      inNumbers = currSub[inNumberIdx].split(/,\s*/g);
      if (!eventToFullTextArray[calendarId]) {
        eventToFullTextArray[calendarId] = {
          descArray: [],
          usedInTitlePrefix: [],
          eventNameArray: [],
        };
      }

      for (var j = 0; j < inNumbers.length; j++) {
        iterPrefix = inNumbers[j];
        if (iterPrefix) {
          // Add all event's text to event
          iterRow = submissionInfo.data[i - inNumbers.length + j + 1];
          iterDescription = iterRow[textIdx].trim();
          eventToFullTextArray[calendarId].descArray.push(iterDescription);
          eventToFullTextArray[calendarId].eventNameArray.push(iterRow[eventIdIdx]);

          // Add this iterprefix to calendar event's usedInTitlePrefix
          iterCalId = submissionInfo.titlePrefixToRow[iterPrefix + ':'][calendarEventIdx];
          eventToFullTextArray[iterCalId].usedInTitlePrefix.push(titlePrefix);
        }
      }

      eventToFullTextArray[calendarId].descArray.push(currSub[textIdx] || '');
      eventToFullTextArray[calendarId].eventNameArray.push(currSub[eventIdIdx]);
    }
  }

  /**
   * Incrementally gets only updated events
   * Skips events that have summaryHeader in the title
   */
  function getEvents(changedCalId, fullSync) {
    var options = {
      maxResults: 30,
    };

    var syncIdx = scriptInfo.index.SyncToken;
    var syncToken = scriptInfo.data[scriptLength][syncIdx];
    if (syncToken && !fullSync) {
      options.syncToken = syncToken;
    } else {
      options.timeMin = getRelativeDate(-30, 0).toISOString();
    }

    // Retrieve events one page at a time.
    var events;
    var pageToken;
    do {
      try {
        options.pageToken = pageToken;
        events = Calendar.Events.list(changedCalId, options);
      } catch (e) {
        // Check to see if the sync token was invalidated by the server;
        // if so, perform a full sync instead.
        console.log(e.message);
        if (e.message === 'Sync token is no longer valid, a full sync is required.') {
          scriptInfo.data[scriptLength][syncIdx] = '';
          scriptInfo.range.setValues(scriptInfo.data);
          getEvents(changedCalId, true);
          return;
        }

        throw new Error(e.message);
      }

      if (events.items && events.items.length > 0) {
        for (var i = 0; i < events.items.length; i++) {
          var event = events.items[i];
          if (event.status === 'cancelled') {
            console.log('Event id %s was cancelled.', event.id);
          } else if (event.summary.indexOf(summaryHeader) !== 0) {
            // All-day event if event.start.date
            // Events that don't last all day; they have defined start times.
            updateEventIfChanged(event.id, event.summary, event.description)
          } else {
            console.log('Event id %s is summary.', event.id);
          }
        }
      }

      pageToken = events.nextPageToken;
    } while (pageToken);

    if (events && events.nextSyncToken) {
      newToken = true;
      scriptInfo.data[scriptLength][syncIdx] = events.nextSyncToken;
    }
  }

  /**
   * Process current event and see if it's text has been changed
   * If the event starts with latestEventPrefix, save as lastEvent
   * and make the event all-day
  */
  function updateEventIfChanged(eventId, eventTitle, eventDescription) {
    // Check to see if description changed
    var currIdx = submissionInfo.eventNameToRow[eventTitle];
    var currRow = submissionInfo.data[currIdx];
    var dividerRegex = new RegExp('\\s*' + divider + '+?\\s*');

    if (!eventDescription || eventDescription.length < prefixCharLength) {
      return;
    }

    // Check every section of current eventTitle,
    // update submissionInfo and related calendar events as necessary
    // If last event, set it as lastEvent and send email
    var isChanged = false;
    var calendarSectionsNew = eventDescription.split(dividerRegex);
    var calendarSectionsOld = eventToFullTextArray[eventId];
    console.log(calendarSectionsOld);

    var currSectionNew, currSectionOld, currSectionIdx, currSectionRow;

    // Recursively update description for all affected calendar events
    function updateAllCalendarForRow(inNumbers, idxFromEnd) {
      console.log('updateAllCalendarForRow for %s, %s', inNumbers.toString(), idxFromEnd);
      var calID = currSectionRow[calendarEventIdx];

      if (idxFromEnd && idxFromEnd > inNumbers.length) {
        console.log('updateAllCalendarForRow Done');
        return;
      } else if (idxFromEnd) {
        var inNumberRow = submissionInfo.titlePrefixToRow[inNumbers[idxFromEnd - 1]];
        calID = inNumberRow[calendarEventIdx];
      } else {
        idxFromEnd = 0;
      }

      if (!updatedEventIdToTextArray[calID]) {
        updatedEventIdToTextArray[calID] = eventToFullTextArray[calID].descArray;
      }

      var updatedDescArray = updatedEventIdToTextArray[calID];
      var idxToUpdate = updatedDescArray.length - idxFromEnd - 1;
      updatedDescArray[idxToUpdate] = currSectionNew;

      console.log('idxToUpdate: %s', idxToUpdate);

      updateAllCalendarForRow(inNumbers, ++idxFromEnd);
    }

    var totalWordsWrote = 0;
    for (var i = 0; i < calendarSectionsNew.length; i++) {
      console.log('calendarSectionsNew i: %s', i);
      currSectionNew = cleanHtmlFromDescription(calendarSectionsNew[i]);
      currSectionOld = calendarSectionsOld.descArray[i].trim();
      totalWordsWrote += (Math.max(0, getWordCount(currSectionNew) - getWordCount(currSectionOld)));

      if (currSectionNew !== currSectionOld) {
        isChanged = true;
        submissionInfoNeedsUpdating = true;
        currSectionIdx = submissionInfo.eventNameToRow[calendarSectionsOld.eventNameArray[i]];
        currSectionRow = submissionInfo.data[currSectionIdx];
        console.log('currSectionIdx: %s', currSectionIdx);
        console.log('currSectionRow: %s', currSectionRow.toString());

        // Update text
        submissionInfo.note[currSectionIdx][textIdx] += (noteDivider + new Date().toLocaleString() + ' overwrote:\n' + currSectionRow[textIdx] + '\n');
        currSectionRow[textIdx] = currSectionNew;

        // Update word and character count
        submissionInfo.note[currSectionIdx][wordsIdx] += (new Date().toLocaleString() + ' overwrote:\n' + currSectionRow[wordsIdx] + '\n');
        currSectionRow[wordsIdx] = getWordCount(currSectionNew);
        submissionInfo.note[currSectionIdx][charIdx] += (new Date().toLocaleString() + ' overwrote:\n' + currSectionRow[charIdx] + '\n');
        currSectionRow[charIdx] = currSectionNew.length;

        if (currSectionRow[editedDateIdx]) {
          submissionInfo.note[currSectionIdx][editedDateIdx] += (new Date().toLocaleString() + ' overwrote:\n' + currSectionRow[editedDateIdx] + '\n');
        }

        currSectionRow[editedDateIdx] = new Date();

        updateAllCalendarForRow(eventToFullTextArray[currSectionRow[calendarEventIdx]].usedInTitlePrefix);
      }
    }

    if (isChanged && eventTitle.indexOf(latestEventPrefix) === 0) {
      lastEvent = writingCalendar.getEventById(eventId);

      var currWordsWrote = getWordCount(calendarSectionsNew[calendarSectionsNew.length - 1]);
      if (!currWordsWrote || currWordsWrote < prefixCharLength) {
        console.log('Nothing added to section');
        lastEvent = undefined;
        return;
      }

      // Send email to user letting them now their current contribution and how many words they wrote
      lastEvent.setAllDayDate(lastEvent.getStartTime());
      console.log('Last event changed, sending email');
      MailApp.sendEmail({
        to: currRow[emailIdx],
        subject: emailPrefix + 'Thanks for writing ' + currWordsWrote + ' words today! (' + new Date().toDateString() + ')',
        body: 'Prompt:\n\n' + lastEvent.getTitle() + noteDivider +
          cleanHtmlFromDescription(eventDescription) +
          '\n\nNew Count: ' + currWordsWrote + '/' + (currWordsWrote + totalWordsWrote) +
          '\n\nTotal Count: ' + getWordCount(eventDescription) +
          '\n' + noteDivider +
          'Link: ' + writingSpreadsheetUrl,
      });
    }
  }

  /**
   * Creates new writing event
   */
  function createEventAndNewRow(config) {
    var title = config.title; // calendar title
    var text = config.text; // calendar description
    var currentText = config.currentText // if there's prefix that should always be added
    var startDate = config.startDate; // calendar start date
    var guests = config.guests; // calendar attendees
    var isAllDay = config.isAllDay; // will not create new row
    var inNumbers = config.inNumbers; // not needed if isAllDay
    var addNewRow = config.addNewRow; // not needed if isAllDay

    // Remove API guests
    if (isLLM(guest)) {
      guests = calendarId;
    }

    var event;
    var eventOptions = {
      description: text,
      location: writingSpreadsheetUrl,
      guests: getEmailOnly(guests),
    };

    if (isAllDay) {
      event = writingCalendar.createAllDayEvent(title, startDate, eventOptions);
    } else {
      var endDate = getEndTime(startDate);
      event = writingCalendar.createEvent(title, startDate, endDate, eventOptions);
    }

    event.setGuestsCanModify(true);

    if (addNewRow) {
      // Now add this new row to "Submission" spreadsheet
      // Get range by row, column, row length, column length
      var newRow = [];
      for (var i = 0; i < submissionInfo.data[0].length;) newRow[i++] = '';

      newRow[submissionInfo.index.ParticipantEmail] = guests;
      newRow[subPromptId] = scriptInfo.data[scriptLength][promptIdIdx];
      newRow[subCurrNumIdx] = scriptInfo.data[scriptLength][currNumberIdx];
      newRow[eventIdIdx] = title;
      newRow[calendarEventIdx] = event.getId().replace('@google.com', '');
      newRow[submissionInfo.index.InNumbers] = inNumbers || '';
      newRow[submissionInfo.index.CreatedDate] = new Date();
      newRow[textIdx] = currentText || '';
      newRow[wordsIdx] = 0;
      newRow[charIdx] = 0;
      lastSubmissionIdx++;
      var cells = submissionInfo.sheet.getRange(lastSubmissionIdx, 1, 1, newRow.length);
      cells.setValues([newRow])
    }

    return event;
  }
}


// ==========================================
// GENERIC HELPER FUNCTIONS
// ==========================================

function getEmailUser(email) {
  return getEmailOnly(email).split('@')[0];
}

function getEmailOnly(emails) {
  return emails.replaceAll(/\+.+?@/g, '@');
}

// Get sheet information - sheet, data, and index
var activeSpreadsheet;

function getSheetInformation(sheetName, includeNote) {
  if (!activeSpreadsheet) {
    activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  var result = {
    sheet: activeSpreadsheet.getSheetByName(sheetName)
  }
  result.range = result.sheet.getDataRange();
  result.data = result.range.getValues();
  result.index = indexSheet(result.data);

  if (includeNote) {
    result.note = result.range.getNotes();
  }

  return result;
}

// Create hash with column name keys pointing to column index
// For greater flexibility (columns can be moved around)
function indexSheet(sheetData) {
  var result = {},
    length = sheetData[0].length;

  for (var i = 0; i < length; i++) {
    result[sheetData[0][i]] = i;
  }

  return result;
}

function changeDate(dateObj, change) {
  dateObj.setDate(dateObj.getDate() + change);
}

function getTitlePrefix(promptId, currentNumber) {
  return promptId + '.' + currentNumber + ':';
}

function getWordCount(text) {
  var matches = text.match(/\b(\w+)\b/g) || '';
  return matches.length;
}

/**
 * Helper function to get a new Date object relative to the current date.
 * @param {number} daysOffset The number of days in the future for the new date.
 * @param {number} hour The hour of the day for the new date, in the time zone
 *     of the script.
 * @return {Date} The new date.
 */
function getRelativeDate(daysOffset, hour) {
  var date = new Date();
  date.setDate(date.getDate() + daysOffset);
  date.setHours(hour);
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}

function getEndTime(startTime) {
  var newDate = new Date(startTime);
  newDate.setHours(startTime.getHours() + lengthEvent);
  return newDate;
}

function getDayDifference(day1, day2) {
  var millisecondsPerDay = 24 * 60 * 60 * 1000;
  return Math.abs((day2 - day1) / millisecondsPerDay);
}

function cleanHtmlFromDescription(text) {
  return text.replaceAll('<br>', '\n').replaceAll(/<.+?>/g, '').trim();
}

// Calculate next participant and update dateObj to that person's ideal time
function calculateNextParticipant(participantInfo, currNumberTotal, dateObj) {
  var numberParticipants = participantInfo.data.length - 1;
  var nextParticipantIdx = currNumberTotal % numberParticipants || numberParticipants;
  var nextParticipantRow = participantInfo.data[nextParticipantIdx];

  // Set correct time for nextStartTime
  dateObj.setHours(nextParticipantRow[participantInfo.index.BestTimeET]);
  dateObj.setMinutes(0);
  dateObj.setMilliseconds(0);

  return nextParticipantRow;
}

function isLLM(text) {
  return text.startsWith('https://');
}

function runLLM(prompt, text, currentNumber, totalNumber) {
  let progress = 'Write as if this was the beginning of the story';
  if (currentNumber >= totalNumber - 2) {
    progress = 'Start writing a conclusion for the story.';
  } else if (currentNumber >= totalNumber / 2) {
    progress = 'Write as if this was the middle of the story.';
  }

  const beginningPrompt = text ?
    `Based on the user input, give grammar and plot feedback, as well as write another ${lengthOfResponse} words that extends the story. ${progress}` :
    `Write ${lengthOfResponse} words that starts a story based on the prompt.`;
  const example = text ?
    `### USER TEXT
The sun has finally decided to peek through, abashed after a month-long absence.

*Maybe nobody noticed I was gone...* It thinks hopefully. That hope dies a quick death as it sees the crowds of peple peppering Hyde Park, arrayed in bathing suits and lounge chairs, morphing the park into what looks like a stretch of strangely green beach.

The sun puffs itself out in guilty pride, *People missed me; they surely, truly did. And here I've been too lethargic to show my face, and then too ashamed at my lethargy to change. I resolve to be better about appearing from now on. I ~will~ change. I really will.*

And for a week the sun shone steadily all day, every day. When it feels tired, it simply remembers the colorfully bedecked people, waiting to bask in its glow. *Can't let them down* it thinks with renewed resolve.

However, one day, after a particularly difficult rise, it notices something has changed. Rather than being greeted with people hastening to enjoy it's warmth, the sun was greeted with grim forebearance. People close themselves inside the moment the sun appears, and only begin cautiously creeping out when it starts to sink.

### RESPONSE
{
      "grammarFeedback": "- \"peple\" is misspelled. It should be \"people\"- Maintain consistency in the tense of the story; For example, \"the sun shone steadily\" and \"the sun was greeted\" is past tense, while \"The sun puffs itself\" is present tense.\n- Consider adding commas after introductory phrases for clarity (e.g., 'And for a week, the sun shone steadily...').\n- Instead of \"stretch of strangely green beach\", consider modifying to \"strange stretch of green beach\" for better flow.\n- Use of contractions like \"it's\" vs. \"its\" should be checked to ensure correct usage (eg: \"enjoy it's warmth\").\n- Preserve plural agreement by changing '*the grass lengthen*,' to ‘the grass lengthens.’",
      "plotFeedback": "- It is usually rainy for months in London. Consider changing \"month-long absence\" to \"multi-month absence\"\n- It would take more than a week of sun to make people dislike it. Consider changing \"And for a week\" to \"And for a month\".\n- Introducing some human-like conflict or dilemma related to this change could add tension—perhaps the sun feels unappreciated or confused by people's behavior.\n- You might also explore how other elements (clouds, flowers, trees) react to the sun being gone or present; it can enrich world-building and humor.\n\n- For example: maybe when the sun comes back flowers suddenly bloom, only to wilt later, or maybe there's a humorous scene where children hide from sunlight because they think it's too hot now.",
      "newWriting": "When they do come out, it is now with hats and sunglasses, not bathing suits and lounge chairs. The sun puzzles over this change, staying out longer and longer to observe the people more carefully. Yet after all that observation, it can only conclude with horror that it is no longer wanted. Nobody cheers when it arrives anymore, instead they hide away in fear. Nobody misses it when it sets anymore, instead they breath a sigh of relief.\nThe sun sadly withdraws, staying out for slightly shorter lengths of time each day. With an unhappy grimace, the sun thinks *I've worked so hard, for just a little appreciation, and now I don't even get that...*\nThe next time the sun has difficulty rising, it doesn't bother getting dressed. Instead, it lazily drifts across the sky in bed, nestled comfortably under its fluffy blanket of clouds.\nThe sun continues to half-heartedly rise each day, staying up for briefer periods and not troubling to come out of its blankets half the time.\nOne day, it decides to not even bother rising. *Will people really notice?* the sun grumbles to itself. Sleep and his inviting nest of blankets urge him to believe the people won't notice a thing.\nOne day of not rising seems to trigger a desire to hibernate even longer. The sun can't seem to get out of bed anymore; and when it does, it does so lethargically, with eyes half-shut and body securely wrapped up in a burrito of blankets.\nThis goes on for one month, then two, and just as the sun thinks everyone is enjoying his absence, a tired bird flops onto his bed."` :
    `### RESPONSE
{
      "grammarFeedback": "N/A",
      "plotFeedback": "N/A",
      "newWriting": "The sun has finally decided to peek through, abashed after a multi-month absence.\n*Maybe nobody noticed I was gone...* It thinks hopefully. That hope dies a quick death as it sees the crowds of people peppering Hyde Park, arrayed in bathing suits and lounge chairs, morphing the park into what looks like a strange stretch of green beach.\nThe sun puffs itself out in guilty pride, *People missed me; they surely, truly did. And here I've been too lethargic to show my face, and then too ashamed at my lethargy to change. I resolve to be better about appearing from now on. I ~will~ change. I really will.*\nAnd for a month, the sun shines steadily all day, every day. When it feels tired, it simply remembers the colorfully bedecked people, waiting to bask in its glow. *Can't let them down* it thinks with renewed resolve. Flowers bloom and trees turn verdant; the grass lengthen and birds abound.\nHowever, one day, after a particularly difficult rise, the sun notices something has changed. Rather than being greeted with people hastening to enjoy its warmth, the sun is greeted with grim forebearance. People close themselves inside the moment the sun appears, and only begin cautiously creeping out when it starts to sink."`
  const required = [
    'grammarFeedback',
    'plotFeedback',
    'newWriting',
  ];

  const data = {
    model: 'gpt-5',
    messages: [
      {
        role: 'system',
        content: [
          {
            type: 'text',
            text: `You are an experienced fiction editor who will catch grammar mistakes. You are participating in a writing round robin exercise. The writing prompt is: ${prompt}`
          }
        ]
      },
      {
        role: 'assistant',
        content: [
          {
            type: 'text',
            text: `${beginningPrompt}.

## EXAMPLE
### PROMPT
1.1: Outside the Window: What’s the weather outside your window doing right now? If that’s not inspiring, what’s the weather like somewhere you wish you could be?

${example}
}`
          },
        ],
      },
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text,
          },
        ],
      },
    ],
    response_format: {
      type: 'json_schema',
      json_schema: {
        name: 'writing_feedback',
        description: 'Writing partner that gives feedback on previous writing as well as generates new writing to continue the story in the same style',
        strict: true,
        schema: {
          type: 'object',
          additionalProperties: false,
          required,
          properties: {
            grammarFeedback: {
              type: 'string',
              description: 'Feedback on grammar changes for user provided content. Uses bullets with explanations and example. Do NOT suggest purely stylistic changes. Focus on true grammar or spelling mistakes.',
            },
            plotFeedback: {
              type: 'string',
              description: 'Suggestions to improve flow of the plot, ideas for future plot direction, or areas that could be expanded. Uses bullets with explanations and short examples',
            },
            newWriting: {
              type: 'string',
              description: `Continue the story for another approximate ${lengthOfResponse} words. Try not to repeat previously used words or phrases. ${progress}`,
            },
          },
        },
      },
    },
    temperature: 1, // Not supported for GTP5, otherwise would be using 0.8
    // frequency_penalty: 1.5, // Not supported for GTP5, prefer less repitition
    presence_penalty: 0
  };

  console.log('LLM Payload: ', data);
  data.messages.forEach((m) => console.log('LLM Messages: ', m.role, '\n', m.content[0].text));

  const result = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'get',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${token}`,
    },
    payload: JSON.stringify(data),
  }).getContentText();
  console.log('LLM Result: ', result);
  return JSON.parse(result);
}

function getResultFromLLM(jsonResult) {
  return JSON.parse(jsonResult.choices[0].message.content)
}
