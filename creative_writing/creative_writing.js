var writingCalendar;

// https://support.google.com/calendar/forum/AAAAd3GaXpEoHbbc3DuDt0/?hl=en
var charLimit = 8148;
var graceLimit = 500; // we allow for average length + gracelimit before dropping 1 day's writing

var warningDay = 5;
var moveDay = 8;

var lengthEvent = 3; // Hours writing event should last

var noteDivider = '\n===================\n';
var summaryHeader = 'Final: '

// ==========================================
// Runs at 8 AM UK Time
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
    writingCalendar =  CalendarApp.getCalendarById(calendarId);
  }

  var events = writingCalendar.getEventsForDay(searchDate);
  var promptId = scriptInfo.data[scriptLength][scriptInfo.index.PromptID]
  var lastDateIdx = scriptInfo.index.LastDate;
  var lastDate = scriptInfo.data[scriptLength][lastDateIdx]
  var currentNumber = scriptInfo.data[scriptLength][scriptInfo.index.CurrentNumber]
  var promptPrefix = getTitlePrefix(promptId, currentNumber);
  var promptEvent, currEvent, currEventTitle;
  var participantInfo, partEmailIdx, submissionInfo, finaleSent;

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
        var daysSince = getDayDifference(lastDate, new Date());
        if (daysSince >= moveDay) {
          if (!participantInfo) {
            participantInfo = getSheetInformation('Participants');
            partEmailIdx = participantInfo.index.Email;
            submissionInfo = getSheetInformation('Submission');
          }

          var partEmailIdx = participantInfo.index.Email;
          var currNumberTotalIdx = scriptInfo.index.CurrentNumberTotal;
          var currentNumberTotal = scriptInfo.data[scriptLength][currNumberTotalIdx] + 1;
          var nextParticipantRow = calculateNextParticipant(participantInfo, currentNumberTotal, startTime);
          endTime.setHours(startTime.getHours() + lengthEvent);
          var nextGuest = nextParticipantRow[partEmailIdx];
          promptEvent.removeGuest(currEmail);
          promptEvent.addGuest(nextGuest);

          // Update ScriptInfo sheet
          scriptInfo.data[scriptLength][lastDateIdx] = new Date();
          scriptInfo.data[scriptLength][currNumberTotalIdx] = currentNumberTotal;
          scriptInfo.data[scriptLength][currParticipantIdx] = nextGuest;
          scriptInfo.range.setValues(scriptInfo.data);

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

          MailApp.sendEmail({
            to: sendEmail,
            subject: '[CreativeWriting] ' + currEmail.split('@')[0] + ' Update event!',
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

        MailApp.sendEmail({
          to: myEmail,
          subject: '[CreativeWriting] Event moved to next day',
          body: 'Event moved to next day for #' + promptPrefix +
                ' (' + currEventTitle + ') originally on ' +
                searchDate.toLocaleString() +
                ', moved to ' + startTime.toLocaleString() +
                '\n' + noteDivider +
                'Link: ' + writingSpreadsheetUrl,
        });
      }
    } else if (!finaleSent && currEventTitle.indexOf(summaryHeader) === 0) {
      // If it's the finale -- give it 1 day to be updated, then send it out to everyone!
      var textIdx, promptIdIdx;
      if (!participantInfo) {
        participantInfo = getSheetInformation('Participants');
        partEmailIdx = participantInfo.index.Email;
        submissionInfo = getSheetInformation('Submission');
        textIdx = submissionInfo.index.Text;
        promptIdIdx = submissionInfo.index.PromptID;
      }

      var allParts = [];
      for (var i = 1; i < participantInfo.data.length; i++) {
        allParts.push(participantInfo.data[i][partEmailIdx]);
      }

      var allStory = [];
      for (var j = 1; j < submissionInfo.data.length; j++) {
        if (participantInfo.data[j][promptIdIdx] === promptId) {
          allStory.push(participantInfo.data[j][textIdx]);
        }
      }

      MailApp.sendEmail({
        to: allParts.join(',') + ',' + scriptInfo.data[scriptInfo.index.AdditionalEmails],
        subject: '[CreativeWriting] ' + currEventTitle,
        body: allStory.join('\n\n') +
              noteDivider +
              'Google Doc Link: ' + totalDoc +
              'Spreadsheet Link: ' + writingSpreadsheetUrl,
      });

      finaleSent = true;
    }
  }
}

// ==========================================
// Run on change --
// 1. Check text -- update "Submission" tab
//      Figure out which parts are changed and update text in spreadsheet (as identified by CalendarEventId) on Submission tab
//      Update "EditedDate"
// 2.1. update nextSyncToken
// 2.2. If is last event: increment currentNumber/CurrentNumberTotal
// 3.1. If currentNumber is 2x rounds -- reset currentNumber and randomly identify new PromptID, create new all-day calendar event, invite both participants, create new row for information
// 3.2. Identify new participant, create new calendar event with updated text in note, create new row with "---------" between each day's note
// ==========================================
function runOnChange() {
  if (!writingCalendar) {
    writingCalendar =  CalendarApp.getCalendarById(calendarId);
  }

  // Calculate latest event
  var scriptInfo = getSheetInformation('ScriptInfo');
  var promptIdIdx = scriptInfo.index.PromptID;
  var currNumberIdx = scriptInfo.index.CurrentNumber;
  var defaultRoundIdx = scriptInfo.index.defaultRounds;
  var currParticipantIdx = scriptInfo.index.CurrentParticipantEmail;
  var scriptLength = scriptInfo.data.length - 1;
  var promptId = scriptInfo.data[scriptLength][promptIdIdx]
  var currentNumber = scriptInfo.data[scriptLength][currNumberIdx]
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
  var textIdx = submissionInfo.index.Text;
  var wordsIdx = submissionInfo.index.Words;
  var editedDateIdx = submissionInfo.index.EditedDate;
  var submissionInfoNeedsUpdating = false;
  var lastSubmissionIdx = submissionInfo.data.length;
  processSubmissions();

  getEvents(calendarId);

  if (submissionInfoNeedsUpdating) {
    submissionInfo.range.setValues(submissionInfo.data);
    submissionInfo.range.setNotes(submissionInfo.note);
  }

  // Handle cases when new section has been added
  if (lastEvent) {
    var currNumberTotalIdx = scriptInfo.index.CurrentNumberTotal;
    var newCurrNumberTotal = scriptInfo.data[scriptLength][currNumberTotalIdx] + 1;
    scriptInfo.data[scriptLength][currNumberTotalIdx] = newCurrNumberTotal;
    scriptInfo.data[scriptLength][scriptInfo.index.LastDate] = new Date();

    // get next day
    var nextStartTime = new Date();

    if (nextStartTime.getDate() === lastEvent.getStartTime().getDate()) {
      changeDate(nextStartTime, 1);
    }

    // Add new row to submissionsheet
    var newNumber = currentNumber + 1;
    scriptInfo.data[scriptLength][currNumberIdx] = newNumber;

    // Figure out next guest
    var participantInfo = getSheetInformation('Participants');
    var partEmailIdx = participantInfo.index.Email;
    var numberParticipants = participantInfo.data.length - 1;
    var nextParticipantRow = calculateNextParticipant(participantInfo, newCurrNumberTotal, nextStartTime);
    var guest = nextParticipantRow[partEmailIdx];

    var title, text;
    var currRoundIdx = scriptInfo.index.currentRounds;
    var currentRound = scriptInfo.data[scriptLength][currRoundIdx];
    var offset = 1; // So that the same person isn't always the one starting prompts
    if (newNumber > (numberParticipants * currentRound + 1)) {
      // Set new currenRound and promptId
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
      promptInfo.sheet.getRange(promptToUse, dateIdx, 1, 2)
                      .setValues([[new Date(), newCurrNumberTotal]]);

      // Define the title/text for new prompt
      promptId = newPrompt[promptInfo.index.Prompt];
      scriptInfo.data[scriptLength][currNumberIdx] = 1;
      scriptInfo.data[scriptLength][promptIdIdx] = promptId;
      scriptInfo.data[scriptLength][currParticipantIdx] = guest;
      scriptInfo.data[scriptLength][currRoundIdx] = scriptInfo.data[scriptLength][defaultRoundIdx];
      title = getTitlePrefix(promptId, 1) + ' ' + newPrompt[promptInfo.index.Prompt];
      text = '';

      // Create overview events for the last writing prompt, keeping within calendar description limit
      var currIndex = 1;
      var totalCharCount = 0;
      var originalTitle = lastEvent.getTitle();

      function getNewTitle() {
        return originalTitle.replace(RegExp('^' + latestEventPrefix), summaryHeader + currIndex + ': ');
      }

      var overviewTitle = getNewTitle();
      var allParts = [];
      var currText;
      for (var i = 1; i < participantInfo.data.length; i++) {
        currText = participantInfo.data[i][partEmailIdx];
        totalCharCount += currText.length;
        if (totalCharCount > charLimit) {
          createEventAndNewRow({
            title: overviewTitle,
            text: lastEvent.getDescription(),
            startDate: nextStartTime,
            guests: allParts.join(','),
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
        createEventAndNewRow({
          title: overviewTitle,
          text: lastEvent.getDescription(),
          startDate: nextStartTime,
          guests: allParts.join(','),
          isAllDay: true,
        });
      }
    } else {
      var newPrefix = getTitlePrefix(promptId, newNumber);
      scriptInfo.data[scriptLength][currNumberIdx] = newNumber;
      title = lastEvent.getTitle().replace(RegExp('^' + latestEventPrefix), newPrefix);
      text = lastEvent.getDescription() + noteDivider + '\n';
    }

    // If text is longer than (charLimit - average length - graceLimit), then remove one section
    // Only need to calculate for events that have text (ie: not new prompts)
    var inNumbers = '';
    if (text.length) {
      var avgCharIdx = scriptInfo.index.AverageCharacters;
      var avgChars = scriptInfo.data[scriptLength][avgCharIdx];
      var firstSectionRegexp = new Regexp('^[\\s\\S]*?' + noteDivider + '\\s*')
      inNumbers = submissionInfo.titlePrefixToRow[latestEventPrefix][submissionInfo.index.InNumbers] +
                  latestEventPrefix + ', ';
      if (text.length >= (charLimit - avgChars - graceLimit)) {
        text = text.replace(firstSectionRegexp, '');
        inNumbers = inNumbers.replace(/[0-9\.]+,\s*/, '');
      }
    }

    createEventAndNewRow({
      title: title,
      text: text,
      startDate: nextStartTime,
      guests: guest,
      inNumbers: inNumbers,
      addNewRow: true,
    });
  } // end of "if (lastEvent) {"

  if (lastEvent || newToken) {
    // Update scriptInfo
    scriptInfo.range.setValues(scriptInfo.data);
  }

  // Helper functions
  /**
   * Incrementally gets only updated tasts
   */
  function processSubmissions() {
    submissionInfo.eventNameToRow = {};
    submissionInfo.titlePrefixToRow = {};
    for (var i = 1; i < submissionInfo.data.length; i++) {
      submissionInfo.eventNameToRow[submissionInfo.data[i][eventIdIdx]] = i;

      var titlePrefix = getTitlePrefix(submissionInfo.data[i][subPromptId],
                                       submissionInfo.data[i][subCurrNumIdx]);

      submissionInfo.titlePrefixToRow[titlePrefix] = submissionInfo.data[i];
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
    var currText = currRow[textIdx];
    if (eventDescription !== currText) {
      submissionInfoNeedsUpdating = true;

      // Update text
      submissionInfo.note[currIdx][textIdx] += (noteDivider + new Date().toLocaleString() + ' overwrote:\n' + currRow[textIdx] + '\n');
      currRow[textIdx] = eventDescription;

      // Update word count
      submissionInfo.note[currIdx][wordsIdx] += (new Date().toLocaleString() + ' overwrote:\n' + currRow[wordsIdx] + '\n');
      currRow[wordsIdx] = getWordCount(eventDescription);

      if (currRow[editedDateIdx]) {
        submissionInfo.note[currIdx][editedDateIdx] += (new Date().toLocaleString() + ' overwrote:\n' + currRow[editedDateIdx] + '\n');
      }

      currRow[editedDateIdx] = new Date();

      if (eventTitle.indexOf(latestEventPrefix) === 0) {
        lastEvent = writingCalendar.getEventById(eventId);

        if (lastEvent.getDescription() !== eventDescription) {
          // Don't process last event if it's just cascading update
          lastEvent.setDescription(eventDescription);
          lastEvent = undefined;
        } else {
          // Send email to user letting them now their current contribution and how many words they wrote
          lastEvent.setAllDayDate(lastEvent.getStartTime());
          var splitByDays = eventDescription.split('------');
          var wordsWrote, currIdx = splitByDays.length - 1;
          while (!wordsWrote && currIdx >= 0) {
            wordsWrote = getWordCount(splitByDays[currIdx]);
            currIdx--;
          }

          MailApp.sendEmail({
            to: currRow[emailIdx],
            subject: '[CreativeWriting] Thanks for writing ' + wordsWrote + ' words today! (' + currRow[editedDateIdx].toDateString() + ')',
            body: 'Prompt:\n\n' + lastEvent.getTitle() + noteDivider +
                  eventDescription +
                  '\n\nNew Count: ' + wordsWrote +
                  '\n\nTotal Count: ' + currRow[wordsIdx] +
                  '\n' + noteDivider +
                  'Link: ' + writingSpreadsheetUrl,
          });
        }
      } else {
        // If older event was edited, cascade changes
        var currPromptId = currRow[subPromptId];
        var nextNumber = currRow[subCurrNumIdx] + 1;
        var nextRow = submissionInfo.titlePrefixToRow[getTitlePrefix(currPromptId, nextNumber)];
        var text = nextRow && nextRow[textIdx];
        var changedText = nextRow && text.replace(currText, eventDescription);

        if (text !== changedText) {
          updateEventIfChanged(nextRow[calendarEventIdx], nextRow[eventIdIdx], changedText);
        }
      }
    }
  }

  /**
   * Creates new writing event
   */
  function createEventAndNewRow(config) {
    var title = config.title; // calendar title
    var text = config.text; // calendar description
    var startDate = config.startDate; // calendar start date
    var guests = config.guests; // calendar attendees
    var isAllDay = config.isAllDay; // will not create new row
    var inNumbers = config.inNumbers; // not needed if isAllDay
    var addNewRow = config.addNewRow; // not needed if isAllDay

    var event;
    var eventOptions = {
      description: text,
      location: writingSpreadsheetUrl,
      guests: guests,
    };

    if (isAllDay) {
      event = writingCalendar.createAllDayEvent(title, startDate, eventOptions);
    } else {
      var endDate = new Date(startDate);
      endDate.setHours(endDate.getHours() + lengthEvent);
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
      newRow[calendarEventIdx] = event.getId();
      newRow[submissionInfo.index.InNumbers] = inNumbers || '';
      newRow[submissionInfo.index.CreatedDate] = new Date();
      newRow[textIdx] = text;
      newRow[wordsIdx] = getWordCount(text);
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

function getDayDifference(day1, day2) {
  var millisecondsPerDay = 24 * 60 * 60 * 1000;
  return Math.abs((day2 - day1) / millisecondsPerDay);
}

// Calculate next participant and update dateObj to that person's ideal time
function calculateNextParticipant(participantInfo, currNumberTotal, dateObj) {
  var partEmailIdx = participantInfo.index.Email;
  var numberParticipants = participantInfo.data.length - 1;
  var nextParticipantIdx = currNumberTotal % numberParticipants || numberParticipants;
  var nextParticipantRow = participantInfo.data[nextParticipantIdx];

  // Set correct time for nextStartTime
  dateObj.setHours(nextParticipantRow[participantInfo.index.BestTimeUK]);
  dateObj.setMinutes(0);
  dateObj.setMilliseconds(0);

  return nextParticipantRow;
}
