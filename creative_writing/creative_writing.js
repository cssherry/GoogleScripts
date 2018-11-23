var writingCalendar;

// ==========================================
// Runs at 8 AM UK Time
// 1. If there isn't an all day event for correct prompt today, move today's event to next day
// ==========================================
function checkDaysProgress() {
  var searchDate = new Date();
  changeDate(searchDate, -1);
  var scriptInfo = getSheetInformation('ScriptInfo');

  if (!writingCalendar) {
    writingCalendar =  CalendarApp.getCalendarById(calendarId);
  }

  var events = writingCalendar.getEventsForDay(searchDate);
  var promptId = scriptInfo.data[scriptInfo.data.length - 1][scriptInfo.index.PromptID]
  var currentNumber = scriptInfo.data[scriptInfo.data.length - 1][scriptInfo.index.CurrentNumber]
  var promptPrefix = getTitlePrefix(promptId, currentNumber);
  var promptEvent, currEvent, currEventTitle;

  for (var i = 0; i < events.length; i++) {
    currEvent = events[i];
    currEventTitle = currEvent.getTitle();
    if (currEventTitle.indexOf(promptPrefix) === 0) {
      // Move out yesterday's event if it's not all-day
      promptEvent = currEvent;
      if (!promptEvent.isAllDayEvent()) {
        var startTime = promptEvent.getStartTime();
        var endTime = promptEvent.getEndTime();
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
                '\n\n---------\n' +
                'Link: ' + writingSpreadsheetUrl,
        });
      }
    } else if (currEventTitle.indexOf('Final:') === 0) {
      // If it's the finale -- give it 1 day to be updated, then send it out to everyone!
      var participantInfo = getSheetInformation('Participants');
      var partEmailIdx = participantInfo.index.Email;
      var allParts = [];

      for (var i = 1; i < participantInfo.data.length; i++) {
        allParts.push([participantInfo.data[i][partEmailIdx]]);
      }

      MailApp.sendEmail({
        to: allParts.join(',') + ',' + scriptInfo.data[scriptInfo.index.AdditionalEmails],
        subject: '[CreativeWriting] ' + currEventTitle,
        body: currEvent.getDescription() +
              '\n\n---------\n' +
              'Link: ' + writingSpreadsheetUrl,
      });
    }
  }
}

// ==========================================
// Run on change --
// 1. Check text -- if changed, update text in spreadsheet (as identified by date) on Submission tab and write "EditedDate"
// 2. update nextSyncToken, increment currentNumber
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

  if (lastEvent) {
    var currNumberTotalIdx = scriptInfo.index.CurrentNumberTotal;
    var newCurrNumberTotal = scriptInfo.data[scriptLength][currNumberTotalIdx] + 1;
    scriptInfo.data[scriptLength][currNumberTotalIdx] = newCurrNumberTotal;

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
    var nextParticipantIdx = newCurrNumberTotal % numberParticipants || numberParticipants;
    var nextParticipantRow = participantInfo.data[nextParticipantIdx];
    var guest = nextParticipantRow[partEmailIdx];

    // Set correct time for nextStartTime
    nextStartTime.setHours(nextParticipantRow[participantInfo.index.BestTimeUK]);
    nextStartTime.setMinutes(0);
    nextStartTime.setMilliseconds(0);

    var title, text;
    var currRoundIdx = scriptInfo.index.currentRounds;
    var currentRound = scriptInfo.data[scriptLength][currRoundIdx];
    if (newNumber > numberParticipants * currentRound) {
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
      promptInfo.sheet.getRange(promptToUse, dateIdx, 1, 2).setValues([[new Date(), newCurrNumberTotal]]);

      // Define the title/text
      var promptId = newPrompt[promptInfo.index.Prompt];
      scriptInfo.data[scriptLength][currNumberIdx] = 1;
      scriptInfo.data[scriptLength][promptIdIdx] = promptId;
      scriptInfo.data[scriptLength][currRoundIdx] = scriptInfo.data[scriptLength][scriptInfo.index.defaultRounds];
      title = newPrompt[promptInfo.index.Prompt];
      text = newPrompt[promptInfo.index.Category] + ': ' + newPrompt[promptInfo.index.Source] + '\n\n';

      // Create an overview event for the last writing prompt
      var overviewTitle = lastEvent.getTitle().replace(RegExp('^' + latestEventPrefix), 'Final: ');
      var allParts = [];

      for (var i = 1; i < participantInfo.data.length; i++) {
        allParts.push([participantInfo.data[i][partEmailIdx]]);
      }

      createEventAndNewRow(overviewTitle, lastEvent.getDescription(), nextStartTime, allParts.join(','), true);
    } else {
      var newPrefix = getTitlePrefix(promptId, newNumber);
      scriptInfo.data[scriptLength][currNumberIdx] = newNumber;
      title = lastEvent.getTitle().replace(RegExp('^' + latestEventPrefix), newPrefix);
      text = lastEvent.getDescription() + '\n------------------------\n\n';
    }

    createEventAndNewRow(title, text, nextStartTime, guest);
  }

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
   * Incrementally gets only updated tasts
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
          } else {
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
      submissionInfo.note[currIdx][textIdx] += ('\n===================\n' + new Date().toLocaleString() + ' overwrote:\n' + currRow[textIdx] + '\n');
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
            body: 'Prompt:\n\n' + lastEvent.getTitle() + '\n==================\n' +
                  eventDescription +
                  '\n\nNew Count: ' + wordsWrote +
                  '\n\nTotal Count: ' + currRow[wordsIdx] +
                  '\n\n---------\n' +
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
  function createEventAndNewRow(title, text, startDate, guests, isAllDay) {
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
      endDate.setHours(endDate.getHours() + 3);
      event = writingCalendar.createEvent(title, startDate, endDate, eventOptions);
    }

    event.setGuestsCanModify(true);

    // Now add this new row to "Submission" spreadsheet
    // Get range by row, column, row length, column length
    var newRow = [];
    for (var i = 0; i < submissionInfo.data[0].length;) newRow[i++] = '';

    newRow[submissionInfo.index.ParticipantEmail] = guests;
    newRow[subPromptId] = scriptInfo.data[scriptLength][promptIdIdx];
    newRow[subCurrNumIdx] = scriptInfo.data[scriptLength][currNumberIdx];
    newRow[eventIdIdx] = title;
    newRow[calendarEventIdx] = event.getId();
    newRow[submissionInfo.index.CreatedDate] = new Date();
    newRow[textIdx] = text;
    newRow[wordsIdx] = getWordCount(text);
    lastSubmissionIdx++;
    var cells = submissionInfo.sheet.getRange(lastSubmissionIdx, 1, 1, newRow.length);
    cells.setValues([newRow])

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
