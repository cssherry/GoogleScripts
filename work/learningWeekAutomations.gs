const currentSheetName = 'FY25_Form';
const settingSheetName = 'Settings';
const teamCalendar = 'Viz Team Calendar';
// const teamCalendar = 'test';
const tilManager = 'amedgyesy@cloudera.com';
const manager = 'szhou@cloudera.com';

// Add to calendar + calculate quarter
function onFormSubmit() {
  const settingSheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName(settingSheetName);
  const settingSheetData = settingSheet.getDataRange().getValues();
  const settingSheetIndex = indexSheet(settingSheetData);

  let hasChanged = false;

  // Get sheet
  const currFormSheet = SpreadsheetApp.getActiveSpreadsheet()
                                      .getSheetByName(currentSheetName);
  const scheduleSheetRange = currFormSheet.getDataRange()
  const scheduleSheetData = scheduleSheetRange.getValues();
  const scheduleSheetIndex = indexSheet(scheduleSheetData);

  // Get info for quarter calculation
  const quarters = parseQuarters(settingSheetData, settingSheetIndex.Quarter);
  
  // Get info for calendar update
  const eventIdIdx = scheduleSheetIndex['Calendar Event'];
  const calendar = CalendarApp.getCalendarsByName(teamCalendar)[0];
  const emailToName = parseEmails(settingSheetData, settingSheetIndex);

  // Go through for every row
  scheduleSheetData.forEach((row, idx) => {
    if (!row[0] || idx === 0) return;

    // Update quarter by end date
    const endDateIdx = scheduleSheetIndex['End Date'];
    const endDate = row[endDateIdx];

    const quarterIdx = scheduleSheetIndex.Quarter;
    const quarter = calculateQuarter(endDate, quarters);
    isChanged = false;
    if (row[quarterIdx] !== quarter) {
      isChanged = true;
      row[quarterIdx] = quarter;
      hasChanged = true;
    }

    // Add calendar event
    const calendarID = addCalendarEvent(scheduleSheetIndex, emailToName, row, calendar, eventIdIdx, isChanged);

    if (calendarID) {
      row[eventIdIdx] = calendarID;
      hasChanged = true;
    }
  });

  // Update spreadsheet
  if (hasChanged) {
    scheduleSheetRange.setValues(scheduleSheetData);
  }
 }

function addCalendarEvent(scheduleSheetIndex, emailToName, values, calendar, eventIdIdx, isChange) {
  const eventId = values[eventIdIdx];
    if (eventId && !isChange) return;

  const startDateIdx = scheduleSheetIndex['Start Date'];
  const endDateIdx = scheduleSheetIndex['End Date'];
  const startDate = values[startDateIdx];
  const endDate = values[endDateIdx];

  if (eventId) {
    const event = calendar.getEventById(eventId);
    const calendarStart = event.getAllDayStartDate();
    const calendarEnd = event.getAllDayEndDate();

    if (isSameDay(startDate, calendarStart) && isSameDay(endDate, calendarEnd)) return;
    event.setAllDayDates(startDate, endDate)
    return;
  }

  const emailIdx = scheduleSheetIndex['Email Address'];


  const topicIdx = scheduleSheetIndex['What is the topic of the learning week?'];
  const lessonsLearnedIdx = scheduleSheetIndex['To help share learning from the week, make a copy of https://docs.google.com/document/d/1_rtoKe1N2ov5QBMf6Msj8ghd1yltgrDwUzmz1TSIcW4/edit#heading=h.7cnulwrfpy0m for the "Training" folder and share here'];
  const resourceIdx = scheduleSheetIndex['Link to resource(s)'];

  const guests = values[emailIdx];
  const name = emailToName[guests];
  const topic = values[topicIdx];
  const beginningTopic = topic.split(' ').slice(0, 5).join(' ');
  const title = `${name}'s Learning Week: ${beginningTopic}${beginningTopic !== topic ? '...' : ''}`;
  const description = `Topic:\ntopic\n\nResources:\n${values[resourceIdx]}\n\nLink to learnings doc: ${values[lessonsLearnedIdx]}`;

  let event;
  if (isSameDay(startDate, endDate)) {
    event = calendar.createAllDayEvent(title, startDate, {description, guests, sendInvites: true});
  } else {
    event = calendar.createAllDayEvent(title, startDate, endDate, {description, guests, sendInvites: true});
  }
  
  return event.getId();
}


// Every week, remind if there's no TIL scheduled
function everyWeek() {
    const currFormSheet = SpreadsheetApp.getActiveSpreadsheet()
                                        .getSheetByName(currentSheetName);

    const scheduleSheetData = currFormSheet.getDataRange().getValues();
    const scheduleSheetIndex = indexSheet(scheduleSheetData);
    const addedDateIdx = scheduleSheetIndex.Timestamp;
    const endDateIdx = scheduleSheetIndex['End Date'];
    const presentedToTeamIdx = scheduleSheetIndex['Presented to Team'];
    const now = Date.now();
    const sevenDaysAgo = new Date(now - 7 * 24 * 60 * 60 * 1000);

    const lessonsLearnedIdx = scheduleSheetIndex['To help share learning from the week, make a copy of https://docs.google.com/document/d/1_rtoKe1N2ov5QBMf6Msj8ghd1yltgrDwUzmz1TSIcW4/edit#heading=h.7cnulwrfpy0m for the "Training" folder and share here'];
    const emailIdx = scheduleSheetIndex['Email Address'];
    const topicIdx = scheduleSheetIndex['What is the topic of the learning week?'];
    const resourceIdx = scheduleSheetIndex['Link to resource(s)'];

    let toScheduleCount = 0;
    let toSchedule = '';
    let newRowsCount = 0;
    let newRows = '';
    scheduleSheetData.forEach((scheduledLearning) => {
      const addedDate = scheduledLearning[addedDateIdx];
      const endDate = scheduledLearning[endDateIdx];
      const presentedToTeam = scheduledLearning[presentedToTeamIdx];

      const email = scheduledLearning[emailIdx];
      const topic = scheduledLearning[topicIdx];
      const resource = scheduledLearning[resourceIdx];
      const lessonsLearned = scheduledLearning[lessonsLearnedIdx];
      
      const rowText = `${email}: ${topic}<br>${resource}<br>Link: ${lessonsLearned}<br><br>`;
      if (addedDate && addedDate > sevenDaysAgo) {
        newRowsCount += 1;
        newRows += `${newRowsCount}) ${rowText}`;
      }

      if (endDate < now && !presentedToTeam) {
        toScheduleCount += 1;
        toSchedule += `${toScheduleCount}) ${rowText}`;
      }
    });

    const toBeScheduledText = toScheduleCount ? `<h1>${toScheduleCount} TILs To Be Scheduled</h1>${toSchedule}<br><br>` : '';
    const newLearningWeekText = newRowsCount ? `<h1>${newRowsCount} New Learning Weeks Planned</h1>${newRows}<br><br>` : '';

    if (!toBeScheduledText && !newLearningWeekText) return;

    MailApp.sendEmail({
    to: `${tilManager}, ${manager}`,
    subject: `Learning Week Update (${toScheduleCount} To Schedule + ${newRowsCount} New)`,
    htmlBody:
      toBeScheduledText +
      newLearningWeekText +
      `<p><em>Link: <a href="https://docs.google.com/spreadsheets/d/1E1XofPrQMsWf5AbEgMaVdk0OCJUTURF-3jQrmUN1nIY/edit?gid=1171670265#gid=1171670265">TIL Sheet</a></em></p>`,
  });
}

// TODO: Every first Monday of the month, remind people to schedule learning week
function everyMonth() {

}