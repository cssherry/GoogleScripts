function sendAbbreviatedSummaryEmail() {
  sendSummaryEmail(false, GLOBALS_VARIABLES.myEmailsAbb);
}

function sendSummaryEmail(includeNotes = true, emails = GLOBALS_VARIABLES.summaryEmails) {
  const summaryDate = getPastDate(1);
  const summaryDateFormatted = parseDate(summaryDate);

  const currDayData = [];

  const allSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = allSheet.getSheetByName('Tracker');
  const currRange = sheet.getDataRange();
  const data = currRange.getValues();
  const dataIdx = indexSheet(data);

  const dateIdx = dataIdx.DateWithTimezone;
  const actionIdx = dataIdx.Action;
  const sleepTimeIdx = dataIdx.TotalTime;
  const descriptionIdx = dataIdx.Note;

  // Number feeding, poo, pee, sleep
  const dailyNums = {
    feed: 0,
    poo: 0,
    pee: 0,
    sleep: 0,
  }

  data.forEach((dataRow, idx) => {
    if (idx === 0) return;
    const rowDate = parseDate(dataRow[dateIdx]);
    if (rowDate === summaryDateFormatted) {
      const dateTime = dataRow[dateIdx];
      const action = dataRow[actionIdx];
      const description = dataRow[descriptionIdx];

      if (includeNotes || action !== 'Note') {
        currDayData.push({
          date: dateTime,
          action,
          value: description,
        });
      }

      // Add to feed count
      if (action === 'Famly.Daycare:MealRegistration' || action === 'Feed') {
        dailyNums.feed += 1;
      }

      // Add to poo and pee count
      if (description === 'Ishaan: Nappy Change - Wet') {
        dailyNums.pee += 1;
      }

      else if (description === 'Ishaan: Nappy Change - BM') {
        dailyNums.poo += 1;
        dailyNums.pee += 1;
      }


      if (action === 'Pee/Poo') {
        let isLogged = false;
        if (description.match(/\bpoo\b/i)) {
          isLogged = true;
          dailyNums.poo += 1
        }

        if (description.match(/\bpee\b/i)) {
          isLogged = true;
          dailyNums.pee += 1
        }

        if (!isLogged) {
          dailyNums.pee += 1;
        }
      }


      // Add to sleep
      if (dataRow[sleepTimeIdx]) {
        dailyNums.sleep += dataRow[sleepTimeIdx];
      }
    }
  });

  dailyNums.sleep = (dailyNums.sleep / 60).toFixed(2);

  let currDaySummary = `Average Values for ${summaryDate}:\n`;

  for (let dailyItem in dailyNums) {
    currDaySummary += `${dailyItem}: ${dailyNums[dailyItem]}\n`;
  }

  currDaySummary += lineSeparators;

  currDayData.sort((day1, day2) => day1.date - day2.date);
  const currDayText = currDayData.map((day) => `${day.date}:\n${day.action}\n${day.value}`).join(lineSeparators);

  // Add weekly summary on Monday
  // Include weight, height, head, number of baths (and whether he needs to be re-measured)
  const weekNum = summaryDate.getDay();
  const includeSummary = weekNum === 1;
  const summaryText = includeSummary && includeNotes
    ? calculateSummary({ allSheet })
    : calculateSummary({
      allSheet,
      numDays: includeSummary ? 7 : 1,
      headerText: 'Summary from overview tab:\n',
      includeNotes: false,
    });

  const weeklyText = includeSummary ? ' & Weekly' : '';
  MailApp.sendEmail({
    to: emails.join(','),
    subject: `[Famly] Daily${weeklyText} Summary ${summaryDateFormatted}`,
    body: summaryText + currDaySummary + currDayText,
  });
}

function calculateSummary({ allSheet, numDays = 7, headerText = 'Weekly Averages and summaries:\n', includeNotes = true }) {
  let summaryDateEnd = getPastDate(0);
  let summaryDateStart = getPastDate(numDays);

  allSheet = allSheet || SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = allSheet.getSheetByName('overview w/ notes');
  const summaryRange = summarySheet.getDataRange();
  const summaryData = summaryRange.getValues();

  // Store all data
  const weeklyNums = {
    feed: '',
    poo: '',
    pee: '',
    sleep: '',

    weight: '',
    height: '',
    head: '',
    bath: 0,
  }

  // Get indexs for all data
  const summaryIdx = indexSheet(summaryData);

  const weeklyIdx = {
    feed: summaryIdx['# feeding'],
    poo: summaryIdx['# poos'],
    pee: summaryIdx['# pees'],
    sleep: summaryIdx['Sleep (only after Tracker)'],
  };

  const measurementIdx = {
    weight: summaryIdx.WeightMeasured,
    head: summaryIdx.HeadMeasured,
    height: summaryIdx.HeightMeasured,
  }

  const summaryDateIdx = summaryIdx.Date;
  const bathIdx = summaryIdx.Bath;
  const generalNoteIdx = summaryIdx['General Day Notes'];

  const includedDatesNotes = [];

  // Start adding info from the week
  summaryData.forEach((chronologicalData, idx) => {
    if (idx === 0) return;
    const sumDate = chronologicalData[summaryDateIdx];
    if (sumDate && sumDate < summaryDateEnd && sumDate >= summaryDateStart) {
      // Add note if it exists
      const note = chronologicalData[generalNoteIdx];
      if (note && includeNotes) {
        includedDatesNotes.push(`${chronologicalData[summaryDateIdx]}\n${note}`);
      }

      // Add measurements if they exist
      for (let measIdx in measurementIdx) {
        const val = chronologicalData[measIdx];
        if (!isNaN(val)) {
          weeklyNums[measIdx] = appendValue(weeklyNums[measIdx], val);
        }
      }

      // Add bath if it exists
      const bath = chronologicalData[bathIdx];
      if (bath) {
        weeklyNums.bath += 1;
      }

      // Add generic measurements from the day
      for (let sumIdx in weeklyIdx) {
        weeklyNums[sumIdx] = appendValue(weeklyNums[sumIdx], chronologicalData[weeklyIdx[sumIdx]]);
      }
    }
  });

  let weeklyAvg = '';
  for (let avgType in weeklyNums) {
    weeklyAvg += `${avgType}: ${weeklyNums[avgType]}\n`;
  }

  // Add summary data to top
  let summaryNotes = includedDatesNotes.filter((dateNote) => !!dateNote).join(lineSeparators);
  summaryNotes = summaryNotes ? `${lineSeparators}${summaryNotes}\n\n=============================\n\n` : summaryNotes;
  return `${headerText}${weeklyAvg}${summaryNotes}`;
}

// HELPER FUNCTIONS
function appendValue(previousString, newString, joiningString = ' -> ') {
  const intermediate = previousString ? joiningString : '';
  return `${previousString}${intermediate}${newString}`;
}

function getPastDate(daysAgo) {
  let newDate = new Date();
  newDate.setDate(newDate.getDate() - daysAgo);
  newDate.setHours(0, 0, 0, 0)
  return newDate;
}