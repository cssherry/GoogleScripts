const weeklyPlanningSheetName = 'Week Planner';
const habitsSheetName = 'Habits Tracker';
const notStarted = 'Not Started';

function appendToCellOnTheRight(spreadsheet, currentRange, appendText) {
  const currentColumn = currentRange.getColumn();
  const currentRow = currentRange.getRow();
  const nextCell = spreadsheet.getRange(
    parseInt(currentRow),
    currentColumn + 1
  );

  const richText = nextCell.getRichTextValue();
  const richTextBuilder = SpreadsheetApp.newRichTextValue();
  richTextBuilder.setText(`${richText.getText()} (${appendText}: ${new Date().toLocaleString()})`);

  let totalIdx = 0;
  richText.getRuns().map((currRichText) => {
    const url = currRichText.getLinkUrl();
    const text = currRichText.getText();
    if (url) {
      richTextBuilder.setLinkUrl(totalIdx, totalIdx + text.length, url);
    }

    totalIdx += text.length;
  });

  const richTextNew = richTextBuilder.build();

  nextCell.setRichTextValue(richTextNew);
}

function onEdit(e) {
  const allSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = allSheet.getSheetByName(weeklyPlanningSheetName);

  if (e.source.getSheetName() !== weeklyPlanningSheetName) return;

  if (e.value === '✅') {
    appendToCellOnTheRight(sheet, e.range, `Completed`);
  }

  if (e.value === 'In Progress') {
    appendToCellOnTheRight(sheet, e.range, `Started`);
  }

  if (e.value === '💬') {
    appendToCellOnTheRight(sheet, e.range, `Waiting`);
  }
}

function addIncompleteItems() {
  const allSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = allSheet.getSheetByName(weeklyPlanningSheetName);
  const weekNum = getCurrWeek();
  const currWeekRange = sheet.getRange(
    `A${26 * weekNum + 3}:B${26 * weekNum + 22 + 3}`
  );
  const currWeekValues = currWeekRange.getValues();

  const startRow = 26 * (weekNum - 1) + 2;
  const taskSummaryRange = sheet.getRange(`L${startRow}:N${startRow + 4}`);
  const taskSummaryRangeValues = taskSummaryRange.getValues();

  let currIdx = 0;
  const incompleteReport = {};
  let incompleteNum = 0;

  for (let rowNum = 0; rowNum < 3; rowNum++) {
    const incompleteItemValue = taskSummaryRangeValues[rowNum][2];
    if (incompleteItemValue && incompleteItemValue !== '#N/A') {
      const sign = taskSummaryRangeValues[rowNum][0];
      incompleteReport[sign] = incompleteItemValue;

      incompleteItemValue.split('\n').forEach((item) => {
        incompleteNum++;
        while (
          currIdx < currWeekValues.length &&
          (currWeekValues[currIdx][1] || currWeekValues[currIdx][0])
        ) {
          currIdx++;
        }

        if (currIdx < currWeekValues.length) {
          if (sign !== notStarted) {
            currWeekValues[currIdx][0] = sign;
          }

          currWeekValues[currIdx][1] = item;
        } else {
          console.log('Unable to find spot for: ', item);
        }
      });
    }
  }

  const preservedStyles = currWeekRange.getTextStyles();
  currWeekRange.setValues(currWeekValues);
  currWeekRange.setTextStyles(preservedStyles);

  const notes = sheet.getRange(`F${26 * (weekNum - 1) + 19}`).getValue();

  sendReport(
    incompleteNum,
    incompleteReport,
    taskSummaryRangeValues[3][1],
    taskSummaryRangeValues[3][2],
    notes,
    allSheet
  );
}

function sendReport(
  incompleteNum,
  incompleteReport,
  completedNumber,
  completedItems,
  notes,
  allSheet
) {
  function formatListSection(header, list) {
    const listItems = list.split('\n').map((text) => `<li>${text.replace('- ', '').replace(/\((.*)\)$/, '<small style="color:gray;""><em>($1)</em></small>')}</li>`);
    return `<h2>${header} (${listItems.length}):</h2>\n<ul>${listItems.join('\n')}</ul>`;
  }

    // Returns back correct styling given actual versus expected number
  function getColorStyle (percentage) {
    if (percentage >= .90) {
      styleColor = 'darkgreen';
    } else if (percentage >= .80) {
      styleColor = 'green';
    } else if (percentage >= .60) {
      styleColor = 'orange';
    } else {
      styleColor = 'mediumvioletred';
    }

    return 'style="color:' + styleColor + ';"';
  }

  const incompleteText = Object.keys(incompleteReport)
    .map((keyVal) => formatListSection(keyVal, incompleteReport[keyVal]))
    .join('\n\n');
  const sheet = allSheet.getSheetByName(habitsSheetName);

  const monthNum = new Date().getMonth();
  const startRow = 20 * monthNum + 3;
  const endRow = 20 * monthNum + 8;
  const monthHabitRange = sheet
    .getRange(`I${startRow}:AN${endRow}`)
    .getValues();
  const monthGoalRange = sheet
    .getRange(`AS${startRow}:AS${endRow}`)
    .getValues();

  let habitCompleted = 0;
  const currDate = new Date().getDate();
  const habitFormatted = monthHabitRange.map((row, idx) => {
    let currHabitComplete = 0;
    let currHabitIncomplete = 0;
    for (let dayIdx = 1; dayIdx < row.length; dayIdx++) {
      const dayInfo = row[dayIdx];
      if (dayInfo === true) {
        currHabitComplete += 1;
        habitCompleted += 1;
      } else if (dayInfo === false) {
        currHabitIncomplete += 1;
      }
    }

    const percentageOfGoal =
      monthGoalRange[idx][0] / (currHabitComplete + currHabitIncomplete);
    const minKeepOnTrack = currDate * percentageOfGoal
    const completedPercentage = currHabitComplete / minKeepOnTrack;

    return {
      text: `<h2>${
            row[0]
          }\n</h2>Completed: <b ${getColorStyle(completedPercentage)}>${currHabitComplete.toFixed(2)}</b><br/>Needed: <b>${(
            minKeepOnTrack
          ).toFixed(2)}</b>`,
      completedPercentage,
    };
  });

  const habitText = habitFormatted.sort((a, b) => a.completedPercentage - b.completedPercentage).map((item) => item.text).join('\n');
  const percentageCompletedGoals = completedNumber / (incompleteNum + completedNumber);

  MailApp.sendEmail({
    to: myEmail,
    subject: `[Habit + Goals] Completed ${completedNumber} | Incomplete ${incompleteNum} | ${habitCompleted} Habits Completed (${new Date().toLocaleString()})`,
    htmlBody:
      `<h1>Habits</h1>${habitText}\n\n\n<h1>Goals <small><em ${getColorStyle(percentageCompletedGoals)}>(${(percentageCompletedGoals * 100).toFixed(1)}%)</em></small></h1>${incompleteText}\n\n${formatListSection('COMPLETED', completedItems)}<h1>NOTES: </h1>${notes.replaceAll('\n', '<br/>') || 'None'}` +
      `<p><em>Link: ${excelLink}</em></p>`,
  });
}

function getCurrWeek() {
  const weekString = Utilities.formatDate(new Date(), 'GMT', 'w');
  return parseInt(weekString);
}
