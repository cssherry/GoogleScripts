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
  nextCell.setValue(
    `${nextCell.getValue()} (${appendText}: ${new Date().toLocaleString()})`
  );
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
  const weekString = Utilities.formatDate(new Date(), 'GMT', 'w');
  const weekNum = parseInt(weekString);
  const startRow = 26 * (weekNum - 1) + 2;
  const currWeekRange = sheet.getRange(
    `A${26 * weekNum + 3}:B${26 * weekNum + 22 + 3}`
  );
  const currWeekValues = currWeekRange.getValues();

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

  currWeekRange.setValues(currWeekValues);

  sendReport(
    incompleteNum,
    incompleteReport,
    taskSummaryRangeValues[3][1],
    taskSummaryRangeValues[3][2],
    allSheet
  );
}


function formatBullets(list) {
  const listItems = list.split('\n').map((text) => `<li>${text.replace('- ', '')}</li>`).join('\n');
  return `<ul>${listItems}</ul>`;
}

  // Returns back correct styling given actual versus expected number
function getColorStyle (percentage) {
  if (percentage > .95) {
    styleColor = 'darkgreen';
  } else if (percentage > .85) {
    styleColor = 'green';
  } else if (percentage > .60) {
    styleColor = 'orange';
  } else {
    styleColor = 'mediumvioletred';
  }

  return 'style="color:' + styleColor + ';"';
}

function sendReport(
  incompleteNum,
  incompleteReport,
  completedNumber,
  completedItems,
  allSheet
) {
  const incompleteText = Object.keys(incompleteReport)
    .map((keyVal) => `<h2>${keyVal}:</h2>\n${formatBullets(incompleteReport[keyVal])}`)
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

  let habitText = '';
  let habitCompleted = 0;
  const currDate = new Date().getDate();
  monthHabitRange.forEach((row, idx) => {
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

    habitText += `<h2>${
      row[0]
    }\n</h2>Completed: <b ${getColorStyle(currHabitComplete / minKeepOnTrack)}>${currHabitComplete.toFixed(2)}</b><br/>Needed: <b>${(
      minKeepOnTrack
    ).toFixed(2)}</b>`;
  });

  MailApp.sendEmail({
    to: myEmail,
    subject: `[Habit + Goals] Completed ${completedNumber} | Incomplete ${incompleteNum} | ${habitCompleted} Habits Completed (${new Date().toLocaleString()})`,
    htmlBody:
      `<h1>Habits</h1>${habitText}\n\n\n<h1>Goals</h1>${incompleteText}\n\n<h2>COMPLETED:</h2>\n${formatBullets(completedItems)}` +
      `<p><em>Link: ${excelLink}</em></p>`,
  });
}
