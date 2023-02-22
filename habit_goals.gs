const weeklyPlanningSheetName = 'Week Planner';
const habitsSheetName = 'Habits Tracker';
const notStarted = 'Not Started';

function recreateText(richText, prefix = '', suffix = '') {
  const richTextBuilder = SpreadsheetApp.newRichTextValue();
  richTextBuilder.setText(`${prefix}${richText.getText()}${suffix}`);

  let totalIdx = prefix.length;
  richText.getRuns().map((currRichText) => {
    const url = currRichText.getLinkUrl();
    const text = currRichText.getText();
    if (url) {
      richTextBuilder.setLinkUrl(totalIdx, totalIdx + text.length, url);
    }

    totalIdx += text.length;
  });

  return richTextBuilder.build();
}

function appendToCellOnTheRight(spreadsheet, currentRange, appendText) {
  const currentColumn = currentRange.getColumn();
  const currentRow = currentRange.getRow();
  const nextCell = spreadsheet.getRange(
    parseInt(currentRow),
    currentColumn + 1
  );

  const richText = nextCell.getRichTextValue();
  const richTextNew = recreateText(richText, '', ` (${appendText}: ${new Date().toLocaleString()})`);
  nextCell.setRichTextValue(richTextNew);
}

function getTodos(offset = 0) {
  const allSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = allSheet.getSheetByName(weeklyPlanningSheetName);
  const weekNum = getCurrWeek() + offset;

  return {
    Monday: sheet.getRange(`A${26 * weekNum + 3}:B${26 * weekNum + 14 + 3}`),
    Tuesday: sheet.getRange(`C${26 * weekNum + 3}:D${26 * weekNum + 14 + 3}`),
    Wednesday: sheet.getRange(`E${26 * weekNum + 3}:F${26 * weekNum + 14 + 3}`),
    Thursday: sheet.getRange(`G${26 * weekNum + 3}:H${26 * weekNum + 14 + 3}`),
    Friday: sheet.getRange(`I${26 * weekNum + 3}:J${26 * weekNum + 14 + 3}`),
    Saturday: sheet.getRange(`A${26 * weekNum + 14 + 3 + 2}:B${26 * weekNum + 14 + 3 + 2 + 6}`),
    Sunday: sheet.getRange(`C${26 * weekNum + 14 + 3 + 2}:D${26 * weekNum + 14 + 3 + 2 + 6}`),
  }
}

const statusOrder = ['💬', 'In Progress', '', '✅'];
function sortTasks([iconA, _A], [iconB, _B]) {
  if (_A.getText() === '' && _B.getText() !== '') return 1;
  if (_A.getText() !== '' && _B.getText() === '') return -1;
  const iconTextA = iconA.getText();
  const iconTextB = iconB.getText();
  const valueA = statusOrder.indexOf(iconTextA);
  const valueB = statusOrder.indexOf(iconTextB)
  return valueA - valueB;
}

function reorderEvents() {
  const todos = getTodos()
  for (let day in todos) {
    const richText = todos[day].getRichTextValues()
    richText.sort(sortTasks);
    todos[day].setRichTextValues(richText);
  }
}

function onEdit(e) {
  const allSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = allSheet.getSheetByName(weeklyPlanningSheetName);

  if (!e || !e.range) {
    const range = sheet.getActiveRange();
    e = {
      source: sheet,
      value: range.getValue(),
      range,
    };
  }

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
  const currWeekValues = currWeekRange.getRichTextValues();

  const todos = getTodos(-1)

  let currIdx = 0;
  const incompleteReport = {};
  let incompleteNum = 0;
  let completedNum = 0;
  let completedItems = '';

  for (let day in todos) {
    const richTexts = todos[day].getRichTextValues();
    richTexts.forEach(([sign, task]) => {
      const signText = sign.getText();
      const signTextForReport = signText || notStarted;
      const taskText = task.getText();

      if (signText === '💬' || signText === 'In Progress' || (!signText && !!taskText)) {
        incompleteNum++;
        while (
          currIdx < currWeekValues.length &&
          (currWeekValues[currIdx][1].getText() || currWeekValues[currIdx][0].getText())
        ) {
          currIdx++;
        }

        if (!incompleteReport[signTextForReport]) {
          incompleteReport[signTextForReport] = '';
        }

        incompleteReport[signTextForReport] += `- ${taskText}\n`;

        if (currIdx < currWeekValues.length) {
          currWeekValues[currIdx][0] = sign;
          currWeekValues[currIdx][1] = recreateText(task, '- ');
        } else {
          console.log(`Unable to find spot for: ${signText || notStarted}: ${taskText}`);
        }
      } else if (!!taskText) {
        completedNum += 1;
        completedItems += `- ${taskText}\n`;
      }
    });
  }

  const firstDay = currWeekValues.slice(0, 15);
  const weekend = currWeekValues.slice(16)
  firstDay.sort(sortTasks);
  weekend.sort(sortTasks);

  const newWeek = [...firstDay, currWeekValues[15], ...weekend];
  currWeekRange.setRichTextValues(newWeek);

  const notes = sheet.getRange(`F${26 * (weekNum - 1) + 19}`).getValue();

  sendReport(
    incompleteNum,
    incompleteReport,
    completedNum,
    completedItems,
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
    const listItems = list.split('\n').map((text) => text ? `<li>${text.replace('- ', '').replace(/\((.*)\)$/, '<small style="color:gray;""><em>($1)</em></small>')}</li>` : '');
    return `<h2>${header} (${listItems.length}):</h2>\n<ol>${listItems.join('\n')}</ol>`;
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
  const startRow = 10 * monthNum + 3;
  const endRow = 10 * monthNum + 8;
  const monthHabitRange = sheet
    .getRange(`I${startRow}:AN${endRow}`)
    .getValues();
  const monthGoalRange = sheet
    .getRange(`AU${startRow}:AU${endRow}`)
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
      `<h1>Habits</h1>${habitText}\n\n\n<h1>Goals <small><em ${getColorStyle(percentageCompletedGoals)}>(${(percentageCompletedGoals * 100).toFixed(1)}% | ${timeTakenHours(completedItems, 'avg')} hrs avg)</em></small></h1>${incompleteText}\n\n${formatListSection('COMPLETED', completedItems)}<h1>NOTES: </h1>${notes.replaceAll('\n', '<br/>') || 'None'}` +
      `<p><em>Link: ${excelLink}</em></p>`,
  });
}

function getCurrWeek() {
  const weekString = Utilities.formatDate(new Date(), 'GMT', 'w');
  return parseInt(weekString);
}

// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Manage Habits')
    .addItem('Order Habits', 'reorderEvents')
    .addSeparator()
    .addItem('Add current status time', 'onEdit')
    .addToUi();
}


// FUNCTIONS
// Converts text in format "1/6/2023, 3:00:35 PM" to date
function convertToDate(text) {
  return new Date(text.replace(/:\d+\s(PM|AM)/, ' $1'));
}

function timeTakenHours(text, operation) {
  let sum = 0;
  let count = 0;
  text.split('\n').forEach((currText) => {
    const strippedText = currText.trim();
    if (!strippedText.length) return;

    if (operation === 'count') {
      count += 1;
    } else {
      // Return
      // 0: "(Waiting: 1/6/2023, 2:11:27 PM) (Completed: 1/6/2023, 3:00:35 PM)"
      // 1: "1/6/2023, 2:11:27 PM"
      // 2: ", 2:11:27 PM"
      // 3: "PM"
      // 4: "1/6/2023, 3:00:35 PM"
      // 5: ", 3:00:35 PM"
      // 6: "PM"
      const timesMatch = strippedText.match(/\(.+?: (\d+\/\d+\/\d+(,\s+\d+:\d+:\d+\s+(PM|AM))?).*?(\d+\/\d+\/\d+(,\s+\d+:\d+:\d+\s+(PM|AM))?)?\)$/);
      if (!timesMatch || !timesMatch[4]) return;
      const timeEnd = convertToDate(timesMatch[4]);
      const timeStart = convertToDate(timesMatch[1]);
      const diff = timeEnd - timeStart;

      sum += diff;
      count += 1;
    }
  })

  switch (operation) {
    case 'avg':
      return sum / count / 60 / 60 / 1000;
    case 'sum':
      return sum / 60 / 60 / 1000;
    default:
      return count;
  }
}