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
  const richTextNew = getNewRichTextValue(richText, appendText);
  nextCell.setRichTextValue(richTextNew);
}

function getNewRichTextValue(richText, appendText) {
  return recreateText(richText, '', ` (${appendText}: ${new Date().toLocaleString()})`);
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

const statusOrder = ['💬', 'In Progress', '❗️', '', '✅', '😵'];
function sortTasks([iconA, _A], [iconB, _B]) {
  if (_A.getText() === '' && _B.getText() !== '') return 1;
  if (_A.getText() !== '' && _B.getText() === '') return -1;
  const iconTextA = iconA.getText();
  const iconTextB = iconB.getText();
  const valueA = statusOrder.indexOf(iconTextA);
  const valueB = statusOrder.indexOf(iconTextB)
  return valueA - valueB;
}

function reorderEventsOnSelection() {
    const selection = SpreadsheetApp.getSelection().getCurrentCell();
    const currentCell = selection.getRow();
    const row = currentCell;
    const weekSelected = Math.floor((row - 3) / 26);
    const currWeek = getCurrWeek();
    reorderEvents(weekSelected - currWeek);
}

function reorderEvents(weekOffset = 0) {
  console.log(`Weekoffset is ${weekOffset}`);
  if (isNaN(parseFloat(weekOffset))) {
    weekOffset = 0;
  }

  console.log(`Using offset ${weekOffset}`);
  const todos = getTodos(weekOffset)
  for (let day in todos) {
    const richText = todos[day].getRichTextValues();
    richText.sort(sortTasks);
    todos[day].setRichTextValues(richText);
  }
}

const iconToText = {
  '✅': `Completed`,
  'In Progress': `Started`,
  '💬': `Waiting`,
  '❗️': `On Notice`,
  '😵': `Abandoned`,
};

const iconKeys = Object.keys(iconToText);

function updateStatusTime() {
  const todos = getTodos()
  let isChanged = false;
  for (let day in todos) {
    const richTextTasks = todos[day].getRichTextValues();

    for (let task of richTextTasks) {
      const [icon, taskName] = task;
      const iconText = icon.getText();
      const taskNameText = taskName.getText();
      if (!iconText) continue;
      const regexpTest = new RegExp(`\\(${iconToText[iconText]}: [0-9\\/,:APM\\s]+\\)$`);
      if (!regexpTest.test(taskNameText)) {
        task[1] = getNewRichTextValue(taskName, iconToText[iconText]);
        isChanged = true;
      }
    }

    if (isChanged) {
      todos[day].setRichTextValues(richTextTasks);
    }
  }
}

// Automatically add date where status changed
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

  for (let icon of iconKeys) {
    if (e.value.includes(icon)) {
      appendToCellOnTheRight(sheet, e.range, iconToText[icon]);
      return;
    }
  }
}

function setNewWeekDay(weekRange, weekValues) {
  weekValues.sort(sortTasks);
  weekRange.setRichTextValues(weekValues);
}

// Runs weekly to move incomplete items to next week
function addIncompleteItems() {
  const todos = getTodos(-1)
  const newTodos = getTodos();

  const days = Object.keys(newTodos);
  const weekends = days.splice(5);
  const weekdayRange = newTodos[days[0]];
  const weekendRange = newTodos[weekends[0]];
  const tracker = {
    weekday: {
      days,
      dayIdx: 0,
      todoIdx: 0,
      newDayRange: weekdayRange,
      newDay: weekdayRange.getRichTextValues(),
    },
    weekend: {
      days: weekends,
      dayIdx: 0,
      todoIdx: 0,
      newDayRange: weekendRange,
      newDay: weekendRange.getRichTextValues(),
    }
  };

  const incompleteReport = {};
  let incompleteNum = 0;
  let completedNum = 0;
  let completedItems = '';

  for (let day in todos) {
    let newDayTracker = tracker.weekday;
    if (weekends.includes(day)) {
      newDayTracker = tracker.weekend;
    }

    const richTexts = todos[day].getRichTextValues();
    richTexts.forEach(([sign, task]) => {
      const signText = sign.getText();
      const signTextForReport = signText || notStarted;
      const taskText = task.getText();

      if (signText === '💬' || signText === 'In Progress' || signText === '❗️' || (!signText && !!taskText)) {
        incompleteNum++;
        let currweekText = newDayTracker.newDay[newDayTracker.todoIdx][1].getText();
        let currweekIcon = newDayTracker.newDay[newDayTracker.todoIdx][0].getText();
        let isTooLong = newDayTracker.todoIdx >= newDayTracker.newDay.length;
        while (
          (!isTooLong || newDayTracker.dayIdx < newDayTracker.days.length) &&
          (!!currweekText || !!currweekIcon)
        ) {
          isTooLong = newDayTracker.todoIdx >= newDayTracker.newDay.length - 1;
          if (isTooLong) {
            setNewWeekDay(newDayTracker.newDayRange, newDayTracker.newDay);
            newDayTracker.dayIdx += 1;
            newDayTracker.todoIdx = 0;
            newDayTracker.newDayRange = newTodos[newDayTracker.days[newDayTracker.dayIdx]];
            newDayTracker.newDay = newDayTracker.newDayRange.getRichTextValues();
            isTooLong = false;
          } else {
            newDayTracker.todoIdx++;
          }

          currweekText = newDayTracker.newDay[newDayTracker.todoIdx][1].getText();
          currweekIcon = newDayTracker.newDay[newDayTracker.todoIdx][0].getText();
        }

        if (!incompleteReport[signTextForReport]) {
          incompleteReport[signTextForReport] = '';
        }

        if (!isTooLong) {
          newDayTracker.newDay[newDayTracker.todoIdx][0] = sign;
          newDayTracker.newDay[newDayTracker.todoIdx][1] = recreateText(task, '- ');
        } else {
          console.log(`Unable to find spot for: ${signText || notStarted}: ${taskText}`);
          taskText += ' (UNABLE TO FIND SPOT)';
        }


        incompleteReport[signTextForReport] += `- ${taskText}\n`;
      } else if (!!taskText) {
        completedNum += 1;
        completedItems += `- ${taskText}\n`;
      }
    });
  }

  setNewWeekDay(tracker.weekday.newDayRange, tracker.weekday.newDay);
  setNewWeekDay(tracker.weekend.newDayRange, tracker.weekend.newDay);

  const allSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = allSheet.getSheetByName(weeklyPlanningSheetName);
  const weekNum = getCurrWeek();
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
    const listItems = list.split('\n').filter((text) => !!text).map((text) => text ? `<li>${text.replace('- ', '').replace(/\((.*)\)$/, '<small style="color:gray;""><em>($1)</em></small>')}</li>` : '');
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
  return parseInt(weekString) - 1;
}

// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Manage Habits')
    .addItem('Order Habits', 'reorderEventsOnSelection')
    .addSeparator()
    .addItem('Add current status time', 'onEdit')
    .addItem('Update tasks with time if missing', 'updateStatusTime')
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

// Adds notes from cell groups to cell
function addNotes(cellRange, spreadsheetName, joinText = '\n') {
  const allSheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheet = spreadsheetName ? allSheet.getSheetByName(spreadsheetName) : allSheet.getActiveSheet();
  try {
    const notes = spreadsheet.getRange(cellRange).getNotes();
    return notes.flat().filter(el => !!el).join(joinText)
  } catch (e) {
    const errorMessage = `"${spreadsheetName}"!${cellRange} : Error ${e}`;
    console.error(erro);
    return errorMessage;
  }
}