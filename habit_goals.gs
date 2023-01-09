const weeklyPlanningSheetName = 'Week Planner';
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
    `A${26 * weekNum + 3}:B${26 * weekNum + 23}`
  );
  const currWeekValues = currWeekRange.getValues();
  let currIdx = 0;

  for (let rowNum = startRow; rowNum < startRow + 3; rowNum++) {
    const incompleteItems = sheet.getRange(`N${rowNum}`);
    const incompleteItemValue = incompleteItems.getValue();
    if (incompleteItemValue && incompleteItemValue !== '#N/A') {
      const sign = sheet.getRange(`L${rowNum}`).getValue();

      incompleteItemValue.split('\n').forEach((item) => {
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
}
