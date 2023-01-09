const weeklyPlanningSheetName = 'Week Planner';
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
