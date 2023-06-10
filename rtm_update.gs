// Expects rtm_settings.rtm_feed url
const rtmContext = {
  taskFormat: /.*?\(Updated: (.+?)\s+\|\|\s+(.+?)\)/,
  dateInfo: {},
  dateIdx: {},
  currYear: new Date().getFullYear(),
  isChanged: false,
  offSet: `GMT+${(-1 * new Date().getTimezoneOffset()) / 60}`,
};
const yearSheetName = 'Overview Year';
function updateWithRTM() {
  const allSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = allSheet.getSheetByName(yearSheetName);
  const eventRange = sheet.getRange(`A2:C`);
  const eventData = eventRange.getValues();
  const eventNotes = eventRange.getNotes();
  rtmContext.eventData = eventData;
  rtmContext.eventNotes = eventNotes;
  const dateInfo = rtmContext.dateInfo;
  eventData.forEach((dayInfo, idx) => {
    const date = dayInfo[0];
    const dateTasks = {};
    const localeTime = convertToLocalTime(date);
    dateInfo[localeTime] = dateTasks;
    rtmContext.dateIdx[localeTime] = idx;
    const tasks = dayInfo[1];
    if (tasks) {
      tasks.split('\n').forEach((task) => {
        const regexpResults = task.match(rtmContext.taskFormat);
        dateTasks[regexpResults[2]] = new Date(regexpResults[1]);
      });
    }
  });

  const xml = UrlFetchApp.fetch(rtm_settings.rtm_feed).getContentText();
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const ns = root.getNamespace();
  const entries = root.getChildren('entry', ns);
  entries.forEach(parseRTMTask);

  const loggingSheet = SpreadsheetApp.openById(
    rtm_settings.logging_sheet_id
  ).getSheetByName(rtm_settings.logging_sheet_name);
  const logData = loggingSheet.getRange(`A2:P`).getValues();
  logData.forEach((data) => {
    const currDate = data[3];
    const currNote = data[15];
    const currIdx =
      currDate && rtmContext.dateIdx[convertToLocalTime(currDate, 'GMT')];

    if (currDate && currNote && currIdx !== undefined) {
      const previousNote = rtmContext.eventData[currIdx][2];
      if (previousNote !== currNote) {
        if (previousNote) {
          rtmContext.eventNotes[
            currIdx
          ][2] = `Updated ${new Date()}\n${previousNote}\n-----------\n`;
        }

        rtmContext.eventData[currIdx][2] = currNote;
      }

      rtmContext.isChanged = true;
    }
  });

  if (rtmContext.isChanged) {
    eventRange.setNotes(eventNotes);
    eventRange.setValues(eventData);
  }
}

function parseRTMTask(entry) {
  const value = entry.getValue();
  const ns = entry.getNamespace();
  let content = entry.getChild('content', ns);
  content = content.getChildren();
  if (!content.length) return;
  const contentChildren = content[0].getChildren();
  let dueDate, note;
  contentChildren.forEach((el) => {
    const elClass = el.getAttribute('class').getValue();
    if (elClass === 'rtm_due') {
      const tagValue = el
        .getChildren()
        .filter(
          (span) => span.getAttribute('class').getValue() === 'rtm_due_value'
        );
      dueDate = tagValue[0].getValue();
    } else if (elClass === 'rtm_notes') {
      note = el;
    }
  });

  const dueDateObj = new Date(dueDate);
  if (isNaN(dueDateObj) || dueDateObj.getFullYear() !== rtmContext.currYear)
    return;

  // 1. Check if task has been recently added or updated
  const id = entry.getChild('id', ns).getValue();
  let updatedDate = entry.getChild('updated', ns).getValue();
  updatedDate = new Date(updatedDate);
  const dueDateString = convertToLocalTime(dueDateObj, 'GMT');
  const lastUpdated = rtmContext.dateInfo[dueDateString][id];

  if (!lastUpdated || updatedDate > lastUpdated) {
    rtmContext.dateInfo[dueDateString][id] = updatedDate;

    const dateIdx = rtmContext.dateIdx[dueDateString];

    // 2a. Replace text in cell
    const title = entry.getChild('title', ns).getValue();
    const taskText = `${title} (Updated: ${updatedDate.toISOString()} || ${id})`;
    if (!lastUpdated) {
      const separator = rtmContext.eventData[dateIdx][1] ? '\n' : '';
      rtmContext.eventData[dateIdx][1] += `${separator}${taskText}`;
    } else {
      const oldText = rtmContext.eventData[dateIdx][1];
      const escapedId = id.replaceAll(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
      rtmContext.eventData[dateIdx][1] = oldText.replace(
        new RegExp(`.*? || ${escapedId}`),
        taskText
      );

      if (oldText === rtmContext.eventData[dateIdx]) {
        throw new Error(
          `No changes in text: ${oldText}, ID: ${id}, escapedID: ${escapedId}, lastUpdated ${lastUpdated.toString()}, updatedDate ${updatedDate.toString()}`
        );
      }
    }

    rtmContext.isChanged = true;

    // 2b. Add new notes to note. Assume notes are only updated, not edited (ie, edits will be saved separately)
    const currNote = rtmContext.eventNotes[dateIdx][1];
    if (!note) return;
    note.getChildren().forEach((rtm_note) => {
      let noteContent, noteUpdated;
      rtm_note.getChildren().forEach((note_container) => {
        const containerClass = note_container.getAttribute('class').getValue();
        if (containerClass === 'rtm_note_content') {
          noteContent = note_container.getValue();
        } else if (containerClass === 'rtm_note_updated_container') {
          noteUpdated = note_container.getValue();
        }
      });

      if (!currNote.includes(noteUpdated)) {
        rtmContext.eventNotes[
          dateIdx
        ][1] = `${noteUpdated}\n${noteContent}\n----------\n${rtmContext.eventNotes[dateIdx][1]}`;
      }
    });

    rtmContext.isChanged = true;
  }
}

function convertToLocalTime(dateObj, locale = rtmContext.offSet) {
  return Utilities.formatDate(dateObj, locale, 'yyyy-MM-dd');
}
