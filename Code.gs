const TASKS_SPREADSHEET_PROPERTY = 'TASKS_SPREADSHEET_ID';
const TASKS_SHEET_NAME = 'Tarefas';
const SETTINGS_KEY = 'APP_SETTINGS_JSON';

function getSpreadsheetId_() {
  const props = PropertiesService.getScriptProperties();
  const legacyKeys = ['SPREADSHEET_ID', 'SPREADSHEETID', 'SHEET_ID'];
  const storedId = props.getProperty(TASKS_SPREADSHEET_PROPERTY);
  if (storedId) return storedId;

  for (const key of legacyKeys) {
    const legacyId = props.getProperty(key);
    if (legacyId) {
      props.setProperty(TASKS_SPREADSHEET_PROPERTY, legacyId);
      return legacyId;
    }
  }

  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet) {
    const activeId = activeSpreadsheet.getId();
    props.setProperty(TASKS_SPREADSHEET_PROPERTY, activeId);
    return activeId;
  }

  throw new Error('Defina a propriedade de script TASKS_SPREADSHEET_ID com o ID da planilha alvo.');
}

function doGet() {
  return HtmlService.createTemplateFromFile('IndexG')
    .evaluate()
    .setTitle('Todo List - Apps Script')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet_() {
  const id = getSpreadsheetId_();
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName(TASKS_SHEET_NAME) || ss.insertSheet(TASKS_SHEET_NAME);
  const header = ['id', 'title', 'date', 'deadline', 'status', 'priority', 'difficulty', 'project', 'notes', 'tags', 'subtasks',
    'recurrence', 'alertLevel', 'metadata'];
  const headerRange = sheet.getRange(1, 1, 1, header.length);
  const currentHeader = headerRange.getValues()[0];
  const hasHeader = currentHeader.some(value => value !== '');
  const headerMismatch = header.some((h, i) => currentHeader[i] !== h);

  if (!hasHeader || headerMismatch) {
    headerRange.setValues([header]);
  }
  return sheet;
}

function listData() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues() : [];

  const parseJsonSafe = (value, fallback) => {
    if (!value) return fallback;
    try {
      return JSON.parse(value);
    } catch (err) {
      return fallback;
    }
  };

  const formatDateCell = (value) => {
    if (value instanceof Date) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    return value || '';
  };

  const tasks = values.map(row => ({
    id: row[0],
    title: row[1] || '',
    date: formatDateCell(row[2]),
    deadline: formatDateCell(row[3]),
    status: row[4] || 'Pendente',
    priority: row[5],
    difficulty: row[6],
    project: row[7] || '',
    notes: row[8] || '',
    tags: row[9] ? row[9].split(',').map(s => s.trim()).filter(Boolean) : [],
    subtasks: parseJsonSafe(row[10], []),
    recurrence: parseJsonSafe(row[11], null),
    alertLevel: row[12],
    metadata: parseJsonSafe(row[13], {}),
  }));

  const settingsJson = PropertiesService.getScriptProperties().getProperty(SETTINGS_KEY) || '{}';
  const settings = JSON.parse(settingsJson);

  return { tasks, settings };
}

function serializeTask_(task) {
  return [
    task.id,
    task.title || '',
    task.date || '',
    task.deadline || '',
    task.status || 'Pendente',
    task.priority || 1,
    task.difficulty || 1,
    task.project || '',
    task.notes || '',
    (task.tags || []).join(', '),
    JSON.stringify(task.subtasks || []),
    JSON.stringify(task.recurrence || null),
    task.alertLevel || '',
    JSON.stringify(task.metadata || {}),
  ];
}

function writeTasks_(sheet, tasks) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  const serializedTasks = (tasks || []).map(serializeTask_);
  if (serializedTasks.length) {
    sheet.getRange(2, 1, serializedTasks.length, serializedTasks[0].length).setValues(serializedTasks);
  }
}

function saveTask(task) {
  if (!task || !task.id) {
    throw new Error('Uma tarefa válida com ID é necessária.');
  }
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const range = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
  const ids = range.map(r => r[0]);
  const idx = ids.indexOf(task.id);
  const rowIndex = idx >= 0 ? idx + 2 : lastRow + 1;
  const serialized = [
    task.id,
    task.title || '',
    task.date || '',
    task.deadline || '',
    task.status || 'Pendente',
    task.priority || 1,
    task.difficulty || 1,
    task.project || '',
    task.notes || '',
    (task.tags || []).join(', '),
    JSON.stringify(task.subtasks || []),
    JSON.stringify(task.recurrence || null),
    task.alertLevel || '',
    JSON.stringify(task.metadata || {}),
  ];
  sheet.getRange(rowIndex, 1, 1, serialized.length).setValues([serialized]);
  return task;
}

function deleteTask(taskId) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => r[0]);
  const idx = ids.indexOf(taskId);
  if (idx === -1) return false;
  sheet.deleteRow(idx + 2);
  return true;
}

function saveSettings(settings) {
  const json = JSON.stringify(settings || {});
  PropertiesService.getScriptProperties().setProperty(SETTINGS_KEY, json);
  return settings;
}

function syncAllData(payload) {
  const data = payload || {};
  const tasks = Array.isArray(data.tasks) ? data.tasks : [];
  const settings = data.settings || {};

  const sheet = getSheet_();
  const hasExistingTasks = sheet.getLastRow() > 1;

  if (tasks.length === 0 && hasExistingTasks) {
    saveSettings(settings);
    return { savedTasks: 0, skippedOverwrite: true };
  }

  writeTasks_(sheet, tasks);
  saveSettings(settings);

  return { savedTasks: tasks.length };
}

function resetStorage() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  PropertiesService.getScriptProperties().deleteProperty(SETTINGS_KEY);
  return true;
}
