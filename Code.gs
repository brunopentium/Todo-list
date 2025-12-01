const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('TASKS_SPREADSHEET_ID');
const TASKS_SHEET_NAME = 'Tarefas';
const SETTINGS_KEY = 'APP_SETTINGS_JSON';

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
  const id = SPREADSHEET_ID;
  if (!id) {
    throw new Error('Defina a propriedade de script TASKS_SPREADSHEET_ID com o ID da planilha alvo.');
  }
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName(TASKS_SHEET_NAME) || ss.insertSheet(TASKS_SHEET_NAME);
  const header = ['id', 'title', 'date', 'deadline', 'status', 'priority', 'difficulty', 'project', 'notes', 'tags', 'subtasks', 'recurrence', 'alertLevel', 'metadata'];
  const currentHeader = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  const needsHeader = currentHeader.join('') === '' || header.some((h, i) => currentHeader[i] !== h);
  if (needsHeader) {
    sheet.clear();
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
  return sheet;
}

function listData() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues() : [];
  const tasks = values.map(row => ({
    id: row[0],
    title: row[1],
    date: row[2],
    deadline: row[3],
    status: row[4],
    priority: row[5],
    difficulty: row[6],
    project: row[7],
    notes: row[8],
    tags: row[9] ? row[9].split(',').map(s => s.trim()).filter(Boolean) : [],
    subtasks: row[10] ? JSON.parse(row[10]) : [],
    recurrence: row[11] ? JSON.parse(row[11]) : null,
    alertLevel: row[12],
    metadata: row[13] ? JSON.parse(row[13]) : {},
  }));

  const settingsJson = PropertiesService.getScriptProperties().getProperty(SETTINGS_KEY) || '{}';
  const settings = JSON.parse(settingsJson);

  return { tasks, settings };
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

function resetStorage() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  PropertiesService.getScriptProperties().deleteProperty(SETTINGS_KEY);
  return true;
}
