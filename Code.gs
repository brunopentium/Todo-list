/**
 * Script principal do aplicativo de tarefas.
 * Usa uma planilha do Google Sheets como base de dados.
 */

const SHEET_NAME = 'Tarefas';
const HEADER_ROW = [
  'ID',
  'Título',
  'Descrição',
  'Projeto',
  'Status',
  'Prioridade',
  'Esforço',
  'Data FUP',
  'Data Limite',
  'Tipo Recorrência',
  'Configuração Recorrência',
  'Observações',
  'Criado em',
  'Atualizado em',
  'Próxima Ocorrência'
];

const STATUS = {
  ACTIVE: 'Em execução',
  DONE: 'Concluída',
  CANCELED: 'Cancelada',
  RECURRING: 'Recorrente'
};

const PRIORITY_ORDER = ['Alta', 'Média', 'Baixa'];

/**
 * Executado quando a planilha é aberta.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestor de Tarefas')
    .addItem('Adicionar tarefa', 'showTaskSidebar')
    .addItem('Aplicar filtros', 'showFilterSidebar')
    .addItem('Classificar (padrão)', 'applyDefaultSort')
    .addItem('Atualizar recorrências', 'processRecurringTasks')
    .addToUi();

  setupSheet_();
}

/**
 * Garante que a planilha de dados exista e possua o cabeçalho esperado.
 */
function setupSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  const currentHeader = sheet.getRange(1, 1, 1, HEADER_ROW.length).getValues()[0];
  const needsHeader = currentHeader.some((cell, index) => cell !== HEADER_ROW[index]);
  if (needsHeader) {
    sheet.clear();
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
    sheet.setFrozenRows(1);
  }
}

/**
 * Abre o formulário de inclusão/edição de tarefas.
 */
function showTaskSidebar(taskId) {
  setupSheet_();
  const template = HtmlService.createTemplateFromFile('TaskForm');
  template.projects = getDistinctValues_(4);
  template.defaultStatus = STATUS.ACTIVE;
  template.task = taskId ? getTaskById_(taskId) : null;
  SpreadsheetApp.getUi()
    .showSidebar(template.evaluate().setTitle(taskId ? 'Editar tarefa' : 'Nova tarefa'));
}

/**
 * Abre a barra lateral de filtros.
 */
function showFilterSidebar() {
  setupSheet_();
  const template = HtmlService.createTemplateFromFile('FilterSidebar');
  template.projects = getDistinctValues_(4);
  template.statuses = [STATUS.ACTIVE, STATUS.DONE, STATUS.CANCELED, STATUS.RECURRING];
  template.priorities = PRIORITY_ORDER;
  SpreadsheetApp.getUi().showSidebar(template.evaluate().setTitle('Filtros'));
}

/**
 * Retorna os valores distintos de uma coluna.
 */
function getDistinctValues_(column) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  const set = new Set();
  values.forEach(row => {
    const value = row[0];
    if (value !== '' && value !== null) {
      set.add(value);
    }
  });
  return Array.from(set).sort();
}

/**
 * Recupera uma tarefa pelo ID.
 */
function getTaskById_(taskId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), HEADER_ROW.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(taskId)) {
      const task = {};
      HEADER_ROW.forEach((key, index) => {
        task[key] = rows[i][index];
      });
      task.rowNumber = i + 2;
      return task;
    }
  }
  return null;
}

/**
 * Cria ou atualiza uma tarefa.
 */
function submitTask(formObject) {
  setupSheet_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  const now = new Date();
  const id = formObject.id && formObject.id !== '' ? formObject.id : generateId_();

  const fupDate = formObject.fupDate ? new Date(formObject.fupDate) : '';
  const deadline = formObject.deadline ? new Date(formObject.deadline) : '';
  const recurrenceType = formObject.status === STATUS.RECURRING ? formObject.recurrenceType : '';
  const recurrenceConfig = formObject.status === STATUS.RECURRING ? buildRecurrenceConfig_(formObject) : '';
  const effort = formObject.effort ? Number(formObject.effort) : '';

  let createdAt = now;
  let nextOccurrence = '';
  if (formObject.rowNumber) {
    const existingRow = sheet.getRange(Number(formObject.rowNumber), 1, 1, HEADER_ROW.length).getValues()[0];
    createdAt = existingRow[12] instanceof Date && !isNaN(existingRow[12].getTime()) ? existingRow[12] : existingRow[12];
  }

  if (formObject.status === STATUS.RECURRING) {
    nextOccurrence = calculateNextOccurrence_(fupDate, deadline, recurrenceType, recurrenceConfig);
  }

  const rowValues = [
    id,
    formObject.title,
    formObject.description,
    formObject.project,
    formObject.status,
    formObject.priority,
    effort,
    fupDate,
    deadline,
    recurrenceType,
    recurrenceConfig,
    formObject.notes,
    createdAt,
    now,
    nextOccurrence
  ];

  if (formObject.rowNumber) {
    sheet.getRange(Number(formObject.rowNumber), 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  applyDefaultSort();
  SpreadsheetApp.getActive().toast('Tarefa salva com sucesso!', 'Gestor de Tarefas', 3);
  return id;
}

/**
 * Constrói a configuração de recorrência com base no formulário.
 */
function buildRecurrenceConfig_(formObject) {
  if (formObject.recurrenceType === 'Semanal') {
    const days = Array.isArray(formObject.recurrenceDays)
      ? formObject.recurrenceDays
      : formObject.recurrenceDays
      ? [formObject.recurrenceDays]
      : [];
    const frequency = Number(formObject.weeklyInterval || 1);
    return JSON.stringify({ type: 'weekly', days, frequency });
  }
  if (formObject.recurrenceType === 'Mensal') {
    const mode = formObject.monthlyMode || 'sameDay';
    const day = formObject.monthlyDay ? Number(formObject.monthlyDay) : null;
    return JSON.stringify({ type: 'monthly', mode, day });
  }
  if (formObject.recurrenceType === 'Diária') {
    const interval = Number(formObject.dailyInterval || 1);
    return JSON.stringify({ type: 'daily', interval });
  }
  return '';
}

/**
 * Calcula a próxima ocorrência para uma tarefa recorrente.
 */
function calculateNextOccurrence_(fupDate, deadline, recurrenceType, recurrenceConfig) {
  if (!(fupDate instanceof Date) || isNaN(fupDate.getTime())) {
    return '';
  }
  if (!recurrenceType || !recurrenceConfig) {
    return '';
  }
  try {
    const config = JSON.parse(recurrenceConfig);
    let nextDate = new Date(fupDate.getTime());
    if (config.type === 'daily') {
      const interval = Math.max(1, Number(config.interval || 1));
      nextDate = addDays_(nextDate, interval);
    } else if (config.type === 'weekly') {
      const days = (config.days || []).map(Number).filter(d => !isNaN(d));
      const frequency = Math.max(1, Number(config.frequency || 1));
      nextDate = getNextWeeklyOccurrence_(nextDate, days, frequency);
    } else if (config.type === 'monthly') {
      nextDate = getNextMonthlyOccurrence_(nextDate, config);
    }
    return nextDate;
  } catch (err) {
    return '';
  }
}

function addDays_(date, days) {
  const newDate = new Date(date.getTime());
  newDate.setDate(newDate.getDate() + days);
  return newDate;
}

function getNextWeeklyOccurrence_(startDate, days, frequency) {
  if (!days.length) {
    return addDays_(startDate, 7 * frequency);
  }
  const sortedDays = days.sort((a, b) => a - b);
  const startDay = startDate.getDay();
  for (let i = 0; i < sortedDays.length; i++) {
    const day = sortedDays[i];
    if (day > startDay) {
      return addDays_(startDate, day - startDay);
    }
  }
  const firstDay = sortedDays[0];
  const daysUntilNext = (7 * frequency) - (startDay - firstDay);
  return addDays_(startDate, daysUntilNext);
}

function getNextMonthlyOccurrence_(startDate, config) {
  const next = new Date(startDate.getTime());
  next.setMonth(next.getMonth() + 1);
  if (config.mode === 'fixedDay' && config.day) {
    const day = Math.min(config.day, daysInMonth_(next.getFullYear(), next.getMonth()));
    next.setDate(day);
  } else {
    const originalDay = startDate.getDate();
    const maxDay = daysInMonth_(next.getFullYear(), next.getMonth());
    next.setDate(Math.min(originalDay, maxDay));
  }
  return next;
}

function daysInMonth_(year, month) {
  return new Date(year, month + 1, 0).getDate();
}

/**
 * Aplica a classificação e filtros padrão.
 */
function applyDefaultSort() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  const range = sheet.getRange(2, 1, lastRow - 1, HEADER_ROW.length);
  const data = range.getValues();
  data.sort(function (a, b) {
    const projectA = (a[3] || '').toString().toLowerCase();
    const projectB = (b[3] || '').toString().toLowerCase();
    if (projectA < projectB) return -1;
    if (projectA > projectB) return 1;

    const dateA = a[7] instanceof Date && !isNaN(a[7].getTime()) ? a[7].getTime() : Number.MAX_SAFE_INTEGER;
    const dateB = b[7] instanceof Date && !isNaN(b[7].getTime()) ? b[7].getTime() : Number.MAX_SAFE_INTEGER;
    if (dateA !== dateB) {
      return dateA - dateB;
    }

    const priorityDiff = getPriorityRank_(a[5]) - getPriorityRank_(b[5]);
    if (priorityDiff !== 0) {
      return priorityDiff;
    }

    const effortA = isNaN(Number(a[6])) ? Number.MAX_SAFE_INTEGER : Number(a[6]);
    const effortB = isNaN(Number(b[6])) ? Number.MAX_SAFE_INTEGER : Number(b[6]);
    return effortA - effortB;
  });
  range.setValues(data);

  const filterRange = sheet.getRange(1, 1, lastRow, HEADER_ROW.length);
  let filter = sheet.getFilter();
  if (!filter) {
    filter = filterRange.createFilter();
  } else {
    filter.remove();
    filter = filterRange.createFilter();
  }
  const excludedStatuses = [STATUS.DONE, STATUS.CANCELED];
  filter.setColumnFilterCriteria(5, SpreadsheetApp.newFilterCriteria().setHiddenValues(excludedStatuses).build());
}

function getPriorityRank_(priority) {
  const index = PRIORITY_ORDER.indexOf(priority);
  return index === -1 ? PRIORITY_ORDER.length : index;
}

/**
 * Aplica filtros personalizados.
 */
function applyFilters(filters) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }
  const filterRange = sheet.getRange(1, 1, lastRow, HEADER_ROW.length);
  let filter = sheet.getFilter();
  if (!filter) {
    filter = filterRange.createFilter();
  }
  if (filters.reset) {
    filter.remove();
    applyDefaultSort();
    return;
  }

  filter.remove();
  filter = filterRange.createFilter();

  if (filters.project) {
    filter.setColumnFilterCriteria(4, SpreadsheetApp.newFilterCriteria().whenTextEqualTo(filters.project).build());
  }
  if (filters.status && filters.status.length) {
    filter.setColumnFilterCriteria(5, SpreadsheetApp.newFilterCriteria().setVisibleValues(filters.status).build());
  }
  if (filters.priority && filters.priority.length) {
    filter.setColumnFilterCriteria(6, SpreadsheetApp.newFilterCriteria().setVisibleValues(filters.priority).build());
  }
  if (filters.effortMin || filters.effortMax) {
    const min = filters.effortMin ? Number(filters.effortMin) : null;
    const max = filters.effortMax ? Number(filters.effortMax) : null;
    const criteriaBuilder = SpreadsheetApp.newFilterCriteria();
    if (min !== null && max !== null) {
      criteriaBuilder.whenNumberBetween(min, max);
    } else if (min !== null) {
      criteriaBuilder.whenNumberGreaterThanOrEqualTo(min);
    } else if (max !== null) {
      criteriaBuilder.whenNumberLessThanOrEqualTo(max);
    }
    filter.setColumnFilterCriteria(7, criteriaBuilder.build());
  }
  if (filters.fupStart || filters.fupEnd) {
    const start = filters.fupStart ? new Date(filters.fupStart) : null;
    const end = filters.fupEnd ? new Date(filters.fupEnd) : null;
    const dateCriteria = SpreadsheetApp.newFilterCriteria();
    if (start && end) {
      dateCriteria.whenDateBetween(start, end);
    } else if (start) {
      dateCriteria.whenDateOnOrAfter(start);
    } else if (end) {
      dateCriteria.whenDateOnOrBefore(end);
    }
    filter.setColumnFilterCriteria(8, dateCriteria.build());
  }
}

/**
 * Atualiza as tarefas recorrentes, avançando a próxima ocorrência quando necessário.
 */
function processRecurringTasks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }
  const range = sheet.getRange(2, 1, lastRow - 1, HEADER_ROW.length);
  const values = range.getValues();
  const today = new Date();
  let updated = false;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[4] !== STATUS.RECURRING) {
      continue;
    }
    const fupDate = row[7];
    const deadline = row[8];
    const recurrenceType = row[9];
    const recurrenceConfig = row[10];
    const computedNext = calculateNextOccurrence_(fupDate, deadline, recurrenceType, recurrenceConfig);
    const nextOccurrence = row[14];

    if (!(nextOccurrence instanceof Date) && computedNext) {
      row[14] = computedNext;
      row[13] = new Date();
      values[i] = row;
      updated = true;
      continue;
    }

    if (computedNext && nextOccurrence instanceof Date && nextOccurrence <= today) {
      row[7] = computedNext;
      if (deadline instanceof Date && !isNaN(deadline.getTime())) {
        const diff = deadline.getTime() - fupDate.getTime();
        row[8] = new Date(computedNext.getTime() + diff);
      }
      row[14] = calculateNextOccurrence_(row[7], row[8], recurrenceType, recurrenceConfig);
      row[13] = new Date();
      values[i] = row;
      updated = true;
    }
  }

  if (updated) {
    range.setValues(values);
    applyDefaultSort();
  }
}

/**
 * Gera um identificador simples.
 */
function generateId_() {
  const props = PropertiesService.getDocumentProperties();
  const current = Number(props.getProperty('TASK_ID_COUNTER') || 0) + 1;
  props.setProperty('TASK_ID_COUNTER', String(current));
  return current;
}

/**
 * Inclui arquivos HTML como templates.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Fornece dados básicos para as interfaces.
 */
function getMetadata() {
  return {
    projects: getDistinctValues_(4),
    statuses: [STATUS.ACTIVE, STATUS.DONE, STATUS.CANCELED, STATUS.RECURRING],
    priorities: PRIORITY_ORDER
  };
}
