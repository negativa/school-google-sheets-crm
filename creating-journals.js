/**
 * Refactored Google Apps Script for creating subject journal sheets.
 *
 * Main idea:
 * - Subjects come from the "Предмети" sheet.
 * - Students come from the "Students list" sheet.
 * - Student block templates come from the "Template" sheet by subject Group.
 * - "Маппінг предметів" is not used.
 */

const CONFIG = {
  SHEETS: {
    TEMPLATE: 'Template',
    STUDENTS: 'Students list',
    SUBJECTS: 'Предмети',
    LOG: 'LOG_ЖУРНАЛИ'
  },

  UI: {
    MENU: 'Журнали',
    CREATE: 'Створити вкладки з Предмети',
    RECREATE: 'Пересоздати вкладки з Предмети'
  },

  TEMPLATE: {
    TITLE_TEXT: 'Назва предмету',
    TOTAL_LABEL: 'Загальна оцінка',
    CLEAR_GRADE_VALUES: true,

    MARKERS: {
      STUDENT_ID: 'STUDENT_ID',
      CLASS: 'CLASS',
      NAME: 'NAME',
      DESCRIPTION: 'DESCRIPTION'
    }
  },

  HEADERS: {
    SUBJECTS: {
      JOURNAL_SUBJECT: [
        'назва предмету ua',
        'предмет в журналі',
        'предмет'
      ],
      CERTIFICATE_SUBJECT: [
        'назва предмету в свідоцтві',
        'предмет в свідоцтві'
      ],
      GROUP: [
        'group',
        'група',
        'освітня галузь'
      ]
    },

    STUDENTS: {
      ID: ['id'],
      CLASS: ['клас'],
      NAME: [
        'імʼя прізвище',
        'ім’я прізвище',
        'імя прізвище',
        'name'
      ]
    }
  }
};

const STATUS = {
  CREATED: 'СТВОРЕНО',
  SKIPPED: 'ПРОПУЩЕНО',
  ERROR: 'ПОМИЛКА',
  STACK: 'STACK'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.UI.MENU)
    .addItem(CONFIG.UI.CREATE, 'createSubjectSheetsFromSubjectsList')
    .addItem(CONFIG.UI.RECREATE, 'recreateSubjectSheetsFromSubjectsList')
    .addToUi();
}

function createSubjectSheetsFromSubjectsList() {
  JournalCreationApp.run({ recreateExisting: false });
}

function recreateSubjectSheetsFromSubjectsList() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Пересоздати вкладки?',
    'Будуть пересоздані тільки вкладки, які є у списку "Предмети". Оцінки на цих вкладках будуть стерті.',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  JournalCreationApp.run({ recreateExisting: true });
}

const JournalCreationApp = {
  run(options) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = new LogWriter();

    try {
      const context = loadContext(ss);
      const result = createJournalSheets(context, options, log);

      log.writeTo(ss);

      alertUser(
        'Готово.\n' +
        'Створено вкладок: ' + result.created + '\n' +
        'Пропущено: ' + result.skipped + '\n' +
        'Деталі дивись у вкладці ' + CONFIG.SHEETS.LOG
      );
    } catch (error) {
      log.error(error);
      log.writeTo(ss);

      alertUser(
        'Помилка: ' + error.message + '\n\n' +
        'Деталі дивись у вкладці ' + CONFIG.SHEETS.LOG
      );

      throw error;
    }
  }
};

function loadContext(ss) {
  const sheets = {
    template: getRequiredSheet(ss, CONFIG.SHEETS.TEMPLATE),
    students: getRequiredSheet(ss, CONFIG.SHEETS.STUDENTS),
    subjects: getRequiredSheet(ss, CONFIG.SHEETS.SUBJECTS)
  };

  const data = {
    students: StudentRepository.read(sheets.students),
    subjects: SubjectRepository.read(sheets.subjects),
    template: TemplateRepository.analyze(sheets.template)
  };

  assertNotEmpty(data.students, 'Не знайдено учнів у вкладці "' + CONFIG.SHEETS.STUDENTS + '".');
  assertNotEmpty(data.subjects, 'Не знайдено предметів у вкладці "' + CONFIG.SHEETS.SUBJECTS + '".');

  return { ss, sheets, data };
}

function createJournalSheets(context, options, log) {
  let created = 0;
  let skipped = 0;

  context.data.subjects.forEach(subject => {
    const result = createOneJournalSheet(context, subject, options, log);

    if (result.created) created++;
    if (result.skipped) skipped++;
  });

  return { created, skipped };
}

function createOneJournalSheet(context, subject, options, log) {
  const sheetName = makeSafeSheetName(subject.journalSubject);

  if (isSystemSheetName(sheetName)) {
    log.row(subject.journalSubject, STATUS.SKIPPED, 'Назва збігається зі службовою вкладкою');
    return { skipped: true };
  }

  const existingSheet = findSheetByName(context.ss, sheetName);

  if (existingSheet && !options.recreateExisting) {
    log.row(subject.journalSubject, STATUS.SKIPPED, 'Вкладка вже існує');
    return { skipped: true };
  }

  const templateBlock = TemplateRepository.findByGroup(context.data.template, subject.group);

  if (!templateBlock) {
    log.row(
      subject.journalSubject,
      STATUS.SKIPPED,
      'Не знайдено шаблон у Template для Group:',
      subject.group
    );

    return { skipped: true };
  }

  if (existingSheet && options.recreateExisting) {
    context.ss.deleteSheet(existingSheet);
  }

  const targetSheet = context.sheets.template.copyTo(context.ss);
  targetSheet.setName(sheetName);

  JournalSheetBuilder.build({
    sourceTemplateSheet: context.sheets.template,
    targetSheet,
    subject,
    students: context.data.students,
    templateInfo: context.data.template,
    templateBlock
  });

  log.row(
    subject.journalSubject,
    STATUS.CREATED,
    'Group: ' + subject.group,
    'Висота блоку учня: ' + templateBlock.blockHeight
  );

  return { created: true };
}

const StudentRepository = {
  read(sheet) {
    const values = sheet.getDataRange().getDisplayValues();
    const headerRowIndex = Header.findRow(values, [
      CONFIG.HEADERS.STUDENTS.ID,
      CONFIG.HEADERS.STUDENTS.CLASS,
      CONFIG.HEADERS.STUDENTS.NAME
    ]);

    const header = values[headerRowIndex];

    const cols = {
      id: Header.findCol(header, CONFIG.HEADERS.STUDENTS.ID),
      className: Header.findCol(header, CONFIG.HEADERS.STUDENTS.CLASS),
      name: Header.findCol(header, CONFIG.HEADERS.STUDENTS.NAME)
    };

    return values
      .slice(headerRowIndex + 1)
      .map(row => ({
        id: text(row[cols.id]),
        className: text(row[cols.className]),
        name: text(row[cols.name])
      }))
      .filter(student => student.id && student.className && student.name)
      .filter(student => /^\d+$/.test(student.id));
  }
};

const SubjectRepository = {
  read(sheet) {
    const values = sheet.getDataRange().getDisplayValues();
    const headerRowIndex = Header.findRow(values, [
      CONFIG.HEADERS.SUBJECTS.JOURNAL_SUBJECT,
      CONFIG.HEADERS.SUBJECTS.GROUP
    ]);

    const header = values[headerRowIndex];

    const cols = {
      journalSubject: Header.findCol(header, CONFIG.HEADERS.SUBJECTS.JOURNAL_SUBJECT),
      certificateSubject: Header.findColOptional(header, CONFIG.HEADERS.SUBJECTS.CERTIFICATE_SUBJECT),
      group: Header.findCol(header, CONFIG.HEADERS.SUBJECTS.GROUP)
    };

    const subjects = [];
    const seen = new Set();

    values.slice(headerRowIndex + 1).forEach(row => {
      const journalSubject = text(row[cols.journalSubject]);
      const group = text(row[cols.group]);

      if (!journalSubject || !group) return;

      const key = normalize(journalSubject);
      if (seen.has(key)) return;

      seen.add(key);

      const certificateSubject = cols.certificateSubject === -1
        ? journalSubject
        : text(row[cols.certificateSubject]) || journalSubject;

      subjects.push({ journalSubject, certificateSubject, group });
    });

    return subjects;
  }
};

const TemplateRepository = {
  analyze(sheet) {
    const values = sheet.getDataRange().getDisplayValues();
    const markerRowIndex = findTemplateMarkerRow(values);

    if (markerRowIndex === -1) {
      throw new Error('На вкладці Template не знайдено рядок з кодами STUDENT_ID, CLASS, NAME, DESCRIPTION.');
    }

    const markerRow = values[markerRowIndex];
    const markerCols = {
      studentId: Header.findCol(markerRow, [CONFIG.TEMPLATE.MARKERS.STUDENT_ID]),
      className: Header.findCol(markerRow, [CONFIG.TEMPLATE.MARKERS.CLASS]),
      name: Header.findCol(markerRow, [CONFIG.TEMPLATE.MARKERS.NAME]),
      description: Header.findCol(markerRow, [CONFIG.TEMPLATE.MARKERS.DESCRIPTION])
    };

    const templateBlocks = detectTemplateBlocks(values, markerRowIndex, markerCols.description);

    if (templateBlocks.length === 0) {
      throw new Error('У Template не знайдено жодного шаблону учня після рядка "Загальна оцінка".');
    }

    return {
      markerRow: markerRowIndex + 1,
      lastCol: sheet.getLastColumn(),
      cols: toSheetCols(markerCols),
      blocks: templateBlocks
    };
  },

  findByGroup(templateInfo, groupName) {
    const groupKey = normalize(groupName);
    if (!groupKey) return null;

    const exact = templateInfo.blocks.find(block => normalize(block.groupName) === groupKey);
    if (exact) return exact;

    return templateInfo.blocks.find(block => {
      const blockGroup = normalize(block.groupName);
      return blockGroup.includes(groupKey) || groupKey.includes(blockGroup);
    }) || null;
  }
};

const JournalSheetBuilder = {
  build({ sourceTemplateSheet, targetSheet, subject, students, templateInfo, templateBlock }) {
    replaceSubjectTitle(targetSheet, subject.journalSubject);
    replaceTopGroupTitle(targetSheet, subject.group, templateInfo.markerRow);

    const firstStudentRow = templateInfo.markerRow + 1;
    const blockHeight = templateBlock.blockHeight;
    const requiredRows = templateInfo.markerRow + students.length * blockHeight;

    ensureSheetRows(targetSheet, requiredRows);
    clearRowsBelow(targetSheet, templateInfo.markerRow, templateInfo.lastCol);

    if (students.length === 0) return;

    this.copyFirstStudentBlock({
      sourceTemplateSheet,
      targetSheet,
      templateBlock,
      targetStartRow: firstStudentRow,
      templateInfo
    });

    this.setStudentData(targetSheet, firstStudentRow, blockHeight, templateInfo, students[0]);

    const firstBlockRange = targetSheet.getRange(
      firstStudentRow,
      1,
      blockHeight,
      templateInfo.lastCol
    );

    for (let i = 1; i < students.length; i++) {
      const targetStartRow = firstStudentRow + i * blockHeight;

      firstBlockRange.copyTo(
        targetSheet.getRange(targetStartRow, 1, blockHeight, templateInfo.lastCol),
        { contentsOnly: false }
      );

      copyRowHeights(sourceTemplateSheet, targetSheet, templateBlock.blockStartRow, targetStartRow, blockHeight);

      this.setStudentData(targetSheet, targetStartRow, blockHeight, templateInfo, students[i]);
    }

    targetSheet.setFrozenRows(templateInfo.markerRow);
  },

  copyFirstStudentBlock({ sourceTemplateSheet, targetSheet, templateBlock, targetStartRow, templateInfo }) {
    const sourceRange = sourceTemplateSheet.getRange(
      templateBlock.blockStartRow,
      1,
      templateBlock.blockHeight,
      templateInfo.lastCol
    );

    const targetRange = targetSheet.getRange(
      targetStartRow,
      1,
      templateBlock.blockHeight,
      templateInfo.lastCol
    );

    sourceRange.copyTo(targetRange, { contentsOnly: false });

    copyRowHeights(
      sourceTemplateSheet,
      targetSheet,
      templateBlock.blockStartRow,
      targetStartRow,
      templateBlock.blockHeight
    );

    if (CONFIG.TEMPLATE.CLEAR_GRADE_VALUES) {
      clearEditableValues(targetSheet, targetStartRow, templateBlock.blockHeight, templateInfo);
    }
  },

  setStudentData(sheet, startRow, blockHeight, templateInfo, student) {
    setMergedValue(sheet, startRow, blockHeight, templateInfo.cols.studentId, student.id);
    setMergedValue(sheet, startRow, blockHeight, templateInfo.cols.className, student.className);
    setMergedValue(sheet, startRow, blockHeight, templateInfo.cols.name, student.name);
  }
};

const Header = {
  findRow(values, requiredHeaderGroups) {
    for (let r = 0; r < values.length; r++) {
      const normalizedRow = values[r].map(normalize);

      const hasAllHeaders = requiredHeaderGroups.every(group => {
        return this.findColInNormalizedRow(normalizedRow, group) !== -1;
      });

      if (hasAllHeaders) return r;
    }

    throw new Error('Не знайдено рядок із потрібними заголовками.');
  },

  findCol(row, aliases) {
    const normalizedRow = row.map(normalize);
    const col = this.findColInNormalizedRow(normalizedRow, aliases);

    if (col === -1) {
      throw new Error('Не знайдено колонку: ' + aliases.join(' / '));
    }

    return col;
  },

  findColOptional(row, aliases) {
    return this.findColInNormalizedRow(row.map(normalize), aliases);
  },

  findColInNormalizedRow(normalizedRow, aliases) {
    const normalizedAliases = aliases.map(normalize);

    for (let c = 0; c < normalizedRow.length; c++) {
      const cell = normalizedRow[c];

      if (normalizedAliases.some(alias => cell === alias || cell.includes(alias))) {
        return c;
      }
    }

    return -1;
  }
};

function detectTemplateBlocks(values, markerRowIndex, descriptionColIndex) {
  const blocks = [];
  const usedGroups = new Set();

  for (let r = markerRowIndex + 1; r < values.length; r++) {
    const description = normalize(values[r][descriptionColIndex]);

    if (description !== normalize(CONFIG.TEMPLATE.TOTAL_LABEL)) continue;

    const groupRowIndex = findGroupHeaderAboveTotalRow(
      values,
      markerRowIndex,
      r,
      descriptionColIndex
    );

    if (groupRowIndex === -1) continue;

    const groupName = getFirstNonEmptyCell(values[groupRowIndex]);
    const groupKey = normalize(groupName);

    if (!groupName || usedGroups.has(groupKey)) continue;

    usedGroups.add(groupKey);

    const blockEndRowIndex = detectStudentBlockEndRow(
      values,
      r,
      descriptionColIndex
    );

    blocks.push({
      groupName,
      groupRow: groupRowIndex + 1,
      groupLabelCol: getFirstNonEmptyCol(values[groupRowIndex]) + 1,
      blockStartRow: r + 1,
      blockEndRow: blockEndRowIndex + 1,
      blockHeight: blockEndRowIndex - r + 1
    });
  }

  return blocks;
}

function findGroupHeaderAboveTotalRow(values, markerRowIndex, totalRowIndex, descriptionColIndex) {
  for (let r = totalRowIndex - 1; r > markerRowIndex; r--) {
    const rowText = normalize(values[r].join(' '));
    const description = normalize(values[r][descriptionColIndex]);

    if (!rowText) continue;
    if (description) continue;
    if (rowText.includes('student_id')) continue;
    if (rowText.includes('description')) continue;
    if (rowText.includes(normalize(CONFIG.TEMPLATE.TITLE_TEXT))) continue;

    return r;
  }

  return -1;
}

function detectStudentBlockEndRow(values, totalRowIndex, descriptionColIndex) {
  let endRow = totalRowIndex;

  for (let r = totalRowIndex; r < values.length; r++) {
    const description = text(values[r][descriptionColIndex]);
    const rowText = normalize(values[r].join(' '));

    if (!description) break;

    if (r > totalRowIndex && rowText.includes('student_id')) break;
    if (r > totalRowIndex && rowText.includes('освітня галузь')) break;

    endRow = r;
  }

  return endRow;
}

function findTemplateMarkerRow(values) {
  const requiredMarkers = [
    [CONFIG.TEMPLATE.MARKERS.STUDENT_ID],
    [CONFIG.TEMPLATE.MARKERS.CLASS],
    [CONFIG.TEMPLATE.MARKERS.NAME],
    [CONFIG.TEMPLATE.MARKERS.DESCRIPTION]
  ];

  try {
    return Header.findRow(values, requiredMarkers);
  } catch (error) {
    return -1;
  }
}

function replaceSubjectTitle(sheet, subjectName) {
  replaceTextInUsedRange(sheet, CONFIG.TEMPLATE.TITLE_TEXT, subjectName);
}

function replaceTopGroupTitle(sheet, groupName, markerRow) {
  if (!groupName) return;

  const topRows = Math.max(markerRow - 1, 1);
  const range = sheet.getRange(1, 1, topRows, sheet.getLastColumn());
  const values = range.getDisplayValues();

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = normalize(values[r][c]);

      if (cell.includes('освітня галузь')) {
        sheet.getRange(r + 1, c + 1).setValue(groupName);
        return;
      }
    }
  }

  sheet.getRange(1, 1).setValue(groupName);
}

function replaceTextInUsedRange(sheet, searchText, replacement) {
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  const needle = normalize(searchText);

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (normalize(values[r][c]) === needle) {
        sheet.getRange(r + 1, c + 1).setValue(replacement);
        return true;
      }
    }
  }

  return false;
}

function ensureSheetRows(sheet, requiredRows) {
  const currentRows = sheet.getMaxRows();

  if (currentRows < requiredRows) {
    sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
    return;
  }

  if (currentRows > requiredRows) {
    sheet.deleteRows(requiredRows + 1, currentRows - requiredRows);
  }
}

function clearRowsBelow(sheet, markerRow, lastCol) {
  const rowsToClear = sheet.getMaxRows() - markerRow;

  if (rowsToClear <= 0) return;

  const range = sheet.getRange(markerRow + 1, 1, rowsToClear, lastCol);

  range.breakApart();
  range.clearContent();
  range.clearFormat();
  range.clearDataValidations();
}

function clearEditableValues(sheet, startRow, blockHeight, templateInfo) {
  const range = sheet.getRange(startRow, 1, blockHeight, templateInfo.lastCol);

  range.breakApart();

  const values = range.getValues();
  const formulas = range.getFormulas();
  const protectedCols = new Set([
    templateInfo.cols.studentId,
    templateInfo.cols.className,
    templateInfo.cols.name,
    templateInfo.cols.description
  ]);

  const clearedValues = values.map((row, r) => {
    return row.map((cell, c) => {
      const col = c + 1;

      if (formulas[r][c]) return formulas[r][c];
      if (protectedCols.has(col)) return cell;

      return '';
    });
  });

  range.setValues(clearedValues);
}

function setMergedValue(sheet, startRow, blockHeight, col, value) {
  const range = sheet.getRange(startRow, col, blockHeight, 1);

  range.breakApart();

  if (blockHeight > 1) {
    range.mergeVertically();
  }

  range
    .setValue(value)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
}

function copyRowHeights(sourceSheet, targetSheet, sourceStartRow, targetStartRow, rowCount) {
  for (let i = 0; i < rowCount; i++) {
    targetSheet.setRowHeight(
      targetStartRow + i,
      sourceSheet.getRowHeight(sourceStartRow + i)
    );
  }
}

class LogWriter {
  constructor() {
    this.rows = [];
  }

  row(...values) {
    this.rows.push(values);
  }

  error(error) {
    this.row(STATUS.ERROR, error.message);
    this.row(STATUS.STACK, error.stack || '');
  }

  writeTo(spreadsheet) {
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.LOG);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.SHEETS.LOG);
    }

    sheet.clear();

    if (this.rows.length === 0) {
      sheet.getRange(1, 1).setValue('Лог порожній');
      return;
    }

    const width = Math.max(...this.rows.map(row => row.length), 1);
    const prepared = this.rows.map(row => padRow(row, width));

    sheet.getRange(1, 1, prepared.length, width).setValues(prepared);
    sheet.autoResizeColumns(1, width);
  }
}

function getRequiredSheet(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Не знайдено вкладку: ' + sheetName);
  }

  return sheet;
}

function findSheetByName(spreadsheet, sheetName) {
  const exact = spreadsheet.getSheetByName(sheetName);

  if (exact) return exact;

  const needle = normalize(sheetName);

  return spreadsheet
    .getSheets()
    .find(sheet => normalize(sheet.getName()) === needle) || null;
}

function isSystemSheetName(sheetName) {
  const systemSheetNames = [
    CONFIG.SHEETS.TEMPLATE,
    CONFIG.SHEETS.STUDENTS,
    CONFIG.SHEETS.SUBJECTS,
    CONFIG.SHEETS.LOG,
    'Маппінг предметів'
  ];

  return systemSheetNames.some(name => normalize(name) === normalize(sheetName));
}

function makeSafeSheetName(name) {
  return text(name)
    .replace(/[\\/?*\[\]:]/g, '-')
    .substring(0, 99);
}

function toSheetCols(zeroBasedCols) {
  const result = {};

  Object.keys(zeroBasedCols).forEach(key => {
    result[key] = zeroBasedCols[key] + 1;
  });

  return result;
}

function assertNotEmpty(items, message) {
  if (!items || items.length === 0) {
    throw new Error(message);
  }
}

function alertUser(message) {
  SpreadsheetApp.getUi().alert(message);
}

function padRow(row, width) {
  const result = row.slice();

  while (result.length < width) {
    result.push('');
  }

  return result;
}

function getFirstNonEmptyCell(row) {
  for (let c = 0; c < row.length; c++) {
    const value = text(row[c]);

    if (value) return value;
  }

  return '';
}

function getFirstNonEmptyCol(row) {
  for (let c = 0; c < row.length; c++) {
    if (text(row[c])) return c;
  }

  return 0;
}

function text(value) {
  return String(value || '').trim();
}

function normalize(value) {
  return text(value)
    .toLowerCase()
    .replace(/[’ʼ`´]/g, "'")
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, ' ');
}
