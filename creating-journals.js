const CONFIG = {
  TEMPLATE_SHEET_NAME: 'Template',
  STUDENTS_SHEET_NAME: 'Students list',
  SUBJECTS_SHEET_NAME: 'Предмети',

  TEMPLATE_TITLE_TEXT: 'Назва предмету',
  TOTAL_LABEL: 'Загальна оцінка',

  LOG_SHEET_NAME: 'LOG_ЖУРНАЛИ',

  CLEAR_GRADE_VALUES: true,

  TEMPLATE_MARKERS: {
    STUDENT_ID: 'STUDENT_ID',
    CLASS: 'CLASS',
    NAME: 'NAME',
    DESCRIPTION: 'DESCRIPTION'
  },

  SUBJECT_HEADERS: {
    CREATE: [
      'створити',
      'створювати',
      'створити журнал',
      'create',
      'checkbox'
    ],

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
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Журнали')
    .addItem('Створити вибрані вкладки', 'createSubjectSheetsFromSubjectsList')
    .addItem('Пересоздати вибрані вкладки', 'recreateSubjectSheetsFromSubjectsList')
    .addToUi();
}

function createSubjectSheetsFromSubjectsList() {
  buildSubjectSheets(false);
}

function recreateSubjectSheetsFromSubjectsList() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Пересоздати вибрані вкладки?',
    'Будуть пересоздані тільки ті вкладки, навпроти яких у списку "Предмети" стоїть галочка. Оцінки на цих вкладках будуть стерті.',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  buildSubjectSheets(true);
}

function buildSubjectSheets(recreateExisting) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  try {
    const templateSheet = getRequiredSheet(ss, CONFIG.TEMPLATE_SHEET_NAME);
    const studentsSheet = getRequiredSheet(ss, CONFIG.STUDENTS_SHEET_NAME);
    const subjectsSheet = getRequiredSheet(ss, CONFIG.SUBJECTS_SHEET_NAME);

    const students = readStudents(studentsSheet);
    const subjects = readSubjectsList(subjectsSheet);
    const templateInfo = analyzeTemplate(templateSheet);

    if (students.length === 0) {
      throw new Error('Не знайдено учнів у вкладці "' + CONFIG.STUDENTS_SHEET_NAME + '".');
    }

    if (subjects.length === 0) {
      throw new Error(
        'Не вибрано жодного предмета. Постав галочку в колонці "Створити" на вкладці "' +
        CONFIG.SUBJECTS_SHEET_NAME +
        '".'
      );
    }

    let created = 0;
    let skipped = 0;

    subjects.forEach(subject => {
      const sheetName = makeSafeSheetName(subject.journalSubject);

      if (isSystemSheetName(sheetName)) {
        log.push([
          subject.journalSubject,
          'ПРОПУЩЕНО',
          'Назва збігається зі службовою вкладкою'
        ]);
        skipped++;
        return;
      }

      const existingSheet = getSheetByNameLoose(ss, sheetName);

      if (existingSheet && !recreateExisting) {
        log.push([
          subject.journalSubject,
          'ПРОПУЩЕНО',
          'Вкладка вже існує'
        ]);
        skipped++;
        return;
      }

      if (existingSheet && recreateExisting) {
        ss.deleteSheet(existingSheet);
      }

      const groupTemplate = findTemplateByGroup(templateInfo, subject.group);

      if (!groupTemplate) {
        log.push([
          subject.journalSubject,
          'ПРОПУЩЕНО',
          'Не знайдено шаблон у Template для Group:',
          subject.group
        ]);
        skipped++;
        return;
      }

      const newSheet = templateSheet.copyTo(ss);
      newSheet.setName(sheetName);

      fillSubjectSheetFromTemplateBlock(
        templateSheet,
        newSheet,
        subject,
        students,
        templateInfo,
        groupTemplate
      );

      created++;

      log.push([
        subject.journalSubject,
        'СТВОРЕНО',
        'Group: ' + subject.group,
        'Висота блоку учня: ' + groupTemplate.blockHeight
      ]);
    });

    writeLog(ss, log);

    SpreadsheetApp.getUi().alert(
      'Готово.\n' +
      'Створено вкладок: ' + created + '\n' +
      'Пропущено: ' + skipped + '\n' +
      'Деталі дивись у вкладці ' + CONFIG.LOG_SHEET_NAME
    );

  } catch (error) {
    log.push(['ПОМИЛКА', error.message]);
    log.push(['STACK', error.stack || '']);

    writeLog(ss, log);

    SpreadsheetApp.getUi().alert(
      'Помилка: ' + error.message + '\n\n' +
      'Деталі дивись у вкладці ' + CONFIG.LOG_SHEET_NAME
    );

    throw error;
  }
}

function fillSubjectSheetFromTemplateBlock(
  templateSheet,
  sheet,
  subject,
  students,
  templateInfo,
  groupTemplate
) {
  replaceSubjectTitle(sheet, subject.journalSubject);

  // Освітня галузь показується вгорі над назвою предмета
  replaceTopGroupTitle(sheet, subject.group, templateInfo.markerRow);

  const markerRow = templateInfo.markerRow;
  const lastCol = templateInfo.lastCol;
  const blockHeight = groupTemplate.blockHeight;

  // В списку учнів освітню галузь НЕ вставляємо
  const totalRowsNeeded = students.length * blockHeight;
  const requiredTotalRows = markerRow + totalRowsNeeded;

  ensureSheetRowCount(sheet, requiredTotalRows);
  clearRowsBelowMarker(sheet, markerRow, lastCol);

  if (students.length === 0) return;

  const firstStudentBlockRow = markerRow + 1;

  // Копіюємо шаблон блоку першого учня з відповідної групи Template
  copyTemplateStudentBlock(
    templateSheet,
    sheet,
    groupTemplate,
    firstStudentBlockRow,
    lastCol
  );

  if (CONFIG.CLEAR_GRADE_VALUES) {
    clearGradeValuesInStudentBlock(
      sheet,
      firstStudentBlockRow,
      blockHeight,
      lastCol,
      templateInfo
    );
  }

  setStudentFieldsOnly(
    sheet,
    firstStudentBlockRow,
    blockHeight,
    templateInfo,
    students[0]
  );

  const firstBlockRange = sheet.getRange(
    firstStudentBlockRow,
    1,
    blockHeight,
    lastCol
  );

  // Далі копіюємо готовий блок першого учня для всіх інших учнів
  for (let i = 1; i < students.length; i++) {
    const targetStartRow = firstStudentBlockRow + i * blockHeight;

    firstBlockRange.copyTo(
      sheet.getRange(targetStartRow, 1, blockHeight, lastCol),
      {
        contentsOnly: false
      }
    );

    setStudentFieldsOnly(
      sheet,
      targetStartRow,
      blockHeight,
      templateInfo,
      students[i]
    );
  }

  sheet.setFrozenRows(markerRow);
}

function ensureSheetRowCount(sheet, requiredTotalRows) {
  const currentRows = sheet.getMaxRows();

  if (currentRows < requiredTotalRows) {
    sheet.insertRowsAfter(
      currentRows,
      requiredTotalRows - currentRows
    );
  }

  if (currentRows > requiredTotalRows) {
    const rowsToDelete = currentRows - requiredTotalRows;

    if (rowsToDelete > 0 && requiredTotalRows > 0) {
      sheet.deleteRows(requiredTotalRows + 1, rowsToDelete);
    }
  }
}

function copyTemplateStudentBlock(templateSheet, targetSheet, groupTemplate, targetStartRow, lastCol) {
  const sourceRange = templateSheet.getRange(
    groupTemplate.blockStartRow,
    1,
    groupTemplate.blockHeight,
    lastCol
  );

  const targetRange = targetSheet.getRange(
    targetStartRow,
    1,
    groupTemplate.blockHeight,
    lastCol
  );

  sourceRange.copyTo(targetRange, {
    contentsOnly: false
  });

  for (let i = 0; i < groupTemplate.blockHeight; i++) {
    targetSheet.setRowHeight(
      targetStartRow + i,
      templateSheet.getRowHeight(groupTemplate.blockStartRow + i)
    );
  }
}

function clearGradeValuesInStudentBlock(sheet, startRow, blockHeight, lastCol, templateInfo) {
  const range = sheet.getRange(startRow, 1, blockHeight, lastCol);

  range.breakApart();

  const values = range.getValues();
  const formulas = range.getFormulas();

  const protectedCols = new Set([
    templateInfo.cols.studentId,
    templateInfo.cols.class,
    templateInfo.cols.name,
    templateInfo.cols.description
  ]);

  const newValues = values.map((row, r) => {
    return row.map((cell, c) => {
      const colNumber = c + 1;

      if (formulas[r][c]) {
        return formulas[r][c];
      }

      if (protectedCols.has(colNumber)) {
        return cell;
      }

      return '';
    });
  });

  range.setValues(newValues);
}

function setStudentFieldsOnly(sheet, startRow, blockHeight, templateInfo, student) {
  mergeStudentField(
    sheet,
    startRow,
    blockHeight,
    templateInfo.cols.studentId,
    student.id
  );

  mergeStudentField(
    sheet,
    startRow,
    blockHeight,
    templateInfo.cols.class,
    student.class
  );

  mergeStudentField(
    sheet,
    startRow,
    blockHeight,
    templateInfo.cols.name,
    student.name
  );
}

function mergeStudentField(sheet, startRow, blockHeight, col, value) {
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

function clearRowsBelowMarker(sheet, markerRow, lastCol) {
  const rowsToClear = sheet.getMaxRows() - markerRow;

  if (rowsToClear <= 0) return;

  const range = sheet.getRange(markerRow + 1, 1, rowsToClear, lastCol);

  range.breakApart();
  range.clearContent();
  range.clearFormat();
  range.clearDataValidations();
}

function analyzeTemplate(sheet) {
  const values = sheet.getDataRange().getDisplayValues();

  const markerRowIndex = findTemplateMarkerRow(values);

  if (markerRowIndex === -1) {
    throw new Error(
      'На вкладці Template не знайдено рядок з кодами STUDENT_ID, CLASS, NAME, DESCRIPTION.'
    );
  }

  const markerRow = values[markerRowIndex];

  const studentIdColIndex = findColByExactText(markerRow, CONFIG.TEMPLATE_MARKERS.STUDENT_ID);
  const classColIndex = findColByExactText(markerRow, CONFIG.TEMPLATE_MARKERS.CLASS);
  const nameColIndex = findColByExactText(markerRow, CONFIG.TEMPLATE_MARKERS.NAME);
  const descriptionColIndex = findColByExactText(markerRow, CONFIG.TEMPLATE_MARKERS.DESCRIPTION);

  const groupTemplates = detectGroupTemplates(values, markerRowIndex, descriptionColIndex);

  if (groupTemplates.length === 0) {
    throw new Error('У Template не знайдено жодної групи з текстом "ОСВІТНЯ ГАЛУЗЬ".');
  }

  return {
    markerRow: markerRowIndex + 1,
    lastCol: sheet.getLastColumn(),

    cols: {
      studentId: studentIdColIndex + 1,
      class: classColIndex + 1,
      name: nameColIndex + 1,
      description: descriptionColIndex + 1
    },

    groupTemplates
  };
}

function detectGroupTemplates(values, markerRowIndex, descriptionColIndex) {
  const groupRows = [];

  for (let r = markerRowIndex + 1; r < values.length; r++) {
    const rowText = normalize(values[r].join(' '));

    if (rowText.includes('освітня галузь')) {
      groupRows.push(r);
    }
  }

  const templates = [];

  groupRows.forEach((groupRowIndex, index) => {
    const nextGroupRowIndex = index + 1 < groupRows.length
      ? groupRows[index + 1]
      : values.length;

    const groupName = getFirstNonEmptyCell(values[groupRowIndex]);
    const groupLabelCol = getFirstNonEmptyCol(values[groupRowIndex]) + 1;

    const totalRowIndex = findFirstTotalRowBetween(
      values,
      groupRowIndex + 1,
      nextGroupRowIndex - 1,
      descriptionColIndex
    );

    if (totalRowIndex === -1) return;

    const blockEndRowIndex = detectTemplateBlockEndRow(
      values,
      totalRowIndex,
      nextGroupRowIndex - 1,
      descriptionColIndex
    );

    if (blockEndRowIndex < totalRowIndex) return;

    templates.push({
      groupName,
      groupRow: groupRowIndex + 1,
      groupLabelCol,
      blockStartRow: totalRowIndex + 1,
      blockEndRow: blockEndRowIndex + 1,
      blockHeight: blockEndRowIndex - totalRowIndex + 1
    });
  });

  return templates;
}

function findFirstTotalRowBetween(values, startRowIndex, endRowIndex, descriptionColIndex) {
  const safeEnd = Math.min(endRowIndex, values.length - 1);

  for (let r = startRowIndex; r <= safeEnd; r++) {
    const description = normalize(values[r][descriptionColIndex]);

    if (description === normalize(CONFIG.TOTAL_LABEL)) {
      return r;
    }
  }

  return -1;
}

function detectTemplateBlockEndRow(values, totalRowIndex, endLimitRowIndex, descriptionColIndex) {
  let endRow = totalRowIndex;

  for (let r = totalRowIndex; r <= endLimitRowIndex; r++) {
    const description = String(values[r][descriptionColIndex] || '').trim();

    if (!description) break;

    const rowText = normalize(values[r].join(' '));

    if (r > totalRowIndex && rowText.includes('освітня галузь')) {
      break;
    }

    endRow = r;
  }

  return endRow;
}

function findTemplateByGroup(templateInfo, groupName) {
  const groupKey = normalize(groupName);

  if (!groupKey) return null;

  const exact = templateInfo.groupTemplates.find(template => {
    return normalize(template.groupName) === groupKey;
  });

  if (exact) return exact;

  const partial = templateInfo.groupTemplates.find(template => {
    const templateGroup = normalize(template.groupName);

    return templateGroup.includes(groupKey) || groupKey.includes(templateGroup);
  });

  return partial || null;
}

function replaceTopGroupTitle(sheet, groupName, markerRow) {
  if (!groupName) return;

  const topRowsCount = Math.max(markerRow - 1, 1);
  const range = sheet.getRange(1, 1, topRowsCount, sheet.getLastColumn());
  const values = range.getDisplayValues();

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cellText = normalize(values[r][c]);

      if (cellText.includes('освітня галузь')) {
        sheet.getRange(r + 1, c + 1).setValue(groupName);
        return;
      }
    }
  }

  sheet.getRange(1, 1).setValue(groupName);
}

function readSubjectsList(sheet) {
  const range = sheet.getDataRange();

  const displayValues = range.getDisplayValues();
  const rawValues = range.getValues();

  const headerRowIndex = findSubjectsHeaderRow(displayValues);

  if (headerRowIndex === -1) {
    throw new Error('Не знайдено заголовки у вкладці "Предмети".');
  }

  const headerRow = displayValues[headerRowIndex];

  const createCol = findColByHeaderLike(
    headerRow,
    CONFIG.SUBJECT_HEADERS.CREATE
  );

  const journalSubjectCol = findColByHeaderLike(
    headerRow,
    CONFIG.SUBJECT_HEADERS.JOURNAL_SUBJECT
  );

  const certificateSubjectCol = findColByHeaderLikeOptional(
    headerRow,
    CONFIG.SUBJECT_HEADERS.CERTIFICATE_SUBJECT
  );

  const groupCol = findColByHeaderLike(
    headerRow,
    CONFIG.SUBJECT_HEADERS.GROUP
  );

  const subjects = [];
  const seen = new Set();

  for (let r = headerRowIndex + 1; r < displayValues.length; r++) {
    const shouldCreate = isCheckedValue(
      rawValues[r][createCol],
      displayValues[r][createCol]
    );

    // Створюємо тільки ті предмети, де стоїть галочка
    if (!shouldCreate) continue;

    const journalSubject = String(displayValues[r][journalSubjectCol] || '').trim();

    if (!journalSubject) continue;

    const certificateSubject = certificateSubjectCol !== -1
      ? String(displayValues[r][certificateSubjectCol] || '').trim()
      : journalSubject;

    const group = String(displayValues[r][groupCol] || '').trim();

    if (!group) continue;

    const key = normalize(journalSubject);

    if (seen.has(key)) continue;

    seen.add(key);

    subjects.push({
      journalSubject,
      certificateSubject: certificateSubject || journalSubject,
      group
    });
  }

  return subjects;
}

function findSubjectsHeaderRow(values) {
  for (let r = 0; r < values.length; r++) {
    const row = values[r].map(normalize);

    const hasCreate = hasAnyHeader(row, CONFIG.SUBJECT_HEADERS.CREATE);
    const hasSubject = hasAnyHeader(row, CONFIG.SUBJECT_HEADERS.JOURNAL_SUBJECT);
    const hasGroup = hasAnyHeader(row, CONFIG.SUBJECT_HEADERS.GROUP);

    if (hasCreate && hasSubject && hasGroup) {
      return r;
    }
  }

  return -1;
}

function readStudents(sheet) {
  const values = sheet.getDataRange().getDisplayValues();

  const headerRowIndex = findStudentsHeaderRow(values);

  if (headerRowIndex === -1) {
    throw new Error('Не знайдено заголовки id, клас, Імʼя Прізвище.');
  }

  const headerRow = values[headerRowIndex];

  const idCol = findColByHeaderLike(headerRow, ['id']);
  const classCol = findColByHeaderLike(headerRow, ['клас']);
  const nameCol = findColByHeaderLike(headerRow, [
    'імʼя прізвище',
    'ім’я прізвище',
    'імя прізвище',
    'name'
  ]);

  const students = [];

  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const id = String(values[r][idCol] || '').trim();
    const studentClass = String(values[r][classCol] || '').trim();
    const name = String(values[r][nameCol] || '').trim();

    if (!id || !studentClass || !name) continue;
    if (!/^\d+$/.test(id)) continue;

    students.push({
      id,
      class: studentClass,
      name
    });
  }

  return students;
}

function findStudentsHeaderRow(values) {
  for (let r = 0; r < values.length; r++) {
    const row = values[r].map(normalize);

    const hasId = row.some(cell => cell === 'id');
    const hasClass = row.some(cell => cell === 'клас');
    const hasName = row.some(cell =>
      cell.includes('ім') ||
      cell.includes('прізвище') ||
      cell.includes('name')
    );

    if (hasId && hasClass && hasName) {
      return r;
    }
  }

  return -1;
}

function findTemplateMarkerRow(values) {
  for (let r = 0; r < values.length; r++) {
    const row = values[r].map(normalize);

    const hasStudentId = row.includes(normalize(CONFIG.TEMPLATE_MARKERS.STUDENT_ID));
    const hasClass = row.includes(normalize(CONFIG.TEMPLATE_MARKERS.CLASS));
    const hasName = row.includes(normalize(CONFIG.TEMPLATE_MARKERS.NAME));
    const hasDescription = row.includes(normalize(CONFIG.TEMPLATE_MARKERS.DESCRIPTION));

    if (hasStudentId && hasClass && hasName && hasDescription) {
      return r;
    }
  }

  return -1;
}

function replaceSubjectTitle(sheet, subjectName) {
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (normalize(values[r][c]) === normalize(CONFIG.TEMPLATE_TITLE_TEXT)) {
        sheet.getRange(r + 1, c + 1).setValue(subjectName);
        return;
      }
    }
  }
}

function writeLog(spreadsheet, log) {
  let sheet = spreadsheet.getSheetByName(CONFIG.LOG_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.LOG_SHEET_NAME);
  }

  sheet.clear();

  if (log.length === 0) {
    sheet.getRange(1, 1).setValue('Лог порожній');
    return;
  }

  const width = 6;

  const prepared = log.map(row => {
    const arr = Array.isArray(row) ? row : [row];

    while (arr.length < width) {
      arr.push('');
    }

    return arr.slice(0, width);
  });

  sheet.getRange(1, 1, prepared.length, width).setValues(prepared);
  sheet.autoResizeColumns(1, width);
}

function getRequiredSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Не знайдено вкладку: ' + sheetName);
  }

  return sheet;
}

function getSheetByNameLoose(ss, sheetName) {
  const exact = ss.getSheetByName(sheetName);

  if (exact) return exact;

  const needle = normalize(sheetName);

  return ss.getSheets().find(sheet => normalize(sheet.getName()) === needle) || null;
}

function isSystemSheetName(sheetName) {
  const systemNames = [
    CONFIG.TEMPLATE_SHEET_NAME,
    CONFIG.STUDENTS_SHEET_NAME,
    CONFIG.SUBJECTS_SHEET_NAME,
    CONFIG.LOG_SHEET_NAME,
    'Маппінг предметів'
  ];

  return systemNames.some(name => normalize(name) === normalize(sheetName));
}

function findColByExactText(row, text) {
  const needle = normalize(text);

  for (let c = 0; c < row.length; c++) {
    if (normalize(row[c]) === needle) {
      return c;
    }
  }

  throw new Error('Не знайдено колонку з текстом: ' + text);
}

function findColByHeaderLike(row, variants) {
  const normalizedVariants = variants.map(normalize);

  for (let c = 0; c < row.length; c++) {
    const cell = normalize(row[c]);

    if (normalizedVariants.some(v => cell === v || cell.includes(v))) {
      return c;
    }
  }

  throw new Error('Не знайдено колонку: ' + variants.join(' / '));
}

function findColByHeaderLikeOptional(row, variants) {
  const normalizedVariants = variants.map(normalize);

  for (let c = 0; c < row.length; c++) {
    const cell = normalize(row[c]);

    if (normalizedVariants.some(v => cell === v || cell.includes(v))) {
      return c;
    }
  }

  return -1;
}

function findIndexByHeaderLike(normalizedRow, variants) {
  const normalizedVariants = variants.map(normalize);

  for (let c = 0; c < normalizedRow.length; c++) {
    const cell = normalizedRow[c];

    if (normalizedVariants.some(v => cell === v || cell.includes(v))) {
      return c;
    }
  }

  return -1;
}

function hasAnyHeader(normalizedRow, variants) {
  return findIndexByHeaderLike(normalizedRow, variants) !== -1;
}

function isCheckedValue(rawValue, displayValue) {
  if (rawValue === true) return true;

  const textValue = normalize(displayValue);

  return (
    textValue === 'true' ||
    textValue === 'так' ||
    textValue === 'yes' ||
    textValue === '1'
  );
}

function makeSafeSheetName(name) {
  return String(name)
    .trim()
    .replace(/[\\/?*\[\]:]/g, '-')
    .substring(0, 99);
}

function getFirstNonEmptyCell(row) {
  for (let c = 0; c < row.length; c++) {
    const value = String(row[c] || '').trim();

    if (value) return value;
  }

  return '';
}

function getFirstNonEmptyCol(row) {
  for (let c = 0; c < row.length; c++) {
    const value = String(row[c] || '').trim();

    if (value) return c;
  }

  return 0;
}

function normalize(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[’ʼ`´]/g, "'")
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, ' ');
}
