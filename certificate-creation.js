const CONFIG = {
  SOURCE_SPREADSHEET_ID: '1oaSN2T1y4s9Cnzl4iFRDHhIwG9aUpbmIeHhfzdcahU4',
  TARGET_SPREADSHEET_ID: '1nr7E-22X_ep4yBtRu-rR_akW9E_uY4G8a4FuhElnoCY',

  TARGET_SHEET_NAME: 'Свідоцтво досягнень',

  // ID учня береться з цієї ячейки у свідоцтві
  STUDENT_ID_CELL: 'A1',

  MARKERS: {
    STUDENT_ID: 'STUDENT_ID',
    NAME: 'NAME',
    S1: 'S1',
    S2: 'S2',
    YEAR: 'YEAR'
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Оцінки')
    .addItem('Оновити всі предмети', 'updateAllSubjectsFromJournals')
    .addSeparator()
    .addItem('Тест створення логу', 'testLog')
    .addToUi();
}

function updateAllSubjectsFromJournals() {
  const log = [];

  try {
    log.push(['СТАРТ', new Date()]);

    const sourceSS = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
    const targetSS = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);

    const targetSheet = targetSS.getSheetByName(CONFIG.TARGET_SHEET_NAME);

    if (!targetSheet) {
      throw new Error('Не знайдено аркуш свідоцтва: ' + CONFIG.TARGET_SHEET_NAME);
    }

    const studentId = getStudentIdFromCertificate(targetSheet);

    log.push(['ID учня з ячейки ' + CONFIG.STUDENT_ID_CELL, studentId]);

    const targetValues = targetSheet.getDataRange().getDisplayValues();
    const certificateCols = findCertificatePeriodColumns(targetValues);

    log.push(['Таблиця журналів', sourceSS.getName()]);
    log.push(['Таблиця свідоцтва', targetSS.getName()]);
    log.push(['Аркуш свідоцтва', targetSheet.getName()]);
    log.push(['Колонка S1 / І семестр у свідоцтві', certificateCols.S1 + 1]);
    log.push(['Колонка S2 / ІІ семестр у свідоцтві', certificateCols.S2 + 1]);
    log.push(['Колонка YEAR / За рік у свідоцтві', certificateCols.YEAR + 1]);

    sourceSS.getSheets().forEach(journalSheet => {
      processJournalSheet(
        journalSheet,
        targetSheet,
        targetValues,
        certificateCols,
        log,
        studentId
      );
    });

    SpreadsheetApp.flush();

    log.push(['ФІНІШ', 'Оновлення завершено']);

    writeLog(targetSS, log);

    SpreadsheetApp.getUi().alert(
      'Готово. Оцінки оновлено для ID: ' + studentId + '. Перевір аркуш LOG_ОНОВЛЕННЯ.'
    );
  } catch (error) {
    log.push(['ПОМИЛКА', error.message]);
    log.push(['STACK', error.stack || '']);

    const targetSS = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
    writeLog(targetSS, log);

    SpreadsheetApp.getUi().alert(
      'Помилка: ' + error.message + '\n\nДеталі дивись в LOG_ОНОВЛЕННЯ.'
    );

    throw error;
  }
}

function getStudentIdFromCertificate(sheet) {
  const studentId = String(sheet.getRange(CONFIG.STUDENT_ID_CELL).getDisplayValue()).trim();

  if (!studentId) {
    throw new Error('У ячейці ' + CONFIG.STUDENT_ID_CELL + ' не вказано ID учня.');
  }

  return studentId;
}

function processJournalSheet(journalSheet, targetSheet, targetValues, certificateCols, log, studentId) {
  const subjectName = journalSheet.getName().trim();

  log.push(['---']);
  log.push(['Перевіряю вкладку журналу', subjectName]);

  const journalValues = journalSheet.getDataRange().getDisplayValues();

  const markerRow = findJournalMarkerRowOrNull(journalValues);

  if (markerRow === null) {
    log.push([
      'ПРОПУЩЕНО',
      'На вкладці немає службових кодів STUDENT_ID, NAME, S1, S2, YEAR'
    ]);
    return;
  }

  const studentIdCol = findColInRow(journalValues[markerRow], CONFIG.MARKERS.STUDENT_ID);
  const nameCol = findColInRow(journalValues[markerRow], CONFIG.MARKERS.NAME);

  const journalCols = {
    S1: findColInRow(journalValues[markerRow], CONFIG.MARKERS.S1),
    S2: findColInRow(journalValues[markerRow], CONFIG.MARKERS.S2),
    YEAR: findColInRow(journalValues[markerRow], CONFIG.MARKERS.YEAR)
  };

  const studentRow = findStudentRowOrNull(journalValues, studentIdCol, studentId);

  if (studentRow === null) {
    log.push([
      'ПРОПУЩЕНО',
      'Учня з ID ' + studentId + ' немає на вкладці ' + subjectName
    ]);
    return;
  }

  const studentName = journalValues[studentRow][nameCol];

  const criteriaCol = findCriteriaColumn(
    journalValues,
    studentRow,
    nameCol,
    journalCols.S1
  );

  log.push(['Предмет', subjectName]);
  log.push(['Рядок службових кодів', markerRow + 1]);
  log.push(['Учень', studentName]);
  log.push(['Перший рядок учня в журналі', studentRow + 1]);
  log.push(['Колонка STUDENT_ID', studentIdCol + 1]);
  log.push(['Колонка NAME', nameCol + 1]);
  log.push(['Колонка критерію', criteriaCol + 1]);
  log.push(['Колонка S1', journalCols.S1 + 1]);
  log.push(['Колонка S2', journalCols.S2 + 1]);
  log.push(['Колонка YEAR', journalCols.YEAR + 1]);

  const journalCriteria = readStudentCriteriaFromJournal(
    journalValues,
    studentRow,
    studentIdCol,
    criteriaCol,
    journalCols
  );

  if (journalCriteria.length === 0) {
    log.push(['ПРОПУЩЕНО', 'Не знайдено жодного критерію для учня']);
    return;
  }

  const subjectCell = findSubjectCellInCertificateOrNull(targetValues, subjectName);

  if (!subjectCell) {
    log.push([
      'ПРОПУЩЕНО',
      'У свідоцтві не знайдено предмет: ' + subjectName
    ]);
    return;
  }

  const subjectBlock = getCertificateSubjectBlock(targetValues, subjectCell);

  log.push(['Рядок предмета у свідоцтві', subjectCell.row + 1]);
  log.push(['Колонка предмета у свідоцтві', subjectCell.col + 1]);
  log.push(['Кінець блоку предмета', subjectBlock.endRow + 1]);

  journalCriteria.forEach(item => {
    const targetRow = findCriterionRowInCertificate(
      targetValues,
      subjectBlock.startRow,
      subjectBlock.endRow,
      item.label
    );

    if (targetRow === null) {
      log.push([
        subjectName,
        'НЕ ЗНАЙДЕНО КРИТЕРІЙ У СВІДОЦТВІ',
        item.label,
        'S1: ' + item.S1,
        'S2: ' + item.S2,
        'YEAR: ' + item.YEAR
      ]);
      return;
    }

    targetSheet.getRange(targetRow + 1, certificateCols.S1 + 1).setValue(item.S1);
    targetSheet.getRange(targetRow + 1, certificateCols.S2 + 1).setValue(item.S2);
    targetSheet.getRange(targetRow + 1, certificateCols.YEAR + 1).setValue(item.YEAR);

    log.push([
      subjectName,
      item.label,
      'рядок свідоцтва: ' + (targetRow + 1),
      'S1: ' + item.S1,
      'S2: ' + item.S2,
      'YEAR: ' + item.YEAR
    ]);
  });
}

function readStudentCriteriaFromJournal(values, studentRow, studentIdCol, criteriaCol, journalCols) {
  const result = [];

  const endRow = findNextStudentRow(values, studentRow, studentIdCol);

  for (let r = studentRow; r < endRow; r++) {
    const label = String(values[r][criteriaCol] || '').trim();

    if (!label) continue;

    result.push({
      label: label,
      S1: cleanGrade(values[r][journalCols.S1]),
      S2: cleanGrade(values[r][journalCols.S2]),
      YEAR: cleanGrade(values[r][journalCols.YEAR])
    });
  }

  return result;
}

function findNextStudentRow(values, studentRow, studentIdCol) {
  for (let r = studentRow + 1; r < values.length; r++) {
    const value = String(values[r][studentIdCol] || '').trim();

    if (value !== '') {
      return r;
    }
  }

  return values.length;
}

function findCriteriaColumn(values, studentRow, nameCol, firstGradeCol) {
  for (let c = nameCol + 1; c < firstGradeCol; c++) {
    const value = normalize(values[studentRow][c]);

    if (value.includes('загальна оцінка')) {
      return c;
    }
  }

  for (let c = nameCol + 1; c < firstGradeCol; c++) {
    const value = String(values[studentRow][c] || '').trim();

    if (value !== '' && isNaN(Number(value))) {
      return c;
    }
  }

  throw new Error('Не знайдено колонку з критеріями біля рядка учня.');
}

function findJournalMarkerRowOrNull(values) {
  for (let r = 0; r < values.length; r++) {
    const row = values[r].map(normalize);

    const hasStudentId = row.includes(normalize(CONFIG.MARKERS.STUDENT_ID));
    const hasName = row.includes(normalize(CONFIG.MARKERS.NAME));
    const hasS1 = row.includes(normalize(CONFIG.MARKERS.S1));
    const hasS2 = row.includes(normalize(CONFIG.MARKERS.S2));
    const hasYear = row.includes(normalize(CONFIG.MARKERS.YEAR));

    if (hasStudentId && hasName && hasS1 && hasS2 && hasYear) {
      return r;
    }
  }

  return null;
}

function findCertificatePeriodColumns(values) {
  for (let r = 0; r < values.length; r++) {
    const row = values[r].map(normalize);

    const s1 = row.indexOf('s1');
    const s2 = row.indexOf('s2');
    const year = row.indexOf('year');

    if (s1 !== -1 && s2 !== -1 && year !== -1) {
      return {
        S1: s1,
        S2: s2,
        YEAR: year
      };
    }
  }

  for (let r = 0; r < values.length; r++) {
    const row = values[r].map(normalize);

    let s1 = -1;
    let s2 = -1;
    let year = -1;

    for (let c = 0; c < row.length; c++) {
      const cell = row[c];

      if (
        [
          'і с',
          'i с',
          '1 с',
          'і семестр',
          'i семестр',
          '1 семестр',
          'перший семестр'
        ].includes(cell)
      ) {
        s1 = c;
      }

      if (
        [
          'іі с',
          'ii с',
          '2 с',
          'іі семестр',
          'ii семестр',
          '2 семестр',
          'другий семестр'
        ].includes(cell)
      ) {
        s2 = c;
      }

      if (
        [
          'за рік',
          'за рiк',
          'річна',
          'річний бал',
          'рік'
        ].includes(cell)
      ) {
        year = c;
      }
    }

    if (s1 !== -1 && s2 !== -1 && year !== -1) {
      return {
        S1: s1,
        S2: s2,
        YEAR: year
      };
    }
  }

  throw new Error('Не знайдено колонки S1, S2, YEAR або І с, ІІ с, За рік у свідоцтві.');
}

function findSubjectCellInCertificateOrNull(values, subjectName) {
  const needle = normalize(subjectName);

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (normalize(values[r][c]) === needle) {
        return {
          row: r,
          col: c
        };
      }
    }
  }

  for (let r = 0; r < values.length; r++) {
    const rowText = normalize(values[r].join(' '));

    if (rowText.includes(needle)) {
      return {
        row: r,
        col: 0
      };
    }
  }

  return null;
}

function getCertificateSubjectBlock(values, subjectCell) {
  const startRow = subjectCell.row;
  const subjectCol = subjectCell.col;

  let endRow = Math.min(values.length - 1, startRow + 25);

  for (let r = startRow + 1; r < values.length; r++) {
    const value = String(values[r][subjectCol] || '').trim();

    if (value !== '') {
      endRow = r - 1;
      break;
    }
  }

  return {
    startRow: startRow,
    endRow: endRow
  };
}

function findCriterionRowInCertificate(values, startRow, endRow, journalCriterionLabel) {
  for (let r = startRow; r <= endRow; r++) {
    const rowText = values[r].join(' ');

    if (isCriterionMatch(journalCriterionLabel, rowText)) {
      return r;
    }
  }

  return null;
}

function isCriterionMatch(journalLabel, certificateRowText) {
  const source = normalize(journalLabel);
  const target = normalize(certificateRowText);

  if (!source || !target) return false;

  if (source === target) return true;
  if (target.includes(source)) return true;
  if (source.includes(target)) return true;

  const sourceTokens = getImportantTokens(source);
  const targetTokens = getImportantTokens(target);

  if (sourceTokens.length === 0) return false;

  const matched = sourceTokens.filter(token => targetTokens.includes(token)).length;

  if (sourceTokens.length <= 2) {
    return matched === sourceTokens.length;
  }

  return matched / sourceTokens.length >= 0.7;
}

function getImportantTokens(text) {
  const stopWords = [
    'і',
    'й',
    'та',
    'з',
    'зі',
    'із',
    'у',
    'в',
    'на',
    'до',
    'за',
    'по',
    'що',
    'як'
  ];

  return normalize(text)
    .split(/[^а-щьюяґєіїa-z0-9']+/i)
    .map(t => t.trim())
    .filter(t => t.length >= 3)
    .filter(t => !stopWords.includes(t));
}

function findColInRow(row, text) {
  const needle = normalize(text);

  for (let c = 0; c < row.length; c++) {
    if (normalize(row[c]) === needle) {
      return c;
    }
  }

  throw new Error('Не знайдено колонку з міткою: ' + text);
}

function findStudentRowOrNull(values, studentIdCol, studentId) {
  const needle = String(studentId).trim();

  for (let r = 0; r < values.length; r++) {
    if (String(values[r][studentIdCol]).trim() === needle) {
      return r;
    }
  }

  return null;
}

function cleanGrade(value) {
  const text = String(value || '').trim();

  if (
    text === '' ||
    text.includes('#DIV') ||
    text.includes('#DZIEL') ||
    text.includes('#VALUE') ||
    text.includes('#N/A') ||
    text.includes('#ERROR') ||
    text.includes('#REF')
  ) {
    return '';
  }

  return text;
}

function normalize(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[’ʼ`´]/g, "'")
    .replace(/\s+/g, ' ');
}

function writeLog(spreadsheet, log) {
  let sheet = spreadsheet.getSheetByName('LOG_ОНОВЛЕННЯ');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('LOG_ОНОВЛЕННЯ');
  }

  sheet.clear();

  const width = 7;

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

function testLog() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);

  writeLog(ss, [
    ['ТЕСТ', 'Скрипт працює'],
    ['Дата', new Date()],
    ['Таблиця', ss.getName()],
    ['ID береться з ячейки', CONFIG.STUDENT_ID_CELL]
  ]);

  SpreadsheetApp.getUi().alert('Тестовий лог створено.');
}
