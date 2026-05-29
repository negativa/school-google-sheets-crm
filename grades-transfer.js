const SAFE_COPY_CONFIG = {
  SOURCE_SPREADSHEET_ID: '16TmHEPb8OoFyarwyRAj9QLChJE2vcGimEaVWK7Wgekg',
  TARGET_SPREADSHEET_ID: '1oaSN2T1y4s9Cnzl4iFRDHhIwG9aUpbmIeHhfzdcahU4',

  LOG_SPREADSHEET_ID: '1To1UG3Ay4Fs5z02xsOyAjjwiae3SU-upxGUGaBxNIDY',
  LOG_SHEET_NAME: 'LOG_ПЕРЕНОС_ОЦІНОК',

  MARKERS: {
    STUDENT_ID: 'STUDENT_ID',
    NAME: 'NAME',
    DESCRIPTION: 'DESCRIPTION',

    AVG1: 'AVG1',
    CERT1: 'CERT1',
    S1: 'S1',

    AVG2: 'AVG2',
    CERT2: 'CERT2',
    S2: 'S2',

    YEAR: 'YEAR'
  },

  SYSTEM_SHEETS: [
    'Template',
    'Students list',
    'Предмети',
    'Маппінг предметів',
    'LOG_ЖУРНАЛИ',
    'LOG_ПЕРЕНОС_ОЦІНОК'
  ],

  TOTAL_LABEL: 'Загальна оцінка',

  COPY_ONLY_NON_EMPTY: true,
  SKIP_ERROR_VALUES: true,

  // ВАЖЛИВО: формули в новому журналі не перезаписуємо
  DO_NOT_OVERWRITE_FORMULAS: true
};

// Якщо в тебе вже є onOpen(), не створюй другий.
// Просто додай всередину свого onOpen() рядок:
// addSafeGradeCopyMenu();

function addSafeGradeCopyMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Перенос оцінок')
    .addItem('Безпечно перенести оцінки', 'copyGradesWithoutDeletingFormulas')
    .addToUi();
}

function copyGradesWithoutDeletingFormulas() {
  const sourceSS = SpreadsheetApp.openById(SAFE_COPY_CONFIG.SOURCE_SPREADSHEET_ID);
  const targetSS = SpreadsheetApp.openById(SAFE_COPY_CONFIG.TARGET_SPREADSHEET_ID);

  const log = [];

  let updatedSheets = 0;
  let skippedSheets = 0;
  let copiedCells = 0;

  const targetSheets = targetSS
    .getSheets()
    .filter(sheet => !isSystemSheet(sheet.getName()));

  targetSheets.forEach(targetSheet => {
    const subjectName = targetSheet.getName();
    const sourceSheet = findSheetByNameLoose(sourceSS, subjectName);

    if (!sourceSheet) {
      skippedSheets++;
      log.push([subjectName, 'ПРОПУЩЕНО', 'У старій таблиці немає вкладки з такою назвою']);
      return;
    }

    try {
      const result = copyOneSubjectSafely(sourceSheet, targetSheet, log);

      copiedCells += result.copiedCells;

      if (result.copiedCells > 0) {
        updatedSheets++;
      }

      log.push([
        subjectName,
        'ГОТОВО',
        'Скопійовано клітинок: ' + result.copiedCells,
        'Не знайдено рядків: ' + result.missedRows
      ]);
    } catch (error) {
      skippedSheets++;
      log.push([subjectName, 'ПОМИЛКА', error.message]);
    }
  });

  writeSafeCopyLog(log);

  SpreadsheetApp.getUi().alert(
    'Готово.\n' +
    'Оновлено вкладок: ' + updatedSheets + '\n' +
    'Пропущено вкладок: ' + skippedSheets + '\n' +
    'Скопійовано оцінок: ' + copiedCells + '\n\n' +
    'Формули в новому журналі не перезаписувались.\n' +
    'Деталі дивись у таблиці логів.'
  );
}

function copyOneSubjectSafely(sourceSheet, targetSheet, log) {
  const sourceValues = sourceSheet.getDataRange().getDisplayValues();

  const targetRange = targetSheet.getDataRange();
  const targetDisplayValues = targetRange.getDisplayValues();
  const targetFormulas = targetRange.getFormulas();

  const sourceLayout = detectLayout(sourceValues, sourceSheet.getName());
  const targetLayout = detectLayout(targetDisplayValues, targetSheet.getName());

  const sourceRows = buildRowsIndex(sourceValues, sourceLayout);
  const targetRows = buildRowsList(targetDisplayValues, targetLayout);

  const columnPairs = [
    ...buildRegularColumnPairs(sourceValues, targetDisplayValues, sourceLayout, targetLayout, 'S1'),
    ...buildRegularColumnPairs(sourceValues, targetDisplayValues, sourceLayout, targetLayout, 'S2'),

    // AVG копіюємо тільки якщо в новій клітинці немає формули
    buildMarkerColumnPair(sourceLayout, targetLayout, 'AVG1'),
    buildMarkerColumnPair(sourceLayout, targetLayout, 'AVG2'),

    // Атестацію копіюємо також тільки якщо там немає формули
    buildMarkerColumnPair(sourceLayout, targetLayout, 'CERT1'),
    buildMarkerColumnPair(sourceLayout, targetLayout, 'CERT2')
  ].filter(Boolean);

  let copiedCells = 0;
  let missedRows = 0;

  targetRows.forEach(targetRow => {
    const sourceRow = sourceRows.get(targetRow.key);

    if (!sourceRow) {
      missedRows++;
      return;
    }

    columnPairs.forEach(pair => {
      const sourceValue = sourceValues[sourceRow.rowIndex][pair.sourceCol];

      if (!shouldCopyValue(sourceValue)) return;

      const targetFormula = targetFormulas[targetRow.rowIndex][pair.targetCol];

      // Головний захист: якщо в новій клітинці формула — не чіпаємо її
      if (SAFE_COPY_CONFIG.DO_NOT_OVERWRITE_FORMULAS && targetFormula) {
        return;
      }

      targetSheet
        .getRange(targetRow.rowIndex + 1, pair.targetCol + 1)
        .setValue(normalizeCopiedValue(sourceValue));

      copiedCells++;
    });
  });

  return {
    copiedCells,
    missedRows
  };
}

function buildRegularColumnPairs(sourceValues, targetValues, sourceLayout, targetLayout, semester) {
  const sourceCols = getRegularColumnsBySemester(sourceLayout, semester);
  const targetCols = getRegularColumnsBySemester(targetLayout, semester);

  const count = Math.min(sourceCols.length, targetCols.length);
  const pairs = [];

  for (let i = 0; i < count; i++) {
    pairs.push({
      sourceCol: sourceCols[i],
      targetCol: targetCols[i],
      type: semester
    });
  }

  return pairs;
}

function getRegularColumnsBySemester(layout, semester) {
  const m = layout.markerCols;

  let startCol = -1;
  let endCol = -1;

  if (semester === 'S1') {
    startCol = layout.descriptionCol + 1;
    endCol = firstExistingCol([
      m.AVG1,
      m.CERT1,
      m.S1
    ]) - 1;
  }

  if (semester === 'S2') {
    startCol = m.S1 + 1;
    endCol = firstExistingCol([
      m.AVG2,
      m.CERT2,
      m.S2
    ]) - 1;
  }

  if (startCol < 0 || endCol < startCol) return [];

  const cols = [];

  for (let c = startCol; c <= endCol; c++) {
    cols.push(c);
  }

  return cols;
}

function buildMarkerColumnPair(sourceLayout, targetLayout, markerName) {
  const sourceCol = sourceLayout.markerCols[markerName];
  const targetCol = targetLayout.markerCols[markerName];

  if (sourceCol === -1 || targetCol === -1) return null;

  return {
    sourceCol,
    targetCol,
    type: markerName
  };
}

function detectLayout(values, sheetName) {
  const markerRowIndex = findMarkerRow(values);

  if (markerRowIndex === -1) {
    throw new Error('Не знайдено службовий рядок NAME / DESCRIPTION на вкладці: ' + sheetName);
  }

  const markerRow = values[markerRowIndex].map(normalize);

  const nameCol = markerRow.indexOf(normalize(SAFE_COPY_CONFIG.MARKERS.NAME));
  const descriptionCol = markerRow.indexOf(normalize(SAFE_COPY_CONFIG.MARKERS.DESCRIPTION));
  const studentIdCol = markerRow.indexOf(normalize(SAFE_COPY_CONFIG.MARKERS.STUDENT_ID));

  if (nameCol === -1 || descriptionCol === -1) {
    throw new Error('У службовому рядку не знайдено NAME або DESCRIPTION на вкладці: ' + sheetName);
  }

  const markerCols = {};

  Object.keys(SAFE_COPY_CONFIG.MARKERS).forEach(key => {
    markerCols[key] = markerRow.indexOf(normalize(SAFE_COPY_CONFIG.MARKERS[key]));
  });

  return {
    markerRowIndex,
    nameCol,
    descriptionCol,
    studentIdCol,
    markerCols
  };
}

function findMarkerRow(values) {
  for (let r = 0; r < values.length; r++) {
    const row = values[r].map(normalize);

    const hasName = row.includes(normalize(SAFE_COPY_CONFIG.MARKERS.NAME));
    const hasDescription = row.includes(normalize(SAFE_COPY_CONFIG.MARKERS.DESCRIPTION));

    if (hasName && hasDescription) {
      return r;
    }
  }

  return -1;
}

function buildRowsIndex(values, layout) {
  const map = new Map();

  buildRowsList(values, layout).forEach(item => {
    if (!map.has(item.key)) {
      map.set(item.key, item);
    }
  });

  return map;
}

function buildRowsList(values, layout) {
  const rows = [];

  let currentStudentName = '';

  for (let r = 0; r < values.length; r++) {
    const rawName = values[r][layout.nameCol];
    const studentName = cleanStudentName(rawName);

    if (studentName && looksLikeStudentName(studentName)) {
      currentStudentName = studentName;
    }

    const description = cleanDescription(values[r][layout.descriptionCol]);

    if (!currentStudentName || !description) continue;

    rows.push({
      rowIndex: r,
      studentName: currentStudentName,
      description,
      key: makeRowKey(currentStudentName, description)
    });
  }

  return rows;
}

function makeRowKey(studentName, description) {
  return normalize(cleanStudentName(studentName)) + '||' + normalize(cleanDescription(description));
}

function cleanStudentName(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ');
}

function cleanDescription(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ');
}

function looksLikeStudentName(value) {
  const textValue = String(value || '').trim();
  const normalized = normalize(textValue);

  if (!textValue) return false;
  if (textValue.length < 5) return false;
  if (/^\d+$/.test(textValue)) return false;

  if (normalized.includes(normalize(SAFE_COPY_CONFIG.TOTAL_LABEL))) return false;
  if (normalized.includes('description')) return false;
  if (normalized.includes('name')) return false;
  if (normalized.includes('дата')) return false;
  if (normalized.includes('семестр')) return false;
  if (normalized.includes('середній бал')) return false;

  return textValue.split(/\s+/).length >= 2;
}

function shouldCopyValue(value) {
  const textValue = String(value || '').trim();

  if (SAFE_COPY_CONFIG.COPY_ONLY_NON_EMPTY && textValue === '') {
    return false;
  }

  if (SAFE_COPY_CONFIG.SKIP_ERROR_VALUES && isErrorValue(textValue)) {
    return false;
  }

  return true;
}

function isErrorValue(value) {
  const textValue = String(value || '').trim().toUpperCase();

  return (
    textValue.includes('#DIV') ||
    textValue.includes('#VALUE') ||
    textValue.includes('#N/A') ||
    textValue.includes('#REF') ||
    textValue.includes('#ERROR') ||
    textValue.includes('#NAME')
  );
}

function normalizeCopiedValue(value) {
  const textValue = String(value || '').trim();

  if (/^-?\d+([.,]\d+)?$/.test(textValue)) {
    return Number(textValue.replace(',', '.'));
  }

  return textValue;
}

function firstExistingCol(cols) {
  return cols
    .filter(col => col !== -1 && col !== undefined && col !== null)
    .sort((a, b) => a - b)[0] ?? -1;
}

function findSheetByNameLoose(spreadsheet, sheetName) {
  const exact = spreadsheet.getSheetByName(sheetName);

  if (exact) return exact;

  const needle = normalize(sheetName);

  return spreadsheet
    .getSheets()
    .find(sheet => normalize(sheet.getName()) === needle) || null;
}

function isSystemSheet(sheetName) {
  return SAFE_COPY_CONFIG.SYSTEM_SHEETS.some(name => normalize(name) === normalize(sheetName));
}

function writeSafeCopyLog(log) {
  const logSS = SpreadsheetApp.openById(SAFE_COPY_CONFIG.LOG_SPREADSHEET_ID);

  let sheet = logSS.getSheetByName(SAFE_COPY_CONFIG.LOG_SHEET_NAME);

  if (!sheet) {
    sheet = logSS.insertSheet(SAFE_COPY_CONFIG.LOG_SHEET_NAME);
  }

  sheet.clear();

  const header = [
    [
      'Дата',
      'Час',
      'Подія / Предмет',
      'Статус',
      'Деталі 1',
      'Деталі 2',
      'Деталі 3'
    ]
  ];

  const now = new Date();
  const timezone = Session.getScriptTimeZone();

  const prepared = log.length
    ? log.map(row => {
        const arr = Array.isArray(row) ? row : [row];

        return [
          Utilities.formatDate(now, timezone, 'yyyy-MM-dd'),
          Utilities.formatDate(now, timezone, 'HH:mm:ss'),
          arr[0] || '',
          arr[1] || '',
          arr[2] || '',
          arr[3] || '',
          arr[4] || ''
        ];
      })
    : [[
        Utilities.formatDate(now, timezone, 'yyyy-MM-dd'),
        Utilities.formatDate(now, timezone, 'HH:mm:ss'),
        'Лог порожній',
        '',
        '',
        '',
        ''
      ]];

  const output = header.concat(prepared);

  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.autoResizeColumns(1, output[0].length);
}

function normalize(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[’ʼ`´]/g, "'")
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, ' ');
}
