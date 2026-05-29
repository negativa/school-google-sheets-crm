function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Галочки')
    .addItem('Рандомно 50%', 'randomCheckboxes50')
    .addItem('Рандомно з відсотком', 'randomCheckboxesByPercent')
    .addItem('Поставити точну кількість', 'randomCheckboxesExactCount')
    .addToUi();
}

// Рандомно ставить галочки приблизно у 50% чекбоксів
function randomCheckboxes50() {
  randomizeCheckboxesInSelectedRange(50);
}

// Запитує відсоток, наприклад 30 або 70
function randomCheckboxesByPercent() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Рандомні галочки',
    'Введи відсоток галочок, наприклад 30 або 50:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const percent = parseFloat(
    response.getResponseText().replace(',', '.')
  );

  if (isNaN(percent) || percent < 0 || percent > 100) {
    ui.alert('Введи число від 0 до 100.');
    return;
  }

  randomizeCheckboxesInSelectedRange(percent);
}

// Ставить галочки у випадкових клітинках виділеного діапазону
function randomizeCheckboxesInSelectedRange(percent) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();

  if (!range) {
    SpreadsheetApp.getUi().alert('Спочатку виділи діапазон з чекбоксами.');
    return;
  }

  const validations = range.getDataValidations();
  const values = range.getValues();

  let changed = 0;

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (isCheckbox(validations[r][c])) {
        values[r][c] = Math.random() * 100 < percent;
        changed++;
      }
    }
  }

  range.setValues(values);

  SpreadsheetApp.getUi().alert(
    'Готово. Оновлено чекбоксів: ' + changed
  );
}

// Ставить точну кількість галочок у випадковому порядку
function randomCheckboxesExactCount() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();

  if (!range) {
    ui.alert('Спочатку виділи діапазон з чекбоксами.');
    return;
  }

  const response = ui.prompt(
    'Точна кількість галочок',
    'Скільки галочок потрібно поставити?',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const count = parseInt(response.getResponseText(), 10);

  if (isNaN(count) || count < 0) {
    ui.alert('Введи нормальне число.');
    return;
  }

  const validations = range.getDataValidations();
  const values = range.getValues();

  const checkboxCells = [];

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (isCheckbox(validations[r][c])) {
        checkboxCells.push({ row: r, col: c });
        values[r][c] = false;
      }
    }
  }

  if (count > checkboxCells.length) {
    ui.alert(
      'У виділеному діапазоні лише ' +
      checkboxCells.length +
      ' чекбоксів.'
    );
    return;
  }

  shuffleArray(checkboxCells);

  for (let i = 0; i < count; i++) {
    const cell = checkboxCells[i];
    values[cell.row][cell.col] = true;
  }

  range.setValues(values);

  ui.alert(
    'Готово. Поставлено галочок: ' + count
  );
}

// Перевіряє, чи клітинка є чекбоксом
function isCheckbox(validation) {
  return (
    validation &&
    validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX
  );
}

// Перемішує масив випадковим чином
function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));

    const temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
}
