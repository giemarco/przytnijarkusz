//autor: sztukawdanych.pl

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🪄 Magia arkusza')
    .addItem('Przytnij arkusz', 'trimSheet')
    .addToUi();
}

function trimSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();

  // zakres CAŁEGO arkusza
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  const range = sheet.getRange(1, 1, maxRows, maxCols);
  const values = range.getValues();
  const formulas = range.getFormulas();

  let lastUsedRow = 0;
  let lastUsedCol = 0;

  // szukamy ostatniego faktycznie użytego wiersza i kolumny
  for (let r = 0; r < maxRows; r++) {
    for (let c = 0; c < maxCols; c++) {
      if (values[r][c] !== '' || formulas[r][c] !== '') {
        if (r + 1 > lastUsedRow) lastUsedRow = r + 1;
        if (c + 1 > lastUsedCol) lastUsedCol = c + 1;
      }
    }
  }

  ss.toast('Przycinam arkusz…', 'Magia arkusza', 5);
  SpreadsheetApp.flush();

  // USUWANIE WIERSZY NA DOLE
  if (lastUsedRow < maxRows) {
    sheet.deleteRows(lastUsedRow + 1, maxRows - lastUsedRow);
  }

  // USUWANIE KOLUMN PO PRAWEJ
  if (lastUsedCol < maxCols) {
    sheet.deleteColumns(lastUsedCol + 1, maxCols - lastUsedCol);
  }

  ss.toast(
    `Gotowe ✅\nUsunięto:\n• ${maxRows - lastUsedRow} wierszy\n• ${maxCols - lastUsedCol} kolumn`,
    'Magia arkusza',
    6
  );
}
