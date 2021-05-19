function selectionDuplicates() {
  // найти строки различающиеся ростами и если разные цены - сообщить пользователю
  var a2 = SpreadsheetApp.getActiveSpreadsheet().getSelection().getActiveRange().getValues();
  // в одномерный массив
  a2 = a2.flat(Infinity);

  //console.log(a2);


  var duplicates = [];

  /* отсортировать массив, а затем проверить, совпадает ли «следующий элемент» с текущим элементом, и поместить его в массив: */
  var tempArray = [...a2].sort();
  //console.log(tempArray);

  for (let i = 0; i < tempArray.length; i++) {
    if (tempArray[i + 1] === tempArray[i]) {
      duplicates.push(tempArray[i]);
    }
  }

// массив оставляю уникальные
  duplicates = duplicates.filter(onlyUnique); 

  Browser.msgBox(duplicates);
  //console.log(duplicates);

}

function onlyUnique(value, index, self) {
  //проверяет, является ли данное значение первым встречающимся. Если нет, то это дубликат и не будет скопирован.
  return self.indexOf(value) === index;
}

function onOpen() {

  var ui = SpreadsheetApp.getUi();  // Or DocumentApp or FormApp.

  ui.createMenu('Прайсы')

    .addItem('Обрамить', 'formulaCodeFind')

    .addItem('Дубликаты', 'selectionDuplicates')

    .addSeparator()

    .addSubMenu(ui.createMenu('Sub-menu')

      .addItem('Тест', 'sheetActive'))

    .addToUi();

}

function formulaCodeFind() {

  // ячейки выделенные обрамить слева и справа
  const column = columnBySheet();
  if (column === undefined) { return; }

  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();
  var rowsCount = range.getNumRows();
  var colsCount = range.getNumColumns();

  var cell;
  var cellValue;
  var formula;
  for (var row = 1; row <= rowsCount; row++) {
    for (var col = 1; col <= colsCount; col++) {

      cell = range.getCell(row, col);

      formula = cell.getFormula()
      if (formula != "") {
        return;
      }

      cellValue = cell.getValue();

      if (cellValue == '') {
        return;
      }

      if (!IsNumeric(cellValue)) {

        // нечисла добавить кавычки
        cellValue = '"' + cellValue + '"';
      }

      cellValue = "=IFError(Index('сводная таблица'!" + column + ";MATCH("
        + cellValue + ";'сводная таблица'!$A:$A;0);1);\"код НЕ найден\")";

      cell.setValue(cellValue);

    }
  }
}


function menuItem2() {

  SpreadsheetApp.getUi().alert('You clicked the second menu item!');

  // DocumentApp.getUi().alert('You clicked the second menu item!'); - for DocumentApp

}

function IsNumeric(stringIN) {
  return isFinite(parseFloat(stringIN));
}


function columnBySheet() {
  // столбец в зависимости от имени листа
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  //Browser.msgBox(sheetName);
  if (sheetName == "Прайс без НДС") { return "$G:$G"; }
  if (sheetName == "Прайс для партнеров без НДС") { return "$F:$F"; }
  if (sheetName == "Прайс СНГ") { return "$I:$I"; }
  if (sheetName == "Прайс для партнеров СНГ") { return "$H:$H"; }
}

function sheetActive() {
  Browser.msgBox(SpreadsheetApp.getActiveSheet().getName())
}