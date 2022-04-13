//https://docs.google.com/spreadsheets/d/1gEQ2Yz3lC9DRY4y2z_Ni6G6LWek_xTNDmBJEnrt5z7Y/edit?pli=1#gid=0
// Пушкин: затея с массивами формул - провалилась - Google Apps Script не умеет этого.

// ТЗ
// в указанном столбце
// ячейки пустые заполнить формулой.
// Решение:
// Так как строк много, то на листе работать проблемно - гугл может не выдержать по времени.
// Диапазон столбца определить
// Google Apps Script не умеет копировать в массив формулы и значения одной командой
// скопировать формулы в массив
// скопировать значения в массив другой
// в формуле =ЕСЛИОШИБКА(ЕСЛИ(W2 > 0; (N2 - W2) / (N2);"");"")
// буду заменять буквы с цифрами на эти же буквы с номерами сооответствующих строк
// проходом по массиву в пустые элементы
// вставить скорректированные формулы
// массив на лист



// const global_Function = '=ЕСЛИОШИБКА(ЕСЛИ(W2 > 0; (N2 - W2) / N2;"");"")'
// const globalColumn28 = 28;

function columnPriceDownSheetActive_Test() {
  //  SpreadsheetApp.getActive().getSheetByName('Тест скрипта').activate();
  var sheet = SpreadsheetApp.getActive().getSheetByName('Тест скрипта');
  sheet.activate();
  columnPriceDownSheetActive();
  var cellValue = sheet.getRange(2, 28).getValue()
  if (cellValue === '') {
    Logger.log('Увы, в ячейке (2,28) пусто')
  }
}

function columnPriceDownSheetActive() {

  // убедиться, что на листе активном 
  // название столбца в нужном месте
  // вызываю из меню скриптов для активного листа

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell_ = sheet.getRange(1, globalColumn28)
  var value = cell_.getValue();

  if (value !== 'снижение от НМЦ') {
    Browser.msgBox('Номер столбца ячейки PriceDown ожидается ' + globalColumn28);
  }
  else {
    // диапазон в массив
    var rng = rangeColumnByNumber(sheet, globalColumn28)
    var a2dValues = rng.getValues();
    var a2dFormul = rng.getFormulas();
    var a2dValuesAndFormulas = array2dCopy2LeftFromRightEmptyNot(a2dValues, 0, a2dFormul, 0);
    a2dColumnreplace(a2dValuesAndFormulas, 0, global_Function);
    // вернуть массив на лист
    array2d2RangeFormulas(cell_, a2dValuesAndFormulas)
  }
}



function a2dColumnreplace(a2d, col, formula) {
  // в массиве 2мерном заменить 
  // пустые элементы на строку с 
  // текущим номером строки
  for (var row = 0; row < a2d.length; row++) {
    if (a2d[row][col] === '') {
      a2d[row][col] = formulaRowsNumbsReplace(global_Function, row + 1);
    }
  }
}



function rangeColumnByNumb_test() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var colnm = randomInteger(1, 9)
  var rangeColumn = rangeColumnByNumber(sheet, colnm);

  if (colnm !== rangeColumn.getColumn()) {
    Logger.log('Номер столбца !== ' + colnm);
  }
  else {
    Logger.log('rangeColumnByNumber_test = OK');
  }
  return true;
}

function rangeColumnByNumber(sheet, numb) {
  // вернуть диапазон столбца по номеру столбца
  var range = sheet.getRange("A:A");
  var rowsCount = range.getNumRows();
  return sheet.getRange(1, numb, rowsCount)
}

function randomInteger(min, max) {
  // случайное число от min до (max+1)
  let rand = min + Math.random() * (max + 1 - min);
  return Math.floor(rand);
}


function getsheetById_test() {
  id = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getGridId();
  var sheet = sheetById(id);
  Logger.log(sheet.getName());
}

function sheetById(id) {
  // вернуть лист по id
  return SpreadsheetApp.getActive().getSheets().filter(
    function (s) {
      return s.getSheetId() === id;
    }
  )[0];
}

function array2d2RangeFormulas_Test() {

  var sheet = SpreadsheetApp.getActive().getSheetByName('Тест скрипта');
  var cellu = sheet.getRange(1, 29);
  var range = sheet.getRange(1, 28, 3);
  var a2dim = range.getFormulas();
  // var a2dim = [
  //   ['=1'],
  //   ['=2']];

  array2d2RangeFormulas(cellu, a2dim);
}

function array2d2RangeFormulas(cell, a2d) {

  // массив 2мерный вставить на лист формулы

  var sheet_id = cell.getGridId();
  var sheet_ob = sheetById(sheet_id);
  var row_numb = cell.getRow();
  var col_numb = cell.getColumn();
  var range = sheet_ob.getRange(row_numb, col_numb, a2d.length, a2d[0].length);
  range.activate();
  range.setFormulas(a2d);
}
