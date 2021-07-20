// https://docs.google.com/spreadsheets/d/1gEQ2Yz3lC9DRY4y2z_Ni6G6LWek_xTNDmBJEnrt5z7Y/edit?pli=1#gid=0
const global_Function = '=IFERROR(IF(W2 > 0; (N2 - W2) / N2;);)'
const globalColumn28 = 28;

function priceDownFillRUN() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('ПРОМ contract');
  sheet.activate();
  priceDownFill();
}

function priceDownFill_Test() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Тест скрипта');
  sheet.activate();
  priceDownFill();
}

function priceDownFill() {
  // заполнить на активном листе
  //  в столбце "снижение от НМЦ"
  // пустые ячейки формулами.
  // диапазон в массив 2мерный
  // из него создать 1мерный с номерами пустых строк
  // по 1мерному массиву заполнить ячейки
  var sheet = SpreadsheetApp.getActive().getActiveSheet();
  var rangeColumn = rangeColumnByNumber(sheet, globalColumn28);
  var a2dFormul = rangeColumn.getFormulas();
  var a2dValues = rangeColumn.getValues();
  var a2dFormulasAndValues = array2dCopy2LeftFromRightEmptyNot(a2dFormul,0, a2dValues,0);
  var a1d = array2dEmpty2array1d(a2dFormulasAndValues, 0);
  rangeCellsEmptyFormulasFill(sheet, a1d, globalColumn28, global_Function);
}

function array2dEmpty2array1d_Test(){

}
function array2dEmpty2array1d(a2d, column) {
  // из массива 2мерного создать 
  // 1мерный из номеров строк пустых элементов
  var a1d = [];
  for (var row = 0; row < a2d.length; row++) {
    if (a2d[row][column] === '') {
      row += 1;
      a1d.push(row);
    }
  }
  return a1d;
}

function rangeCellsEmptyFormulasFill_Test() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Тест скрипта');
  var array1dRowsNumbers = [2, 3];
  rangeCellsEmptyFormulasFill(sheet, array1dRowsNumbers, 28, global_Function);
  if (sheet.getRange(2, 28).getFormula() === '') {
    Logger.log("getRange(2,28).getFormula() === ''")
  }
}

function rangeCellsEmptyFormulasFill(sheet, array1dRowsNumbers, columnNumber, formula) {
  // проходом по массиву номеров строк
  // пустым ячейкам проставить формулу с подобранным номером строки
  var formulaRow = '';
  var rowNumber = 0;
  var cell;
  for (var idx = 0; idx < array1dRowsNumbers.length; idx++) {
    rowNumber = array1dRowsNumbers[idx]
    formulaRow = formulaRowsNumbsReplace(formula, rowNumber)
    cell = sheet.getRange(rowNumber, columnNumber);
    cell.setFormula(formulaRow);
  }
}

function formulaRowsNumbsReplaceTest() {
  var ret = formulaRowsNumbsReplace('=ЕСЛИОШИБКА(ЕСЛИ(W2 > 0; (N2 - W2) / (N2);"");"")', 99);
  if (ret !== '=ЕСЛИОШИБКА(ЕСЛИ(W99 > 0; (N99 - W99) / (N99);"");"")') {
    Logger.log("Не совпало");
  }
  else {
    Logger.log("ОК. Совпало");
  }
}
function formulaRowsNumbsReplace(formula, rowNumb) {
  //специфично для задачи
  // заменить буква латинская с цифрами везде на эту же букву с rowNumb
  var re = /([A-z])\d+/g;
  var reOk = re.test(formula);
  if (reOk) {
    return formula.replace(re, "$1" + rowNumb);
  }
  return formula
}


function array2dCopy2LeftFromRightEmptyNot_Test() {
  var a2dLeft_ = [
    ['1'],
    ['']];
  var a2dRight = [
    [],
    ['=']];
  var colLeft_ = 0;
  var colRight = 0;
  a2d_All = array2dCopy2LeftFromRightEmptyNot(a2dLeft_, colLeft_, a2dRight, colRight)
  if (a2d_All[1][0] != '=') {
    Logger.log('array2dCopy2LeftFromRightEmptyNot_Test false');
  }
  else {
    Logger.log('true  array2dCopy2LeftFromRightEmptyNot_Test ')
  }
}

function array2dCopy2LeftFromRightEmptyNot(a2dLeft_, colLeft_, a2dRight, colRight) {
  // копировать из столбца правого массива 2мерного 
  // в левый 
  // непустые значения
  // массивы должны быть одинаковой длины
  if (a2dLeft_.length !== a2dRight.length) {
    Logger.log('array2dCopy2LeftFromRightEmptyNot: a2dLeft_.length !== a2dRight.length');
  }
  else {
    // массив скопировать
    var a2dNew = a2dLeft_.map(function (arr) {
      return arr.slice();
    });
    for (var row = 0; row < a2dNew.length; row++) {
      // если слева пусто, а справа != пусто
      if (a2dNew[row][colLeft_] === '') {
        if (a2dRight[row][colRight] !== '') {
          a2dNew[row][colLeft_] = a2dRight[row][colRight]
        }
      }
    }
  }
  return a2dNew;
}
