function Range_Update_by_Heads_RUN_Test() {

  // Pivot_Duplicate();

  SpreadsheetApp.getActive().getSheetByName('ЗГ Обновление').activate();

  Range_Update_by_Heads_RUN();
};

function Pivot_Duplicate() {

  // создать копию листа основная таблица

  // лист удалить
  var spread = SpreadsheetApp.getActive();
  var sheet_ = spread.getSheetByName('сводная таблица (копия)');
  if (sheet_) {
    spread.deleteSheet(sheet_);
  }

  // лист создать
  sheet_ = spread.getSheetByName('сводная таблица');
  spread.setActiveSheet(sheet_);
  spread.duplicateActiveSheet();
}


function Range_Update_by_Heads_RUN() {
  // вызываю по кнопке
  // подхватить с активного листа - диапазон назначения обновления - в массив Старый
  // подхватить с активного листа имя листа - источник
  // подхватить с активного листа имя столбца - код (одинаков для обоих)
  // подхватить с активного листа имя ячейки на листе источнике и закинуть диапазона в массив Новый
  // из диапазонов заголовков создать массив пар совпадений по названию
  // пусть пользователь проверит названия найденных столбцов совпадающих
  // вызвать функцию работы с диапазонами, которая:
  // из столбца кода источника создать словарь - код->номер строки в массиве источнике
  // проходом по столбцу кода назначения обновления,
  //   в словаре код есть ?
  //     если есть взять номер строки массива Источника
  //       проходом по массиву Столбцы - обновить значения в массиве Старый
  // вставить массив Старый на лист активный


  var spread = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_New = spread.getActiveSheet();

  // диапазон источника обновления вокруг ячейки, CurrentRegion
  var range_New = sheet_New.getRange("Range_New").getDataRegion();
  // ячейкам текстовый формат, чтобы значения вида '1,1,2005' не преобразовывались в строку даты время
  range_New.setNumberFormat('@');

  var sheet_Old_name = sheet_New.getRange("Sheet_Old").getValue();
  var sheet_Old = spread.getSheetByName(sheet_Old_name);

  if (!sheet_Old) {
    Browser.msgBox('Выход! Не найден лист ' + sheet_Old_name);
  } else {

    // UsedRange
    var range_Old = sheet_Old.getDataRange();

    // создать таблицу подстановки по именам столбцов
    // беру заговолок диапазона rng_New
    var range_row_1 = Range_Rows(range_New, 1);
    var array1d_Heads_New = range_row_1.getValues().flat();

    // беру заговолок диапазона rng_Update
    range_row_1 = Range_Rows(range_Old, 1);
    var array1d_Heads_Old = range_row_1.getValues().flat();

    // из двух 1мерных массив сделать 2мерный массив соответствия номеров столбцов
    var a2_Columns = Array1D_2_HeadNumbers_LookUp(array1d_Heads_Old, array1d_Heads_New);

    // для проверки пользователем
    var a1_Columns_Heads = Arrays1D_ValuesEqual(array1d_Heads_Old, array1d_Heads_New);
    var string_columns = a1_Columns_Heads.join(',\n');

    if (string_columns.length < 1) {

      Browser.msgBox('Выход. Не найдены совпадения в названиях столбцов');
      // Выход

    } else {

      // Пусть пользователь проверит соответствия номеров столбцов.
      var choice = ''
      // choice = Browser.msgBox('Найдены столбцы в обоих диапазонах: ' + string_columns, Browser.Buttons.YES_NO);
      if (choice == 'no') {

        // Выход по выбору пользователя

      } else {

        // найти в первой строке источника

        var a2_heads = Range_Rows(range_New, 1).getValues();
        var key_stri = sheet_New.getRange('Code_Name').getValue();

        var column_Key_New = Array2D_Column_Find_In_Row(a2_heads, 0, key_stri);
        if (column_Key_New < 0) {
          Logger.log('Выход. Не найден столбец кода в 1ой строке диапазоне источника');
        } else {

          // найти в первой строке назначения

          a2_heads = Range_Rows(range_Old, 1).getValues();
          key_stri = sheet_New.getRange('Code_Name').getValue();

          var column_Key_Old = Array2D_Column_Find_In_Row(a2_heads, 0, key_stri);
          if (column_Key_New < 0) {
            Logger.log('Выход. Не найден столбец кода в 1ой строке диапазоне назначения');
          } else {

            Range_Update_by_Heads(range_Old, column_Key_Old, range_New, column_Key_New, a2_Columns, 'Log')

            choice = Browser.msgBox('Показать лист Log', Browser.Buttons.YES_NO);
            if (choice == 'yes') {
              var sheet_act = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
              sheet_act.activate();
            }
          }
        }
      }
    }
  }
}


function Range_Update_by_Heads(rng_Old, column_Key_Old, rng_New, column_Key_New, a2d_columns, log_make) {

  // Обновить диапазон по совпадению в ключевых столбцах с учётом наименований столбцов
  // диапазоны в массивы

  // 2021-09-10
  // var a2d_Old = rng_Old.getValues();
  // var a2d_New = rng_New.getValues();
  var a2d_Old = range_2_ArrayValuesFormulas(rng_Old);
  var a2d_New = range_2_ArrayValuesFormulas(rng_New);

  // в листе "сводная таблица" в столбце Код нулей лидирующих нет,
  // поэтому удаляю нули из нового. Нет не буду удалять, сделаю нули в ячейках.
  //  array2dColumnSymbolsLeading(a2d_New, column_Key_New, '0');

  var map_Sea = Array2D_2_Map(a2d_New, column_Key_New);

  // основное действие

  var a2d_Ret = Array2D_Update_by_Map(a2d_New, a2d_Old,
    column_Key_Old, map_Sea, a2d_columns, 'Log');

  // массив положить на лист
  array2d2Range(rng_Old, a2d_Ret);
}

function rangeZerosAddFrontByFormat_Test() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Ошибки');
  //   var sheetName = sheet.getName();

  //   var cell = sheet.getRange(1, 1);
  //   cell.setValue("'1"); 
  //   var cellFormat_old = cell.getNumberFormat();
  //   cell.setNumberFormat('00000000000');
  // // чтобы формат обновился стопудово, его надо считать
  //   var cellFormat_new = cell.getNumberFormat();// НО - апсостроф исчезнет

  //   cell = sheet.getRange(2, 1);
  //   var cellFormat_old = cell.getNumberFormat();
  //   cell.setValue("'02");
  //   // cell.setNumberFormat('00000000000');
  //   var cellFormat_new = cell.getNumberFormat();

  var rng = sheet.getRange('A1:A3');
  rangeZerosAddFrontByFormat(rng, '00000000000');
  rng.setNumberFormat('@');
  rng.getNumberFormat(); // чтобы проявился setNumberFormats
}

function selectionNullFormatted() {
  var choice = Browser.msgBox('В выделенном диапазоне, \n ячейкам с форматом 00000000000 (11 нулей) \n добавит слева апостроф и недостающие нули.\n НЕ быстро.', Browser.Buttons.YES_NO);
  if (choice === 'yes') {
    var rng = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
    rangeZerosAddFrontByFormat(rng, '00000000000');
  }
}

function rangeZerosAddFrontByFormat(rng, formatPattern) {

  const start = new Date().getTime();

  // диапазону добавить лидирующие нули по формату
  // пусть пользователь сначала присвоит диапазону формат

  var sheet = sheetById(rng.getGridId());
  var rowStart = rng.getRow();
  var rowStop_ = rowStart + rng.getNumRows() - 1;
  var colStart = rng.getRow();
  var colStop_ = colStart + rng.getNumColumns() - 1;

  for (var row = rowStart; row <= rowStop_; row++) {
    for (var col = colStart; col <= colStop_; col++) {

      var cell = sheet.getRange(row, col);
      var formatCell = cell.getNumberFormat();

      if (formatCell === formatPattern) {

        var cellValue = String(cell.getValue());

        if (cellValue.length > 0) {

          var lenDiff = formatPattern.length - cellValue.length;

          if (lenDiff > 0) {

            var symb = formatPattern[0];

            var symbRepeat = symb.repeat(lenDiff);

            cellValue = "'" + symbRepeat + cellValue;

            cell.setValue(cellValue);
          }
        }
      }
    }
  }
  const end = new Date().getTime();
  Logger.log('rangeZerosAddFrontByFormat время работы: ${end - start}ms');
  Browser.msgBox('Время работы: ' + String((end - start) / 1000) + ', сек')
}

function range_2_ArrayValuesFormulas_Test() {

  var sheet = SpreadsheetApp.getActive().getSheetByName('Ошибки');
  sheet.getRange(1, 1).setValue("значение старое");
  sheet.getRange(1, 2).setFormula("=SUM(B3:B4)");

  var range = sheet.getRange("A1:B1");
  range.clearContent;

  var a2d = range_2_ArrayValuesFormulas(range);

  range.setValues(a2d);

  if (String(a2d[0][1]).charAt(0) !== '=') {
    Logger.log("range_2_ArrayValuesFormulas_Test(): String(a2d[0][1]).charAt(0) !== '='");
  }
}

function range_2_ArrayValuesFormulas(range) {

  // из диапазона вернуть массив значений и формул

  var a2d_formul = range.getFormulas();
  var a2d_values = range.getValues();
  a2d_Values_add_Formulas(a2d_values, a2d_formul);

  return a2d_values;
}

function a2d_Values_add_Formulas_test() {
  var a2d_values = [
    ["Tom", 1],
    ["Bill", 1],
  ];
  var a2d_formul = [
    ["=SUM(D3:D4)", 2],
    ["", 2],
  ];

  a2d_Values_add_Formulas(a2d_values, a2d_formul);

  if (a2d_values[0][0] == 'Tom') {
    Logger.log("a2d_Values_add_Formulas_test() - a2d_values[0][0] !== 'Tom'");
  }
}

function a2d_Values_add_Formulas(a2d_values, a2d_formul) {

  // в массив 2мерный значений вставить формулы
  // скопировать ячейки начинающиеся с =
  // массивы должны быть одинакового размера
  // потом сделать метод который копирует значения, если в позиции есть символ

  for (let row = 0; row < a2d_values.length; row++) {
    for (let col = 0; col < a2d_values[0].length; col++) {

      if (String(a2d_formul[row][col]).charAt(0) == '=') {

        a2d_values[row][col] = a2d_formul[row][col];

      }
    }
  }
}

function cells_Compare_Test() {

  // сравнить две ячейки
  var sheet = SpreadsheetApp.getActive().getSheetByName('Log')

  var cell_01 = sheet.getRange('D4');
  var valu_01 = parseFloat(cell_01.getValue());
  var cell_02 = sheet.getRange('E4');
  var valu_02 = parseFloat(cell_02.getValue());

  Logger.log(typeof valu_01);
  Logger.log(typeof valu_02);
}
