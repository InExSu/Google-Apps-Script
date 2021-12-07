// Библиотека методов InExSu

const DIGITS_COMMA_POINT = '0123456789,.';
const DIGITS_COMMA_POINT_SPACE = '0123456789 ,.';

function Array2D_Update_by_Map_Test() {
  var a2d_New = [
    ['CodeSour', 'ValueSour'],
    ['0', 'Новый '],
    ['1', 'Новый 2']
  ];
  var a2d_Old = [
    ['CodeUpda', 'ValueUpda'],
    ['0', 'Старый'],
    ['1', 'Старый 2']
  ];
  var column_code = 0;
  var map_codes = Array2D_2_Map(a2d_New, 0);
  var array2d_columns = [[1, 1]];
  Logger.log('a2d_Update ' + a2d_Old);

  var a2d_New = Array2D_Update_by_Map(a2d_New, a2d_Old, column_code, map_codes, array2d_columns, 'Log');

  // Logger.log('a2d_New ' + a2d_New);
  // Logger.log('a2d_Update ' + a2d_Old);
}


function Array2D_Update_by_Map(array2d_New, array2d_Old,
  column_code, map_codes, array2d_columns, sheet_log_name) {
  // обновить массив из другого массива по коду и соответствия столбцов
  // Проходом по столбцу ключа в массиве назначения				
  // 	Найти код в столбце источнике (словарь)			
  // 		Если найден		
  // 			Проходом по массиву соответствия номеров столбцов	
  // 				Обновить значения элементов текущей строки массива назначения

  // массив 2мерный копировать не просто
  var array2d_ret = JSON.parse(JSON.stringify(array2d_Old))

  var code = '';
  var row_New = -1;
  var row_Old = 0;
  var col_New = -1;
  var col_Old = -1;

  //var a2d_log = [['Код', 'Строка', 'Столбец', 'Было', 'Стало']];
  var a2d_log = [['Лог обновления', '', Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss' мск'"), '', '']];
  a2d_log.push(['', '', '', '', '']);
  a2d_log.push(['Код', 'Строка', 'Столбец', 'Было', 'Стало']);

  var col = '';
  var was_new = '';
  var now_new = '';
  var was_old = '';
  var now_old = '';

  for (row_Old = 0; row_Old < array2d_ret.length; row_Old++) {

    code = String(array2d_ret[row_Old][column_code]);

    if (map_codes.has(code)) {

      row_New = map_codes.get(code);

      // проход по строкам массива соответствия номеров столбцов
      for (var row_columns = 0; row_columns < array2d_columns.length; row_columns++) {

        col_Old = array2d_columns[row_columns][0];
        col_New = array2d_columns[row_columns][1];

        // было и стало в отчёт
        was_old = array2d_ret[row_Old][col_Old];
        now_old = array2d_New[row_New][col_New];

        was_new = String(was_old);
        now_new = String(now_old);

        // из Excel вставляются числа с пробелами
        was_new = string_2_float_if(was_new);
        now_new = string_2_float_if(now_new);

        // гугл таблицы творчески меняют форматы при обмене массива с диапазоном
        was_new = replaceIfEnds(was_new, ',00', '');
        now_new = replaceIfEnds(now_new, ',00', '');

        was_new = convert2FloatCommaPointIfPossible(was_new);
        now_new = convert2FloatCommaPointIfPossible(now_new);

        // в массив попадает то #VALUE!, то #ЗНАЧ!
        if (was_new == '#ЗНАЧ!') { was_new = '#VALUE!' };
        if (now_new == '#ЗНАЧ!') { now_new = '#VALUE!' };

        if (was_new != now_new) {

          // гуглтаблица значения с двумя запятыми стремится преобразовать во время
          // now_new = symbolsMore1RepeatsReplace(now_new, ',', ' / ');
          // now_new = apostropheIfSymbolsMoreRepeats(now_new, ',', 1);

          // заголовок столбца в отчёт
          col = array2d_New[0][col_New];

          if (sheet_log_name) {
            a2d_log.push([code, row_Old + 1, col, was_old, now_new]);
          }

          array2d_ret[row_Old][col_Old] = now_new;
        }
      }
    }
  }

  if (sheet_log_name) {
    // массив лога на лист
    var sheet_logit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_log_name);
    sheet_logit.clear();

    cell = sheet_logit.getRange(1, 1);
    array2d2Range(cell, a2d_log);
  }

  return array2d_ret;

}

function SheetNameExists(sheetName) {
  /* существует ли лист*/
  var spread = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spread.getSheetByName(sheetName);
  if (sheet) {
    return True;
  }
};

function SheetDuplicate(sheetName) {
  /*  var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('сводная таблица (копия)'), true);
    spreadsheet.deleteActiveSheet();
    spreadsheet.duplicateActiveSheet();*/
  if (SheetNameExists(sheetName)) {
    SheetNameDelete(sheetName);
    return spreadsheet.copy(sheetName);
  }
};

function SheetNameDelete(sheetName) {
  /* удалить лист по имени, если он есть*/
  var spread = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spread.getSheetByName(sheetName);
  if (sheet) {
    spread.deleteSheet(sheet);
  }
};

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


function Array2D_Column_Find_In_Row_Test() {
  a2 = [
    ['1', '2']
  ]
  col = Array2D_Column_Find_In_Row(a2, 0, '3');
  //  Logger.log(col);
  col = Array2D_Column_Find_In_Row(a2, 0, '2');
  //  Logger.log(col);
  a2 = [
    ['1', '2', '33'],
    ['4', '4', '3']
  ]
  col = Array2D_Column_Find_In_Row(a2, 1, '3');
  Logger.log(col);
}

function Array2D_Column_Find_In_Row(array2d, row, string_find) {
  // в двумернном массиве, в строке найти значение, вернуть номер столбца или -1
  var val = ''
  for (var column = 0; column < array2d[0].length; column++) {
    val = array2d[row][column];
    // Logger.log(val);
    if (array2d[row][column] == string_find) {
      return column;
    }
  }
  return -1;
}

function Range_Rows_Test() {
  var ssheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('сводная таблица (копия)');
  var rrange = ssheet.getRange('B2:D9');
  Logger.log(Range_Rows(rrange, 1).getValues());
}

function Range_Rows(range_In, rows_count) {

  // вернуть строки диапазона

  // Parent двухходовочка
  var sheet_id = range_In.getGridId();
  var sheet_ob = sheetById(sheet_id);

  var row_number = range_In.getRow();
  var column_number = range_In.getColumn(); //starting column position for this range
  var columns_count = range_In.getNumColumns();

  return sheet_ob.getRange(row_number, column_number, rows_count, columns_count);
}

function Map_from_2_Arrays1D_Test() {
  a1_sour = ['1', '2', '3'];
  a1_upda = ['4', '3', '2'];
  var map = Map_from_2_Arrays1D(a1_sour, a1_upda);
}

function Map_from_2_Arrays1D(array1d_Update_Heads, array1d_Source_Heads) {
  // создать массив ассоциативный из двух массивов одномерных
  var index = -1;
  var map_return = new Map();
  var val = '';

  for (var idx = 0; idx < array1d_Update_Heads.length; idx++) {

    val = String(array1d_Update_Heads[idx]);

    if (val.length > 0) {
      index = array1d_Source_Heads.indexOf(val);
      if (index > -1) {
        // если ключ повторяется, то обновится значение
        map_return.set(val, index);
      }
    }
  }
  return map_return;
}

function Array2D_2_Map_Test() {
  // тест создания массива ассоциативного из 2мерного
  var a2 = [
    [0, 1, 2], // строка 0  
    [3, 4, 5] // строка 1  
  ];
  var map = Array2D_2_Map(a2, 0);
  if (map.size == 2) {
    Logger.log('Array2D_2_Map_Test = OK');
  } else {
    Logger.log('Array2D_2_Map_Test = Ошибка');
  }
  // тестирую повтор ключа
  var a2 = [
    [0, 1, 2], // строка 0  
    [0, 4, 5] // строка 1  
  ];
  var map = Array2D_2_Map(a2, 0);
  if (map.size == 1) {
    Logger.log('Array2D_2_Map_Test повтор = OK');
  } else {
    Logger.log('Array2D_2_Map_Test повтор = Ошибка');
  }
  // тестирую регистр символов
  var a2 = [
    ["Z", 1, 2], // строка 0  
    ["z", 4, 5] // строка 1  
  ];
  var map = Array2D_2_Map(a2, 0);
  if (map.size == 2) {
    Logger.log('Array2D_2_Map_Test регистр = OK');
  } else {
    Logger.log('Array2D_2_Map_Test регистр = Ошибка');
  }

}

function Array2D_2_Map(array2d, column_key) {
  // из массива 2мерного вернуть словарь - массив ассоциативный: значение столбца и номер строки
  var map_return = new Map();
  var val = '';
  for (var row = 0; row < array2d.length; row++) {
    val = String(array2d[row][column_key]);
    if (val.length > 0) {
      // если ключ повторяется, то обновится значение
      map_return.set(val, row);
    }
  }
  return map_return;
}

function Sheet2Array2DTest() {
  const oSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log(oSheet.getName())
  const array2 = Sheet2Array2D(oSheet);
  Logger.log(array2)
}

function Sheet2Array2D(oSheet) {
  // лист все данные в массив двумерный
  return oSheet.getDataRange().getValues();
}

function Array1D_2_HeadNumbers_LookUp_Test() {
  var a2_Old = ['1', '2'];
  var a2_New = ['2', '3'];
  var a2_Ret = Array1D_2_HeadNumbers_LookUp(a2_Old, a2_New);
  Logger.log(a2_Ret);
}

function Array1D_2_HeadNumbers_LookUp(array1d_Old, array1d_New) {

  // из двух 1мерных массивов создать массив 2мерный с соответствия номеров столбцов

  var value;
  var row_new;
  var array2D = [];

  for (var row_old = 0; row_old < array1d_Old.length; row_old++) {

    value = array1d_Old[row_old];

    if (String(value).length > 0) {

      row_new = array1d_New.indexOf(value);

      if (row_new > -1) {
        array2D.push([row_old, row_new]);
      }
    }
  }

  return array2D;
}

function Array2D_Column_2_String_Test() {
  var array2d = [
    [1, 1, 1],
    [2, 2, 2]
  ];
  var separat = '\n';
  var str_ret = Array2D_Column_2_String(array2d, 0, separat);
  Logger.log(str_ret);
}

function Array2D_Column_2_String(array2d, column, separator) {
  // вернуть строку из столбца массива 2мерного

  var string_col = '';
  var string_new = '';

  for (var row = 0; row < array2d.length; row++) {
    string_col = array2d[row][column] + separator;
    string_new += string_col;
  }

  return string_new;
}

function array2d2Range_Test() {

  var sheet = SpreadsheetApp.getActive().getSheetByName('Ошибки');
  var cellu = sheet.getRange(1, 1);

  var a2dim = [
    [1, 2],
    [3, 4]
  ];

  array2d2Range(cellu, a2dim);
}

function array2d2Range(cell, a2d) {

  // массив 2мерный вставить на лист

  var sheet_id = cell.getGridId();
  var sheet_ob = sheetById(sheet_id);
  var row_numb = cell.getRow();
  var col_numb = cell.getColumn();

  sheet_ob.getRange(row_numb, col_numb, a2d.length, a2d[0].length).setValues(a2d);
}


function Array2d_ColumnsEquals_RowsDelete_Test(a2d) {

  var a2d_Old = [
    [1, 2],
    [2, 2],
    [3, 2],
    [3, 3]
  ];

  var a2d_New = Array2d_ColumnsEquals_RowsDelete(a2d_Old);

  Logger.log(a2d_New);
}


function Array2d_ColumnsEquals_RowsDelete(a2d_In) {

  // массив удалить строки массива 2мерного с одинаковыми значениями

  // копировать массив 2мерный не просто
  var a2d = JSON.parse(JSON.stringify(a2d_In))

  var val = '';
  var equ = true;

  for (var row = a2d.length - 1; row >= 0; row--) {

    val = String(a2d[row][0]);

    for (var col = 1; col < a2d[0].length; col++) {

      if (val !== String(a2d[row][col])) {

        equ = false;
        break;

      }
    }
    if (equ) {
      // удалить текущую строку
      a2d.splice(row, 1); // remove row, 1 - колво строк
    }

    equ = true;

  }

  return a2d;
}


function Arrays1D_ValuesEqual_Test() {

  var a1a = ['Весна', 'Зима', 'Лето', 'Осень'];
  var a1b = ['Добро', 'Зима', 'Собака'];

  var a1z = Arrays1D_ValuesEqual(a1a, a1b);

  Logger.log(a1z)
}

function Arrays1D_ValuesEqual(a1d_1, a1d_2) {

  // вернуть массив совпавших значений в 1мерных массивах

  return a1d_1.filter(function (obj) { return a1d_2.indexOf(obj) >= 0; });

}

function Array1D_2_String(a1d, sepa) {

  // массив 1мерный в строку

}

function symbols_by_template(string_in, string_check) {

  // вернуть строку из символов string_in, которые ЕСТЬ в string_chek 
  // float = DIGITS_COMMA_POINT

  var str_ret = '';
  var str_idx = '';

  for (var i = 0; i < string_in.length; i++) {

    str_idx = String(string_in[i]);

    if (string_check.indexOf(str_idx) > -1) {

      str_ret += str_idx;

    }
  }

  return str_ret;
}


function symbols_NOT_in_template(string_in, string_chek) {

  // вернуть строку из символов string_in, которых НЕТ в string_chek 

  var str_ret = '';
  var str_idx = '';

  for (var i = 0; i < string_in.length; i++) {

    str_idx = String(string_in[i]);

    if (string_chek.indexOf(str_idx) == -1) {

      str_ret += str_idx;

    }
  }
  return str_ret;
}


function string_2_float_if_Test() {
  Logger.log('90000013547');
}

function string_2_float_if(string_in) {
  // определить, что строка число
  // если похоже на число, вернуть число, 
  // иначе вернуть оригинальную строку

  // сначала определяю наличие не нужных символов
  var other = symbols_NOT_in_template(string_in, DIGITS_COMMA_POINT_SPACE);

  if (other.length > 0) {

    return string_in;

  }

  return symbols_by_template(string_in, DIGITS_COMMA_POINT);
}

function Date_Time_Local() {
  // набросок
  var formattedDate = Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss");
  Logger.log(formattedDate);
}


function replaceIfEnds_Test() {
  Logger.log(replaceIfEnds('1,00', ',00', ''));
  Logger.log(replaceIfEnds('1,01', ',00', ''));
}

function replaceIfEnds(stri, what, for_) {
  // заменить, если оканчивается
  if (stri.endsWith(what)) {
    return stri.replace(what, for_);
  }
  return stri;
}

function symbolsMore1RepeatsReplace_Test() {
  Logger.log(symbolsMore1RepeatsReplace("1,0", ',', ';'));
  Logger.log(symbolsMore1RepeatsReplace("1,2,0", ',', ';'));
}

function symbolsMore1RepeatsReplace(stri, find, repl) {

  //  если find > 1, заменить на repl

  var count = stri.split(find).length - 1;

  if (count > 1) {
    // replaceAll не поддержалась
    return stri.split(find).join(repl);
  }
  return stri;
}

function apostropheIfSymbolsMore1Repeats_Test() {
  Logger.log(apostropheIfSymbolsMoreRepeats("1,0", ',', 1));
  Logger.log(apostropheIfSymbolsMoreRepeats("1,2,0", ',', 1));
}

function apostropheIfSymbolsMoreRepeats(stri, find, mini) {

  //  если find встречается > mini, довавить в начало апостроф

  var count = stri.split(find).length - 1;

  if (count > mini) {
    return "'" + stri;
  }
  return stri;
}


function Array2DFormRangeWithApostorphes(rng_New_In) {

  // гуглтаблица, при вставке диапазона в массив (getValues) 
  // пытается преобразовать значения в двойными запятыми в формат даты, 
  // копировать диапазон в новый лист и всем ячейкам, не пустым, проставить апостроф
  // имменно в ДИАПАЗОНЕ (ибо в массив попадут уже "улучшенные" значения).
  // вернуть массив с апострофами, а лист удалить

  var spreadSh = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTmp = spreadSh.insertSheet();
  var rangeTmp = sheetTmp.getRange(1, 1);
  rng_New_In.copyTo(rangeTmp);

  // UsedRange
  var rng = sheetTmp.getDataRange();

}

function rangeApostropheAddIfMoreOne_Test() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Ошибки');
  // sheet.getRange(2, 2).setValue(',');
  // sheet.getRange(2, 3).setValue(',,');
  // var rng = sheet.getRange('B2:C2')
  // rangeApostropheAddIfMoreOne(rng, ',');
  var rng = sheet.getRange('E1')
  rangeApostropheAddIfMoreOne(rng, ',');

  Logger.log(sheet.getRange('E1').getValue());
}

function rangeApostropheAddIfMoreOne(rng, symb) {

  // проходом по ячейкам диапазона
  // значениям c двумя и более symb добавить спереди апостроф

  var sh_id = rng.getGridId();
  var sheet = sheetById(sh_id);

  var val = '';
  var rowStart = rng.getRow();
  var colStart = rng.getColumn();
  var row_Stop = rowStart + rng.getNumRows() - 1;
  var col_Stop = colStart + rng.getNumColumns() - 1;
  var pos_Frst = -1;
  var pos_Last = -1;

  for (var row = rowStart; row <= row_Stop; row++) {
    for (var col = colStart; col <= col_Stop; col++) {

      val = sheet.getRange(row, col).getValue();

      Logger.log(val);

      if (val.length > 0) {

        pos_Frst = val.indexOf(symb);
        pos_Last = val.lastIndexOf(symb);

        if (pos_Frst != pos_Last) {

          sheet.getRange(row, col).setValue("'" + val);

        }
      }
    }
  }
  return rng;
}


function textFinder_test() {
  // набросок
  var sheet = SpreadsheetApp.getActive().getSheetByName('Ошибки');
  var textFinder = sheet.createTextFinder(',')
    .matchEntireCell(false)
    .useRegularExpression(true);

  var a1_rng = textFinder.findAll();
  for (var key in a1_rng) {  // OK in V8
    var key = a1_rng[key];
    var val = key.getValue();
    Logger.log("val = %s", val);
  }
}

//  ==='
function digitsSpacesKiller() {

  // в выделенных ячейках,содержащих только цифры, пробелы, системный разделитель десятичных чисел,
  // удалить пробел

  var rng = SpreadsheetApp.getActiveRange();

  if (rng === null) {
    // нет выделенного диапазона
  } else {

    var a2d = rng.getValues();

    a2d = arrayXdDigitsSpaceKiller(a2d, DIGITS_COMMA_POINT_SPACE);

    // вставить массив на лист

    array2d2Range(rng, a2d);
  }
}


function array2dDigitsSpaceKiller_Test() {
  var a1d = ['1 ,1', '', '1', 'z1'];
  a1d = arrayXdDigitsSpaceKiller(a1d, DIGITS_COMMA_POINT_SPACE);
  Logger.log(a1d);
}


function arrayXdDigitsSpaceKiller(aXd, tmp) {

  // в массиве, в элементах, содержащих только:
  // цифры, пробелы, системный разделитель десятичных чисел - 
  // удалить пробел 

  var ele = '';

  for (var idx = 0; idx < aXd.length; idx++) {

    ele = String(aXd[idx]);

    if (digitWithSpace(ele, tmp)) {

      aXd[idx] = ele.replace(' ', '');

    }
  }

  return aXd;

}


function digitWithSpace_Test() {
  Logger.log(digitWithSpace('', DIGITS_COMMA_POINT_SPACE));
  Logger.log(digitWithSpace('1', DIGITS_COMMA_POINT_SPACE));
  Logger.log(digitWithSpace('1 ,', DIGITS_COMMA_POINT_SPACE));
  Logger.log(digitWithSpace('z1 ,', DIGITS_COMMA_POINT_SPACE));
}

function digitWithSpace(str, tmp) {

  // строка похожа на число с пробелом ?

  var smb = '';

  for (var pos = 0; pos < str.length; pos++) {

    smb = str[pos];

    if (!symbolInString(smb, tmp)) {

      return false;
    }
  }

  return true;

}


function symbolInString(smb, str) {

  // символ в строке ?

  if (str.indexOf(smb) < 0) {

    return false;
  }

  return true;
}

function array2dColumnSymbolsLeading_Test() {

  var a2d = [
    ['01', '02'],
    ['03', '04']
  ];

  array2dColumnSymbolsLeading(a2d, 1, 0);

  return a2d
}

function array2dColumnSymbolsLeading(array2d, column, symbol) {
  // проходом по массиву, по столбцу, убрать лидирующие символы
  for (var row = 0; row < array2d.length; row++) {
    array2d[row][column] = stringSymbolsLeadingDelete(array2d[row][column], symbol)
  }
}


function stringSymbolsLeadingDelete(value, symbol) {

  // лидирующие символы удалить

  var stringReturn = '';

  var stringValue = String(value);

  var regexp = new RegExp('^' + String(symbol) + '+');

  if (stringValue[0] === String(symbol)) {

    stringReturn = stringValue.replace(regexp, '')

  } else {

    stringReturn = stringValue;

  }

  return stringReturn;

}

function cellActiveInfo() {

  // информация об активной ячейки активного листа

  sheet = SpreadsheetApp.getActive().getActiveSheet();
  sheetName = sheet.getName()
  cell = sheet.getActiveCell();

  Logger.log('Лист:' + sheetName + ', формат активной ячейки ' + cell.getNumberFormat());
  Logger.log('getA1Notation(): ' + cell.getA1Notation());
  Logger.log('getValue(): ' + cell.getValue());
}

function getRangeColumnByNumb_test() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var colnm = randomInteger(1, 9)
  var rangeColumn = getRangeColumnByNumb(sheet, colnm);

  if (colnm !== rangeColumn.getColumn()) {
    Logger.log('Номер столбца !== ' + colnm);
  }
  else {
    Logger.log('getRangeColumnByNumb_test = OK');
  }
  return true;
}

function getRangeColumnByNumb(sheet, numb) {
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

function convertIfPossible_Test() {
  var value = '1,1';
  var wante = 1;
  var conve = convertIfPossible(value, parseFloat)
  if (conve != wante) {
    Logger.log('convert2FloatCommaPointIfPossible: %s != %s', conve, wante);
  }

  value = '2,1z';
  wante = 2;
  conve = convertIfPossible(value, parseFloat)
  if (conve != wante) {
    Logger.log('convertIfPossible: %s != %s', conve, wante);
  }

  value = '3.1';
  wante = 3.1;
  conve = convertIfPossible(value, parseFloat)
  if (conve != wante) {
    Logger.log('convertIfPossible: %s != %s', conve, wante);
  }

  value = '4.1 Z';
  wante = 4.1;
  conve = convertIfPossible(value, parseFloat)
  if (conve != wante) {
    Logger.log('convertIfPossible: %s != %s', conve, wante);
  }

  value = 'Z';
  wante = 'Z';
  conve = convertIfPossible(value, parseFloat)
  if (conve != wante) {
    Logger.log('convertIfPossible: %s != %s', conve, wante);
  }
}


function convertIfPossible(value, method) {
  // преобразовать, испрользуя method, иначе вернуть value.
  var convert = method(value);
  return isNaN(convert) ? value : convert;
}

function convert2FloatCommaPointIfPossible_Test() {
  var value = '1,1';
  var wante = 1.1;
  var conve = convert2FloatCommaPointIfPossible(value);
  if (conve != wante) {
    Logger.log('convert2FloatCommaPointIfPossible: %s != %s', conve, wante);
  }

  value = '2,1z';
  wante = '2,1z';
  conve = convert2FloatCommaPointIfPossible(value);
  if (conve != wante) {
    Logger.log('convert2FloatCommaPointIfPossible: %s != %s', conve, wante);
  }

  value = '3.1';
  wante = 3.1;
  conve = convert2FloatCommaPointIfPossible(value);
  if (conve != wante) {
    Logger.log('convert2FloatCommaPointIfPossible: %s != %s', conve, wante);
  }

  value = '4.1 Z';
  wante = '4.1 Z';
  conve = convert2FloatCommaPointIfPossible(value);
  if (conve != wante) {
    Logger.log('convert2FloatCommaPointIfPossible: %s != %s', conve, wante);
  }

  value = 'Z';
  wante = 'Z';
  conve = convert2FloatCommaPointIfPossible(value);
  if (conve != wante) {
    Logger.log('convert2FloatCommaPointIfPossible: %s != %s', conve, wante);
  }
}

function convert2FloatCommaPointIfPossible(value_old) {
  // конвертировать в число с плавающей точкой,
  // с учётом запятой и точки
  // сначала убедиться, что в строке только нужные символы
  
  if (digitsCommaPointSpace(value_old)) {
    var value_new = value_old.replace(",", ".");
    value_new = convertIfPossible(value_new, parseFloat);
    return value_new;
  }
  return value_old;
}


function digitsCommaPointSpace(str) {

  // строка похожа на число(с запятой, точкой, пробелом) ?

  var smb = '';

  for (var pos = 0; pos < str.length; pos++) {

    smb = str[pos];

    if (!symbolInString(smb, DIGITS_COMMA_POINT_SPACE)) {

      return false;
    }
  }

  return true;

}

function columnLetters2Number_Test() {

  var numb = columnLetters2Number("AM");
  Logger.log(numb);

};

function columnLetters2Number(letters) {
  // по буквам столбца вернуть номер
  var addre = letters + 1;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var numbe = sheet.getRange(addre).getColumn();
  return numbe;
};


function sheetColumnValueRowLastNumber(range) {
  // принимает  диапазон,
  // возвращает номер последней непустой строки
  // идёт снизу вверх по массиву

  var array1d = range.getValues();
  for (var i = array1d.length - 1; i >= 0; i--) {
    if (array1d[i][0] != null && array1d[i][0] != '') {
      return i + 1;
    };
  };
};

function sheetsList_Test() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Оглавление');
  const cell = sheet.getRange(3, 1);
  const a1_except = ['Оглавление']
  sheetsList(cell, a1_except);
}

function sheetsList(cell, a1_except) {
  // вставить список листов начиная с ячейки cell
  // исключая названия в a1_except

  var shee = SpreadsheetApp.getActive().getSheets();
  var row_ = 0;

  for (var i = 0; i < shee.length; i++) {
    var nCel = cell.offset(row_, 0);
    var name = shee[i].getName();
    if (a1_except.indexOf(name)) {
      cellLink2Sheet(nCel, name);
      row_++;
    }
  }
}

function cellLink2Sheet_Test() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Оглавление')
  const cell = sheet.getRange(1, 3);
  const name = '1С';
  const rnge = 'A1';
  cellLink2Sheet(cell, name, rnge);
}

function cellLink2Sheet(cell, name, range) {
  // вставить в ячейку ссылку на лист и диапазон

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  var richText = SpreadsheetApp.newRichTextValue()
    .setText(name)
    //    .setLinkUrl('#gid=' + sheet.getSheetId() + '&range=' + range)
    .setLinkUrl('#gid=' + sheet.getSheetId())
    .build();
  cell.setRichTextValue(richText);
}

