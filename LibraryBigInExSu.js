// Библиотека методов InExSu

const DIGITS_COMMA_POINT = '0123456789,.';
const DIGITS_COMMA_POINT_SPACE = '0123456789,. ';

function Range_Rows_Test() {
  let ssheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');

  let rColu1 = ssheet.getRange('D1:D2');
  let a2 = rColu1.getValues();
  Logger.log(a2);

  let rColu2 = ssheet.getRange('D1:E2');
  a2 = rColu2.getValues();
  Logger.log(a2);

}

function Array2D_Update_by_Map_Test() {
  let a2d_New = [
    ['CodeSour', 'ValueSour'],
    ['0', 'Новый '],
    ['1', 'Новый 2']
  ];
  let a2d_Old = [
    ['CodeUpda', 'ValueUpda'],
    ['0', 'Старый'],
    ['1', 'Старый 2']
  ];
  let column_code = 0;
  let map_codes = Array2D_Column_2_Map(a2d_New, 0);
  let array2d_columns = [[1, 1]];
  // Logger.log('a2d_Update ' + a2d_Old);

  a2d_New = Array2D_Update_by_Map(a2d_New, a2d_Old, column_code, map_codes, array2d_columns, 'Log');

  // Logger.log('a2d_New ' + a2d_New);
  // Logger.log('a2d_Old ' + a2d_Old);
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
  let array2d_ret = JSON.parse(JSON.stringify(array2d_Old));

  let code = '';
  let row_New = -1;
  let row_Old = 0;
  let col_New = -1;
  let col_Old = -1;

  //var a2d_log = [['Код', 'Строка', 'Столбец', 'Было', 'Стало']];
  let a2d_log = [['Лог обновления', '', Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss' мск'"), '', '']];
  a2d_log.push(['', '', '', '', '']);
  a2d_log.push(['Код', 'Строка', 'Столбец', 'Было', 'Стало']);

  let col = '';
  let was_new = '';
  let now_new = '';
  let was_old = '';
  let now_old = '';

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
    let sheet_logit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_log_name);
    sheet_logit.clear();

    cell = sheet_logit.getRange(1, 1);
    array2d2Range(cell, a2d_log);
  }

  return array2d_ret;

}

function SheetNameExists(sheetName) {
  /* существует ли лист*/
  let spread = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spread.getSheetByName(sheetName);
  if (sheet) {
    return True;
  }
};

function SheetDuplicate(sheetName) {
  /*  let spreadsheet = SpreadsheetApp.getActive();
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
  let spread = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spread.getSheetByName(sheetName);
  if (sheet) {
    spread.deleteSheet(sheet);
  }
};

function getsheetById_test() {
  id = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getGridId();
  let sheet = sheetById(id);
  // Logger.log(sheet.getName());
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
  let a2 = [
    ['1', '2']
  ]
  let col = Array2D_Column_Find_In_Row(a2, 0, '3');
  //  // Logger.log(col);
  col = Array2D_Column_Find_In_Row(a2, 0, '2');
  //  // Logger.log(col);
  a2 = [
    ['1', '2', '33'],
    ['4', '4', '3']
  ]
  col = Array2D_Column_Find_In_Row(a2, 1, '3');
  // Logger.log(col);
}

function Array2D_Column_Find_In_Row(array2d, row, string_find) {
  // в двумернном массиве, в строке найти значение, вернуть номер столбца или -1
  let val = ''
  for (var column = 0; column < array2d[0].length; column++) {
    val = array2d[row][column];
    // // Logger.log(val);
    if (array2d[row][column] == string_find) {
      return column;
    }
  }
  return -1;
}

function Range_Rows(range_In, rows_count) {

  // вернуть строки диапазона

  // Parent двухходовочка
  let sheet_id = range_In.getGridId();
  let sheet_ob = sheetById(sheet_id);

  let row_number = range_In.getRow();
  let column_number = range_In.getColumn(); //starting column position for this range
  let columns_count = range_In.getNumColumns();

  return sheet_ob.getRange(row_number, column_number, rows_count, columns_count);
}

function Map_from_2_Arrays1D_Test() {
  a1_sour = ['1', '2', '3'];
  a1_upda = ['4', '3', '2'];
  let map = Map_from_2_Arrays1D(a1_sour, a1_upda);
}

function Map_from_2_Arrays1D(array1d_Update_Heads, array1d_Source_Heads) {
  // создать массив ассоциативный из двух массивов одномерных
  let index = -1;
  let map_return = new Map();
  let val = '';

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
  let a2 = [
    [0, 1, 2], // строка 0  
    [3, 4, 5] // строка 1  
  ];
  let map = Array2D_Column_2_Map(a2, 0);
  if (map.size == 2) {
    // Logger.log('Array2D_2_Map_Test = OK');
  } else {
    // Logger.log('Array2D_2_Map_Test = Ошибка');
  }
  // тестирую повтор ключа
  a2 = [
    [0, 1, 2], // строка 0  
    [0, 4, 5] // строка 1  
  ];
  map = Array2D_Column_2_Map(a2, 0);
  if (map.size == 1) {
    // Logger.log('Array2D_2_Map_Test повтор = OK');
  } else {
    // Logger.log('Array2D_2_Map_Test повтор = Ошибка');
  }
  // тестирую регистр символов
  a2 = [
    ["Z", 1, 2], // строка 0  
    ["z", 4, 5] // строка 1  
  ];
  map = Array2D_Column_2_Map(a2, 0);
  if (map.size == 2) {
    // Logger.log('Array2D_2_Map_Test регистр = OK');
  } else {
    // Logger.log('Array2D_2_Map_Test регистр = Ошибка');
  }

}

function Array2D_Column_2_Map(array2d, column_key) {
  // из массива 2мерного вернуть словарь - массив ассоциативный: значение столбца и номер строки
  let map_return = new Map();
  let val = '';
  for (var row = 0; row < array2d.length; row++) {
    val = String(array2d[row][column_key]);
    if (val.length > 0) {
      // если ключ повторяется, то обновится значение
      map_return.set(val, row);
    }
  }
  return map_return;
}

function Array2D_ColumnS_2_Map_Test() {
  // тест создания массива ассоциативного из 2мерного
  let a2 = [
    [0, 1, 2], // строка 0  
    [3, 4, 5]  // строка 1  
  ];
  let map = Array2D_ColumnS_2_Map(a2, 0, 2);
  if (map.size == 2) {
    // Logger.log('Array2D_2_Map_Test = OK');
  } else {
    // Logger.log('Array2D_2_Map_Test = Ошибка');
  }
  // тестирую повтор ключа
  a2 = [
    [0, 1, 2], // строка 0  
    [0, 4, 5] // строка 1  
  ];
  map = Array2D_ColumnS_2_Map(a2, 0, 1);
  if (map.size == 1) {
    // Logger.log('Array2D_2_Map_Test повтор = OK');
  } else {
    // Logger.log('Array2D_2_Map_Test повтор = Ошибка');
  }
  // тестирую регистр символов
  a2 = [
    ["Z", 1, 2], // строка 0  
    ["z", 4, 5] // строка 1  
  ];
  map = Array2D_ColumnS_2_Map(a2, 0, 2);
  if (map.size == 2) {
    // Logger.log('Array2D_2_Map_Test регистр = OK');
  } else {
    // Logger.log('Array2D_2_Map_Test регистр = Ошибка');
  }

}
function Array2D_ColumnS_2_Map(array2d, column_key, column_Value) {
  // из массива 2мерного вернуть словарь - массив ассоциативный: 
  // значение столбца ключа и значение в столбце column_Value

  let map_return = new Map();
  let val = '';

  for (var row = 0; row < array2d.length; row++) {

    val = String(array2d[row][column_key]);

    if (val.length > 0) {

      // если ключ повторяется, то обновится значение
      map_return.set(val, array2d[row][column_Value]);
    }
  }
  return map_return;
}


function Sheet2Array2DTest() {
  const oSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Logger.log(oSheet.getName())
  const array2 = Sheet2Array2D(oSheet);
  // Logger.log(array2)
}

function Sheet2Array2D(oSheet) {
  // лист все данные в массив двумерный
  return oSheet.getDataRange().getValues();
}

function Array1D_2_HeadNumbers_LookUp_Test() {
  let a2_Old = ['1', '2'];
  let a2_New = ['2', '3'];
  let a2_Ret = Array1D_2_HeadNumbers_LookUp(a2_Old, a2_New);
  // Logger.log(a2_Ret);
}

function Array1D_2_HeadNumbers_LookUp(array1d_Old, array1d_New) {

  // из двух 1мерных массивов создать массив 2мерный с соответствия номеров столбцов

  let value;
  let row_new;
  let array2D = [];

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
  let array2d = [
    [1, 1, 1],
    [2, 2, 2]
  ];
  let separat = '\n';
  let str_ret = Array2D_Column_2_String(array2d, 0, separat);
  // Logger.log(str_ret);
}

function Array2D_Column_2_String(array2d, column, separator) {
  // вернуть строку из столбца массива 2мерного

  let string_col = '';
  let string_new = '';

  for (var row = 0; row < array2d.length; row++) {
    string_col = array2d[row][column] + separator;
    string_new += string_col;
  }

  return string_new;
}

function array2d2Range_Test() {

  let sheet = SpreadsheetApp.getActive().getSheetByName('Ошибки');
  let cellu = sheet.getRange(1, 1);

  let a2dim = [
    [1, 2],
    [3, 4]
  ];

  array2d2Range(cellu, a2dim);
}

function array2d2Range(cell, a2d) {

  // массив 2мерный вставить на лист

  let sheet_id = cell.getGridId();
  let sheet_ob = sheetById(sheet_id);
  let row_numb = cell.getRow();
  let col_numb = cell.getColumn();

  sheet_ob.getRange(row_numb, col_numb, a2d.length, a2d[0].length).setValues(a2d);
}


function Array2d_ColumnsEquals_RowsDelete_Test(a2d) {

  let a2d_Old = [
    [1, 2],
    [2, 2],
    [3, 2],
    [3, 3]
  ];

  let a2d_New = Array2d_ColumnsEquals_RowsDelete(a2d_Old);

  // Logger.log(a2d_New);
}


function Array2d_ColumnsEquals_RowsDelete(a2d_In) {

  // массив удалить строки массива 2мерного с одинаковыми значениями

  // копировать массив 2мерный не просто
  let a2d = JSON.parse(JSON.stringify(a2d_In))

  let val = '';
  let equ = true;

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

  let a1a = ['Весна', 'Зима', 'Лето', 'Осень'];
  let a1b = ['Добро', 'Зима', 'Собака'];

  let a1z = Arrays1D_ValuesEqual(a1a, a1b);

  // Logger.log(a1z)
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

  let str_ret = '';
  let str_idx = '';

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

  let str_ret = '';
  let str_idx = '';

  for (var i = 0; i < string_in.length; i++) {

    str_idx = String(string_in[i]);

    if (string_chek.indexOf(str_idx) == -1) {

      str_ret += str_idx;

    }
  }
  return str_ret;
}


function string_2_float_if_Test() {
  // Logger.log('90000013547');
}

function string_2_float_if(string_in) {
  // определить, что строка число
  // если похоже на число, вернуть число, 
  // иначе вернуть оригинальную строку

  // сначала определяю наличие не нужных символов
  let other = symbols_NOT_in_template(string_in, DIGITS_COMMA_POINT_SPACE);

  if (other.length > 0) {

    return string_in;

  }

  return symbols_by_template(string_in, DIGITS_COMMA_POINT);
}

function Date_Time_Local() {
  // набросок
  let formattedDate = Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss");
  // Logger.log(formattedDate);
}


function replaceIfEnds_Test() {
  // Logger.log(replaceIfEnds('1,00', ',00', ''));
  // Logger.log(replaceIfEnds('1,01', ',00', ''));
}

function replaceIfEnds(stri, what, for_) {
  // заменить, если оканчивается
  if (stri.endsWith(what)) {
    return stri.replace(what, for_);
  }
  return stri;
}

function symbolsMore1RepeatsReplace_Test() {
  // Logger.log(symbolsMore1RepeatsReplace("1,0", ',', ';'));
  // Logger.log(symbolsMore1RepeatsReplace("1,2,0", ',', ';'));
}

function symbolsMore1RepeatsReplace(stri, find, repl) {

  //  если find > 1, заменить на repl

  let count = stri.split(find).length - 1;

  if (count > 1) {
    // replaceAll не поддержалась
    return stri.split(find).join(repl);
  }
  return stri;
}

function apostropheIfSymbolsMore1Repeats_Test() {
  // Logger.log(apostropheIfSymbolsMoreRepeats("1,0", ',', 1));
  // Logger.log(apostropheIfSymbolsMoreRepeats("1,2,0", ',', 1));
}

function apostropheIfSymbolsMoreRepeats(stri, find, mini) {

  //  если find встречается > mini, довавить в начало апостроф

  let count = stri.split(find).length - 1;

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

  let spreadSh = SpreadsheetApp.getActiveSpreadsheet();
  let sheetTmp = spreadSh.insertSheet();
  let rangeTmp = sheetTmp.getRange(1, 1);
  rng_New_In.copyTo(rangeTmp);

  // UsedRange
  let rng = sheetTmp.getDataRange();

}

function rangeApostropheAddIfMoreOne_Test() {
  let sheet = SpreadsheetApp.getActive().getSheetByName('Ошибки');
  // sheet.getRange(2, 2).setValue(',');
  // sheet.getRange(2, 3).setValue(',,');
  // let rng = sheet.getRange('B2:C2')
  // rangeApostropheAddIfMoreOne(rng, ',');
  let rng = sheet.getRange('E1')
  rangeApostropheAddIfMoreOne(rng, ',');

  // Logger.log(sheet.getRange('E1').getValue());
}

function rangeApostropheAddIfMoreOne(rng, symb) {

  // проходом по ячейкам диапазона
  // значениям c двумя и более symb добавить спереди апостроф

  let sh_id = rng.getGridId();
  let sheet = sheetById(sh_id);

  let val = '';
  let rowStart = rng.getRow();
  let colStart = rng.getColumn();
  let row_Stop = rowStart + rng.getNumRows() - 1;
  let col_Stop = colStart + rng.getNumColumns() - 1;
  let pos_Frst = -1;
  let pos_Last = -1;

  for (var row = rowStart; row <= row_Stop; row++) {
    for (var col = colStart; col <= col_Stop; col++) {

      val = sheet.getRange(row, col).getValue();

      // Logger.log(val);

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
  let sheet = SpreadsheetApp.getActive().getSheetByName('Ошибки');
  let textFinder = sheet.createTextFinder(',')
    .matchEntireCell(false)
    .useRegularExpression(true);

  let a1_rng = textFinder.findAll();
  for (var key in a1_rng) {  // OK in V8
    let key = a1_rng[key];
    let val = key.getValue();
    // Logger.log("val = %s", val);
  }
}

//  ==='
function digitsSpacesKiller() {

  // в выделенных ячейках,содержащих только цифры, пробелы, системный разделитель десятичных чисел,
  // удалить пробел

  let rng = SpreadsheetApp.getActiveRange();

  if (rng === null) {
    // нет выделенного диапазона
  } else {

    let a2d = rng.getValues();

    a2d = arrayXdDigitsSpaceKiller(a2d, DIGITS_COMMA_POINT_SPACE);

    // вставить массив на лист

    array2d2Range(rng, a2d);
  }
}


function array2dDigitsSpaceKiller_Test() {
  let a1d = ['1 ,1', '', '1', 'z1'];
  a1d = arrayXdDigitsSpaceKiller(a1d, DIGITS_COMMA_POINT_SPACE);
  // Logger.log(a1d);
}


function arrayXdDigitsSpaceKiller(aXd, tmp) {

  // в массиве, в элементах, содержащих только:
  // цифры, пробелы, системный разделитель десятичных чисел - 
  // удалить пробел 

  let ele = '';

  for (var idx = 0; idx < aXd.length; idx++) {

    ele = String(aXd[idx]);

    if (digitWithSpace(ele, tmp)) {

      aXd[idx] = ele.replace(' ', '');

    }
  }

  return aXd;

}


function digitWithSpace_Test() {
  // Logger.log(digitWithSpace('', DIGITS_COMMA_POINT_SPACE));
  // Logger.log(digitWithSpace('1', DIGITS_COMMA_POINT_SPACE));
  // Logger.log(digitWithSpace('1 ,', DIGITS_COMMA_POINT_SPACE));
  // Logger.log(digitWithSpace('z1 ,', DIGITS_COMMA_POINT_SPACE));
}

function digitWithSpace(str, tmp) {

  // строка похожа на число с пробелом ?

  let smb = '';

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

  let a2d = [
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

  let stringReturn = '';

  let stringValue = String(value);

  let regexp = new RegExp('^' + String(symbol) + '+');

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

  // Logger.log('Лист:' + sheetName + ', формат активной ячейки ' + cell.getNumberFormat());
  // Logger.log('getA1Notation(): ' + cell.getA1Notation());
  // Logger.log('getValue(): ' + cell.getValue());
}

function getRangeColumnByNumb_test() {

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let colnm = randomInteger(1, 9)
  let rangeColumn = getRangeColumnByNumb(sheet, colnm);

  if (colnm !== rangeColumn.getColumn()) {
    // Logger.log('Номер столбца !== ' + colnm);
  }
  else {
    // Logger.log('getRangeColumnByNumb_test = OK');
  }
  return true;
}

function getRangeColumnByNumb(sheet, numb) {
  // вернуть диапазон столбца по номеру столбца
  let range = sheet.getRange("A:A");
  let rowsCount = range.getNumRows();
  return sheet.getRange(1, numb, rowsCount)
}

function randomInteger(min, max) {
  // случайное число от min до (max+1)
  let rand = min + Math.random() * (max + 1 - min);
  return Math.floor(rand);
}

function convertIfPossible_Test() {
  let value = '1,1';
  let wante = 1;
  let conve = convertIfPossible(value, parseFloat)
  if (conve != wante) {
    // Logger.log('convert2FloatCommaPointIfPossible: %s != %s', conve, wante);
  }

  value = '2,1z';
  wante = 2;
  conve = convertIfPossible(value, parseFloat)
  if (conve != wante) {
    // Logger.log('convertIfPossible: %s != %s', conve, wante);
  }

  value = '3.1';
  wante = 3.1;
  conve = convertIfPossible(value, parseFloat)
  if (conve != wante) {
    // Logger.log('convertIfPossible: %s != %s', conve, wante);
  }

  value = '4.1 Z';
  wante = 4.1;
  conve = convertIfPossible(value, parseFloat)
  if (conve != wante) {
    // Logger.log('convertIfPossible: %s != %s', conve, wante);
  }

  value = 'Z';
  wante = 'Z';
  conve = convertIfPossible(value, parseFloat)
  if (conve != wante) {
    // Logger.log('convertIfPossible: %s != %s', conve, wante);
  }
}


function convertIfPossible(value, method) {
  // преобразовать, испрользуя method, иначе вернуть value.
  let convert = method(value);
  return isNaN(convert) ? value : convert;
}

function convert2FloatCommaPointIfPossible_Test() {
  let value = '2 100 830,00';
  let wante = 2100830;
  let conve = convert2FloatCommaPointIfPossible(value);
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

  // для использования в map массива из диапазона
  if (Array.isArray(value_old)) {
    value_old = value_old.join();
  }

  if (digitsCommaPointSpace(value_old)) {
    if (value_old % 1 != 0) {
      let value_new = value_old.replace(/\s/g, '');
      value_new = value_new.replace(",", ".");
      value_new = convertIfPossible(value_new, parseFloat);
      return value_new;
    }
  }
  return value_old;
}

function isNumber_Test() {
  console.log('число  860', isNumber(860));
  console.log('строка 860', isNumber('860'));
  console.log('строка 860.0', isNumber('860.0'));
  console.log('строка 2 100 830,00', isNumber('2 100 830,00'));
}

function isNumber(str) {
  // является ли строка числом

  // if (typeof str != "string") return false // we only process strings!
  // // could also coerce to string: str = ""+str
  // return !isNaN(str) && !isNaN(parseFloat(str))

  // if (str % 1 == 0)
  //   return true;
  // else
  //   return false;

  return (str % 1 == 0);
}

function digitsCommaPointSpace(str) {

  // строка похожа на число(с запятой, точкой, пробелом) ?

  let smb = '';

  for (var pos = 0; pos < str.length; pos++) {

    smb = str[pos];

    if (!symbolInString(smb, DIGITS_COMMA_POINT_SPACE)) {

      return false;
    }
  }

  return true;

}

function a2Duplicates2a1_Test() {
  let a2 = [
    [1, 2],
    [2, 1]
  ];

  let a1 = a2Duplicates2a1(a2);

}

function a2Duplicates2a1(a2) {
  // поиск дубликатов в 2мерном массиве

  let a1 = a2.flat(Infinity);
  return a1Duplicates2a1(a1);
}

function a1Duplicates2a1_Test() {
  let a1 = [1, 1, 2]
  let du = a1Duplicates2a1(a1);
}

function a1Duplicates2a1(a1) {
  // вернуть дубликаты 1мерного массива в 1мерном массиве

  let dict = new Map();
  let dupl = [];
  let valu = '';

  for (let i = 0; i < a1.length; i++) {
    valu = a1[i];
    if (dict.has(valu)) {
      dupl.push(valu);
    } else {
      dict.set(valu, valu);
    }
  }
  return dupl;
}

function a2FindRowCol_Test() {
  let a2 = [
    ['1', '33'],
    ['4', '3']
  ]
  let a1 = a2FindRowCol(a2, '33');
  // Logger.log('a1 = ' + a1);
}

function a2FindRowCol(a2, val) {
  // найти в массиве val вернуть номера строк и столбца

  for (let row = 0; row <= a2.length; row++) {
    for (let col = 0; col <= a2[0].length; col++) {
      // Logger.log(a2[row][col]);
      if (a2[row][col] == val) {
        return [row, col];
      }
    }
  }
}

function rangeValues_2_Array() {
  // как выгдядит массив 2мерный из дипазона? вот так
  // [[Код 1С], []]
  const oSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let array2 = oSheet.getRange("A1:A2").getValues();
  Logger.log(array2);
  array2 = oSheet.getRange("A1:B2").getValues();
  Logger.log(array2);
  Logger.log(array2[0]);
  Logger.log(array2[0][0]);
  console.log(array2[0][0]);
}

function cellDigit_Test() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ошибки');
  const cell_ = sheet.getRange('A1')
  let val = '2 2,3';

  // val = val.replace(' ','');
  // val = val.replace(',','.');
  // val = parseFloat(val);
  val = convert2FloatCommaPointIfPossible(val);
  cell_.setValue(val);
}
