// Библиотека методов InExSu

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

  Logger.log('a2d_New ' + a2d_New);
  Logger.log('a2d_Update ' + a2d_Old);
}


function Array2D_Update_by_Map(array2d_New, array2d_Old,
  column_code, map_codes, array2d_columns, sheet_log_name) {
  // обновить массив из другого массива по коду и соответствия столбцов
  // Проходом по столбцу ключа в массиве назначения				
  // 	Найти код в столбце источнике (словарь)			
  // 		Если найден		
  // 			Проходом по массиву соответствия номеров столбцов	
  // 				Обновить значения элементов текущей строки массива назначения

  // копировать массив 2мерный не просто
  var array2d_return = JSON.parse(JSON.stringify(array2d_Old))

  var row_New = -1;
  var row_Old = 0;
  var code = '';
  var column_New = -1;
  var column_Old = -1;

  //var a2d_log = [['Код', 'Строка', 'Столбец', 'Было', 'Стало']];
  var a2d_log = [['Лог обновления', '', Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss' мск'"), '', '']];
  a2d_log.push(['', '', '', '', '']);
  a2d_log.push(['Код', 'Строка', 'Столбец', 'Было', 'Стало']);

  var col = '';
  var was_new = '';
  var now_new = '';
  var was_orig = '';
  var now_orig = '';

  for (row_Old = 0; row_Old < array2d_return.length; row_Old++) {

    code = String(array2d_return[row_Old][column_code]);

    if (map_codes.has(code)) {

      row_New = map_codes.get(code);

      // проход по строкам массива соответствия номеров столбцов
      for (var row_columns = 0; row_columns < array2d_columns.length; row_columns++) {

        column_Old = array2d_columns[row_columns][0];
        column_New = array2d_columns[row_columns][1];

        if (code == '_0000019005' && column_New == 13) {
          var stop = '';
        }

        // было и стало в отчёт
        was_orig = array2d_return[row_Old][column_Old];
        now_orig = array2d_New[row_New][column_New];


        was_new = String(was_orig);
        now_new = String(now_orig);

        // из Excel вставляются числа с пробелами
        was_new = string_2_float_if(was_new);
        now_new = string_2_float_if(now_new);

        // гугл таблицы творчески меняют форматы при обмене массива с диапазоном
        was_new = replaceIfEnds(was_new, ',00', '');
        now_new = replaceIfEnds(now_new, ',00', '');

        // в массив попадает то #VALUE!, то #ЗНАЧ!
        if (was_new == '#ЗНАЧ!') { was_new = '#VALUE!' };
        if (now_new == '#ЗНАЧ!') { now_new = '#VALUE!' };

        if (was_new != now_new) {

          // гуглтаблица значения с двумя запятыми стремится преобразовать во время
          now_new = symbolsMore1RepeatsReplace(now_new, ',', ' / ');
          now_new = apostropheIfSymbolsMoreRepeats(now_new, ',', 1);

          if ((row_Old + 1) == 1177 && column_Old == 15 + 1) {
            var stop = '';
          }
          // заголовок столбца в отчёт
          col = array2d_New[0][column_New];

          if (sheet_log_name) {
            a2d_log.push([code, row_Old + 1, col, was_orig, now_new]);
          }

          array2d_return[row_Old][column_Old] = now_new;
        }
      }
    }
  }

  if (sheet_log_name) {
    // массив лога на лист
    var sheet_logit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_log_name);
    sheet_logit.clear();

    cell = sheet_logit.getRange(1, 1);
    // удалить строки из одинаковых значений по столбцам
    // a2d_log = Array2d_ColumnsEquals_RowsDelete(a2d_log);
    Array2d_2_Range(cell, a2d_log);
  }

  return array2d_return;

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

function getSheetById_test() {
  id = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getGridId();
  var sheet = SheetById(id);
  Logger.log(sheet.getName());
}

function SheetById(id) {
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
  var sheet_id = range_In.getGridId();
  var sheet_ob = SheetById(sheet_id);
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

function Array2d_2_Range_Test() {

  var sheet = SpreadsheetApp.getActive().getSheetByName('Ошибки');
  var cellu = sheet.getRange(1, 1);

  var a2dim = [
    [1, 2],
    [3, 4]
  ];

  Array2d_2_Range(cellu, a2dim);
}

function Array2d_2_Range(cell, a2d) {

  // массив 2мерный вставить на лист

  var sheet_id = cell.getGridId();
  var sheet_ob = SheetById(sheet_id);
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
  // float = '0123456789,'

  var str_ret = ''
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
  var other = symbols_NOT_in_template(string_in, '0123456789 ,');

  if (other.length > 0) {

    return string_in;

  }

  return symbols_by_template(string_in, '0123456789,');
}

function Date_Time_Local() {
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
