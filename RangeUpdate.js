function Range_Update_by_Heads_RUN_Test() {

  Pivot_Duplicate();

  SpreadsheetApp.getActive().getSheetByName('ЗГ Обновление (копия)').activate();

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

  var a2d_Old = rng_Old.getValues();
  var a2d_New = rng_New.getValues();
  var map_Sea = Array2D_2_Map(a2d_New, column_Key_New);

  // основное действие

  var a2d_Ret = Array2D_Update_by_Map(a2d_New, a2d_Old,
    column_Key_Old, map_Sea, a2d_columns, 'Log');

  // массив положить на лист
  Array2d_2_Range(rng_Old, a2d_Ret);
}