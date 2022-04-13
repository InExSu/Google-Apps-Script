function ZG_grouping_Fast_RUN() {
  var user_choice = Browser.msgBox('Запустить обновление группировок ?')
  if (user_choice == "ok") {
    ZG_grouping_Fast();
    user_choice = Browser.msgBox('Результаты работы на листе ' + 'Log' + ' Активировать ?');
    if (user_choice == "ok") {
      SpreadsheetApp.getActive().getSheetByName('Log').activate();
    }
  }
}

function ZG_grouping_prepare_RUN() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B:B').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('сводная таблица (копия)'), true);
  spreadsheet.getRange('B:B').activate();
  spreadsheet.getRange('\'сводная таблица\'!B:B').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  ZG_grouping_Fast();
}
function ZG_grouping_Fast() {
  // Пройтись по кодам массива "ЗГ Группировка"
  // найти код в массиве "основная таблица"
  // заменить значение в ячейке столбца "Группировка ЗГ"
  // значением из массива столбца a2_column_groupce

  // 2021-05-06 Попов

  var sheet_group = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ЗГ Группировки');
  var sheet_pivot = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('сводная таблица (копия)');
  var sheet_logit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');

  sheet_logit.clearContents();
  var date_time = Date();
  sheet_logit.getRange(1, 1).setValue(date_time);
  sheet_logit.getRange(1, 2).setValue('ZG_grouping');
  sheet_logit.getRange(3, 1).setValue('Строка');
  sheet_logit.getRange(3, 2).setValue('Было');
  sheet_logit.getRange(3, 3).setValue('Стало');
  sheet_logit.setColumnWidth(1, 170);
  sheet_logit.setColumnWidth(2, 170);
  sheet_logit.setColumnWidth(3, 170);

  var range_pivot = sheet_pivot.getDataRange();
  var cell_value = range_pivot.getCell(1, 2).getValue();
  if (cell_value != 'Группировка ЗГ') {
    Logger.log('Выход. На листе сводная таблица, в ячейке 1,2 нет значения Группировка ЗГ');
    return;
  }

  range_pivot = sheet_pivot.getRange('A:B'); // лишнего не возьмёт
  range_group = sheet_group.getRange('A:E'); // лишнего не возьмёт
  var array_pivot = range_pivot.getValues();
  var array_group = range_group.getValues();

  // массив ассоциативный: код, строка номер
  var map_group_code1s = Array2D_Column_2_Map(array_group, 1);
  var map_group_code1s_size = map_group_code1s.size;
  Logger.log('map_group_code1s_rows = ' + map_group_code1s_size);
  var array_pivot_length = array_pivot.length;
  Logger.log('array_pivot_length = ' + array_pivot_length);
  var product_code1s = '';
  var row_group = 0;
  var cell_pivot;
  var product_group = '';
  var value_old = '';
  var row_free = 0;
  var array_log = [['', '', '']];
  array_log.pop();
  var log_rows;
  var log_cols;

  // проход по массиву сводной таблицы
  for (var row_pivot = 0; row_pivot < array_pivot_length; row_pivot++) {

    product_code1s = String(array_pivot[row_pivot][0]);

    if (product_code1s.length > 0) {

      // есть ли в словаре код продукта
      if (map_group_code1s.has(product_code1s)) {

        // взять из словаря номер строки
        row_group = map_group_code1s.get(product_code1s);

        // взять из словаря название группировки
        product_group = array_group[row_group][3] + ' ' + array_group[row_group][4];

        // определяю на сводной ячейку группировки 
        //cell_pivot = sheet_pivot.getRange(row_pivot, 2);
        // cell_old_value = cell_pivot.getValue()
        value_old = array_pivot[row_pivot][1];

        if (value_old !== product_group) {
          //cell_pivot.setValue(product_group);
          array_pivot[row_pivot][1] = product_group;
          // Logger.log('В строке ' + row_pivot + ' заменил: ' + cell_old_value + '\n на \n' + product_group);

          // записать на лист Log
          /*           row_free = 1 + sheet_logit.getLastRow();
                    sheet_logit.getRange(row_free, 1).setValue(row_pivot);
                    sheet_logit.getRange(row_free, 2).setValue(cell_old_value);
                    sheet_logit.getRange(row_free, 3).setValue(product_group);
           */
          // добавляю строку к массиву 2мерному
          array_log.push([row_pivot, value_old, product_group]);
        }
      }
    }
  }
  // массивы на листы
  sheet_pivot.getRange(1, 1, array_pivot.length, array_pivot[0].length).setValues(array_pivot);
  sheet_logit.getRange(4, 1, array_log.length, array_log[0].length).setValues(array_log);
}

