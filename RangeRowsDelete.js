function Range_Rows_Delete_by_Range_Test() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ошибки');
  sheet.getRange(1, 1).setValue('1');
  sheet.getRange(2, 1).setValue('2');
  var rng_Search = sheet.getRange('A1:A2')

  sheet.getRange(4, 1).setValue('2');
  sheet.getRange(5, 1).setValue('3');
  var rng_Delete = sheet.getRange('A4:A5');

  Range_Rows_Delete_by_Range(rng_Delete, 0, rng_Search, 0);

}

function EXMZ_Delete() {
  var rng_Delete = SpreadsheetApp.getActive().getSheetByName("сводная таблица (копия)").getRange('A:A')
  var rng_Search = SpreadsheetApp.getActive().getSheetByName("ЭХМЗ удалить").getRange('A:A')
  // запуск
  var arr_val_Delete = Range_Rows_Delete_by_Range(rng_Delete, 0, rng_Search, 0);
  // логи на лист
  var sheet_logit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  sheet_logit.getRange(1, 1).setValue('ЭХМЗ коды 1с строки удалены');
  // массив на листы
  sheet_logit.clear();
  if (arr_val_Delete[0] !== undefined) {
    sheet_logit.getRange(2, 1, arr_val_Delete.length, arr_val_Delete[0].length).setValues(arr_val_Delete);
  }
}

function Range_Rows_Delete_by_Range(rng_Delete, column_Delete, rng_Search, column_Search) {
  // удалить строки в диапазоне rng_Delete по совпаденю значений в column_Search
  // с использованием словаря - массива ассоциативного

  // диапазоны в массив
  var arr_Delete = rng_Delete.getValues();
  var arr_Search = rng_Search.getValues();

  // массив поиска в ассоциативный массив
  var map_Search = Array2D_2_Map(arr_Search, column_Search);

  var val_Delete = '';
  var row_Delete = '';
  var row_in_map = false;
  var id = rng_Delete.getGridId();
  var sheet_Dele = SheetById(id);
  var array1d_value_Deleted = [];
  var row_1st = rng_Delete.getRow();
  var row_Curr = 0;

  // удалять строки надо снизу
  for (var row = arr_Delete.length - 1; row >= 0; row--) {

    // row не отображается в отладчике
    row_Curr = row;

    val_Delete = String(arr_Delete[row][0]);

    if (val_Delete.length > 0) {

      row_in_map = map_Search.has(String(val_Delete));

      if (row_in_map) {
        // если диапазон начинается не с первой строки
        row_Delete = row + row_1st;
        // добавляю строку к массиву 2мерному
        array1d_value_Deleted.push([val_Delete]);
        sheet_Dele.deleteRow(row_Delete);
      }
    }
  }
  return array1d_value_Deleted;
}