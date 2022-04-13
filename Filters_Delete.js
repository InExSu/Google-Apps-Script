// https://ru.stackoverflow.com/questions/1175594/%D0%9A%D0%B0%D0%BA-%D1%83%D0%B4%D0%B0%D0%BB%D0%B8%D1%82%D1%8C-%D1%81%D0%BE%D1%85%D1%80%D0%B0%D0%BD%D1%8F%D0%B5%D0%BC%D1%8B%D0%B5-%D1%84%D0%B8%D0%BB%D1%8C%D1%82%D1%80%D1%8B-google-%D1%82%D0%B0%D0%B1%D0%BB%D0%B8%D1%86%D1%8B-google-sheet-api

function filters_delete() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var id = ss.getId();

  var id_sheet = ss.getActiveSheet().getSheetId();
  var substr = "Фильтр ";
  var arr_filterViewId = [];

  var json_obj = Sheets.Spreadsheets.get(id, {
    fields: 'sheets/filterViews',
  });

  for (var i = 0; i < json_obj.sheets.length; i++) {
    if (json_obj.sheets[i].filterViews != null) {
      for (var i2 = 0; i2 < json_obj.sheets[i].filterViews.length; i2++) {
        var current_filterView = json_obj.sheets[i].filterViews[i2];
        var sheetId = current_filterView.range.sheetId;
        if (sheetId === id_sheet) {
          var str_title = current_filterView.title;
          if (str_title.includes(substr)) {
            arr_filterViewId.push(JSON.stringify(current_filterView.filterViewId));
          }
        }

      } // end of 2 for
    }
  }

  Logger.log(arr_filterViewId);

  for (var x = 0; x < arr_filterViewId.length; x++) {
    var value_filterId = parseInt(arr_filterViewId[x], 10);
    Sheets.Spreadsheets.batchUpdate({
      "requests": [
        {
          "deleteFilterView": {
            "filterId": value_filterId
          }
        }
      ]
    }, id);
  }

  //Logger.log(JSON.stringify(json_obj.sheets[3].filterViews[0].range.sheetId));
}