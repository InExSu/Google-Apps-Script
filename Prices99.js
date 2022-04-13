// // Обновление прайса
// // из формул с кодом сделать копию с артикулами.
// // На листе будут три таблицы:
// // слева    версия для печати - цены вбиваются пользователями.
// // в центре версия с артикулами.
// // По этим версиям, скрипт обновления цен в листе "сводная таблица"
// // создаст справа отчёт работы в третью таблицу.

// function priceArticoolPivot_RUN() {
//   // вызываю по кнопке на листе

//   let book_ = SpreadsheetApp.getActive();
//   let sheetPrice = book_.getSheetByName("Прайс без НДС");
//   let sheetPivot = book_.getSheetByName("сводная таблица");

//   // Убеждаюсь, что диапазон цен и артикулов на месте
//   if (priceRangeArticoolsCheck(sheetPrice)) {
//     if (priceRangePriceCheck(sheetPrice)) {

//       let a2PriceRange = sheetPrice.getRange('C:H').getValues();
//       let a2ArtiCRange = sheetPrice.getRange('L:Q').getValues();

//       if (sheetPivot.getRange("B1").getValue() == 'Артикул') {
//         if (sheetPivot.getRange("J1").getValue() == 'Цена, руб (Без НДС)') {

//           let a2ArtiColumn = sheetPivot.getRange('B:B').getValues();
//           let a2PricColumn = sheetPivot.getRange('L:L').getValues();

//           let a2PriceNew = a2RangeColumnSubst(a2PriceRange, a2ArtiCRange, a2PricColumn, a2ArtiColumn);

//           // создать лог различий старого и нового столбца

//           // столбец обновлённых цен на лист
//           // sheetPivot.getRange('L:L').setValues(a2PriceNew);

//         }
//       }
//     } else { Browser.msgBox('Диапазон цен не похож на ожидаемый'); }
//   } else { Browser.msgBox('Диапазон артикулов не похож на ожидаемый'); }
// }

// function a2RangeColumnSubst(a2PriceRange, a2ArtiCRange, a2PricColumn, a2ArtiColumn) {
//   // Скрипт получит на вход:
//   // - массив прайса с ценами, равноразмерный с 
//   // - массив прайса с артикулами
//   // - массив столбец цен "сводная таблица"
//   // - массив столбец артикулов "сводная таблица"

//   // Проходом по массиву диапазона прайса с артикулами,
//   // ищет значение ячейки в массиве артикулов "сводная таблица".
//   // Если находит, то берёт значение из массива цен и
//   // вставляет в массив цен "сводная таблица"
//   // вернёт массив цен "сводная таблица".

//   // массив 2мерный копировать не просто
//   let a2Return = JSON.parse(JSON.stringify(a2PricColumn));

//   for (let row = 0; row < a2ArtiCRange.length; row++) {
//     for (let col = 0; col < a2ArtiCRange[0].length; col++) {
//       let artic = a2ArtiCRange[row][col];
//       if (artic !== '') {

//       }
//     }
//   }
// }


// function priceRangeArticoolsCheck(sheet) {

//   let value = sheet.getRange("L7").getValue();
//   if (value !== 'ШМП-1') { return false };

//   value = sheet.getRange("Q7").getValue();
//   if (value !== 'МАГ-4') { return false };

//   return true

// }


// function priceRangePriceCheck(sheet) {

//   let value = sheet.getRange("C7").getValue();
//   if (value !== 'ШМП-1') { return false };

//   value = sheet.getRange("H7").getValue();
//   if (value !== 'МАГ-4') { return false };

//   return true

// }




// function rangeCtrlShiftDown_Test() {
//   let cell = SpreadsheetApp.getActive().getRange('S1:Y1');
//   rangeCtrlShiftDown(cell).activate();
// }
// function rangeCtrlShiftDown(range) {
//   // вернуть прямоугольный диапазон от range по последнюю строку со значеними
//   let sheet = sheetByRange(range);
//   let row_First = 1;
//   let row_Last_ = 1;
//   // ToDo: продолжить
//   // let row_Last_ = sheetColumnValueRowLastNumber()
//   // return sheet.getRange(range);
// }


// function sheetByRange(cell) {
//   // вернуть лист по диапазону
//   // как в Excel range.Parent

//   return sheetById(cell.getGridId());
// }


// function sheetById(id) {
//   // вернуть лист по id

//   return SpreadsheetApp.getActive().getSheets().filter(
//     function (s) {
//       return s.getSheetId() === id;
//     }
//   )[0];
// }


// function sheetColumnValueRowLastNumber(range) {
//   // принимает  диапазон,
//   // возвращает номер последней непустой строки
//   // идёт снизу вверх по массиву

//   let array1d = range.getValues();
//   for (let i = array1d.length - 1; i >= 0; i--) {
//     if (array1d[i][0] != null && array1d[i][0] != '') {
//       return i + 1;
//     };
//   };
// };

