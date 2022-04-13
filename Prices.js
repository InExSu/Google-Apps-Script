function rangePriceColumnUpade() {
  // Переделка прайса
  // из формул с кодом разово сделать копию с артикулами.
  // На листе будут три таблицы:
  // слева    версия для печати - цены вбиваются пользователями.
  // в центре версия с артикулами.
  // По этим  версиям, скрипт обновления цен в листе "сводная таблица"
  // создаст справа отчёт работы в третью таблицу.

  const spread = SpreadsheetApp.getActive();

  const sheet_Price_bez_NDS = spread.getSheetByName('Прайс без НДС');

  let sheet_Svodnaya = spread.getSheetByName('сводная таблица');

  const sheet_Logg = spread.getSheetByName('Log');

  if (headersOk(sheet_Price_bez_NDS, sheet_Svodnaya) === false) {

    let warning = 'Ожидаемые заголовки НЕ совпали. \n Выход!'
    Browser.msgBox(warning);
    Logger.log(warning);

  } else {

    const a2_Price_bez_NDS_Prices_LQ = sheet_Price_bez_NDS.getRange('L:Q').getValues();
    const a2_Price_bez_NDS_Artics_CH = sheet_Price_bez_NDS.getRange('C:H').getValues();

    const a2_Svodnaya_Artics_B = sheet_Svodnaya.getRange('B:B').getValues();
    const map_Artics_Svodnaya_B = Array2D_2_Map(a2_Svodnaya_Artics_B);

    const range_Svodnaya_J = sheet_Svodnaya.getRange('J:J');
    let a2_Column_Prices_J = range_Svodnaya_J.getValues();

    // копировать массив 2мерный 
    let a2_Column_Prices_Old = JSON.parse(JSON.stringify(a2_Column_Prices_J))

    a2PriceColumnUpdate(a2_Price_bez_NDS_Prices_LQ, a2_Price_bez_NDS_Artics_CH, map_Artics_Svodnaya_B, a2_Column_Prices_J);

    // В "Прайс без НДС" для разных ростов указан один артикул.
    // нужно по этому артикулу установить туже цену для других ростов
    const a2_Svodnya_BD = sheet_Svodnaya.getRange('B:D').getValues();
    priceGrowths(a2_Price_bez_NDS_Prices_LQ, a2_Svodnya_BD, a2_Column_Prices_J);

    // проставить цены по массиву артикулов соответствия 2022-04-13
    artiCoolsPriceOne(a2_Svodnaya_Artics_B, a2_Column_Prices_J, a2Artics4One())

    range_Svodnaya_J.setValues(a2_Column_Prices_J);

    rangePriceColumnUpade_Log(sheet_Logg, a2_Svodnaya_Artics_B, a2_Column_Prices_Old, a2_Column_Prices_J);

    sheet_Logg.activate();

  }
}

function priceGrowths_Test() {
  let a2_Price_bez_NDS_Prices_LQ = [['0', '1', '2', '3', '4', 'z', 'артик1']];
  let a2_Svodnya_BD = [
    ['0', '1', '2'],
    ['артик1', '', 'артик1 назв рост 1'],
    ['артик2', '', 'артик1 назв рост 2'],
    ['артик3', '', 'артик3 без р о с т а']];
  let a2_Column_Prices_J = [
    [0],
    [9],
    [9],
    [3]];

  priceGrowths(a2_Price_bez_NDS_Prices_LQ, a2_Svodnya_BD, a2_Column_Prices_J);

  console_log(a2_Column_Prices_J);

}

function priceGrowths(a2_Price_bez_NDS_Prices_LQ, a2_Svodnya_BD, a2_Column_Prices_J) {
  // Проходом по диапазону артикулов листа a2_Price_bez_NDS_Prices_LQ ,
  // найти артикул в массиве a2_Svodnya_BD,
  // взять номер строки a2_Svodnya_BD,
  // взять наименование
  // в наименовании отсечь по /\d\sрост или по "рост".
  // по номеру строки a2_Svodnya_BD взять новую цену из a2_Column_Prices_J.
  // Проходом по столбцу название,
  // если наименование начинаеся со значения без роста и в нём есть слово "рост",
  // то в эту же строку столбца цена проставить цену

  let artic = '';
  let name_ = '';
  let price = 0;
  const map_Artics_Row = Array2D_Column_2_Map(a2_Svodnya_BD, 0)

  let nameCut = '';
  const COL_2 = 2;

  for (let row = 0; row < a2_Price_bez_NDS_Prices_LQ.length; row++) {
    for (let col = 0; col < a2_Price_bez_NDS_Prices_LQ[0].length; col++) {

      artic = a2_Price_bez_NDS_Prices_LQ[row][col];

      if (map_Artics_Row.has(artic)) {

        let row_Found = map_Artics_Row.get(artic);

        if (row_Found === 'undefined') {
          debugger;
        }

        name_ = a2_Svodnya_BD[row_Found][2];

        if (name_.search(/рост/) > -1) {

          price = a2_Column_Prices_J[row_Found][0];

          nameCut = nameGrowths(name_);

          if (nameCut !== '') {

            // проход по "Полное наименование" - проставить price
            A2s_Match(nameCut, a2_Svodnya_BD, COL_2, a2_Column_Prices_J, price);

          } else {
            debugger;
          }
        }
      }
    }
  }
}

function A2s_Match_Test() {

  let a2_Svodnya_BD = [
    ['0', '1', '2'],
    ['артик1', '', 'артик1 рост 1'],
    ['артик2', '', 'артик1 рост 2'],
    ['артик3', '', 'без р о с т а']];
  let a2_Column_Prices_J = [
    [0],
    [1],
    [2],
    [3]];

  let nameCut = 'артик1';
  const COL_2 = 2;
  let price = 234.56;

  A2s_Match(nameCut, a2_Svodnya_BD, COL_2, a2_Column_Prices_J, price);

  console_log(a2_Column_Prices_J);
}

function console_log(a2_Column_Prices_J) {
  console.log('a2_Column_Prices_J[0] = ' + a2_Column_Prices_J[0]);
  console.log('a2_Column_Prices_J[1] = ' + a2_Column_Prices_J[1]);
  console.log('a2_Column_Prices_J[2] = ' + a2_Column_Prices_J[2]);
  console.log('a2_Column_Prices_J[3] = ' + a2_Column_Prices_J[3]);

}

function A2s_Match(nameCut, a2_Svodnya_BD, COL_2, a2_Column_Prices_J, price) {
  // если значение начинается с nameCut, подставить цену в a2_Column_Prices_J
  // a2_Svodnya_BD и a2_Column_Prices_J одинаковы по высоте

  let str = '';
  let lft = '';

  for (let row = 0; row < a2_Svodnya_BD.length; row++) {

    str = a2_Svodnya_BD[row][COL_2];
    lft = str.slice(0, nameCut.length);

    if (lft.toUpperCase() === nameCut.toUpperCase()) {

      a2_Column_Prices_J[row][0] = price;

    }
  }
}

function nameGrowths_Test() {
  let name = '';
  let noGr = '';

  // name = 'Издел (0 РосТ)';
  // noGr = nameGrowths(name);
  // console.log(noGr);

  name = 'Издел РосТ 2';
  noGr = nameGrowths(name);
  console.log(noGr);

  // name = 'Издел без р о с т а';
  // noGr = nameGrowths(name);
  // console.log(noGr);

}

function nameGrowths(stringIn) {
  // вернуть слева от роста
  // для случаев:
  // (3 рост)
  // РОСТ 1

  let a1 = [];
  let noGrowth = '';

  a1 = stringIn.split(/\d\sрост/i);

  if (a1.length > 1) {

    noGrowth = a1[0];

  } else {

    a1 = stringIn.split(/\sрост\s/i)

    if (a1.length > 1)
      noGrowth = a1[0];

  }

  return noGrowth;
}

function a2PriceColumnUpdate(a2_Arti_Range, a2_Price_Range, map_Arti, a2_Price_Colum) {
  // Словарь артикулов - артикул: номер строки

  // Проходом по массиву артикулов
  // 	Если артикул есть в словаре
  // 		Взять цену из массива цен в координатах артикула
  // 			Взять номер строки из словаря
  // 				Вставить в массив цен цену по номеру строки

  let artic = '';

  for (let row = 0; row < a2_Arti_Range.length; row++) {
    for (let col = 0; col < a2_Arti_Range[0].length; col++) {

      artic = a2_Arti_Range[row][col];

      if (artic.search(/\d{3}-\d{3}-\d{4}/) > -1) {

        if (map_Arti.has(artic)) {

          let row_Price = map_Arti.get(artic);
          let price = a2_Price_Range[row][col];

          let stop = 'да';
          a2_Price_Colum[row_Price][0] = convert2FloatCommaPointIfPossible(price);
        }
      }
    }
  }
}

function rangePriceColumnUpade_Log_Test() {

  const spread = SpreadsheetApp.getActive();
  const sheet_Logg = spread.getSheetByName('Log');

  const a2_Column_Artics = [['a1'], ['a2']];
  const a2_Column_Prices_Old = [[1], [2]];
  const a2_Column_Prices_New = [[11], [22]];

  rangePriceColumnUpade_Log(sheet_Logg, a2_Column_Artics, a2_Column_Prices_Old, a2_Column_Prices_New)
}

function rangePriceColumnUpade_Log(sheet_Logg, a2_Column_Artics, a2_Column_Prices_Old, a2_Column_Prices_J) {

  sheet_Logg.clear();

  let cell = sheet_Logg.getRange('A1')
  let valu = 'Лог обновления "сводная таблица" столбец "Прайс без НДС" ' + Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss' мск'");
  cell.setValue(valu);

  cell = sheet_Logg.getRange('A4');
  array2d2Range(cell, a2_Column_Artics);

  a2_Column_Prices_Old[0][0] = 'Было';
  cell = sheet_Logg.getRange('B4');
  array2d2Range(cell, a2_Column_Prices_Old);

  a2_Column_Prices_J[0][0] = 'Стало';
  cell = sheet_Logg.getRange('C4');
  array2d2Range(cell, a2_Column_Prices_J);

  // копировать массив 2мерный
  const a2_Diff = JSON.parse(JSON.stringify(a2_Column_Prices_J))
  // заменить в столбце все значения на формулу
  arrayColumFillFormula(a2_Diff, 1, 0, 4);
  a2_Diff[0][0] = 'Сравнение';

  cell = sheet_Logg.getRange('D4');
  array2d2Range(cell, a2_Diff);
}

function arrayColumFillFormula(a2, rowStart, col, shift) {
  // заменить в столбце все значения на формулу

  for (let row = rowStart; row < a2.length; row++) {
    let rowFormu = row + shift;
    let formula_ = '=B' + rowFormu + '=C' + rowFormu;
    a2[row][col] = formula_;
  }
}

function headersOk(sheet_Sour, sheet_Dest) {
  // проверка значений ячеек

  return cell_Value(sheet_Dest.getRange('B1'), 'Артикул') &&
    cell_Value(sheet_Dest.getRange('J1'), 'Цена, руб (Без НДС)') &&
    cell_Value(sheet_Sour.getRange('C7'), 'ШМП-1') &&
    cell_Value(sheet_Sour.getRange('L7'), 'ШМП-1')

}

function cell_Value_Test() {

  const sheet = SpreadsheetApp.getActive().getSheetByName('Прайс без НДС');
  const cell_ = sheet.getRange('C7')

  let value = 'ШМП-1';
  let resul = cell_Value(cell_, value);
  Logger.log(resul);

  value = '';
  resul = cell_Value(cell_, value);
  Logger.log(resul);
}

function cell_Value(cell, value) {
  if (cell.getValue() !== value) {
    const sheet = sheetByRange(cell);
    Logger.log('На листе ' + sheet.getName() + ' в ячейке ' +
      cell.getA1Notation() + ' !== ' + value);
    return false;
  }
  return true;
}


function A2PriceColumnUpdate_Test() {

  let a2_Arti_Range = [
    ['1', ''],
    ['3', '4']];
  let a2_Price_Range = [
    [11, 22],
    [33, 44]];

  let a2_Arti__Colum = ['5', '4', '3', '2', '1'];
  let a2_Price_Colum = [5, 4, 3, 2, 1];

  let map_Arti = Array2D_Column_2_Map(a2_Arti__Colum, 0);

  A2PriceColumnUpdate(a2_Arti_Range, a2_Price_Range, map_Arti, a2_Price_Colum);

  if (a2_Price_Colum[4][0] !== 11) {
    Logger.log('a2_Price_Colum[4][0] !== 11');
  }
}




function sheetByRange(cell) {
  // вернуть лист по диапазону
  // как в Excel range.Parent

  return sheetById(cell.getGridId());
}


function sheetById(id) {
  // вернуть лист по id

  return SpreadsheetApp.getActive().getSheets().filter(
    function (s) {
      return s.getSheetId() === id;
    }
  )[0];
}



function price2VendorCode_Test() {

  cell = SpreadsheetApp.getActive().getSheetByName('Прайс без НДС (копия) с формулами').getRange(8, 4).getFormula();

  price2VendorCode(cell);
}

function price2VendorCode(formu) {
  // Разовый этап создания из таблицы с формулами таблицы с артикулами
  // UDF из тексту формулы извлекает код 1С,
  // по коду 1С ищет на листе строку с этим кодом,
  // если в строке есть артикул вернёт его или cell
  // формула на листе, использующая эту формулу
  // =ЕСЛИ(ЕОШИБКА(FORMULATEXT(C4));C4;
  //    ПОДСТАВИТЬ(
  //      price2VendorCode(FORMULATEXT(C4));
  // СИМВОЛ(34);""))

  // нужны формулы без пробелов
  formu.replaceAll(' ', '');
  let code1 = extractBetween(formu, 'MATCH(', ";'");
  if (code1 == '') {
    code1 = formu;
  }
  else {
    //Logger.log('code1 = ' + code1);
  }
  return code1;
}

function extractBetween_Test() {

  let result = extractBetween('123', '0', '3');
  if (result !== '') {
    Logger.log('extractBetween_Test ошибка: ждал пусто, пришло ' + result);
  }

  result = extractBetween('123', '1', '3');
  if (result !== '2') {
    Logger.log('extractBetween_Test ошибка: ждал 2, пришло ' + result);
  }
  result = extractBetween('12345', '12', '45');
  if (result !== '3') {
    Logger.log('extractBetween_Test ошибка: ждал 3, пришло ' + result);
  }
}

function extractBetween(sMain, sLeft, sRigh) {
  // из строки извлечь строку между подстроками
  // InExSu 

  // добавил 1, чтобы стало возможным условие проверки на 0
  let idxBeg = sMain.indexOf(sLeft) + 1;
  let idxEnd = sMain.indexOf(sRigh) + 1;
  let strOut = '';

  if ((idxBeg * idxEnd) > 0) {
    idxBeg = idxBeg + sLeft.length - 1;
    idxEnd = idxEnd - 1;
    strOut = sMain.slice(idxBeg, idxEnd);
  }
  return strOut;
}

function Array2D_Column_2_Map(array2d, column_key) {
  // из массива 2мерного вернуть словарь - массив ассоциативный: 
  // значение столбца и номер строки

  let map_return = new Map();
  let val = '';

  for (let row = 0; row < array2d.length; row++) {

    val = String(array2d[row][column_key]);

    if (val.length > 0) {
      // если ключ повторяется, то обновится значение
      map_return.set(val, row);
    }
  }
  return map_return;
}

function Array2D_2_Map(array2d) {
  // из массива 2мерного вернуть словарь массив ассоциативный

  let map_return = new Map();
  let val = '';

  for (let row = 0; row < array2d.length; row++) {
    for (let col = 0; col < array2d[0].length; col++) {

      val = String(array2d[row][col]);

      if (val.length > 0) {
        // если ключ повторяется, то обновится значение
        map_return.set(val, row);
      }
    }
  }
  return map_return;
}

function array2d2Range(cell, a2d) {

  // массив 2мерный вставить на лист

  let sheet_id = cell.getGridId();
  let sheet_ob = sheetById(sheet_id);
  const row_numb = cell.getRow();
  const col_numb = cell.getColumn();

  sheet_ob.getRange(row_numb, col_numb, a2d.length, a2d[0].length).setValues(a2d);
}

function artiCoolsCheck() {
  // проверить артикулы листа "Прайс без НДС" в листе "сводная таблица"

  const sheetPivot = SpreadsheetApp.getActive().getSheetByName('сводная таблица');
  const sheetPrice = SpreadsheetApp.getActive().getSheetByName('Прайс без НДС');
  const rangeArticPivot = sheetPivot.getRange("B:B");
  const rangeArticPrice = sheetPrice.getRange("L:Q");
  const a2Pivot = rangeArticPivot.getValues();
  const a2Price = rangeArticPrice.getValues();

  let artics = '';
  let value = ''
  const a2Map = Array2D_Column_2_Map(a2Pivot, 0);

  for (let row = 0; row < a2Price.length; row++) {
    for (let col = 0; col < a2Price[0].length; col++) {

      value = a2Price[row][col];

      if (/\d{3}-\d{3}-\d{4}/.test(value)) {

        if (a2Map.has(value) == false) {

          artics += value + "; ";

        }
      }
    }
  }
  if (artics.length === 0) {
    Browser.msgBox("Артикулы все найдены");
    console.log("Артикулы все найдены")
  } else {
    Browser.msgBox("Отсутствуют в 'сводная таблица:\n" + artics);
    console.log("Отсутствуют в 'сводная таблица:\n" + artics);
  }
}

function artiCoolsPriceOne_Test() {

  const a2ColumnArtics = [['102-011-0017'], ['102-011-0056']];
  let a2ColumnPrices = [[11], [22]];
  let a2Artics = a2Artics4One();

  artiCoolsPriceOne(a2ColumnArtics, a2ColumnPrices, a2Artics);

  if (a2ColumnPrices[1][0] == 11) {
    console.log('artiCoolsPriceOne_Test OK!');
  } else {
    console.log('artiCoolsPriceOne_Test ОШИБКА! ожидалось 11, получил ' + a2ColumnPrices[1][0]);
  }
}

function artiCoolsPriceOne(a2ColumnArtics, a2ColumnPrices, a2Artics) {
  // проходом по столбцу артикулов "сводная таблица"
  // взять цену из строки массива цен
  // искать артикул в массиве массивов одинаковых артикулов
  // проходом по вложенному массиву по столбцу артикулов
  // проставить цену в столбец цен

  let artic = '';
  let a1Art = [];
  let price = 0;
  let mapArtics = Array2D_2_Map(a2ColumnArtics)

  for (let rowA = 0; rowA < a2ColumnArtics.length; rowA++) {

    artic = a2ColumnArtics[rowA][0];
    a1Art = a2FindA1(a2Artics, artic);

    if (typeof a1Art === 'object') {

      price = a2ColumnPrices[rowA][0];

      price2Artics(mapArtics, a2ColumnPrices, a1Art, price);

    }
  }
}

function price2Artics_Test() {
  const a2Artics = [['1-1'], ['2-2']];
  const mapArtics = Array2D_2_Map(a2Artics);
  let a2ColumnPrices = [[11], [22]];;
  const a1Art = ['3-3', '2-2'];
  const price = 1;
  price2Artics(mapArtics, a2ColumnPrices, a1Art, price)
  if (a2ColumnPrices[1][0] !== price) {
    console.log('price2Artics_Test, ошибка ожидалось 1, получил' + a2ColumnPrices[1][0]);
  } else {
    console.log('price2Artics_Test Ok!');
  }
}
function price2Artics(mapArtics, a2ColumnPrices, a1Art, price) {
  // расставить артикулам цены

  artic = '';
  row = -1;

  for (let index = 0; index < a1Art.length; index++) {

    artic = a1Art[index];

    if (mapArtics.has(artic)) {

      row = mapArtics.get(artic);

      a2ColumnPrices[row][0] = price;
    }
  }
}


function a2Artics4One() {
  // вернуть массив артикулов одинаковых цен
  return [
    ['102-011-0017', '102-011-0056'],
    ['102-011-0012', '102-011-0079'],
    ['101-011-0003', '101-011-0004'],
    ['101-011-0005', '101-011-0006'],
    ['101-011-0007', '101-011-0008'],
    ['102-044-0001', '102-044-0002'],
    ['102-044-0003', '102-044-0004'],
    ['102-044-0005', '102-044-0006'],
    ['102-044-0007', '102-044-0008'],
    ['102-044-0009', '102-044-0010'],
    ['102-044-0011', '102-044-0012'],
    ['102-044-0013', '102-044-0014'],
    ['102-044-0015', '102-044-0016'],
    ['102-044-0017', '102-044-0018'],
    ['102-024-0003', '102-024-0004', '102-024-0005'],
    ['102-025-0001', '102-025-0002', '102-025-0003'],
    ['102-025-0004', '102-025-0005', '102-025-0006'],
    ['102-025-0007', '102-025-0008', '102-025-0009'],
    ['302-122-0001', '302-122-0002', '302-122-0003', '302-122-0004', '302-122-0005'],
    ['302-122-0006', '302-122-0007', '302-122-0008', '302-122-0009'],
    ['302-122-0010', '302-122-0011', '302-122-0012', '302-122-0013'],
    ['302-123-0003', '302-123-0004', '302-123-0005', '302-123-0006', '302-123-0007'],
    ['302-123-0008', '302-123-0009']
  ];
}

function a2FindA1_Test() {
  let a2 = a2Artics4One();

  let a1 = a2FindA1(a2, '102-011-0017');
  console.log(typeof a1);

  a1 = a2FindA1(a2, '302-123-0007');
  console.log(typeof a1);

  a1 = a2FindA1(a2, '');
  console.log(typeof a1);
}

function a2FindA1(a2, value) {
  //  найти в массиве массивов и вернуть массив 1мерный

  if (value !== '') {

    let a1 = [];

    for (let row = 0; row < a2.length; row++) {

      a1 = a2[row];

      for (let idx = 0; idx < a1.length; idx++) {

        if (value == a1[idx]) {

          return a2[row];
        }
      }
    }
  }
}

// price2Artics_Test();
// artiCoolsPriceOne_Test();
