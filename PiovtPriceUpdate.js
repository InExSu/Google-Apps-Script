function aMain() {
  // по диапазонам цен и артикулов
  // обновить столбец цен

  // let sheetPrice = SpreadsheetApp.getActive().getSheetByName('Прайс без НДС');
  // let sheetPivot = SpreadsheetApp.getActive().getSheetByName('сводная таблица');

  // function диапазонАрти
  // function диапазонЦена  
  // function столбНаЛист_
  // словарьАртиЦена = диапазоныВСловарь(диапазонАрти, ДиапазонЦена)
  // массивСтолбЦена = столбОбновить(массивСтолбЦена, массивАрти, словарьАртиЦена)
  // столбНаЛист(массивСтолбЦена, лист, ячейкаСтарт, артиУникСтоп)

  // logPriceChanged(a2ColumnPrice(sheetPrice), 
  // columnOnSheet(sheetPivot,
  // columnPriceUpdate(
  //   a2ColumnPrice(sheetPrice),
  //   a2ColumnArtic(sheetPrice),
  //   dictRangeArticPrice(sheetPrice),
  //   articUniq(sheetPrice))));
}


function logPriceChanged(a2Old, a2New) {
  // положить в лог сравнение столбцов прайса

  let a2 = a2ColumnsClue(a2Old, 0, a2New, 0);

}

function a2ColumnsClue_Test() {
  let a2Old = [[11], [12]];
  let a2New = [[21], [22]];
  let coOld = 0;
  let coNew = 0;

  let a2 = a2ColumnsClue(a2Old, coOld, a2New, coNew)
}

function a2ColumnsClue(a2Old, colOld, a2New, colNew) {
  // склеить столбцы массивов 2мерных

  let a2n = [];
  let dif = 0;
  let old = 0;
  let neu = 0;

  for (row = 0; row < a2Old; row++) {

    old = a2Old[row][colOld];
    neu - a2New[row][colNew];
    dif = old - neu;
    a2n.push([old, neu, dif]);

  }

  return a2n;
}

function columnOnSheet(sheet, a2) {
  // положить цены на лист и вернуть входящий a2
  // sheet.getRange('J:J').setValues(a2);
  return a2;
}

function columnPriceUpdate_Test() {

  let sheet = SpreadsheetApp.getActive().getSheetByName('сводная таблица');
  let a2Artic = a2ColumnArtic(sheet);
  let a2Price = a2ColumnPrice(sheet);
  let sheetPrice = SpreadsheetApp.getActive().getSheetByName('Прайс без НДС');
  let dict = dictRangeArticPrice(sheetPrice);
  let a2 = columnPriceUpdate(a2Price, a2Artic, dict)
}

function columnPriceUpdate(a2Price, a2Artic, dict) {
  // проходом по массив артикулов, если он есть в словаре, изменить цену в этой же строке

  let key = '';
  for (let row = 0; row < a2Artic.length; row++) {
    key = a2Artic[row][0]
    if (dict.has(key)) {
      val = dict.get(key);
      a2Price[row][0] = val;
    }
  }
  return a2Price;
}

function a2ColumnArtic_Test() {
  let a2 = a2ColumnArtic(SpreadsheetApp.getActive().getSheetByName('сводная таблица'))
}
function a2ColumnArtic(sheet) {
  // вернуть массив 2мерный артикулов
  return sheet.getRange('B:B').getValues();
}

function a2ColumnPrice(sheet) {
  // вернуть массив 2мерный артикулов
  return sheet.getRange('J:J').getValues();
}

function dictRangeArticPrice_Test() {
  let sheetPrice = SpreadsheetApp.getActive().getSheetByName('Прайс без НДС');
  let dict = dictRangeArticPrice(sheetPrice);
}

function dictRangeArticPrice(sheet) {
  // сделать массив ассоциированный Артикул-Цена из двух диапазонов 

  let a2Price = sheet.getRange('C:H').getValues();
  let a2Artic = sheet.getRange('L:R').getValues();
  let dict = new Map();
  let key = '';

  for (let row = 0; row < a2Artic.length; row++) {
    for (let col = 0; col < a2Artic[0].length; col++) {

      key = a2Artic[row][col];

      if (/\d{3}-\d{3}-\d{4}/.test(key)) {
        val = a2Price[row][col];
        dict.set(key, val);
      }
    }
  }
  return dict;
}

function articUniq_Test() {
  let sheetPrice = SpreadsheetApp.getActive().getSheetByName('Прайс без НДС');
  let a2 = sheetPrice.getRange('L:R').getValues();
  let a1 = articUniq(a2);
  let a12str = a1.reduce((a12str, value) => {
    return a12str + value + '\n';
  })
  Logger.log(a12str);
}

function articUniqCheck(sheetPrice) {
  // проверить уникальность артикулов в диапазоне
  let a2 = sheetPrice.getRange('L:R').getValues();
  // // вытянуть 2мерный массив в одномерный
  // let a1 = a2.flat(Infinity);
  // // оставить только артикулы
  // let a1Arti = a1.filter(item => /\d{3}-\d{3}-\d{4}/.test(item));
  // let a1Dupl = a1Duplicates2a1(a1Arti);

  // передача регулярного через аргумент не прокатит
  // let reg = /\d{3}-\d{3}-\d{4}/
  // let a1Dupl = articUniq(a2, reg);
  let a1Dupl = articUniq(a2)

  return a1Dupl.length == 0 ? true : false;
}

function articUniq(a2) {
  // вернуть масссив 1мерный одинаковых артикулов

  // вытянуть 2мерный массив в одномерный
  let a1 = a2.flat(Infinity);
  // оставить только артикулы
  let a1Arti = a1.filter(item => /\d{3}-\d{3}-\d{4}/.test(item));
  return a1Duplicates2a1(a1Arti);
}

function regParam_Test() {
  let re = /\d/;
  let a1 = [1, "z"];
  let fi = a1.filter(item => re.test(item));
  Logger.log(fi);
}








