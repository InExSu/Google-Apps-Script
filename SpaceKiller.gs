// https://docs.google.com/spreadsheets/d/1gEQ2Yz3lC9DRY4y2z_Ni6G6LWek_xTNDmBJEnrt5z7Y/edit?pli=1#gid=0

function digitsSpacesKiller() {

  // в выделенных ячейках,содержащих только цифры, пробелы, системный разделитель десятичных чисел,
  // удалить пробел

  // нужно дополнительное тестирование!

  var rng = SpreadsheetApp.getActiveRange();

  if (rng === null) {
    // нет выделенного диапазона
  } else {

    var a2d = rng.getValues();

    a2d = arrayXdDigitsSpaceKiller(a2d, '0123456789 ,');

    // вставить массив на лист

    Array2d_2_Range(rng, a2d);
  }
}

function array2dDigitsSpaceKiller(a2d, tmp) {

  // в массиве, в элементах, содержащих только:
  // цифры, пробелы, системный разделитель десятичных чисел - 
  // удалить пробел 

  var ele = '';

  for (var row = 0; row < a2d.length; row++) {
    for (var col = 0; col < a2d[0].length; col++) {

      ele = String(a2d[row][col]);

      if (digitWithSpace(ele, tmp)) {

        a2d[row][col] = ele.replace(' ', '');

      }

    }
  }

  return a2d;

}


function digitWithSpace_Test() {

  Logger.log('');

}

function digitWithSpace(str, tmp) {

  // строка похожа на число с пробелом ?

  var smb = '';

  for (var pos = 0; pos < str.length; pos++) {

    smb = str[pos];

    if (symbolInString(smb, tmp)) {

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

