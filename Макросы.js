function FormatSet2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A3').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('00000000000');
};

function TextTutti() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('@');
};

function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('S2').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('S2:S8').activate();
  spreadsheet.getRange('S2:S8').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(true)
    .requireValueInRange(spreadsheet.getRange('\'Тары\'!$A$2:$A$99'), true)
    .build());
};

function URL_Fetch_Code_Response_Test() {
  const url = "https://cdn-ru.bitrix24.ru/b6361393/iblock/435/4357badb4a38a6253c39a80ec0626d31/SERTIFIKAT_MZS_VK_EKRAN.pdf";

  // var resp = URL_Fetch_Code_Response('url');
  // / The code below logs the HTTP headers from the response
  // // received when fetching the Google home page.
  var response = UrlFetchApp.fetch(url);
  Logger.log(response.getAllHeaders());
}

function URL_Fetch_Code_Response(url) {
  // UDF вернёт ответ сервера
  var url_trimmed = url.trim();
  // Check if script cache has a cached status code for the given url
  var cache = CacheService.getScriptCache();
  var result = cache.get(url_trimmed);

  // If value is not in cache/or cache is expired fetch a new request to the url
  if (!result) {

    var options = {
      'muteHttpExceptions': true,
      'followRedirects': false
    };
    var response = UrlFetchApp.fetch(url_trimmed, options);
    var responseCode = response.getResponseCode();

    // Store the response code for the url in script cache for subsequent retrievals
    cache.put(url_trimmed, responseCode, 21600); // cache maximum storage duration is 6 hours
    result = responseCode;
  }

  return result;
}

function debuggerTest(){
  debugger
  let s = 1;
}