// XXX: Add YOUR Application ID
var appId = "xxx";

function addTitles() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getRange("B:B");
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var results = [];
  results[0] = [];
  for(var i = 0; i <= numRows; i++) {
    sheet.getRange("A" + (i + 1)).setValue(getTitleFromIsbn(values[i]));
  }
 };

function getTitleFromIsbn(isbn) {
  if(isbn == "") {
    return ""
  }
  var response = UrlFetchApp.fetch(
    "https://app.rakuten.co.jp/services/api/BooksBook/Search/20130522?" +
    "applicationId=" + appId +
    "&isbn=" + isbn +
    "&outOfStockFlag=1"
  );
  var data = JSON.parse(response.getContentText());
  if(data.count >= 1) {
    result = data.Items[0].Item.title;
  } else {
    result = "";
  }
  return result;
}

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Add titles",
    functionName : "addTitles"
  }];
  sheet.addMenu("Scripts", entries);
}
