function postLineNotify(message){
  if (!message) return;

  var token = PropertiesService.getScriptProperties().getProperty('LINE_NOTIFY_TOKEN');
  var options = {
     "method"  : "post",
     "payload" : "message=" + message,
     "headers" : {"Authorization" : "Bearer "+ token}
   };
   UrlFetchApp.fetch("https://notify-api.line.me/api/notify",options);
}

function doEventEach(startTime, endTime, callBackEvent) {
  var scriptProp = PropertiesService.getScriptProperties();
  var cal = CalendarApp.getCalendarById(scriptProp.getProperty('CAL_ID'));
  var events = cal.getEvents(startTime, endTime);

  for (e in events) {
    callBackEvent(events[e]);
  }
}

function execYql(url, xpath) {
  if (!url) {
    return;
  }

  var yqlUrl = "https://query.yahooapis.com/v1/public/yql?muteHttpExceptions=true&format=json&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys&q="
  var yql = "select * from htmlstring where url = '" + url + "' and xpath = '" + xpath + "'";
  var queryURL = yqlUrl + encodeURIComponent(yql);
  var jsonString = UrlFetchApp.fetch(queryURL).getContentText();
  return JSON.parse(jsonString).query.results.result;
}

/**
 * objListの各要素に対しpropKeyArrayの要素順にvalueを取得して配列化する。
 * objListにpropertyが存在しない場合は commonProp から取得する。
 * 
 * @param {Array.<string>} propKeyArray 
 * @param {Array.<object>} objList 
 * @param {Object} commonProp
 * @return {Array.<Array.<string>>} 
 */
function toSpreadSheetValues(propKeyArray, objList, commonProp) {
  if (!objList || !objList.length) {
    return;
  }
  var ssValues = [];
  for (var i = 0; i < objList.length; i++) {
    var values = [];
    for (var p = 0; p < propKeyArray.length; p++) {
      if (objList[i].hasOwnProperty(propKeyArray[p])) {
        values.push(objList[i][propKeyArray[p]]);
        continue;
      }
      values.push(commonProp[propKeyArray[p]]);
    }
    ssValues.push(values);
  }
  return ssValues;
}

/**
 * ssId の存在するshNameシートに対し、resultsを上部固定行より下に行挿入する。
 * 数式はその下の行からコピーする。
 * 
 * @param {string} ssId spreadSheetId
 * @param {string} shName シート名
 * @param {Array.<Array.<object>>} results Range.setValues する値
 * @param {Object} opt resultPutStartColにより設定開始列を指定可能
 */
function insertResults(ssId, shName, results, opt) {
  var option = opt || {resultPutStartCol: 1};
  if (!results || !results.length) {
    return;
  }
  var ss = SpreadsheetApp.openById(ssId);
  var sh = ss.getSheetByName(shName);
  var frozenRowCnt = sh.getFrozenRows();
  sh.insertRowsBefore(frozenRowCnt + 1, results.length);
  sh.getRange(frozenRowCnt + 1 + results.length, 1, 1, sh.getLastColumn()).copyTo(
    sh.getRange(frozenRowCnt + 1, 1, results.length, sh.getLastColumn()));
  sh.getRange(frozenRowCnt + 1, option.resultPutStartCol, results.length, results[0].length).setValues(results);
  SpreadsheetApp.flush();
}

