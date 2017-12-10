function collectSchelpaCos() {
  var scriptProp = PropertiesService.getScriptProperties();
  var ssId = scriptProp.getProperty('SS_ID');
  var propBaseDate = scriptProp.getProperty('BASE_DATE');
  var baseDate = propBaseDate ? new Date(propBaseDate) : new Date();
  var periodDate = scriptProp.getProperty('PERIOD_DATE') || 7;

  // 日曜の夜実行の週トリガーで完了済みイベントの出欠を収集する
  var cal = CalendarApp.getCalendarById(scriptProp.getProperty('CAL_ID'));
  var events = cal.getEvents(new Date(baseDate.getTime() - (24 * 60 * 60 * 1000 * periodDate)), baseDate);

  for (e in events) {
    var eventProp = {
      "url": events[e].getDescription().match(/.*schelpa.*/),
      "title": events[e].getTitle(),
      "date": events[e].getStartTime().toLocaleDateString()
    };

    if (!eventProp["url"]) {
      continue;
    }
    
    var titleAtSplit = eventProp["title"].split(/[@＠]/);
    var eventPropKeys = ["date", "title", "url"];
    var dbPropKeys = ["name", "co", "title", "date", "url"];
    if (titleAtSplit.length == 2) {
      eventProp["kind"] = titleAtSplit[1].replace("千代田区スポーツセンター", "体育館");
      eventPropKeys.push("kind");
      dbPropKeys.push("kind");
    }

    insertResults(ssId, 'event',
      toSpreadSheetValues(eventPropKeys, [eventProp]));

    insertResults(ssId, 'db',
      toSpreadSheetValues(
        dbPropKeys,
        execYql(eventProp["url"]),
        eventProp),
      {resultPutStartCol: 4});
  }
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

function execYql(url) {
  if (!url) {
    return;
  }

  var xpath = {
    path: '//*[@id="userLine"]//span[@class="fontName"] | //*[@id="userLine"]//div[@class="boxSanka"]//td',
    parse: function(yqlResults) {
      var coList = [];
      
      // <span class="fontName">NAME</span>\n<td class="length_0 table_5560 sankaList_CO"/>\n
      var spanElem = yqlResults.split('<span ');
      spanElem.forEach(function(elem) {
        if (!elem) return;
        var oneLine = elem.replace(/\n/g,'');
        coList.push({
          name: oneLine.replace(/^[^>]+>([^<]+)<.*/,'$1'),
          co: oneLine.replace(/.*sankaList_(ok|ng|)\b.*/,'$1')
        });
      });
//      yqlResults.span.forEach(function(span) {
//        coList.push({name: span.content});
//      });
//      yqlResults.td.forEach(function(td, idx) {
//        coList[idx].co = td["class"].replace(/.*sankaList_/, '');
//      });
      return coList;
    }
  };
  var yqlUrl = "https://query.yahooapis.com/v1/public/yql?muteHttpExceptions=true&format=json&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys&q="
  var yql = "select * from htmlstring where url = '" + url + "' and xpath = '" + xpath.path + "'";
  var queryURL = yqlUrl + encodeURIComponent(yql);
  var jsonString = UrlFetchApp.fetch(queryURL).getContentText();
  var json = JSON.parse(jsonString);
  return xpath.parse(json.query.results.result);
}

//function test() {
////  var htmlString = UrlFetchApp.fetch('http://www.densuke.biz/list?cd=xEK2p7BeYB33SdD4&pw=').getContentText();
////  Logger.log(htmlString);
//  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SS_ID'));
//  var shDb = ss.getSheetByName('db');
//  var frozenRowCnt = shDb.getFrozenRows();
////  shDb.insertRowsBefore(frozenRowCnt + 1, 0);
//  var vals = shDb.getRange(frozenRowCnt + 1 + 0, 1, 3, shDb.getLastColumn()).getValues();
//  for (var r in vals) {
//    for (var c in vals[r]) {
//      Logger.log(vals[r][c]);
//    }
//  }
//}
