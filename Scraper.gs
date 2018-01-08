function collectSchelpaCos() {
  var scriptProp = PropertiesService.getScriptProperties();
  var ssId = scriptProp.getProperty('SS_ID');
  var propBaseDate = scriptProp.getProperty('BASE_DATE');
  var baseDate = propBaseDate ? new Date(propBaseDate) : new Date();
  var periodDate = scriptProp.getProperty('PERIOD_DATE') || 7;

  // 日曜の夜実行の週トリガーで完了済みイベントの出欠を収集する
  doEventEach(new Date(baseDate.getTime() - (24 * 60 * 60 * 1000 * periodDate)), baseDate, function(event) {
    var eventProp = {
      "url": event.getDescription().match(/.*schelpa.*/),
      "title": event.getTitle(),
      "date": event.getStartTime().toLocaleDateString()
    };

    if (!eventProp["url"]) {
      return;
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
        parseCoToObjList(eventProp["url"]),
        eventProp),
      {resultPutStartCol: 4});
  });
}

function parseCoToObjList(url) {
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
      return coList;
    }
  };
  var jsonResult = execYql(url, xpath.path);
  return xpath.parse(jsonResult);
}
