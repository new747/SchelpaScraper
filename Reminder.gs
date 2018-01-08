function pushRemind() {
  var scriptProp = PropertiesService.getScriptProperties();
  var propBaseDate = scriptProp.getProperty('REMINDER_BASE_DATE');
  var baseDate = propBaseDate ? new Date(propBaseDate) : new Date();
  var periodDate = scriptProp.getProperty('REMINDER_PERIOD_DATE') || 14;

  // 日曜の夜実行の週トリガーで2週間先までのイベントのCO状況を通知する
  doEventEach(baseDate, new Date(baseDate.getTime() + (24 * 60 * 60 * 1000 * periodDate)), function(event) {
    var eventProp = {
      "url": event.getDescription().match(/.*schelpa.*/),
      "title": event.getTitle(),
      "date": event.getStartTime().toLocaleDateString()
    };
    if (!eventProp["url"]) return;
    
    var coSummary = parseCoSummary(eventProp["url"]);
    if (!coSummary) return;
    postLineNotify(eventProp["date"] + "の" + eventProp["title"] + "の参加人数は" + coSummary["okCount"] + "\n\n"
              + "表明者は" + coSummary["totalCount"] + "(" + coSummary["coNames"]+ ") です。\n\n"
              + "未COの方は早めのCOにご協力お願いします。\n\n" + eventProp["url"]);
  });
}

function parseCoSummary(url) {
  var xpath = {
    path: '//*[@id="stsList_id"]',
    parse: function(yqlResults) {
      var coSummary;

      // <span class="tips" title="coname1 coname2 separated by space">15人</span>
      var spanElem = yqlResults.split(/\n/g);
      spanElem.forEach(function(elem) {

        if (!elem || !elem.match(/<span class="tips" title="([^"]+)">([0-9]+人)/)) return;
        if (!coSummary) {
          coSummary = {
            coNames: RegExp.$1,
            totalCount: RegExp.$2
          };
        } else {
          coSummary["okCount"] = RegExp.$2;
        }
      });
      return coSummary;
    }
  };
  var jsonResult = execYql(url, xpath.path);
  return xpath.parse(jsonResult);
}
