var targetUrl = "http://cal.syoboi.jp/db.php";
var cmdCh = "?Command=ChLookup";
var cmdPr = "?Command=ProgLookup&StTime=";
var cmdTi = "?Command=TitleLookup&TID=";

var oneDayTime = 24 * 3600 * 1000;
var oneMinTime = 60 * 1000;

function test() {
  var today = new Date;
  Logger.log(loadProgData([19], today, new Date(today.getTime() + oneDayTime)));
}

//この関数を每日動作させてください
function main() {
  var today = new Date();
  today = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  //設定の読込
  var spsheet  = SpreadsheetApp.getActive();
  var sheet    = spsheet.getSheets()[0];
  var settings = sheet.getRange(2, 2, 5, 1).getValues();
  
  //カレンダーの取得
  var cal = CalendarApp.getCalendarsByName(settings[0][0]);
  if (cal.length == 0) { return false; }
  else { cal = cal[0]; }

  //通知日の設定
  var remindDay = settings[1][0] || "每日";
  var dayList = sheet.getRange("b3").getDataValidation().getCriteriaValues()[0];
  remindDay = dayList.indexOf(remindDay);
  var tomorrow = new Date(today.getTime() + oneDayTime);
  remindDay = (remindDay == 7 ? 0 : (7 + remindDay - tomorrow.getDay())%7) * oneDayTime;
  remindDay = tomorrow.getTime() + remindDay;
  var remindTime = settings[2][0];
  remindTime = (remindTime.getHours()*60 + remindTime.getMinutes()) * oneMinTime
               + remindDay;

  //重複回避機能
  var invalidDup = (settings[3][0] == "OFF") ? false : true;

  //最終更新日
  var beginDay  = new Date((settings[4][0] || today).getTime() + oneMinTime);
  var endDay = new Date(today.getTime() + 8 * oneDayTime - oneMinTime);
  if (beginDay.getTime() >= endDay.getTime()) { return false; }
  
  //チャンネル設定の読込
  var cData  = spsheet.getSheets()[1].getDataRange().getValues();
  var cIdList= [];
  var cNameList  = {};
  cData.shift();
  cData.forEach(function(ch){
    if(ch[2] == "ON") {
      cIdList.push(ch[0]);
      cNameList[ch[0]] = ch[1];
    }
  });
  
  //番組データの読込
  var programs = loadProgData(cIdList, beginDay, endDay);
  
  //カレンダーに番組情報を入力
  programs.forEach(function(pr) {
    var subRemindTime = remindTime;
    while(subRemindTime > pr["StTime"]) { subRemindTime -= oneDayTime; }
    cal.createEvent( pr["Title"], pr["StTime"], pr["EdTime"],
                    { description: pr["Comment"], location: cNameList[pr["ChID"]] }
    ).addEmailReminder((pr["StTime"].getTime() - subRemindTime) / oneMinTime);
  });
  
  //チェックした日付を記録
  var range = sheet.getRange("b6");
  range.setValue(endDay);
  SpreadsheetApp.flush();
}


function loadProgData(chList, beginTime, endTime) {
  //指定された時間の番組表データを持ってくる
  var url = targetUrl + cmdPr + dateToString(beginTime) + '-' + dateToString(endTime)
          + '&ChID=' + chList.join(',');
  Logger.log(url);
  var xml = UrlFetchApp.fetch(url).getContentText();
  xml = XmlService.parse(xml).getRootElement();
  if (xml.getChild("Result").getChildText("Code") != "200") { return []; }
  var programs = xml.getChild("ProgItems").getChildren("ProgItem");
  
  //番組表データから一話目のみを抽出し尚且つTIDの配列を用意する
  var tidList = [];
  var extracts = ["TID", "ChID", "StTime", "EdTime"];
  programs = programs.filter(function(pr) {
    return pr.getChildText("Count") == "1";
  }).map(function(pr) {
    var result = {}
    extracts.forEach(function (key) {
      result[key] = pr.getChild(key).getText();
    });
    var tid = result["TID"];
    if (tidList.indexOf(tid) == -1) { tidList.push(tid); }
    return result;
  });
  
  if (programs.length > 0) {
    //タイトルデータを取ってくる
    var titles = {};
    xml = UrlFetchApp.fetch(targetUrl + cmdTi + tidList.join(',')).getContentText();
    xml = XmlService.parse(xml).getRootElement();
    if (xml.getChild("Result").getChildText("Code") != "200") { return []; }
    xml.getChild("TitleItems").getChildren("TitleItem").forEach(function(ti) {
      var url = ti.getChildText("Comment").match(/ (http:[^\]]+)\]/);
      titles[ti.getChildText("TID")] = {
        "string":  ti.getChildText("Title"),
        "comment": (url ? url[1] : "")
      };
    });
    
    //タイトルデータを当てはめる
    for(var i = 0; i < programs.length; i++) {
      var pr = programs[i];
      var tData = titles[pr["TID"]];
      pr["Title"]   = tData["string"];
      pr["Comment"] = tData["comment"];
      pr["StTime"] = new Date(pr["StTime"].replace(/-/g, '/'));
      pr["EdTime"] = new Date(pr["EdTime"].replace(/-/g, '/'));
    }
  }
  return programs;
}

function to2dig(num) { return ((num < 10) ? "0" : "") + num; }

function dateToString(date) {
  return "" + date.getFullYear()
         + [date.getMonth() + 1, date.getDate()].map(to2dig).join('')
         + "_"
         + [date.getHours(),date.getMinutes(), 0].map(to2dig).join('');
}

function reloadChannnel() {
  //元データの読込
  var sheet = SpreadsheetApp.getActive().getSheets()[1];
  var cData  = sheet.getDataRange().getValues();
  var cList = {};
  cData.shift();
  cData.forEach(function(ch){cList[ch[0]] = [ch[1], ch[2]];});
  var rule = sheet.getRange("c2").getDataValidation();
  if (rule == null) {
    rule = SpreadsheetApp
             .newDataValidation()
             .requireValueInList(['ON', 'OFF'], true);
  }

  //チャンネルデータの読込
  var xml = UrlFetchApp.fetch(targetUrl + cmdCh).getContentText();
  XmlService.parse(xml)
            .getRootElement()
            .getChild("ChItems")
            .getChildren("ChItem")
            .forEach(function(elem){
                var chId = elem.getChild("ChID").getText();
                var chName = elem.getChild("ChName").getText();
                if (chId in cList) { cList[chId][0] = chName; }
                else { cList[chId] = [chName, ""]; }
            });
  cData = [];
  for (var chId in cList) {
    var chDie = cList[chId];
    cData.push([chId, chDie[0], chDie[1]]);
  }
  
  //データの書込
  var range = sheet.getRange(2, 1, cData.length, 3);
  range.setValues(cData);
  range = sheet.getRange(2, 3, cData.length, 1);
  range.setDataValidation(rule);
  
  SpreadsheetApp.flush();
}
