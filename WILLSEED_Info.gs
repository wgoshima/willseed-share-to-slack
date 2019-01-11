function sendWillseedInformation() {
  // 取り扱いたいスプレッドシートのURL
  var spreadSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1MvXjKO0vhANloQfBEOddyeEowDRZ4DGDYaVxo4Uwr5o/edit#gid=429064649");
  var sheet = spreadSheet.getSheetByName("■事例共有リスト");

  // 送信対象の事例を先頭行から探す
  var row = 1;

  do {
    status = sheet.getRange(++row, 2).getValue();
    url = sheet.getRange(row, 6).getValue();
  } while ((status != ""　||  url == "") && row < 100);
  
  if (row >= 100) {
    Browser.msgBox("送信対象が見当たらないため終了しました。");
    return;
  }
  
  // 共有情報の整形
  var dataArray = sheet.getRange(row, 2, 1, 7).getValues(); //共有情報の行の2列目～8列目を取得
  
  var channel = dataArray[0][1];  // 投稿先チャンネル
  var unfurlFlg = true;  // サムネイル展開の有無
  var message = "@here: \n";
  
//  Logger.log(dataArray);

  // サムネイル展開の列がOFFのときに展開しないよう設定
  if(dataArray[0][6] == "OFF") {
    unfurlFlg = false;
  }
  
  for (i = 2; i < 6; i++) {
    data = dataArray[0][i];
//    Logger.log("data" + i + " : " + data);
    if(data != "") {
      message += data + "\n";
    }
  }
  
//  Logger.log(message);

  var yourSelections = Browser.msgBox("Slackに送信してよろしいですか？ :" + dataArray[0][4], Browser.Buttons.OK_CANCEL);
  if(yourSelections == "cancel") {
    Browser.msgBox("キャンセルしました。");
    return;
  }

  // Slackに送信
  sendToSlack(message, channel, unfurlFlg);
  
  // スプレッドシート上のステータスを送信済にする
  sheet.getRange(row, 2).setValue("送信済");
  sheet.getRange(row, 9).setValue(new Date());
  sheet.getRange(row, 10).setValue(Session.getActiveUser().getEmail());

  Browser.msgBox("送信しました！");
}

//Slackに送信
function sendToSlack( messages, channel, unfurlFlg ){
  var url        = "https://slack.com/api/chat.postMessage";
  var token      = "xoxp-2334628001-6550917015-119976036933-e05eacbcde8b11ce8e800f1e71f15955"; //取得したtoken
  var channel    = "#" + channel; //　メッセージ投稿先のチャネル名
  var botname   = "☆WILLSEED事務局からのお知らせ☆"; //BOTの名前
  var parse      = "full";
  var icon_emoji = ":willseed1:"; // Botのアイコンを指定
  var method     = "post";
 
  var payload = {
      "token"      : token,
      "channel"    : channel,
      "text"       : messages,
      "username"   : botname,
      "parse"      : parse,
      "icon_emoji" : icon_emoji,
    "unfurl_links" : unfurlFlg,
    "unfurl_media" : unfurlFlg,
  };
  
  var params = {
      "method" : method,
      "payload" : payload
  };
 
  //slackにポスト
  var response = UrlFetchApp.fetch(url, params);
}

