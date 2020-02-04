//slackのアクセストークンを取得
var slackAccessToken = 'xoxp-139567565152-680949140533-859343407392-c0687b83037ecfae9276e325026511ca';
///対象チャンネルoption上下で入れ替え
var channelId = "#あさのの技術部屋";
var channelId = '#mj_library';



//slackにメッセージを送信
function sendMessage(message) {
  var slackApp = SlackApp.create(slackAccessToken);
  var options = {
    // 投稿するユーザーの名前
    username: "MJ文庫リマインダー"
  }
  slackApp.postMessage(channelId, message, options);
}



//#dateクラス、日数:変更された日付;dateクラスの二週間後の日付を返す関数
function changeDate(date,day){
  var changedDate = new Date(date.getYear(), date.getMonth(), date.getDate() + day);
  return changedDate;
}

//#シート、キーワード:キーワードの列;1列目の特定のキーワードがある列を取得
function getKeyCol(sheet,key) {
  //タイトルの配列を取得
  var titles = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues();
  //タイトルがある列を検索
  var key_col = titles[0].indexOf(key) + 1;
  //列の番号を返す
  return key_col;
}

//#シート、キーワード:キーワードの列;1列目の特定のキーワードがある列の配列での番号を取得
function getKeyNum(sheet,key) {
  //タイトルの配列を取得
  var titles = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues();
  //タイトルがある列を検索
  var key_num = titles[0].indexOf(key);
  //列の番号を返す
  return key_num;
}

//#メールアドレス:SlackID;メアドとSlackIDのhashを取得し、メアドに対応するIDを返す
function getMailToSlackId(mailAddress){
  var sheet = SpreadsheetApp.openById('1GqTU4fCCcXn2Rlzfwj5oPxPG0Pzz1KPFP_WeDxMmcTw').getActiveSheet();
  //配列の初期化
  var mail_slackIdHash = {};
  for(var i = 1; i <= sheet.getLastRow(); i++){
    //Hashのキーと値を取得
    mail_slackIdHash[sheet.getRange(i,1).getValue()] = sheet.getRange(i,3).getValue();
  }
  var mention = '<@' + mail_slackIdHash[mailAddress] + '>'
  return mention;
  /*
  for ( var key in mail_slackIdHash){
    Logger.log(key + ':' + mail_slackIdHash);
  }*/
}



//#メールアドレス:名前;メアドと名前のhashを取得
function getMailToName(mailAddress){
  var sheet = SpreadsheetApp.openById('1GqTU4fCCcXn2Rlzfwj5oPxPG0Pzz1KPFP_WeDxMmcTw').getActiveSheet();
  var lastRow = sheet.getLastRow();
  // 配列の初期化
  var mail_nameHash = {};
    for(var i = 2; i <= lastRow; i++){
      if(true){
        // 配列のkeyに対し値を設定する
        mail_nameHash[sheet.getRange(i, 1).getValue()] = sheet.getRange(i, 2).getValue();
      }
    }
  /*
  for (var key in mail_nameHash){
    Logger.log(key + ':' mail_nameHash[key]);
  }
  */
  //mailAddress = 'yuya.asano@aiesec.jp';
  return mail_nameHash[mailAddress];
}

//#シート:カテゴリタイトルは除くシートのデータ;空要素は除去
function getFormData(){
  var sheet = ss.getSheetByName('回答');
  var data = sheet.getSheetValues(2,1,sheet.getLastRow() - 1,sheet.getLastColumn());
  data.forEach(function (value, index) {
    data[index] = value.filter(Boolean);
    //Logger.log(data[index]);
  });
  //Logger.log(data);
  return data;
}

//#シート:カテゴリタイトルは除くシートのデータ;
function getOriginalFormData(){
  var sheet = ss.getSheetByName('回答');
  var data = sheet.getSheetValues(2,1,sheet.getLastRow() - 1,sheet.getLastColumn());
  //Logger.log(data);
  return data;
}


/*#メアドと名前のhashを取得
function getHash(){   
  var sheet = SpreadsheetApp.openById('1GqTU4fCCcXn2Rlzfwj5oPxPG0Pzz1KPFP_WeDxMmcTw').getActiveSheet();
  var lastRow = sheet.getLastRow();
  // 配列の初期化
  var hashColor2 = {}; 
  for(var i = 2; i <= lastRow; i++){
    if(true){  
      // 配列のkeyに対し値を設定する
      hashColor2[sheet.getRange(i, 1).getValue()] = sheet.getRange(i, 2).getValue();
    }
  } 
  return hashColor2;
}
*/