var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('回答');


//------------------------------------------------------
//リマインドプログラム
function mentionTest(){
  var sh = SpreadsheetApp.openById('1GqTU4fCCcXn2Rlzfwj5oPxPG0Pzz1KPFP_WeDxMmcTw').getActiveSheet();
  var IDs = sh.getRange(1,3,47).getValues();
  var message = 'メンションテストを行います\n';
  for(var i = 0; i < IDs.length; i++){
    if(IDs[i] != 'undefined'){
      var mention = '<@' + IDs[i] + '>さん\n';
      message += mention;
    }
  }
  //Logger.log(message);
  var ore = '<@' + getMailToSlackId('yuya.asano@aiesec.jp') + '>';
  sendMessage(ore);
}
//空無し
function getMessageInfo(array,owner){
  //カテゴリー(表示しない)
  //var category = getCategory(lastArray);
  //貸出者
  var borrower = getMailToSlackId(array[1 +0]) + 'さん';
  //貸出本
  var book = array[0 +0];
  //所有者※あとで書く
  //var owner = getOwner(category,book);
  //var owner = ;
  //貸出日
  //var checkOutDay = value[2];
  var checkOutDay = Utilities.formatDate(array[2 +0],'JST','MM月dd日');
  //返却期限
  //var deadLine = value[3];
  var deadLine = Utilities.formatDate(array[3 +0],'JST','MM月dd日');  
  //メッセージ
  var message = 
  '\n[貸出者]' + borrower +
  '\n[所有者]' + owner +
  '\n[貸出本]' + book + 
  '\n[貸出日]' + checkOutDay + 
  '\n[返却期限]'　+ deadLine;
  return message;
}

//#フォームデータを解析して期限が過ぎた日時の配列の本のデータと貸出者を取得
function deadLineNotify(){
  //#空要素除去前フォームデータ
  var originalSheetData = getOriginalFormData();
  //#フォームデータを呼び出し
  var sheetData = getFormData();
  var dtLimit = new Date();  // 現在時刻を取得
  //#メアドの列を取得
  var mailCol = getKeyCol(sheet,'メールアドレス');
  var done = 'にAIESECMJ#mj_libraryへ送信完了';

  // シートの各行ごとにデータを取り出す
  sheetData.forEach(function(value, index) {
    // 送信完了していない、かつ送信予定日時が現在時刻より前ならば、メールを送信する
    if (!value[4 +0] && (new Date(value[3 +0])).getTime() < dtLimit.getTime()) {
      //MailApp.sendEmail(value[0], value[1], value[2]);  // メールを送信する
      sheet.getRange(2 + index,mailCol + 3 +0).setValue(dtLimit + done);  // 送信完了日時をシートに書く
      //空要素除去前の配列を用いて所有者を取得
      var oriAry = originalSheetData[index];
      var category = getCategory(oriAry);
      var book = getBook(oriAry);
      var owner = getOwner(category,book);
      messageContent = getMessageInfo(value,owner);
      message  =  '※テスト\n【返却期限のお知らせ】' + messageContent;
      
      
      /*
      //貸出者
      //var borrower = hash[value[1 +0]];
      var borrower = getMailToName(value[1 +0]);
      //貸出本
      var book = value[0 +0];
      //貸出日
      //var checkOutDay = value[2];
      var checkOutDay = Utilities.formatDate(value[2 +0],'JST','MM月dd日');
      //返却期限
      //var deadLine = value[3];
      var deadLine = Utilities.formatDate(value[3 +0],'JST','MM月dd日');
      //メッセージ
      var message = '※テスト\n【返却催促プログラム作動】\n<@' + borrower + '>\n[貸出本]' + book + '\n[貸出日]' + checkOutDay + '\n[返却期限]'　+ deadLine;
      //Logger.log(message);
      */
      //#メッセージをSlackに送信
      sendMessage(message);
      Logger.log(value + 'について返却リマインドプログラムを作動させました'); //どの行について処理したかログを出す
    }
  });
  //var sheetData = sheet.getSheetValues(startRows, 1, sheet.getLastRow(), sheet.getLastColumn());  // シートのデータを取得（2次元配列）  
}


//#フォームが送信される度に最終列に二週間後の日付をセットする関数
function set2weekAhead(){
  //#空要素を抹消したデータを取得
  var sheetData = getFormData();
  //一番最近に送られれてきたフォームデータのタイムスタンプを取得
  var lastArray = sheetData[sheetData.length - 1];
  var checkOutDate = lastArray[2];
  //Logger.log(checkOutDate);
  //#期日を2週間後に定義
  var deadLine = changeDate(checkOutDate,14);
  var deadCol = getKeyCol(sheet,'返却期限');
  var range = sheet.getRange(sheet.getLastRow(),deadCol);
  range.setValue(deadLine); 
}
//#フォーム送信時に本の所有者をセット

//#フォーム送信時の貸出通知
function checkOutNotify(){
  //空要素除去前データを取得
  var originalSheetData = getOriginalFormData();
  var originalLastArray = originalSheetData[originalSheetData.length -1];
  //#空配列消したデータを取得
  var sheetData = getFormData();
  //カテゴリを取得
  var category = getCategory(originalLastArray);
  //一番最近に送られてきたフォームデータを取得
  var lastArray = sheetData[sheetData.length - 1];
  //本を取得
  var book = lastArray[0 +0];
  //所有者
  var owner = getOwner(category,book)
  var messageContent = getMessageInfo(lastArray,owner);
  var message = '※テスト\n【貸出のお知らせ】' + messageContent;
  /*
  //貸出者
  var borrower = getMailToName(lastArray[1]);
  //貸出本
  var book = lastArray[0];
  //貸出日
  //var checkOutDay = value[2];
  var checkOutDay = Utilities.formatDate(lastArray[2],'JST','MM月dd日');
  //返却期限
  //var deadLine = value[3];
  var deadLine = Utilities.formatDate(lastArray[3],'JST','MM月dd日');
  //メッセージ
  var message = '※テスト\n【貸出通知プログラム作動】\n<@' + borrower + '>\n[貸出本]' + book + '\n[貸出日]' + checkOutDay + '\n[返却期限]'　+ deadLine;
  */
  //Logger.log(message);
  //#メッセージをSlackに送信
  sendMessage(message);
  Logger.log(lastArray + 'について貸出通知プログラムを作動させました'); //どの行について処理したかログを出す

}

function set2weekAhead_checkOutNotify(){
   set2weekAhead();
   checkOutNotify();
}
