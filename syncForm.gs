
//----------------------------------------------------
//フォーム同期プログラム

//#カテゴリ名の本の配列を取得
function getCategoryArray(category){
  //カテゴリ名のシートを取得
  var sheet = ss.getSheetByName(category);
  // A行の2行目からコンテンツをもつ最後の行までの値を配列で取得する
  var categoryArray = sheet.getRange(3,2,sheet.getLastRow() - 2).getValues();
  Logger.log(categoryArray);
  return categoryArray;
}


//#フォームの選択肢にカテゴリ別の本をぶち込む
function overwriteList(category,categoryArray) {

  /**
  // Googleフォームのプルダウン内の値を上書きする
  //
  **/

  // GoogleフォームのIDを設定　→「https://docs.google.com/forms/d/〇〇〇/edit」の〇〇〇を↓に記述
  var form = FormApp.openById('1QHUgJYF-13DBRR_PHrfTzeJxIAJvL2_yIQsyweE9UU8');

  // 質問項目がプルダウンのもののみ取得
  var items = form.getItems(FormApp.ItemType.LIST);

  items.forEach(function(item){
    // 質問項目が「好きなDJを選択して下さい」を含むものに対して、スプレッドシートの内容を反映する
    if(item.getTitle().match(category)){
      var listItemQuestion = item.asListItem();
      var choices = [];

      categoryArray.forEach(function(name){
        if(name != ""){
          choices.push(listItemQuestion.createChoice(name));
        }
      });

      // プルダウンの選択肢を上書きする
      listItemQuestion.setChoices(choices);
    }
  });
}

//#一斉同期
function sync(){
  //カテゴリー一覧を取得
  var glanceSheet = ss.getSheetByName('使い方＆貸出者一覧');
  var categorys = glanceSheet.getRange(7,2,1,glanceSheet.getLastColumn() - 1).getValues()[0];
  //Logger.log(categorys);
  //Logger.log(categorys.length);
  //カテゴリー別の本のリストを作成
  var categorysArrays = [];
  for(i = 0;i < categorys.length;i++){
    categorysArrays.push(getCategoryArray(categorys[i]));
  }
  //Logger.log(categorysArrays);
  for(i = 0;i < categorysArrays.length;i++){
    overwriteList(categorys[i],categorysArrays[i]);
  }
}

