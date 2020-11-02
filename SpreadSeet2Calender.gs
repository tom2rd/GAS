/* スプレッド表示の際に呼出し　*/
/* メニューに追加 */
function onOpen() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 
  //スプレッドシートのメニューにカスタムメニュー「カレンダー連携 > 実行」を作成
  var subMenus = [];
  subMenus.push({
    name: "実行",
    functionName: "createSchedule"  //実行で呼び出す関数を指定
  });
  ss.addMenu("カレンダー連携", subMenus);
}
 
/**
 * 予定を作成する
 */
function createSchedule() {
 
  // 連携するアカウント
  const gAccount = PropertiesService.getScriptProperties().getProperty('CALENDER_ID');  // ★★ここに連携するカレンダーのアドレスを入れる
  
  // 読み取り範囲（表の始まり行と終わり列）
  const topRow = 2;
  const lastCol = 8;
 
  // 0始まりで列を指定しておく
  const statusCellNum = 0;
  const dayCellNum = 1;
  const startCellNum = 3;
  const endCellNum = 4;
  const titleCellNum = 5;
  const locationCellNum = 6;
  const descriptionCellNum = 7;
 
  // シートを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 
  // 予定の最終行を取得
  var lastRow = sheet.getLastRow();
  
  //予定の一覧を取得
  var contents = sheet.getRange(topRow, 1, sheet.getLastRow(), lastCol).getValues();
 
  // googleカレンダーの取得
  var calender = CalendarApp.getCalendarById(gAccount);
 
  //順に予定を作成（今回は正しい値が来ることを想定）
  for (i = 0; i <= lastRow - topRow; i++) {
 
    //「済」っぽいのか、空の場合は飛ばす
    var status = contents[i][statusCellNum]; 
    if (
      status == "済" ||
      status == "済み" ||
      status == "OK" ||
      contents[i][dayCellNum] == ""
    ) {
      continue;
    }
 
    // 値をセット 日時はフォーマットして保持
    var day = new Date(contents[i][dayCellNum]);
    var startTime = contents[i][startCellNum];
    var endTime = contents[i][endCellNum];
    var title = contents[i][titleCellNum];
    // 場所と詳細をセット
    var options = {location: contents[i][locationCellNum], description: contents[i][descriptionCellNum]};
    
    try {
      // 開始終了が無ければ終日で設定
      if (startTime == '' || endTime == '') {
        //予定を作成
        calender.createAllDayEvent(
          title,
          new Date(day),
          options
        );
        
      // 開始終了時間があれば範囲で設定
      } else {
        // 開始日時をフォーマット
        var startDate = new Date(day);
        startDate.setHours(startTime.getHours())
        startDate.setMinutes(startTime.getMinutes());
        // 終了日時をフォーマット
        var endDate = new Date(day);
        endDate.setHours(endTime.getHours())
        endDate.setMinutes(endTime.getMinutes());
        // 予定を作成
        calender.createEvent(
          title,
          startDate,
          endDate,
          options
        );
      }
 
      //無事に予定が作成されたら「済」にする
      sheet.getRange(topRow + i, 1).setValue("済");
 
    // エラーの場合（今回はログ出力のみ）
    } catch(e) {
      Logger.log(e);
    }
    
  }
  // ブラウザへ完了通知
  Browser.msgBox("完了");
}
