// シート取得
const spreadSheet = SpreadsheetApp.getActiveSpreadsheet(); 
const sheetSchedules = spreadSheet.getSheetByName('予定').clear(); // シート取得と同時にクリア
const sheetSettings = spreadSheet.getSheetByName('設定');

// 設定取得
const calendarIds = sheetSettings.getRange(1, 2).getValue().split(); // カレンダーID配列
const reExclusion = new RegExp(sheetSettings.getRange(2, 2).getValue()); // 除外ワード
const allDayExclusion = sheetSettings.getRange(3, 2).getValue(); // 終日の予定除外
const startDate = new Date(sheetSettings.getRange(4, 2).getValue()); // 取得開始日
const endDate = new Date(sheetSettings.getRange(5, 2).getValue()); // 取得終了日
const breakTimeThreshold = sheetSettings.getRange(6, 2).getValue(); // 休憩時間付与のしきい値

// ヘッダー行定義
const FIRST_ROW = ['カレンダーID', '件名', '開始日時', '終了日時', '月', '時間数', '詳細']; // 取得結果1行目 


/**
 * メニュー追加
 */
function onOpen() {
  
  const ui = SpreadsheetApp.getUi(); // UIクラス取得
  const menu = ui.createMenu('GAS'); // メニュー名セット
  menu.addItem('Googleカレンダー取得開始', "runFetchSchedules"); // 関数セット
  menu.addToUi(); // スプレッドシートに反映
  
}


/**
 * 実行
 */
function runFetchSchedules(){
  
  const result = fetchSchedules(calendarIds, reExclusion, allDayExclusion, startDate, endDate); // 実行
  sheetSchedules.getRange(1, 1, result.length, result[0].length).setValues(result); // シートに出力
  spreadSheet.toast(`${result.length - 1}件の予定を取得しました。`, 'Googleカレンダー取得完了', 5); // 完了メッセージ表示
  
}


/**
 * GoogleカレンダーからgetEvents
 * @param {Array} calendarIds 取得対象カレンダーIDの配列
 * @param {RegExp} reExclusion 除外キーワードの正規表現オブジェクト
 * @param {Boolean} allDayExclusion 終日予定の除外フラグ
 * @param {data} startDate 個別のCalendarEventクラス
 * @param {data} endDate 個別のCalendarEventクラス
 * @return {Array} 取得結果の二次元配列
 */
function fetchSchedules(calendarIds, reExclusion, allDayExclusion, startDate, endDate) {
  
  const table = new Array(); // 配列初期化
        table.push(FIRST_ROW); // ヘッダー追加
  
  // カレンダー数分ループ処理
  for(let i = 0; i < calendarIds.length; i++) {
    const calendar = CalendarApp.getCalendarById(calendarIds[i]); // カレンダー
    const calendarName = calendar.getName(); // カレンダー名
    const events = calendar.getEvents(startDate, endDate); // 範囲内の予定を取得
    
    // 各予定のデータを配列に追加
    for(let ii = 0; ii < events.length; ii++) {
      
      if(isExclusion(events[ii], reExclusion ,allDayExclusion)) continue;
      
      let start = events[ii].getStartTime();
      let end = events[ii].getEndTime();
      
      let event = new Array();
          event.push(calendarName); // カレンダーID
          event.push(events[ii].getTitle()); // 件名
          event.push(start); // 開始日時
          event.push(end); // 終了日時
          event.push(start.getMonth() + 1); // 月
          event.push(getOperatingTime(start, end)); // 時間数
          event.push(events[ii].getDescription()); // 詳細
      
      table.push(event); // 配列に追加

    };
  };
  
  return table;
  
}


/**
 * 取得対象の切り分け
 * @param {CalendarEvent} schedule 個別のCalendarEventクラス
 * @return {boolean} 真偽値
 */
function isExclusion(event, reExclusion, allDayExclusion){
  
  if(allDayExclusion && event.isAllDayEvent()) return true; // 終日イベントはスキップ
  if(reExclusion.test(event.getTitle())) return true; // 除外ワードを含む場合はスキップ
  return false;
  
}


/**
 * 経過時間数の計算
 * @param {Date} start 予定の開始日時
 * @param {Date} end 予定の終了日時 
 * @return {number} 経過時間数
 */
function getOperatingTime(start, end){
  
  const time = (end - start) / 1000 / 60 / 60; // 時間数算出
  const operatingTime = time >= breakTimeThreshold ? time - 1 : time; // 休憩時間の減算
  return operatingTime;
  
}

