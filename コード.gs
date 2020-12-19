// シート取得
const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
const sheetSchedules = spreadSheet.getSheetByName("予定");
const sheetSettings = spreadSheet.getSheetByName("設定");

// 設定値取得
const calendarIds = sheetSettings.getRange(1, 2).getValue().split(","); // カレンダーID配列
const reExclusion = new RegExp(sheetSettings.getRange(2, 2).getValue()); // 除外ワード
const allDayExclusion = sheetSettings.getRange(3, 2).getValue(); // 終日の予定除外
const startDate = new Date(sheetSettings.getRange(4, 2).getValue()); // 取得開始日
const endDate = new Date(sheetSettings.getRange(5, 2).getValue()); // 取得終了日
const breakTimeThreshold = sheetSettings.getRange(6, 2).getValue(); // 休憩時間付与のしきい値

/**
 * メニュー追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi(); // UIクラス取得
  const menu = ui.createMenu("GAS"); // メニュー名セット
  menu.addItem("Googleカレンダー取得", "getGoogleCalendar"); // 関数セット
  menu.addToUi(); // スプレッドシートに反映
}

/**
 * 実行
 */
function getGoogleCalendar() {
  // シート2行名以下をクリア
  let lastRow = sheetSchedules.getLastRow();
  let lastColumn = sheetSchedules.getLastColumn();
  sheetSchedules.getRange(2, 1, lastRow, lastColumn).clearContent();

  // 配列初期化
  let table = new Array();

  // カレンダー数分ループ処理
  for (let i = 0; i < calendarIds.length; i++) {
    // 取得結果の配列を追記
    table = table.concat(
      fetchSchedules(calendarIds[i],reExclusion,allDayExclusion,startDate,endDate)
    );
  }

  if (table.length) {
    sheetSchedules.getRange(2, 1, table.length, table[0].length).setValues(table); // シートに出力
    Logger.log(`${table.length}件の予定を取得しました。`);
    spreadSheet.toast(`${table.length}件の予定を取得しました。`, 'Googleカレンダー取得完了', 5); // 完了メッセージ表示
  } else {
    Logger.log(`${table.length}件の予定を取得しました。`);
    spreadSheet.toast('取得結果が0件です。', 'Googleカレンダー取得完了', 5); // エラーメッセージ表示
  }
}

/**
 * GoogleカレンダーからgetEvents
 * @param {Array} calendarIds 取得対象カレンダーIDの配列
 * @return {Array} 取得結果の二次元配列
 */
function fetchSchedules(calendarId) {
  const schedules = new Array(); // 配列初期化
  const calendar = CalendarApp.getCalendarById(calendarId); // カレンダー
  const calendarName = calendar.getName(); // カレンダー名
  const events = calendar.getEvents(startDate, endDate); // 範囲内の予定を取得

  // 各予定のデータを配列に追加
  for (let i = 0; i < events.length; i++) {
    // 除外対象の場合は処理をスキップ
    if (isExclusion(events[i], reExclusion, allDayExclusion)) continue;

    let start = events[i].getStartTime();
    let end = events[i].getEndTime();

    let event = new Array();
    event.push(calendarName); // カレンダー名
    event.push(events[i].getTitle()); // 件名
    event.push(start); // 開始日時
    event.push(end); // 終了日時
    event.push(start.getMonth() + 1); // 月
    event.push(getOperatingTime(start, end)); // 時間数
    event.push(events[i].getDescription()); // 詳細

    schedules.push(event); // 配列に追加
  }

  return schedules;
}

/**
 * 取得対象の切り分け
 * @param {CalendarEvent} schedule 個別のCalendarEventクラス
 * @return {boolean} 真偽値
 */
function isExclusion(event) {
  // 終日イベントはスキップ
  if (allDayExclusion && event.isAllDayEvent()) return true;

  // 除外ワードを含む場合はスキップ
  if (reExclusion.test(event.getTitle())) return true;

  return false;
}

/**
 * 経過時間数の計算
 * @param {Date} start 予定の開始日時
 * @param {Date} end 予定の終了日時
 * @return {number} 経過時間数
 */
function getOperatingTime(start, end) {
  // 時間数算出
  const time = (end - start) / 1000 / 60 / 60;

  // 休憩時間の減算
  const operatingTime = time >= breakTimeThreshold ? time - 1 : time;

  return operatingTime;
}