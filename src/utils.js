/**
 * Googleカレンダーの予定の件名を基にカテゴリを判定する関数
 * @param {String} summary - カレンダー予定の件名
 * @returns {String} category - 判定されたカテゴリ（マッチなしの場合は "未分類"）
 */
function categorize(summary, categories) {
  for (const row of categories) {
    const [strRegExp, category] = row;
    const regex = new RegExp(strRegExp, "g");
    if (regex.test(summary)) {
      return category;
    }
  }
  return "未分類";
}

/**
 * Google カレンダー API の Event オブジェクトが終日イベントかどうかを判定する
 * 
 * @param {Object} event - Google カレンダー API で取得したイベントオブジェクト
 * @param {Object} [event.start] - イベントの開始情報
 * @param {String} [event.start.date] - 終日イベントの場合の日付（例: "2025-02-09"）
 * @param {Object} [event.start.dateTime] - 通常イベントの場合の開始日時
 * @returns {Boolean} - 終日イベントなら true、そうでなければ false
 */
function isAllDayEvent(event) {
  return !!event.start?.date;
}

/**
 * dayjsを用いて開始日時と終了日時の経過時間数を計算する関数
 * @param {String} startTime - 開始日時（ISO 8601 形式）
 * @param {String} endTime - 終了日時（ISO 8601 形式）
 * @returns {Number} 経過時間数
 */
function calculateElapsedHours(start, end) {
  const startDate = dayjs.dayjs(start);
  const endDate = dayjs.dayjs(end);
  const elapsedHours = endDate.diff(startDate, "second", true);
  return elapsedHours / (60 * 60 * 24);
}

/**
 * シートとセルの存在を確認し、存在しない場合はエラーを通知する関数
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {String} sheetName - シート名
 */
function getSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート '${sheetName}' が見つかりません。`);
  }
  return sheet;
}

/**
 * シートとセルの存在を確認し、存在しない場合はエラーを通知する関数
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {String} cellName - 確認するセル範囲名
 */
function getCell(ss, cellName) {
  try {
    return ss.getRange(cellName);
  } catch (e) {
    throw new Error(`セル '${cellName}' が見つかりません。`);
  }
}

/**
 * シートのn行目以降をクリアする関数
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象のシート
 * @param {Integer} startRow - 開始行番号（default: 2）
 */
function clearSheetRows(sheet = sheets.events, startRow = 2) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(startRow, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  ss.toast(`シート「${sheet.getName()}」のデータをクリアしました。`, "データクリア")
}
