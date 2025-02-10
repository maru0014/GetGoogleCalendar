// シート定義
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheets = {
  settingsCalendars: getSheet(ss, "設定_カレンダー"),
  settingsSearch: getSheet(ss, "設定_検索"),
  settingsCategory: getSheet(ss, "設定_カテゴリ"),
  events: getSheet(ss, "イベント"),
}

// 設定読み込み
const settings = {
  startDate: new Date(getCell(ss, "設定_取得開始日").getValue()),
  endDate: new Date(getCell(ss, "設定_取得終了日").getValue()),
  query: getCell(ss, "設定_検索キーワード").getValue(),
  putAllDayEvent: getCell(ss, "設定_終日イベントを出力する").getValue(),
  putDescription: getCell(ss, "設定_概要を出力する").getValue(),
  putAttendees: getCell(ss, "設定_参加者を出力する").getValue(),
}

// カテゴリ定義を読み込み
const settingsCategory = sheets.settingsCategory.getDataRange().getValues().slice(1);

/**
 * 必須項目チェック
 */ 
function validateSettings() {
  if (!settings.startDate) {
    getCell(ss, "設定_取得開始日").activate();
    Browser.msgBox("設定_取得開始日 を入力してください");
    throw new Error("設定_取得開始日 を入力してください");
  }
  
  if (!settings.endDate) {
    getCell(ss, "設定_取得終了日").activate();
    Browser.msgBox("設定_取得終了日 を入力してください");
    throw new Error("設定_取得終了日 を入力してください");
  }

  if (!(settings.endDate instanceof Date) || isNaN(settings.endDate.getTime())) {
    getCell(ss, "設定_取得終了日").activate();
    Browser.msgBox("設定_取得終了日 に日付を入力してください");
    throw new Error("設定_取得終了日 に日付を入力してください");
  }
}
