/**
* メニュー追加
*/
function onOpen() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu("GAS");
  menu.addItem("イベントデータクリア", "clearSheetRows");
  menu.addItem("イベント取得", "menu_main");
  menu.addItem("イベント全件再取得", "menu_main2");
  menu.addItem("カテゴリー再判定", "menu_reCategorize");
  menu.addToUi();
}

function menu_main() {
  const menuName = "イベント取得";
  const confirm = Browser.msgBox(`${menuName} を実行しますか？`, Browser.Buttons.OK_CANCEL);
  if (confirm !== "ok") return;

  // 入力値チェック
  validateSettings();

  // 実行
  const resultCount = main();
  SpreadsheetApp.flush();
  if (resultCount === 0) {
    console.warn("[エラー] 取得対象カレンダーが0件でした。再取得する場合は処理結果の列をクリアしてください。");
    Browser.msgBox("[エラー] 取得対象カレンダーが0件でした。再取得する場合は処理結果の列をクリアしてください。");
    sheets.settingsCalendars.getRange("B2").activate();
  } else {
    Browser.msgBox(`[完了] ${menuName}。${resultCount}件のカレンダーを取得しました。`);
  }
}

function menu_main2() {
  const menuName = "イベント全件再取得";
  const confirm = Browser.msgBox(`${menuName} を実行しますか？`, Browser.Buttons.OK_CANCEL);
  if (confirm !== "ok") return;

  // 入力値チェック
  validateSettings();

  // イベントデータクリア
  clearSheetRows();

  // カレンダー処理結果列クリア
  sheets.settingsCalendars.getRange(2, 3, sheets.settingsCalendars.getLastRow(), 1).clearContent();

  // 実行
  const resultCount = main();
  SpreadsheetApp.flush();
  if (resultCount === 0) {
    console.warn("[エラー] 取得対象カレンダーが0件でした。再取得する場合は処理結果の列をクリアしてください。");
    Browser.msgBox("[エラー] 取得対象カレンダーが0件でした。再取得する場合は処理結果の列をクリアしてください。");
    sheets.settingsCalendars.getRange("B2").activate();
  } else {
    Browser.msgBox(`[完了] ${menuName}。${resultCount}件のカレンダーを取得しました。`);
  }
}

function menu_reCategorize() {
  const menuName = "カテゴリー再判定";
  const confirm = Browser.msgBox(`${menuName} を実行しますか？`, Browser.Buttons.OK_CANCEL);
  if (confirm !== "ok") return;

  const result = reCategorize();
  SpreadsheetApp.flush();
  if (!result) {
    Browser.msgBox("[エラー] 取得済みのイベントが0件です。");
  } else {
    Browser.msgBox(`[完了] ${menuName}`);
  }
}
