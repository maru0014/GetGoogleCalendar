/**
 * Googleカレンダーからイベントを取得し、シートに出力します。
 * @throws {Error} 必須項目が入力されていない場合に例外が発生します。
 */
function main() {
  let resultCount = 0;

  // 設定_取得終了日に23:59:59をセット
  settings.endDate.setHours(23, 59, 59);

  // Start Loop カレンダー
  const lastRow = sheets.settingsCalendars.getLastRow();
  for (let i = 2; i <= lastRow; i++) {
    // 実行済みの場合はスキップ
    const status = sheets.settingsCalendars.getRange(i, 3).getValue();
    if (status) continue;

    // calendarIdを指定してイベント取得を実行
    const calendarId = sheets.settingsCalendars.getRange(i, 1).getValue();
    const calendarName = sheets.settingsCalendars.getRange(i, 2).getValue();
    try {
      const result = getCalendarEvents(calendarId, calendarName);
      sheets.settingsCalendars.getRange(i, 3).setValue(result);
      resultCount++;
    } catch (e) {
      console.error(calendarId, e);
      sheets.settingsCalendars.getRange(i, 3).setValue(`Error: ${e}`);
    }
  }
  // End Loop カレンダー

  // 処理件数を返す
  return resultCount;
}

/**
 * 指定されたカレンダーIDのイベントを取得してシートに出力します。
 * @param {String} calendarId - カレンダーID
 * @param {String} calendarName - カレンダー名
 * @return {String|Error} 取得結果
 */
function getCalendarEvents(calendarId, calendarName) {
  console.time(calendarId);
  ss.toast(`${calendarName} (${calendarId})`, "取得中...")
  let dataRows = [];
  const events = getEvents(calendarId, settings.startDate, settings.endDate, settings.query);

  // Start Loop イベント
  for (const event of events) {
    let { summary, start, end, location, htmlLink, description, attendees } = event;
    // 終日イベントを出力しない、かつ終日イベントの場合はスキップ
    if (!settings.putAllDayEvent && isAllDayEvent(event)) continue;
    // カテゴリ判定
    const category = categorize(summary, settingsCategory);
    // 日時表記を調整
    const startDateTime = dayjs.dayjs(start.dateTime).format("YYYY/MM/DD HH:mm");
    const endDateTime = dayjs.dayjs(end.dateTime).format("YYYY/MM/DD HH:mm");
    // 経過時間を計算
    const elapsed = calculateElapsedHours(start.dateTime, end.dateTime);
    // 概要を出力しない場合はクリア
    if (!settings.putDescription) description = "";
    // 参加者を出力する場合はカンマ区切りでセット
    const attendeesEmails = settings.putAttendees && attendees ? attendees.map(a => a.email).join(",") : "";
    // URLをリンクに変換
    const url = `=HYPERLINK("${htmlLink}","予定を開く")`;
    // データ行追加
    const row = [calendarId, calendarName, summary, category, startDateTime, endDateTime, elapsed, description, url, location, attendeesEmails];
    dataRows.push(row);
  }
  // End Loop イベント

  // 結果が0件でなければイベントシートに出力
  if (dataRows.length) {
    const lastRow = sheets.events.getLastRow();
    sheets.events.getRange(lastRow + 1, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
  }

  ss.toast(`${calendarName} (${calendarId}) の予定を ${dataRows.length} 件取得しました。`, "取得完了")
  console.timeEnd(calendarId);
  return `Success: get ${dataRows.length} events.`;
}

/**
 * 取得済みのイベントシートに対して再度カテゴリを設定する
 */
function reCategorize() {
  const table = sheets.events.getDataRange().getValues();
  if (table.length <= 1) {
    Browser.msgBox("イベントデータがありません。");
    return false;
  }

  const rows = table.slice(1);
  const tableBody = rows.map(row => {
    const summary = row[2];
    row[3] = categorize(summary, settingsCategory);
    return row;
  });

  sheets.events.getRange(2, 1, tableBody.length, tableBody[0].length).setValues(tableBody);
  return true;
}


/**
 * 指定した時刻範囲に予定があるかどうかを返します
 * @param {String} calendarId - カレンダーID
 * @param {Date} start - 開始日時
 * @param {Date} end - 終了日時
 * @param {String} keyword - キーワード
 * @return {Array} 予定配列
 */
function getEvents(calendarId, start, end, keyword) {
  const FORMAT_RFC3339 = "YYYY-MM-DDTHH:mm:ssZ";
  const timeZone = Session.getScriptTimeZone();
  const timeMin = dayjs.dayjs(start).format(FORMAT_RFC3339);
  const timeMax = dayjs.dayjs(end).format(FORMAT_RFC3339);
  const query = {
    timeZone, timeMin, timeMax,
    maxResults: 2500, // 最大2500 デフォルト250
    showDeleted: false, // 削除された予定も表示する
    singleEvents: true, // 繰り返しの予定を対象外にする
    orderBy: "startTime"
  };
  if (keyword) query["q"] = keyword; // キーワードがあれば追加

  let allEvents = [];
  let nextPageToken = null;

  do {
    const events = Calendar.Events.list(calendarId, { ...query, pageToken: nextPageToken });
    if (events.items) {
      allEvents = [...allEvents, ...events.items];
    }
    nextPageToken = events.nextPageToken;
  } while (nextPageToken);

  return allEvents;
}
