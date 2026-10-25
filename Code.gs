/**
 * ピックライナー イベント参加同意書システム（バックエンド）
 *
 * 役割:
 *   - LIFF アプリから送られた同意データをスプレッドシート「同意記録」に追記
 *   - Messaging API の push で、本人のLINEトークに控えを送信
 *
 * 設定:
 *   - スクリプトプロパティ CHANNEL_ACCESS_TOKEN に Messaging API のチャネルアクセストークンを登録
 *     （未設定の場合は push をスキップし、記録のみ行う）
 */

var SHEET_NAME = "同意記録";
var HEADERS = ["受付日時", "受付番号", "氏名", "LINE表示名", "LINE userId", "同意書バージョン", "端末情報"];

/**
 * 動作確認用。ブラウザで GAS URL を開くと表示される。
 */
function doGet() {
  return ContentService
    .createTextOutput("consent endpoint ok")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * LIFF アプリからの POST を受け取る。
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var id = makeReceiptId();
    var now = new Date();
    var tz = "Asia/Tokyo";
    var receivedAt = Utilities.formatDate(now, tz, "yyyy/MM/dd HH:mm:ss");

    var name = data.name || "";
    var displayName = data.displayName || "";
    var userId = data.userId || "";
    var version = data.version || "";
    var ua = data.ua || "";

    // --- スプレッドシートへ記録 ---
    var sheet = getSheet();
    sheet.appendRow([receivedAt, id, name, displayName, userId, version, ua]);

    // --- 本人へ控えを push（トークンがあれば） ---
    pushReceipt(userId, name, version, id, receivedAt);

    return jsonOutput({ ok: true, id: id });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

/**
 * シートを取得。無ければヘッダー付きで自動作成する。
 */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

/**
 * 8桁の英数字の受付番号を作る（UUID の先頭を大文字化）。
 */
function makeReceiptId() {
  var uuid = Utilities.getUuid().replace(/-/g, "");
  return uuid.substring(0, 8).toUpperCase();
}

/**
 * Messaging API の push で本人に控えを送る。
 * CHANNEL_ACCESS_TOKEN が未設定なら何もしない。
 */
function pushReceipt(userId, name, version, id, receivedAt) {
  var token = PropertiesService.getScriptProperties().getProperty("CHANNEL_ACCESS_TOKEN");
  if (!token || !userId) {
    return; // トークン未設定、または userId 不明なら push しない
  }

  var text =
    "【参加同意を受け付けました】\n" +
    "お名前：" + name + "\n" +
    "同意書バージョン：" + version + "\n" +
    "受付番号：" + id + "\n" +
    "受付日時：" + receivedAt + "（日本時間）\n\n" +
    "当日はこの画面を受付でご提示ください。";

  var payload = {
    to: userId,
    messages: [{ type: "text", text: text }]
  };

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

/**
 * JSON レスポンスを返す。
 */
function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
