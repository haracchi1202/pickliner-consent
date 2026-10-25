/**
 * ピックライナー イベント参加同意書システム（バックエンド）
 *
 * 役割:
 *   - LIFF アプリから送られた同意データをスプレッドシート「同意記録」に追記
 *   - Messaging API の push で、本人のLINEトークに控えを送信
 *   - 控え送信の成否と理由を「控え送信結果」列に記録（届かない原因の切り分け用）
 *
 * 設定:
 *   - スクリプトプロパティ CHANNEL_ACCESS_TOKEN に Messaging API のチャネルアクセストークンを登録
 *     （未設定の場合は push をスキップし、記録のみ行う）
 */

var SHEET_NAME = "同意記録";
var HEADERS = ["受付日時", "受付番号", "氏名", "LINE表示名", "LINE userId", "同意書バージョン", "端末情報", "控え送信結果"];

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

    // --- 本人へ控えを push（トークンがあれば）。結果文字列を受け取る ---
    var pushResult = pushReceipt(userId, name, version, id, receivedAt);

    // --- スプレッドシートへ記録（控え送信結果も残す）---
    var sheet = getSheet();
    sheet.appendRow([receivedAt, id, name, displayName, userId, version, ua, pushResult]);

    return jsonOutput({ ok: true, id: id, push: pushResult });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

/**
 * シートを取得。無ければヘッダー付きで自動作成する。
 * 既存シートのヘッダーが古い（列が足りない）場合は上書きして最新にする。
 */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    return sheet;
  }
  // 1行目のヘッダーが最新の列数に満たなければ、ヘッダーだけ貼り直す
  var lastCol = sheet.getLastColumn();
  if (lastCol < HEADERS.length) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
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
 * 送信の成否と理由を表す文字列を返す（スプレッドシート記録用）。
 */
function pushReceipt(userId, name, version, id, receivedAt) {
  var token = PropertiesService.getScriptProperties().getProperty("CHANNEL_ACCESS_TOKEN");
  if (!token) {
    return "送信スキップ：トークン未設定（CHANNEL_ACCESS_TOKEN）";
  }
  if (!userId) {
    return "送信スキップ：userIdなし";
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

  try {
    var res = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code === 200) {
      return "送信OK";
    }
    // 失敗時は LINE が返すエラー本文をそのまま残す（原因特定用）
    return "送信失敗 HTTP" + code + "：" + res.getContentText();
  } catch (err) {
    return "送信例外：" + String(err);
  }
}

/**
 * JSON レスポンスを返す。
 */
function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
