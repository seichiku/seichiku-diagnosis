/**
 * 成竹カラダMBTI診断ツール - データ蓄積用 Google Apps Script
 * ---------------------------------------------------------
 * セットアップ手順（5分で完了）:
 *
 * 1. Googleドライブで新しいスプレッドシートを作成
 *    名前は「成竹診断ログ」など分かりやすく
 *
 * 2. メニュー「拡張機能」→「Apps Script」を開く
 *
 * 3. 表示されたエディタの中身を全部消して、このファイルの中身を貼り付け
 *
 * 4. ファイル名は何でもOK。「保存」(💾)を押す
 *
 * 5. 「デプロイ」(右上の青ボタン) → 「新しいデプロイ」
 *      種類: ウェブアプリ
 *      説明: 診断ツール用
 *      次のユーザーとして実行: 自分
 *      アクセス権: 全員（重要！匿名アクセス必須）
 *    → 「デプロイ」を押す
 *
 * 6. 表示された「ウェブアプリのURL」(https://script.google.com/.../exec) をコピー
 *
 * 7. index.html を開き、以下の行を編集:
 *      const GAS_ENDPOINT = "";
 *    ↓
 *      const GAS_ENDPOINT = "コピーしたURL";
 *
 * 8. 完了！診断結果が自動でスプレッドシートに溜まります。
 */

const SHEET_NAME = "log";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "timestamp", "session_id", "event",
        "type_code", "type_name", "answers"
      ]);
      sheet.setFrozenRows(1);
      sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#0c4a6e").setFontColor("#fff");
    }
    sheet.appendRow([
      data.ts || new Date(),
      data.sid || "",
      data.event || "",
      data.code || "",
      data.typeName || "",
      data.answers || ""
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput("成竹診断ツール ログ受信エンドポイント (OK)");
}

/**
 * 集計用: ファネル数（手動実行 or トリガーで定期実行）
 * 実行すると「summary」シートに集計結果を出力
 */
function generateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAME);
  if (!logSheet) return;
  const data = logSheet.getDataRange().getValues();
  const headers = data.shift();
  const evIdx = headers.indexOf("event");
  const typeIdx = headers.indexOf("type_name");

  const counts = {};
  data.forEach(row => {
    const ev = row[evIdx];
    counts[ev] = (counts[ev] || 0) + 1;
  });
  const typeCounts = {};
  data.filter(r => r[evIdx] === "diagnosed").forEach(r => {
    const t = r[typeIdx];
    typeCounts[t] = (typeCounts[t] || 0) + 1;
  });

  let sum = ss.getSheetByName("summary");
  if (!sum) sum = ss.insertSheet("summary");
  sum.clear();
  sum.appendRow(["更新日時", new Date()]);
  sum.appendRow([]);
  sum.appendRow(["ファネル", "件数", "コンバージョン率"]);
  const diagnosed = counts.diagnosed || 0;
  const order = ["page_view", "start", "diagnosed", "click_line", "click_phone", "click_hp", "click_map"];
  order.forEach(ev => {
    const c = counts[ev] || 0;
    const cv = diagnosed > 0 ? (c / diagnosed * 100).toFixed(1) + "%" : "-";
    sum.appendRow([ev, c, cv]);
  });
  sum.appendRow([]);
  sum.appendRow(["タイプ分布", "件数"]);
  Object.entries(typeCounts)
    .sort((a, b) => b[1] - a[1])
    .forEach(([t, c]) => sum.appendRow([t, c]));
}
