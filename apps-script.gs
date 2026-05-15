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
 * ★ 初回1回だけ実行 ★
 * リアルタイム集計シート「ダッシュボード」を作成（COUNTIF式で自動更新）
 * Apps Scriptエディタで関数名「setupDashboard」を選んで実行
 */
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName("ダッシュボード");
  if (!dash) dash = ss.insertSheet("ダッシュボード", 0); // 先頭に表示
  dash.clear();

  // タイトル
  dash.getRange("A1:C1").merge()
    .setValue("📊 成竹診断ツール  リアルタイム集計")
    .setFontSize(16).setFontWeight("bold")
    .setBackground("#0c4a6e").setFontColor("#fff")
    .setHorizontalAlignment("center");
  dash.setRowHeight(1, 36);

  // ファネルセクション
  dash.getRange("A3").setValue("🔻 ファネル（人数の流れ）")
    .setFontWeight("bold").setFontSize(13).setBackground("#fde047");
  dash.getRange("A3:C3").merge();

  dash.getRange("A4:C4").setValues([["イベント", "人数", "診断完了比"]])
    .setFontWeight("bold").setBackground("#f3f4f6");

  const funnelRows = [
    ["1. ページ閲覧",          "page_view"],
    ["2. 診断スタート",        "start"],
    ["3. 診断完了 ⭐",          "diagnosed"],
    ["4-A. LINEへ移行",        "click_line"],
    ["4-B. 電話で予約クリック", "click_phone"],
    ["4-C. HPで詳しく見る",    "click_hp"],
    ["4-D. 地図を確認",        "click_map"],
  ];

  funnelRows.forEach((row, i) => {
    const r = 5 + i;
    dash.getRange(r, 1).setValue(row[0]);
    dash.getRange(r, 2).setFormula(`=COUNTIF(log!C:C, "${row[1]}")`);
    // 「診断完了」を分母にしたコンバージョン率
    dash.getRange(r, 3).setFormula(`=IFERROR(IF(B${r}=0,"-",TEXT(B${r}/B$7, "0.0%")), "-")`);
  });

  // 強調: 診断完了行
  dash.getRange("A7:C7").setBackground("#fef3c7").setFontWeight("bold");
  // 強調: LINE移行行
  dash.getRange("A8:C8").setBackground("#dcfce7");

  // タイプ分布セクション
  const tStart = 14;
  dash.getRange(tStart, 1).setValue("🦁 タイプ別 診断分布")
    .setFontWeight("bold").setFontSize(13).setBackground("#fde047");
  dash.getRange(tStart, 1, 1, 3).merge();

  dash.getRange(tStart + 1, 1, 1, 2).setValues([["タイプ名", "人数"]])
    .setFontWeight("bold").setBackground("#f3f4f6");

  const types = [
    "ぎっくり腰オフィスゴリラ", "在宅ぎっくり仔猫", "PC疲れ・美の白鳥", "肩こり美ナマケモノ",
    "万年肩こり鋼鉄カブト", "だるおも在宅パンダ", "むくみ脱出スワン", "冷えとり美ペンギン",
    "スポーツ捻挫チーター", "やんちゃ突発ジャガー", "動ける美のシカ", "日焼け美フラミンゴ",
    "働きすぎ鋼の馬", "サボリ上手のキツネ", "キレイ目アスリート狼", "自由気ままな美ライオン",
  ];
  types.forEach((t, i) => {
    const r = tStart + 2 + i;
    dash.getRange(r, 1).setValue(t);
    dash.getRange(r, 2).setFormula(`=COUNTIF(log!E:E, "${t}")`);
  });

  // 列幅
  dash.setColumnWidth(1, 240);
  dash.setColumnWidth(2, 90);
  dash.setColumnWidth(3, 110);

  // logシートのテスト行（test-session）を削除
  const log = ss.getSheetByName("log");
  if (log) {
    const data = log.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === "test-session" || data[i][2] === "test") {
        log.deleteRow(i + 1);
      }
    }
  }

  SpreadsheetApp.getUi().alert("✅ ダッシュボードを作成しました！\n左下のシートタブから「ダッシュボード」をご確認ください。");
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
