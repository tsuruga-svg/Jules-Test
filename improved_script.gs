// ==========================================
// ★初期設定エリア (ここを書き換えてください)
// ==========================================

// 1. PDFとCSVを入れる「GoogleドライブのフォルダID」
const FOLDER_ID = 'ここにフォルダIDを貼り付けてください';

// 2. デプロイした「ウェブアプリのURL」
const WEB_APP_URL = 'ここにウェブアプリのURLを貼り付けてください';

// 3. 処理済みのCSVを移動させるフォルダ名
const ARCHIVE_FOLDER_NAME = '処理済みCSV';

// 4. シート名の設定
const MAIN_SHEET_NAME = '発注一覧';
const MASTER_SHEET_NAME = '仕入先マスタ';


// ==========================================
// ★列名定義エリア (列を追加・変更した場合はここを修正)
// ==========================================
// ヘッダーの定義を一元管理し、コード内での「マジックナンバー」を排除します。
const SYSTEM_HEADERS = [
  "入力日", "発注NO", "処理区分", "発注日", "仕入先", "仕入先略称", "仕入・加工", "仕入・加工名",
  "部門", "部門名", "担当者", "担当者名", "相手先NO", "伝票種別", "納品先", "納品先略称",
  "入庫倉庫", "入庫倉庫名", "伝票備考", "取引区分", "商品", "商品名1", "商品名2", "ロットNO",
  "有効期限", "ロットNO備考", "税率区分", "税率", "発注数量", "単位", "発注単価", "発注金額",
  "受付NO", "単価更新区分", "明細備考1", "明細備考2", "指定納期", "納期回答日", "入荷完了区分",
  "仕入完了区分", "受注NO"
];

const APP_HEADERS = [
  "PDFリンク", "確認用リンク", "確認状況", "確認日時", "送付先メアド", "送信日時"
];

const ALL_HEADERS = SYSTEM_HEADERS.concat(APP_HEADERS);

// ==========================================
// ★ユーティリティ関数
// ==========================================

/**
 * シートのヘッダー行から「列名:列番号(1始まり)」の対応表オブジェクトを生成する
 * @param {Sheet} sheet 対象のシート
 * @returns {Object} {'列名1': 1, '列名2': 2, ...}
 */
function getColumnIndexes(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndexes = {};
  headers.forEach((header, index) => {
    if (header) {
      colIndexes[header] = index + 1; // 1始まりの列番号
    }
  });
  return colIndexes;
}


// ==========================================
// ★メイン機能
// ==========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('★発注管理メニュー')
    .addItem('1. 初期設定（ヘッダー作成）', 'setupHeaders')
    .addItem('2. データ取込（CSV & PDF）', 'importCSVAndLinkPDF')
    .addItem('3. メール一斉送信', 'sendEmailsToSuppliers')
    .addToUi();
}

// ---------------------------------------------------------
// 機能1: ヘッダー作成
// ---------------------------------------------------------
function setupHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MAIN_SHEET_NAME);
  }

  sheet.getRange(1, 1, 1, ALL_HEADERS.length).setValues([ALL_HEADERS]);

  // 色分け装飾
  sheet.getRange(1, 1, 1, SYSTEM_HEADERS.length).setBackground("#d9ead3");
  sheet.getRange(1, SYSTEM_HEADERS.length + 1, 1, APP_HEADERS.length).setBackground("#fff2cc");

  // マスタシート
  let masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    masterSheet = ss.insertSheet(MASTER_SHEET_NAME);
    masterSheet.getRange(1, 1, 1, 2).setValues([["仕入先名", "メールアドレス"]]).setBackground("#c9daf8");
  }

  SpreadsheetApp.getUi().alert("初期設定が完了しました。");
}

// ---------------------------------------------------------
// 機能2: CSV取込 & PDF自動紐付け
// ---------------------------------------------------------
function importCSVAndLinkPDF() {
  const ui = SpreadsheetApp.getUi();
  if (!FOLDER_ID || FOLDER_ID === 'ここにフォルダIDを貼り付けてください') {
    ui.alert('エラー: スクリプト内の FOLDER_ID が設定されていません。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const folder = DriveApp.getFolderById(FOLDER_ID);

  // 列番号マップを取得
  const col = getColumnIndexes(sheet);

  // --- 事前準備 ---
  // 1. 仕入先マスタ
  const masterData = masterSheet.getDataRange().getValues();
  const emailMap = masterData.slice(1).reduce((obj, row) => {
    const name = String(row[0]).trim();
    const mail = String(row[1]).trim();
    if (name) obj[name] = mail;
    return obj;
  }, {});

  // 2. PDFファイル一覧
  const pdfFiles = folder.getFilesByType(MimeType.PDF);
  const pdfMap = {};
  while (pdfFiles.hasNext()) {
    const file = pdfFiles.next();
    pdfMap[file.getName()] = file.getUrl();
  }

  // 3. 既存の発注NO一覧 (重複チェック用)
  const lastRow = sheet.getLastRow();
  const existingIds = (lastRow > 1) ?
    sheet.getRange(2, col["発注NO"], lastRow - 1, 1).getValues().flat().map(String) : [];

  // --- CSV処理 ---
  const csvFiles = folder.getFilesByType(MimeType.CSV);
  if (!csvFiles.hasNext()) {
    ui.alert('フォルダにCSVファイルが見つかりません。');
    return;
  }

  let totalNewRows = [];
  while (csvFiles.hasNext()) {
    const csvFile = csvFiles.next();
    const data = csvFile.getBlob().getDataAsString('Shift_JIS');
    const csvData = Utilities.parseCsv(data);

    // CSVの列定義 (0始まり) - 販売管理システムの仕様書に合わせて調整
    const CSV_COL_ORDER_NO = 1;
    const CSV_COL_SUPPLIER = 4;

    const newRows = csvData.slice(1).map(row => {
      const orderNo = String(row[CSV_COL_ORDER_NO]);
      const supplier = String(row[CSV_COL_SUPPLIER]);

      if (!orderNo || existingIds.includes(orderNo)) return null;

      const fullRow = new Array(ALL_HEADERS.length).fill("");
      row.forEach((cell, i) => {
        if (i < SYSTEM_HEADERS.length) fullRow[i] = cell;
      });

      // PDFリンクと確認用リンク
      for (const fileName in pdfMap) {
        if (fileName.includes(orderNo)) {
          const pdfUrl = pdfMap[fileName];
          fullRow[col["PDFリンク"] - 1] = pdfUrl;
          if (WEB_APP_URL) {
            fullRow[col["確認用リンク"] - 1] = `=HYPERLINK("${WEB_APP_URL}?id=${orderNo}&url="&ENCODEURL("${pdfUrl}"), "注文書を確認")`;
          }
          break;
        }
      }

      // メールアドレス
      if (emailMap[supplier]) {
        fullRow[col["送付先メアド"] - 1] = emailMap[supplier];
      }

      existingIds.push(orderNo); // 追加するデータも重複チェック対象に加える
      return fullRow;

    }).filter(row => row !== null); // null(スキップされた行)を除外

    totalNewRows = totalNewRows.concat(newRows);

    // 処理済みCSVを別フォルダへ移動
    let archiveFolder = folder.getFoldersByName(ARCHIVE_FOLDER_NAME);
    archiveFolder = archiveFolder.hasNext() ? archiveFolder.next() : folder.createFolder(ARCHIVE_FOLDER_NAME);
    csvFile.moveTo(archiveFolder);
  }

  // --- シートへの一括書き込み ---
  if (totalNewRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, totalNewRows.length, ALL_HEADERS.length).setValues(totalNewRows);
    ui.alert(`${totalNewRows.length} 件のデータを取り込みました。`);
  } else {
    ui.alert('新しい発注データはありませんでした。');
  }
}


// ---------------------------------------------------------
// 機能3: メール一斉送信
// ---------------------------------------------------------
function sendEmailsToSuppliers() {
  const ui = SpreadsheetApp.getUi();
  if (!WEB_APP_URL || WEB_APP_URL === 'ここにウェブアプリのURLを貼り付けてください') {
    ui.alert("エラー: WEB_APP_URL が設定されていません。");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (sheet.getLastRow() < 2) {
    ui.alert("送信対象のデータがありません。");
    return;
  }

  // --- データの一括読み込み ---
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const data = dataRange.getValues();
  const col = getColumnIndexes(sheet); // 列番号マップ

  let sentCount = 0;
  const now = new Date();

  // --- メイン処理 (配列内で行う) ---
  data.forEach((row, i) => {
    const email = row[col["送付先メアド"] - 1];
    const sentDate = row[col["送信日時"] - 1];
    const pdfUrl = row[col["PDFリンク"] - 1];

    if (email && !sentDate && pdfUrl) {
      const orderNo = row[col["発注NO"] - 1];
      const supplier = row[col["仕入先"] - 1];
      const confirmLink = `${WEB_APP_URL}?id=${orderNo}&url=${encodeURIComponent(pdfUrl)}`;

      const subject = `【注文書送付】株式会社アキツ 発注番号:${orderNo}`;
      const body = `${supplier} 御中\n\n`
                 + `いつも大変お世話になっております。\n株式会社アキツです。\n\n`
                 + `以下のリンクより注文書(PDF)をご確認ください。\n\n`
                 + `--------------------------------------------------\n`
                 + `■発注番号: ${orderNo}\n`
                 + `■注文書確認リンク:\n${confirmLink}\n`
                 + `--------------------------------------------------\n\n`
                 + `※リンクを開くと、弊社側に受領通知が届きます。\n（メールの返信や電話連絡は不要です）\n\n`
                 + `ご確認よろしくお願いいたします。`;

      try {
        GmailApp.sendEmail(email, subject, body);
        data[i][col["送信日時"] - 1] = now; // 配列内のデータを更新
        sentCount++;
      } catch (e) {
        console.error(`送信エラー: 発注NO ${orderNo}, 仕入先 ${supplier}, エラー: ${e.message}`);
      }
    }
  });

  // --- 更新結果の一括書き込み ---
  if (sentCount > 0) {
    dataRange.setValues(data); // 変更された配列をシートに一括で書き戻す
    ui.alert(`${sentCount} 件のメールを送信しました。`);
  } else {
    ui.alert('送信対象はありませんでした。');
  }
}

// ---------------------------------------------------------
// 機能4: Webアプリ (受領確認)
// ---------------------------------------------------------
function doGet(e) {
  const { id, url } = e.parameter;

  if (!id || !url) {
    return ContentService.createTextOutput("エラー: リンクのパラメータが無効です。");
  }

  try {
    const sheet = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId()).getSheetByName(MAIN_SHEET_NAME);
    const col = getColumnIndexes(sheet);
    const idColumn = col["発注NO"];

    const idData = sheet.getRange(2, idColumn, sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = idData.findIndex(orderNo => String(orderNo) === String(id));

    if (rowIndex !== -1) {
      const rowNum = rowIndex + 2; // ヘッダー分と0-index分を考慮
      const timestamp = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm");
      sheet.getRange(rowNum, col["確認状況"]).setValue("確認済");
      sheet.getRange(rowNum, col["確認日時"]).setValue(timestamp);
    }
  } catch(err) {
    console.error(`doGetエラー: 発注NO ${id} の処理中にエラーが発生しました。詳細: ${err.message}`);
  }

  // PDFへリダイレクトするHTMLを生成
  const html = `
    <!DOCTYPE html><html><head><title>注文書確認</title>
    <style>body{font-family:sans-serif;text-align:center;padding-top:50px;}.loader{border:5px solid #f3f3f3;border-top:5px solid #3498db;border-radius:50%;width:50px;height:50px;animation:spin 1s linear infinite;margin:20px auto;}@keyframes spin{0%{transform:rotate(0deg);}100%{transform:rotate(360deg);}}</style>
    </head><body><h2>受領確認を記録しました</h2><p>注文書(PDF)を開いています...</p><div class="loader"></div>
    <script>setTimeout(() => window.top.location.href = "${url}", 1000);</script>
    </body></html>`;

  return HtmlService.createHtmlOutput(html);
}
