// ==========================================
// ★初期設定エリア (ここを書き換えてください)
// ==========================================

/** 1. CSVとPDFが保存されているGoogleドライブのフォルダID */
const SOURCE_FOLDER_ID = 'ここにフォルダIDを貼り付けてください';

/** 2. デプロイしたウェブアプリのURL */
const WEB_APP_URL = 'ここにウェブアプリのURLを貼り付けてください';

/** 3. 処理済みのCSVを移動させるフォルダ名 (SOURCE_FOLDER_IDの中に自動作成) */
const ARCHIVE_FOLDER_NAME = '処理済みCSV';

/** 4. シート名の設定 */
const MAIN_SHEET_NAME = '発注一覧';
const MASTER_SHEET_NAME = '仕入先マスタ';

/** 5. CSVファイル内で「発注NO」と「仕入先」が何列目にあるか (0から数えた番号) */
const CSV_COLUMN_INDEX = {
  ORDER_NO: 1,  // 発注NO (例: B列にある場合は 1)
  SUPPLIER: 4   // 仕入先 (例: E列にある場合は 4)
};

// ==========================================
// ★システムヘッダー定義エリア
// ==========================================
// スプレッドシートのヘッダー構成です。列の追加や名称変更はここを修正してください。
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
 * シートのヘッダー行から「列名:列番号(1始まり)」の対応表オブジェクトを生成します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象のシート
 * @returns {Object.<string, number>} {'列名1': 1, '列名2': 2, ...}
 */
function getColumnIndexes(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce((obj, header, index) => {
    if (header) obj[header] = index + 1;
    return obj;
  }, {});
}

/**
 * 2次元配列の各要素から不要な空白や制御文字を削除します。
 * @param {string[][]} data 処理対象の2次元配列
 * @returns {string[][]} クレンジング済みの2次元配列
 */
function cleanData(data) {
  const controlCharRegex = /[\uFEFF\u0000-\u001F\u007F-\u009F]/g;
  return data.map(row =>
    row.map(cell => (typeof cell === 'string') ? cell.replace(controlCharRegex, '').trim() : cell)
  );
}

// ==========================================
// ★メイン機能
// ==========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('★発注管理システム')
    .addItem('1. 初期設定（ヘッダー作成）', 'setupHeaders')
    .addItem('2. CSVデータ取込', 'importFromCsvAndLinkPdf')
    .addItem('3. メール一斉送信', 'sendEmailsToSuppliers')
    .addToUi();
}

/**
 * 機能1: ヘッダー作成
 */
function setupHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(MAIN_SHEET_NAME);

  sheet.getRange(1, 1, 1, ALL_HEADERS.length).setValues([ALL_HEADERS]);
  sheet.getRange(1, 1, 1, SYSTEM_HEADERS.length).setBackground("#d9ead3");
  sheet.getRange(1, SYSTEM_HEADERS.length + 1, 1, APP_HEADERS.length).setBackground("#fff2cc");

  let masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    masterSheet = ss.insertSheet(MASTER_SHEET_NAME);
    masterSheet.getRange(1, 1, 1, 2).setValues([["仕入先名", "メールアドレス"]]).setBackground("#c9daf8");
  }
  SpreadsheetApp.getUi().alert("初期設定が完了しました。");
}

/**
 * 機能2: CSV取込 & PDF自動紐付け
 */
function importFromCsvAndLinkPdf() {
  const ui = SpreadsheetApp.getUi();
  if (!SOURCE_FOLDER_ID || SOURCE_FOLDER_ID.includes('貼り付け')) {
    ui.alert('エラー: スクリプト内の SOURCE_FOLDER_ID が設定されていません。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || !masterSheet) {
    ui.alert('エラー: 「発注一覧」または「仕入先マスタ」シートが見つかりません。');
    return;
  }

  try {
    const sourceFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
    const csvFiles = sourceFolder.getFilesByType(MimeType.CSV);
    if (!csvFiles.hasNext()) {
      ui.alert('フォルダにCSVファイルが見つかりません。');
      return;
    }

    // --- 事前準備 ---
    const col = getColumnIndexes(sheet);
    if (!col["発注NO"]) {
       ui.alert(`エラー: シート「${MAIN_SHEET_NAME}」に「発注NO」の列が見つかりません。`);
       return;
    }

    const masterData = masterSheet.getDataRange().getValues();
    const emailMap = masterData.slice(1).reduce((obj, row) => {
      const name = String(row[0]).trim();
      if (name) obj[name] = String(row[1]).trim();
      return obj;
    }, {});

    const pdfFiles = sourceFolder.getFilesByType(MimeType.PDF);
    const pdfMap = {};
    while (pdfFiles.hasNext()) {
      const file = pdfFiles.next();
      pdfMap[file.getName()] = file.getUrl();
    }

    const lastRow = sheet.getLastRow();
    const existingIds = (lastRow > 1) ?
      new Set(sheet.getRange(2, col["発注NO"], lastRow - 1, 1).getValues().flat().map(String)) : new Set();

    let archiveFolder = sourceFolder.getFoldersByName(ARCHIVE_FOLDER_NAME);
    archiveFolder = archiveFolder.hasNext() ? archiveFolder.next() : sourceFolder.createFolder(ARCHIVE_FOLDER_NAME);

    // --- CSV処理 ---
    let totalNewRows = [];
    let processedFileNames = [];

    while (csvFiles.hasNext()) {
      const csvFile = csvFiles.next();
      const csvBlob = csvFile.getBlob().getDataAsString('Shift_JIS');
      let csvData = Utilities.parseCsv(csvBlob);

      if (csvData.length > 1) {
        csvData.shift(); // ヘッダー除去
        const cleanedData = cleanData(csvData);

        const newRows = cleanedData.map(row => {
          const orderNo = row[CSV_COLUMN_INDEX.ORDER_NO];
          if (!orderNo || existingIds.has(orderNo)) return null;

          const supplier = row[CSV_COLUMN_INDEX.SUPPLIER];
          const fullRow = new Array(ALL_HEADERS.length).fill("");
          row.forEach((cell, i) => { if (i < SYSTEM_HEADERS.length) fullRow[i] = cell; });

          for (const fileName in pdfMap) {
            if (fileName.includes(orderNo)) {
              const pdfUrl = pdfMap[fileName];
              fullRow[col["PDFリンク"] - 1] = pdfUrl;
              if (WEB_APP_URL && !WEB_APP_URL.includes('貼り付け')) {
                fullRow[col["確認用リンク"] - 1] = `=HYPERLINK("${WEB_APP_URL}?id=${orderNo}&url="&ENCODEURL("${pdfUrl}"), "注文書を確認")`;
              }
              break;
            }
          }
          if (emailMap[supplier]) {
            fullRow[col["送付先メアド"] - 1] = emailMap[supplier];
          }

          existingIds.add(orderNo);
          return fullRow;
        }).filter(Boolean);

        totalNewRows = totalNewRows.concat(newRows);
      }

      csvFile.moveTo(archiveFolder);
      processedFileNames.push(csvFile.getName());
    }

    // --- シートへの一括書き込み ---
    if (totalNewRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, totalNewRows.length, ALL_HEADERS.length).setValues(totalNewRows);
    }

    // --- 最終結果の報告 ---
    let message = totalNewRows.length > 0 ? `${totalNewRows.length} 件の新規データを転記しました。\n` : '新しいデータはありませんでした。\n';
    if (processedFileNames.length > 0) message += `\n処理済みファイル:\n- ${processedFileNames.join('\n- ')}`;
    ui.alert(message);

  } catch(e) {
    Logger.log(e);
    ui.alert(`エラーが発生しました: ${e.message}`);
  }
}

/**
 * 機能3: メール一斉送信
 */
function sendEmailsToSuppliers() {
  const ui = SpreadsheetApp.getUi();
  if (!WEB_APP_URL || WEB_APP_URL.includes('貼り付け')) {
    ui.alert("エラー: WEB_APP_URL が設定されていません。");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (sheet.getLastRow() < 2) return ui.alert("送信対象のデータがありません。");

  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const data = dataRange.getValues();
  const col = getColumnIndexes(sheet);

  let sentCount = 0;
  const now = new Date();

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
        data[i][col["送信日時"] - 1] = now;
        sentCount++;
      } catch (e) {
        console.error(`送信エラー: 発注NO ${orderNo}, エラー: ${e.message}`);
      }
    }
  });

  if (sentCount > 0) {
    dataRange.setValues(data);
    ui.alert(`${sentCount} 件のメールを送信しました。`);
  } else {
    ui.alert('送信対象はありませんでした。');
  }
}

/**
 * 機能4: Webアプリ (受領確認)
 */
function doGet(e) {
  const { id, url } = e.parameter;
  if (!id || !url) return ContentService.createTextOutput("エラー: リンクのパラメータが無効です。");

  try {
    const sheet = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId()).getSheetByName(MAIN_SHEET_NAME);
    const col = getColumnIndexes(sheet);
    const idData = sheet.getRange(2, col["発注NO"], sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = idData.findIndex(orderNo => String(orderNo) === String(id));

    if (rowIndex !== -1) {
      const rowNum = rowIndex + 2;
      const timestamp = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm");
      sheet.getRange(rowNum, col["確認状況"]).setValue("確認済");
      sheet.getRange(rowNum, col["確認日時"]).setValue(timestamp);
    }
  } catch(err) {
    console.error(`doGetエラー: 発注NO ${id}, 詳細: ${err.message}`);
  }

  const html = `<!DOCTYPE html><html><head><title>注文書確認</title><style>body{font-family:sans-serif;text-align:center;padding-top:50px;}.loader{border:5px solid #f3f3f3;border-top:5px solid #3498db;border-radius:50%;width:50px;height:50px;animation:spin 1s linear infinite;margin:20px auto;}@keyframes spin{0%{transform:rotate(0deg);}100%{transform:rotate(360deg);}}</style></head><body><h2>受領確認を記録しました</h2><p>注文書(PDF)を開いています...</p><div class="loader"></div><script>setTimeout(()=>window.top.location.href="${url}",1e3)</script></body></html>`;
  return HtmlService.createHtmlOutput(html);
}
