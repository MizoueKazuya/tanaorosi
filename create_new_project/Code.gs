/**
 * 3行目からスプレッドシートのA列とB列を別シートにコピペする関数
 */
function copyColumnsABToAnotherSheet() {
  // 現在のスプレッドシートを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 現在のシート（コピー元）を取得
  const sourceSheet = spreadsheet.getActiveSheet();
  
  // コピー先のシート名（必要に応じて変更してください）
  const targetSheetName = "コピー先シート";
  
  // コピー先のシートを取得または作成
  let targetSheet = spreadsheet.getSheetByName(targetSheetName);
  if (!targetSheet) {
    targetSheet = spreadsheet.insertSheet(targetSheetName);
  }
  
  // コピー元のA列のデータを取得（3行目から最後の行まで）
  const lastRow = sourceSheet.getLastRow();
  
  // 3行目より下にデータがない場合は処理を終了
  if (lastRow < 3) {
    console.log("3行目以降にデータがありません");
    return;
  }
  
  // 3行目から最後の行までのA列とB列のデータを取得
  const sourceRange = sourceSheet.getRange(3, 1, lastRow - 2, 2); // 2列分取得
  const sourceData = sourceRange.getValues();
  
  // コピー先シートのA列とB列にデータを貼り付け
  const targetRange = targetSheet.getRange(1, 1, sourceData.length, sourceData[0].length);
  targetRange.setValues(sourceData);
  
  console.log(`${sourceData.length}行のデータをコピーしました`);
}

/**
 * より柔軟なバージョン：シート名と開始行を指定可能
 * @param {string} sourceSheetName - コピー元のシート名
 * @param {string} targetSheetName - コピー先のシート名
 * @param {number} startRow - 開始行（デフォルト: 3）
 */
function copyColumnAToAnotherSheetAdvanced(sourceSheetName, targetSheetName, startRow = 3) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // コピー元のシートを取得
  const sourceSheet = spreadsheet.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    throw new Error(`シート "${sourceSheetName}" が見つかりません`);
  }
  
  // コピー先のシートを取得または作成
  let targetSheet = spreadsheet.getSheetByName(targetSheetName);
  if (!targetSheet) {
    targetSheet = spreadsheet.insertSheet(targetSheetName);
  }
  
  // コピー元のA列のデータを取得
  const lastRow = sourceSheet.getLastRow();
  
  if (lastRow < startRow) {
    console.log(`${startRow}行目以降にデータがありません`);
    return;
  }
  
  // 指定した行から最後の行までのA列のデータを取得
  const sourceRange = sourceSheet.getRange(startRow, 1, lastRow - startRow + 1, 1);
  const sourceData = sourceRange.getValues();
  
  // コピー先シートのA列にデータを貼り付け
  const targetRange = targetSheet.getRange(1, 1, sourceData.length, 1);
  targetRange.setValues(sourceData);
  
  console.log(`${sourceData.length}行のデータを "${sourceSheetName}" から "${targetSheetName}" にコピーしました`);
}

/**
 * トリガー設定用の関数（手動実行用）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('A列とB列を別シートにコピー', 'copyColumnsABToAnotherSheet')
    .addToUi();
}
