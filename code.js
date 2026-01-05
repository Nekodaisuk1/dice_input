const SPREADSHEET_ID = 'your_spreadsheet_id_here'; // 同じブックにバインドするなら null でもOK
const DEFAULT_SHEET_NAME = 'シート1';   // デフォルト（UIで別シート選べる）
const DEFAULT_TARGET_COLUMN = 3;     // デフォルト列（UIで上書きされる）

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('番号入力')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function resetRowCache(){
  CacheService.getScriptCache().remove("lastRow");
  Logger.log("キャッシュ初期化完了");
}

function getSheetNames() {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().map(s => s.getName());
}

function getColumns(sheetName) {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName || DEFAULT_SHEET_NAME);
  const lastCol = sheet.getLastColumn();
  const cols = [];
  for (let i = 1; i <= lastCol; i++) {
    cols.push({
      name: String.fromCharCode(64 + i), // A,B,C...
      index: i
    });
  }
  return cols;
}

function appendNumber(num, sheetName, columnIndex) {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  const targetSheetName = sheetName || DEFAULT_SHEET_NAME;
  const targetCol = columnIndex ? Number(columnIndex) : DEFAULT_TARGET_COLUMN;

  const sheet = ss.getSheetByName(targetSheetName);

  const colValues = sheet.getRange(1, targetCol, sheet.getLastRow()).getValues();
  let lastFilledRow = 0;
  for (let i = colValues.length - 1; i >= 0; i--) {
    if (colValues[i][0] !== "" && colValues[i][0] != null) {
      lastFilledRow = i + 1;
      break;
    }
  }

  const rowToWrite = lastFilledRow + 1;

  sheet.getRange(rowToWrite, targetCol).setValue(num);

  return `シート「${targetSheetName}」の ${String.fromCharCode(64+targetCol)}列${rowToWrite}行 に ${num} を記録しました`;
}

