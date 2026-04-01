// ============================================================
// Code.gs － 主程式
// 功能：讀取 Google Drive 上的 PDF、解析行程、寫入 Sheet、提供網頁
// ============================================================

// ── 設定區（請依實際情況修改）──────────────────────────────
const CONFIG = {
  PDF_FILE_NAME : 'Singapore_-_行程.pdf',   // Drive 上的 PDF 檔名
  SHEET_NAME    : 'Itinerary',              // 工作表名稱（不存在時自動建立）
  TRIP_TITLE    : 'Singapore 行程',
};

// ── 欄位定義（A-F，共 6 欄）────────────────────────────────
const HEADERS = ['Day', 'Type', 'Time', 'Location', 'Description', 'Weather'];

// ── 入口：doGet ─────────────────────────────────────────────
function doGet() {
  // 1. 解析 PDF → 取得結構化資料
  const rows = parsePdfItinerary();

  // 2. 寫入 Google Sheet
  if (rows.length > 0) {
    writeToSheet(rows);
  }

  // 3. 回傳 index.html 網頁（允許跨域嵌入）
  const html = HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG.TRIP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return html;
}

// ── 寫入 Google Sheet ────────────────────────────────────────
function writeToSheet(rows) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet  = ss.getSheetByName(CONFIG.SHEET_NAME);

  // 若工作表不存在，新建一張
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  sheet.clearContents();

  // 寫入標題列
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  headerRange.setFontWeight('bold')
             .setBackground('#1a3a5c')
             .setFontColor('#ffffff');

  // 寫入資料列
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  }

  // 自動調整欄寬
  HEADERS.forEach((_, i) => sheet.autoResizeColumn(i + 1));

  // 凍結標題列
  sheet.setFrozenRows(1);

  Logger.log(`✅ 已寫入 ${rows.length} 筆資料到工作表「${CONFIG.SHEET_NAME}」`);
}

// ── 供 index.html 呼叫：取得所有行程資料（JSON）──────────────
function getItineraryData() {
  const rows = parsePdfItinerary();
  return JSON.stringify(rows);
}

// ── 供 index.html 呼叫：取得 Sheet 現有資料 ─────────────────
function getSheetData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return JSON.stringify([]);

  const data = sheet.getDataRange().getValues();
  return JSON.stringify(data); // 第一列為 header
}

// ── 手動執行入口（不走 doGet，直接更新 Sheet）────────────────
function manualSync() {
  const rows = parsePdfItinerary();
  if (rows.length > 0) {
    writeToSheet(rows);
    SpreadsheetApp.getUi().alert(`✅ 同步完成，共 ${rows.length} 筆行程資料已寫入工作表。`);
  } else {
    SpreadsheetApp.getUi().alert('⚠️ 未解析到任何資料，請確認 PDF 檔案名稱與內容。');
  }
}
