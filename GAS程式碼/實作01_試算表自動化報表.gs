/**
 * ============================================================
 * 實作 1：試算表自動化報表
 * ============================================================
 *
 * 功能說明：
 *   1. 學習自訂函式（Custom Function）
 *   2. 學習 appendRow() 和 setValues() 批次寫入
 *   3. 學習格式化試算表（粗體、背景色、欄寬、凍結列）
 *   4. 產生統計報表並輸出到試算表
 *
 * 試算表結構（自動建立）：
 *   工作表「銷售表」：品項名稱 | 數量 | 單價 | 小計
 *   工作表「簽到表」：序號 | 時間戳記 | 姓名 | 狀態 | 備註
 *
 * ⚠️ 注意：此程式需在 Google 試算表的「綁定專案」中執行
 *   （從試算表 → 擴充功能 → Apps Script 開啟的才是綁定專案）
 *
 * 使用方式：
 *   - 執行 createSampleData() → 建立範例銷售資料
 *   - 執行 generateReport() → 產生統計報表
 *   - 執行 createSignInSheet() → 建立簽到表
 *   - 執行 batchSignIn() → 批次匯入簽到
 *   - 執行 quickTest() → 一鍵完整測試
 * ============================================================
 */

// ============================================================
// 第一部分：自訂函式（Custom Function）
// ============================================================

/**
 * 自訂函式：計算小計（數量 × 單價）
 *
 * 在試算表中使用方式：
 *   在 D2 儲存格輸入 =calcSubtotal(B2, C2)
 *   就會自動計算 B2 × C2 的結果
 *
 * @param {number} quantity 數量
 * @param {number} price 單價
 * @return {number} 小計金額
 * @customfunction
 */
function calcSubtotal(quantity, price) {
  if (!quantity || !price) return 0;
  return quantity * price;
}

/**
 * 自訂函式：根據金額顯示等級
 *
 * 在試算表中使用方式：
 *   在某個儲存格輸入 =getLevel(D2)
 *
 * @param {number} amount 金額
 * @return {string} 等級文字
 * @customfunction
 */
function getLevel(amount) {
  if (amount >= 1000) return '高額';
  if (amount >= 500) return '中等';
  if (amount > 0) return '一般';
  return '-';
}

// ============================================================
// 第二部分：建立範例資料
// ============================================================

/**
 * 建立範例銷售資料（含格式化）
 */
function createSampleData() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // 清除現有資料
  sheet.clear();

  // 寫入標題
  var headers = ['品項名稱', '數量', '單價', '小計'];
  sheet.getRange(1, 1, 1, 4).setValues([headers]);

  // 格式化標題列
  var headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(12);
  headerRange.setBackground('#FF9800');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');
  headerRange.setBorder(true, true, true, true, true, true);

  // 寫入範例品項
  var items = [
    ['美式咖啡', 10, 60],
    ['拿鐵咖啡', 8, 80],
    ['珍珠奶茶', 15, 55],
    ['綠茶', 12, 30],
    ['紅茶', 6, 25]
  ];

  for (var i = 0; i < items.length; i++) {
    var row = i + 2;
    sheet.getRange(row, 1).setValue(items[i][0]);
    sheet.getRange(row, 2).setValue(items[i][1]);
    sheet.getRange(row, 3).setValue(items[i][2]);
    // D 欄使用自訂函式
    sheet.getRange(row, 4).setFormula('=calcSubtotal(B' + row + ',C' + row + ')');
  }

  // 設定欄寬
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 100);

  // 凍結標題列
  sheet.setFrozenRows(1);

  Logger.log('範例資料建立完成！共 ' + items.length + ' 筆品項。');
}

// ============================================================
// 第三部分：產生統計報表
// ============================================================

/**
 * 讀取銷售資料，產生統計報表寫入新工作表
 */
function generateReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    Logger.log('試算表中尚無資料，請先執行 createSampleData()');
    return;
  }

  // 讀取所有資料（從第 2 列開始，跳過標題）
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  var totalAmount = 0;
  var totalItems = 0;
  var maxItem = { name: '', amount: 0 };
  var itemList = [];

  for (var i = 0; i < data.length; i++) {
    var name = data[i][0];
    var qty = data[i][1];
    var price = data[i][2];
    var subtotal = qty * price;

    if (name === '') continue;

    totalAmount += subtotal;
    totalItems += qty;
    itemList.push({ name: name, qty: qty, subtotal: subtotal });

    if (subtotal > maxItem.amount) {
      maxItem = { name: name, amount: subtotal };
    }
  }

  // 建立報表工作表
  var reportName = '報表_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'MMdd');
  var reportSheet = ss.getSheetByName(reportName);
  if (reportSheet) ss.deleteSheet(reportSheet);
  reportSheet = ss.insertSheet(reportName);

  // 寫入報表標題
  reportSheet.getRange('A1').setValue('銷售統計報表');
  reportSheet.getRange('A1').setFontSize(16).setFontWeight('bold');

  var now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
  reportSheet.getRange('A2').setValue('產出時間：' + now);

  // 寫入統計數據
  reportSheet.getRange('A4').setValue('總金額');
  reportSheet.getRange('B4').setValue('$' + totalAmount.toLocaleString());
  reportSheet.getRange('A5').setValue('總數量');
  reportSheet.getRange('B5').setValue(totalItems + ' 個');
  reportSheet.getRange('A6').setValue('最高品項');
  reportSheet.getRange('B6').setValue(maxItem.name + '（$' + maxItem.amount + '）');

  // 格式化統計區
  reportSheet.getRange('A4:A6').setFontWeight('bold').setBackground('#E8F5E9');

  // 寫入明細
  reportSheet.getRange('A8').setValue('品項明細').setFontWeight('bold').setFontSize(12);
  var detailHeaders = ['品項', '數量', '小計'];
  reportSheet.getRange(9, 1, 1, 3).setValues([detailHeaders]);
  reportSheet.getRange(9, 1, 1, 3).setFontWeight('bold').setBackground('#1565C0').setFontColor('#fff');

  for (var j = 0; j < itemList.length; j++) {
    reportSheet.getRange(10 + j, 1).setValue(itemList[j].name);
    reportSheet.getRange(10 + j, 2).setValue(itemList[j].qty);
    reportSheet.getRange(10 + j, 3).setValue(itemList[j].subtotal);
  }

  // 設定欄寬
  reportSheet.setColumnWidth(1, 150);
  reportSheet.setColumnWidth(2, 100);
  reportSheet.setColumnWidth(3, 100);

  Logger.log('報表已產生至工作表「' + reportName + '」');
}

// ============================================================
// 第四部分：簽到表系統（appendRow + 格式化）
// ============================================================

var SIGNIN_SHEET_NAME = '簽到表';

/**
 * 建立簽到表（含格式化）
 */
function createSignInSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SIGNIN_SHEET_NAME);

  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(SIGNIN_SHEET_NAME);

  var headers = ['序號', '時間戳記', '姓名', '狀態', '備註'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(12);
  headerRange.setBackground('#1565C0');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');
  headerRange.setBorder(true, true, true, true, true, true);

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 200);
  sheet.setFrozenRows(1);

  Logger.log('簽到表建立完成！');
}

/**
 * 單筆簽到 — 使用 appendRow()
 * @param {string} name - 簽到者姓名
 * @param {string} note - 備註（選填）
 */
function signIn(name, note) {
  name = name || '未填姓名';
  note = note || '';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SIGNIN_SHEET_NAME);

  if (!sheet) {
    createSignInSheet();
    sheet = ss.getSheetByName(SIGNIN_SHEET_NAME);
  }

  var serialNumber = sheet.getLastRow();
  var now = new Date();
  var timestamp = now.toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  var hour = now.getHours();
  var minute = now.getMinutes();
  var status = (hour < 9 || (hour === 9 && minute === 0)) ? '準時' : '遲到';

  sheet.appendRow([serialNumber, timestamp, name, status, note]);

  var newRow = sheet.getLastRow();
  var newRange = sheet.getRange(newRow, 1, 1, 5);
  newRange.setHorizontalAlignment('center');
  newRange.setBorder(true, true, true, true, true, true);
  if (status === '遲到') newRange.setBackground('#FFEBEE');

  Logger.log('簽到成功：' + name + '（' + status + '）');
}

/**
 * 批次匯入多筆簽到 — 使用 setValues()
 */
function batchSignIn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SIGNIN_SHEET_NAME);

  if (!sheet) {
    createSignInSheet();
    sheet = ss.getSheetByName(SIGNIN_SHEET_NAME);
  }

  var now = new Date();
  var timestamp = now.toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  var startSerial = sheet.getLastRow();

  var batchData = [
    [startSerial + 0, timestamp, '陳大文', '準時', ''],
    [startSerial + 1, timestamp, '林小美', '準時', ''],
    [startSerial + 2, timestamp, '張志明', '遲到', '塞車'],
    [startSerial + 3, timestamp, '王小華', '準時', ''],
    [startSerial + 4, timestamp, '李建國', '遲到', '家裡有事']
  ];

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, batchData.length, 5).setValues(batchData);

  var newRange = sheet.getRange(startRow, 1, batchData.length, 5);
  newRange.setHorizontalAlignment('center');
  newRange.setBorder(true, true, true, true, true, true);

  for (var i = 0; i < batchData.length; i++) {
    if (batchData[i][3] === '遲到') {
      sheet.getRange(startRow + i, 1, 1, 5).setBackground('#FFEBEE');
    }
  }

  Logger.log('批次匯入完成：' + batchData.length + ' 筆');
}

// ============================================================
// 第五部分：快速測試
// ============================================================

/**
 * 一鍵測試：建立範例資料 → 產生報表
 */
function quickTest() {
  createSampleData();
  Utilities.sleep(1000);
  generateReport();
  Logger.log('測試完成！請查看新建的報表工作表。');
}
