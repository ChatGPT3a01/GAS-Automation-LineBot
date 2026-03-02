/**
 * ============================================================
 * 實作 3：產出第一個 GAS + Line Bot 通知
 * ============================================================
 *
 * 功能說明：
 *   1. 學習 GAS 最基本的試算表讀寫操作
 *   2. 學習自訂函式（Custom Function）
 *   3. 將試算表中的統計結果透過 Line Bot 推播
 *
 * 試算表結構（請先手動建立）：
 *   A 欄：品項名稱（例：咖啡、紅茶、綠茶）
 *   B 欄：數量（例：10、5、8）
 *   C 欄：單價（例：60、30、25）
 *   D 欄：小計（使用自訂函式計算）
 *
 * 使用方式：
 *   1. 在試算表中輸入品項、數量、單價
 *   2. 在 D2 儲存格輸入 =calcSubtotal(B2, C2) 計算小計
 *   3. 執行 sendDailyReport() 將統計結果推播到 Line
 * ============================================================
 */

// ========== Line Bot 設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';

// ============================================================
// 第一部分：基本讀寫操作
// ============================================================

/**
 * 讀取單一儲存格的值
 * 示範：讀取 A1 儲存格
 */
function readCell() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var value = sheet.getRange('A1').getValue();
  Logger.log('A1 的值是：' + value);

  // 推播到 Line
  pushText(LINE_USER_ID, '📖 讀取結果\nA1 儲存格的值是：' + value);
}

/**
 * 寫入單一儲存格
 * 示範：在 E1 儲存格寫入文字
 */
function writeCell() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('E1').setValue('這是 GAS 寫入的文字');
  sheet.getRange('E2').setValue(new Date());   // 寫入目前時間
  sheet.getRange('E3').setValue(12345);         // 寫入數字

  Logger.log('已成功寫入 E1~E3');

  // 推播到 Line
  pushText(LINE_USER_ID, '✏️ 寫入完成！\n已在 E1~E3 儲存格寫入資料。');
}

/**
 * 讀取一個範圍的資料
 * 示範：讀取 A1:C5 的所有資料
 */
function readRange() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getRange('A1:C5').getValues();

  // data 是一個二維陣列
  Logger.log('共讀取 ' + data.length + ' 列資料');

  var message = '📊 試算表資料：\n\n';
  for (var i = 0; i < data.length; i++) {
    // data[i] 是第 i 列的陣列
    if (data[i][0] !== '') {
      message += data[i][0] + ' | ' + data[i][1] + ' | ' + data[i][2] + '\n';
    }
  }

  Logger.log(message);

  // 推播到 Line
  pushText(LINE_USER_ID, message);
}

// ============================================================
// 第二部分：自訂函式（Custom Function）
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
  if (amount >= 1000) return '🔴 高額';
  if (amount >= 500) return '🟡 中等';
  if (amount > 0) return '🟢 一般';
  return '-';
}

// ============================================================
// 第三部分：統計並推播報表
// ============================================================

/**
 * 統計試算表資料並透過 Line Bot 推播每日報表
 *
 * 這個函式會：
 *   1. 讀取試算表中所有品項資料
 *   2. 計算總金額、最高金額品項
 *   3. 將報表透過 Line Bot 推播
 */
function sendDailyReport() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();

  // 如果沒有資料（只有標題列）
  if (lastRow <= 1) {
    pushText(LINE_USER_ID, '📊 報表通知：試算表中尚無資料。');
    return;
  }

  // 讀取所有資料（從第 2 列開始，跳過標題）
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  var totalAmount = 0;
  var totalItems = 0;
  var maxItem = { name: '', amount: 0 };
  var itemList = [];

  for (var i = 0; i < data.length; i++) {
    var name = data[i][0];     // A 欄：品項名稱
    var qty = data[i][1];      // B 欄：數量
    var price = data[i][2];    // C 欄：單價
    var subtotal = qty * price; // 計算小計

    if (name === '') continue;

    totalAmount += subtotal;
    totalItems += qty;
    itemList.push(name + ' ×' + qty + ' = $' + subtotal);

    if (subtotal > maxItem.amount) {
      maxItem = { name: name, amount: subtotal };
    }
  }

  // 組合報表訊息
  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  var report = '📊 每日銷售報表\n';
  report += '━━━━━━━━━━━━\n';
  report += '📅 ' + now + '\n\n';

  // 各品項明細
  for (var j = 0; j < itemList.length; j++) {
    report += '  • ' + itemList[j] + '\n';
  }

  report += '\n━━━━━━━━━━━━\n';
  report += '💰 總金額：$' + totalAmount.toLocaleString() + '\n';
  report += '📦 總數量：' + totalItems + ' 個\n';
  report += '🏆 最高品項：' + maxItem.name + '（$' + maxItem.amount + '）';

  // 推播到 Line
  pushText(LINE_USER_ID, report);

  Logger.log('報表已推播至 Line');
  Logger.log(report);
}

// ============================================================
// 第四部分：寫入範例資料（方便測試）
// ============================================================

/**
 * 建立範例資料（執行一次即可）
 */
function createSampleData() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // 清除現有資料
  sheet.clear();

  // 寫入標題
  var headers = ['品項名稱', '數量', '單價', '小計'];
  sheet.getRange(1, 1, 1, 4).setValues([headers]);
  sheet.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#FF9800')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');

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

  Logger.log('範例資料建立完成！');

  // 推播通知
  pushText(LINE_USER_ID, '✅ 範例資料已建立完成！\n共 ' + items.length + ' 筆品項。\n\n請執行 sendDailyReport() 來測試報表推播。');
}

// ========== Line Bot 推播函式（共用模組）==========

function pushText(to, text) {
  pushLine(to, [{ type: 'text', text: text }]);
}

function pushLine(to, messages) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({ to: to, messages: messages }),
    muteHttpExceptions: true
  });
}
