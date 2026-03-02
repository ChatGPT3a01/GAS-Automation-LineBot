/**
 * ============================================================
 * 實作 5：將結果寫入試算表 + Line Bot 簽到回報
 * ============================================================
 *
 * 功能說明：
 *   1. 學習 appendRow() 和 getRange().setValues() 批次寫入
 *   2. 學習日期時間處理
 *   3. 學習格式化試算表（粗體、背景色、欄寬、凍結列）
 *   4. 簽到完成後透過 Line Bot 推播確認
 *
 * 試算表結構（自動建立）：
 *   工作表「簽到表」：序號 | 時間戳記 | 姓名 | 狀態 | 備註
 *
 * 使用方式：
 *   - 執行 signIn("王小明") → 自動簽到 + 推播確認到 Line
 *   - 執行 createSignInSheet() → 建立簽到表格式
 *   - 執行 batchSignIn() → 批次匯入多筆簽到
 *   - 執行 sendSignInSummary() → 推播今日簽到統計
 * ============================================================
 */

// ========== Line Bot 設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';

// ========== 試算表設定 ==========
var SIGNIN_SHEET_NAME = '簽到表';

// ============================================================
// 第一部分：建立簽到表
// ============================================================

/**
 * 建立簽到表（含格式化）
 * 示範：setFontWeight、setBackground、setColumnWidth、setFrozenRows
 */
function createSignInSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SIGNIN_SHEET_NAME);

  // 如果已存在，先刪除
  if (sheet) {
    ss.deleteSheet(sheet);
  }

  // 建立新的工作表
  sheet = ss.insertSheet(SIGNIN_SHEET_NAME);

  // ===== 標題列 =====
  var headers = ['序號', '時間戳記', '姓名', '狀態', '備註'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 格式化標題列
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');              // 粗體
  headerRange.setFontSize(12);                     // 字體大小
  headerRange.setBackground('#1565C0');            // 藍色背景
  headerRange.setFontColor('#FFFFFF');             // 白色文字
  headerRange.setHorizontalAlignment('center');    // 置中對齊
  headerRange.setBorder(true, true, true, true, true, true); // 框線

  // 設定欄寬
  sheet.setColumnWidth(1, 60);    // 序號
  sheet.setColumnWidth(2, 200);   // 時間戳記
  sheet.setColumnWidth(3, 120);   // 姓名
  sheet.setColumnWidth(4, 80);    // 狀態
  sheet.setColumnWidth(5, 200);   // 備註

  // 凍結標題列（捲動時標題不會消失）
  sheet.setFrozenRows(1);

  Logger.log('簽到表建立完成！');
  pushText(LINE_USER_ID, '✅ 簽到表已建立完成！\n\n可以開始使用 signIn() 函式進行簽到。');
}

// ============================================================
// 第二部分：單筆簽到（appendRow）
// ============================================================

/**
 * 單筆簽到 — 使用 appendRow() 在最後一列新增資料
 *
 * @param {string} name - 簽到者姓名
 * @param {string} note - 備註（選填）
 */
function signIn(name, note) {
  // 預設值
  name = name || '未填姓名';
  note = note || '';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SIGNIN_SHEET_NAME);

  // 如果簽到表不存在，先建立
  if (!sheet) {
    createSignInSheet();
    sheet = ss.getSheetByName(SIGNIN_SHEET_NAME);
  }

  // 計算序號（目前最後一列 - 1 = 資料筆數，再 + 1 = 新序號）
  var serialNumber = sheet.getLastRow();

  // 取得目前時間
  var now = new Date();
  var timestamp = now.toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

  // 判斷簽到狀態（9:00 前為準時，之後為遲到）
  var hour = now.getHours();
  var minute = now.getMinutes();
  var status = (hour < 9 || (hour === 9 && minute === 0)) ? '✅ 準時' : '⚠️ 遲到';

  // 使用 appendRow() 新增一列
  sheet.appendRow([serialNumber, timestamp, name, status, note]);

  // 格式化新加入的那一列
  var newRow = sheet.getLastRow();
  var newRange = sheet.getRange(newRow, 1, 1, 5);
  newRange.setHorizontalAlignment('center');
  newRange.setBorder(true, true, true, true, true, true);

  // 如果遲到，用淡紅色標記
  if (status.indexOf('遲到') !== -1) {
    newRange.setBackground('#FFEBEE');
  }

  Logger.log('簽到成功：' + name + '（' + status + '）');

  // 推播 Line Bot 簽到確認
  var msg = '📋 簽到成功！\n';
  msg += '━━━━━━━━━━━━\n';
  msg += '👤 姓名：' + name + '\n';
  msg += '⏰ 時間：' + timestamp + '\n';
  msg += '📌 狀態：' + status + '\n';
  msg += '📊 今日第 ' + serialNumber + ' 位簽到';

  pushText(LINE_USER_ID, msg);
}

// ============================================================
// 第三部分：批次寫入（getRange().setValues()）
// ============================================================

/**
 * 批次匯入多筆簽到 — 使用 setValues() 一次寫入多列
 * 比起逐筆 appendRow()，效率高出許多
 */
function batchSignIn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SIGNIN_SHEET_NAME);

  if (!sheet) {
    createSignInSheet();
    sheet = ss.getSheetByName(SIGNIN_SHEET_NAME);
  }

  // 準備批次資料
  var now = new Date();
  var timestamp = now.toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  var startSerial = sheet.getLastRow();

  var batchData = [
    [startSerial + 0, timestamp, '陳大文', '✅ 準時', ''],
    [startSerial + 1, timestamp, '林小美', '✅ 準時', ''],
    [startSerial + 2, timestamp, '張志明', '⚠️ 遲到', '塞車'],
    [startSerial + 3, timestamp, '王小華', '✅ 準時', ''],
    [startSerial + 4, timestamp, '李建國', '⚠️ 遲到', '家裡有事']
  ];

  // 使用 setValues() 一次寫入所有資料
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, batchData.length, 5).setValues(batchData);

  // 格式化新寫入的資料
  var newRange = sheet.getRange(startRow, 1, batchData.length, 5);
  newRange.setHorizontalAlignment('center');
  newRange.setBorder(true, true, true, true, true, true);

  // 遲到的列標記淡紅色
  for (var i = 0; i < batchData.length; i++) {
    if (batchData[i][3].indexOf('遲到') !== -1) {
      sheet.getRange(startRow + i, 1, 1, 5).setBackground('#FFEBEE');
    }
  }

  Logger.log('批次匯入完成：' + batchData.length + ' 筆');

  // 推播通知
  pushText(LINE_USER_ID, '📋 批次簽到完成！\n\n共匯入 ' + batchData.length + ' 筆簽到紀錄。');
}

// ============================================================
// 第四部分：統計與推播
// ============================================================

/**
 * 推播今日簽到統計
 */
function sendSignInSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SIGNIN_SHEET_NAME);

  if (!sheet || sheet.getLastRow() <= 1) {
    pushText(LINE_USER_ID, '📋 簽到統計：尚無簽到紀錄。');
    return;
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

  var total = data.length;
  var onTime = 0;
  var late = 0;

  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][3]);
    if (status.indexOf('準時') !== -1) {
      onTime++;
    } else if (status.indexOf('遲到') !== -1) {
      late++;
    }
  }

  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  var summary = '📊 簽到統計報表\n';
  summary += '━━━━━━━━━━━━\n';
  summary += '📅 ' + now + '\n\n';
  summary += '👥 總簽到人數：' + total + ' 人\n';
  summary += '✅ 準時：' + onTime + ' 人\n';
  summary += '⚠️ 遲到：' + late + ' 人\n';
  summary += '📈 出席率：' + Math.round(total / total * 100) + '%\n';
  summary += '⏱ 準時率：' + Math.round(onTime / total * 100) + '%';

  pushText(LINE_USER_ID, summary);
  Logger.log('統計已推播');
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

// ========== 快速測試 ==========

/**
 * 一鍵測試：建立簽到表 → 批次匯入 → 推播統計
 */
function quickTest() {
  createSignInSheet();
  Utilities.sleep(1000);  // 等待 1 秒
  batchSignIn();
  Utilities.sleep(1000);
  sendSignInSummary();
}
