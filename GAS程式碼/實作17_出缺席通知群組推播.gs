/**
 * ============================================================
 * 實作 10：到班出缺席通知 → Line Bot 群組推播
 * ============================================================
 *
 * 功能說明：
 *   1. 從試算表讀取出缺席記錄
 *   2. 自動篩選今日缺席名單
 *   3. 統計出缺席資料（出席率、缺席次數）
 *   4. 格式化缺席通知訊息
 *   5. 透過 Line Bot 推播到指定群組（群組推播）
 *   6. 只在有缺席時才推播（條件式觸發）
 *
 * 事前準備：
 *   1. 建立出缺席記錄試算表
 *      欄位：姓名 | 日期 | 狀態（出席/缺席/請假）| 備註
 *   2. 取得試算表 ID（網址列 /d/ 和 /edit 之間的那串）
 *   3. 取得 Line Bot 的 Channel Access Token
 *   4. 取得 Line 群組 ID（或個人 User ID）
 *
 * 如何取得群組 ID：
 *   1. 將 Line Bot 加入群組
 *   2. 使用「訊息收集器」（實作 2）記錄群組訊息
 *   3. 從試算表中找到 groupId 欄位
 *   或：在 doPost() 中 Logger.log(event.source.groupId)
 *
 * 設定排程自動執行：
 *   1. 在 Apps Script 編輯器中，點選左側「觸發條件」（鬧鐘圖示）
 *   2. 點選「新增觸發條件」
 *   3. 選擇函式：sendAttendanceReport
 *   4. 事件來源：時間驅動
 *   5. 時間型觸發條件類型：每日計時器
 *   6. 時段：下午 4 點到 5 點
 *   7. 儲存
 * ============================================================
 */

// ========== Line Bot 設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';

// 群組推播目標（可改為群組 ID）
// 個人推播用 User ID，群組推播用 Group ID
var PUSH_TARGET = LINE_USER_ID;  // 改成群組 ID 就能推到群組

// ========== 試算表設定 ==========
var SPREADSHEET_ID = '在此貼上你的試算表 ID';
var SHEET_NAME = '出缺席記錄';

// ============================================================
// 第一部分：Line Bot 推播函式
// ============================================================

/**
 * 推播文字訊息
 * 可以推給個人（User ID）或群組（Group ID）
 */
function pushText(to, text) {
  pushLine(to, [{ type: 'text', text: text }]);
}

/**
 * 推播底層函式
 */
function pushLine(to, messages) {
  var url = 'https://api.line.me/v2/bot/message/push';
  var options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({
      to: to,
      messages: messages
    }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();

  if (code !== 200) {
    Logger.log('推播失敗，狀態碼：' + code);
    Logger.log('錯誤內容：' + response.getContentText());
  } else {
    Logger.log('推播成功！（目標：' + to.substring(0, 8) + '...）');
  }
}

// ============================================================
// 第二部分：試算表操作
// ============================================================

/**
 * 初始化試算表（建立標題列）
 * 如果工作表不存在就自動建立
 */
function initSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // 設定標題列
    var headers = ['姓名', '日期', '狀態', '備註'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1b5e20')
      .setFontColor('white');
    sheet.setFrozenRows(1);

    Logger.log('已建立工作表：' + SHEET_NAME);
  }

  return sheet;
}

/**
 * 新增出缺席記錄（手動新增用）
 *
 * @param {string} name - 姓名
 * @param {string} status - 狀態（出席/缺席/請假）
 * @param {string} note - 備註
 */
function addRecord(name, status, note) {
  var sheet = initSheet();
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
  sheet.appendRow([name, today, status, note || '']);
  Logger.log('已新增記錄：' + name + ' - ' + status);
}

/**
 * 取得今日的出缺席記錄
 *
 * @returns {Object} { present: [], absent: [], leave: [], total: number }
 */
function getTodayAttendance() {
  var sheet = initSheet();
  var data = sheet.getDataRange().getValues();
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');

  var result = {
    present: [],  // 出席
    absent: [],   // 缺席
    leave: [],    // 請假
    total: 0
  };

  // 從第 2 列開始（跳過標題列）
  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    var date = data[i][1];
    var status = data[i][2];
    var note = data[i][3];

    // 日期比對（處理日期格式）
    var recordDate;
    if (date instanceof Date) {
      recordDate = Utilities.formatDate(date, 'Asia/Taipei', 'yyyy/MM/dd');
    } else {
      recordDate = String(date);
    }

    if (recordDate === today) {
      result.total++;
      if (status === '出席') {
        result.present.push(name);
      } else if (status === '缺席') {
        result.absent.push({ name: name, note: note });
      } else if (status === '請假') {
        result.leave.push({ name: name, note: note });
      }
    }
  }

  return result;
}

/**
 * 統計每個人的缺席次數
 *
 * @returns {Object} { '姓名': 缺席次數 }
 */
function getAbsenceCount() {
  var sheet = initSheet();
  var data = sheet.getDataRange().getValues();
  var counts = {};

  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    var status = data[i][2];

    if (status === '缺席') {
      counts[name] = (counts[name] || 0) + 1;
    }
  }

  return counts;
}

// ============================================================
// 第三部分：組合通知訊息
// ============================================================

/**
 * 組合出缺席通知訊息
 *
 * @returns {string|null} 通知訊息，沒有缺席時回傳 null
 */
function buildAttendanceMessage() {
  var attendance = getTodayAttendance();
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd (EEE)');

  // 如果今天沒有任何記錄
  if (attendance.total === 0) {
    Logger.log('今日沒有出缺席記錄');
    return null;
  }

  // 如果沒有缺席和請假，不推播
  if (attendance.absent.length === 0 && attendance.leave.length === 0) {
    Logger.log('今日全員出席，不需推播');
    return null;
  }

  var absenceCounts = getAbsenceCount();

  var message = '📋 出缺席通知\n';
  message += '📅 ' + today + '\n';
  message += '━━━━━━━━━━━━━━━\n';

  // 統計資訊
  var attendanceRate = Math.round(attendance.present.length / attendance.total * 100);
  message += '\n📊 今日統計：\n';
  message += '  ✅ 出席：' + attendance.present.length + ' 人\n';
  message += '  ❌ 缺席：' + attendance.absent.length + ' 人\n';
  message += '  📝 請假：' + attendance.leave.length + ' 人\n';
  message += '  📈 出席率：' + attendanceRate + '%\n';

  // 缺席名單
  if (attendance.absent.length > 0) {
    message += '\n❌ 缺席名單：\n';
    for (var i = 0; i < attendance.absent.length; i++) {
      var item = attendance.absent[i];
      var count = absenceCounts[item.name] || 1;
      message += '  • ' + item.name + '（累計 ' + count + ' 次）';
      if (item.note) message += '（' + item.note + '）';
      message += '\n';
    }
  }

  // 請假名單
  if (attendance.leave.length > 0) {
    message += '\n📝 請假名單：\n';
    for (var j = 0; j < attendance.leave.length; j++) {
      var leaveItem = attendance.leave[j];
      message += '  • ' + leaveItem.name;
      if (leaveItem.note) message += '（' + leaveItem.note + '）';
      message += '\n';
    }
  }

  message += '\n━━━━━━━━━━━━━━━';
  message += '\n💡 由 GAS 自動推播';

  return message;
}

// ============================================================
// 第四部分：主要功能 — 發送出缺席通知
// ============================================================

/**
 * 【主函式】發送出缺席通知到 Line
 * 排程觸發時執行這個函式
 *
 * 只在有缺席或請假時才推播（條件式觸發）
 */
function sendAttendanceReport() {
  Logger.log('===== 開始檢查出缺席 =====');

  var message = buildAttendanceMessage();

  if (message) {
    // 有缺席或請假，推播通知
    pushText(PUSH_TARGET, message);
    Logger.log('已推播出缺席通知');
  } else {
    Logger.log('全員出席或無記錄，不推播');
  }

  Logger.log('===== 出缺席檢查完成 =====');
}

// ============================================================
// 第五部分：測試函式
// ============================================================

/**
 * 測試：新增測試用的出缺席資料
 */
function testAddSampleData() {
  addRecord('王小明', '出席', '');
  addRecord('李大華', '出席', '');
  addRecord('張美美', '缺席', '身體不適');
  addRecord('陳志明', '出席', '');
  addRecord('林小芬', '請假', '家庭因素');
  Logger.log('已新增 5 筆測試資料');
}

/**
 * 測試：查看今日出缺席狀況
 */
function testTodayAttendance() {
  var result = getTodayAttendance();
  Logger.log('出席：' + result.present.join(', '));
  Logger.log('缺席：' + JSON.stringify(result.absent));
  Logger.log('請假：' + JSON.stringify(result.leave));
  Logger.log('總人數：' + result.total);
}

/**
 * 測試：發送出缺席通知
 */
function testSendReport() {
  sendAttendanceReport();
}

/**
 * 測試：推播文字
 */
function testPush() {
  pushText(LINE_USER_ID, '🧪 出缺席通知 Line Bot 連線測試成功！');
}
