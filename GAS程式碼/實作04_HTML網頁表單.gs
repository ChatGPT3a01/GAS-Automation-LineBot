/**
 * ============================================================
 * 實作 4：HTML 表單 + Line Bot 即時通知
 * ============================================================
 *
 * 功能說明：
 *   1. 用 HtmlService 建立一個公開的 HTML 表單網頁
 *   2. 使用者填寫表單 → 資料自動寫入 Google 試算表
 *   3. 每次有人填寫 → Line Bot 即時推播通知管理員
 *
 * 檔案結構（需在 Apps Script 中建立兩個檔案）：
 *   - Code.gs：此檔案（後端程式）
 *   - form.html：前端網頁
 *     ⚠️ 在 Apps Script 中必須命名為「form」（不含 .html）
 *     內容請參考下載包中的 04_HTML表單_前端.html
 *
 * 試算表結構（自動建立）：
 *   時間戳記 | 姓名 | Email | 訊息內容
 *
 * 部署步驟：
 *   1. 建立一個 Google 試算表
 *   2. 開啟 Apps Script
 *   3. 在 Code.gs 貼上此程式碼
 *   4. 新增 HTML 檔案：點選「+」→「HTML」→ 命名為「form」
 *   5. 將 04_HTML表單_前端.html 的內容貼入 form.html
 *   6. 修改 LINE_TOKEN 和 LINE_USER_ID
 *   7. 部署為網頁應用程式（存取權限：所有人）
 *   8. 開啟部署網址即可看到表單
 * ============================================================
 */

// ========== Line Bot 設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
var LINE_GROUP_ID = '在此貼上你的 Group ID（不需要群組通知就留空白）';

// ========== 試算表設定 ==========
var FORM_SHEET_NAME = '表單回覆';

// ============================================================
// 第一部分：提供 HTML 頁面
// ============================================================

/**
 * 處理 GET 請求 — 顯示 HTML 表單頁面
 * 當使用者開啟部署網址時，就會看到這個表單
 */
function doGet(e) {
  var html = HtmlService.createHtmlOutputFromFile('form');
  html.setTitle('聯絡表單');
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

// ============================================================
// 第二部分：處理表單提交
// ============================================================

/**
 * 處理表單提交的資料
 * 由前端 HTML 透過 google.script.run.processForm(data) 呼叫
 *
 * @param {Object} formData - 表單資料 { name, email, message }
 * @returns {Object} 處理結果 { success, message }
 */
function processForm(formData) {
  try {
    // 基本輸入驗證
    if (!formData.name || !formData.email || !formData.message) {
      return { success: false, message: '請填寫所有必填欄位。' };
    }
    if (formData.email.indexOf('@') === -1) {
      return { success: false, message: 'Email 格式不正確。' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(FORM_SHEET_NAME);

    // 如果工作表不存在，建立新的
    if (!sheet) {
      sheet = createFormSheet(ss);
    }

    // 取得目前時間
    var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

    // 寫入試算表
    sheet.appendRow([
      now,
      formData.name,
      formData.email,
      formData.message
    ]);

    // 取得目前是第幾筆
    var rowCount = sheet.getLastRow() - 1; // 扣掉標題列

    // 透過 Line Bot 推播通知管理員
    var notification = '📬 收到新的表單填寫！\n';
    notification += '━━━━━━━━━━━━\n';
    notification += '👤 姓名：' + formData.name + '\n';
    notification += '📧 Email：' + formData.email + '\n';
    notification += '💬 訊息：' + formData.message + '\n';
    notification += '━━━━━━━━━━━━\n';
    notification += '📅 時間：' + now + '\n';
    notification += '📊 累計第 ' + rowCount + ' 筆';

    // 推播給個人
    pushText(LINE_USER_ID, notification);

    // 推播到群組（如果有設定 groupId）
    if (LINE_GROUP_ID && LINE_GROUP_ID.indexOf('C') === 0) {
      pushText(LINE_GROUP_ID, notification);
    }

    Logger.log('表單資料已記錄，已推播 Line 通知');

    return {
      success: true,
      message: '提交成功！感謝你的填寫。'
    };
  } catch (error) {
    Logger.log('processForm 錯誤：' + error.toString());
    return {
      success: false,
      message: '提交失敗：' + error.toString()
    };
  }
}

// ============================================================
// 第三部分：工具函式
// ============================================================

/**
 * 建立表單回覆工作表
 */
function createFormSheet(ss) {
  var sheet = ss.insertSheet(FORM_SHEET_NAME);

  var headers = ['時間戳記', '姓名', 'Email', '訊息內容'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#3F51B5')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 400);
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * 取得所有表單回覆（可選功能）
 */
function getAllSubmissions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(FORM_SHEET_NAME);

  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return data.map(function (row) {
    return {
      time: row[0],
      name: row[1],
      email: row[2],
      message: row[3]
    };
  });
}

// ========== Line Bot 推播函式（共用模組）==========

function pushText(to, text) {
  pushLine(to, [{ type: 'text', text: text }]);
}

function pushLine(to, messages) {
  var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({ to: to, messages: messages }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log('推播失敗，狀態碼：' + response.getResponseCode());
    Logger.log('錯誤內容：' + response.getContentText());
  }
}

// ========== 測試函式 ==========

/**
 * 測試表單處理（模擬一筆表單提交）
 */
function testFormSubmit() {
  var testData = {
    name: '測試同學',
    email: 'test@example.com',
    message: '這是一筆測試資料，用來確認系統運作正常。'
  };

  var result = processForm(testData);
  Logger.log('測試結果：' + JSON.stringify(result));
}
