/**
 * ============================================================
 * 實作 11：自動化活動報名系統 → Line Bot 即時通知
 * ============================================================
 *
 * 功能說明：
 *   1. 使用 Google 表單收集報名資料
 *   2. 每次有人報名 → Line Bot 即時推播通知管理員
 *   3. 自動統計報名人數
 *   4. 達到人數上限時，自動關閉表單 + Line Bot 通知
 *   5. 所有報名記錄自動寫入試算表
 *
 * 事前準備：
 *   1. 建立 Google 表單（活動報名表）
 *   2. 取得表單 ID（網址列 /d/ 和 /edit 之間的那串）
 *   3. 表單必須連結到 Google 試算表（回覆 → 試算表）
 *   4. 取得 Line Bot 的 Channel Access Token
 *   5. 取得你的 Line User ID
 *
 * 設定表單提交觸發條件：
 *   1. 在 Apps Script 編輯器中，點選左側「觸發條件」（鬧鐘圖示）
 *   2. 點選「新增觸發條件」
 *   3. 選擇函式：onFormSubmit
 *   4. 事件來源：試算表
 *   5. 事件類型：提交表單時
 *   6. 儲存
 * ============================================================
 */

// ========== Line Bot 設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';

// ========== 報名系統設定 ==========
// Google 表單 ID（網址中 /d/ 和 /edit 之間的那串）
var FORM_ID = '在此貼上你的 Google 表單 ID';

// 報名人數上限（設為 0 表示不限制）
var MAX_PARTICIPANTS = 30;

// 活動名稱
var EVENT_NAME = 'GAS 自動化工作坊';

// ============================================================
// 第一部分：Line Bot 推播函式
// ============================================================

/**
 * 推播文字訊息
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
    Logger.log('推播成功！');
  }
}

// ============================================================
// 第二部分：表單提交觸發（核心功能）
// ============================================================

/**
 * 【核心函式】表單提交時自動執行
 * 需要設定觸發條件：試算表 → 提交表單時
 *
 * @param {Object} e - 表單提交事件物件
 */
function onFormSubmit(e) {
  try {
    Logger.log('===== 收到新報名 =====');

    // 取得提交的資料
    var responses = e.namedValues;  // { '姓名': ['王小明'], 'Email': ['test@gmail.com'], ... }
    var timestamp = e.namedValues['時間戳記'] ? e.namedValues['時間戳記'][0] : new Date().toLocaleString('zh-TW');

    // 組合報名資訊
    var info = '';
    var keys = Object.keys(responses);
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (key !== '時間戳記') {
        info += '  • ' + key + '：' + responses[key][0] + '\n';
      }
    }

    // 取得目前報名人數
    var currentCount = getCurrentCount();

    // 組合通知訊息
    var message = '🎉 收到新報名！\n';
    message += '━━━━━━━━━━━━━━━\n';
    message += '📌 活動：' + EVENT_NAME + '\n';
    message += '🕐 時間：' + timestamp + '\n\n';
    message += '📝 報名資料：\n' + info + '\n';
    message += '📊 目前報名人數：' + currentCount;

    if (MAX_PARTICIPANTS > 0) {
      message += ' / ' + MAX_PARTICIPANTS;
      var remaining = MAX_PARTICIPANTS - currentCount;
      message += '\n📋 剩餘名額：' + remaining + ' 位';
    }

    // 推播通知管理員
    pushText(LINE_USER_ID, message);

    // 檢查是否額滿
    if (MAX_PARTICIPANTS > 0 && currentCount >= MAX_PARTICIPANTS) {
      closeForm();
    }

    Logger.log('===== 報名處理完成 =====');

  } catch (error) {
    Logger.log('onFormSubmit 錯誤：' + error.message);
    pushText(LINE_USER_ID, '⚠️ 報名系統發生錯誤：' + error.message);
  }
}

// ============================================================
// 第三部分：報名人數統計
// ============================================================

/**
 * 取得目前報名人數
 *
 * @returns {number} 目前報名人數
 */
function getCurrentCount() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();
  // 扣掉標題列
  var count = lastRow > 0 ? lastRow - 1 : 0;
  Logger.log('目前報名人數：' + count);
  return count;
}

// ============================================================
// 第四部分：表單控制
// ============================================================

/**
 * 關閉表單（報名額滿時呼叫）
 */
function closeForm() {
  try {
    var form = FormApp.openById(FORM_ID);
    form.setAcceptingResponses(false);
    form.setCustomClosedFormMessage(
      '感謝您的關注！「' + EVENT_NAME + '」報名已額滿，期待下次活動見面！'
    );

    Logger.log('表單已關閉（額滿）');

    // 推播額滿通知
    var message = '🔔 報名額滿通知\n';
    message += '━━━━━━━━━━━━━━━\n';
    message += '📌 活動：' + EVENT_NAME + '\n';
    message += '👥 已報名：' + MAX_PARTICIPANTS + ' / ' + MAX_PARTICIPANTS + '\n\n';
    message += '✅ 報名表單已自動關閉！';

    pushText(LINE_USER_ID, message);

  } catch (error) {
    Logger.log('關閉表單失敗：' + error.message);
  }
}

/**
 * 重新開啟表單（手動操作）
 */
function openForm() {
  var form = FormApp.openById(FORM_ID);
  form.setAcceptingResponses(true);
  Logger.log('表單已重新開啟');
  pushText(LINE_USER_ID, '✅ 「' + EVENT_NAME + '」報名表單已重新開啟！');
}

// ============================================================
// 第五部分：報名統計報告
// ============================================================

/**
 * 產生報名統計報告並推播
 */
function sendRegistrationReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();
  var count = lastRow > 0 ? lastRow - 1 : 0;

  var now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm');

  var message = '📊 報名統計報告\n';
  message += '━━━━━━━━━━━━━━━\n';
  message += '📌 活動：' + EVENT_NAME + '\n';
  message += '🕐 統計時間：' + now + '\n\n';
  message += '👥 已報名人數：' + count;

  if (MAX_PARTICIPANTS > 0) {
    message += ' / ' + MAX_PARTICIPANTS + '\n';
    message += '📋 剩餘名額：' + (MAX_PARTICIPANTS - count) + ' 位\n';
    message += '📈 報名率：' + Math.round(count / MAX_PARTICIPANTS * 100) + '%';
  }

  pushText(LINE_USER_ID, message);
}

// ============================================================
// 第六部分：快速建立報名表單
// ============================================================

/**
 * 快速建立一個活動報名表單
 * 執行後會在 Google Drive 中建立新表單
 */
function createRegistrationForm() {
  var form = FormApp.create(EVENT_NAME + ' — 報名表');

  // 必填欄位
  form.addTextItem().setTitle('姓名').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  form.addTextItem().setTitle('聯絡電話').setRequired(true);
  form.addMultipleChoiceItem()
    .setTitle('用餐需求')
    .setChoiceValues(['葷食', '素食', '不需要'])
    .setRequired(true);
  form.addParagraphTextItem().setTitle('備註（選填）');

  // 設定表單
  form.setConfirmationMessage('報名成功！我們會盡快與您聯繫。');
  form.setCollectEmail(false);

  // 建立回覆試算表
  form.setDestination(FormApp.DestinationType.SPREADSHEET,
    SpreadsheetApp.getActiveSpreadsheet().getId());

  Logger.log('表單已建立！');
  Logger.log('表單網址：' + form.getPublishedUrl());
  Logger.log('表單編輯：' + form.getEditUrl());
  Logger.log('表單 ID：' + form.getId());

  // 推播表單連結
  var message = '✅ 報名表單已建立！\n';
  message += '━━━━━━━━━━━━━━━\n';
  message += '📌 活動：' + EVENT_NAME + '\n';
  message += '📝 表單連結：\n' + form.getPublishedUrl() + '\n\n';
  message += '⚠️ 請記得更新程式碼中的 FORM_ID：\n' + form.getId();

  pushText(LINE_USER_ID, message);
}

// ============================================================
// 第七部分：測試函式
// ============================================================

/**
 * 測試：查看目前報名人數
 */
function testCurrentCount() {
  var count = getCurrentCount();
  Logger.log('目前報名人數：' + count);
}

/**
 * 測試：發送統計報告
 */
function testSendReport() {
  sendRegistrationReport();
}

/**
 * 測試：推播文字
 */
function testPush() {
  pushText(LINE_USER_ID, '🧪 活動報名系統 Line Bot 連線測試成功！');
}
