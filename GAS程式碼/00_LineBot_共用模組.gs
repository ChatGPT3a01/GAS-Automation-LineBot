/**
 * ============================================================
 * 00_LineBot 共用推播模組
 * ============================================================
 *
 * 用途：所有實作都會使用的 Line Bot 推播功能
 *
 * 使用方式：
 *   方法 A（推薦）：在 Apps Script 編輯器中，點選「+」新增檔案，
 *            命名為「LineBot」，將此程式碼貼入。
 *   方法 B：直接貼在主程式的最下方。
 *
 * 修改步驟：
 *   1. 將 LINE_TOKEN 換成你的 Channel Access Token
 *   2. 將 LINE_USER_ID 換成你的 User ID
 *
 * Line Developers Console 網址：https://developers.line.biz/
 * ============================================================
 */

// ========== 請修改以下兩個值 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';

// ========== 推播函式 ==========

/**
 * 推播文字訊息給指定對象
 * @param {string} to - 接收者的 userId 或 groupId
 * @param {string} text - 文字內容
 */
function pushText(to, text) {
  pushLine(to, [{ type: 'text', text: text }]);
}

/**
 * 推播圖片訊息給指定對象
 * @param {string} to - 接收者的 userId 或 groupId
 * @param {string} imageUrl - 圖片的公開網址（必須是 https 開頭）
 */
function pushImage(to, imageUrl) {
  pushLine(to, [{
    type: 'image',
    originalContentUrl: imageUrl,
    previewImageUrl: imageUrl
  }]);
}

/**
 * 推播多則訊息（底層函式）
 * Line Bot 每次最多推播 5 則訊息
 * @param {string} to - 接收者的 userId 或 groupId
 * @param {Array} messages - 訊息陣列，每個元素是一個訊息物件
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

// ========== 回覆函式（用於 Webhook 場景）==========

/**
 * 回覆文字訊息（當 Line Bot 收到使用者訊息時使用）
 * @param {string} replyToken - Line 平台提供的回覆用 token（每次事件不同）
 * @param {string} text - 回覆的文字內容
 */
function replyText(replyToken, text) {
  replyLine(replyToken, [{ type: 'text', text: text }]);
}

/**
 * 回覆多則訊息（底層函式）
 * @param {string} replyToken - Line 平台提供的回覆用 token
 * @param {Array} messages - 訊息陣列
 */
function replyLine(replyToken, messages) {
  var url = 'https://api.line.me/v2/bot/message/reply';
  var options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: messages
    }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();

  if (code !== 200) {
    Logger.log('回覆失敗，狀態碼：' + code);
    Logger.log('錯誤內容：' + response.getContentText());
  }
}

// ========== 測試函式 ==========

/**
 * 測試推播 — 在 Apps Script 編輯器中直接執行此函式
 * 執行後你的 Line 應該會收到一則測試訊息
 */
function testPush() {
  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  pushText(LINE_USER_ID, '🎉 Line Bot 連線測試成功！\n\n目前時間：' + now + '\n\n如果你看到這則訊息，表示 GAS 與 Line Bot 已成功串接。');
}
