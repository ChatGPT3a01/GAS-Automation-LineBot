/**
 * ============================================================
 * 實作 6：Line Bot 建立與推播測試
 * ============================================================
 *
 * 功能說明：
 *   1. 認識 Line Bot 的運作原理
 *   2. 設定 Channel Access Token 和 User ID
 *   3. 學會推播文字訊息和圖片訊息
 *   4. 建立可重用的 Line Bot 共用模組
 *
 * 使用方式：
 *   1. 到 Line Developers Console 建立 Messaging API Channel
 *   2. 取得 Channel Access Token 和 User ID
 *   3. 填入下方設定區
 *   4. 執行 testPush() 測試推播
 *
 * Line Developers Console 網址：https://developers.line.biz/
 * ============================================================
 */

// ========== 請修改以下設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
var LINE_GROUP_ID = '在此貼上你的 Group ID（不需要群組通知就留空白）';

// ============================================================
// 第一部分：推播函式（共用模組）
// ============================================================

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

// ============================================================
// 第二部分：回覆函式（用於 Webhook 場景）
// ============================================================

/**
 * 回覆文字訊息（當 Line Bot 收到使用者訊息時使用）
 * @param {string} replyToken - Line 平台提供的回覆用 token
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

// ============================================================
// 第三部分：測試函式
// ============================================================

/**
 * 測試推播 — 在 Apps Script 編輯器中直接執行此函式
 * 執行後你的 Line 應該會收到一則測試訊息
 */
function testPush() {
  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  var msg = '🎉 Line Bot 連線測試成功！\n\n目前時間：' + now + '\n\n如果你看到這則訊息，表示 GAS 與 Line Bot 已成功串接。';
  pushText(LINE_USER_ID, msg);
  if (LINE_GROUP_ID && LINE_GROUP_ID.indexOf('C') === 0) {
    pushText(LINE_GROUP_ID, msg);
  }
}

/**
 * 測試推播圖片
 */
function testPushImage() {
  // 使用 Google 的範例圖片
  var imageUrl = 'https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_272x92dp.png';
  pushImage(LINE_USER_ID, imageUrl);
  if (LINE_GROUP_ID && LINE_GROUP_ID.indexOf('C') === 0) {
    pushImage(LINE_GROUP_ID, imageUrl);
  }
  Logger.log('圖片推播測試完成！');
}

/**
 * 測試推播多則訊息
 */
function testMultiMessage() {
  var messages = [
    { type: 'text', text: '第 1 則：這是文字訊息' },
    { type: 'text', text: '第 2 則：Line Bot 一次最多推播 5 則' },
    { type: 'text', text: '第 3 則：恭喜你成功了！🎉' }
  ];
  pushLine(LINE_USER_ID, messages);
  if (LINE_GROUP_ID && LINE_GROUP_ID.indexOf('C') === 0) {
    pushLine(LINE_GROUP_ID, messages);
  }
}
