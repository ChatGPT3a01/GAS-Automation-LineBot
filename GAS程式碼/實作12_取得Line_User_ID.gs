/**
 * ============================================================
 * 實作 12：取得 Line User ID
 * ============================================================
 *
 * 功能說明：
 *   建立一個「ID 查詢 Bot」，傳任何訊息給 Bot，
 *   Bot 就自動回覆你的 userId 和 groupId。
 *   拿到 ID 後，就可以用在其他實作的推播功能。
 *
 * 使用方式：
 *   1. 把這段程式碼貼到 GAS 編輯器
 *   2. 填入你的 Line Bot Channel Access Token
 *   3. 部署為網頁應用程式
 *   4. 把部署網址設定為 Line Bot 的 Webhook URL
 *   5. 用手機傳一則訊息給 Bot，就會收到你的 ID
 *
 * ⚠️ 注意：部署後會覆蓋原本的 Webhook。
 *   取得 ID 後，記得把 Webhook 改回原本的網址。
 * ============================================================
 */

// ========== 設定區 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';

// ============================================================
// 第一部分：接收 Line 訊息，回覆 User ID
// ============================================================

/**
 * doPost — 接收 Line Bot Webhook
 * 使用者傳任何訊息，Bot 就回覆他的 userId
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var events = data.events;

    for (var i = 0; i < events.length; i++) {
      var event = events[i];

      if (event.type === 'message') {
        var replyToken = event.replyToken;
        var source = event.source;

        // 組合回覆訊息
        var reply = '📋 你的 Line ID 資訊\n';
        reply += '━━━━━━━━━━━━━━━\n\n';

        // userId（個人 ID）
        reply += '👤 你的 userId：\n';
        reply += source.userId + '\n\n';

        // 判斷是個人聊天還是群組
        if (source.type === 'group') {
          reply += '👥 群組 groupId：\n';
          reply += source.groupId + '\n\n';
          reply += '💡 推播給個人 → 用 userId\n';
          reply += '💡 推播到群組 → 用 groupId\n';
        } else if (source.type === 'room') {
          reply += '🏠 聊天室 roomId：\n';
          reply += source.roomId + '\n\n';
        } else {
          reply += '💡 這是個人聊天\n';
          reply += '💡 把上面的 userId 複製起來\n';
          reply += '💡 貼到其他實作的 LINE_USER_ID 設定中\n';
        }

        reply += '\n━━━━━━━━━━━━━━━\n';
        reply += '✅ 複製 ID 後，記得把 Webhook\n';
        reply += '改回原本的網址喔！';

        // 回覆訊息
        replyLine(replyToken, [{ type: 'text', text: reply }]);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('錯誤：' + error.message);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// 第二部分：Line Bot 回覆函式
// ============================================================

/**
 * 回覆訊息
 */
function replyLine(replyToken, messages) {
  var url = 'https://api.line.me/v2/bot/message/reply';
  var options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({ replyToken: replyToken, messages: messages }),
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(url, options);
}
