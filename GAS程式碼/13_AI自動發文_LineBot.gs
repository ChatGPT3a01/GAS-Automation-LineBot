/**
 * ============================================================
 * 實作 13：AI 自動發文 → Line Bot AI 推播
 * ============================================================
 *
 * 功能說明：
 *   1. 串接 Gemini API（Google AI）生成文章
 *   2. 設計 Prompt 自動生成指定主題的社群貼文
 *   3. AI 回覆寫入試算表（含標題、內文、標籤）
 *   4. 透過 Line Bot 推播 AI 生成的貼文
 *   5. 設定每天自動產生一篇文章並推播
 *
 * 事前準備：
 *   1. 到 Google AI Studio 申請 Gemini API Key
 *      網址：https://aistudio.google.com/apikey
 *   2. 取得 Line Bot 的 Channel Access Token
 *   3. 取得你的 Line User ID
 *   4. 建立 Google 試算表（用於記錄 AI 生成的文章）
 *   5. 取得試算表 ID
 *
 * 設定排程自動執行：
 *   1. 在 Apps Script 編輯器中，點選左側「觸發條件」（鬧鐘圖示）
 *   2. 點選「新增觸發條件」
 *   3. 選擇函式：dailyAutoPost
 *   4. 事件來源：時間驅動
 *   5. 時間型觸發條件類型：每日計時器
 *   6. 時段：上午 9 點到 10 點
 *   7. 儲存
 * ============================================================
 */

// ========== Line Bot 設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';

// ========== AI 設定 ==========
// 到 https://aistudio.google.com/apikey 申請
var GEMINI_API_KEY = '在此貼上你的 Gemini API Key';

// Gemini 模型選項：
// gemini-2.5-flash（推薦，速度快、免費額度高）
// gemini-2.5-pro（更強，免費額度較低）
var GEMINI_MODEL = 'gemini-2.5-flash';

// ========== 試算表設定 ==========
var SPREADSHEET_ID = '在此貼上你的試算表 ID';
var SHEET_NAME = 'AI 發文記錄';

// ========== 發文主題設定 ==========
// 你可以自行修改這些主題
var TOPICS = [
  '科技新趨勢',
  '生活小技巧',
  '健康養生知識',
  '職場溝通技巧',
  '教育創新方法',
  '環保永續行動',
  '時間管理術'
];

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
// 第二部分：Gemini AI API 串接
// ============================================================

/**
 * 呼叫 Gemini API 生成文字
 *
 * @param {string} prompt - 提示詞
 * @returns {string} AI 生成的回覆文字
 */
function callGemini(prompt) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/'
    + GEMINI_MODEL + ':generateContent?key=' + GEMINI_API_KEY;

  var payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.8,      // 創意度（0~1，越高越有創意）
      maxOutputTokens: 1024  // 最大回覆長度
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code !== 200) {
      Logger.log('Gemini API 呼叫失敗，狀態碼：' + code);
      Logger.log('錯誤內容：' + response.getContentText());
      return null;
    }

    var json = JSON.parse(response.getContentText());

    // 取得 AI 回覆的文字
    if (json.candidates && json.candidates[0] &&
        json.candidates[0].content && json.candidates[0].content.parts) {
      var text = json.candidates[0].content.parts[0].text;
      Logger.log('Gemini 回覆：' + text.substring(0, 100) + '...');
      return text;
    }

    Logger.log('Gemini 回覆格式異常：' + JSON.stringify(json));
    return null;

  } catch (error) {
    Logger.log('Gemini API 錯誤：' + error.message);
    return null;
  }
}

// ============================================================
// 第三部分：Prompt 設計（提示詞工程）
// ============================================================

/**
 * 產生社群貼文的 Prompt
 *
 * @param {string} topic - 文章主題
 * @returns {string} 完整的提示詞
 */
function buildPrompt(topic) {
  var prompt = '你是一位專業的社群內容創作者。\n\n';
  prompt += '請幫我撰寫一篇關於「' + topic + '」的社群貼文。\n\n';
  prompt += '要求：\n';
  prompt += '1. 標題：吸引人、簡潔有力（15字以內）\n';
  prompt += '2. 內文：200~300字，口語化、易讀\n';
  prompt += '3. 包含 2~3 個實用的重點或建議\n';
  prompt += '4. 結尾加上 3~5 個相關的 hashtag\n';
  prompt += '5. 適當使用 emoji 讓文章更生動\n\n';
  prompt += '請用以下格式回覆：\n';
  prompt += '【標題】你的標題\n';
  prompt += '【內文】你的內文\n';
  prompt += '【標籤】#tag1 #tag2 #tag3';

  return prompt;
}

/**
 * 解析 AI 回覆，提取標題、內文、標籤
 *
 * @param {string} response - AI 的回覆文字
 * @returns {Object} { title, content, tags }
 */
function parseResponse(response) {
  var result = {
    title: '',
    content: '',
    tags: ''
  };

  // 提取標題
  var titleMatch = response.match(/【標題】(.+?)(\n|【)/);
  if (titleMatch) {
    result.title = titleMatch[1].trim();
  }

  // 提取內文
  var contentMatch = response.match(/【內文】([\s\S]+?)【標籤】/);
  if (contentMatch) {
    result.content = contentMatch[1].trim();
  }

  // 提取標籤
  var tagsMatch = response.match(/【標籤】(.+)/);
  if (tagsMatch) {
    result.tags = tagsMatch[1].trim();
  }

  // 如果格式解析失敗，用整篇文字作為內文
  if (!result.title && !result.content) {
    result.title = '每日分享';
    result.content = response;
    result.tags = '';
  }

  return result;
}

// ============================================================
// 第四部分：試算表記錄
// ============================================================

/**
 * 初始化記錄表
 */
function initRecordSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var headers = ['日期', '主題', '標題', '內文', '標籤', '狀態'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1b5e20')
      .setFontColor('white');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 將 AI 生成的文章記錄到試算表
 *
 * @param {string} topic - 主題
 * @param {Object} article - { title, content, tags }
 */
function saveToSheet(topic, article) {
  var sheet = initRecordSheet();
  var now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm');

  sheet.appendRow([
    now,
    topic,
    article.title,
    article.content,
    article.tags,
    '已推播'
  ]);

  Logger.log('文章已記錄到試算表');
}

// ============================================================
// 第五部分：主要功能 — AI 自動發文
// ============================================================

/**
 * 產生一篇 AI 文章並推播到 Line
 *
 * @param {string} topic - 指定主題（選填，不填則隨機）
 */
function generateAndPost(topic) {
  Logger.log('===== 開始 AI 自動發文 =====');

  // 如果沒指定主題，隨機選一個
  if (!topic) {
    topic = TOPICS[Math.floor(Math.random() * TOPICS.length)];
  }
  Logger.log('主題：' + topic);

  // 步驟 1：建立 Prompt
  var prompt = buildPrompt(topic);
  Logger.log('Prompt 已建立');

  // 步驟 2：呼叫 Gemini API
  var aiResponse = callGemini(prompt);

  if (!aiResponse) {
    pushText(LINE_USER_ID, '⚠️ AI 自動發文失敗：無法取得 AI 回覆，請檢查 API Key。');
    return;
  }

  // 步驟 3：解析回覆
  var article = parseResponse(aiResponse);
  Logger.log('標題：' + article.title);

  // 步驟 4：記錄到試算表
  saveToSheet(topic, article);

  // 步驟 5：推播到 Line
  var message = '🤖 AI 每日好文\n';
  message += '━━━━━━━━━━━━━━━\n\n';
  message += '📌 ' + article.title + '\n\n';
  message += article.content + '\n\n';
  if (article.tags) {
    message += article.tags + '\n\n';
  }
  message += '━━━━━━━━━━━━━━━\n';
  message += '📝 主題：' + topic + '\n';
  message += '💡 由 Gemini AI + GAS 自動產生';

  // Line 單則訊息上限 5000 字
  if (message.length > 4900) {
    message = message.substring(0, 4900) + '\n\n...（已截斷）';
  }

  pushText(LINE_USER_ID, message);

  Logger.log('===== AI 自動發文完成 =====');
}

/**
 * 【排程函式】每天自動產生一篇文章並推播
 */
function dailyAutoPost() {
  generateAndPost();  // 隨機主題
}

// ============================================================
// 第六部分：測試函式
// ============================================================

/**
 * 測試：呼叫 Gemini API
 */
function testGemini() {
  var response = callGemini('用一句話介紹 Google Apps Script');
  Logger.log('AI 回覆：' + response);
}

/**
 * 測試：產生並推播一篇 AI 文章
 */
function testGenerateAndPost() {
  generateAndPost('科技新趨勢');
}

/**
 * 測試：指定主題產生文章
 */
function testCustomTopic() {
  generateAndPost('教育創新方法');
}

/**
 * 測試：推播文字
 */
function testPush() {
  pushText(LINE_USER_ID, '🧪 AI 自動發文 Line Bot 連線測試成功！');
}
