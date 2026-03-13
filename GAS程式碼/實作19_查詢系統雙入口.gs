/**
 * ============================================================
 * 實作 12：GAS 查詢系統 → Line Bot 互動查詢
 * ============================================================
 *
 * 功能說明：
 *   1. 建立 HTML 查詢介面（搜尋框 + 結果區）
 *   2. GAS 後端實作搜尋邏輯（模糊搜尋、篩選）
 *   3. 前端用 google.script.run 非同步取得結果
 *   4. 部署為公開 Web App 查詢系統
 *   5. 同步建立 Line Bot 查詢入口（雙入口系統）
 *
 * 事前準備：
 *   1. 建立資料試算表（至少有：姓名、部門、Email 等欄位）
 *   2. 取得試算表 ID
 *   3. 取得 Line Bot 的 Channel Access Token
 *   4. 取得你的 Line User ID
 *   5. 部署後設定 Line Bot Webhook URL
 *
 * 檔案結構（需在 Apps Script 中建立兩個檔案）：
 *   - Code.gs：此檔案（後端程式）
 *   - 查詢介面.html：前端網頁
 *     ⚠️ 在 Apps Script 中必須命名為「查詢介面」（不含 .html）
 *     內容請參考下載包中的 實作19_查詢系統雙入口_前端.html
 *
 * 部署步驟：
 *   1. 點選右上角「部署」→「新增部署作業」
 *   2. 類型選擇「網頁應用程式」
 *   3. 執行身分：我
 *   4. 存取權：所有人
 *   5. 部署
 *   6. 複製網址 → 到 Line Developers Console 設定 Webhook URL
 * ============================================================
 */

// ========== Line Bot 設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';

// ========== 試算表設定 ==========
var SPREADSHEET_ID = '在此貼上你的試算表 ID';
var SHEET_NAME = '人員資料';

// ============================================================
// 第一部分：Web App 入口
// ============================================================

/**
 * doGet — 顯示查詢網頁介面
 * 當使用者用瀏覽器打開 Web App 時執行
 */
function doGet(e) {
  // 如果有查詢參數，回傳 JSON（API 模式）
  if (e && e.parameter && e.parameter.action === 'search') {
    var keyword = e.parameter.q || '';
    var results = searchData(keyword);
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      count: results.length,
      data: results
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 否則顯示 HTML 查詢頁面
  return HtmlService.createHtmlOutputFromFile('查詢介面')
    .setTitle('資料查詢系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * doPost — 接收 Line Bot Webhook
 * 使用者在 Line 傳訊息時執行
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var events = data.events;

    for (var i = 0; i < events.length; i++) {
      var event = events[i];

      if (event.type === 'message' && event.message.type === 'text') {
        var userMessage = event.message.text.trim();
        var replyToken = event.replyToken;

        // 判斷是否為查詢指令
        if (userMessage.indexOf('查') === 0 || userMessage.indexOf('找') === 0) {
          // 去掉「查」或「找」字
          var keyword = userMessage.substring(1).trim();
          handleLineSearch(replyToken, keyword);
        } else if (userMessage === '使用說明' || userMessage === '幫助') {
          replyText(replyToken,
            '📖 查詢系統使用說明\n\n' +
            '傳送以下格式即可查詢：\n' +
            '• 查 王小明\n' +
            '• 找 業務部\n' +
            '• 查 test@email\n\n' +
            '支援模糊搜尋，輸入部分關鍵字即可！'
          );
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('doPost 錯誤：' + error.message);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// 第二部分：搜尋邏輯
// ============================================================

/**
 * 搜尋試算表資料（模糊搜尋）
 *
 * @param {string} keyword - 搜尋關鍵字
 * @returns {Object[]} 符合條件的資料陣列
 */
function searchData(keyword) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log('找不到工作表：' + SHEET_NAME);
    return [];
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];  // 只有標題列或空的

  var headers = data[0];  // 第一列為標題
  var results = [];
  var searchKey = keyword.toLowerCase();

  // 從第二列開始搜尋
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var match = false;

    // 檢查每一欄是否包含關鍵字
    for (var j = 0; j < row.length; j++) {
      var cellValue = String(row[j]).toLowerCase();
      if (cellValue.indexOf(searchKey) !== -1) {
        match = true;
        break;
      }
    }

    if (match) {
      var record = {};
      for (var k = 0; k < headers.length; k++) {
        record[headers[k]] = row[k];
      }
      results.push(record);
    }
  }

  Logger.log('搜尋「' + keyword + '」找到 ' + results.length + ' 筆資料');
  return results;
}

// ============================================================
// 第三部分：Line Bot 查詢回覆
// ============================================================

/**
 * 處理 Line Bot 的查詢請求
 *
 * @param {string} replyToken - Line 回覆用 Token
 * @param {string} keyword - 搜尋關鍵字
 */
function handleLineSearch(replyToken, keyword) {
  if (!keyword) {
    replyText(replyToken, '❓ 請輸入要查詢的關鍵字\n\n範例：查 王小明');
    return;
  }

  var results = searchData(keyword);

  if (results.length === 0) {
    replyText(replyToken, '🔍 查詢「' + keyword + '」\n\n找不到符合的資料。\n請嘗試其他關鍵字。');
    return;
  }

  // 組合回覆訊息
  var message = '🔍 查詢「' + keyword + '」\n';
  message += '📊 找到 ' + results.length + ' 筆資料\n';
  message += '━━━━━━━━━━━━━━━\n';

  // 最多顯示 5 筆（避免訊息太長）
  var showCount = Math.min(results.length, 5);

  for (var i = 0; i < showCount; i++) {
    var record = results[i];
    var keys = Object.keys(record);
    message += '\n📋 第 ' + (i + 1) + ' 筆：\n';

    for (var j = 0; j < keys.length; j++) {
      message += '  • ' + keys[j] + '：' + record[keys[j]] + '\n';
    }
  }

  if (results.length > showCount) {
    message += '\n...還有 ' + (results.length - showCount) + ' 筆，請縮小搜尋範圍。';
  }

  replyText(replyToken, message);
}

// ============================================================
// 第四部分：Line Bot 推播與回覆函式
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
    payload: JSON.stringify({ to: to, messages: messages }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    Logger.log('推播失敗：' + response.getContentText());
  }
}

/**
 * 回覆文字訊息（用於 Webhook）
 */
function replyText(replyToken, text) {
  replyLine(replyToken, [{ type: 'text', text: text }]);
}

/**
 * 回覆底層函式
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

// ============================================================
// 第五部分：試算表初始化
// ============================================================

/**
 * 建立範例資料（首次使用時執行）
 */
function initSampleData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    sheet.clear();
  }

  // 標題列
  var headers = ['姓名', '部門', '職稱', 'Email', '分機'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1b5e20')
    .setFontColor('white');

  // 範例資料
  var sampleData = [
    ['王小明', '業務部', '業務經理', 'wang@example.com', '101'],
    ['李大華', '技術部', '工程師', 'lee@example.com', '201'],
    ['張美美', '人資部', '人資專員', 'chang@example.com', '301'],
    ['陳志明', '業務部', '業務專員', 'chen@example.com', '102'],
    ['林小芬', '技術部', '前端工程師', 'lin@example.com', '202'],
    ['黃大偉', '行政部', '行政主管', 'huang@example.com', '401'],
    ['趙小燕', '技術部', '後端工程師', 'chao@example.com', '203'],
    ['周美玲', '人資部', '人資主管', 'chou@example.com', '302']
  ];

  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  sheet.setFrozenRows(1);

  // 自動調整欄寬
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  Logger.log('已建立 ' + sampleData.length + ' 筆範例資料');
  pushText(LINE_USER_ID, '✅ 查詢系統範例資料已建立（' + sampleData.length + ' 筆）');
}

// ============================================================
// 第六部分：測試函式
// ============================================================

/**
 * 測試：搜尋功能
 */
function testSearch() {
  var results = searchData('業務');
  Logger.log('找到 ' + results.length + ' 筆');
  for (var i = 0; i < results.length; i++) {
    Logger.log(JSON.stringify(results[i]));
  }
}

/**
 * 測試：推播文字
 */
function testPush() {
  pushText(LINE_USER_ID, '🧪 查詢系統 Line Bot 連線測試成功！\n\n在 Line 中傳「查 業務」試試看。');
}
