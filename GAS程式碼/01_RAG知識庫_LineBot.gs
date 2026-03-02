/**
 * ============================================================
 * 實作 1：RAG 知識庫 API + Line Bot 查詢回覆
 * ============================================================
 *
 * 功能說明：
 *   1. 提供 REST API（用瀏覽器或程式存取知識庫）
 *   2. 整合 Line Bot Webhook（在 Line 中輸入關鍵字即可查詢）
 *   3. 支援知識庫的 CRUD 操作（新增、查詢、更新、刪除）
 *
 * 試算表結構（自動建立）：
 *   工作表「知識庫」：問題 | 答案 | 分類 | 更新時間 | 來源
 *   工作表「查詢記錄」：查詢時間 | 查詢內容 | 結果數量
 *
 * 使用方式：
 *   A. REST API：部署後用瀏覽器存取
 *      - 列出全部：https://你的網址/exec?action=list
 *      - 搜尋：https://你的網址/exec?action=search&query=關鍵字
 *      - 統計：https://你的網址/exec?action=stats
 *   B. Line Bot：在 Line 中傳送訊息給 Bot
 *      - 輸入任意文字 → Bot 自動搜尋知識庫並回覆
 *
 * 部署步驟：
 *   1. 開啟 Google 試算表 → 擴充功能 → Apps Script
 *   2. 貼上此程式碼
 *   3. 修改下方的 LINE_TOKEN 和 LINE_USER_ID
 *   4. 點選「部署」→「新增部署作業」
 *   5. 類型選「網頁應用程式」
 *   6. 存取權限選「所有人」
 *   7. 部署後取得網址
 *   8. 將網址貼到 Line Developers Console 的 Webhook URL
 * ============================================================
 */

// ========== Line Bot 設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';

// ========== 知識庫設定 ==========
var SHEET_NAME = '知識庫';
var LOG_SHEET_NAME = '查詢記錄';

// ============================================================
// 第一部分：處理 GET 請求（REST API 查詢用）
// ============================================================

/**
 * 處理 GET 請求 — 用瀏覽器存取時觸發
 * 支援三種 action：list（列出全部）、search（搜尋）、stats（統計）
 */
function doGet(e) {
  try {
    var params = e.parameter;
    var action = params.action || 'list';

    switch (action) {
      case 'list':
        return getKnowledgeBase(params);
      case 'search':
        return searchKnowledgeBase(params);
      case 'stats':
        return getStatistics();
      default:
        return createResponse(false, '不支援的操作：' + action, null);
    }
  } catch (error) {
    return createResponse(false, '發生錯誤：' + error.toString(), null);
  }
}

// ============================================================
// 第二部分：處理 POST 請求（REST API 寫入 + Line Bot Webhook）
// ============================================================

/**
 * 處理 POST 請求
 * 會自動判斷是 Line Bot Webhook 還是 REST API 呼叫
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // 判斷是否為 Line Bot Webhook 事件
    if (data.events) {
      return handleLineWebhook(data);
    }

    // 否則當作 REST API 處理
    var action = data.action;
    switch (action) {
      case 'add':
        return addKnowledge(data);
      case 'update':
        return updateKnowledge(data);
      case 'delete':
        return deleteKnowledge(data);
      case 'batch_add':
        return batchAddKnowledge(data);
      default:
        return createResponse(false, '不支援的操作：' + action, null);
    }
  } catch (error) {
    return createResponse(false, '發生錯誤：' + error.toString(), null);
  }
}

// ============================================================
// 第三部分：Line Bot Webhook 處理
// ============================================================

/**
 * 處理 Line Bot Webhook 事件
 * 使用者在 Line 傳訊息 → 搜尋知識庫 → 回覆結果
 */
function handleLineWebhook(data) {
  for (var i = 0; i < data.events.length; i++) {
    var event = data.events[i];

    // 只處理文字訊息
    if (event.type !== 'message' || event.message.type !== 'text') {
      continue;
    }

    var userMessage = event.message.text.trim();
    var replyToken = event.replyToken;

    // 搜尋知識庫
    var results = searchKnowledgeData(userMessage);

    // 組合回覆訊息
    var replyMessage = '';
    if (results.length === 0) {
      replyMessage = '🔍 找不到與「' + userMessage + '」相關的知識。\n\n請試試其他關鍵字！';
    } else {
      replyMessage = '🔍 查詢「' + userMessage + '」的結果：\n\n';
      // 最多顯示 3 筆
      var showCount = Math.min(results.length, 3);
      for (var j = 0; j < showCount; j++) {
        replyMessage += '📌 ' + results[j].question + '\n';
        replyMessage += '💡 ' + results[j].answer + '\n\n';
      }
      if (results.length > 3) {
        replyMessage += '（還有 ' + (results.length - 3) + ' 筆結果未顯示）';
      }
    }

    // 回覆使用者
    replyText(replyToken, replyMessage);

    // 記錄查詢
    logQuery(userMessage, results.length);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 第四部分：知識庫 CRUD 操作
// ============================================================

/**
 * 取得知識庫列表
 */
function getKnowledgeBase(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = createKnowledgeBaseSheet(ss);
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);

  var knowledgeBase = rows.map(function (row, index) {
    var obj = { rowIndex: index + 2 };
    headers.forEach(function (header, i) {
      obj[header] = row[i];
    });
    return obj;
  });

  // 支援分類篩選
  var filtered = knowledgeBase;
  if (params.category) {
    filtered = filtered.filter(function (item) {
      return item['分類'] === params.category;
    });
  }

  return createResponse(true, '取得成功', {
    data: filtered,
    count: filtered.length,
    timestamp: new Date().toISOString()
  });
}

/**
 * 搜尋知識庫（REST API 版本，回傳 JSON）
 */
function searchKnowledgeBase(params) {
  var query = params.query || '';
  var results = searchKnowledgeData(query);

  logQuery(query, results.length);

  return createResponse(true, '搜尋成功', {
    results: results,
    count: results.length,
    query: query
  });
}

/**
 * 搜尋知識庫核心邏輯（內部使用，回傳陣列）
 */
function searchKnowledgeData(query) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  var searchQuery = query.toLowerCase();

  var results = [];
  for (var i = 0; i < rows.length; i++) {
    var question = String(rows[i][0] || '').toLowerCase();
    var answer = String(rows[i][1] || '').toLowerCase();
    var category = String(rows[i][2] || '').toLowerCase();

    if (question.indexOf(searchQuery) !== -1 ||
        answer.indexOf(searchQuery) !== -1 ||
        category.indexOf(searchQuery) !== -1) {
      results.push({
        rowIndex: i + 2,
        question: rows[i][0],
        answer: rows[i][1],
        category: rows[i][2],
        updateTime: rows[i][3],
        source: rows[i][4]
      });
    }
  }

  return results;
}

/**
 * 新增一筆知識
 */
function addKnowledge(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = createKnowledgeBaseSheet(ss);
  }

  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  var newRow = [
    data.question || '',
    data.answer || '',
    data.category || '一般',
    now,
    data.source || '手動新增'
  ];

  sheet.appendRow(newRow);

  // 推播通知
  pushText(LINE_USER_ID, '📚 知識庫新增成功！\n\n問題：' + newRow[0] + '\n分類：' + newRow[2]);

  return createResponse(true, '新增成功', { row: newRow });
}

/**
 * 更新一筆知識
 */
function updateKnowledge(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    return createResponse(false, '知識庫工作表不存在', null);
  }

  var rowIndex = data.rowIndex;
  if (!rowIndex || rowIndex < 2) {
    return createResponse(false, '無效的列索引', null);
  }

  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  var range = sheet.getRange(rowIndex, 1, 1, 5);
  range.setValues([[
    data.question || '',
    data.answer || '',
    data.category || '一般',
    now,
    '已更新'
  ]]);

  return createResponse(true, '更新成功', null);
}

/**
 * 刪除一筆知識
 */
function deleteKnowledge(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    return createResponse(false, '知識庫工作表不存在', null);
  }

  var rowIndex = data.rowIndex;
  if (!rowIndex || rowIndex < 2) {
    return createResponse(false, '無效的列索引', null);
  }

  sheet.deleteRow(rowIndex);

  return createResponse(true, '刪除成功', null);
}

/**
 * 批次新增知識
 */
function batchAddKnowledge(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = createKnowledgeBaseSheet(ss);
  }

  var items = data.items || [];
  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

  var newRows = items.map(function (item) {
    return [
      item.question || '',
      item.answer || '',
      item.category || '一般',
      now,
      '批次匯入'
    ];
  });

  if (newRows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, 5).setValues(newRows);
  }

  // 推播通知
  pushText(LINE_USER_ID, '📚 批次匯入完成！\n共新增 ' + newRows.length + ' 筆知識。');

  return createResponse(true, '批次新增成功', { addedCount: newRows.length });
}

// ============================================================
// 第五部分：統計與記錄
// ============================================================

/**
 * 取得統計資料
 */
function getStatistics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    return createResponse(true, '取得成功', {
      totalCount: 0,
      categories: {},
      lastUpdated: null
    });
  }

  var data = sheet.getDataRange().getValues();
  var rows = data.slice(1);

  var categories = {};
  for (var i = 0; i < rows.length; i++) {
    var category = rows[i][2] || '一般';
    categories[category] = (categories[category] || 0) + 1;
  }

  var lastUpdated = null;
  if (rows.length > 0) {
    lastUpdated = rows[rows.length - 1][3];
  }

  return createResponse(true, '取得成功', {
    totalCount: rows.length,
    categories: categories,
    lastUpdated: lastUpdated,
    timestamp: new Date().toISOString()
  });
}

/**
 * 記錄查詢歷史
 */
function logQuery(query, resultCount) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(LOG_SHEET_NAME);

    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET_NAME);
      var headers = ['查詢時間', '查詢內容', '結果數量'];
      logSheet.getRange(1, 1, 1, 3).setValues([headers]);
      logSheet.getRange(1, 1, 1, 3)
        .setFontWeight('bold')
        .setBackground('#2196F3')
        .setFontColor('#FFFFFF');
      logSheet.setFrozenRows(1);
    }

    var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
    logSheet.appendRow([now, query, resultCount]);
  } catch (error) {
    Logger.log('記錄查詢失敗：' + error);
  }
}

// ============================================================
// 第六部分：工具函式
// ============================================================

/**
 * 建立知識庫工作表（首次使用自動建立）
 */
function createKnowledgeBaseSheet(ss) {
  var sheet = ss.insertSheet(SHEET_NAME);

  // 設定標題列
  var headers = ['問題', '答案', '分類', '更新時間', '來源'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 格式化標題列
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4CAF50');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  // 設定欄寬
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 500);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 100);

  // 凍結標題列
  sheet.setFrozenRows(1);

  // 新增範例資料
  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  var examples = [
    ['什麼是 Google Apps Script？', 'Google Apps Script（GAS）是 Google 提供的雲端腳本語言，以 JavaScript 為基礎，可以自動化 Google 試算表、文件、表單等服務的操作。', '程式設計', now, '範例資料'],
    ['什麼是 Line Bot？', 'Line Bot 是 Line 平台上的聊天機器人，透過 Messaging API 可以接收使用者訊息並自動回覆，也可以主動推播訊息給使用者。', '通訊應用', now, '範例資料'],
    ['如何串接外部 API？', '在 GAS 中使用 UrlFetchApp.fetch(url, options) 即可呼叫外部 API。options 可設定 method、headers、payload 等參數。回傳結果用 JSON.parse() 解析。', '程式設計', now, '範例資料'],
    ['什麼是 Webhook？', 'Webhook 是一種伺服器間的即時通知機制。當事件發生時（例如使用者傳訊息），平台會主動發送 HTTP POST 請求到你指定的 URL，讓你的程式即時處理。', '網路技術', now, '範例資料'],
    ['GAS 有什麼限制？', 'GAS 免費版主要限制：每次執行最長 6 分鐘、每天 UrlFetchApp 呼叫上限 20,000 次、每天寄信上限 100 封、觸發條件執行總時長每天 90 分鐘。', '程式設計', now, '範例資料']
  ];

  sheet.getRange(2, 1, examples.length, 5).setValues(examples);

  return sheet;
}

/**
 * 建立標準 JSON 回應格式
 */
function createResponse(success, message, data) {
  var response = {
    success: success,
    message: message,
    data: data,
    timestamp: new Date().toISOString()
  };

  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== Line Bot 推播函式（共用模組）==========

function pushText(to, text) {
  pushLine(to, [{ type: 'text', text: text }]);
}

function pushImage(to, imageUrl) {
  pushLine(to, [{ type: 'image', originalContentUrl: imageUrl, previewImageUrl: imageUrl }]);
}

function pushLine(to, messages) {
  var url = 'https://api.line.me/v2/bot/message/push';
  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({ to: to, messages: messages }),
    muteHttpExceptions: true
  });
}

function replyText(replyToken, text) {
  replyLine(replyToken, [{ type: 'text', text: text }]);
}

function replyLine(replyToken, messages) {
  var url = 'https://api.line.me/v2/bot/message/reply';
  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({ replyToken: replyToken, messages: messages }),
    muteHttpExceptions: true
  });
}

// ========== 手動測試函式 ==========

/**
 * 在編輯器中執行此函式，測試知識庫是否正常運作
 */
function testKnowledgeBase() {
  // 測試搜尋
  var results = searchKnowledgeData('GAS');
  Logger.log('搜尋 "GAS" 找到 ' + results.length + ' 筆結果');
  for (var i = 0; i < results.length; i++) {
    Logger.log('  - ' + results[i].question);
  }

  // 推播測試結果到 Line
  if (results.length > 0) {
    pushText(LINE_USER_ID, '📚 知識庫測試成功！\n搜尋 "GAS" 找到 ' + results.length + ' 筆結果。\n\n第一筆：' + results[0].question);
  } else {
    pushText(LINE_USER_ID, '📚 知識庫測試：搜尋 "GAS" 沒有找到結果，請先新增範例資料。');
  }
}
