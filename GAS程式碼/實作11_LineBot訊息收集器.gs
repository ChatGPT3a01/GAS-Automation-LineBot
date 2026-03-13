/**
 * ============================================================
 * 實作 2：Line Bot 訊息收集器
 * ============================================================
 *
 * 功能說明：
 *   1. 接收 Line Bot Webhook 事件（使用者傳送的所有訊息）
 *   2. 將訊息自動記錄到 Google 試算表
 *   3. 將圖片、影片、檔案自動存入 Google 雲端硬碟
 *   4. 收到訊息後自動回覆確認
 *
 * 試算表結構（自動記錄）：
 *   時間 | 群組ID | 群組名稱 | 用戶ID | 用戶名稱 | 訊息內容 | 檔案位置
 *
 * 支援的訊息類型：
 *   - 文字訊息：直接記錄文字
 *   - 貼圖：記錄 Package ID 和 Sticker ID
 *   - 位置：記錄經緯度和地址
 *   - 圖片/影片/音檔/檔案：下載到 Google Drive 並記錄連結
 *
 * 部署步驟：
 *   1. 建立一個 Google 試算表
 *   2. 在 Google Drive 建立一個資料夾（用來存檔案），複製資料夾 ID
 *   3. 開啟 Apps Script，貼上此程式碼
 *   4. 修改下方的四個設定值
 *   5. 部署為網頁應用程式（存取權限：所有人）
 *   6. 將部署網址貼到 Line Developers Console 的 Webhook URL
 *   7. 開啟「Use webhook」
 * ============================================================
 */

// ========== 設定區（請修改以下四個值）==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
var ROOT_FOLDER_ID = '在此貼上 Google Drive 資料夾 ID';
var SPREADSHEET_ID = '在此貼上 Google 試算表 ID';

// 不需要處理的事件類型
var IGNORE_SOURCE_TYPE = ['join', 'leave', 'follow', 'unfollow'];

// ============================================================
// 主要處理函式
// ============================================================

/**
 * 處理 Line Bot Webhook 的 POST 請求
 * 每當使用者傳訊息給 Bot，Line 平台就會呼叫此函式
 */
function doPost(e) {
  try {
    // 檢查 postData 是否存在
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'no postData' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var userData = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();

    // 檢查是否有標題列，沒有就建立
    if (sheet.getLastRow() === 0) {
      createHeaderRow(sheet);
    }

    // 逐一處理每個事件
    for (var i = 0; i < userData.events.length; i++) {
      var event = userData.events[i];

      // 跳過非訊息事件
      if (event.type !== 'message') {
        continue;
      }

      processMessage(event, sheet);
    }
  } catch (error) {
    Logger.log('doPost 錯誤：' + error.toString());
  }

  // 回傳 200 OK 給 Line 平台，避免 Webhook 重試
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 處理單一訊息事件
 */
function processMessage(event, sheet) {
  var now = new Date();
  var timestamp = formatTimestamp(now);

  // 取得來源資訊（個人/群組/聊天室）
  var sourceInfo = getSourceInfo(event);

  // 取得用戶資訊
  var userProfile = getUserProfile(sourceInfo.sourceType, sourceInfo.groupId, event.source.userId);
  var displayName = userProfile ? userProfile.displayName : '未知用戶';

  // 根據訊息類型處理
  var messageResult = processMessageType(event, sourceInfo, now);

  // 寫入試算表
  sheet.appendRow([
    now,                          // A 欄：時間
    sourceInfo.groupId,           // B 欄：群組 ID
    sourceInfo.groupName,         // C 欄：群組名稱
    event.source.userId,          // D 欄：用戶 ID
    displayName,                  // E 欄：用戶名稱
    messageResult.text,           // F 欄：訊息內容
    messageResult.fileLocation    // G 欄：檔案位置
  ]);

  // 回覆確認（可選：取消註解下面這行即可啟用自動回覆）
  // replyText(event.replyToken, '✅ 已記錄你的訊息！');

  Logger.log('已記錄：' + displayName + ' - ' + messageResult.text);
}

// ============================================================
// 訊息類型處理
// ============================================================

/**
 * 根據不同訊息類型進行處理
 * @returns {Object} { text: 訊息文字, fileLocation: 檔案連結 }
 */
function processMessageType(event, sourceInfo, now) {
  var text = '';
  var fileLocation = '';
  var message = event.message;

  switch (message.type) {
    // 文字訊息
    case 'text':
      text = message.text;
      break;

    // 貼圖
    case 'sticker':
      text = '【貼圖】';
      fileLocation = 'Package：' + message.packageId + '、Sticker：' + message.stickerId;
      break;

    // 位置資訊
    case 'location':
      text = '【位置】經度：' + message.longitude + '、緯度：' + message.latitude;
      if (message.title) { text += '、' + message.title; }
      if (message.address) { text += '、' + message.address; }
      break;

    // 圖片、影片、音檔、檔案 → 存到 Google Drive
    case 'image':
    case 'video':
    case 'audio':
    case 'file':
      text = '【' + getMessageTypeName(message.type) + '】';
      fileLocation = saveFileToDrive(event, sourceInfo, now);
      break;

    default:
      text = '【不支援的訊息類型：' + message.type + '】';
  }

  return { text: text, fileLocation: fileLocation };
}

/**
 * 取得訊息類型的中文名稱
 */
function getMessageTypeName(type) {
  var names = {
    'image': '圖片',
    'video': '影片',
    'audio': '音檔',
    'file': '檔案'
  };
  return names[type] || type;
}

// ============================================================
// 檔案儲存功能
// ============================================================

/**
 * 將 Line 傳送的檔案存到 Google Drive
 * @returns {string} 檔案的 Google Drive 連結
 */
function saveFileToDrive(event, sourceInfo, now) {
  try {
    var rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);

    // 決定子資料夾名稱（用群組ID或用戶ID）
    var folderName = sourceInfo.groupId || event.source.userId;
    var destinationFolder = getOrCreateSubfolder(rootFolder, folderName);

    // 從 Line 下載檔案
    var messageId = event.message.id;
    var fileBlob = getLineFileData(messageId).getBlob();

    // 建立檔案
    var file = destinationFolder.createFile(fileBlob);

    // 設定檔案名稱
    var prefix = formatFilePrefix(now);
    if (event.message.type === 'file' && event.message.fileName) {
      file.setName(prefix + event.message.fileName);
    } else {
      var ext = file.getName().split('.').pop();
      file.setName(prefix + messageId + '.' + ext);
    }

    return 'https://drive.google.com/open?id=' + file.getId();
  } catch (error) {
    Logger.log('檔案儲存失敗：' + error.toString());
    return '儲存失敗：' + error.toString();
  }
}

/**
 * 取得或建立子資料夾
 */
function getOrCreateSubfolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

/**
 * 從 Line 伺服器下載檔案內容
 */
function getLineFileData(messageId) {
  var url = 'https://api-data.line.me/v2/bot/message/' + messageId + '/content';
  return UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    method: 'get'
  });
}

// ============================================================
// Line API 輔助函式
// ============================================================

/**
 * 取得來源資訊（判斷是個人、群組、還是聊天室）
 */
function getSourceInfo(event) {
  var groupId = '';
  var groupName = '';
  var sourceType = event.source.type || '';

  if (sourceType === 'group') {
    groupId = event.source.groupId;
    var groupProfile = getGroupProfile(groupId);
    groupName = groupProfile ? groupProfile.groupName : '未知群組';
  } else if (sourceType === 'room') {
    groupId = event.source.roomId;
    groupName = '聊天室（無名稱）';
  }

  return {
    groupId: groupId,
    groupName: groupName,
    sourceType: sourceType
  };
}

/**
 * 取得 Line 用戶的個人資料
 */
function getUserProfile(sourceType, groupId, userId) {
  try {
    var url = '';
    switch (sourceType) {
      case 'group':
        url = 'https://api.line.me/v2/bot/group/' + groupId + '/member/' + userId;
        break;
      case 'room':
        url = 'https://api.line.me/v2/bot/room/' + groupId + '/member/' + userId;
        break;
      default:
        url = 'https://api.line.me/v2/bot/profile/' + userId;
    }

    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + LINE_TOKEN },
      method: 'get',
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }
    return null;
  } catch (error) {
    Logger.log('取得用戶資料失敗：' + error);
    return null;
  }
}

/**
 * 取得 Line 群組的基本資料
 */
function getGroupProfile(groupId) {
  try {
    var url = 'https://api.line.me/v2/bot/group/' + groupId + '/summary';
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + LINE_TOKEN },
      method: 'get',
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }
    return null;
  } catch (error) {
    Logger.log('取得群組資料失敗：' + error);
    return null;
  }
}

// ============================================================
// 工具函式
// ============================================================

/**
 * 建立試算表標題列
 */
function createHeaderRow(sheet) {
  var headers = ['時間', '群組ID', '群組名稱', '用戶ID', '用戶名稱', '訊息內容', '檔案位置'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1B5E20')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 400);
  sheet.setColumnWidth(7, 300);
}

/**
 * 格式化時間戳記（用於檔案名稱前綴）
 */
function formatFilePrefix(date) {
  var y = date.getFullYear();
  var m = padZero(date.getMonth() + 1);
  var d = padZero(date.getDate());
  var h = padZero(date.getHours());
  var min = padZero(date.getMinutes());
  var s = padZero(date.getSeconds());
  return y + m + d + h + min + s + '-';
}

/**
 * 格式化時間戳記（用於日誌）
 */
function formatTimestamp(date) {
  return date.toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
}

/**
 * 數字補零
 */
function padZero(num) {
  return num < 10 ? '0' + num : '' + num;
}

// ========== Line Bot 推播函式（共用模組）==========

function pushText(to, text) {
  pushLine(to, [{ type: 'text', text: text }]);
}

function pushImage(to, imageUrl) {
  pushLine(to, [{ type: 'image', originalContentUrl: imageUrl, previewImageUrl: imageUrl }]);
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

function replyText(replyToken, text) {
  replyLine(replyToken, [{ type: 'text', text: text }]);
}

function replyLine(replyToken, messages) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({ replyToken: replyToken, messages: messages }),
    muteHttpExceptions: true
  });
}

// ========== 測試函式 ==========

/**
 * 測試：推播一則訊息確認系統運作正常
 */
function testSystem() {
  pushText(LINE_USER_ID, '🔧 訊息收集器測試成功！\n\n系統已就緒，請在 Line 傳送訊息給 Bot 進行測試。');
  Logger.log('測試推播已發送');
}
