/**
 * ============================================================
 * 實作 8：每天來張早安圖 → Line Bot 圖片推播
 * ============================================================
 *
 * 功能說明：
 *   1. 從 Google Drive 指定資料夾中讀取所有圖片
 *   2. 隨機挑選一張圖片
 *   3. 將圖片設為「任何人皆可檢視」（產生公開連結）
 *   4. 透過 Line Bot pushImage() 推播早安圖
 *   5. 設定每天早上自動發送
 *
 * 事前準備：
 *   1. 在 Google Drive 建立一個「早安圖」資料夾
 *   2. 上傳 3~5 張早安圖到該資料夾（支援 jpg、png、gif）
 *   3. 複製該資料夾的 ID（網址列 folders/ 後面那串）
 *   4. 取得 Line Bot 的 Channel Access Token
 *   5. 取得你的 Line User ID
 *
 * 設定排程自動執行：
 *   1. 在 Apps Script 編輯器中，點選左側「觸發條件」（鬧鐘圖示）
 *   2. 點選「新增觸發條件」
 *   3. 選擇函式：sendMorningImage
 *   4. 事件來源：時間驅動
 *   5. 時間型觸發條件類型：每日計時器
 *   6. 時段：上午 7 點到 8 點
 *   7. 儲存
 * ============================================================
 */

// ========== Line Bot 設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
var LINE_GROUP_ID = '在此貼上你的 Group ID（不需要群組通知就留空白）';

// ========== Google Drive 設定 ==========
// 在 Google Drive 建立一個資料夾，把早安圖放進去
// 資料夾 ID 就是網址中 folders/ 後面的那串英數字
var MORNING_FOLDER_ID = '在此貼上你的早安圖資料夾 ID';

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
 * 推播圖片訊息
 * 注意：圖片網址必須是 https 開頭，且可公開存取
 */
function pushImage(to, imageUrl) {
  pushLine(to, [{
    type: 'image',
    originalContentUrl: imageUrl,
    previewImageUrl: imageUrl
  }]);
}

/**
 * 推播底層函式
 */
/**
 * 同時推播給個人和群組（如果有設定 GROUP_ID）
 */
function pushToAll(text) {
  pushText(LINE_USER_ID, text);
  if (LINE_GROUP_ID && LINE_GROUP_ID.indexOf('C') === 0) {
    pushText(LINE_GROUP_ID, text);
  }
}

function pushImageToAll(imageUrl) {
  pushImage(LINE_USER_ID, imageUrl);
  if (LINE_GROUP_ID && LINE_GROUP_ID.indexOf('C') === 0) {
    pushImage(LINE_GROUP_ID, imageUrl);
  }
}

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
// 第二部分：Google Drive 圖片操作
// ============================================================

/**
 * 從指定資料夾取得所有圖片檔案
 *
 * @param {string} folderId - Google Drive 資料夾 ID
 * @returns {File[]} 圖片檔案陣列
 */
function getImagesFromFolder(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var images = [];

  while (files.hasNext()) {
    var file = files.next();
    var mimeType = file.getMimeType();

    // 只取圖片檔案（jpg、png、gif、webp）
    if (mimeType.indexOf('image/') === 0) {
      images.push(file);
      Logger.log('找到圖片：' + file.getName() + '（' + mimeType + '）');
    }
  }

  Logger.log('共找到 ' + images.length + ' 張圖片');
  return images;
}

/**
 * 隨機挑選一張圖片
 *
 * @param {File[]} images - 圖片檔案陣列
 * @returns {File} 隨機選中的圖片
 */
function pickRandomImage(images) {
  var index = Math.floor(Math.random() * images.length);
  var selected = images[index];
  Logger.log('隨機選中：' + selected.getName());
  return selected;
}

/**
 * 取得圖片的公開連結
 * 會自動將圖片設為「任何人皆可檢視」
 *
 * @param {File} file - Google Drive 檔案物件
 * @returns {string} 可直接存取的圖片網址
 */
function getPublicImageUrl(file) {
  // 設定為「任何人皆可檢視」
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 取得檔案 ID
  var fileId = file.getId();

  // 組合直接存取的圖片網址
  // 這個網址格式可以讓 Line Bot 直接顯示圖片
  var imageUrl = 'https://lh3.googleusercontent.com/d/' + fileId;

  Logger.log('圖片公開網址：' + imageUrl);
  return imageUrl;
}

// ============================================================
// 第三部分：主要功能 — 發送早安圖
// ============================================================

/**
 * 【主函式】發送早安圖到 Line
 * 排程觸發時執行這個函式
 */
function sendMorningImage() {
  Logger.log('===== 開始發送早安圖 =====');

  // 步驟 1：取得資料夾中的所有圖片
  var images = getImagesFromFolder(MORNING_FOLDER_ID);

  if (images.length === 0) {
    Logger.log('資料夾中沒有圖片！');
    pushToAll('⚠️ 早安圖資料夾中沒有圖片，請上傳一些圖片。');
    return;
  }

  // 步驟 2：隨機挑選一張
  var selectedImage = pickRandomImage(images);

  // 步驟 3：取得公開連結
  var imageUrl = getPublicImageUrl(selectedImage);

  // 步驟 4：推播圖片到 Line（個人 + 群組）
  pushImageToAll(imageUrl);

  // 步驟 5：附帶早安文字訊息
  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd (EEE)');
  var greetings = [
    '🌅 早安！新的一天，新的開始！',
    '☀️ 早安！今天也要元氣滿滿喔！',
    '🌸 早安！美好的一天從微笑開始！',
    '🌈 早安！願你今天一切順利！',
    '💪 早安！今天也要加油喔！'
  ];
  var greeting = greetings[Math.floor(Math.random() * greetings.length)];

  pushToAll(greeting + '\n\n📅 ' + dateStr + '\n📸 今日早安圖：' + selectedImage.getName());

  Logger.log('===== 早安圖發送完成 =====');
}

// ============================================================
// 第四部分：測試函式
// ============================================================

/**
 * 測試：列出資料夾中的所有圖片
 */
function testListImages() {
  var images = getImagesFromFolder(MORNING_FOLDER_ID);
  for (var i = 0; i < images.length; i++) {
    Logger.log((i + 1) + '. ' + images[i].getName() + ' (' + images[i].getMimeType() + ')');
  }
}

/**
 * 測試：發送一張早安圖
 */
function testSendMorningImage() {
  sendMorningImage();
}

/**
 * 測試：推播文字
 */
function testPush() {
  pushToAll('🧪 早安圖 Line Bot 連線測試成功！');
}
