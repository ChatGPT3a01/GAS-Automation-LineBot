/**
 * ============================================
 * 📝 自動化評量系統 — Line Bot 自動評語
 * ============================================
 *
 * 功能：學生透過 Line Bot 上傳 PDF 學習成果，
 *       系統自動提取文字、呼叫 AI 生成評語，
 *       回傳個人專屬評語並記錄到試算表。
 *
 * 使用技術：
 *   ✅ LINE Messaging API（Webhook 接收 + Push 推送）
 *   ✅ Google Drive API（儲存 PDF + 文字提取）
 *   ✅ Gemini / OpenAI API（AI 自動生成評語）
 *   ✅ Google Sheets（評量記錄存檔）
 *   ✅ PropertiesService（使用者狀態管理）
 *
 * 事前準備：
 *   1. 建立 LINE Bot → 取得 Channel Access Token
 *   2. 建立 Google Drive 資料夾 → 取得資料夾 ID
 *   3. 建立 Google 試算表 → 取得試算表 ID
 *   4. 取得 AI API Key（Gemini 免費 / OpenAI 付費）
 *   5. 在 GAS 編輯器 →「服務」→ 啟用 Drive API（v2）
 *   6. 部署為網路應用程式 → 設定 LINE Webhook URL
 *
 * ⚠️ 重要：必須在 GAS 啟用「Drive API」進階服務！
 *   步驟：編輯器左側「服務 +」→ 搜尋 Drive API → 啟用
 * ============================================
 */

// ============================================
// 第一部分：設定常數（請填入你自己的值）
// ============================================

const LINE_TOKEN     = '你的 LINE Channel Access Token';
const FOLDER_ID      = '你的 Google Drive 資料夾 ID';
const SPREADSHEET_ID = '你的 Google 試算表 ID';

// === AI 設定 ===
// 選擇 AI 提供者：'gemini'（推薦，免費）或 'openai'
const AI_PROVIDER = 'gemini';
const AI_API_KEY  = '你的 AI API Key';

// Gemini 模型選項：gemini-2.5-flash（推薦）、gemini-2.5-pro
const GEMINI_MODEL = 'gemini-2.5-flash';

// OpenAI 模型選項：gpt-5.2（推薦）、gpt-5.1
const OPENAI_MODEL = 'gpt-5.2';

// 使用者狀態存儲
const scriptProps = PropertiesService.getScriptProperties();

// ============================================
// 第二部分：Webhook 入口
// ============================================

/** LINE Webhook 接收點 */
function doPost(e) {
  try {
    const events = JSON.parse(e.postData.contents).events;

    for (let i = 0; i < events.length; i++) {
      handleEvent(events[i]);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error('doPost 錯誤:', error);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** Webhook 驗證用 */
function doGet(e) {
  return ContentService.createTextOutput('自動化評量系統運作中！');
}

// ============================================
// 第三部分：事件處理
// ============================================

function handleEvent(event) {
  const userId = event.source.userId;
  const replyToken = event.replyToken;

  // ─── 加入好友歡迎訊息 ───
  if (event.type === 'follow') {
    replyText(replyToken,
      '👋 歡迎使用「自動化評量系統」！\n\n' +
      '📝 使用方式：\n' +
      '輸入「幫我評語」即可開始\n\n' +
      '我會根據你的學習成果給予專屬評語！'
    );
    return;
  }

  // ─── 處理文字訊息 ───
  if (event.type === 'message' && event.message.type === 'text') {
    const text = event.message.text.trim();

    if (text === '幫我評語') {
      setUserState(userId, 'waiting_for_file');
      replyText(replyToken,
        '📄 請上傳你的自主學習成果 PDF 檔案\n\n' +
        '⚠️ 注意事項：\n' +
        '• 僅支援 PDF 格式\n' +
        '• 建議 10MB 以內\n\n' +
        '💡 檔案太大？可先到 iLovePDF 壓縮：\n' +
        'https://www.ilovepdf.com/zh-tw/compress_pdf\n\n' +
        '上傳後我會仔細閱讀並給你評語！'
      );
      return;
    }

    // 其他文字 → 提示使用方式
    replyText(replyToken,
      '🤖 我是自動化評量機器人\n\n請輸入「幫我評語」開始使用！'
    );
    return;
  }

  // ─── 處理檔案訊息 ───
  if (event.type === 'message' && event.message.type === 'file') {
    const state = getUserState(userId);

    if (state !== 'waiting_for_file') {
      replyText(replyToken, '請先輸入「幫我評語」再上傳檔案喔！');
      return;
    }

    const fileName = event.message.fileName;
    const messageId = event.message.id;

    // 檢查 PDF 格式
    if (!fileName.toLowerCase().endsWith('.pdf')) {
      replyText(replyToken, '⚠️ 請上傳 PDF 格式的檔案！');
      return;
    }

    // 回覆「處理中」
    replyText(replyToken, '⏳ 正在處理你的檔案，約需 30 秒到 1 分鐘...');

    // 清除狀態 & 處理檔案
    clearUserState(userId);
    processFile(userId, messageId, fileName);
  }
}

// ============================================
// 第四部分：檔案處理（核心流程）
// ============================================

function processFile(userId, messageId, fileName) {
  try {
    // Step 1：從 LINE 下載檔案
    const fileBlob = downloadFromLine(messageId);

    // Step 2：存到 Google Drive
    const driveFile = saveToDrive(fileBlob, fileName, userId);

    // Step 3：提取 PDF 文字
    const pdfText = extractPdfText(driveFile);
    if (!pdfText || pdfText.trim().length < 10) {
      pushText(userId,
        '⚠️ 無法讀取 PDF 內容\n\n' +
        '可能原因：掃描檔或圖片型 PDF\n' +
        '請確保 PDF 內含可選取的文字。'
      );
      return;
    }

    // Step 4：AI 生成評語
    const feedback = generateFeedback(pdfText);

    // Step 5：記錄到試算表
    logToSheet(userId, fileName, feedback);

    // Step 6：回傳評語
    pushText(userId,
      '✨ 你的自主學習評語 ✨\n\n' +
      feedback + '\n\n' +
      '───\n📌 繼續加油！歡迎再次分享學習成果！'
    );

  } catch (error) {
    console.error('processFile 錯誤:', error);
    pushText(userId, '❌ 處理時發生錯誤：' + error.message + '\n\n請稍後再試。');
  }
}

// ============================================
// 第五部分：LINE API 函式
// ============================================

/** 從 LINE 下載使用者上傳的檔案 */
function downloadFromLine(messageId) {
  const url = 'https://api-data.line.me/v2/bot/message/' + messageId + '/content';
  const res = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + LINE_TOKEN },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('LINE 檔案下載失敗');
  }
  return res.getBlob();
}

/** 回覆訊息（使用 replyToken） */
function replyText(replyToken, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{ type: 'text', text: text }]
    }),
    muteHttpExceptions: true
  });
}

/** 主動推送訊息（使用 userId） */
function pushText(userId, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'text', text: text }]
    }),
    muteHttpExceptions: true
  });
}

// ============================================
// 第六部分：Google Drive 函式
// ============================================

/** 儲存檔案到 Google Drive */
function saveToDrive(blob, fileName, userId) {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss');
  const newName = timestamp + '_' + userId.substring(0, 8) + '_' + fileName;

  blob.setName(newName);
  return folder.createFile(blob);
}

/**
 * 提取 PDF 文字內容
 * 原理：透過 Drive API 將 PDF 轉換為 Google Docs，再讀取文字
 * ⚠️ 需要啟用 Drive API 進階服務
 */
function extractPdfText(file) {
  try {
    const blob = file.getBlob();
    const tempName = 'temp_pdf_' + Date.now();

    // 將 PDF 轉換為 Google Docs
    const docFile = Drive.Files.insert(
      { title: tempName, mimeType: 'application/vnd.google-apps.document' },
      blob,
      { convert: true }
    );

    // 讀取轉換後的文字
    const doc = DocumentApp.openById(docFile.id);
    const text = doc.getBody().getText();

    // 清除暫存檔
    DriveApp.getFileById(docFile.id).setTrashed(true);

    console.log('PDF 提取成功，文字長度：' + text.length);
    return text;

  } catch (error) {
    console.error('PDF 提取失敗：' + error.message);
    return '';
  }
}

// ============================================
// 第七部分：AI 評語生成
// ============================================

/** 評語的 System Prompt */
const SYSTEM_PROMPT = [
  '你是一位溫暖且專業的教育評語專家。',
  '根據學生的自主學習成果，給予正向鼓勵的評語。',
  '',
  '評語要求：',
  '1. 字數：100-150 字',
  '2. 語氣：溫暖、正向、鼓勵',
  '3. 結構：肯定努力 → 指出優點 → 未來建議',
  '4. 使用繁體中文',
  '5. 避免制式化語句'
].join('\n');

/** 根據 PDF 內容生成評語 */
function generateFeedback(pdfContent) {
  // 截取內容避免超過 token 限制
  const maxLen = 8000;
  const content = pdfContent.length > maxLen
    ? pdfContent.substring(0, maxLen) + '...(內容已截取)'
    : pdfContent;

  const userPrompt =
    '請根據以下學生的自主學習成果，撰寫 100-150 字的正向評語：\n\n' +
    '===學生作品===\n' + content + '\n===============';

  if (AI_PROVIDER === 'gemini') {
    return callGemini(SYSTEM_PROMPT, userPrompt);
  } else {
    return callOpenAI(SYSTEM_PROMPT, userPrompt);
  }
}

/** 呼叫 Gemini API */
function callGemini(systemPrompt, userPrompt) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/'
    + GEMINI_MODEL + ':generateContent?key=' + AI_API_KEY;

  const payload = {
    system_instruction: { parts: [{ text: systemPrompt }] },
    contents: [{ parts: [{ text: userPrompt }] }],
    generationConfig: { temperature: 0.7, maxOutputTokens: 500 }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const result = JSON.parse(res.getContentText());

  if (result.error) {
    throw new Error('Gemini 錯誤：' + result.error.message);
  }

  if (!result.candidates || !result.candidates[0] ||
      !result.candidates[0].content || !result.candidates[0].content.parts) {
    throw new Error('Gemini 回覆格式異常');
  }

  return result.candidates[0].content.parts[0].text.trim();
}

/** 呼叫 OpenAI API */
function callOpenAI(systemPrompt, userPrompt) {
  const url = 'https://api.openai.com/v1/chat/completions';

  const payload = {
    model: OPENAI_MODEL,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    max_tokens: 500,
    temperature: 0.7
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + AI_API_KEY
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const result = JSON.parse(res.getContentText());

  if (result.error) {
    throw new Error('OpenAI 錯誤：' + result.error.message);
  }

  if (!result.choices || !result.choices[0] || !result.choices[0].message) {
    throw new Error('OpenAI 回覆格式異常');
  }

  return result.choices[0].message.content.trim();
}

// ============================================
// 第八部分：試算表記錄
// ============================================

function logToSheet(userId, fileName, feedback) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('評語記錄');

    // 若工作表不存在，建立並加標題
    if (!sheet) {
      sheet = ss.insertSheet('評語記錄');
      sheet.getRange(1, 1, 1, 5).setValues([
        ['時間', '使用者 ID', '檔案名稱', '評語', '狀態']
      ]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    const maskedId = userId.substring(0, 10) + '***';

    sheet.appendRow([timestamp, maskedId, fileName, feedback, '完成']);
  } catch (error) {
    console.error('試算表記錄錯誤:', error);
  }
}

// ============================================
// 第九部分：使用者狀態管理
// ============================================

function setUserState(userId, state) {
  scriptProps.setProperty('state_' + userId, state);
}

function getUserState(userId) {
  return scriptProps.getProperty('state_' + userId) || '';
}

function clearUserState(userId) {
  scriptProps.deleteProperty('state_' + userId);
}

// ============================================
// 第十部分：測試函式
// ============================================

/** 測試基本設定 */
function testBot() {
  console.log('=== 自動化評量系統測試 ===');
  console.log('LINE Token 長度:', LINE_TOKEN.length);
  console.log('AI Provider:', AI_PROVIDER);
  console.log('AI Model:', AI_PROVIDER === 'gemini' ? GEMINI_MODEL : OPENAI_MODEL);
  console.log('Drive Folder ID:', FOLDER_ID);
  console.log('Spreadsheet ID:', SPREADSHEET_ID);
  console.log('✅ 基本設定確認完成');
}

/** 測試 AI 評語生成 */
function testAI() {
  const testContent = '這份報告探討了再生能源在台灣的發展現況。' +
    '學生分析了太陽能和風力發電的效率，並提出校園節能方案。';

  console.log('測試 AI 評語生成...');
  console.log('AI Provider:', AI_PROVIDER);

  try {
    const feedback = generateFeedback(testContent);
    console.log('✅ 成功！');
    console.log('評語：\n' + feedback);
  } catch (error) {
    console.log('❌ 失敗：' + error.message);
  }
}

/** 測試 PDF 提取（從 Drive 資料夾取最新 PDF） */
function testPdfExtraction() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFilesByType('application/pdf');

  if (!files.hasNext()) {
    console.log('資料夾中沒有 PDF，請先上傳一個來測試');
    return;
  }

  const file = files.next();
  console.log('測試檔案：' + file.getName());

  const text = extractPdfText(file);
  if (text && text.length > 0) {
    console.log('✅ 提取成功！文字長度：' + text.length);
    console.log('前 300 字：\n' + text.substring(0, 300));
  } else {
    console.log('❌ 提取失敗');
  }
}
