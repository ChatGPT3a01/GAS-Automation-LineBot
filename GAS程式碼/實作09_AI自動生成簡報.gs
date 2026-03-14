/**
 * ============================================================
 * 實作 9：AI 自動生成簡報（升級版）
 * ============================================================
 *
 * 功能說明：
 *   1. 從試算表 A1 讀取簡報主題
 *   2. 從 B1、B2、B3... 讀取關鍵字 / 重點
 *   3. 從 C1 讀取風格編號（1~6），或用側邊欄選擇
 *   4. 用 Gemini AI 根據主題 + 關鍵字生成大綱
 *   5. 用 SlidesApp 自動建立完整 Google 簡報
 *
 * ============ 使用方式 ============
 *
 * 【方法一】用擴充選單（推薦）
 *   1. 重新整理試算表 → 上方出現「🎨 AI 簡報」選單
 *   2. 點擊「開啟簡報生成器」→ 右側出現操作面板
 *   3. 選擇風格 → 點擊「生成簡報」
 *
 * 【方法二】直接執行
 *   1. 在試算表填好 A1 主題、B 欄關鍵字、C1 風格
 *   2. 在 GAS 編輯器執行 generatePresentation()
 *
 * ============ 試算表填寫方式 ============
 *   A1 = 簡報主題（例如：「人工智慧入門」）
 *   B1 = 關鍵字1（例如：「機器學習」）
 *   B2 = 關鍵字2（例如：「深度學習」）
 *   B3 = 關鍵字3（例如：「應用案例」）
 *   ...可繼續往下填，沒有上限
 *   C1 = 風格編號 1~6（選填，預設為 1）
 *
 * ============ 6 種簡報風格 ============
 *   1 = 📘 專業商務（深藍白底）
 *   2 = 🌿 清新自然（綠色系）
 *   3 = 🚀 科技未來（暗黑霓虹）
 *   4 = 🧡 溫暖教育（橘色系）
 *   5 = ⬛ 簡約黑白（極簡風）
 *   6 = 💜 活力繽紛（紫色系）
 *
 * ============ 設定區（只需改這裡）============
 *
 * Google AI Studio：https://aistudio.google.com/apikey
 * ============================================================
 */

var GEMINI_API_KEY = '在此貼上你的 Gemini API Key';
var GEMINI_MODEL = 'gemini-2.5-flash';

// ============================================================
// 6 種簡報風格定義
// ============================================================

var STYLES = {
  1: {
    name: '📘 專業商務',
    coverBg: '#1a237e', headerBg: '#283593',
    titleColor: '#ffffff', subtitleColor: '#90caf9',
    contentBg: '#f5f5f5', headingColor: '#1a237e',
    bulletColor: '#333333', accentBar: '#1565c0', endBg: '#1a237e'
  },
  2: {
    name: '🌿 清新自然',
    coverBg: '#1b5e20', headerBg: '#2e7d32',
    titleColor: '#ffffff', subtitleColor: '#c8e6c9',
    contentBg: '#f1f8e9', headingColor: '#1b5e20',
    bulletColor: '#333333', accentBar: '#43a047', endBg: '#1b5e20'
  },
  3: {
    name: '🚀 科技未來',
    coverBg: '#212121', headerBg: '#37474f',
    titleColor: '#00e5ff', subtitleColor: '#80deea',
    contentBg: '#263238', headingColor: '#00e5ff',
    bulletColor: '#eceff1', accentBar: '#00bcd4', endBg: '#212121'
  },
  4: {
    name: '🧡 溫暖教育',
    coverBg: '#e65100', headerBg: '#f57c00',
    titleColor: '#ffffff', subtitleColor: '#ffe0b2',
    contentBg: '#fff8e1', headingColor: '#e65100',
    bulletColor: '#333333', accentBar: '#ff9800', endBg: '#e65100'
  },
  5: {
    name: '⬛ 簡約黑白',
    coverBg: '#212121', headerBg: '#424242',
    titleColor: '#ffffff', subtitleColor: '#bdbdbd',
    contentBg: '#fafafa', headingColor: '#212121',
    bulletColor: '#424242', accentBar: '#757575', endBg: '#212121'
  },
  6: {
    name: '💜 活力繽紛',
    coverBg: '#6a1b9a', headerBg: '#8e24aa',
    titleColor: '#ffffff', subtitleColor: '#e1bee7',
    contentBg: '#f3e5f5', headingColor: '#6a1b9a',
    bulletColor: '#333333', accentBar: '#ab47bc', endBg: '#6a1b9a'
  }
};

// ============================================================
// 擴充選單（重新整理試算表後出現）
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎨 AI 簡報')
    .addItem('🚀 開啟簡報生成器', 'showSidebar')
    .addItem('▶ 直接生成簡報', 'generatePresentation')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('09_簡報風格選擇')
    .setTitle('🎨 AI 簡報生成器')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// 從試算表讀取資料
// ============================================================

function getInputFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // A1 = 主題
  var topic = sheet.getRange('A1').getValue().toString().trim();

  // B 欄 = 關鍵字列表
  var keywords = [];
  var lastRow = sheet.getLastRow();
  for (var i = 1; i <= lastRow; i++) {
    var val = sheet.getRange('B' + i).getValue().toString().trim();
    if (val) keywords.push(val);
  }

  // C1 = 風格編號（1~6）
  var styleNum = parseInt(sheet.getRange('C1').getValue()) || 1;
  if (styleNum < 1 || styleNum > 6) styleNum = 1;

  return { topic: topic, keywords: keywords, styleNum: styleNum };
}

// 給側邊欄用：取得目前試算表資料
function getSheetData() {
  return getInputFromSheet();
}

// ============================================================
// 呼叫 Gemini AI
// ============================================================

function callGemini(prompt) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/'
          + GEMINI_MODEL + ':generateContent?key=' + GEMINI_API_KEY;

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.7 }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();

  if (code !== 200) {
    throw new Error('Gemini API 錯誤（狀態碼：' + code + '）');
  }

  var json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error('Gemini API 錯誤：' + json.error.message);
  }

  if (!json.candidates || !json.candidates[0] ||
      !json.candidates[0].content || !json.candidates[0].content.parts) {
    throw new Error('Gemini 回覆格式異常');
  }

  return json.candidates[0].content.parts[0].text;
}

// ============================================================
// 建立 Prompt（含關鍵字）
// ============================================================

function buildPrompt(topic, keywords) {
  var prompt = '請幫我製作一份關於「' + topic + '」的簡報大綱。\n\n';

  if (keywords.length > 0) {
    prompt += '請特別包含以下重點或關鍵字：\n';
    for (var i = 0; i < keywords.length; i++) {
      prompt += '• ' + keywords[i] + '\n';
    }
    prompt += '\n';
  }

  prompt += '回覆格式必須是純 JSON（不要加 ```json 標記），結構如下：\n';
  prompt += '{\n';
  prompt += '  "title": "簡報標題",\n';
  prompt += '  "subtitle": "副標題",\n';
  prompt += '  "slides": [\n';
  prompt += '    {\n';
  prompt += '      "heading": "這一頁的標題",\n';
  prompt += '      "bullets": ["重點1", "重點2", "重點3"]\n';
  prompt += '    }\n';
  prompt += '  ],\n';
  prompt += '  "conclusion": "一句話結語"\n';
  prompt += '}\n\n';
  prompt += '請產生 5~8 頁內容頁，每頁 3~4 個重點，每個重點控制在 25 字以內。使用繁體中文。';

  return prompt;
}

// ============================================================
// 建立封面頁
// ============================================================

function createCoverSlide(presentation, title, subtitle, style) {
  var slide = presentation.getSlides()[0];

  // 設定背景色
  slide.getBackground().setSolidFill(style.coverBg);

  // 上方裝飾線
  var topLine = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 60, 90, 600, 3);
  topLine.getFill().setSolidFill(style.subtitleColor);
  topLine.getBorder().setTransparent();

  // 標題
  var titleBox = slide.insertTextBox(title, 60, 105, 600, 90);
  titleBox.getText().getTextStyle()
    .setFontSize(42).setBold(true).setForegroundColor(style.titleColor);
  titleBox.getText().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // 副標題
  var subBox = slide.insertTextBox(subtitle, 100, 220, 520, 60);
  subBox.getText().getTextStyle()
    .setFontSize(20).setForegroundColor(style.subtitleColor);
  subBox.getText().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // 日期
  var dateText = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy / MM / dd');
  var dateBox = slide.insertTextBox(dateText, 250, 310, 220, 30);
  dateBox.getText().getTextStyle()
    .setFontSize(12).setForegroundColor(style.subtitleColor);
  dateBox.getText().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // 底部裝飾線
  var bottomLine = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 250, 305, 220, 2);
  bottomLine.getFill().setSolidFill(style.subtitleColor);
  bottomLine.getBorder().setTransparent();
}

// ============================================================
// 建立內容頁（迴圈產生多頁）
// ============================================================

function createContentSlides(presentation, slidesData, style) {
  for (var i = 0; i < slidesData.length; i++) {
    var data = slidesData[i];
    var slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

    // 設定內容頁背景
    slide.getBackground().setSolidFill(style.contentBg);

    // 上方標題橫條
    var headerBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 720, 60);
    headerBg.getFill().setSolidFill(style.headerBg);
    headerBg.getBorder().setTransparent();

    // 標題文字
    var headerText = slide.insertTextBox(data.heading, 20, 10, 680, 45);
    headerText.getText().getTextStyle()
      .setFontSize(24).setBold(true).setForegroundColor('#ffffff');

    // 左側裝飾條
    var accentBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 30, 80, 5, 280);
    accentBar.getFill().setSolidFill(style.accentBar);
    accentBar.getBorder().setTransparent();

    // 重點內容
    var bulletText = '';
    var bullets = data.bullets || data.points || [];
    for (var j = 0; j < bullets.length; j++) {
      bulletText += '● ' + bullets[j] + '\n\n';
    }

    var contentBox = slide.insertTextBox(bulletText, 50, 80, 630, 290);
    contentBox.getText().getTextStyle()
      .setFontSize(18).setForegroundColor(style.bulletColor);

    // 頁碼
    var pageNum = slide.insertTextBox((i + 1) + ' / ' + slidesData.length, 620, 370, 80, 25);
    pageNum.getText().getTextStyle()
      .setFontSize(11).setForegroundColor(style.accentBar);
    pageNum.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.RIGHT);
  }
}

// ============================================================
// 建立結尾頁
// ============================================================

function createEndSlide(presentation, conclusion, style) {
  var slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(style.endBg);

  // 結語
  var conclusionBox = slide.insertTextBox(conclusion || '感謝聆聽！', 60, 100, 600, 80);
  conclusionBox.getText().getTextStyle()
    .setFontSize(28).setBold(true).setForegroundColor(style.titleColor);
  conclusionBox.getText().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Thank You
  var thankBox = slide.insertTextBox('Thank You', 60, 200, 600, 70);
  thankBox.getText().getTextStyle()
    .setFontSize(48).setBold(true).setForegroundColor(style.titleColor);
  thankBox.getText().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // 底部資訊
  var infoBox = slide.insertTextBox('Powered by AI + Google Apps Script', 100, 320, 520, 30);
  infoBox.getText().getTextStyle()
    .setFontSize(14).setForegroundColor(style.subtitleColor);
  infoBox.getText().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================================
// 主函式：生成簡報
// ============================================================

function generatePresentation(styleOverride) {
  // 讀取試算表資料
  var input = getInputFromSheet();

  if (!input.topic) {
    SpreadsheetApp.getUi().alert('❌ 請在 A1 欄位輸入簡報主題！');
    return { success: false, message: '請在 A1 輸入主題' };
  }

  var styleNum = styleOverride || input.styleNum;
  var style = STYLES[styleNum] || STYLES[1];

  Logger.log('📝 主題：' + input.topic);
  Logger.log('🔑 關鍵字：' + (input.keywords.length > 0 ? input.keywords.join('、') : '無'));
  Logger.log('🎨 風格：' + style.name);

  // Step 1：用 AI 生成大綱
  Logger.log('正在請 AI 生成大綱...');
  var prompt = buildPrompt(input.topic, input.keywords);
  var aiReply = callGemini(prompt);

  // Step 2：清理 AI 回覆（移除可能的 markdown 標記）
  aiReply = aiReply.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();

  // Step 3：解析 JSON
  var outline = JSON.parse(aiReply);
  Logger.log('大綱標題：' + outline.title);
  Logger.log('共 ' + outline.slides.length + ' 頁內容');

  // Step 4：建立簡報
  var presentation = SlidesApp.create(outline.title);

  // Step 5：建立封面頁
  createCoverSlide(presentation, outline.title, outline.subtitle, style);

  // Step 6：建立內容頁
  createContentSlides(presentation, outline.slides, style);

  // Step 7：建立結尾頁
  createEndSlide(presentation, outline.conclusion, style);

  // 取得簡報連結
  var url = presentation.getUrl();
  var totalPages = outline.slides.length + 2;

  // 寫入試算表記錄
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow() + 2;
  sheet.getRange(lastRow, 1).setValue('📊 簡報標題').setFontWeight('bold');
  sheet.getRange(lastRow, 2).setValue(outline.title);
  sheet.getRange(lastRow + 1, 1).setValue('📎 簡報連結').setFontWeight('bold');
  sheet.getRange(lastRow + 1, 2).setValue(url);
  sheet.getRange(lastRow + 2, 1).setValue('📄 頁數').setFontWeight('bold');
  sheet.getRange(lastRow + 2, 2).setValue(totalPages + ' 頁');
  sheet.getRange(lastRow + 3, 1).setValue('🎨 風格').setFontWeight('bold');
  sheet.getRange(lastRow + 3, 2).setValue(style.name);

  Logger.log('========================================');
  Logger.log('✅ 簡報生成完成！');
  Logger.log('標題：' + outline.title);
  Logger.log('頁數：' + totalPages + ' 頁');
  Logger.log('風格：' + style.name);
  Logger.log('📎 連結：' + url);
  Logger.log('========================================');

  return { success: true, url: url, title: outline.title, pages: totalPages };
}

// ============================================================
// 測試函式
// ============================================================

function testGenerate() {
  var result = generatePresentation();
  if (result && result.success) {
    SpreadsheetApp.getUi().alert(
      '✅ 簡報生成完成！\n\n' +
      '標題：' + result.title + '\n' +
      '頁數：' + result.pages + ' 頁\n\n' +
      '📎 連結：' + result.url
    );
  }
}
