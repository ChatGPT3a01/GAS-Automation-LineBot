/**
 * ============================================================
 * 實作 9：AI 自動生成簡報
 * ============================================================
 *
 * 功能說明：
 *   1. 用 Gemini AI 根據主題生成簡報大綱（標題 + 要點）
 *   2. 用 SlidesApp 建立新的 Google 簡報
 *   3. 自動產生封面頁、內容頁、結語頁
 *   4. 設定字體、顏色、版面配置
 *   5. 簡報連結寫入試算表
 *
 * 設定步驟：
 *   1. 到 Google AI Studio 取得 Gemini API Key
 *   2. 將 API Key 填入下方 GEMINI_API_KEY
 *   3. 執行 generatePresentation('你想要的主題')
 *
 * Google AI Studio：https://aistudio.google.com/apikey
 * ============================================================
 */

// ========== 設定區 ==========
var GEMINI_API_KEY = '在此貼上你的 Gemini API Key';
var GEMINI_MODEL = 'gemini-2.5-flash';

// ============================================================
// 第一部分：呼叫 Gemini AI 生成大綱
// ============================================================

/**
 * 呼叫 Gemini AI 生成簡報大綱
 * @param {string} topic - 簡報主題
 * @return {Object} 包含 title, slides 陣列的結構化資料
 */
function generateOutline(topic) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + GEMINI_API_KEY;

  var prompt = '請根據以下主題生成一份簡報大綱，回傳 JSON 格式：\n\n';
  prompt += '主題：' + topic + '\n\n';
  prompt += '請回傳以下 JSON 格式（不要加 markdown 標記）：\n';
  prompt += '{\n';
  prompt += '  "title": "簡報標題",\n';
  prompt += '  "subtitle": "副標題",\n';
  prompt += '  "slides": [\n';
  prompt += '    { "heading": "第一頁標題", "points": ["要點1", "要點2", "要點3"] },\n';
  prompt += '    { "heading": "第二頁標題", "points": ["要點1", "要點2", "要點3"] }\n';
  prompt += '  ],\n';
  prompt += '  "conclusion": "結語"\n';
  prompt += '}\n\n';
  prompt += '請生成 4-6 頁內容頁。每頁 3 個要點。使用繁體中文。';

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

  try {
    var json = JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log('API 回應解析失敗：' + e.message);
    return null;
  }

  if (json.error) {
    Logger.log('Gemini API 錯誤：' + json.error.message);
    return null;
  }

  // 檢查 candidates 是否存在
  if (!json.candidates || !json.candidates[0] ||
      !json.candidates[0].content || !json.candidates[0].content.parts) {
    Logger.log('Gemini 回覆格式異常：' + JSON.stringify(json).substring(0, 200));
    return null;
  }

  var text = json.candidates[0].content.parts[0].text;

  // 移除可能的 markdown 標記
  text = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();

  try {
    return JSON.parse(text);
  } catch (e) {
    Logger.log('AI 回傳的 JSON 格式錯誤：' + e.message);
    Logger.log('原始回覆：' + text.substring(0, 300));
    return null;
  }
}

// ============================================================
// 第二部分：用 SlidesApp 建立簡報
// ============================================================

/**
 * 建立簡報的主函式
 * @param {string} topic - 簡報主題（預設為「AI 在教育的應用」）
 */
function generatePresentation(topic) {
  topic = topic || 'AI 在教育領域的應用';

  Logger.log('正在用 AI 生成大綱...');
  var outline = generateOutline(topic);

  if (!outline) {
    Logger.log('大綱生成失敗，請檢查 API Key');
    return;
  }

  Logger.log('大綱生成成功，正在建立簡報...');

  // 建立新簡報
  var presentation = SlidesApp.create(outline.title);

  // 取得第一頁（預設空白頁）作為封面
  var slides = presentation.getSlides();
  var coverSlide = slides[0];

  // ===== 封面頁 =====
  coverSlide.getBackground().setSolidFill('#1565C0');

  coverSlide.insertTextBox(outline.title, 50, 120, 620, 80)
    .getText()
    .getTextStyle()
    .setFontSize(36)
    .setBold(true)
    .setForegroundColor('#FFFFFF');

  coverSlide.insertTextBox(outline.subtitle || topic, 50, 220, 620, 50)
    .getText()
    .getTextStyle()
    .setFontSize(18)
    .setForegroundColor('#BBDEFB');

  coverSlide.insertTextBox('由 AI 自動生成 | ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd'), 50, 350, 620, 30)
    .getText()
    .getTextStyle()
    .setFontSize(12)
    .setForegroundColor('#90CAF9');

  // ===== 內容頁 =====
  for (var i = 0; i < outline.slides.length; i++) {
    var slideData = outline.slides[i];
    var slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

    // 頁碼色條
    slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 720, 60)
      .getFill()
      .setSolidFill('#1565C0');

    // 頁標題
    slide.insertTextBox(slideData.heading, 20, 8, 680, 45)
      .getText()
      .getTextStyle()
      .setFontSize(22)
      .setBold(true)
      .setForegroundColor('#FFFFFF');

    // 要點
    var pointsText = '';
    for (var j = 0; j < slideData.points.length; j++) {
      pointsText += '• ' + slideData.points[j] + '\n\n';
    }

    slide.insertTextBox(pointsText, 50, 90, 620, 300)
      .getText()
      .getTextStyle()
      .setFontSize(16)
      .setForegroundColor('#333333');
  }

  // ===== 結語頁 =====
  var endSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  endSlide.getBackground().setSolidFill('#1565C0');

  endSlide.insertTextBox('總結', 50, 100, 620, 60)
    .getText()
    .getTextStyle()
    .setFontSize(32)
    .setBold(true)
    .setForegroundColor('#FFFFFF');

  endSlide.insertTextBox(outline.conclusion || '感謝聆聽！', 50, 180, 620, 100)
    .getText()
    .getTextStyle()
    .setFontSize(18)
    .setForegroundColor('#BBDEFB');

  endSlide.insertTextBox('Thank You', 50, 320, 620, 50)
    .getText()
    .getTextStyle()
    .setFontSize(24)
    .setBold(true)
    .setForegroundColor('#FFFFFF');

  // 取得簡報連結
  var presentationUrl = presentation.getUrl();

  // 寫入試算表
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow() + 2;
  sheet.getRange(lastRow, 1).setValue('簡報主題').setFontWeight('bold');
  sheet.getRange(lastRow, 2).setValue(outline.title);
  sheet.getRange(lastRow + 1, 1).setValue('簡報連結').setFontWeight('bold');
  sheet.getRange(lastRow + 1, 2).setValue(presentationUrl);
  sheet.getRange(lastRow + 2, 1).setValue('頁數').setFontWeight('bold');
  sheet.getRange(lastRow + 2, 2).setValue(presentation.getSlides().length + ' 頁');

  Logger.log('========================================');
  Logger.log('簡報建立完成！');
  Logger.log('標題：' + outline.title);
  Logger.log('頁數：' + presentation.getSlides().length + ' 頁');
  Logger.log('連結：' + presentationUrl);
  Logger.log('========================================');
}

// ============================================================
// 第三部分：快速測試
// ============================================================

/**
 * 快速測試 — 用預設主題生成簡報
 */
function testGeneratePresentation() {
  generatePresentation('AI 在教育領域的應用');
}
