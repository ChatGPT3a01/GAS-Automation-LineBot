/**
 * ============================================================
 * 實作 20：AI 段考審題助手
 * ============================================================
 *
 * 功能說明：
 *   1. 從 Google Docs 讀取段考試題
 *   2. 呼叫 Gemini AI 審查題目品質
 *   3. AI 分析：題目清晰度、難度分布、選項品質
 *   4. 將審查結果寫入試算表
 *
 * 使用方式：
 *   1. 準備一份 Google 文件，貼上段考題目
 *   2. 將文件 ID 填入下方 DOC_ID
 *   3. 填入 Gemini API Key
 *   4. 執行 reviewExamQuestions() 開始審題
 *   或執行 createSampleExam() 建立範例試卷
 *
 * Google AI Studio：https://aistudio.google.com/apikey
 * ============================================================
 */

// ========== 設定區 ==========
var GEMINI_API_KEY = '在此貼上你的 Gemini API Key';
var GEMINI_MODEL = 'gemini-2.5-flash';
var DOC_ID = '在此貼上你的 Google 文件 ID';

// ============================================================
// 第一部分：建立範例試卷（Google Docs）
// ============================================================

/**
 * 建立一份範例段考試卷（Google Docs）
 */
function createSampleExam() {
  var doc = DocumentApp.create('範例段考試卷 - AI 審題用');
  var body = doc.getBody();

  body.appendParagraph('國中八年級 自然科 第一次段考')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('');

  var questions = [
    {
      q: '1. 下列哪一個是光合作用的產物？',
      options: '(A) 二氧化碳  (B) 水  (C) 葡萄糖  (D) 氧氣和葡萄糖'
    },
    {
      q: '2. 植物的哪個部位主要負責吸收水分？',
      options: '(A) 葉子  (B) 莖  (C) 根  (D) 花'
    },
    {
      q: '3. 以下關於細胞的敘述，何者正確？',
      options: '(A) 所有細胞都有細胞壁  (B) 動物細胞沒有粒線體  (C) 植物細胞有葉綠體  (D) 細胞核存在於所有細胞中'
    },
    {
      q: '4. 下面那個東西比較重？',
      options: '(A) 鐵  (B) 棉花  (C) 不一定  (D) 以上皆非'
    },
    {
      q: '5. 請說明光合作用的過程及其重要性。',
      options: '（問答題）'
    }
  ];

  for (var i = 0; i < questions.length; i++) {
    body.appendParagraph(questions[i].q);
    body.appendParagraph(questions[i].options);
    body.appendParagraph('');
  }

  doc.saveAndClose();

  var docUrl = doc.getUrl();
  var docId = doc.getId();

  // 記錄到試算表
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('A1').setValue('範例試卷 ID').setFontWeight('bold');
  sheet.getRange('B1').setValue(docId);
  sheet.getRange('A2').setValue('範例試卷連結').setFontWeight('bold');
  sheet.getRange('B2').setValue(docUrl);

  Logger.log('範例試卷已建立！');
  Logger.log('文件 ID：' + docId);
  Logger.log('連結：' + docUrl);
  Logger.log('請將此 ID 填入 DOC_ID 變數，再執行 reviewExamQuestions()');
}

// ============================================================
// 第二部分：讀取 Google Docs 試卷內容
// ============================================================

/**
 * 從 Google Docs 讀取試卷內容
 * @param {string} docId - Google 文件 ID
 * @return {string} 試卷文字內容
 */
function readExamDoc(docId) {
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  var text = body.getText();
  Logger.log('已讀取試卷，共 ' + text.length + ' 個字');
  return text;
}

// ============================================================
// 第三部分：呼叫 Gemini AI 審查試題
// ============================================================

/**
 * 用 Gemini AI 審查段考試題
 * @param {string} examContent - 試卷文字內容
 * @return {Object} AI 審查結果
 */
function aiReviewExam(examContent) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + GEMINI_API_KEY;

  var prompt = '你是一位資深的國中教師和試題審查專家。請審查以下段考試卷，回傳 JSON 格式的審查報告。\n\n';
  prompt += '試卷內容：\n' + examContent + '\n\n';
  prompt += '請回傳以下 JSON 格式（不要加 markdown 標記）：\n';
  prompt += '{\n';
  prompt += '  "totalQuestions": 題目總數,\n';
  prompt += '  "overallScore": 整體品質分數(1-10),\n';
  prompt += '  "overallComment": "整體評語",\n';
  prompt += '  "difficultyDistribution": "難度分布說明",\n';
  prompt += '  "questions": [\n';
  prompt += '    {\n';
  prompt += '      "number": 題號,\n';
  prompt += '      "clarity": 清晰度(1-10),\n';
  prompt += '      "difficulty": "簡單/中等/困難",\n';
  prompt += '      "optionQuality": 選項品質(1-10),\n';
  prompt += '      "issue": "問題描述（無問題則留空）",\n';
  prompt += '      "suggestion": "改善建議（無則留空）"\n';
  prompt += '    }\n';
  prompt += '  ],\n';
  prompt += '  "suggestions": ["整體建議1", "整體建議2"]\n';
  prompt += '}\n\n';
  prompt += '請使用繁體中文。';

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.3 }
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
// 第四部分：將審查結果寫入試算表
// ============================================================

/**
 * 主流程：讀取試卷 → AI 審查 → 寫入試算表
 */
function reviewExamQuestions() {
  // 讀取試卷
  Logger.log('正在讀取試卷...');
  var examContent = readExamDoc(DOC_ID);

  // AI 審查
  Logger.log('正在用 AI 審查試題...');
  var review = aiReviewExam(examContent);

  if (!review) {
    Logger.log('AI 審查失敗，請檢查 API Key 和文件 ID');
    return;
  }

  // 寫入試算表
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = '審題結果_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'MMdd_HHmm');
  var sheet = ss.insertSheet(sheetName);

  // 標題區
  sheet.getRange('A1').setValue('AI 段考審題報告').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2').setValue('產出時間：' + new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }));
  sheet.getRange('A3').setValue('整體評分：' + review.overallScore + ' / 10').setFontWeight('bold');
  sheet.getRange('A4').setValue('整體評語：' + review.overallComment);
  sheet.getRange('A5').setValue('難度分布：' + review.difficultyDistribution);

  // 逐題審查表
  var headers = ['題號', '清晰度', '難度', '選項品質', '問題', '改善建議'];
  sheet.getRange(7, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(7, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#FF6F00')
    .setFontColor('#FFFFFF');

  if (review.questions) {
    for (var i = 0; i < review.questions.length; i++) {
      var q = review.questions[i];
      var row = 8 + i;
      sheet.getRange(row, 1).setValue(q.number || (i + 1));
      sheet.getRange(row, 2).setValue((q.clarity || '-') + '/10');
      sheet.getRange(row, 3).setValue(q.difficulty || '-');
      sheet.getRange(row, 4).setValue((q.optionQuality || '-') + '/10');
      sheet.getRange(row, 5).setValue(q.issue || '無');
      sheet.getRange(row, 6).setValue(q.suggestion || '無');

      // 品質低的用紅色標記
      if (q.clarity < 6 || q.optionQuality < 6) {
        sheet.getRange(row, 1, 1, 6).setBackground('#FFEBEE');
      }
    }
  }

  // 整體建議
  var sugRow = 8 + (review.questions ? review.questions.length : 0) + 1;
  sheet.getRange(sugRow, 1).setValue('整體建議').setFontWeight('bold').setFontSize(12);
  if (review.suggestions) {
    for (var j = 0; j < review.suggestions.length; j++) {
      sheet.getRange(sugRow + 1 + j, 1).setValue('• ' + review.suggestions[j]);
    }
  }

  // 設定欄寬
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 300);
  sheet.setColumnWidth(6, 300);

  Logger.log('========================================');
  Logger.log('審題完成！結果已寫入工作表「' + sheetName + '」');
  Logger.log('整體評分：' + review.overallScore + '/10');
  Logger.log('========================================');
}
