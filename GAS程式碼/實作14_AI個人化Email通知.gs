/**
 * ============================================================
 * 實作 14：AI 個人化 Email 通知
 * ============================================================
 *
 * 功能說明：
 *   1. 從試算表讀取學生名單和成績
 *   2. 呼叫 Gemini AI 根據成績生成個人化鼓勵語
 *   3. 用 GmailApp.sendEmail() 寄出 HTML 格式信件
 *   4. 寄信進度和結果回寫試算表
 *
 * 試算表結構（執行 createSampleStudents 自動建立）：
 *   A 欄：姓名
 *   B 欄：Email
 *   C 欄：國文
 *   D 欄：英文
 *   E 欄：數學
 *   F 欄：平均
 *   G 欄：AI 評語
 *   H 欄：寄送狀態
 *
 * 設定步驟：
 *   1. 到 Google AI Studio 取得 Gemini API Key
 *   2. 將 API Key 填入下方 GEMINI_API_KEY
 *   3. 執行 createSampleStudents() → 建立範例資料
 *   4. 執行 generateAndSendAll() → AI 生成評語 + 寄信
 *
 * Google AI Studio：https://aistudio.google.com/apikey
 * ============================================================
 */

// ========== 設定區 ==========
var GEMINI_API_KEY = '在此貼上你的 Gemini API Key';
var GEMINI_MODEL = 'gemini-2.5-flash';

// ============================================================
// 第一部分：建立範例學生資料
// ============================================================

/**
 * 建立範例學生名單與成績
 */
function createSampleStudents() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();

  var headers = ['姓名', 'Email', '國文', '英文', '數學', '平均', 'AI 評語', '寄送狀態'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#7B1FA2')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');

  // 範例學生（Email 請改成真實地址）
  var students = [
    ['陳小明', 'your-email@gmail.com', 85, 72, 90],
    ['林小美', 'your-email@gmail.com', 92, 88, 78],
    ['張大衛', 'your-email@gmail.com', 65, 58, 45],
    ['王小華', 'your-email@gmail.com', 78, 82, 95],
    ['李建國', 'your-email@gmail.com', 55, 48, 62]
  ];

  for (var i = 0; i < students.length; i++) {
    var row = i + 2;
    sheet.getRange(row, 1, 1, 5).setValues([students[i]]);
    // F 欄：平均（公式）
    sheet.getRange(row, 6).setFormula('=ROUND(AVERAGE(C' + row + ':E' + row + '),1)');
  }

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 60);
  sheet.setColumnWidth(6, 60);
  sheet.setColumnWidth(7, 400);
  sheet.setColumnWidth(8, 100);
  sheet.setFrozenRows(1);

  Logger.log('範例學生資料建立完成！請將 Email 改成真實地址。');
}

// ============================================================
// 第二部分：呼叫 Gemini AI 生成個人化評語
// ============================================================

/**
 * 用 Gemini AI 生成個人化鼓勵語
 * @param {string} name - 學生姓名
 * @param {number} chinese - 國文分數
 * @param {number} english - 英文分數
 * @param {number} math - 數學分數
 * @return {string} AI 生成的個人化評語
 */
function generateComment(name, chinese, english, math) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + GEMINI_API_KEY;

  var avg = Math.round((chinese + english + math) / 3 * 10) / 10;

  var prompt = '你是一位溫暖的導師。請根據以下成績，為學生寫一段 50-80 字的個人化鼓勵語。\n\n';
  prompt += '學生：' + name + '\n';
  prompt += '國文：' + chinese + ' 分\n';
  prompt += '英文：' + english + ' 分\n';
  prompt += '數學：' + math + ' 分\n';
  prompt += '平均：' + avg + ' 分\n\n';
  prompt += '要求：\n';
  prompt += '- 使用繁體中文\n';
  prompt += '- 肯定優秀科目，鼓勵弱科進步\n';
  prompt += '- 語氣溫暖正面\n';
  prompt += '- 直接回覆評語內容，不要加標題或格式';

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.8, maxOutputTokens: 200 }
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
    return '（AI 評語生成失敗：回應格式錯誤）';
  }

  if (json.error) {
    Logger.log('Gemini API 錯誤：' + json.error.message);
    return '（AI 評語生成失敗）';
  }

  if (!json.candidates || !json.candidates[0] ||
      !json.candidates[0].content || !json.candidates[0].content.parts) {
    Logger.log('Gemini 回覆格式異常');
    return '（AI 評語生成失敗：回覆格式異常）';
  }

  return json.candidates[0].content.parts[0].text.trim();
}

// ============================================================
// 第三部分：寄送 HTML 格式個人化 Email
// ============================================================

/**
 * 寄送個人化成績通知 Email
 */
function sendPersonalizedEmail(name, email, chinese, english, math, comment) {
  var avg = Math.round((chinese + english + math) / 3 * 10) / 10;

  var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">';
  htmlBody += '<h2 style="color: #7B1FA2; border-bottom: 3px solid #7B1FA2; padding-bottom: 10px;">📊 個人成績通知</h2>';
  htmlBody += '<p>親愛的 <strong>' + name + '</strong> 同學，你好！</p>';
  htmlBody += '<p>以下是你這次的成績：</p>';

  // 成績表格
  htmlBody += '<table style="border-collapse: collapse; width: 100%; margin: 16px 0;">';
  htmlBody += '<tr style="background: #7B1FA2; color: white;">';
  htmlBody += '<th style="padding: 10px; border: 1px solid #ddd;">科目</th>';
  htmlBody += '<th style="padding: 10px; border: 1px solid #ddd;">分數</th></tr>';

  var subjects = [['國文', chinese], ['英文', english], ['數學', math]];
  for (var i = 0; i < subjects.length; i++) {
    var color = subjects[i][1] >= 80 ? '#E8F5E9' : (subjects[i][1] >= 60 ? '#FFF3E0' : '#FFEBEE');
    htmlBody += '<tr style="background: ' + color + ';">';
    htmlBody += '<td style="padding: 8px; border: 1px solid #ddd; text-align: center;">' + subjects[i][0] + '</td>';
    htmlBody += '<td style="padding: 8px; border: 1px solid #ddd; text-align: center; font-weight: bold;">' + subjects[i][1] + '</td></tr>';
  }

  htmlBody += '<tr style="background: #F3E5F5; font-weight: bold;">';
  htmlBody += '<td style="padding: 8px; border: 1px solid #ddd; text-align: center;">平均</td>';
  htmlBody += '<td style="padding: 8px; border: 1px solid #ddd; text-align: center;">' + avg + '</td></tr>';
  htmlBody += '</table>';

  // AI 評語
  htmlBody += '<div style="background: #F3E5F5; border-left: 4px solid #7B1FA2; padding: 15px; margin: 16px 0; border-radius: 4px;">';
  htmlBody += '<p style="margin: 0; font-style: italic;">💬 老師的話：</p>';
  htmlBody += '<p style="margin: 8px 0 0 0;">' + comment + '</p>';
  htmlBody += '</div>';

  htmlBody += '<p style="color: #666; font-size: 12px; margin-top: 30px;">此信件由 GAS 自動化系統寄出</p>';
  htmlBody += '</div>';

  GmailApp.sendEmail(email, name + ' 同學 — 成績通知與鼓勵', '', { htmlBody: htmlBody });
}

// ============================================================
// 第四部分：主流程 — AI 生成 + 寄信
// ============================================================

/**
 * 對所有學生生成 AI 評語並寄送 Email
 */
function generateAndSendAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    Logger.log('沒有學生資料，請先執行 createSampleStudents()');
    return;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = i + 2;
    var name = data[i][0];
    var email = data[i][1];
    var chinese = data[i][2];
    var english = data[i][3];
    var math = data[i][4];

    if (!email || email.indexOf('@') === -1) {
      sheet.getRange(row, 8).setValue('略過');
      continue;
    }

    try {
      // 1. AI 生成評語
      sheet.getRange(row, 8).setValue('生成中...');
      var comment = generateComment(name, chinese, english, math);
      sheet.getRange(row, 7).setValue(comment);

      // 2. 寄送 Email
      sheet.getRange(row, 8).setValue('寄送中...');
      sendPersonalizedEmail(name, email, chinese, english, math, comment);

      sheet.getRange(row, 8).setValue('已寄出').setBackground('#E8F5E9');
      Logger.log('完成：' + name);

      // 避免 API 過載，間隔 2 秒
      Utilities.sleep(2000);

    } catch (e) {
      sheet.getRange(row, 8).setValue('失敗').setBackground('#FFEBEE');
      Logger.log('失敗：' + name + ' - ' + e.message);
    }
  }

  Logger.log('全部處理完成！');
}

// ============================================================
// 第五部分：單筆測試
// ============================================================

/**
 * 測試 AI 評語生成（不寄信）
 */
function testGenerateComment() {
  var comment = generateComment('測試同學', 85, 72, 90);
  Logger.log('AI 評語：\n' + comment);
}
