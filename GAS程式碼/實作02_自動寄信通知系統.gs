/**
 * ============================================================
 * 實作 2：自動寄信通知系統
 * ============================================================
 *
 * 功能說明：
 *   1. 從試算表讀取收件人名單（姓名、Email、主題）
 *   2. 用 MailApp.sendEmail() 逐一寄出通知信
 *   3. 支援 HTML 格式信件（粗體、顏色、表格）
 *   4. 寄信記錄回寫試算表
 *
 * 試算表結構（請先手動建立或執行 createSampleRecipients）：
 *   A 欄：姓名
 *   B 欄：Email
 *   C 欄：主題（選填，空白則用預設）
 *   D 欄：寄送狀態（程式自動填入）
 *   E 欄：寄送時間（程式自動填入）
 *
 * 使用方式：
 *   1. 執行 createSampleRecipients() → 建立範例名單
 *   2. 執行 sendNotifications() → 逐一寄信
 *   3. 執行 sendHtmlEmail() → 寄送 HTML 格式信件
 *   4. 執行 checkQuota() → 查看今日剩餘寄信額度
 *
 * ⚠️ 注意：
 *   - 此程式需在 Google 試算表的「綁定專案」中執行
 *   - 範例 Email 請改成真實地址才能測試
 *   - 免費帳號每天上限 100 封，Google Workspace 每天 1500 封
 * ============================================================
 */

// ============================================================
// 第一部分：建立範例名單
// ============================================================

/**
 * 建立範例收件人名單
 */
function createSampleRecipients() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();

  // 標題列
  var headers = ['姓名', 'Email', '主題', '寄送狀態', '寄送時間'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1565C0')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');

  // 範例名單（請改成真實 Email 才能測試）
  var recipients = [
    ['王小明', 'your-email@gmail.com', '活動通知'],
    ['林小美', 'your-email@gmail.com', '會議提醒'],
    ['張大衛', 'your-email@gmail.com', '']
  ];

  sheet.getRange(2, 1, recipients.length, 3).setValues(recipients);

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 200);
  sheet.setFrozenRows(1);

  Logger.log('範例名單建立完成！請將 Email 改成真實地址再寄送。');
}

// ============================================================
// 第二部分：逐一寄信
// ============================================================

/**
 * 從試算表讀取名單，逐一寄出通知信
 */
function sendNotifications() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    Logger.log('沒有收件人資料，請先執行 createSampleRecipients()');
    return;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var successCount = 0;
  var failCount = 0;

  for (var i = 0; i < data.length; i++) {
    var name = data[i][0];
    var email = data[i][1];
    var subject = data[i][2] || '通知信件';
    var row = i + 2;

    if (!email || email.indexOf('@') === -1) {
      sheet.getRange(row, 4).setValue('略過');
      sheet.getRange(row, 5).setValue('Email 格式無效');
      failCount++;
      continue;
    }

    try {
      var body = name + ' 您好，\n\n';
      body += '這是一封由 Google Apps Script 自動寄出的通知信。\n\n';
      body += '此信件用於測試 GAS 自動寄信功能。\n\n';
      body += '祝好，\nGAS 自動化系統';

      MailApp.sendEmail(email, subject, body);

      var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
      sheet.getRange(row, 4).setValue('已寄出');
      sheet.getRange(row, 4).setBackground('#E8F5E9');
      sheet.getRange(row, 5).setValue(now);
      successCount++;

    } catch (e) {
      sheet.getRange(row, 4).setValue('失敗');
      sheet.getRange(row, 4).setBackground('#FFEBEE');
      sheet.getRange(row, 5).setValue(e.message);
      failCount++;
    }
  }

  Logger.log('寄信完成！成功 ' + successCount + ' 封，失敗 ' + failCount + ' 封。');
}

// ============================================================
// 第三部分：HTML 格式信件
// ============================================================

/**
 * 寄送 HTML 格式的通知信
 * 支援粗體、顏色、表格等格式
 */
function sendHtmlEmail() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    Logger.log('沒有收件人資料');
    return;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  for (var i = 0; i < data.length; i++) {
    var name = data[i][0];
    var email = data[i][1];
    var subject = data[i][2] || 'HTML 通知信件';
    var row = i + 2;

    if (!email || email.indexOf('@') === -1) continue;

    // 組合 HTML 內容
    var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">';
    htmlBody += '<h2 style="color: #1565C0;">📬 自動通知</h2>';
    htmlBody += '<p><strong>' + name + '</strong> 您好，</p>';
    htmlBody += '<p>這是一封 <span style="color: #E65100; font-weight: bold;">HTML 格式</span> 的自動通知信。</p>';
    htmlBody += '<table style="border-collapse: collapse; width: 100%; margin: 16px 0;">';
    htmlBody += '<tr style="background: #1565C0; color: white;">';
    htmlBody += '<th style="padding: 8px; border: 1px solid #ddd;">項目</th>';
    htmlBody += '<th style="padding: 8px; border: 1px solid #ddd;">內容</th></tr>';
    htmlBody += '<tr><td style="padding: 8px; border: 1px solid #ddd;">收件人</td>';
    htmlBody += '<td style="padding: 8px; border: 1px solid #ddd;">' + name + '</td></tr>';
    htmlBody += '<tr><td style="padding: 8px; border: 1px solid #ddd;">寄送時間</td>';
    htmlBody += '<td style="padding: 8px; border: 1px solid #ddd;">' + new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }) + '</td></tr>';
    htmlBody += '</table>';
    htmlBody += '<p style="color: #666; font-size: 12px;">此信件由 GAS 自動化系統寄出</p>';
    htmlBody += '</div>';

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlBody,
        body: name + ' 您好，這是自動通知信。' // 純文字備援
      });

      var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
      sheet.getRange(row, 4).setValue('已寄出(HTML)');
      sheet.getRange(row, 4).setBackground('#E8F5E9');
      sheet.getRange(row, 5).setValue(now);

    } catch (e) {
      sheet.getRange(row, 4).setValue('失敗');
      sheet.getRange(row, 4).setBackground('#FFEBEE');
      sheet.getRange(row, 5).setValue(e.message);
    }
  }

  Logger.log('HTML 信件寄送完成！');
}

// ============================================================
// 第四部分：查看寄信額度
// ============================================================

/**
 * 查看今日剩餘寄信額度
 * 免費帳號每天上限 100 封，Google Workspace 每天 1500 封
 */
function checkQuota() {
  var remaining = MailApp.getRemainingDailyQuota();
  Logger.log('今日剩餘寄信額度：' + remaining + ' 封');
}
