/**
 * ============================================================
 * 實作 5：擴充選單與自訂按鈕
 * ============================================================
 *
 * 功能說明：
 *   1. 用 onOpen() 在試算表開啟時自動建立自訂選單
 *   2. 用 createMenu() 加入多個選項連結到自訂函式
 *   3. 用 ui.alert() 顯示確認對話框
 *   4. 用 ui.prompt() 顯示輸入對話框
 *   5. 將試算表變成「有按鈕的應用程式」
 *
 * ⚠️ 注意：此程式只能在 Google 試算表的「綁定專案」中使用
 *   （從試算表 → 擴充功能 → Apps Script 開啟的才是綁定專案）
 *
 * 使用方式：
 *   1. 將此程式碼貼入試算表的 Apps Script
 *   2. 重新整理試算表頁面（或關閉再開啟）
 *   3. 上方選單列會出現「自動化工具」選單
 *   4. 點選各項功能即可執行
 * ============================================================
 */

// ============================================================
// 第一部分：建立自訂選單
// ============================================================

/**
 * onOpen() 是特殊觸發函式 — 試算表開啟時自動執行
 * 用途：在選單列建立自訂選單
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('自動化工具')
    .addItem('建立今日報表', 'menuCreateReport')
    .addItem('快速新增資料', 'menuAddData')
    .addSeparator()
    .addItem('清除所有資料', 'menuClearData')
    .addSeparator()
    .addItem('關於此工具', 'menuAbout')
    .addToUi();

  Logger.log('自訂選單已建立');
}

// ============================================================
// 第二部分：選單功能 — alert 對話框
// ============================================================

/**
 * 功能一：建立今日報表（示範 alert 確認框）
 */
function menuCreateReport() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

  // 彈出確認對話框
  var result = ui.alert(
    '建立報表',
    '是否要建立今日報表？\n這將在新工作表中產生統計。',
    ui.ButtonSet.YES_NO
  );

  // 判斷使用者的選擇
  if (result === ui.Button.YES) {
    // 使用者按了「是」
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reportName = '報表_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'MMdd');

    var reportSheet = ss.getSheetByName(reportName);
    if (reportSheet) ss.deleteSheet(reportSheet);
    reportSheet = ss.insertSheet(reportName);

    reportSheet.getRange('A1').setValue('今日報表').setFontSize(16).setFontWeight('bold');
    reportSheet.getRange('A2').setValue('產出時間：' + new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }));
    reportSheet.getRange('A4').setValue('（此處可加入自動統計邏輯）');

    ui.alert('完成', '報表「' + reportName + '」已建立！', ui.ButtonSet.OK);
  } else {
    Logger.log('使用者取消了報表建立');
  }
}

// ============================================================
// 第三部分：選單功能 — prompt 輸入框
// ============================================================

/**
 * 功能二：快速新增資料（示範 prompt 輸入框）
 */
function menuAddData() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

  // 彈出輸入對話框 — 輸入姓名
  var nameResponse = ui.prompt(
    '新增資料',
    '請輸入姓名：',
    ui.ButtonSet.OK_CANCEL
  );

  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  var name = nameResponse.getResponseText();

  // 彈出第二個輸入框 — 輸入分數
  var scoreResponse = ui.prompt(
    '新增資料',
    '請輸入 ' + name + ' 的分數：',
    ui.ButtonSet.OK_CANCEL
  );

  if (scoreResponse.getSelectedButton() !== ui.Button.OK) return;
  var score = Number(scoreResponse.getResponseText());

  // 寫入試算表
  var lastRow = sheet.getLastRow();

  // 如果是第一筆，先建立標題
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['姓名', '分數', '新增時間']]);
    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setBackground('#4CAF50')
      .setFontColor('#FFFFFF');
    lastRow = 1;
  }

  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  sheet.appendRow([name, score, now]);

  ui.alert('新增成功', name + ' 的資料已新增！\n分數：' + score, ui.ButtonSet.OK);
}

// ============================================================
// 第四部分：選單功能 — 確認刪除
// ============================================================

/**
 * 功能三：清除所有資料（示範危險操作的確認流程）
 */
function menuClearData() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

  // 第一次確認
  var result = ui.alert(
    '⚠️ 警告',
    '確定要清除目前工作表的所有資料嗎？\n此操作無法復原！',
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    // 第二次確認（重要操作雙重確認）
    var confirm = ui.alert(
      '再次確認',
      '真的要清除嗎？請再按一次「是」。',
      ui.ButtonSet.YES_NO
    );

    if (confirm === ui.Button.YES) {
      sheet.clear();
      ui.alert('已清除', '所有資料已清除。', ui.ButtonSet.OK);
      Logger.log('使用者已清除工作表資料');
    }
  }
}

// ============================================================
// 第五部分：關於此工具
// ============================================================

/**
 * 功能四：顯示關於資訊
 */
function menuAbout() {
  var ui = SpreadsheetApp.getUi();

  var message = '自動化工具 v1.0\n\n';
  message += '功能列表：\n';
  message += '• 建立今日報表\n';
  message += '• 快速新增資料\n';
  message += '• 清除所有資料\n\n';
  message += '由 Google Apps Script 驅動\n';
  message += '實作 5：擴充選單與自訂按鈕';

  ui.alert('關於', message, ui.ButtonSet.OK);
}
