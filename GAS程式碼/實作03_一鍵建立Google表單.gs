/**
 * ============================================================
 * 實作 3：一鍵自動建立 Google 表單
 * ============================================================
 *
 * 功能說明：
 *   1. 從試算表讀取題目清單（題目、選項、類型）
 *   2. 用 FormApp.create() 自動建立 Google Form
 *   3. 支援多種題型：單選、多選、文字、下拉
 *   4. 設定表單確認訊息和收集 Email
 *   5. 表單連結回寫試算表
 *
 * 試算表結構（執行 createSampleQuestions 自動建立）：
 *   A 欄：題號
 *   B 欄：題目
 *   C 欄：題型（單選/多選/文字/下拉）
 *   D 欄：選項（用逗號分隔）
 *   E 欄：是否必填（是/否）
 *
 * 使用方式：
 *   1. 執行 createSampleQuestions() → 建立範例題目
 *   2. 執行 createFormFromSheet() → 自動建立 Google 表單
 *   3. 到執行記錄查看表單連結
 *
 * ⚠️ 注意：此程式需在 Google 試算表的「綁定專案」中執行
 * ============================================================
 */

// ============================================================
// 第一部分：建立範例題目清單
// ============================================================

/**
 * 建立範例題目清單
 */
function createSampleQuestions() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();

  // 標題列
  var headers = ['題號', '題目', '題型', '選項（逗號分隔）', '必填'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4CAF50')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');

  // 範例題目
  var questions = [
    [1, '您的姓名', '文字', '', '是'],
    [2, '您的部門', '下拉', '行政處,教務處,學務處,總務處,輔導室', '是'],
    [3, '您想參加哪一場研習？', '單選', '場次A：GAS入門,場次B：AI應用,場次C：數據分析', '是'],
    [4, '您希望學到什麼？（可複選）', '多選', '自動寄信,自動建表單,Line Bot,AI 整合', '否'],
    [5, '其他建議', '文字', '', '否']
  ];

  sheet.getRange(2, 1, questions.length, 5).setValues(questions);

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 350);
  sheet.setColumnWidth(5, 60);
  sheet.setFrozenRows(1);

  Logger.log('範例題目建立完成！共 ' + questions.length + ' 題。');
}

// ============================================================
// 第二部分：自動建立 Google 表單
// ============================================================

/**
 * 從試算表讀取題目，自動建立 Google 表單
 */
function createFormFromSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    Logger.log('沒有題目資料，請先執行 createSampleQuestions()');
    return;
  }

  // 建立新表單
  var formTitle = '自動建立的問卷 - ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'MM/dd HH:mm');
  var form = FormApp.create(formTitle);

  // 設定表單描述
  form.setDescription('此表單由 Google Apps Script 自動建立。');

  // 設定確認訊息
  form.setConfirmationMessage('感謝您的填寫！回覆已成功送出。');

  // 收集 Email
  form.setCollectEmail(true);

  // 讀取題目資料
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  for (var i = 0; i < data.length; i++) {
    var questionText = data[i][1];
    var questionType = data[i][2];
    var optionsStr = data[i][3];
    var required = (data[i][4] === '是');

    if (!questionText) continue;

    var options = optionsStr ? String(optionsStr).split(',').map(function(o) { return o.trim(); }) : [];

    switch (questionType) {
      case '文字':
        form.addTextItem()
          .setTitle(questionText)
          .setRequired(required);
        break;

      case '單選':
        if (options.length > 0) {
          form.addMultipleChoiceItem()
            .setTitle(questionText)
            .setChoiceValues(options)
            .setRequired(required);
        }
        break;

      case '多選':
        if (options.length > 0) {
          form.addCheckboxItem()
            .setTitle(questionText)
            .setChoiceValues(options)
            .setRequired(required);
        }
        break;

      case '下拉':
        if (options.length > 0) {
          form.addListItem()
            .setTitle(questionText)
            .setChoiceValues(options)
            .setRequired(required);
        }
        break;

      default:
        form.addTextItem()
          .setTitle(questionText)
          .setRequired(required);
    }

    Logger.log('已新增題目：' + questionText + '（' + questionType + '）');
  }

  // 取得表單連結
  var formUrl = form.getPublishedUrl();
  var editUrl = form.getEditUrl();

  // 將連結寫回試算表
  var infoRow = lastRow + 2;
  sheet.getRange(infoRow, 1).setValue('表單連結');
  sheet.getRange(infoRow, 2).setValue(formUrl);
  sheet.getRange(infoRow, 1).setFontWeight('bold').setBackground('#E8F5E9');

  sheet.getRange(infoRow + 1, 1).setValue('編輯連結');
  sheet.getRange(infoRow + 1, 2).setValue(editUrl);
  sheet.getRange(infoRow + 1, 1).setFontWeight('bold').setBackground('#FFF3E0');

  Logger.log('========================================');
  Logger.log('表單建立成功！');
  Logger.log('表單連結：' + formUrl);
  Logger.log('編輯連結：' + editUrl);
  Logger.log('========================================');
}
