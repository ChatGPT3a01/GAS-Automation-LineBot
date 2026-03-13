/**
 * ============================================================
 * 實作 0：GAS 環境建立與第一支程式
 * ============================================================
 *
 * 功能說明：
 *   1. 認識 Google Apps Script 編輯器
 *   2. 學會用 Logger.log() 輸出訊息
 *   3. 學會讀取與寫入試算表儲存格
 *   4. 認識 SpreadsheetApp 的基本操作
 *
 * ⚠️ 注意：此程式需在 Google 試算表的「綁定專案」中執行
 *   （從試算表 → 擴充功能 → Apps Script 開啟的才是綁定專案）
 *
 * 使用方式：
 *   1. 打開任意 Google 試算表
 *   2. 點選「擴充功能 → Apps Script」
 *   3. 將此程式碼貼入編輯器
 *   4. 依序執行各個函式，觀察結果
 * ============================================================
 */

// ============================================================
// 第一部分：Hello World — Logger.log()
// ============================================================

/**
 * 你的第一支 GAS 程式
 * 點擊「執行」按鈕，然後到「執行記錄」查看輸出
 */
function helloWorld() {
  Logger.log('Hello World！這是我的第一支 GAS 程式！');
  Logger.log('目前時間：' + new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }));
}

// ============================================================
// 第二部分：讀取試算表
// ============================================================

/**
 * 讀取單一儲存格的值
 * 執行前請先在 A1 儲存格輸入任意文字
 */
function readCell() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var value = sheet.getRange('A1').getValue();
  Logger.log('A1 的值是：' + value);
}

/**
 * 讀取一個範圍的資料
 * 執行前請先在 A1:C3 輸入一些資料
 */
function readRange() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getRange('A1:C3').getValues();

  // data 是一個二維陣列（表格）
  Logger.log('共讀取 ' + data.length + ' 列資料');

  for (var i = 0; i < data.length; i++) {
    Logger.log('第 ' + (i + 1) + ' 列：' + data[i].join(' | '));
  }
}

// ============================================================
// 第三部分：寫入試算表
// ============================================================

/**
 * 寫入單一儲存格
 */
function writeCell() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('E1').setValue('這是 GAS 寫入的文字');
  sheet.getRange('E2').setValue(new Date());   // 寫入目前時間
  sheet.getRange('E3').setValue(12345);         // 寫入數字

  Logger.log('已成功寫入 E1~E3！');
}

/**
 * 寫入多個儲存格（使用 setValues 一次寫入）
 */
function writeRange() {
  var sheet = SpreadsheetApp.getActiveSheet();

  var data = [
    ['姓名', '分數', '等級'],
    ['王小明', 92, '優'],
    ['林小美', 85, '甲'],
    ['張大衛', 78, '乙']
  ];

  sheet.getRange(1, 7, data.length, data[0].length).setValues(data);
  Logger.log('已寫入 ' + data.length + ' 列資料到 G~I 欄！');
}

// ============================================================
// 第四部分：綜合練習 — 讀取再計算
// ============================================================

/**
 * 讀取試算表中的數字，計算總和與平均
 * 執行前請在 A1:A5 輸入五個數字
 */
function calculateSum() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getRange('A1:A5').getValues();

  var sum = 0;
  var count = 0;

  for (var i = 0; i < data.length; i++) {
    if (typeof data[i][0] === 'number') {
      sum += data[i][0];
      count++;
    }
  }

  var avg = count > 0 ? Math.round(sum / count * 10) / 10 : 0;

  Logger.log('總和：' + sum);
  Logger.log('平均：' + avg);
  Logger.log('共 ' + count + ' 筆數字');

  // 將結果寫回試算表
  sheet.getRange('C1').setValue('總和');
  sheet.getRange('D1').setValue(sum);
  sheet.getRange('C2').setValue('平均');
  sheet.getRange('D2').setValue(avg);

  Logger.log('計算結果已寫入 C1:D2！');
}
