/**
 * ============================================================
 * 實作 6：天氣預報 + Line Bot 每日推播
 * ============================================================
 *
 * 功能說明：
 *   1. 串接中央氣象署開放資料 API（36小時天氣預報）
 *   2. 取得指定縣市的天氣預報（天氣狀況、降雨機率、溫度）
 *   3. 格式化成易讀的天氣報告
 *   4. 透過 Line Bot 推播天氣預報
 *   5. 設定每天早上自動推播
 *
 * 事前準備：
 *   1. 到中央氣象署開放資料平台申請 API 授權碼
 *      網址：https://opendata.cwa.gov.tw/
 *      步驟：註冊帳號 → 登入 → 取得授權碼
 *   2. 取得 Line Bot 的 Channel Access Token
 *   3. 取得你的 Line User ID
 *
 * 設定排程自動執行：
 *   1. 在 Apps Script 編輯器中，點選左側「觸發條件」（鬧鐘圖示）
 *   2. 點選「新增觸發條件」
 *   3. 選擇函式：sendWeatherReport
 *   4. 事件來源：時間驅動
 *   5. 時間型觸發條件類型：每日計時器
 *   6. 時段：上午 7 點到 8 點
 *   7. 儲存
 * ============================================================
 */

// ========== Line Bot 設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';

// ========== 氣象 API 設定 ==========
var CWA_API_KEY = '在此貼上你的中央氣象署 API 授權碼';

// 要查詢的縣市（可以修改）
var TARGET_CITY = '臺北市';

// API 資料集代碼（36小時天氣預報）
var DATASET_ID = 'F-C0032-001';

// ============================================================
// 第一部分：取得天氣資料
// ============================================================

/**
 * 從中央氣象署 API 取得天氣預報
 *
 * @param {string} cityName - 縣市名稱（例：臺北市、高雄市）
 * @returns {Object|null} 天氣資料物件，失敗時回傳 null
 */
function getWeatherData(cityName) {
  // 組合 API 網址
  var url = 'https://opendata.cwa.gov.tw/api/v1/rest/datastore/' + DATASET_ID;
  url += '?Authorization=' + CWA_API_KEY;
  url += '&locationName=' + encodeURIComponent(cityName);

  try {
    // 呼叫 API
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var code = response.getResponseCode();

    if (code !== 200) {
      Logger.log('API 呼叫失敗，狀態碼：' + code);
      Logger.log('回應內容：' + response.getContentText());
      return null;
    }

    // 解析 JSON
    var json = JSON.parse(response.getContentText());

    // 檢查 API 回應是否成功
    if (json.success !== 'true') {
      Logger.log('API 回應失敗：' + JSON.stringify(json));
      return null;
    }

    // 取得指定縣市的資料
    var locations = json.records.location;
    if (!locations || locations.length === 0) {
      Logger.log('找不到 ' + cityName + ' 的資料');
      return null;
    }

    return locations[0];
  } catch (error) {
    Logger.log('getWeatherData 錯誤：' + error.toString());
    return null;
  }
}

// ============================================================
// 第二部分：解析天氣資料
// ============================================================

/**
 * 從天氣資料中提取各項天氣元素
 *
 * API 回傳的 weatherElement 包含：
 *   Wx  — 天氣現象（晴、多雲、陰天等）
 *   PoP — 降雨機率（%）
 *   MinT — 最低溫度（°C）
 *   MaxT — 最高溫度（°C）
 *   CI  — 舒適度（舒適、悶熱等）
 *
 * @param {Object} locationData - API 回傳的單一縣市資料
 * @returns {Array} 各時段的天氣資訊陣列
 */
function parseWeatherData(locationData) {
  var elements = locationData.weatherElement;
  var forecasts = [];

  // 找出各個天氣元素
  var wxData = findElement(elements, 'Wx');      // 天氣現象
  var popData = findElement(elements, 'PoP');     // 降雨機率
  var minTData = findElement(elements, 'MinT');   // 最低溫
  var maxTData = findElement(elements, 'MaxT');   // 最高溫
  var ciData = findElement(elements, 'CI');       // 舒適度

  // 取得時段數量（通常是 3 個時段）
  var timeCount = wxData ? wxData.time.length : 0;

  for (var i = 0; i < timeCount; i++) {
    var forecast = {
      startTime: wxData.time[i].startTime,
      endTime: wxData.time[i].endTime,
      weather: wxData ? wxData.time[i].parameter.parameterName : '未知',
      rainProb: popData ? popData.time[i].parameter.parameterName : '未知',
      minTemp: minTData ? minTData.time[i].parameter.parameterName : '未知',
      maxTemp: maxTData ? maxTData.time[i].parameter.parameterName : '未知',
      comfort: ciData ? ciData.time[i].parameter.parameterName : '未知'
    };

    forecasts.push(forecast);
  }

  return forecasts;
}

/**
 * 從 weatherElement 陣列中找到指定名稱的元素
 */
function findElement(elements, elementName) {
  for (var i = 0; i < elements.length; i++) {
    if (elements[i].elementName === elementName) {
      return elements[i];
    }
  }
  return null;
}

// ============================================================
// 第三部分：格式化天氣報告
// ============================================================

/**
 * 將天氣資料格式化為 Line 訊息
 *
 * @param {string} cityName - 縣市名稱
 * @param {Array} forecasts - 天氣預報陣列
 * @returns {string} 格式化後的天氣報告文字
 */
function formatWeatherReport(cityName, forecasts) {
  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

  var report = '🌤 天氣預報 — ' + cityName + '\n';
  report += '━━━━━━━━━━━━━━\n';
  report += '📅 ' + now + '\n\n';

  for (var i = 0; i < forecasts.length; i++) {
    var f = forecasts[i];

    // 格式化時段（只取時間部分）
    var startTime = f.startTime.split(' ')[1].substring(0, 5);
    var endTime = f.endTime.split(' ')[1].substring(0, 5);
    var startDate = f.startTime.split(' ')[0].substring(5); // MM-DD

    // 天氣 emoji
    var weatherEmoji = getWeatherEmoji(f.weather);

    report += '🕐 ' + startDate + ' ' + startTime + ' ~ ' + endTime + '\n';
    report += '   ' + weatherEmoji + ' ' + f.weather + '\n';
    report += '   🌡 ' + f.minTemp + '°C ~ ' + f.maxTemp + '°C\n';
    report += '   🌧 降雨機率：' + f.rainProb + '%\n';
    report += '   😊 ' + f.comfort + '\n\n';
  }

  // 加入提醒
  var firstForecast = forecasts[0];
  var rainProb = parseInt(firstForecast.rainProb);
  if (rainProb >= 50) {
    report += '☂️ 提醒：降雨機率較高，記得帶傘！\n';
  }

  var maxTemp = parseInt(firstForecast.maxTemp);
  if (maxTemp >= 35) {
    report += '🥵 提醒：高溫警報，注意防曬補水！\n';
  } else if (parseInt(firstForecast.minTemp) <= 15) {
    report += '🧥 提醒：天氣較涼，注意保暖！\n';
  }

  return report;
}

/**
 * 根據天氣狀況回傳對應的 emoji
 */
function getWeatherEmoji(weather) {
  if (weather.indexOf('晴') !== -1 && weather.indexOf('雲') === -1) return '☀️';
  if (weather.indexOf('晴') !== -1 && weather.indexOf('雲') !== -1) return '🌤';
  if (weather.indexOf('多雲') !== -1) return '⛅';
  if (weather.indexOf('陰') !== -1) return '☁️';
  if (weather.indexOf('雨') !== -1) return '🌧';
  if (weather.indexOf('雷') !== -1) return '⛈';
  if (weather.indexOf('雪') !== -1) return '🌨';
  if (weather.indexOf('霧') !== -1) return '🌫';
  return '🌈';
}

// ============================================================
// 第四部分：推播天氣預報
// ============================================================

/**
 * 主要執行函式 — 取得天氣 → 格式化 → 推播到 Line
 * 設定排程觸發條件時，選擇此函式
 */
function sendWeatherReport() {
  // 取得天氣資料
  var locationData = getWeatherData(TARGET_CITY);

  if (!locationData) {
    pushText(LINE_USER_ID, '❌ 天氣預報取得失敗\n\n可能原因：\n1. API 授權碼不正確\n2. 縣市名稱不正確\n3. API 服務暫時無法使用');
    return;
  }

  // 解析天氣資料
  var forecasts = parseWeatherData(locationData);

  if (forecasts.length === 0) {
    pushText(LINE_USER_ID, '❌ 天氣預報解析失敗：沒有預報資料');
    return;
  }

  // 格式化報告
  var report = formatWeatherReport(TARGET_CITY, forecasts);

  // 推播到 Line
  pushText(LINE_USER_ID, report);

  // 記錄到試算表
  logWeatherToSheet(TARGET_CITY, forecasts);

  Logger.log('天氣預報已推播');
}

/**
 * 查詢多個縣市的天氣（進階功能）
 */
function sendMultiCityWeather() {
  var cities = ['臺北市', '臺中市', '高雄市'];

  for (var i = 0; i < cities.length; i++) {
    var locationData = getWeatherData(cities[i]);
    if (locationData) {
      var forecasts = parseWeatherData(locationData);
      var report = formatWeatherReport(cities[i], forecasts);
      pushText(LINE_USER_ID, report);

      // 避免短時間內送太多訊息
      Utilities.sleep(1000);
    }
  }
}

// ============================================================
// 第五部分：記錄到試算表
// ============================================================

/**
 * 將天氣資料記錄到試算表（方便日後回顧）
 */
function logWeatherToSheet(cityName, forecasts) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('天氣記錄');

    // 如果工作表不存在，建立新的
    if (!sheet) {
      sheet = ss.insertSheet('天氣記錄');
      var headers = ['記錄時間', '縣市', '時段', '天氣', '最低溫', '最高溫', '降雨機率', '舒適度'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#0D47A1')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    // 寫入資料
    var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

    for (var i = 0; i < forecasts.length; i++) {
      var f = forecasts[i];
      sheet.appendRow([
        now,
        cityName,
        f.startTime + ' ~ ' + f.endTime,
        f.weather,
        f.minTemp,
        f.maxTemp,
        f.rainProb + '%',
        f.comfort
      ]);
    }
  } catch (error) {
    Logger.log('記錄天氣資料失敗：' + error);
  }
}

// ========== Line Bot 推播函式（共用模組）==========

function pushText(to, text) {
  pushLine(to, [{ type: 'text', text: text }]);
}

function pushLine(to, messages) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({ to: to, messages: messages }),
    muteHttpExceptions: true
  });
}

// ========== 測試函式 ==========

/**
 * 測試 API 連線（先確認 API 授權碼是否正確）
 */
function testApiConnection() {
  var url = 'https://opendata.cwa.gov.tw/api/v1/rest/datastore/' + DATASET_ID;
  url += '?Authorization=' + CWA_API_KEY;
  url += '&locationName=' + encodeURIComponent(TARGET_CITY);

  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var code = response.getResponseCode();

  Logger.log('API 狀態碼：' + code);

  if (code === 200) {
    var json = JSON.parse(response.getContentText());
    Logger.log('API 連線成功！資料筆數：' + json.records.location.length);
    pushText(LINE_USER_ID, '✅ 氣象 API 連線測試成功！\n\n已成功取得 ' + TARGET_CITY + ' 的天氣資料。');
  } else {
    Logger.log('API 連線失敗：' + response.getContentText());
    pushText(LINE_USER_ID, '❌ 氣象 API 連線失敗\n狀態碼：' + code + '\n\n請確認 API 授權碼是否正確。');
  }
}
