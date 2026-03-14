/**
 * ============================================================
 * 實作 7：低軌衛星氣象 Line Bot
 * ============================================================
 *
 * 功能說明：
 *   1. 互動式選單 — 使用者輸入「低軌衛星」開啟主選單
 *   2. 天氣查詢 — 串接 OpenWeatherMap API（免費）
 *   3. 洋流查詢 — 串接 Stormglass API（免費）
 *   4. 衛星雲圖 — 提供即時衛星雲圖連結
 *   5. 查詢記錄 — 自動記錄到 Google 試算表
 *
 * 使用的 API：
 *   - OpenWeatherMap（免費，每分鐘 60 次）https://openweathermap.org/api
 *   - Stormglass（免費，每天 10 次）https://stormglass.io/
 *   - Line Messaging API
 *
 * 設定步驟：
 *   1. 註冊 OpenWeatherMap 帳號，取得 API Key
 *   2. 註冊 Stormglass 帳號，取得 API Key
 *   3. 填入下方 4 個設定值
 *   4. 部署為「網頁應用程式」
 *   5. 把部署網址貼到 Line Developers 的 Webhook URL
 *
 * 操作方式：
 *   使用者在 Line 輸入「低軌衛星」→ 顯示選單
 *   輸入 1~5 選擇功能 → 輸入 1~9 選擇地區 → 回傳資料
 *
 * ============================================================
 */

// ========== 設定區（只需改這 4 個）==========

var LINE_TOKEN = '在此貼上你的 Line Channel Access Token';
var OPENWEATHER_API_KEY = '在此貼上你的 OpenWeatherMap API Key';
var STORMGLASS_API_KEY = '在此貼上你的 Stormglass API Key';
var SHEET_ID = '在此貼上你的 Google Sheet ID';

// ========== 工作表名稱（程式會自動建立）==========

var LOG_SHEET = '查詢記錄';
var WEATHER_SHEET = '天氣資料';
var OCEAN_SHEET = '海洋資料';
var SESSION_SHEET = '使用者狀態';

// ========== 預設查詢地區 ==========

var LOCATIONS = {
  '1': { name: '台北', lat: 25.0330, lon: 121.5654 },
  '2': { name: '台中', lat: 24.1477, lon: 120.6736 },
  '3': { name: '高雄', lat: 22.6273, lon: 120.3014 },
  '4': { name: '花蓮', lat: 23.9871, lon: 121.6011 },
  '5': { name: '澎湖', lat: 23.5711, lon: 119.5793 },
  '6': { name: '金門', lat: 24.4893, lon: 118.3713 },
  '7': { name: '台東', lat: 22.7583, lon: 121.1444 },
  '8': { name: '墾丁', lat: 21.9500, lon: 120.8000 }
};

// ============================================================
// Line Bot Webhook 主程式
// ============================================================

/**
 * 處理 Line Webhook POST 請求
 */
function doPost(e) {
  try {
    var events = JSON.parse(e.postData.contents).events;

    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      if (event.type === 'message' && event.message.type === 'text') {
        handleMessage(event);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    logError('doPost', error);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 處理 GET 請求（用於測試部署是否成功）
 */
function doGet(e) {
  return ContentService.createTextOutput('🛰️ 低軌衛星氣象 Line Bot 運作中！')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ============================================================
// 訊息處理邏輯
// ============================================================

/**
 * 處理使用者訊息
 */
function handleMessage(event) {
  var userMessage = event.message.text.trim();
  var userId = event.source.userId;
  var replyToken = event.replyToken;

  // 記錄查詢
  logQuery(userId, userMessage);

  // 取得使用者目前狀態
  var userState = getUserState(userId);

  var replyMessages = [];

  // 判斷輸入
  if (userMessage === '低軌衛星' || userMessage === '低歸衛星' || userMessage.toLowerCase() === 'menu') {
    // 顯示主選單
    replyMessages = getMainMenu();
    setUserState(userId, 'MAIN_MENU', null);
  }
  else if (userState.state === 'MAIN_MENU' && ['1', '2', '3', '4', '5'].indexOf(userMessage) !== -1) {
    // 使用者選擇了功能
    replyMessages = handleMainMenuSelection(userId, userMessage);
  }
  else if (userState.state === 'SELECT_LOCATION') {
    // 使用者選擇了地區
    replyMessages = handleLocationSelection(userId, userMessage, userState.action);
  }
  else if (userState.state === 'CUSTOM_INPUT') {
    // 使用者輸入自訂地點
    replyMessages = handleCustomInput(userId, userMessage, userState.action);
  }
  else if (userMessage === '取消' || userMessage === '返回') {
    replyMessages = getMainMenu();
    setUserState(userId, 'MAIN_MENU', null);
  }
  else {
    replyMessages = getWelcomeMessage();
  }

  // 回覆訊息
  replyToLine(replyToken, replyMessages);
}

// ============================================================
// 選單系統
// ============================================================

/**
 * 主選單
 */
function getMainMenu() {
  var text = '🛰️ 低軌衛星氣象資料系統\n'
    + '━━━━━━━━━━━━━━━━━━\n'
    + '請輸入數字選擇功能：\n\n'
    + '1️⃣  查詢天氣狀況\n'
    + '2️⃣  查詢海洋/洋流資料\n'
    + '3️⃣  取得衛星雲圖連結\n'
    + '4️⃣  查詢所有資料\n'
    + '5️⃣  顯示使用說明\n\n'
    + '━━━━━━━━━━━━━━━━━━\n'
    + '💡 輸入「取消」可隨時返回此選單';

  return [{ type: 'text', text: text }];
}

/**
 * 歡迎訊息
 */
function getWelcomeMessage() {
  var text = '🛰️ 歡迎使用低軌衛星氣象系統！\n\n'
    + '請輸入【低軌衛星】開始使用\n\n'
    + '━━━━━━━━━━━━━━━━━━\n'
    + '可用指令：\n'
    + '• 低軌衛星 - 開啟主選單\n'
    + '• 取消 - 返回主選單';

  return [{ type: 'text', text: text }];
}

/**
 * 地區選單
 */
function getLocationMenu(action) {
  var actionName = '資料';
  if (action === 'weather') actionName = '天氣';
  if (action === 'ocean') actionName = '洋流';
  if (action === 'all') actionName = '所有資料';

  var text = '📍 請選擇要查詢' + actionName + '的地區：\n\n'
    + '1️⃣  台北\n'
    + '2️⃣  台中\n'
    + '3️⃣  高雄\n'
    + '4️⃣  花蓮\n'
    + '5️⃣  澎湖\n'
    + '6️⃣  金門\n'
    + '7️⃣  台東\n'
    + '8️⃣  墾丁\n'
    + '9️⃣  自行輸入地點\n\n'
    + '━━━━━━━━━━━━━━━━━━\n'
    + '💡 輸入數字選擇，或輸入「取消」返回';

  return [{ type: 'text', text: text }];
}

/**
 * 處理主選單選擇
 */
function handleMainMenuSelection(userId, selection) {
  switch (selection) {
    case '1':
      setUserState(userId, 'SELECT_LOCATION', 'weather');
      return getLocationMenu('weather');

    case '2':
      setUserState(userId, 'SELECT_LOCATION', 'ocean');
      return getLocationMenu('ocean');

    case '3':
      setUserState(userId, 'MAIN_MENU', null);
      return getSatelliteResponse();

    case '4':
      setUserState(userId, 'SELECT_LOCATION', 'all');
      return getLocationMenu('all');

    case '5':
      setUserState(userId, 'MAIN_MENU', null);
      return getHelpResponse();

    default:
      return getMainMenu();
  }
}

/**
 * 處理地區選擇
 */
function handleLocationSelection(userId, selection, action) {
  // 選擇 9：自行輸入地點
  if (selection === '9') {
    setUserState(userId, 'CUSTOM_INPUT', action);
    return [{
      type: 'text',
      text: '✏️ 請輸入要查詢的地點名稱\n\n'
        + '範例：\n'
        + '• 新竹\n'
        + '• 宜蘭\n'
        + '• Tokyo\n'
        + '• New York\n\n'
        + '━━━━━━━━━━━━━━━━━━\n'
        + '💡 輸入地點名稱，或輸入「取消」返回主選單'
    }];
  }

  var location = LOCATIONS[selection];

  if (!location) {
    return [{
      type: 'text',
      text: '❌ 無效的選項，請輸入 1-9 的數字，或輸入「取消」返回主選單'
    }];
  }

  // 重置狀態
  setUserState(userId, 'MAIN_MENU', null);

  // 根據 action 執行查詢
  switch (action) {
    case 'weather':
      return getWeatherResponse(location.lat, location.lon, location.name);

    case 'ocean':
      return getOceanResponse(location.lat, location.lon, location.name);

    case 'all':
      var weatherMsg = getWeatherResponse(location.lat, location.lon, location.name);
      var oceanMsg = getOceanResponse(location.lat, location.lon, location.name);
      var satelliteMsg = getSatelliteResponse();
      return weatherMsg.concat(oceanMsg).concat(satelliteMsg);

    default:
      return getMainMenu();
  }
}

/**
 * 處理自訂地點輸入
 */
function handleCustomInput(userId, locationName, action) {
  var geoUrl = 'https://api.openweathermap.org/geo/1.0/direct?q='
    + encodeURIComponent(locationName)
    + '&limit=1&appid=' + OPENWEATHER_API_KEY;

  try {
    var response = UrlFetchApp.fetch(geoUrl);
    var data = JSON.parse(response.getContentText());

    if (!data || data.length === 0) {
      return [{
        type: 'text',
        text: '❌ 找不到「' + locationName + '」這個地點\n\n'
          + '請嘗試：\n'
          + '• 使用中文或英文城市名\n'
          + '• 檢查拼寫是否正確\n\n'
          + '💡 輸入「低軌衛星」返回主選單'
      }];
    }

    var lat = data[0].lat;
    var lon = data[0].lon;
    var name = (data[0].local_names && data[0].local_names.zh) || data[0].name;

    setUserState(userId, 'MAIN_MENU', null);

    switch (action) {
      case 'weather':
        return getWeatherResponse(lat, lon, name);
      case 'ocean':
        return getOceanResponse(lat, lon, name);
      case 'all':
        var w = getWeatherResponse(lat, lon, name);
        var o = getOceanResponse(lat, lon, name);
        var s = getSatelliteResponse();
        return w.concat(o).concat(s);
      default:
        return getMainMenu();
    }
  } catch (error) {
    logError('handleCustomInput', error);
    return [{
      type: 'text',
      text: '❌ 查詢「' + locationName + '」時發生錯誤，請稍後再試。\n\n💡 輸入「低軌衛星」返回主選單'
    }];
  }
}

// ============================================================
// 使用者狀態管理（用試算表儲存）
// ============================================================

/**
 * 取得使用者狀態
 */
function getUserState(userId) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SESSION_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(SESSION_SHEET);
      sheet.appendRow(['使用者ID', '狀態', '動作', '更新時間']);
      return { state: 'NONE', action: null };
    }

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        // 檢查是否過期（30 分鐘）
        var updateTime = new Date(data[i][3]);
        var now = new Date();
        if ((now - updateTime) > 30 * 60 * 1000) {
          return { state: 'NONE', action: null };
        }
        return { state: data[i][1], action: data[i][2] };
      }
    }

    return { state: 'NONE', action: null };
  } catch (error) {
    logError('getUserState', error);
    return { state: 'NONE', action: null };
  }
}

/**
 * 設定使用者狀態
 */
function setUserState(userId, state, action) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SESSION_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(SESSION_SHEET);
      sheet.appendRow(['使用者ID', '狀態', '動作', '更新時間']);
    }

    var data = sheet.getDataRange().getValues();
    var found = false;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        sheet.getRange(i + 1, 2).setValue(state);
        sheet.getRange(i + 1, 3).setValue(action || '');
        sheet.getRange(i + 1, 4).setValue(new Date());
        found = true;
        break;
      }
    }

    if (!found) {
      sheet.appendRow([userId, state, action || '', new Date()]);
    }
  } catch (error) {
    logError('setUserState', error);
  }
}

// ============================================================
// 天氣 API（OpenWeatherMap）
// ============================================================

/**
 * 取得天氣資料
 */
function getWeatherData(lat, lon) {
  var url = 'https://api.openweathermap.org/data/2.5/weather?lat=' + lat
    + '&lon=' + lon
    + '&appid=' + OPENWEATHER_API_KEY
    + '&units=metric&lang=zh_tw';

  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var code = response.getResponseCode();

    if (code !== 200) {
      Logger.log('OpenWeatherMap API 錯誤，狀態碼：' + code);
      return null;
    }

    var data = JSON.parse(response.getContentText());

    // 儲存到試算表
    saveWeatherData(data);

    return {
      location: data.name,
      description: data.weather[0].description,
      temperature: data.main.temp,
      feelsLike: data.main.feels_like,
      humidity: data.main.humidity,
      pressure: data.main.pressure,
      windSpeed: data.wind.speed,
      windDeg: data.wind.deg,
      clouds: data.clouds.all,
      visibility: data.visibility,
      sunrise: new Date(data.sys.sunrise * 1000).toLocaleTimeString('zh-TW'),
      sunset: new Date(data.sys.sunset * 1000).toLocaleTimeString('zh-TW')
    };
  } catch (error) {
    logError('getWeatherData', error);
    return null;
  }
}

/**
 * 產生天氣回覆訊息
 */
function getWeatherResponse(lat, lon, locationName) {
  var d = getWeatherData(lat, lon);

  if (!d) {
    return [{ type: 'text', text: '❌ 取得天氣資料失敗，請稍後再試。' }];
  }

  var text = '🌤️ ' + locationName + ' 天氣資訊\n'
    + '━━━━━━━━━━━━━━━━━━\n'
    + '🌡️ 溫度：' + d.temperature + '°C\n'
    + '🤔 體感：' + d.feelsLike + '°C\n'
    + '☁️ 天氣：' + d.description + '\n'
    + '💧 濕度：' + d.humidity + '%\n'
    + '🌬️ 風速：' + d.windSpeed + ' m/s\n'
    + '🔭 能見度：' + d.visibility + 'm\n'
    + '☁️ 雲量：' + d.clouds + '%\n'
    + '🌅 日出：' + d.sunrise + '\n'
    + '🌇 日落：' + d.sunset + '\n'
    + '━━━━━━━━━━━━━━━━━━\n'
    + '📡 資料來源：OpenWeatherMap\n'
    + '⏰ 更新時間：' + new Date().toLocaleString('zh-TW') + '\n\n'
    + '💡 輸入「低軌衛星」返回主選單';

  return [{ type: 'text', text: text }];
}

// ============================================================
// 洋流 / 海洋 API（Stormglass）
// ============================================================

/**
 * 取得海洋資料
 */
function getOceanData(lat, lon) {
  var url = 'https://api.stormglass.io/v2/weather/point?lat=' + lat
    + '&lng=' + lon
    + '&params=waveHeight,wavePeriod,waveDirection,waterTemperature,currentSpeed,currentDirection,seaLevel';

  try {
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': STORMGLASS_API_KEY },
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code !== 200) {
      Logger.log('Stormglass API 錯誤，狀態碼：' + code);
      return null;
    }

    var data = JSON.parse(response.getContentText());
    var current = data.hours[0];

    // 儲存到試算表
    saveOceanData(current);

    return {
      waveHeight: (current.waveHeight && current.waveHeight.sg) || 'N/A',
      wavePeriod: (current.wavePeriod && current.wavePeriod.sg) || 'N/A',
      waveDirection: (current.waveDirection && current.waveDirection.sg) || 'N/A',
      waterTemperature: (current.waterTemperature && current.waterTemperature.sg) || 'N/A',
      currentSpeed: (current.currentSpeed && current.currentSpeed.sg) || 'N/A',
      currentDirection: (current.currentDirection && current.currentDirection.sg) || 'N/A',
      seaLevel: (current.seaLevel && current.seaLevel.sg) || 'N/A',
      time: current.time
    };
  } catch (error) {
    logError('getOceanData', error);
    return null;
  }
}

/**
 * 產生海洋回覆訊息
 */
function getOceanResponse(lat, lon, locationName) {
  var d = getOceanData(lat, lon);

  if (!d) {
    return [{
      type: 'text',
      text: '❌ 取得 ' + locationName + ' 海洋資料失敗\n\n'
        + '可能原因：\n'
        + '• Stormglass 免費版每日限 10 次請求\n'
        + '• 網路連線問題\n\n'
        + '💡 輸入「低軌衛星」返回主選單'
    }];
  }

  var text = '🌊 ' + locationName + ' 海洋/洋流資訊\n'
    + '━━━━━━━━━━━━━━━━━━\n'
    + '🌊 浪高：' + d.waveHeight + ' m\n'
    + '⏱️ 週期：' + d.wavePeriod + ' s\n'
    + '➡️ 浪向：' + d.waveDirection + '°\n'
    + '🌡️ 水溫：' + d.waterTemperature + '°C\n'
    + '💨 洋流速度：' + d.currentSpeed + ' m/s\n'
    + '🧭 洋流方向：' + d.currentDirection + '°\n'
    + '📊 海平面：' + d.seaLevel + ' m\n'
    + '━━━━━━━━━━━━━━━━━━\n'
    + '📡 資料來源：Stormglass\n'
    + '⏰ 資料時間：' + d.time + '\n\n'
    + '💡 輸入「低軌衛星」返回主選單';

  return [{ type: 'text', text: text }];
}

// ============================================================
// 衛星雲圖
// ============================================================

/**
 * 產生衛星雲圖回覆訊息
 */
function getSatelliteResponse() {
  var text = '🛰️ 衛星雲圖連結\n'
    + '━━━━━━━━━━━━━━━━━━\n'
    + '以下是即時衛星雲圖：\n\n'
    + '📡 Windy 衛星雲圖（互動式）\n'
    + '👉 https://www.windy.com/satellite\n\n'
    + '📡 向日葵衛星（日本氣象廳）\n'
    + '👉 https://www.jma.go.jp/bosai/map.html#5/25/121/&elem=ir&contents=himawari\n\n'
    + '📡 Zoom Earth（全球衛星）\n'
    + '👉 https://zoom.earth/\n\n'
    + '📡 NASA Worldview\n'
    + '👉 https://worldview.earthdata.nasa.gov/\n\n'
    + '━━━━━━━━━━━━━━━━━━\n'
    + '⏰ 查詢時間：' + new Date().toLocaleString('zh-TW') + '\n\n'
    + '💡 輸入「低軌衛星」返回主選單';

  return [{ type: 'text', text: text }];
}

// ============================================================
// 使用說明
// ============================================================

function getHelpResponse() {
  var text = '📖 使用說明\n'
    + '━━━━━━━━━━━━━━━━━━\n\n'
    + '【如何使用】\n'
    + '1️⃣ 輸入「低軌衛星」開啟選單\n'
    + '2️⃣ 輸入數字選擇功能\n'
    + '3️⃣ 選擇查詢地區（或自行輸入）\n'
    + '4️⃣ 系統回傳資料\n\n'
    + '【預設地區】\n'
    + '台北、台中、高雄、花蓮\n'
    + '澎湖、金門、台東、墾丁\n\n'
    + '【自訂地點】\n'
    + '選擇「9. 自行輸入」後\n'
    + '可輸入任何城市名稱查詢\n\n'
    + '【資料來源】\n'
    + '🌤️ 天氣：OpenWeatherMap\n'
    + '🌊 洋流：Stormglass\n'
    + '🛰️ 衛星：向日葵衛星/NASA\n\n'
    + '【注意事項】\n'
    + '• 洋流資料每日限 10 次查詢\n'
    + '• 衛星雲圖為外部連結\n'
    + '• 輸入「取消」可返回主選單\n\n'
    + '━━━━━━━━━━━━━━━━━━\n'
    + '💡 輸入「低軌衛星」返回主選單';

  return [{ type: 'text', text: text }];
}

// ============================================================
// Line API 功能
// ============================================================

/**
 * 回覆 Line 訊息（Webhook 回覆）
 */
function replyToLine(replyToken, messages) {
  var url = 'https://api.line.me/v2/bot/message/reply';

  var payload = {
    replyToken: replyToken,
    messages: messages.slice(0, 5)  // Line 限制最多 5 則
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + LINE_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    UrlFetchApp.fetch(url, options);
  } catch (error) {
    logError('replyToLine', error);
  }
}

/**
 * 主動推播訊息（給測試用）
 */
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

// ============================================================
// Google Sheet 記錄功能
// ============================================================

/**
 * 記錄查詢
 */
function logQuery(userId, query) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(LOG_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(LOG_SHEET);
      sheet.appendRow(['時間戳記', '使用者ID', '查詢內容']);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold')
        .setBackground('#1b5e20').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([new Date(), userId, query]);
  } catch (error) {
    logError('logQuery', error);
  }
}

/**
 * 儲存天氣資料
 */
function saveWeatherData(data) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(WEATHER_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(WEATHER_SHEET);
      sheet.appendRow(['時間戳記', '位置', '溫度', '體感溫度', '濕度', '風速', '天氣描述']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold')
        .setBackground('#0D47A1').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(), data.name, data.main.temp, data.main.feels_like,
      data.main.humidity, data.wind.speed, data.weather[0].description
    ]);
  } catch (error) {
    logError('saveWeatherData', error);
  }
}

/**
 * 儲存海洋資料
 */
function saveOceanData(data) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(OCEAN_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(OCEAN_SHEET);
      sheet.appendRow(['時間戳記', '浪高', '週期', '水溫', '洋流速度', '洋流方向']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold')
        .setBackground('#01579B').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      (data.waveHeight && data.waveHeight.sg) || 'N/A',
      (data.wavePeriod && data.wavePeriod.sg) || 'N/A',
      (data.waterTemperature && data.waterTemperature.sg) || 'N/A',
      (data.currentSpeed && data.currentSpeed.sg) || 'N/A',
      (data.currentDirection && data.currentDirection.sg) || 'N/A'
    ]);
  } catch (error) {
    logError('saveOceanData', error);
  }
}

/**
 * 記錄錯誤
 */
function logError(functionName, error) {
  Logger.log('[' + functionName + '] ' + error.toString());

  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('錯誤記錄');

    if (!sheet) {
      sheet = ss.insertSheet('錯誤記錄');
      sheet.appendRow(['時間戳記', '函式名稱', '錯誤訊息']);
    }

    sheet.appendRow([new Date(), functionName, error.toString()]);
  } catch (e) {
    Logger.log('無法記錄錯誤到 Sheet');
  }
}

// ============================================================
// 測試函式
// ============================================================

/**
 * 測試天氣 API — 查詢台北天氣
 */
function testWeatherAPI() {
  var result = getWeatherData(25.0330, 121.5654);
  if (result) {
    Logger.log('✅ 天氣 API 測試成功！');
    Logger.log('地點：' + result.location);
    Logger.log('溫度：' + result.temperature + '°C');
    Logger.log('天氣：' + result.description);
  } else {
    Logger.log('❌ 天氣 API 測試失敗，請檢查 OPENWEATHER_API_KEY');
  }
}

/**
 * 測試海洋 API — 查詢台北海洋資料
 */
function testOceanAPI() {
  var result = getOceanData(25.0330, 121.5654);
  if (result) {
    Logger.log('✅ 海洋 API 測試成功！');
    Logger.log('浪高：' + result.waveHeight + ' m');
    Logger.log('水溫：' + result.waterTemperature + '°C');
  } else {
    Logger.log('❌ 海洋 API 測試失敗，請檢查 STORMGLASS_API_KEY');
  }
}

/**
 * 清除過期的使用者狀態（可設排程每小時執行）
 */
function cleanupExpiredSessions() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SESSION_SHEET);
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    var now = new Date();

    for (var i = data.length - 1; i >= 1; i--) {
      var updateTime = new Date(data[i][3]);
      if ((now - updateTime) > 60 * 60 * 1000) {
        sheet.deleteRow(i + 1);
      }
    }

    Logger.log('✅ 已清除過期的使用者狀態');
  } catch (error) {
    logError('cleanupExpiredSessions', error);
  }
}
