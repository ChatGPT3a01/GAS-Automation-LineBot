/**
 * ============================================================
 * 實作 9：RSS 新聞推播機器人 → Line Bot 每日新聞
 * ============================================================
 *
 * 功能說明：
 *   1. 用 UrlFetchApp 抓取 RSS XML 資料
 *   2. 用 XmlService.parse() 解析 XML 格式
 *   3. 從 RSS 中提取新聞標題、連結、發布時間
 *   4. 格式化成易讀的新聞摘要
 *   5. 透過 Line Bot 推播每日新聞
 *   6. 將新聞記錄到試算表（避免重複推播）
 *
 * 事前準備：
 *   1. 取得 Line Bot 的 Channel Access Token
 *   2. 取得你的 Line User ID
 *   3. 在 Google 試算表的綁定專案中執行（用於記錄已推播的新聞）
 *
 * 設定排程自動執行：
 *   1. 在 Apps Script 編輯器中，點選左側「觸發條件」（鬧鐘圖示）
 *   2. 點選「新增觸發條件」
 *   3. 選擇函式：sendDailyNews
 *   4. 事件來源：時間驅動
 *   5. 時間型觸發條件類型：每日計時器
 *   6. 時段：上午 8 點到 9 點
 *   7. 儲存
 * ============================================================
 */

// ========== Line Bot 設定 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';

// ========== RSS 來源設定 ==========
// 可以自行新增或修改 RSS 來源
var RSS_FEEDS = [
  { name: '中央社',     url: 'https://feeds.feedburner.com/caboraa' },
  { name: '科技新報',   url: 'https://technews.tw/feed/' },
  { name: '數位時代',   url: 'https://www.bnext.com.tw/rss' }
];

// 每個來源取幾則新聞
var NEWS_COUNT = 3;

// ============================================================
// 第一部分：Line Bot 推播函式
// ============================================================

/**
 * 推播文字訊息
 */
function pushText(to, text) {
  pushLine(to, [{ type: 'text', text: text }]);
}

/**
 * 推播底層函式
 */
function pushLine(to, messages) {
  var url = 'https://api.line.me/v2/bot/message/push';
  var options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({
      to: to,
      messages: messages
    }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();

  if (code !== 200) {
    Logger.log('推播失敗，狀態碼：' + code);
    Logger.log('錯誤內容：' + response.getContentText());
  } else {
    Logger.log('推播成功！');
  }
}

// ============================================================
// 第二部分：試算表記錄與去重
// ============================================================

/**
 * 檢查某則新聞是否已經推播過
 * @param {string} title - 新聞標題
 * @returns {boolean} 是否已推播
 */
function isAlreadyPushed(title) {
  var sheet = getLogSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === title) return true;
  }
  return false;
}

/**
 * 將已推播的新聞記錄到試算表
 * @param {Object[]} newsItems - 新聞項目陣列
 */
function logPushedNews(newsItems) {
  var sheet = getLogSheet();
  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  for (var i = 0; i < newsItems.length; i++) {
    sheet.appendRow([now, newsItems[i].title, newsItems[i].link, newsItems[i].source]);
  }
}

/**
 * 取得或建立新聞記錄工作表
 */
function getLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('推播記錄');
  if (!sheet) {
    sheet = ss.insertSheet('推播記錄');
    sheet.getRange(1, 1, 1, 4).setValues([['推播時間', '標題', '連結', '來源']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================================
// 第三部分：RSS 解析
// ============================================================

/**
 * 抓取並解析 RSS Feed
 *
 * @param {string} feedUrl - RSS Feed 的網址
 * @returns {Object[]} 新聞項目陣列 [{title, link, pubDate}]
 */
function fetchRSS(feedUrl) {
  try {
    // 呼叫 RSS 網址
    var response = UrlFetchApp.fetch(feedUrl, {
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('RSS 抓取失敗：' + feedUrl + '（狀態碼：' + response.getResponseCode() + '）');
      return [];
    }

    // 解析 XML
    var xml = XmlService.parse(response.getContentText());
    var root = xml.getRootElement();

    // RSS 2.0 格式：root > channel > item
    // Atom 格式：root > entry
    var items = [];
    var namespace = root.getNamespace();

    // 嘗試 RSS 2.0 格式
    var channel = root.getChild('channel');
    if (channel) {
      var rssItems = channel.getChildren('item');
      for (var i = 0; i < Math.min(rssItems.length, NEWS_COUNT); i++) {
        var item = rssItems[i];
        items.push({
          title: item.getChildText('title') || '無標題',
          link: item.getChildText('link') || '',
          pubDate: item.getChildText('pubDate') || ''
        });
      }
    }

    // 嘗試 Atom 格式
    if (items.length === 0) {
      var entries = root.getChildren('entry', namespace);
      for (var j = 0; j < Math.min(entries.length, NEWS_COUNT); j++) {
        var entry = entries[j];
        var link = '';
        var links = entry.getChildren('link', namespace);
        for (var k = 0; k < links.length; k++) {
          if (links[k].getAttribute('rel') &&
              links[k].getAttribute('rel').getValue() === 'alternate') {
            link = links[k].getAttribute('href').getValue();
            break;
          }
          if (links[k].getAttribute('href')) {
            link = links[k].getAttribute('href').getValue();
          }
        }
        items.push({
          title: entry.getChildText('title', namespace) || '無標題',
          link: link,
          pubDate: entry.getChildText('published', namespace) ||
                   entry.getChildText('updated', namespace) || ''
        });
      }
    }

    Logger.log('從 ' + feedUrl + ' 取得 ' + items.length + ' 則新聞');
    return items;

  } catch (error) {
    Logger.log('解析 RSS 發生錯誤：' + error.message);
    return [];
  }
}

/**
 * 格式化發布時間
 *
 * @param {string} dateStr - 原始日期字串
 * @returns {string} 格式化後的時間
 */
function formatDate(dateStr) {
  if (!dateStr) return '';
  try {
    var date = new Date(dateStr);
    return Utilities.formatDate(date, 'Asia/Taipei', 'MM/dd HH:mm');
  } catch (e) {
    return dateStr.substring(0, 16);
  }
}

// ============================================================
// 第三部分：組合新聞訊息
// ============================================================

/**
 * 從所有 RSS 來源取得新聞並組合成推播訊息
 * 會自動跳過已推播過的新聞（去重）
 *
 * @returns {Object} { message: 格式化訊息, newItems: 新推播的項目陣列 }
 */
function buildNewsMessage() {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd (EEE)');

  var message = '📰 每日新聞摘要\n';
  message += '📅 ' + dateStr + '\n';
  message += '━━━━━━━━━━━━━━━\n';

  var hasNews = false;
  var newItems = [];  // 記錄本次新推播的項目

  for (var i = 0; i < RSS_FEEDS.length; i++) {
    var feed = RSS_FEEDS[i];
    Logger.log('正在抓取：' + feed.name + '（' + feed.url + '）');

    var items = fetchRSS(feed.url);
    var feedNewItems = [];

    // 過濾已推播過的新聞
    for (var j = 0; j < items.length; j++) {
      if (!isAlreadyPushed(items[j].title)) {
        feedNewItems.push(items[j]);
      }
    }

    if (feedNewItems.length > 0) {
      hasNews = true;
      message += '\n【' + feed.name + '】\n';

      for (var k = 0; k < feedNewItems.length; k++) {
        var item = feedNewItems[k];
        var time = formatDate(item.pubDate);
        message += '\n' + (k + 1) + '. ' + item.title;
        if (time) message += '\n   🕐 ' + time;
        if (item.link) message += '\n   🔗 ' + item.link;
        message += '\n';

        newItems.push({ title: item.title, link: item.link, source: feed.name });
      }
    }
  }

  if (!hasNews) {
    message += '\n目前沒有新的新聞，或所有新聞都已推播過。';
  }

  message += '\n━━━━━━━━━━━━━━━';
  message += '\n💡 由 GAS 自動推播';

  return { message: message, newItems: newItems };
}

// ============================================================
// 第四部分：主要功能 — 發送每日新聞
// ============================================================

/**
 * 【主函式】發送每日新聞到 Line
 * 排程觸發時執行這個函式
 */
function sendDailyNews() {
  Logger.log('===== 開始推播每日新聞 =====');

  // 組合新聞訊息（含去重）
  var result = buildNewsMessage();
  var newsMessage = result.message;

  // Line Bot 單則訊息上限 5000 字
  // 如果超過，只取前 4900 字
  if (newsMessage.length > 4900) {
    newsMessage = newsMessage.substring(0, 4900) + '\n\n...（訊息過長，已截斷）';
  }

  // 推播到 Line
  pushText(LINE_USER_ID, newsMessage);

  // 記錄已推播的新聞到試算表（避免下次重複推播）
  if (result.newItems.length > 0) {
    logPushedNews(result.newItems);
    Logger.log('已記錄 ' + result.newItems.length + ' 則新聞到試算表');
  }

  Logger.log('===== 每日新聞推播完成 =====');
}

// ============================================================
// 第五部分：測試函式
// ============================================================

/**
 * 測試：抓取單一 RSS 來源
 */
function testFetchSingleRSS() {
  var items = fetchRSS(RSS_FEEDS[0].url);
  for (var i = 0; i < items.length; i++) {
    Logger.log((i + 1) + '. ' + items[i].title);
    Logger.log('   連結：' + items[i].link);
    Logger.log('   時間：' + items[i].pubDate);
  }
}

/**
 * 測試：組合新聞訊息（只在日誌中顯示，不推播）
 */
function testBuildMessage() {
  var result = buildNewsMessage();
  Logger.log('訊息長度：' + result.message.length + ' 字');
  Logger.log('新增項目：' + result.newItems.length + ' 則');
  Logger.log(result.message);
}

/**
 * 測試：發送每日新聞
 */
function testSendDailyNews() {
  sendDailyNews();
}

/**
 * 測試：推播文字
 */
function testPush() {
  pushText(LINE_USER_ID, '🧪 RSS 新聞推播 Line Bot 連線測試成功！');
}
