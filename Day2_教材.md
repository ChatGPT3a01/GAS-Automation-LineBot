# Day 2：GAS 進階應用、RSS 推播與 AI 整合

> **Google Apps Script 自動化教學 — 第二天**
>
> 適用對象：已完成 Day 1 的學員（具備 GAS 基礎 + Line Bot 推播能力）
> 教學時間：約 6 小時
> 核心目標：進階 Line Bot 推播技巧 + AI 整合，打造全自動化系統

---

## 教學時程表

| 時間 | 單元 | 時長 | Line Bot 技能 |
|------|------|------|---------------|
| 09:00 - 09:40 | 【實作 8】早安圖 → Line 圖片推播 | 40 min | **圖片推播** |
| 09:40 - 10:30 | 【實作 9】RSS 新聞 → Line 推播 | 50 min | RSS + 長文推播 |
| 10:30 - 10:40 | 休息 | 10 min | |
| 10:40 - 11:20 | 【實作 10】出缺席 → Line 群組推播 | 40 min | **群組推播** |
| 11:20 - 12:10 | 【實作 11】報名系統 → Line 即時通知 | 50 min | 事件驅動推播 |
| 12:10 - 13:10 | 午餐 | 60 min | |
| 13:10 - 14:00 | 【實作 12】查詢系統 → Line 互動查詢 | 50 min | **對話式查詢** |
| 14:00 - 14:50 | 【實作 13】AI 發文 → Line AI 推播 | 50 min | AI + 推播 |
| 14:50 - 15:00 | 休息 | 10 min | |
| 15:00 - 15:20 | Day 2 總結與延伸 | 20 min | |

---

## Day 1 回顧

在開始今天的進階實作之前，讓我們快速回顧 Day 1 學到的核心技能：

| 你已經會的 | 今天要學的新技能 |
|-----------|-----------------|
| `pushText()` 推播文字 | `pushImage()` **推播圖片** |
| 推播給自己（User ID） | 推播到 **Line 群組**（Group ID） |
| 手動觸發函式 | **事件驅動觸發**（表單提交 → 自動推播） |
| `UrlFetchApp` 呼叫 JSON API | `XmlService.parse()` **解析 RSS XML** |
| `SpreadsheetApp` 讀寫試算表 | `FormApp` **操作 Google 表單** |
| 基礎 JavaScript | 串接 **Gemini AI** API |

> 如果你的 Line Bot 還沒設定好（Channel Access Token、User ID），請先回 Day 1 的單元 0 完成設定。

---

## 實作 8：每天來張早安圖 → Line Bot 圖片推播（40 分鐘）

### 學習目標

- 學會 `DriveApp` 從 Google 雲端硬碟取得檔案
- 學會圖片公開分享連結的處理
- 學會隨機選取機制 `Math.random()`
- 學會 **Line Bot 圖片訊息推播**（Image Message type）
- 學會設定每天自動執行的排程

### 成果預覽

完成後你會擁有一個「每日早安圖推播系統」：
- 每天早上 7 點，Line Bot 自動從你的 Google Drive 隨機挑一張早安圖
- 推播到你的 Line，附帶早安問候語
- 完全自動化，不需要手動操作

📄 **程式碼檔案**：`GAS程式碼/08_早安圖_LineBot.gs`

---

### 步驟 1：準備早安圖資料夾

1. 打開 Chrome 瀏覽器，進入 **drive.google.com**
2. 點選左上角的 **「+新增」** 按鈕
3. 選擇 **「新資料夾」**
4. 輸入資料夾名稱：**「早安圖」**
5. 按下 **「建立」**

> 📸 **截圖 D2-01**：（Google Drive 建立新資料夾）
>
> 擷取位置：Google Drive 頁面，點選「+新增」後的下拉選單，顯示「新資料夾」選項
> 用途：讓學員知道在哪裡建立新資料夾

6. 進入剛才建立的「早安圖」資料夾
7. 上傳 3~5 張早安圖片（可以用手邊任何圖片）

> 💡 **小技巧**：你可以直接從電腦拖曳圖片到 Google Drive 的瀏覽器視窗中，就能上傳了。

> 📸 **截圖 D2-02**：（早安圖資料夾中已有圖片）
>
> 擷取位置：Google Drive 「早安圖」資料夾內，顯示已上傳的圖片縮圖
> 用途：讓學員確認圖片已正確上傳

### 步驟 2：取得資料夾 ID

1. 在「早安圖」資料夾頁面
2. 看瀏覽器的**網址列**
3. 網址看起來像這樣：`https://drive.google.com/drive/folders/1aBcDeFgHiJkLmNoPqRsTuVwXyZ`
4. **`folders/` 後面那串英數字**就是資料夾 ID
5. 把它複製下來，等一下程式碼要用

> 📸 **截圖 D2-03**：（瀏覽器網址列標示資料夾 ID 的位置）
>
> 擷取位置：瀏覽器網址列，用框線標出 folders/ 後面的 ID 部分
> 用途：讓學員知道從哪裡複製資料夾 ID

> ⚠️ **注意**：資料夾 ID 是一串像 `1aBcDeFgHiJkLmNoPqRsTuVwXyZ` 的文字，不是整個網址！

### 步驟 3：建立新的 Apps Script 專案

1. 回到你的 Google 試算表（可以新建一個，命名為 **「早安圖系統」**）
2. 點選 **「擴充功能」→「Apps Script」**
3. 刪掉預設的 `myFunction()` 程式碼
4. 把 `GAS程式碼/08_早安圖_LineBot.gs` 的**完整程式碼**貼上

> 📸 **截圖 D2-04**：（Apps Script 編輯器中貼上早安圖程式碼）
>
> 擷取位置：Apps Script 編輯器，顯示已貼上的早安圖程式碼（能看到開頭的註釋區塊）
> 用途：讓學員確認程式碼已正確貼上

### 步驟 4：修改設定值

在程式碼的最上方，找到這三個設定，填入你自己的值：

```javascript
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
var MORNING_FOLDER_ID = '在此貼上你的早安圖資料夾 ID';
```

把它們改成你自己的值（注意保留前後的**單引號**）。

### 步驟 5：理解程式碼重點

讓我們看看這支程式的核心概念：

**1. DriveApp — 操作 Google Drive**

```javascript
// 用資料夾 ID 取得資料夾物件
var folder = DriveApp.getFolderById(folderId);

// 取得資料夾中的所有檔案
var files = folder.getFiles();

// 逐一檢查每個檔案
while (files.hasNext()) {
  var file = files.next();
  var mimeType = file.getMimeType();  // 取得檔案類型
  // 判斷是不是圖片
  if (mimeType.indexOf('image/') === 0) {
    // 是圖片，加入清單
  }
}
```

> `DriveApp` 就是 GAS 用來操作 Google Drive 的工具。`getFolderById()` 可以用資料夾 ID 找到指定的資料夾。

**2. Math.random() — 隨機選取**

```javascript
// 產生 0 到 images.length-1 之間的隨機整數
var index = Math.floor(Math.random() * images.length);
var selected = images[index];
```

> `Math.random()` 會產生 0~1 之間的隨機小數。乘以陣列長度再取整數，就能隨機選中一個項目。

**3. pushImage() — 圖片推播**

```javascript
function pushImage(to, imageUrl) {
  pushLine(to, [{
    type: 'image',                    // 訊息類型改成 image
    originalContentUrl: imageUrl,     // 原始圖片網址
    previewImageUrl: imageUrl         // 預覽圖網址
  }]);
}
```

> 跟 Day 1 的 `pushText()` 不同，`pushImage()` 的 `type` 是 `'image'`，需要提供圖片的公開網址。

**4. 設定圖片為公開**

```javascript
// 讓任何有連結的人都可以查看這張圖片
file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

// 產生直接存取的圖片網址
var imageUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
```

> Line Bot 要能顯示圖片，圖片網址必須是**公開可存取**的。所以我們用 `setSharing()` 把圖片設為公開。

### 步驟 6：測試早安圖推播

1. 在上方的函式選擇下拉選單中，選擇 **`testSendMorningImage`**
2. 點選 **▶ 執行**
3. 如果跳出授權視窗，按照 Day 1 的方式完成授權（審查權限 → 進階 → 前往 → 允許）

> ⚠️ **注意**：因為這次程式要存取 Google Drive，所以授權視窗會顯示「查看、編輯及刪除 Google 雲端硬碟中的檔案」，這是正常的！

> 📸 **截圖 D2-05**：（DriveApp 授權畫面）
>
> 擷取位置：Google 授權視窗，顯示需要的權限清單（含 Google Drive 存取）
> 用途：讓學員知道授權視窗會出現哪些權限要求

4. 等待幾秒鐘...
5. 拿起手機，打開 Line

> 📸 **截圖 D2-06**：（手機 Line 收到早安圖推播的畫面）
>
> 擷取位置：手機 Line 聊天畫面，顯示 Bot 傳來的早安圖和問候文字
> 用途：讓學員看到成功的推播結果

6. 你應該會在 Line 上看到：
   - 一張隨機選中的早安圖
   - 一則附帶日期的早安問候訊息

### 步驟 7：設定每天自動執行

現在讓我們設定排程，讓程式每天早上自動執行：

1. 在 Apps Script 編輯器中，點選左側的 **鬧鐘圖示**（觸發條件）
2. 點選右下角的 **「+ 新增觸發條件」**

> 📸 **截圖 D2-07**：（觸發條件設定頁面）
>
> 擷取位置：Apps Script 觸發條件頁面，標示「+ 新增觸發條件」按鈕位置
> 用途：讓學員知道在哪裡新增觸發條件

3. 在彈出的設定視窗中：
   - 選擇要執行的函式：**`sendMorningImage`**
   - 事件來源：**時間驅動**
   - 時間型觸發條件類型：**每日計時器**
   - 時段：**上午 7 點到 8 點**

> 📸 **截圖 D2-08**：（觸發條件設定細節）
>
> 擷取位置：觸發條件設定對話框，已填好上述設定的畫面
> 用途：讓學員確認每個欄位的設定值

4. 點選 **「儲存」**

> 🎉 **完成！** 從明天開始，每天早上 7~8 點，你的 Line 就會自動收到一張隨機的早安圖！

---

## 實作 9：RSS 新聞推播機器人 → Line Bot 每日新聞（50 分鐘）

### 學習目標

- 認識 RSS（Really Simple Syndication）網站內容訂閱格式
- 學會 `UrlFetchApp` 抓取 RSS XML 資料
- 學會 `XmlService.parse()` 解析 XML 格式
- 學會從 RSS 中提取標題、連結、發布時間
- 學會長文字的格式化與推播

### 成果預覽

完成後你會擁有一個「每日新聞推播機器人」：
- 每天早上 8 點，自動抓取多個新聞來源的最新新聞
- 整理成易讀的新聞摘要
- 透過 Line Bot 推播到你的手機

📄 **程式碼檔案**：`GAS程式碼/09_RSS新聞推播_LineBot.gs`

---

### 步驟 1：認識 RSS

**什麼是 RSS？**

RSS（Really Simple Syndication）是一種讓網站提供內容更新的標準格式。你可以把它想像成「新聞的宅配服務」：

- 以前：你要主動去每個新聞網站看有沒有新文章
- 現在：RSS 就像訂報紙，新文章出來就自動送到你面前

RSS 的資料格式是 **XML**（一種跟 HTML 類似的標記語言）。看起來像這樣：

```xml
<rss version="2.0">
  <channel>
    <title>科技新報</title>
    <item>
      <title>AI 最新發展趨勢</title>
      <link>https://technews.tw/2026/03/...</link>
      <pubDate>Mon, 09 Mar 2026 08:00:00 +0800</pubDate>
    </item>
    <item>
      <title>另一則新聞標題</title>
      ...
    </item>
  </channel>
</rss>
```

> 每個 `<item>` 就是一則新聞，裡面有 `<title>`（標題）、`<link>`（連結）、`<pubDate>`（發布時間）。

### 步驟 2：建立 Apps Script 專案

1. 新建一個 Google 試算表，命名為 **「新聞推播系統」**
2. 點選 **「擴充功能」→「Apps Script」**
3. 刪掉預設程式碼
4. 貼上 `GAS程式碼/09_RSS新聞推播_LineBot.gs` 的完整程式碼

### 步驟 3：修改設定值

```javascript
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
```

> RSS 來源已經預設了幾個台灣的新聞網站，你也可以自行修改 `RSS_FEEDS` 陣列。

### 步驟 4：理解程式碼重點

**1. UrlFetchApp.fetch() — 抓取 RSS 資料**

```javascript
var response = UrlFetchApp.fetch(feedUrl, {
  muteHttpExceptions: true,
  followRedirects: true
});
```

> 這跟 Day 1 呼叫天氣 API 的方式一樣。不同的是，RSS 回傳的不是 JSON 而是 XML。

**2. XmlService.parse() — 解析 XML**

```javascript
// 把文字解析成 XML 物件
var xml = XmlService.parse(response.getContentText());

// 取得根元素
var root = xml.getRootElement();

// 取得 channel 下的所有 item
var channel = root.getChild('channel');
var items = channel.getChildren('item');

// 取得每個 item 的標題和連結
var title = items[0].getChildText('title');
var link = items[0].getChildText('link');
```

> `XmlService` 是 GAS 解析 XML 的工具。它把 XML 文字轉換成一個有階層結構的物件，你可以用 `getChild()`、`getChildText()` 等方法來取得需要的資料。

> 📸 **截圖 D2-09**：（XML 結構說明圖）
>
> 擷取位置：教材製作的 XML 結構示意圖，用縮排展示 rss > channel > item > title/link/pubDate 的層級關係
> 用途：讓學員理解 XML 的階層結構

**3. 長文字的格式化**

```javascript
var message = '📰 每日新聞摘要\n';
message += '📅 ' + dateStr + '\n';
message += '━━━━━━━━━━━━━━━\n';  // 分隔線
```

> 在 Line 推播中，`\n` 代表換行。善用 emoji 和分隔線可以讓長文字更易讀。

### 步驟 5：測試 RSS 抓取

1. 在函式下拉選單選擇 **`testFetchSingleRSS`**
2. 點選 **▶ 執行**
3. 查看**執行記錄**

> 📸 **截圖 D2-10**：（RSS 抓取結果的執行記錄）
>
> 擷取位置：Apps Script 執行記錄，顯示成功抓取到的新聞標題和連結
> 用途：讓學員看到 RSS 抓取成功的結果

4. 你應該會看到類似這樣的輸出：
```
1. AI 最新發展趨勢
   連結：https://technews.tw/2026/03/...
   時間：Mon, 09 Mar 2026 08:00:00 +0800
```

### 步驟 6：測試完整新聞推播

1. 在函式下拉選單選擇 **`testSendDailyNews`**
2. 點選 **▶ 執行**
3. 等待幾秒（需要抓取多個 RSS 來源）
4. 查看手機 Line

> 📸 **截圖 D2-11**：（手機 Line 收到新聞推播的畫面）
>
> 擷取位置：手機 Line 聊天畫面，顯示格式化的每日新聞摘要
> 用途：讓學員看到新聞推播的預期結果

### 步驟 7：設定每天自動推播

1. 點選左側 **鬧鐘圖示**（觸發條件）
2. 點選 **「+ 新增觸發條件」**
3. 設定：
   - 函式：**`sendDailyNews`**
   - 事件來源：**時間驅動**
   - 類型：**每日計時器**
   - 時段：**上午 8 點到 9 點**
4. 儲存

> 🎉 **完成！** 從明天開始，你每天早上都會在 Line 上收到最新的新聞摘要！

### 延伸挑戰

如果你想加入更多 RSS 來源，只要在 `RSS_FEEDS` 陣列中新增即可：

```javascript
var RSS_FEEDS = [
  { name: '中央社',     url: 'https://feeds.feedburner.com/caboraa' },
  { name: '科技新報',   url: 'https://technews.tw/feed/' },
  { name: '數位時代',   url: 'https://www.bnext.com.tw/rss' },
  // 新增你想要的 RSS 來源：
  { name: '你的來源名稱', url: 'RSS 網址' }
];
```

> 💡 **如何找到 RSS 網址？** 大部分新聞網站都有 RSS。通常在網站底部或「訂閱」頁面可以找到，或直接 Google 搜尋「XXX網站 RSS」。

---

## 實作 10：到班出缺席通知 → Line Bot 群組推播（40 分鐘）

### 學習目標

- 學會試算表資料的篩選與比對邏輯
- 學會 **Line Bot 群組推播**（push to groupId）
- 學會條件式通知：只在有缺席時才推播
- 學會資料統計與格式化

### 成果預覽

完成後你會擁有一個「出缺席自動通知系統」：
- 在試算表中記錄學生出缺席狀況
- GAS 每天下午自動檢查有沒有缺席
- **只在有缺席時**才推播通知到 Line 群組
- 包含統計資訊（出席率、缺席次數）

📄 **程式碼檔案**：`GAS程式碼/10_出缺席通知_LineBot.gs`

---

### 步驟 1：建立出缺席試算表

1. 建立新的 Google 試算表，命名為 **「出缺席記錄」**
2. 取得試算表 ID（網址列 `/d/` 和 `/edit` 之間的那串）

> 📸 **截圖 D2-12**：（瀏覽器網址列標示試算表 ID）
>
> 擷取位置：瀏覽器網址列，用框線標出試算表 ID 的位置
> 用途：讓學員知道從哪裡複製試算表 ID

### 步驟 2：建立 Apps Script 專案

1. 在試算表中，點選 **「擴充功能」→「Apps Script」**
2. 刪掉預設程式碼
3. 貼上 `GAS程式碼/10_出缺席通知_LineBot.gs` 的完整程式碼

### 步驟 3：修改設定值

```javascript
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
var SPREADSHEET_ID = '在此貼上你的試算表 ID';
```

> 💡 **關於群組推播**：如果你想推播到 Line 群組，把 `PUSH_TARGET` 改成群組 ID。取得群組 ID 的方法請看程式碼註釋。目前先用自己的 User ID 測試。

### 步驟 4：理解程式碼重點

**1. 個人推播 vs 群組推播**

```javascript
// 推播給個人（User ID 以 U 開頭）
pushText('U1234567890abcdef...', '訊息內容');

// 推播到群組（Group ID 以 C 開頭）
pushText('C1234567890abcdef...', '訊息內容');
```

> `pushText()` 函式的第一個參數可以是 User ID 或 Group ID。推播到群組時，群組中所有成員都會看到訊息。

> 📸 **截圖 D2-13**：（個人推播 vs 群組推播示意圖）
>
> 擷取位置：教材製作的示意圖，左邊顯示個人推播（一對一），右邊顯示群組推播（一對多）
> 用途：讓學員理解兩種推播模式的差異

**2. 條件式觸發 — 只在有缺席時才通知**

```javascript
function sendAttendanceReport() {
  var message = buildAttendanceMessage();

  if (message) {
    // 有缺席或請假，才推播通知
    pushText(PUSH_TARGET, message);
  } else {
    // 全員出席，不推播（省錢又不打擾人）
    Logger.log('全員出席，不推播');
  }
}
```

> 這就是「條件式觸發」的概念：排程每天都會執行，但只有在有缺席的情況下才發出通知。避免每天都收到「全員出席」的無意義訊息。

**3. 日期比對邏輯**

```javascript
// 取得今天的日期字串
var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');

// 比對試算表中的日期
if (recordDate === today) {
  // 這是今天的記錄
}
```

### 步驟 5：新增測試資料

1. 在函式下拉選單選擇 **`testAddSampleData`**
2. 點選 **▶ 執行**
3. 完成授權後，回到試算表查看

> 📸 **截圖 D2-14**：（試算表中的出缺席測試資料）
>
> 擷取位置：Google 試算表，顯示自動建立的「出缺席記錄」工作表，含 5 筆測試資料
> 用途：讓學員看到測試資料的格式

4. 你應該會看到一個新的「出缺席記錄」工作表，裡面有 5 筆測試資料

### 步驟 6：測試出缺席通知

1. 在函式下拉選單選擇 **`testSendReport`**
2. 點選 **▶ 執行**
3. 查看手機 Line

> 📸 **截圖 D2-15**：（手機 Line 收到出缺席通知）
>
> 擷取位置：手機 Line 聊天畫面，顯示格式化的出缺席通知（含統計、缺席名單、請假名單）
> 用途：讓學員看到出缺席通知的預期結果

### 步驟 7：設定每天自動檢查

1. 點選左側 **鬧鐘圖示**
2. 新增觸發條件：
   - 函式：**`sendAttendanceReport`**
   - 時間驅動 → 每日計時器 → **下午 4 點到 5 點**
3. 儲存

### 如何取得 Line 群組 ID

如果你要推播到 Line 群組，需要取得群組 ID：

1. 先把 Line Bot 加入目標群組
2. 使用 Day 1 的「實作 2：訊息收集器」記錄群組訊息
3. 在試算表中找到群組的 `groupId`（以 C 開頭的一串文字）
4. 把 `PUSH_TARGET` 改成這個 Group ID

```javascript
// 改成群組推播
var PUSH_TARGET = 'C1234567890abcdef1234567890abcdef';
```

---

## 實作 11：自動化活動報名系統 → Line Bot 即時通知（50 分鐘）

### 學習目標

- 學會 `FormApp` 操作 Google 表單
- 學會 `onFormSubmit(e)` 表單提交觸發
- 學會報名人數管控（額滿自動關閉）
- 學會**事件驅動推播**（表單提交 → 即時通知）

### 成果預覽

完成後你會擁有一個「自動化活動報名系統」：
- 學員填寫 Google 表單報名
- 每次有人報名 → Line Bot **即時推播**通知你
- 報名額滿 → 表單**自動關閉** + Line Bot 通知

📄 **程式碼檔案**：`GAS程式碼/11_活動報名系統_LineBot.gs`

---

### 步驟 1：建立 Google 表單

有兩種方式建立表單：

**方式 A：手動建立（推薦，學習用）**

1. 打開 Chrome，進入 **forms.google.com**
2. 點選 **「+」空白表單**
3. 填入表單標題：**「GAS 自動化工作坊 — 報名表」**
4. 新增以下問題：
   - **姓名**（簡答，必填）
   - **Email**（簡答，必填）
   - **聯絡電話**（簡答，必填）
   - **用餐需求**（單選：葷食 / 素食 / 不需要，必填）
   - **備註**（段落，選填）

> 📸 **截圖 D2-16**：（Google 表單編輯畫面）
>
> 擷取位置：Google 表單編輯器，顯示已建立的報名表單（含上述問題）
> 用途：讓學員看到表單的預期樣式

5. 點選上方的 **「回覆」** 分頁
6. 點選 **試算表圖示**（「在試算表中查看」）
7. 選擇 **「建立新試算表」**
8. 這樣表單的回覆就會自動存到試算表

> 📸 **截圖 D2-17**：（表單連結試算表的設定）
>
> 擷取位置：表單「回覆」分頁，標示試算表圖示的位置
> 用途：讓學員知道如何連結表單和試算表

**方式 B：用程式自動建立**

也可以執行 `createRegistrationForm` 函式，程式會自動幫你建立表單。

### 步驟 2：取得表單 ID

1. 在表單編輯頁面
2. 看瀏覽器網址列：`https://docs.google.com/forms/d/XXXXXXXXX/edit`
3. `/d/` 和 `/edit` 之間的那串就是表單 ID

### 步驟 3：建立 Apps Script 專案

1. 在**表單連結的那個試算表**中，點選 **「擴充功能」→「Apps Script」**
2. 貼上 `GAS程式碼/11_活動報名系統_LineBot.gs` 的完整程式碼

> ⚠️ **重要**：Apps Script 要建在**表單連結的試算表**裡，不是建在表單裡！

### 步驟 4：修改設定值

```javascript
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
var FORM_ID = '在此貼上你的 Google 表單 ID';
var MAX_PARTICIPANTS = 30;  // 報名人數上限
var EVENT_NAME = 'GAS 自動化工作坊';  // 活動名稱
```

### 步驟 5：設定表單提交觸發條件

這是今天最重要的新概念：**事件驅動觸發**。

跟之前的「時間驅動」不同，事件驅動是：**有人做了某件事（填表單）→ 程式自動執行**。

1. 在 Apps Script 編輯器中，點選左側 **鬧鐘圖示**
2. 點選 **「+ 新增觸發條件」**
3. 設定：
   - 函式：**`onFormSubmit`**
   - 事件來源：**從試算表**
   - 事件類型：**提交表單時**

> 📸 **截圖 D2-18**：（onFormSubmit 觸發條件設定）
>
> 擷取位置：觸發條件設定對話框，事件來源選「從試算表」，事件類型選「提交表單時」
> 用途：讓學員看到事件驅動觸發的設定方式

4. 儲存

### 步驟 6：理解程式碼重點

**1. onFormSubmit(e) — 事件驅動函式**

```javascript
function onFormSubmit(e) {
  // e 是事件物件，包含了表單提交的資料
  var responses = e.namedValues;
  // responses 的格式：{ '姓名': ['王小明'], 'Email': ['test@gmail.com'] }
}
```

> 當有人填寫表單時，GAS 會自動呼叫 `onFormSubmit()`，並把填寫的資料放在 `e` 參數中。

**2. FormApp — 操控表單**

```javascript
// 用表單 ID 開啟表單
var form = FormApp.openById(FORM_ID);

// 關閉表單（不再接受回覆）
form.setAcceptingResponses(false);

// 設定關閉後的訊息
form.setCustomClosedFormMessage('報名已額滿！');

// 重新開啟表單
form.setAcceptingResponses(true);
```

> `FormApp` 可以程式化地控制 Google 表單。最實用的功能就是「額滿自動關閉」。

### 步驟 7：測試報名系統

1. 用手機或另一個瀏覽器視窗打開表單
2. 填寫一筆報名資料，送出

> 📸 **截圖 D2-19**：（填寫 Google 報名表單）
>
> 擷取位置：Google 表單填寫頁面（使用者視角）
> 用途：讓學員看到表單的填寫畫面

3. 等待幾秒鐘...
4. 查看手機 Line

> 📸 **截圖 D2-20**：（手機 Line 收到報名通知）
>
> 擷取位置：手機 Line 聊天畫面，顯示「收到新報名！」的通知（含報名人資料和目前人數）
> 用途：讓學員看到即時推播通知的結果

5. 你應該會在 Line 上看到：
   - 活動名稱
   - 報名者的資料
   - 目前報名人數 / 上限
   - 剩餘名額

> 🎉 **這就是事件驅動的威力！** 有人報名 → 你立刻收到通知，完全即時，不需要一直去看試算表。

---

## 實作 12：GAS 查詢系統 → Line Bot 互動查詢（50 分鐘）

### 學習目標

- 學會 `HtmlService` 進階用法：建立搜尋介面
- 學會前後端分離架構（HTML + GAS）
- 學會 `google.script.run` 前端呼叫後端
- 學會模糊搜尋邏輯
- 學會建立**雙入口系統**（Web App + Line Bot 都能查）

### 成果預覽

完成後你會擁有一個「雙入口查詢系統」：
- **入口 1**：用瀏覽器打開 Web App 網頁，在搜尋框輸入關鍵字查詢
- **入口 2**：在 Line 上傳「查 XX」給 Bot，Bot 回覆查詢結果
- 兩個入口查的是同一份試算表的資料

📄 **程式碼檔案**：`GAS程式碼/12_查詢系統_LineBot.gs` + `GAS程式碼/12_查詢系統_前端.html`

---

### 步驟 1：建立資料試算表

1. 建立新的 Google 試算表，命名為 **「查詢系統」**
2. 取得試算表 ID

### 步驟 2：建立 Apps Script 專案

1. 在試算表中，點選 **「擴充功能」→「Apps Script」**
2. 刪掉預設程式碼
3. 貼上 `GAS程式碼/12_查詢系統_LineBot.gs` 的完整程式碼

### 步驟 3：建立前端 HTML 檔案

1. 在 Apps Script 編輯器左側，點選 **「+」** 號
2. 選擇 **「HTML」**
3. 將檔案命名為 **「查詢介面」**（不需要加 .html）

> 📸 **截圖 D2-21**：（在 Apps Script 中新增 HTML 檔案）
>
> 擷取位置：Apps Script 編輯器左側，點「+」後選擇「HTML」的畫面
> 用途：讓學員知道如何建立 HTML 檔案

4. 把 `GAS程式碼/12_查詢系統_前端.html` 的完整內容貼上

> 📸 **截圖 D2-22**：（Apps Script 中有 Code.gs 和 查詢介面.html 兩個檔案）
>
> 擷取位置：Apps Script 編輯器左側檔案列表，顯示兩個檔案
> 用途：讓學員確認檔案結構正確

### 步驟 4：修改設定值

在 `Code.gs` 中修改：

```javascript
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
var SPREADSHEET_ID = '在此貼上你的試算表 ID';
```

### 步驟 5：建立範例資料

1. 在函式下拉選單選擇 **`initSampleData`**
2. 點選 **▶ 執行**（完成授權）
3. 回到試算表查看

> 📸 **截圖 D2-23**：（試算表中的範例查詢資料）
>
> 擷取位置：Google 試算表，顯示「人員資料」工作表，含 8 筆範例資料（姓名、部門、職稱、Email、分機）
> 用途：讓學員看到範例資料的格式

### 步驟 6：理解核心概念 — google.script.run

```javascript
// 前端 HTML 呼叫後端 GAS 函式
google.script.run
  .withSuccessHandler(showResults)    // 成功時執行 showResults()
  .withFailureHandler(showError)      // 失敗時執行 showError()
  .searchData(keyword);               // 呼叫後端的 searchData() 函式
```

> 📸 **截圖 D2-24**：（前後端架構示意圖）
>
> 擷取位置：教材製作的架構圖，顯示：使用者 → HTML 前端 → google.script.run → GAS 後端 → 試算表
> 用途：讓學員理解前後端分離的資料流

> `google.script.run` 是 GAS 獨有的魔法：讓前端網頁可以直接呼叫後端的 GAS 函式。不需要寫 API、不需要設定伺服器，Google 都幫你處理好了。

### 步驟 7：部署為 Web App

1. 點選右上角 **「部署」→「新增部署作業」**
2. 點選齒輪圖示，選擇 **「網頁應用程式」**
3. 設定：
   - 說明：查詢系統 v1
   - 執行身分：**我**
   - 存取權：**所有人**

> 📸 **截圖 D2-25**：（Web App 部署設定）
>
> 擷取位置：部署設定對話框，標示各個欄位的設定值
> 用途：讓學員正確設定部署參數

4. 點選 **「部署」**
5. 複製部署後的 **網頁應用程式網址**

> ⚠️ **注意**：這個網址就是你的 Web App 查詢系統！同時也是 Line Bot 的 Webhook URL。

### 步驟 8：測試 Web App 查詢

1. 用瀏覽器打開剛才複製的網址
2. 在搜尋框輸入 **「業務」**，點選搜尋

> 📸 **截圖 D2-26**：（Web App 查詢結果畫面）
>
> 擷取位置：瀏覽器中的查詢系統頁面，搜尋「業務」的結果
> 用途：讓學員看到 Web App 的查詢功能

### 步驟 9：設定 Line Bot Webhook

1. 到 **Line Developers Console**
2. 進入你的 Messaging API Channel
3. 找到 **Webhook URL** 欄位
4. 貼上剛才的 Web App 網址
5. 開啟 **Use webhook**

> 📸 **截圖 D2-27**：（Line Developers Console 的 Webhook 設定）
>
> 擷取位置：Line Developers Console，Messaging API 分頁的 Webhook URL 設定區域
> 用途：讓學員知道在哪裡設定 Webhook URL

### 步驟 10：測試 Line Bot 查詢

1. 在 Line 上傳訊息給你的 Bot：**「查 業務」**
2. Bot 應該會回覆查詢結果

> 📸 **截圖 D2-28**：（手機 Line 的查詢互動畫面）
>
> 擷取位置：手機 Line 聊天畫面，顯示使用者傳「查 業務」後 Bot 回覆查詢結果
> 用途：讓學員看到 Line Bot 查詢功能的實際效果

> 🎉 **雙入口系統完成！** 同一份資料，可以用網頁查，也可以用 Line 查！

---

## 實作 13：AI 自動發文 → Line Bot AI 推播（50 分鐘）

### 學習目標

- 學會串接 **Gemini API**（Google AI）
- 學會 **Prompt Engineering** 基礎（提示詞設計）
- 學會 AI 生成內容 → 自動寫入試算表
- 學會 AI 產出 → Line Bot 自動推播
- 學會排程實現**全自動化流程**

### 成果預覽

完成後你會擁有一個「AI 自動發文系統」：
- 每天自動隨機選一個主題
- 呼叫 Gemini AI 產生一篇社群貼文
- 貼文自動記錄到試算表
- 透過 Line Bot 推播 AI 寫的文章
- 全程自動，不需人工介入

📄 **程式碼檔案**：`GAS程式碼/13_AI自動發文_LineBot.gs`

---

### 步驟 1：申請 Gemini API Key

1. 打開 Chrome 瀏覽器
2. 進入 **https://aistudio.google.com/apikey**
3. 登入你的 Google 帳號

> 📸 **截圖 D2-29**：（Google AI Studio API Key 頁面）
>
> 擷取位置：Google AI Studio 的 API Key 管理頁面
> 用途：讓學員找到申請 API Key 的位置

4. 點選 **「Create API key」**（建立 API Key）
5. 選擇一個 Google Cloud 專案（或建立新的）
6. 等待 API Key 產生
7. 複製 API Key，存到記事本

> 📸 **截圖 D2-30**：（API Key 已產生的畫面）
>
> 擷取位置：Google AI Studio 顯示已建立的 API Key（部分遮蔽）
> 用途：讓學員看到 API Key 的格式

> ⚠️ **注意**：API Key 就像密碼，**千萬不要分享給別人**，也不要放在公開的程式碼中！

> 💡 **Gemini API 免費嗎？** 是的！`gemini-2.5-flash` 模型有免費額度，每天可以呼叫很多次。對於我們的教學用途綽綽有餘。

### 步驟 2：建立 Apps Script 專案

1. 建立新的 Google 試算表，命名為 **「AI 自動發文」**
2. 取得試算表 ID
3. 點選 **「擴充功能」→「Apps Script」**
4. 貼上 `GAS程式碼/13_AI自動發文_LineBot.gs` 的完整程式碼

### 步驟 3：修改設定值

```javascript
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
var GEMINI_API_KEY = '在此貼上你的 Gemini API Key';
var SPREADSHEET_ID = '在此貼上你的試算表 ID';
```

### 步驟 4：理解 Prompt Engineering

**什麼是 Prompt Engineering？**

Prompt（提示詞）就是你對 AI 說的話。Prompt Engineering 就是**如何把話說得好**，讓 AI 給出你想要的回答。

看看我們的 Prompt 設計：

```javascript
function buildPrompt(topic) {
  var prompt = '你是一位專業的社群內容創作者。\n\n';      // 角色設定
  prompt += '請幫我撰寫一篇關於「' + topic + '」的社群貼文。\n\n';  // 任務
  prompt += '要求：\n';
  prompt += '1. 標題：吸引人、簡潔有力（15字以內）\n';    // 具體要求
  prompt += '2. 內文：200~300字，口語化、易讀\n';
  prompt += '3. 包含 2~3 個實用的重點或建議\n';
  prompt += '4. 結尾加上 3~5 個相關的 hashtag\n';
  prompt += '5. 適當使用 emoji 讓文章更生動\n\n';
  prompt += '請用以下格式回覆：\n';                        // 輸出格式
  prompt += '【標題】你的標題\n';
  prompt += '【內文】你的內文\n';
  prompt += '【標籤】#tag1 #tag2 #tag3';
  return prompt;
}
```

> **好的 Prompt 包含四個要素：**
> 1. **角色設定**：你是誰？（專業的社群內容創作者）
> 2. **任務描述**：要做什麼？（撰寫社群貼文）
> 3. **具體要求**：有什麼限制？（字數、風格、內容要點）
> 4. **輸出格式**：回覆格式？（用【標題】【內文】【標籤】包裝）

### 步驟 5：理解 Gemini API 呼叫

```javascript
function callGemini(prompt) {
  // 1. 組合 API 網址
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/'
    + GEMINI_MODEL + ':generateContent?key=' + GEMINI_API_KEY;

  // 2. 準備請求內容
  var payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: 0.8,      // 創意度（0~1）
      maxOutputTokens: 1024  // 最大回覆長度
    }
  };

  // 3. 發送 API 請求
  var response = UrlFetchApp.fetch(url, options);

  // 4. 解析 AI 回覆
  var json = JSON.parse(response.getContentText());
  var text = json.candidates[0].content.parts[0].text;
  return text;
}
```

> 這跟 Day 1 呼叫天氣 API 很像，都是用 `UrlFetchApp.fetch()` 發送 HTTP 請求。差別是：
> - 天氣 API 用 `GET` 方法（取得資料）
> - Gemini API 用 `POST` 方法（送出 Prompt，取回 AI 回覆）

> 📸 **截圖 D2-31**：（Gemini API 呼叫流程圖）
>
> 擷取位置：教材製作的流程圖，顯示：GAS → Prompt → Gemini API → AI 回覆 → 解析 → 推播
> 用途：讓學員理解 AI API 的呼叫流程

### 步驟 6：測試 Gemini API

先測試 API Key 是否正確：

1. 在函式下拉選單選擇 **`testGemini`**
2. 點選 **▶ 執行**
3. 查看**執行記錄**

> 📸 **截圖 D2-32**：（Gemini API 測試結果）
>
> 擷取位置：Apps Script 執行記錄，顯示 Gemini 回覆的文字
> 用途：讓學員確認 API Key 設定正確

4. 你應該會在執行記錄中看到 Gemini 的回覆文字

> ⚠️ **如果出現錯誤**：
> - `400`：API Key 格式不正確
> - `403`：API Key 無效或未啟用
> - `429`：超過免費額度（等一下再試）

### 步驟 7：測試完整的 AI 發文流程

1. 在函式下拉選單選擇 **`testGenerateAndPost`**
2. 點選 **▶ 執行**
3. 等待幾秒鐘（AI 需要一點時間思考）
4. 查看手機 Line

> 📸 **截圖 D2-33**：（手機 Line 收到 AI 生成的貼文）
>
> 擷取位置：手機 Line 聊天畫面，顯示 AI 生成的社群貼文（含標題、內文、hashtag）
> 用途：讓學員看到 AI 自動發文的最終成果

5. 同時回到試算表，會看到新增了一筆記錄

> 📸 **截圖 D2-34**：（試算表中的 AI 發文記錄）
>
> 擷取位置：Google 試算表「AI 發文記錄」工作表，顯示日期、主題、標題、內文、標籤、狀態
> 用途：讓學員看到 AI 生成的文章也會記錄到試算表

### 步驟 8：設定每天自動發文

1. 點選左側 **鬧鐘圖示**
2. 新增觸發條件：
   - 函式：**`dailyAutoPost`**
   - 時間驅動 → 每日計時器 → **上午 9 點到 10 點**
3. 儲存

> 🎉 **恭喜！你已經打造了一個完整的 AI 自動化閉環：**
>
> ```
> 排程觸發 → 隨機選主題 → 呼叫 Gemini AI → 生成文章 → 記錄試算表 → Line Bot 推播
> ```
>
> 這就是 **AI + 排程 + Line Bot** 的全自動化系統！

### 延伸挑戰

**修改發文主題：**

```javascript
var TOPICS = [
  '科技新趨勢',
  '生活小技巧',
  // 加入你自己感興趣的主題：
  '烹飪食譜分享',
  '旅遊景點推薦',
  '理財投資入門'
];
```

**調整 AI 的創意度：**

```javascript
generationConfig: {
  temperature: 0.8,  // 0.0 = 很保守，1.0 = 很有創意
}
```

---

## Day 2 總結

### 今天學到的新技能

| 技能類別 | 內容 | 對應實作 |
|----------|------|----------|
| 圖片推播 | `pushImage()` + DriveApp | 實作 8 |
| XML 解析 | `XmlService.parse()` | 實作 9 |
| 群組推播 | push to Group ID | 實作 10 |
| 事件驅動 | `onFormSubmit()` | 實作 11 |
| 前後端分離 | `google.script.run` | 實作 12 |
| AI 整合 | Gemini API | 實作 13 |

### 兩天技能總覽

| Day | 實作 | GAS 服務 | Line Bot 技能 |
|-----|------|----------|---------------|
| 1 | 單元 0 | 環境建置 | 建立 Bot + 推播測試 |
| 1 | 實作 1 | doGet/doPost, SpreadsheetApp | Reply 回覆 |
| 1 | 實作 2 | UrlFetchApp, DriveApp | Webhook 接收 |
| 1 | 實作 3 | SpreadsheetApp | Push 文字 |
| 1 | 實作 4 | HtmlService | 事件推播 |
| 1 | 實作 5 | SpreadsheetApp | 確認推播 |
| 1 | 實作 6 | UrlFetchApp | 排程推播 |
| 1 | 實作 7 | DriveApp, Gemini API | AI + Push |
| 2 | 實作 8 | DriveApp, Math.random | **圖片推播** |
| 2 | 實作 9 | UrlFetchApp, XmlService | **RSS 長文推播** |
| 2 | 實作 10 | SpreadsheetApp | **群組推播** |
| 2 | 實作 11 | FormApp, onFormSubmit | **事件驅動推播** |
| 2 | 實作 12 | HtmlService, doPost | **對話式查詢** |
| 2 | 實作 13 | Gemini API | **AI + 排程推播** |

### 學員帶走的完整成品

**Day 1（8 個成品）：**
- Line Bot + 共用推播模組
- 知識庫查詢機器人
- 訊息收集器系統
- 第一個 GAS 自動化
- HTML 表單 + 通知
- 試算表寫入系統
- 天氣預報每日推播
- AI 自動化評量系統

**Day 2（6 個成品）：**
- 每日早安圖推播系統
- RSS 新聞推播機器人
- 出缺席自動通知系統
- 活動報名自動管理系統
- 雙入口查詢系統（Web + Line）
- AI 自動發文系統

### 延伸學習方向

若學員有興趣深入，可自行研究：

- **Flex Message** — Line Bot 卡片式豐富訊息（進階排版）
- **Line RAG 智慧助理** — 完整 AI 聊天機器人系統
- **空氣品質即時通** — 政府開放資料 + Line Bot 警示
- **簡報自動化生成** — Google Slides API
- **問卷自動化生成** — Google Forms API 進階
- **自主學習評語機器人** — AI + 教育應用
- **便利貼機器人** — Line Bot 進階互動
- **獎狀套印** — Google Docs 合併列印
