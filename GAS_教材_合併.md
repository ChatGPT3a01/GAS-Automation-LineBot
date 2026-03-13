# GAS 自動化 Line Bot 實戰教學

> **Google Apps Script 自動化教學 — 完整教材**
>
> 適用對象：完全零基礎的初學者
> 本教材包含 Part 1（基礎入門）與 Part 2（進階應用）兩大單元，共 14 個實作專案。

\newpage

## Part 1：GAS 基礎入門與 Line Bot 串接

> **Google Apps Script 自動化教學 — Part 1**
>
> 適用對象：完全零基礎的初學者
> 教學時間：約 6 小時
> 核心目標：用 GAS 自動化 Google 生態系，每個實作都串接 Line Bot

---

## 教學時程表

| 時間 | 單元 | 時長 | Line Bot 技能 |
|------|------|------|---------------|
| 09:00 - 09:40 | 【單元 0】GAS 環境 + Line Bot 建立 | 40 min | 建立 Bot + 推播測試 |
| 09:40 - 10:30 | 【實作 1】RAG 知識庫 → Line 查詢回覆 | 50 min | Reply 回覆 |
| 10:30 - 10:40 | 休息 | 10 min | |
| 10:40 - 11:30 | 【實作 2】Line Bot 訊息收集器 | 50 min | Webhook 接收 |
| 11:30 - 12:00 | 【實作 3】第一個 GAS → Line 通知 | 30 min | Push 文字 |
| 12:00 - 13:00 | 午餐 | 60 min | |
| 13:00 - 13:40 | 【實作 4】HTML 表單 → Line 通知 | 40 min | 事件推播 |
| 13:40 - 14:10 | 【實作 5】寫入試算表 → Line 回報 | 30 min | 確認推播 |
| 14:10 - 14:20 | 休息 | 10 min | |
| 14:20 - 15:00 | 【實作 6】天氣預報 → Line 推播 | 40 min | 排程推播 |
| 15:00 - 15:10 | Part 1 小結 | 10 min | |

---

## 前置準備

在開始今天的課程之前，請先確認你已經備妥以下三樣東西：

1. **Google 帳號** — 如果你有 Gmail，就已經有了。如果沒有，請到 [https://accounts.google.com/](https://accounts.google.com/) 免費註冊一個。
2. **手機安裝 Line** — 等一下我們會建立 Line Bot，需要在手機上接收測試訊息。請確認你的 Line 可以正常使用。
3. **Chrome 瀏覽器** — 雖然其他瀏覽器也能用，但 Google 的服務跟 Chrome 最搭配，建議使用 Chrome 來操作。如果還沒安裝，請到 [https://www.google.com/chrome/](https://www.google.com/chrome/) 下載。

> 以上三樣東西準備好了嗎？好，那我們開始吧！

---

## 單元 0：GAS 環境介紹 + Line Bot 建立（40 分鐘）

### Part A — 認識 Google Apps Script（15 分鐘）

#### 什麼是 Google Apps Script？

**Google Apps Script（簡稱 GAS）** 是 Google 提供的免費雲端程式工具。你可以把它想像成一個「超級助手」，它可以幫你自動操作 Google 試算表、Google 文件、Gmail、Google 日曆等所有 Google 的服務。

簡單來說：

- **GAS 是什麼？** → Google 送你的免費自動化工具
- **用什麼語言？** → JavaScript（全世界最多人用的程式語言之一）
- **在哪裡執行？** → 在 Google 的雲端伺服器上，不用安裝任何東西
- **要花錢嗎？** → 完全免費！

舉個例子：你每天要打開試算表，把當天的銷售數據整理好，然後寄 Email 給主管。用 GAS，你可以寫一段程式讓它每天自動幫你做這件事，而且還會用 Line 通知你「報表已送出」。

#### 步驟 1：開啟 Google 試算表

1. 打開 Chrome 瀏覽器
2. 在網址列輸入 **sheets.google.com**，按下 Enter
3. 如果還沒登入 Google 帳號，請先登入
4. 你會看到 Google 試算表的首頁

> 📸 **截圖 01**：（Google 試算表首頁）
>
> 擷取位置：瀏覽器開啟 sheets.google.com 後的畫面
> 用途：讓學員確認自己看到的畫面跟教材一樣

5. 點選左上角的 **「+」空白試算表**（或「建立新試算表」）
6. 一個全新的試算表就會打開
7. 在左上角把試算表名稱改成 **「GAS 練習」**（點一下「未命名的試算表」文字，輸入新名稱）

#### 步驟 2：開啟 Apps Script 編輯器

1. 在試算表的上方選單中，找到 **「擴充功能」**
2. 點一下「擴充功能」
3. 在下拉選單中，點選 **「Apps Script」**

> 📸 **截圖 02**：（擴充功能選單中的 Apps Script 按鈕）
>
> 擷取位置：Google 試算表上方選單列，「擴充功能」下拉選單展開的畫面
> 用途：讓學員找到 Apps Script 的入口位置

4. 瀏覽器會開啟一個新的分頁，這就是 **Apps Script 編輯器**
5. 等待頁面載入完成（大約 3~5 秒）

#### 步驟 3：認識 Apps Script 編輯器介面

編輯器打開後，你會看到一個看起來像程式開發工具的畫面。別緊張，我們一個一個認識：

> 📸 **截圖 03**：（Apps Script 編輯器介面，標註程式碼區、執行按鈕、日誌按鈕）
>
> 擷取位置：Apps Script 編輯器完整畫面，用箭頭標註以下區域：(A) 左側檔案列表、(B) 中間程式碼編輯區、(C) 上方工具列的執行按鈕（三角形）、(D) 上方工具列的日誌按鈕、(E) 左側選單的「觸發條件」鬧鐘圖示
> 用途：讓學員認識編輯器各個區域的功能

- **(A) 左側檔案列表**：顯示你的程式檔案。預設會有一個叫 `Code.gs` 的檔案
- **(B) 中間程式碼編輯區**：這是你寫程式的地方。預設已經有一個空的 `myFunction()` 函式
- **(C) 執行按鈕（▶）**：上方工具列的三角形按鈕，按下去就會執行你的程式
- **(D) 執行記錄**：看程式執行的結果和日誌訊息
- **(E) 觸發條件**：左側選單中的鬧鐘圖示，可以設定自動排程

#### 步驟 4：第一次執行程式 — Hello GAS!

現在我們來寫人生中第一行 GAS 程式碼！

1. 在程式碼編輯區，你會看到預設的程式碼：

```javascript
function myFunction() {

}
```

2. 把它**全部刪掉**（用鍵盤 `Ctrl + A` 全選，然後按 `Delete`）
3. 輸入以下程式碼：

```javascript
function myFunction() {
  Logger.log("Hello GAS！我的第一支程式！");
}
```

> **小知識**：`Logger.log()` 就像是在螢幕上印出文字的意思。它會把括號裡的文字記錄到「執行記錄」中，方便你確認程式有沒有正常運作。

4. 按下 **Ctrl + S** 儲存程式碼（或點選上方的磁碟片圖示）
5. 點選上方的 **▶ 執行** 按鈕

> ⚠️ **注意**：第一次執行時，Google 會跳出**授權視窗**，要求你允許程式存取你的 Google 帳號。這是正常的！
>
> 如果看到授權視窗：
> - 點選「審查權限」
> - 選擇你的 Google 帳號
> - 如果看到「這個應用程式未經 Google 驗證」，點選左下角的「進階」
> - 點選「前往 XXX（不安全）」
> - 點選「允許」
>
> 別擔心，因為這是你自己寫的程式，所以是安全的。

6. 授權完成後，程式會自動執行
7. 點選下方的 **「執行記錄」** 查看結果

> 📸 **截圖 04**：（Logger.log 執行結果的執行記錄畫面）
>
> 擷取位置：Apps Script 編輯器下方的「執行記錄」面板，顯示「Hello GAS！我的第一支程式！」的輸出結果
> 用途：讓學員確認程式有成功執行，並看到 Logger.log 的輸出

8. 你應該會在執行記錄中看到：
```
Hello GAS！我的第一支程式！
```

恭喜你！你已經成功執行了第一支 GAS 程式！🎉

#### 步驟 5：認識觸發條件與部署

在 GAS 中，有兩個非常重要的概念：

**觸發條件（Trigger）** — 讓程式自動執行

點選左側選單中的 **鬧鐘圖示**（觸發條件），你會看到觸發條件設定頁面。

> 📸 **截圖 05**：（觸發條件設定頁面）
>
> 擷取位置：點選 Apps Script 左側選單的鬧鐘圖示後，顯示的觸發條件頁面
> 用途：讓學員認識觸發條件頁面的位置和外觀

觸發條件的功能就像「鬧鐘」一樣：

- **時間驅動觸發**：每天早上 7 點自動執行（例如：每天推播天氣預報）
- **試算表觸發**：有人編輯試算表時自動執行
- **表單提交觸發**：有人填寫 Google 表單時自動執行

我們在後面的實作 6（天氣預報）會實際設定觸發條件。

**部署（Deploy）** — 讓程式變成一個網站

部署就是把你的 GAS 程式發布到網路上，讓其他人（或其他程式）可以透過網址存取它。這個功能在後面的實作 1 和實作 4 會用到。

#### GAS 的限制

最後，了解一下 GAS 有哪些使用限制（免費版）：

| 限制項目 | 限制值 |
|----------|--------|
| 每次執行時間上限 | **6 分鐘** |
| 每天觸發條件總執行時間 | 90 分鐘 |
| 每天 UrlFetchApp 呼叫次數 | 20,000 次 |
| 每天寄送 Email 數量 | 100 封 |
| 每個專案的觸發條件數量 | 20 個 |

> 對於一般使用來說，這些限制綽綽有餘，不用擔心！

---

### Part B — 建立 Line Bot（25 分鐘）

#### 為什麼要用 Line Bot？

在這門課程中，我們的設計理念是：

```
GAS 負責「做事」 ＋ Line Bot 負責「通知」
```

每個實作的最終結果，都會透過 Line Bot 推播到你的手機上。這樣一來，你不用一直盯著電腦螢幕，自動化的結果會自動送到手機的 Line 上。

> ⚠️ **注意**：**Line Notify 已於 2025 年 3 月 31 日停止服務**。以前很多教學都是用 Line Notify，但現在它已經不能用了。所以我們使用 **Line Bot（Messaging API）** 來代替，功能更強大！

#### 步驟 1：登入 Line Developers Console

1. 打開 Chrome 瀏覽器
2. 在網址列輸入：**https://developers.line.biz/**
3. 按下 Enter

> 📸 **截圖 06**：（Line Developers Console 登入頁面）
>
> 擷取位置：瀏覽器開啟 https://developers.line.biz/ 後的登入畫面
> 用途：讓學員確認進入的是正確的網站

4. 點選 **「Log in」** 按鈕（在頁面右上角）
5. 選擇用 **Line 帳號** 登入（如果有的話）或用 Email 登入
6. 輸入你的 Line 帳號密碼
7. 手機上的 Line 可能會跳出驗證碼確認，請在手機上輸入驗證碼
8. 登入成功後，你會進入 Line Developers Console 的主畫面

#### 步驟 2：建立 Provider

Provider 就像是一個「開發團隊」的概念。你可以把它想成一個資料夾，裡面放你建立的 Bot。

1. 在 Console 首頁，點選 **「Create a new provider」**（建立新的 Provider）
2. 輸入 Provider 名稱，例如：**「我的GAS專案」** 或你的名字
3. 點選 **「Create」**

> 📸 **截圖 07**：（建立 Provider 的畫面）
>
> 擷取位置：Line Developers Console 中，點選「Create a new provider」後出現的輸入欄位
> 用途：讓學員看到在哪裡輸入 Provider 名稱

#### 步驟 3：建立 Messaging API Channel

Channel 就是你的 Line Bot 本體。

1. 進入剛剛建立的 Provider
2. 點選 **「Create a Messaging API channel」**
3. 填寫以下資料：

> 📸 **截圖 08**：（建立 Messaging API Channel 的表單）
>
> 擷取位置：建立 Messaging API Channel 的完整表單畫面
> 用途：讓學員看到需要填寫哪些欄位

| 欄位 | 填寫內容 |
|------|----------|
| Channel type | Messaging API（預設） |
| Provider | 剛才建立的 Provider（預設） |
| Company or owner's country or region | **Taiwan** |
| Channel icon | 可以先不設定（之後再改） |
| Channel name | **我的GAS助手**（或你喜歡的名字） |
| Channel description | **GAS 自動化通知機器人** |
| Category | **Education**（教育） |
| Subcategory | 隨便選一個即可 |

4. 勾選底部的兩個同意條款
5. 點選 **「Create」** 按鈕
6. 在確認視窗中，點選 **「OK」**

> ⚠️ **注意**：Channel name 之後還可以修改，所以不用想太久。

#### 步驟 4：取得 Channel Access Token

**Channel Access Token** 是一組很長的密鑰，就像是 Bot 的「身分證」。GAS 程式要透過它來控制你的 Line Bot。

1. 建立 Channel 後，你會進入 Channel 的設定頁面
2. 點選上方的 **「Messaging API」** 分頁
3. 捲動到頁面最下方，找到 **「Channel access token (long-lived)」** 區域
4. 點選 **「Issue」** 按鈕

> 📸 **截圖 09**：（Channel Access Token 的 Issue 按鈕位置）
>
> 擷取位置：Messaging API 分頁最下方的 Channel access token 區域，標示 Issue 按鈕
> 用途：讓學員找到 Issue 按鈕的位置

5. 等待幾秒鐘，一組很長的文字就會出現
6. 點選旁邊的 **「複製」** 按鈕（或手動選取全部文字，按 Ctrl+C 複製）

> ⚠️ **注意**：請把這組 Token 先貼到記事本或其他地方保存好！後面每個實作都會用到。**千萬不要分享給別人**，因為有了這組 Token 就可以操控你的 Bot。

#### 步驟 5：取得 User ID

**User ID** 是你的 Line 帳號專屬 ID。GAS 要知道你的 User ID，才能把訊息推播給你。

1. 點選上方的 **「Basic settings」** 分頁
2. 捲動到頁面 **最下方**
3. 找到 **「Your user ID」** 欄位
4. 點選旁邊的複製按鈕，把 User ID 複製下來

> 📸 **截圖 10**：（User ID 在 Basic settings 的位置）
>
> 擷取位置：Basic settings 分頁最下方的 Your user ID 欄位
> 用途：讓學員找到 User ID 的位置

> ⚠️ **注意**：同樣把 User ID 也存到記事本裡。現在你的記事本上應該有兩組文字：
> - Channel Access Token（很長的一串）
> - User ID（以 U 開頭的一串）

#### 步驟 6：關閉自動回覆和打招呼訊息

Line Bot 預設會自動回覆「感謝加好友」之類的訊息，我們需要把它關掉，否則會跟我們自己的程式衝突。

1. 點選上方的 **「Messaging API」** 分頁
2. 找到 **「Auto-reply messages」** 那一列
3. 點選右邊的 **「Edit」** 連結（會跳轉到 Line Official Account Manager 頁面）
4. 在新頁面中，找到以下兩個設定，把它們都**關閉**：
   - **自動回應訊息**：關閉（切換到「停用」）
   - **加入好友的歡迎訊息**：關閉（切換到「停用」）

> 📸 **截圖 11**：（關閉 Auto-reply 和 Greeting messages 的設定）
>
> 擷取位置：Line Official Account Manager 的回應設定頁面，標示兩個需要關閉的開關
> 用途：讓學員看到哪些開關需要關閉

5. 設定完成後，**關閉這個分頁**，回到 Line Developers Console

#### 步驟 7：加 Bot 為好友

在測試之前，你需要先在手機的 Line 上加入這個 Bot 為好友。

1. 在 Line Developers Console 的 **「Messaging API」** 分頁
2. 往上捲動，找到 **「Bot information」** 區域
3. 你會看到一個 **QR Code**
4. 用手機的 Line 掃描這個 QR Code，加入好友

> ⚠️ **注意**：一定要先加好友！不然 Bot 沒辦法傳訊息給你。

#### 步驟 8：在 GAS 中建立共用推播模組

現在我們要回到 GAS，建立一個「共用推播模組」。這個模組包含推播功能的程式碼，之後每一個實作都會用到它。

📄 **程式碼檔案**：`GAS程式碼/00_LineBot_共用模組.gs`

1. 回到剛才的 Apps Script 編輯器（還記得那個「GAS 練習」的試算表嗎？）
2. 在左側的檔案列表中，你會看到 `Code.gs`
3. 點選 `Code.gs` 旁邊的 **「+」** 號
4. 選擇 **「指令碼」**

> 📸 **截圖 12**：（在 GAS 中新增檔案，點 + 號新增 LineBot.gs）
>
> 擷取位置：Apps Script 編輯器左側檔案列表，滑鼠點選「+」按鈕後出現的選單（顯示「指令碼」和「HTML」選項）
> 用途：讓學員知道怎麼新增程式碼檔案

5. 新檔案會以「未命名」出現，將它命名為 **LineBot**（不需要加 .gs，系統會自動加）
6. 按下 Enter 確認
7. 現在你有兩個檔案：`Code.gs` 和 `LineBot.gs`
8. 點選 **`LineBot.gs`**，把裡面預設的內容全部刪掉
9. 把以下程式碼完整貼上：

```javascript
/**
 * ============================================================
 * LineBot 共用推播模組
 * ============================================================
 * 所有實作都會使用的 Line Bot 推播功能
 */

// ========== 請修改以下兩個值 ==========
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';

// ========== 推播函式 ==========

/**
 * 推播文字訊息給指定對象
 */
function pushText(to, text) {
  pushLine(to, [{ type: 'text', text: text }]);
}

/**
 * 推播圖片訊息給指定對象
 */
function pushImage(to, imageUrl) {
  pushLine(to, [{
    type: 'image',
    originalContentUrl: imageUrl,
    previewImageUrl: imageUrl
  }]);
}

/**
 * 推播多則訊息（底層函式）
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

// ========== 回覆函式（用於 Webhook）==========

function replyText(replyToken, text) {
  replyLine(replyToken, [{ type: 'text', text: text }]);
}

function replyLine(replyToken, messages) {
  var url = 'https://api.line.me/v2/bot/message/reply';
  var options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: messages
    }),
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(url, options);
}

// ========== 測試函式 ==========

function testPush() {
  var now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  pushText(LINE_USER_ID, '🎉 Line Bot 連線測試成功！\n\n目前時間：' + now + '\n\n如果你看到這則訊息，表示 GAS 與 Line Bot 已成功串接。');
}
```

10. 在程式碼中，找到這兩行：

```javascript
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
```

11. 把 `'在此貼上你的 Channel Access Token'` 替換成你剛才複製的 Channel Access Token（注意**保留前後的單引號**）
12. 把 `'在此貼上你的 User ID'` 替換成你剛才複製的 User ID（注意**保留前後的單引號**）

例如：
```javascript
var LINE_TOKEN = 'aB3cD4eF5gH6iJ7kL8mN9oP0qR1sT2uV3wX4yZ...（這是一串很長的文字）';
var LINE_USER_ID = 'U1234567890abcdef1234567890abcdef';
```

> ⚠️ **注意**：Token 和 User ID 前後一定要有**單引號** `'...'`。沒有引號的話程式會出錯！

13. 按下 **Ctrl + S** 儲存

#### 步驟 9：測試推播

1. 在上方工具列中，找到**函式選擇下拉選單**（顯示「myFunction」的那個）
2. 點選下拉選單，選擇 **「testPush」**

> 📸 **截圖 13**：（執行 testPush 函式的下拉選單）
>
> 擷取位置：Apps Script 編輯器上方工具列，展開函式選擇下拉選單，顯示 testPush 選項
> 用途：讓學員知道如何選擇要執行的函式

3. 點選 **▶ 執行** 按鈕
4. 如果跳出授權視窗，請依照之前的步驟完成授權
5. 等待幾秒鐘...
6. 拿起手機，打開 Line

> 📸 **截圖 14**：（手機 Line 收到測試推播的畫面）
>
> 擷取位置：手機 Line 的聊天畫面，顯示 Bot 傳來的「Line Bot 連線測試成功！」訊息
> 用途：讓學員看到成功接收推播的預期結果

7. 你應該會在 Line 上看到一則來自你的 Bot 的訊息：「🎉 Line Bot 連線測試成功！」

> ⚠️ **如果沒收到訊息，請檢查以下幾點：**
> 1. Channel Access Token 是否正確貼上？（有沒有多空格或少字元？）
> 2. User ID 是否正確？（以大寫 U 開頭）
> 3. 是否已在手機上加 Bot 為好友？
> 4. 查看 Apps Script 的執行記錄，看有沒有錯誤訊息

**恭喜！單元 0 完成！** 你現在已經有一個可以運作的 Line Bot 了，接下來的每個實作都會用到它！

---

## 實作 1：RAG 知識庫 API → Line Bot 查詢回覆（50 分鐘）

### 學習目標

- 學會 `doGet()` / `doPost()` — GAS Web App 的核心入口
- 學會 `SpreadsheetApp` 讀寫試算表資料
- 學會 `ContentService` 回傳 JSON 格式
- 學會部署 GAS 為網頁應用程式
- 學會設定 Line Bot Webhook

### 成果預覽

完成後你會擁有一個「知識庫查詢機器人」：
- 用瀏覽器打開網址，就能查看知識庫的所有資料（REST API）
- 在 Line 上傳送關鍵字給 Bot，Bot 會自動搜尋知識庫並回覆結果

📄 **程式碼檔案**：`GAS程式碼/01_RAG知識庫_LineBot.gs`

---

### 步驟 1：建立新的 Google 試算表

1. 在 Chrome 瀏覽器中開啟新分頁
2. 輸入 **sheets.google.com**
3. 點選「+」建立一個新的空白試算表
4. 把試算表名稱改成 **「知識庫系統」**

### 步驟 2：開啟 Apps Script

1. 在試算表的上方選單，點選 **「擴充功能」→「Apps Script」**
2. Apps Script 編輯器會在新分頁中開啟

### 步驟 3：貼上程式碼

1. 在 `Code.gs` 中，把預設的程式碼全部刪掉（`Ctrl + A` 全選，然後 `Delete`）
2. 打開程式碼檔案 `GAS程式碼/01_RAG知識庫_LineBot.gs`
3. 複製裡面的**全部內容**
4. 貼到 Apps Script 的 `Code.gs` 中（`Ctrl + V`）

### 步驟 4：修改設定值

在程式碼的最上方，找到這兩行：

```javascript
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
```

把它們替換成你在單元 0 取得的 Token 和 User ID（跟剛才一樣的操作）。

### 步驟 5：儲存

按下 **Ctrl + S** 儲存程式碼。

### 步驟 6：部署為網頁應用程式

這是很重要的一步！我們要把程式「部署」到網路上，讓它變成一個可以透過網址存取的 Web App。

1. 在 Apps Script 編輯器的上方選單，點選 **「部署」**
2. 在下拉選單中，點選 **「新增部署作業」**

> 📸 **截圖 15**：（新增部署作業的對話框）
>
> 擷取位置：點選「部署」→「新增部署作業」後出現的對話框
> 用途：讓學員看到部署設定對話框的外觀

3. 在彈出的對話框中：
   - 點選左上角的**齒輪圖示**（選取類型旁邊）
   - 選擇 **「網頁應用程式」**

4. 填寫部署設定：

> 📸 **截圖 16**：（部署設定：類型、說明、權限）
>
> 擷取位置：新增部署作業對話框中，已選擇「網頁應用程式」類型，並填好各項設定
> 用途：讓學員看到每個欄位要怎麼填

| 設定項目 | 填寫內容 |
|----------|----------|
| 說明 | **知識庫 API v1** |
| 執行身分 | **我**（預設值） |
| 誰可以存取 | **所有人** |

> ⚠️ **注意**：「誰可以存取」一定要選「**所有人**」！如果選錯，Line Bot 就無法存取你的程式。

5. 點選 **「部署」** 按鈕

### 步驟 7：授權存取

第一次部署時，Google 會要求你授權。

> 📸 **截圖 17**：（授權流程的「審查權限」按鈕）
>
> 擷取位置：部署過程中跳出的授權視窗，標示「審查權限」按鈕
> 用途：讓學員知道要點哪個按鈕

1. 點選 **「授權存取」** 或 **「審查權限」**
2. 選擇你的 Google 帳號
3. 你可能會看到「Google 尚未驗證這個應用程式」的警告頁面
4. 點選左下角的 **「進階」**

> 📸 **截圖 18**：（「前往不安全」的進階連結）
>
> 擷取位置：Google 驗證警告頁面，標示「進階」展開後的「前往 XXX（不安全）」連結
> 用途：讓學員知道在警告頁面中要點哪裡

5. 點選 **「前往 XXX（不安全）」**（XXX 是你的專案名稱）
6. 在權限列表頁面，點選 **「允許」**

### 步驟 8：複製部署網址

授權完成後，你會看到部署成功的畫面，上面有一組**網址**。

> 📸 **截圖 19**：（部署成功後的網址）
>
> 擷取位置：部署成功的對話框，顯示「網頁應用程式」的 URL
> 用途：讓學員知道要複製哪一個網址

1. 點選網址旁邊的 **「複製」** 按鈕
2. 把這個網址貼到記事本保存（後面會用到）
3. 點選 **「完成」**

> ⚠️ **注意**：這個網址非常重要！它長得像這樣：
> `https://script.google.com/macros/s/AKfycb.../exec`
> 請務必保存好。

### 步驟 9：用瀏覽器測試 REST API

現在我們用瀏覽器來測試看看這個 API 能不能正常運作。

1. 開啟一個新的瀏覽器分頁
2. 在網址列貼上你剛才複製的部署網址
3. 在網址最後面加上 **`?action=list`**

完整網址會像這樣：
```
https://script.google.com/macros/s/AKfycb.../exec?action=list
```

4. 按下 Enter

> 📸 **截圖 20**：（瀏覽器中 ?action=list 的 JSON 結果）
>
> 擷取位置：瀏覽器中顯示 JSON 格式的知識庫資料（含範例資料）
> 用途：讓學員看到 API 成功回傳的 JSON 資料

5. 你應該會在瀏覽器中看到一堆 JSON 格式的文字，裡面包含了知識庫的範例資料

6. 接下來測試搜尋功能。把網址改成：
```
https://script.google.com/macros/s/AKfycb.../exec?action=search&query=GAS
```
你會看到搜尋「GAS」的結果

7. 再測試統計功能：
```
https://script.google.com/macros/s/AKfycb.../exec?action=stats
```
你會看到知識庫的統計資料

### 步驟 10：確認試算表

1. 切換回 Google 試算表的分頁（「知識庫系統」）
2. 你會發現底部多了兩個工作表分頁：**「知識庫」** 和 **「查詢記錄」**

> 📸 **截圖 21**：（試算表中自動建立的「知識庫」工作表）
>
> 擷取位置：Google 試算表畫面，顯示「知識庫」工作表的資料（綠色標題列、五筆範例資料）
> 用途：讓學員看到程式自動建立的工作表和範例資料

3. 點選「知識庫」工作表，你會看到五筆範例資料，包括「什麼是 Google Apps Script？」等問答
4. 點選「查詢記錄」工作表，你會看到剛才用瀏覽器測試時留下的查詢記錄

### 步驟 11：設定 Line Bot Webhook URL

現在我們要把 GAS 的部署網址告訴 Line Bot，這樣使用者在 Line 上傳訊息時，Line 平台就會把訊息轉發到你的 GAS 程式。

1. 切換到 **Line Developers Console** 的瀏覽器分頁
2. 進入你的 Channel
3. 點選上方的 **「Messaging API」** 分頁
4. 捲動到 **「Webhook settings」** 區域
5. 在 **「Webhook URL」** 欄位中，貼上你的 GAS 部署網址

> 📸 **截圖 22**：（Line Developers Console 的 Webhook URL 設定）
>
> 擷取位置：Messaging API 分頁中的 Webhook settings 區域，標示 Webhook URL 輸入框和 Verify 按鈕
> 用途：讓學員看到在哪裡貼上 Webhook URL

6. 點選 **「Update」** 按鈕儲存
7. 點選 **「Verify」** 按鈕測試連線
8. 如果看到 **「Success」** 的綠色提示，表示連線成功！
9. 確認 **「Use webhook」** 的開關是**開啟**的狀態（綠色）

> ⚠️ **注意**：如果 Verify 失敗，請檢查：
> 1. 網址是否正確（是 `/exec` 結尾，不是 `/dev` 結尾）
> 2. 部署時的存取權限是否選了「所有人」
> 3. 如果修改了程式碼，需要重新部署（「部署」→「管理部署作業」→ 編輯 → 版本選「新版本」→ 部署）

### 步驟 12：在 Line 上測試

1. 拿起手機，打開 Line
2. 找到你的 Bot（剛才加入好友的那個）
3. 在聊天輸入框中，輸入 **「GAS」** 然後送出

> 📸 **截圖 23**：（Line 上傳送查詢並收到回覆的畫面）
>
> 擷取位置：手機 Line 的聊天畫面，顯示使用者傳送「GAS」後，Bot 回覆知識庫查詢結果
> 用途：讓學員看到 Line Bot 查詢回覆的實際效果

4. Bot 應該會回覆類似這樣的訊息：

```
🔍 查詢「GAS」的結果：

📌 什麼是 Google Apps Script？
💡 Google Apps Script（GAS）是 Google 提供的雲端腳本語言...

📌 GAS 有什麼限制？
💡 GAS 免費版主要限制...
```

5. 你也可以試試輸入其他關鍵字，例如「Line Bot」、「API」、「Webhook」

**恭喜！實作 1 完成！** 你已經建立了一個可以透過 Line 查詢的知識庫機器人！

---

## 實作 2：Line Bot 訊息收集器（50 分鐘）

### 學習目標

- 學會 `doPost()` 接收 Line Webhook 事件
- 學會 `UrlFetchApp` 呼叫 Line Messaging API
- 學會 `DriveApp` 操作 Google 雲端硬碟（建立資料夾、儲存檔案）
- 學會處理不同類型的 Line 訊息（文字、圖片、位置、貼圖）

### 成果預覽

完成後你會擁有一個「訊息收集器」：
- 在 Line 上傳文字 → 自動記錄到 Google 試算表
- 在 Line 上傳圖片 → 自動存到 Google Drive + 試算表記錄連結
- 在 Line 上傳位置 → 記錄經緯度到試算表

📄 **程式碼檔案**：`GAS程式碼/02_LineBot訊息收集器.gs`

---

### 步驟 1：建立新的 Google 試算表

1. 開啟新分頁，到 **sheets.google.com**
2. 建立一個新的空白試算表
3. 命名為 **「Line 訊息記錄」**
4. 記下這個試算表的 **ID**

> **什麼是試算表 ID？**
> 看你的瀏覽器網址列，網址長這樣：
> `https://docs.google.com/spreadsheets/d/`**`1aBcDeFgHiJkLmNoPqRsTuVwXyZ`**`/edit`
>
> 夾在 `/d/` 和 `/edit` 之間的那一串文字，就是試算表 ID。

> 📸 **截圖 25**：（試算表 ID 位置，在網址中標示）
>
> 擷取位置：瀏覽器網址列，用紅色方框標出 /d/ 和 /edit 之間的 ID 部分
> 用途：讓學員學會找到試算表 ID

5. 把試算表 ID 複製到記事本

### 步驟 2：在 Google Drive 建立資料夾

1. 開啟新分頁，到 **drive.google.com**
2. 點選左上角的 **「+ 新增」** 按鈕
3. 選擇 **「新資料夾」**
4. 命名為 **「Line 檔案」**
5. 點選「建立」
6. 點進剛建立的資料夾
7. 看瀏覽器網址列，網址長這樣：
   `https://drive.google.com/drive/folders/`**`1aBcDeFgHiJkLmNoPqRsTuVwXyZ`**

> 📸 **截圖 24**：（Google Drive 資料夾的 ID 位置，在網址中標示）
>
> 擷取位置：瀏覽器網址列，用紅色方框標出 folders/ 後面的 ID 部分
> 用途：讓學員學會找到資料夾 ID

8. `folders/` 後面那串文字就是**資料夾 ID**
9. 把資料夾 ID 也複製到記事本

### 步驟 3：開啟 Apps Script 並貼上程式碼

1. 回到「Line 訊息記錄」試算表
2. 點選 **「擴充功能」→「Apps Script」**
3. 在 `Code.gs` 中，把預設的程式碼全部刪掉
4. 打開程式碼檔案 `GAS程式碼/02_LineBot訊息收集器.gs`
5. 複製全部內容，貼到 `Code.gs` 中

### 步驟 4：修改四個設定值

在程式碼最上方，找到這四行：

```javascript
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
var ROOT_FOLDER_ID = '在此貼上 Google Drive 資料夾 ID';
var SPREADSHEET_ID = '在此貼上 Google 試算表 ID';
```

> 📸 **截圖 26**：（程式碼中四個設定值的位置）
>
> 擷取位置：Apps Script 編輯器中，用紅色方框標出四個需要修改的設定值
> 用途：讓學員清楚看到需要修改哪些地方

把每一個都替換成對應的值：
- `LINE_TOKEN`：你的 Channel Access Token（跟之前一樣）
- `LINE_USER_ID`：你的 User ID（跟之前一樣）
- `ROOT_FOLDER_ID`：剛才建立的 Google Drive 資料夾 ID
- `SPREADSHEET_ID`：剛才建立的試算表 ID

> ⚠️ **注意**：四個值都要放在單引號 `'...'` 裡面。

### 步驟 5：儲存

按下 **Ctrl + S** 儲存。

### 步驟 6：部署為網頁應用程式

1. 點選 **「部署」→「新增部署作業」**
2. 選擇類型：**網頁應用程式**
3. 填寫設定：
   - 說明：**訊息收集器 v1**
   - 執行身分：**我**
   - 誰可以存取：**所有人**
4. 點選 **「部署」**
5. 完成授權流程（跟實作 1 一樣的步驟）
6. **複製部署網址**到記事本

### 步驟 7：設定 Line Bot Webhook URL

> ⚠️ **注意**：因為一個 Line Bot 同一時間只能設定一個 Webhook URL，所以設定新的 URL 會覆蓋掉實作 1 的 URL。這是正常的！後面如果要回去測試實作 1 的 Line Bot 功能，只需要把 Webhook URL 改回去就好。

1. 回到 **Line Developers Console**
2. 進入你的 Channel → **Messaging API** 分頁
3. 在 **Webhook URL** 欄位，貼上這次的新部署網址
4. 點選 **「Update」**
5. 點選 **「Verify」** 確認成功
6. 確認 **「Use webhook」** 是開啟的

### 步驟 8：測試

拿起手機，在 Line 上測試以下幾種訊息：

> 📸 **截圖 27**：（Line 傳送各種訊息的畫面）
>
> 擷取位置：手機 Line 的聊天畫面，顯示使用者傳送文字訊息和圖片的畫面
> 用途：讓學員看到傳送不同類型訊息的示範

**測試 1 — 傳送文字**：
1. 在 Line 上對 Bot 傳送一段文字，例如「你好，這是測試訊息！」
2. 切換到電腦，打開「Line 訊息記錄」試算表
3. 你會看到試算表中自動出現了一筆記錄

> 📸 **截圖 28**：（試算表中自動記錄的訊息列表）
>
> 擷取位置：Google 試算表畫面，顯示自動記錄的訊息（時間、群組ID、用戶名稱、訊息內容等欄位）
> 用途：讓學員看到訊息被自動記錄到試算表的效果

**測試 2 — 傳送圖片**：
1. 在 Line 上對 Bot 傳送一張圖片（從相簿選一張）
2. 等待幾秒鐘
3. 回到試算表，你會看到新的一筆記錄，F 欄寫著「【圖片】」，G 欄有一個 Google Drive 連結
4. 打開 Google Drive，進入「Line 檔案」資料夾

> 📸 **截圖 29**：（Google Drive 中自動儲存的圖片檔）
>
> 擷取位置：Google Drive 的「Line 檔案」資料夾中，顯示自動存入的圖片檔案
> 用途：讓學員看到圖片被自動存到 Drive 的效果

5. 你會看到剛才傳的圖片已經自動存在這裡了！

**測試 3 — 傳送位置**：
1. 在 Line 上，點選聊天輸入框旁邊的 **「+」** 號
2. 選擇 **「位置資訊」**
3. 選擇一個位置後送出
4. 回到試算表，你會看到一筆記錄，內容包含經緯度資訊

**恭喜！實作 2 完成！** 你已經建立了一個自動收集 Line 訊息的系統！

---

## 實作 3：產出第一個 GAS → Line Bot 通知（30 分鐘）

### 學習目標

- 學會 `SpreadsheetApp.getActiveSheet()` 取得工作表
- 學會 `getRange()` / `getValue()` / `setValue()` 讀寫儲存格
- 學會自訂函式（Custom Function）— 在試算表中用 `=函式名稱()` 呼叫
- 學會將試算表資料統計後推播到 Line Bot

### 成果預覽

完成後你會擁有：
- 一個有品項資料的試算表（含自訂函式計算小計）
- 在 Line 上收到銷售報表

📄 **程式碼檔案**：`GAS程式碼/03_第一個GAS_LineBot.gs`

---

### 步驟 1：建立新的 Google 試算表

1. 到 **sheets.google.com** 建立一個新的空白試算表
2. 命名為 **「銷售報表」**

### 步驟 2：開啟 Apps Script 並貼上程式碼

1. 點選 **「擴充功能」→「Apps Script」**
2. 把 `Code.gs` 的預設內容清空
3. 打開程式碼檔案 `GAS程式碼/03_第一個GAS_LineBot.gs`
4. 複製全部內容，貼到 `Code.gs`

### 步驟 3：修改設定值

找到程式碼最上方的：

```javascript
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
```

替換成你的 Token 和 User ID。

### 步驟 4：儲存

按下 **Ctrl + S** 儲存。

### 步驟 5：執行 createSampleData() 建立範例資料

1. 在上方工具列的函式選擇下拉選單中，選擇 **「createSampleData」**
2. 點選 **▶ 執行**
3. 如果跳出授權視窗，請完成授權
4. 等待幾秒鐘
5. 切換到試算表頁面

> 📸 **截圖 30**：（試算表中的範例資料，含自訂函式）
>
> 擷取位置：Google 試算表畫面，顯示品項名稱、數量、單價、小計四個欄位，其中小計欄使用 =calcSubtotal() 自訂函式
> 用途：讓學員看到程式自動建立的範例資料和自訂函式的效果

6. 你會看到試算表中已經自動填入了五筆品項資料：
   - 美式咖啡、拿鐵咖啡、珍珠奶茶、綠茶、紅茶
   - 每個品項都有數量和單價
   - **D 欄（小計）** 使用了自訂函式 `=calcSubtotal(B2,C2)` 來計算

7. 點選 D2 儲存格，在上方的公式列中你會看到 `=calcSubtotal(B2,C2)`

> **小知識**：`calcSubtotal` 是我們用 GAS 程式碼自己定義的函式！只要在程式碼中寫了一個帶有 `@customfunction` 標記的函式，就可以像 Excel 內建函式（如 SUM、AVERAGE）一樣，在儲存格中用 `=函式名稱()` 來呼叫。

### 步驟 6：執行 readCell() — 讀取單一儲存格

1. 在函式下拉選單中，選擇 **「readCell」**

> 📸 **截圖 31**：（選擇要執行的函式下拉選單）
>
> 擷取位置：Apps Script 編輯器上方工具列，展開函式選擇下拉選單，顯示所有可執行的函式列表
> 用途：讓學員看到有哪些函式可以執行

2. 點選 **▶ 執行**
3. 等待幾秒鐘，拿起手機
4. Line 上會收到一則訊息：「📖 讀取結果 — A1 儲存格的值是：品項名稱」

### 步驟 7：執行 readRange() — 讀取一個範圍

1. 選擇 **「readRange」** 函式
2. 點選 **▶ 執行**
3. Line 上會收到一則訊息，列出 A1:C5 範圍內的所有資料

### 步驟 8：執行 sendDailyReport() — 推播每日報表

1. 選擇 **「sendDailyReport」** 函式
2. 點選 **▶ 執行**
3. Line 上會收到一則完整的銷售報表：

> 📸 **截圖 32**：（Line 收到的每日銷售報表）
>
> 擷取位置：手機 Line 的聊天畫面，顯示 Bot 傳來的「每日銷售報表」訊息（含各品項明細、總金額、最高品項）
> 用途：讓學員看到報表推播的實際效果

```
📊 每日銷售報表
━━━━━━━━━━━━
📅 2026/3/3 上午10:30:00

  • 美式咖啡 ×10 = $600
  • 拿鐵咖啡 ×8 = $640
  • 珍珠奶茶 ×15 = $825
  • 綠茶 ×12 = $360
  • 紅茶 ×6 = $150

━━━━━━━━━━━━
💰 總金額：$2,575
📦 總數量：51 個
🏆 最高品項：珍珠奶茶（$825）
```

**恭喜！實作 3 完成！** 你已經學會了試算表的基本讀寫操作和自訂函式！

---

## 實作 4：部署教學 — HTML 表單 → Line Bot 通知（40 分鐘）

### 學習目標

- 學會 `HtmlService` 建立網頁介面
- 學會 `doGet()` 回傳 HTML 頁面
- 學會 `google.script.run` 前端呼叫後端函式
- 學會完整的部署流程
- 學會在 Apps Script 中建立 HTML 檔案

### 成果預覽

完成後你會擁有：
- 一個公開的聯絡表單網頁（任何人都可以填寫）
- 有人填寫表單時，資料自動寫入試算表 + Line 即時推播通知

📄 **程式碼檔案**：
- `GAS程式碼/04_HTML表單_主程式.gs`（後端程式）
- `GAS程式碼/04_HTML表單_前端.html`（前端網頁）

---

### 步驟 1：建立新的 Google 試算表

1. 到 **sheets.google.com** 建立一個新的空白試算表
2. 命名為 **「表單系統」**

### 步驟 2：開啟 Apps Script

1. 點選 **「擴充功能」→「Apps Script」**

### 步驟 3：貼上後端程式碼

1. 在 `Code.gs` 中，把預設內容清空
2. 打開程式碼檔案 `GAS程式碼/04_HTML表單_主程式.gs`
3. 複製全部內容，貼到 `Code.gs`

### 步驟 4：新增 HTML 檔案

這一步很重要！我們需要建立一個新的 HTML 檔案，放前端網頁的程式碼。

1. 在左側檔案列表旁邊，點選 **「+」** 號

> 📸 **截圖 33**：（Apps Script 中新增 HTML 檔案的步驟）
>
> 擷取位置：Apps Script 編輯器左側，點選「+」後出現的下拉選單，標示「HTML」選項
> 用途：讓學員知道如何新增 HTML 檔案

2. 選擇 **「HTML」**
3. 在出現的檔案名稱輸入框中，輸入 **form**（不需要加 `.html`，系統會自動加）
4. 按下 Enter

> 📸 **截圖 34**：（兩個檔案：Code.gs 和 form.html 的檔案列表）
>
> 擷取位置：Apps Script 編輯器左側檔案列表，顯示 Code.gs 和 form.html 兩個檔案
> 用途：讓學員確認檔案結構正確

5. 新建立的 `form.html` 會自動打開，裡面有一些預設的 HTML 程式碼
6. 把預設內容**全部刪掉**
7. 打開程式碼檔案 `GAS程式碼/04_HTML表單_前端.html`
8. 複製全部內容，貼到 `form.html` 中

> ⚠️ **注意**：HTML 檔案的名稱一定要是 **form**（全小寫）！因為程式碼中是用 `HtmlService.createHtmlOutputFromFile('form')` 來載入它的。如果名稱不對，網頁就會無法顯示。

### 步驟 5：修改設定值

1. 點選左側的 **`Code.gs`** 切換回後端程式碼
2. 找到最上方的 LINE_TOKEN 和 LINE_USER_ID
3. 替換成你的值

### 步驟 6：儲存所有檔案

按下 **Ctrl + S** 儲存（兩個檔案都要確認有儲存）。

### 步驟 7：部署為網頁應用程式

1. 點選 **「部署」→「新增部署作業」**
2. 選擇類型：**網頁應用程式**
3. 填寫設定：
   - 說明：**聯絡表單 v1**
   - 執行身分：**我**
   - 誰可以存取：**所有人**
4. 點選 **「部署」**
5. 完成授權流程
6. **複製部署網址**

### 步驟 8：開啟表單網頁

1. 在瀏覽器中開啟一個新分頁
2. 貼上部署網址，按 Enter

> 📸 **截圖 35**：（部署後瀏覽器中顯示的表單網頁）
>
> 擷取位置：瀏覽器中顯示的聯絡表單網頁（紫色漸層背景、白色表單卡片、有姓名/Email/訊息欄位）
> 用途：讓學員看到部署成功的表單外觀

3. 你會看到一個漂亮的聯絡表單頁面，有以下欄位：
   - 姓名
   - Email
   - 訊息內容
   - 送出按鈕

### 步驟 9：填寫表單並送出

1. 在表單中填入測試資料：
   - 姓名：**測試同學**
   - Email：**test@example.com**
   - 訊息內容：**你好！這是我的第一次表單測試。**
2. 點選 **「送出表單」** 按鈕
3. 等待幾秒鐘（按鈕會變成「提交中...」）

> 📸 **截圖 36**：（填寫表單送出成功的畫面）
>
> 擷取位置：瀏覽器中的表單頁面，顯示綠色的「提交成功！感謝你的填寫。」提示
> 用途：讓學員看到表單送出成功的畫面

4. 你會看到綠色的成功提示：「提交成功！感謝你的填寫。」

### 步驟 10：確認試算表

1. 切換到「表單系統」的試算表
2. 底部會出現一個新的工作表分頁：**「表單回覆」**
3. 點選它，你會看到剛才填寫的資料

> 📸 **截圖 37**：（試算表中的表單回覆資料）
>
> 擷取位置：Google 試算表的「表單回覆」工作表，顯示時間戳記、姓名、Email、訊息內容
> 用途：讓學員看到表單資料被自動記錄到試算表

### 步驟 11：檢查 Line 通知

拿起手機，打開 Line：

> 📸 **截圖 38**：（Line 收到的表單通知）
>
> 擷取位置：手機 Line 的聊天畫面，顯示 Bot 傳來的「收到新的表單填寫！」通知（含姓名、Email、訊息內容）
> 用途：讓學員看到 Line Bot 即時推播表單通知的效果

你會看到：
```
📬 收到新的表單填寫！
━━━━━━━━━━━━
👤 姓名：測試同學
📧 Email：test@example.com
💬 訊息：你好！這是我的第一次表單測試。
━━━━━━━━━━━━
📅 時間：2026/3/3 下午1:05:00
📊 累計第 1 筆
```

> **想想看**：這個功能可以用在哪些地方？
> - 活動報名表 → 有人報名就收到通知
> - 客戶聯絡表 → 客戶留言就收到通知
> - 問卷調查 → 有人填寫就收到通知
> - 意見回饋 → 有人提交就收到通知

**恭喜！實作 4 完成！** 你已經學會了如何建立 HTML 表單 + Line Bot 即時通知！

---

## 實作 5：將結果寫入試算表 → Line Bot 簽到回報（30 分鐘）

### 學習目標

- 學會 `appendRow()` 新增一整列資料
- 學會 `getRange().setValues()` 批次寫入多列
- 學會日期時間處理 `new Date()`
- 學會格式化試算表（粗體、背景色、欄寬、凍結列、框線）

### 成果預覽

完成後你會擁有：
- 一個格式漂亮的簽到表（藍色標題列、凍結列、自動格式化）
- 簽到後 Line 收到確認通知
- 遲到的列有淡紅色背景

📄 **程式碼檔案**：`GAS程式碼/05_寫入試算表_LineBot.gs`

---

### 步驟 1：建立新的 Google 試算表

1. 到 **sheets.google.com** 建立一個新的空白試算表
2. 命名為 **「簽到系統」**

### 步驟 2：開啟 Apps Script 並貼上程式碼

1. 點選 **「擴充功能」→「Apps Script」**
2. 清空 `Code.gs` 的內容
3. 打開程式碼檔案 `GAS程式碼/05_寫入試算表_LineBot.gs`
4. 複製全部內容，貼到 `Code.gs`

### 步驟 3：修改設定值

找到程式碼最上方的 LINE_TOKEN 和 LINE_USER_ID，替換成你的值。

### 步驟 4：儲存

按下 **Ctrl + S** 儲存。

### 步驟 5：執行 createSignInSheet() — 建立簽到表

1. 在函式下拉選單中，選擇 **「createSignInSheet」**
2. 點選 **▶ 執行**
3. 完成授權（如果需要的話）
4. 切換到試算表

> 📸 **截圖 39**：（格式化後的簽到表，標題列藍色、凍結列）
>
> 擷取位置：Google 試算表畫面，顯示「簽到表」工作表，有藍色標題列（序號、時間戳記、姓名、狀態、備註）、已凍結的標題列
> 用途：讓學員看到程式自動建立的格式化簽到表

5. 你會看到：
   - 底部多了一個 **「簽到表」** 工作表
   - 標題列是**深藍色背景 + 白色粗體字**
   - 欄位有：序號、時間戳記、姓名、狀態、備註
   - 標題列已經**凍結**（往下捲動時標題不會消失）
   - 每一欄的寬度都已經自動調整好了

6. Line 也會收到通知：「✅ 簽到表已建立完成！」

### 步驟 6：執行 signIn() — 單筆簽到

1. 回到 Apps Script 編輯器
2. 選擇 **「signIn」** 函式
3. 點選 **▶ 執行**
4. 切換到試算表

> 📸 **截圖 40**：（signIn 後新增的簽到紀錄，含遲到的紅色背景）
>
> 擷取位置：Google 試算表的簽到表，顯示新增的簽到記錄列，如果是遲到的列會有淡紅色背景
> 用途：讓學員看到簽到記錄和遲到標記的效果

5. 試算表中會多出一筆簽到記錄
6. Line 會收到簽到確認通知

> **小知識**：程式會自動判斷簽到時間。9:00 前（含）算「準時」，9:00 後算「遲到」。遲到的列會自動標上**淡紅色背景**，非常一目了然！

### 步驟 7：執行 batchSignIn() — 批次簽到

1. 選擇 **「batchSignIn」** 函式
2. 點選 **▶ 執行**
3. 切換到試算表
4. 你會看到一次新增了 5 筆資料！其中「張志明」和「李建國」標記為遲到，有淡紅色背景

> **技術重點**：`batchSignIn()` 使用了 `getRange().setValues()` 一次寫入多列資料，比起用迴圈逐筆 `appendRow()` 效率高出很多。當資料量大的時候，這個技巧非常有用！

### 步驟 8：執行 sendSignInSummary() — 推播統計

1. 選擇 **「sendSignInSummary」** 函式
2. 點選 **▶ 執行**
3. 拿起手機看 Line

> 📸 **截圖 41**：（Line 收到的簽到確認和統計報表）
>
> 擷取位置：手機 Line 的聊天畫面，顯示 Bot 傳來的簽到統計報表（含總人數、準時人數、遲到人數、準時率）
> 用途：讓學員看到統計報表推播的效果

你會收到類似這樣的訊息：
```
📊 簽到統計報表
━━━━━━━━━━━━
📅 2026/3/3 下午1:30:00

👥 總簽到人數：6 人
✅ 準時：4 人
⚠️ 遲到：2 人
📈 出席率：100%
⏱ 準時率：67%
```

**恭喜！實作 5 完成！** 你已經學會了試算表的進階寫入和格式化技巧！

---

## 實作 6：天氣預報 → Line Bot 每日推播（40 分鐘）

### 學習目標

- 學會 `UrlFetchApp.fetch()` 呼叫外部 API
- 學會 `JSON.parse()` 解析 JSON 資料
- 學會串接中央氣象署開放資料 API
- 學會設定**時間驅動觸發條件**（排程自動執行）

### 成果預覽

完成後你會擁有：
- 每天早上自動在 Line 收到天氣預報
- 包含天氣狀況、溫度、降雨機率
- 下雨天自動提醒帶傘

📄 **程式碼檔案**：`GAS程式碼/06_天氣預報_LineBot.gs`

---

### 步驟 1：申請中央氣象署 API 授權碼

1. 打開瀏覽器，輸入網址：**https://opendata.cwa.gov.tw/**

> 📸 **截圖 42**：（中央氣象署開放資料平台首頁）
>
> 擷取位置：瀏覽器開啟 https://opendata.cwa.gov.tw/ 後的首頁畫面
> 用途：讓學員確認進入的是正確的網站

2. 點選右上角的 **「登入/註冊」**（如果已經有帳號就直接登入）
3. 如果是第一次使用，點選 **「註冊」**，填寫以下資料：
   - Email（用你的 Gmail 即可）
   - 密碼
   - 姓名
4. 完成註冊後登入
5. 登入後，點選右上角的你的名字
6. 進入 **「會員中心」**（或「API 授權碼」頁面）

> 📸 **截圖 43**：（API 授權碼取得頁面）
>
> 擷取位置：中央氣象署會員中心的 API 授權碼頁面，標示授權碼的位置
> 用途：讓學員找到並複製 API 授權碼

7. 在頁面中找到 **「API 授權碼」**
8. 點選 **「取得授權碼」** 或直接複製已有的授權碼
9. 把授權碼複製到記事本保存

> ⚠️ **注意**：API 授權碼是免費的，每天有 10,000 次呼叫額度，我們每天只呼叫 1 次，完全夠用！

### 步驟 2：建立新的 Google 試算表

1. 到 **sheets.google.com** 建立一個新的空白試算表
2. 命名為 **「天氣預報」**

### 步驟 3：開啟 Apps Script 並貼上程式碼

1. 點選 **「擴充功能」→「Apps Script」**
2. 清空 `Code.gs`
3. 打開程式碼檔案 `GAS程式碼/06_天氣預報_LineBot.gs`
4. 複製全部內容，貼到 `Code.gs`

### 步驟 4：修改三個設定值

在程式碼最上方，找到這三行：

```javascript
var LINE_TOKEN = '在此貼上你的 Channel Access Token';
var LINE_USER_ID = '在此貼上你的 User ID';
var CWA_API_KEY = '在此貼上你的中央氣象署 API 授權碼';
```

> 📸 **截圖 44**：（三個設定值在程式碼中的位置）
>
> 擷取位置：Apps Script 編輯器中，用紅色方框標出三個需要修改的設定值
> 用途：讓學員清楚看到需要修改的三個值

把它們替換成你的值：
- `LINE_TOKEN`：你的 Channel Access Token
- `LINE_USER_ID`：你的 User ID
- `CWA_API_KEY`：剛才申請的中央氣象署 API 授權碼

### 步驟 5：修改查詢縣市（可選）

在程式碼中找到這一行：

```javascript
var TARGET_CITY = '臺北市';
```

如果你不在臺北，可以改成你所在的縣市。**注意用字**：

| 正確寫法 | 錯誤寫法 |
|----------|----------|
| 臺北市 | 台北市 |
| 臺中市 | 台中市 |
| 臺南市 | 台南市 |
| 臺東縣 | 台東縣 |

> ⚠️ **注意**：氣象署 API 使用「**臺**」而不是「**台**」！如果用錯字，API 會找不到資料。其他縣市如高雄市、新北市、桃園市等則正常使用。

### 步驟 6：儲存

按下 **Ctrl + S** 儲存。

### 步驟 7：執行 testApiConnection() — 測試 API 連線

1. 在函式下拉選單中，選擇 **「testApiConnection」**
2. 點選 **▶ 執行**
3. 完成授權（如果需要的話）
4. 拿起手機看 Line

> 📸 **截圖 45**：（testApiConnection 成功的 Line 通知）
>
> 擷取位置：手機 Line 的聊天畫面，顯示 Bot 傳來的「✅ 氣象 API 連線測試成功！」訊息
> 用途：讓學員確認 API 連線成功

5. 如果看到「✅ 氣象 API 連線測試成功！」，表示 API 授權碼和設定都正確

> ⚠️ **如果看到「❌ 氣象 API 連線失敗」**：
> 1. 檢查 API 授權碼是否正確貼上
> 2. 檢查縣市名稱是否用了「臺」字
> 3. 查看 Apps Script 的執行記錄（點選左側的「執行項目」圖示），看詳細的錯誤訊息

### 步驟 8：執行 sendWeatherReport() — 推播天氣預報

1. 選擇 **「sendWeatherReport」** 函式
2. 點選 **▶ 執行**
3. 拿起手機看 Line

> 📸 **截圖 46**：（完整的天氣預報 Line 訊息）
>
> 擷取位置：手機 Line 的聊天畫面，顯示 Bot 傳來的完整天氣預報（含三個時段的天氣、溫度、降雨機率、舒適度，以及貼心提醒）
> 用途：讓學員看到天氣預報推播的完整效果

你會收到類似這樣的天氣預報訊息：
```
🌤 天氣預報 — 臺北市
━━━━━━━━━━━━━━
📅 2026/3/3 下午2:30:00

🕐 03-03 06:00 ~ 18:00
   ⛅ 多雲時晴
   🌡 18°C ~ 25°C
   🌧 降雨機率：20%
   😊 舒適

🕐 03-03 18:00 ~ 06:00
   ☁️ 多雲
   🌡 16°C ~ 21°C
   🌧 降雨機率：30%
   😊 舒適

🕐 03-04 06:00 ~ 18:00
   🌧 短暫雨
   🌡 15°C ~ 20°C
   🌧 降雨機率：70%
   😊 舒適偏涼

☂️ 提醒：降雨機率較高，記得帶傘！
```

4. 同時，切換到試算表，你會看到底部多了一個 **「天氣記錄」** 工作表，裡面記錄了天氣資料

### 步驟 9：設定每天自動推播

這是整個 Part 1 的最後一個重要技能 — **時間驅動觸發條件**！

1. 在 Apps Script 編輯器的左側選單中，點選 **鬧鐘圖示**（觸發條件）
2. 頁面會切換到觸發條件管理頁面
3. 點選右下角的 **「+ 新增觸發條件」** 按鈕

> 📸 **截圖 47**：（觸發條件設定畫面：選擇函式、時間驅動、每日計時器）
>
> 擷取位置：新增觸發條件的對話框，標示每個下拉選單的選項（函式選 sendWeatherReport、事件來源選時間驅動、類型選每日計時器、時段選上午7點到8點）
> 用途：讓學員看到觸發條件的完整設定步驟

4. 在彈出的設定表單中，依序設定：

| 設定項目 | 選擇 |
|----------|------|
| 選擇要執行的函式 | **sendWeatherReport** |
| 選擇要執行的部署作業 | **主版本** |
| 選取活動來源 | **時間驅動** |
| 選取時間型觸發條件類型 | **每日計時器** |
| 選取每天的時間 | **上午 7 點到 8 點** |

5. 點選 **「儲存」**
6. 如果跳出授權視窗，請完成授權

> 📸 **截圖 48**：（已建立的觸發條件列表）
>
> 擷取位置：觸發條件管理頁面，顯示已建立的觸發條件（函式名稱、部署作業、事件來源、時間）
> 用途：讓學員確認觸發條件已成功建立

7. 你會在觸發條件列表中看到新增的項目

> **這表示什麼？** 從明天開始，每天早上 7 點到 8 點之間，GAS 會自動執行 `sendWeatherReport()` 函式，將天氣預報推播到你的 Line 上。你什麼都不用做，早上起床就能看到天氣預報！

> ⚠️ **注意事項**：
> - 觸發條件的時間範圍是「7 點到 8 點之間」，不是精確的 7:00 整。Google 會在這個區間內隨機選一個時間執行。
> - 如果要刪除觸發條件，在列表中找到該項目，點選右邊的三個點「⋮」→「刪除觸發條件」。
> - 每個 GAS 專案最多可以設定 20 個觸發條件。

**恭喜！實作 6 完成！** 你已經學會了呼叫外部 API 和設定自動排程！

---

## Part 1 小結與回顧

恭喜你完成了 Part 1 所有的實作！讓我們回顧一下今天學到了什麼。

### 今日學到的 7 個核心技能

| 編號 | 技能 | 使用到的實作 |
|------|------|-------------|
| 1 | **GAS 環境操作與部署** | 所有實作 |
| 2 | **doGet / doPost 網頁應用程式** | 實作 1、實作 4 |
| 3 | **SpreadsheetApp 試算表讀寫** | 實作 1、3、4、5 |
| 4 | **UrlFetchApp 外部 API 串接** | 實作 1、2、6 |
| 5 | **Line Bot Webhook 接收訊息** | 實作 1、2 |
| 6 | **Line Bot Push Message 主動推播** | 實作 1、3、4、5、6 |
| 7 | **時間驅動觸發條件** | 實作 6 |

### 今日學到的 GAS 服務

| GAS 服務 | 功能 | 實作 |
|----------|------|------|
| `Logger.log()` | 在執行記錄中輸出訊息 | 單元 0 |
| `SpreadsheetApp` | 讀寫 Google 試算表 | 實作 1、3、4、5、6 |
| `ContentService` | 回傳 JSON 格式資料 | 實作 1 |
| `HtmlService` | 建立 HTML 網頁 | 實作 4 |
| `UrlFetchApp` | 呼叫外部 API | 實作 1、2、6 |
| `DriveApp` | 操作 Google 雲端硬碟 | 實作 2 |

### 你帶走的成品清單

今天你一共建立了以下成品，全部都可以繼續使用：

| 成品 | 說明 |
|------|------|
| 1 個 Line Bot | 可接收訊息、主動推播，全程使用 |
| 1 個知識庫 Line 查詢機器人 | 在 Line 輸入關鍵字就能查詢知識庫 |
| 1 個訊息記錄系統 | Line 訊息自動存入試算表 + 圖片存入 Drive |
| 1 個銷售報表系統 | 自訂函式 + Line 報表推播 |
| 1 個公開表單 + Line 通知 | HTML 表單 + 即時 Line 通知 |
| 1 個簽到系統 + Line 確認 | 格式化試算表 + 簽到 Line 推播 |
| 1 個每日天氣推播 | 每天早上自動推播天氣預報到 Line |

### 重要提醒

1. **保存好你的 Token 和 ID**：Channel Access Token、User ID、氣象署 API 授權碼，這些在 Part 2 還會繼續使用。

2. **Webhook URL 只能設一個**：同一時間只能設定一個 Webhook URL。如果要切換不同的實作，就需要去 Line Developers Console 修改 Webhook URL。

3. **修改程式碼後要重新部署**：如果你修改了已部署的程式碼，需要重新部署才會生效。操作方式：「部署」→「管理部署作業」→ 點選鉛筆圖示 → 版本選「新版本」→ 部署。

4. **Part 2 預告**：明天我們會學習更進階的內容，包括 Line Bot 圖片推播（早安圖）、RSS 新聞推播、出缺席通知、自動化報名系統、查詢系統、以及 AI 自動發文！敬請期待！

---

> **Part 1 教學結束。辛苦了，明天見！** 👋


\newpage

## Part 2：GAS 進階應用、RSS 推播與 AI 整合

> **Google Apps Script 自動化教學 — Part 2**
>
> 適用對象：已完成 Part 1 的學員（具備 GAS 基礎 + Line Bot 推播能力）
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
| 15:00 - 15:20 | Part 2 總結與延伸 | 20 min | |

---

## Part 1 回顧

在開始今天的進階實作之前，讓我們快速回顧 Part 1 學到的核心技能：

| 你已經會的 | 今天要學的新技能 |
|-----------|-----------------|
| `pushText()` 推播文字 | `pushImage()` **推播圖片** |
| 推播給自己（User ID） | 推播到 **Line 群組**（Group ID） |
| 手動觸發函式 | **事件驅動觸發**（表單提交 → 自動推播） |
| `UrlFetchApp` 呼叫 JSON API | `XmlService.parse()` **解析 RSS XML** |
| `SpreadsheetApp` 讀寫試算表 | `FormApp` **操作 Google 表單** |
| 基礎 JavaScript | 串接 **Gemini AI** API |

> 如果你的 Line Bot 還沒設定好（Channel Access Token、User ID），請先回 Part 1 的單元 0 完成設定。

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

> 跟 Part 1 的 `pushText()` 不同，`pushImage()` 的 `type` 是 `'image'`，需要提供圖片的公開網址。

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
3. 如果跳出授權視窗，按照 Part 1 的方式完成授權（審查權限 → 進階 → 前往 → 允許）

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

> 這跟 Part 1 呼叫天氣 API 的方式一樣。不同的是，RSS 回傳的不是 JSON 而是 XML。

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
2. 使用 Part 1 的「實作 2：訊息收集器」記錄群組訊息
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

> 這跟 Part 1 呼叫天氣 API 很像，都是用 `UrlFetchApp.fetch()` 發送 HTTP 請求。差別是：
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

## Part 2 總結

### 今天學到的新技能

| 技能類別 | 內容 | 對應實作 |
|----------|------|----------|
| 圖片推播 | `pushImage()` + DriveApp | 實作 8 |
| XML 解析 | `XmlService.parse()` | 實作 9 |
| 群組推播 | push to Group ID | 實作 10 |
| 事件驅動 | `onFormSubmit()` | 實作 11 |
| 前後端分離 | `google.script.run` | 實作 12 |
| AI 整合 | Gemini API | 實作 13 |

### 兩部分技能總覽

| Part | 實作 | GAS 服務 | Line Bot 技能 |
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

**Part 1（8 個成品）：**
- Line Bot + 共用推播模組
- 知識庫查詢機器人
- 訊息收集器系統
- 第一個 GAS 自動化
- HTML 表單 + 通知
- 試算表寫入系統
- 天氣預報每日推播
- AI 自動化評量系統

**Part 2（6 個成品）：**
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
