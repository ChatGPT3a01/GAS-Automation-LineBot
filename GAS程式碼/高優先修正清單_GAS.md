# GAS 高優先修正清單

## 使用方式
- `必修`：不修就很可能無法部署、無法執行、或有明顯功能落差。
- `建議修`：目前不一定會立刻壞，但有穩定性、維護性、安全性或敘述一致性問題。

## 優先順序總覽

### 第一層：最先處理
- `實作04_HTML網頁表單.gs`
- `實作04_HTML網頁表單_前端.html`
- `實作19_查詢系統雙入口.gs`
- `實作19_查詢系統雙入口_前端.html`
- `實作11_LineBot訊息收集器.gs`
- `實作15_自動化評量AI評語.gs`

### 第二層：很快會卡住執行
- 所有含有 Token / ID / API Key 佔位值的檔案
- 所有 AI 直接 `JSON.parse()` 的檔案

### 第三層：功能可跑，但品質與說明需補正
- `實作10_RSS新聞推播機器人.gs`
- `實作16_RAG知識庫問答機器人.gs`

## 逐檔清單

### 實作00_GAS環境建立.gs
- 必修：
  - 無
- 建議修：
  - 若要作為教學範例，建議補上「需綁定試算表執行」的提示，避免誤以為可在獨立 GAS 專案直接跑所有函式

### 實作01_試算表自動化報表.gs
- 必修：
  - 無
- 建議修：
  - 建議補上「需在 Google 試算表綁定專案中執行」的前置說明
  - 若要提升穩定性，可增加工作表不存在時的防呆訊息

### 實作02_自動寄信通知系統.gs
- 必修：
  - 範例 Email 需改成真實地址，否則功能驗證沒有意義
- 建議修：
  - 建議補寄信配額不足時的處理與提示
  - 建議將收件名單有效性檢查寫清楚

### 實作03_一鍵建立Google表單.gs
- 必修：
  - 無
- 建議修：
  - 建議補充「在綁定試算表專案內執行」的限制說明
  - 建議補錯誤處理，例如表單建立失敗時的提示

### 實作04_HTML網頁表單.gs
- 必修：
  - `createHtmlOutputFromFile('form')` 與目前資料夾 HTML 檔名不一致，需統一
  - `LINE_TOKEN`、`LINE_USER_ID` 必須改成真實值，否則提交通知無法使用
- 建議修：
  - 建議將設定值改用 `PropertiesService` 管理
  - 建議在 `pushLine()` 補上回應碼檢查與錯誤日誌
  - 建議在 `processForm()` 先驗證輸入格式，例如 Email 合法性

### 實作04_HTML網頁表單_前端.html
- 必修：
  - 必須與後端檔案命名規則一致，供 `createHtmlOutputFromFile()` 正確載入
  - 不能當成獨立 HTML 使用，需放進 GAS Web App
- 建議修：
  - 建議在檔案註解明確標示「此檔需命名為 form.html」
  - 建議補前端 Email 格式驗證與輸入長度限制

### 實作05_擴充選單與自訂按鈕.gs
- 必修：
  - 無
- 建議修：
  - 建議明確標示此檔只能在綁定 Google 試算表的專案內使用
  - 建議補上選單功能失敗時的錯誤提示

### 實作06_LineBot建立與推播.gs
- 必修：
  - `LINE_TOKEN`、`LINE_USER_ID` 必須填入真實值
- 建議修：
  - 建議將設定值移到 `PropertiesService`
  - 建議對 `push` / `reply` 的失敗回應加上檢查與日誌

### 實作07_天氣預報LineBot推播.gs
- 必修：
  - `LINE_TOKEN`、`LINE_USER_ID`、`CWA_API_KEY` 必須填入真實值
- 建議修：
  - 建議加強氣象 API 回傳欄位缺失時的防呆
  - 建議將外部設定改為 `PropertiesService`
  - 建議在試算表記錄前確認有綁定 Spreadsheet 環境

### 實作08_早安圖LineBot圖片推播.gs
- 必修：
  - `LINE_TOKEN`、`LINE_USER_ID`、`MORNING_FOLDER_ID` 必須填入真實值
- 建議修：
  - `https://lh3.googleusercontent.com/d/{fileId}` 這種圖片直連方式有相容性風險，建議補可用性驗證
  - 建議在公開分享失敗時提供錯誤訊息
  - 建議將設定值移到 `PropertiesService`

### 實作09_AI自動生成簡報.gs
- 必修：
  - `GEMINI_API_KEY` 必須填入真實值
- 建議修：
  - AI 回覆目前直接 `JSON.parse()`，建議補容錯與格式檢查
  - 建議補 `candidates` 不存在時的防呆
  - 建議在寫回試算表前確認有 `ActiveSheet`

### 實作10_RSS新聞推播機器人.gs
- 必修：
  - `LINE_TOKEN`、`LINE_USER_ID` 必須填入真實值
  - 若註解要成立，需補上「寫入試算表」與「避免重複推播」邏輯；目前文件與程式不一致
- 建議修：
  - 建議加入來源失效或 RSS 格式不一致時的更完整錯誤處理
  - 建議改成可設定是否去重、保留幾天紀錄

### 實作11_LineBot訊息收集器.gs
- 必修：
  - `LINE_TOKEN`、`LINE_USER_ID`、`ROOT_FOLDER_ID`、`SPREADSHEET_ID` 必須填入真實值
  - `doPost()` 建議回傳 `ContentService`，避免 webhook 端點回應不明確
- 建議修：
  - 目前 webhook 流程內同步做下載檔案、存 Drive、查 Profile，建議拆輕或降風險
  - 建議補上事件合法性檢查，例如 `e.postData` 是否存在
  - 建議將設定值移到 `PropertiesService`

### 實作13_活動報名即時通知.gs
- 必修：
  - `LINE_TOKEN`、`LINE_USER_ID`、`FORM_ID` 必須填入真實值
  - 必須建立 `onFormSubmit` 觸發器，否則核心功能不會自動執行
- 建議修：
  - 建議在 `e.namedValues` 結構不符時增加防呆
  - 建議在關閉表單前確認 `FORM_ID` 對應表單存在

### 實作14_AI個人化Email通知.gs
- 必修：
  - `GEMINI_API_KEY` 必須填入真實值
  - 範例學生 Email 必須改為真實地址
- 建議修：
  - 建議補 Gemini 回傳異常時更明確的失敗狀態
  - 建議加入 Gmail 寄送配額或速率的提示

### 實作15_自動化評量AI評語.gs
- 必修：
  - `LINE_TOKEN`、`FOLDER_ID`、`SPREADSHEET_ID`、`AI_API_KEY` 必須填入真實值
  - 必須啟用 Drive API 進階服務，否則 `Drive.Files.insert()` 會失敗
- 建議修：
  - `processFileAsync()` 名稱與實際行為不符，實際上仍是同步流程，建議重構或更名
  - webhook 流程過重，建議拆分，避免超時
  - OpenAI 端點與模型名稱需再次確認是否符合你實際使用的 API 方案
  - 建議補 LINE / AI / Drive 各步驟的回應碼與錯誤明細

### 實作16_RAG知識庫問答機器人.gs
- 必修：
  - `LINE_TOKEN`、`LINE_USER_ID` 必須填入真實值
  - 若名稱與說明要維持「RAG」，需補真正的檢索增強生成流程；否則至少要改文件名稱與定位
- 建議修：
  - 建議將 `ActiveSpreadsheet` 依賴說明清楚
  - 建議補 webhook / REST API 參數驗證
  - 建議區分「知識庫 CRUD API」與「Line 查詢 Bot」兩種用途的文件說明

### 實作17_出缺席通知群組推播.gs
- 必修：
  - `LINE_TOKEN`、`LINE_USER_ID`、`SPREADSHEET_ID` 必須填入真實值
- 建議修：
  - 建議補上工作表不存在時的初始化或錯誤提示
  - 建議將推播目標 `PUSH_TARGET` 與設定值管理方式再整理

### 實作18_AI自動發文系統.gs
- 必修：
  - `LINE_TOKEN`、`LINE_USER_ID`、`GEMINI_API_KEY`、`SPREADSHEET_ID` 必須填入真實值
- 建議修：
  - 建議補 AI 回覆結構異常時的容錯
  - 建議把主題、模型、試算表設定集中管理
  - 建議增加避免重複發文或重複主題的機制

### 實作19_查詢系統雙入口.gs
- 必修：
  - `createHtmlOutputFromFile('查詢介面')` 與目前資料夾 HTML 檔名不一致，需統一
  - `LINE_TOKEN`、`LINE_USER_ID`、`SPREADSHEET_ID` 必須填入真實值
- 建議修：
  - 建議對 `doPost()` 補更多 webhook 事件合法性驗證
  - 建議限制一次回覆的欄位數與字數，避免訊息過長
  - 建議將設定值改為 `PropertiesService`

### 實作19_查詢系統雙入口_前端.html
- 必修：
  - 必須與後端 HTML 引用名稱一致，供 GAS 正確載入
  - 不能獨立使用，必須在 GAS Web App 環境中運作
- 建議修：
  - 目前結果區使用 `innerHTML` 組字串，有 XSS 風險，建議改為安全的 DOM 建構方式
  - 建議在檔案註解中標示「需命名為 查詢介面.html」

### 實作20_AI段考審題助手.gs
- 必修：
  - `GEMINI_API_KEY`、`DOC_ID` 必須填入真實值
- 建議修：
  - 直接 `JSON.parse()` AI 回覆，建議補容錯
  - 建議補 `DocumentApp.openById()` 與試算表寫入失敗時的處理
  - 建議檢查審題結果欄位是否缺漏，避免寫表時資料結構不完整

## 依修正類型分組

### A 類：不修會直接壞
- `實作04_HTML網頁表單.gs`
- `實作04_HTML網頁表單_前端.html`
- `實作19_查詢系統雙入口.gs`
- `實作19_查詢系統雙入口_前端.html`

原因：
- HTML 檔名與 `createHtmlOutputFromFile()` 不一致
- 這是明確、立即、可重現的錯誤

### B 類：不填值就不能跑
- `實作04_HTML網頁表單.gs`
- `實作06_LineBot建立與推播.gs`
- `實作07_天氣預報LineBot推播.gs`
- `實作08_早安圖LineBot圖片推播.gs`
- `實作09_AI自動生成簡報.gs`
- `實作10_RSS新聞推播機器人.gs`
- `實作11_LineBot訊息收集器.gs`
- `實作13_活動報名即時通知.gs`
- `實作14_AI個人化Email通知.gs`
- `實作15_自動化評量AI評語.gs`
- `實作16_RAG知識庫問答機器人.gs`
- `實作17_出缺席通知群組推播.gs`
- `實作18_AI自動發文系統.gs`
- `實作19_查詢系統雙入口.gs`
- `實作20_AI段考審題助手.gs`

### C 類：可跑，但高風險
- `實作09_AI自動生成簡報.gs`
- `實作11_LineBot訊息收集器.gs`
- `實作15_自動化評量AI評語.gs`
- `實作19_查詢系統雙入口_前端.html`
- `實作20_AI段考審題助手.gs`

### D 類：文件與實作不一致
- `實作10_RSS新聞推播機器人.gs`
- `實作16_RAG知識庫問答機器人.gs`

## 最後建議的閱讀順序
- 若你要先決定「哪些一定要修」：
  - 先看 A 類
- 若你要先決定「哪些能最快跑起來」：
  - 再看 B 類
- 若你要先決定「哪些後面容易出事」：
  - 再看 C 類
- 若你要整理教材或對外說明：
  - 最後看 D 類

