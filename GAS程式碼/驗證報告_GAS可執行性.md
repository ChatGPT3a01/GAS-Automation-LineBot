# GAS / HTML 可執行性驗證報告

## 驗證範圍
- `.gs` 檔：20 份
- `.html` 檔：2 份
- `.zip` 封裝檔：4 份

## 驗證方式
- 靜態結構檢查：確認每個 `.gs` 是否把自己用到的函式寫在同一檔內。
- JavaScript 語法檢查：將 20 份 `.gs` 暫時複製成 `.js` 後，以 `node --check` 驗證，未發現 JavaScript 語法錯誤。
- 依賴掃描：檢查是否依賴 `HtmlService`、`google.script.run`、`UrlFetchApp`、`SpreadsheetApp.openById()`、`FormApp.openById()`、`PropertiesService`、`Drive.Files.insert()` 等外部條件。
- 風險檢查：針對 Webhook、HTML 檔名、AI 回傳解析、外部 API 與安全性做人工檢閱。

## 重要限制
- 這次是「靜態驗證 + 結構驗證」，不是在 Google Apps Script 雲端環境中逐支實跑。
- 目前資料夾內沒有 `clasp` 專案設定、`appsscript.json`、測試腳本、部署 URL、真實 API 金鑰，因此無法在本機直接完成真正的 GAS 執行驗證。
- 結論中的「可執行」代表：在 GAS 環境中補齊必要設定後，該檔案本身看起來可獨立成立，不需再引用本資料夾其他 `.gs`。

## 總結
- 大多數 `.gs` 檔案是「單檔自足」：函式與輔助函式都在同一檔內，沒有明顯跨檔呼叫依賴。
- 兩個 HTML Web App 類案例不是完全獨立：
  - `實作04_HTML網頁表單.gs` 依賴一個名為 `form` 的 HTML 檔。
  - `實作19_查詢系統雙入口.gs` 依賴一個名為 `查詢介面` 的 HTML 檔。
- 目前最明確的可執行阻礙，是「HTML 檔名與 GAS 程式引用名稱不一致」。

## 主要問題

### 1. HTML 檔名不一致，部署後會直接失敗
- `實作04_HTML網頁表單.gs:46` 使用 `HtmlService.createHtmlOutputFromFile('form')`
- 但資料夾中的檔名是 `實作04_HTML網頁表單_前端.html`
- 若直接把這整個檔案名稱搬進 Apps Script，不另建 `form.html`，會出現找不到 HTML 檔案

- `實作19_查詢系統雙入口.gs:59` 使用 `HtmlService.createHtmlOutputFromFile('查詢介面')`
- 但資料夾中的檔名是 `實作19_查詢系統雙入口_前端.html`
- 若不手動建立 `查詢介面.html`，部署時同樣會失敗

### 2. 兩個 `.html` 檔都不是獨立執行檔
- `實作04_HTML網頁表單_前端.html` 透過 `google.script.run.processForm(...)` 呼叫後端
- `實作19_查詢系統雙入口_前端.html` 透過 `google.script.run.searchData(...)` 呼叫後端
- 這兩個 HTML 檔若直接用本機瀏覽器開啟，不在 GAS Web App 內執行，`google.script.run` 不存在，功能無法運作

### 3. 多支檔案目前仍保留佔位設定值，未填寫前不可直接執行
- 常見項目：
  - `LINE_TOKEN`
  - `LINE_USER_ID`
  - `GEMINI_API_KEY`
  - `CWA_API_KEY`
  - `SPREADSHEET_ID`
  - `FORM_ID`
  - `DOC_ID`
  - `ROOT_FOLDER_ID`
  - `MORNING_FOLDER_ID`
- 這不是程式語法錯誤，但在目前狀態下屬於「不可直接執行」

### 4. AI JSON 解析邏輯偏脆弱
- `實作09_AI自動生成簡報.gs` 與 `實作20_AI段考審題助手.gs` 都要求 Gemini 回傳 JSON，之後直接 `JSON.parse(text)`
- 若模型回傳多一句說明、欄位缺失、格式偏掉，即可能直接拋錯
- 這類錯誤不是每次都發生，但屬於真實執行風險

### 5. Webhook 類程式存在超時或流程設計風險
- `實作11_LineBot訊息收集器.gs`
  - `doPost()` 沒有明確回傳 `ContentService` 回應
  - 檔案下載、Drive 存檔、Line Profile 查詢都在 webhook 流程內同步完成，量大時有超時風險
- `實作15_自動化評量AI評語.gs`
  - 名稱寫「非同步處理」，但 `processFileAsync()` 其實仍是同步函式呼叫
  - PDF 下載、Drive 轉檔、AI 呼叫都在 webhook 主流程內，風險比一般文字 Bot 高

### 6. 文件敘述與實作內容有落差
- `實作10_RSS新聞推播機器人.gs`
  - 註解寫「將新聞記錄到試算表（避免重複推播）」
  - 目前程式沒有實作寫入試算表，也沒有重複去重邏輯
- `實作16_RAG知識庫問答機器人.gs`
  - 名稱寫 RAG
  - 目前實作本質上是「試算表關鍵字查詢 + CRUD + Line 回覆」，沒有 embedding / vector search / 檢索增強生成

### 7. 前端顯示有一個安全性風險
- `實作19_查詢系統雙入口_前端.html`
  - 搜尋結果以字串拼接後塞入 `innerHTML`
  - 若試算表資料含惡意 HTML / script 字串，存在 XSS 風險

## 逐檔驗證結果

| 檔案 | 類型 | 是否可視為單檔自足 | 目前可直接執行 | 結論 |
|---|---|---|---|---|
| 實作00_GAS環境建立.gs | GAS | 是 | 否 | 依賴 GAS 試算表環境的 `ActiveSheet`，邏輯單純，可獨立使用 |
| 實作01_試算表自動化報表.gs | GAS | 是 | 否 | 依賴 `ActiveSpreadsheet`，檔內函式完整，可獨立使用 |
| 實作02_自動寄信通知系統.gs | GAS | 是 | 否 | 可單檔成立，但需要真實 Email 與 Gmail/Mail 配額 |
| 實作03_一鍵建立Google表單.gs | GAS | 是 | 否 | 可單檔成立，但需在 GAS + Spreadsheet / Form 環境執行 |
| 實作04_HTML網頁表單.gs | GAS | 否 | 否 | 依賴名為 `form` 的 HTML 檔；目前資料夾檔名不對 |
| 實作04_HTML網頁表單_前端.html | HTML | 否 | 否 | 必須嵌入 GAS Web App 才能透過 `google.script.run` 使用 |
| 實作05_擴充選單與自訂按鈕.gs | GAS | 是 | 否 | 可單檔成立，但只適用綁定試算表的 GAS 專案 |
| 實作06_LineBot建立與推播.gs | GAS | 是 | 否 | 可單檔成立，但要先填 LINE Token / User ID |
| 實作07_天氣預報LineBot推播.gs | GAS | 是 | 否 | 可單檔成立，但依賴中央氣象署 API、LINE、試算表 |
| 實作08_早安圖LineBot圖片推播.gs | GAS | 是 | 否 | 可單檔成立，但依賴 Drive 資料夾與公開圖片連結可用性 |
| 實作09_AI自動生成簡報.gs | GAS | 是 | 否 | 可單檔成立，但依賴 Gemini API；JSON 解析較脆弱 |
| 實作10_RSS新聞推播機器人.gs | GAS | 是 | 否 | 可單檔成立，但目前沒有實作註解所寫的去重/記錄 |
| 實作11_LineBot訊息收集器.gs | GAS | 是 | 否 | 可單檔成立，但依賴 LINE、Drive、Spreadsheet；Webhook 風險較高 |
| 實作13_活動報名即時通知.gs | GAS | 是 | 否 | 可單檔成立，但需 `FORM_ID`、Line 設定與表單提交觸發器 |
| 實作14_AI個人化Email通知.gs | GAS | 是 | 否 | 可單檔成立，但需 Gemini API 與真實 Email |
| 實作15_自動化評量AI評語.gs | GAS | 是 | 否 | 可單檔成立，但依賴 LINE、Drive、Spreadsheet、Drive 進階服務、Gemini/OpenAI；Webhook 流程較重 |
| 實作16_RAG知識庫問答機器人.gs | GAS | 是 | 否 | 單檔成立，但實際是關鍵字知識庫查詢，不是完整 RAG |
| 實作17_出缺席通知群組推播.gs | GAS | 是 | 否 | 可單檔成立，但需 LINE 與試算表 ID |
| 實作18_AI自動發文系統.gs | GAS | 是 | 否 | 可單檔成立，但需 Gemini、LINE、試算表 ID |
| 實作19_查詢系統雙入口.gs | GAS | 否 | 否 | 依賴名為 `查詢介面` 的 HTML 檔；目前資料夾檔名不對 |
| 實作19_查詢系統雙入口_前端.html | HTML | 否 | 否 | 必須在 GAS Web App 中執行；另有 `innerHTML` 風險 |
| 實作20_AI段考審題助手.gs | GAS | 是 | 否 | 可單檔成立，但需 Gemini API、Google Docs ID，且 JSON 解析較脆弱 |

## 壓縮檔檢查結果
- `03_第一個GAS_LineBot.zip`
  - 內含：`03_第一個GAS_LineBot.gs`
- `04_HTML表單_前端.zip`
  - 內含：`04_HTML表單_前端.html`、`04_HTML表單_主程式.gs`
- `05_寫入試算表_LineBot.zip`
  - 內含：`05_寫入試算表_LineBot.gs`
- `06_天氣預報_LineBot.zip`
  - 內含：`06_天氣預報_LineBot.gs`

判斷：
- 這 4 份 `.zip` 比較像舊版或教學封裝檔，不是目前資料夾主驗證目標。
- 本次未解壓重驗內容，只確認封裝項目。

## 我對「每個檔案應能獨立、不依賴其他檔案」的結論
- `.gs` 檔部分：
  - 除了 `實作04_HTML網頁表單.gs` 與 `實作19_查詢系統雙入口.gs` 之外，其餘 `.gs` 基本上都屬於「單檔自足」，沒有依賴本資料夾其他 `.gs` 檔案。
- `.html` 檔部分：
  - 兩份 `.html` 都不能算獨立可執行檔，它們是 GAS Web App 的前端片段。
- 整體來看：
  - 若你要追求「每個檔案都能獨立交付」，目前最不符合要求的是兩組 Web App：
    - `實作04_HTML網頁表單.gs` + `實作04_HTML網頁表單_前端.html`
    - `實作19_查詢系統雙入口.gs` + `實作19_查詢系統雙入口_前端.html`

## 建議後續動作
- 第一優先：
  - 先修正兩個 HTML 引用名稱與實際檔名不一致問題
- 第二優先：
  - 把所有金鑰 / ID / Token 改成 `PropertiesService` 或統一設定區，不要直接寫死在程式碼
- 第三優先：
  - 替 AI JSON 回傳加上容錯解析
- 第四優先：
  - 重 webhook 流程改成較輕量，避免在 `doPost()` 內做過多同步工作
- 第五優先：
  - 若你真的要驗證「可執行性」，下一步應建立一個正式的 `clasp` / Apps Script 專案，把這些檔案逐支匯入雲端後實跑

