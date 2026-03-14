# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 專案概述

GAS 自動化實戰教學課程，四篇共 21 個實作的互動式 HTML 簡報與 Google Apps Script 程式碼。部署於 GitHub Pages。

- **GitHub Pages**: `https://chatgpt3a01.github.io/GAS-Automation-LineBot/簡報/index.html`
- **所有回應使用繁體中文**

## 開發與部署

此專案為純靜態 HTML，無建置流程、無套件管理、無測試框架。修改後直接 git push 到 master 即可透過 GitHub Pages 生效。

```bash
git add -A && git commit -m "描述" && git push origin master
```

## 架構

### 簡報系統

每個實作是獨立的單檔 HTML（含內嵌 CSS/JS），共用 `簡報/password-gate.js` 做密碼保護。共四篇 21 個實作。

- **入口**: `簡報/index.html`（課程總覽頁，四篇卡片式佈局）
- **投影片結構**: 每個 `.slide` 元素為一頁，透過 `opacity` + `z-index` 控制切換
- **導航**: 頂部 nav-bar、左右箭頭、底部 indicators、鍵盤左右鍵
- **每個實作最後一頁**: 固定三按鈕（下載程式碼 / 下一個實作 / 回到總覽）

### 四篇課程結構

| 篇 | 主題 | 實作範圍 | 說明 |
|---|------|---------|------|
| 第一篇 | GAS 基礎入門 | 實作 0-5 | 100% 純 GAS，無 Line Bot |
| 第二篇 | Line Bot 串接 | 實作 6-10 | Line Bot + AI 簡報穿插 |
| 第三篇 | 進階互動 | 實作 11-15 | Webhook + AI Email 穿插 |
| 第四篇 | 實戰應用 | 實作 16-20 | 進階整合 + AI 審題穿插 |

### 檔案對應

| 實作 | 簡報 | GAS 程式碼 |
|------|------|-----------|
| 0 | `篇1_實作00_GAS環境建立.html` | `實作00_GAS環境建立.gs` |
| 1 | `篇1_實作01_試算表自動化報表.html` | `實作01_試算表自動化報表.gs` |
| 2 | `篇1_實作02_自動寄信通知系統.html` | `實作02_自動寄信通知系統.gs` |
| 3 | `篇1_實作03_一鍵建立Google表單.html` | `實作03_一鍵建立Google表單.gs` |
| 4 | `篇1_實作04_HTML網頁表單.html` | `實作04_HTML網頁表單.gs` + `04_HTML表單_前端.html` |
| 5 | `篇1_實作05_擴充選單與自訂按鈕.html` | `實作05_擴充選單與自訂按鈕.gs` |
| 6 | `篇2_實作06_LineBot建立與推播.html` | `實作06_LineBot建立與推播.gs` |
| 7 | `篇2_實作07_低軌衛星氣象LineBot.html` | `實作07_低軌衛星氣象LineBot.gs` |
| 8 | `篇2_實作08_早安圖LineBot圖片推播.html` | `實作08_早安圖LineBot圖片推播.gs` |
| 9 | `篇2_實作09_AI自動生成簡報.html` | `實作09_AI自動生成簡報.gs` |
| 10 | `篇2_實作10_RSS新聞推播機器人.html` | `實作10_RSS新聞推播機器人.gs` |
| 11 | `篇3_實作11_LineBot訊息收集器.html` | `實作11_LineBot訊息收集器.gs` |
| 12 | `篇3_實作12_取得Line_User_ID.html` | （無獨立 .gs，教學用） |
| 13 | `篇3_實作13_活動報名即時通知.html` | `實作13_活動報名即時通知.gs` |
| 14 | `篇3_實作14_AI個人化Email通知.html` | `實作14_AI個人化Email通知.gs` |
| 15 | `篇3_實作15_自動化評量AI評語.html` | `實作15_自動化評量AI評語.gs` |
| 16 | `篇4_實作16_RAG知識庫問答機器人.html` | `實作16_RAG知識庫問答機器人.gs` |
| 17 | `篇4_實作17_出缺席通知群組推播.html` | `實作17_出缺席通知群組推播.gs` |
| 18 | `篇4_實作18_AI自動發文系統.html` | `實作18_AI自動發文系統.gs` |
| 19 | `篇4_實作19_查詢系統雙入口.html` | `實作19_查詢系統雙入口.gs` |
| 20 | `篇4_實作20_AI段考審題助手.html` | `實作20_AI段考審題助手.gs` |

### 密碼閘門（password-gate.js）

所有 21 個 HTML 簡報在 `<body>` 開頭或 `</body>` 前載入 `password-gate.js`。

- **學員密碼**: 預設 `GAS2026`，可由管理員修改，存於 `localStorage('gas_course_pwd')`
- **管理員密碼**: `Aa@0981737608`，可修改學員密碼
- **驗證狀態**: `sessionStorage('gas_course_auth')`

### 設計規範

- **主題色**: `#1b5e20`（深綠）、`#004d40`（青綠）、`#0f9b0f`（亮綠）、`#69f0ae`（accent）
- **漸層**: `linear-gradient(135deg, #1b5e20 0%, #004d40 100%)`
- **字體**: `'Microsoft JhengHei', 'Segoe UI', Arial, sans-serif`
- **content-card**: `border-radius: 24px`、`max-height: 88vh`、`overflow-y: auto`
- **info box**: `background: #e8f5e9`、`border-left: 6px solid #4caf50`
- **warning box**: `background: #fff3cd`、`border-left: 6px solid #ffc107`

### GAS 程式碼慣例

- 檔案開頭為區塊註釋（功能說明 + 步驟指南 + 設定區）
- 配置常數：`LINE_TOKEN`、`LINE_USER_ID`、`SPREADSHEET_ID`、`GEMINI_API_KEY` 等
- 區塊用 `====` 分隔線
- Line Bot 共用推播函式：`pushText`、`pushImage`、`pushLine`
- 回覆函式：`replyText`、`replyLine`
- 測試函式：`testPush`、`testSystem`
- AI 實作使用 Gemini API（`gemini-2.5-flash` 模型）

### AI 穿插實作（「參差」設計）

以下 4 個實作為 AI 主題，穿插於各篇之間：
- 實作 9：AI 自動生成簡報（SlidesApp + Gemini）
- 實作 14：AI 個人化 Email 通知（GmailApp + Gemini）
- 實作 15：自動化評量 AI 評語（Gemini + 試算表）
- 實作 20：AI 段考審題助手（DocumentApp + Gemini）

## 注意事項

- `參考/`、`截圖說明/`、`兩天GAS教學規劃.md` 在 `.gitignore` 中，不會推送
- `GAS程式碼/` 下有 zip 檔（打包版），在 `.gitignore` 中排除
- 簡報 HTML 檔案很大（30-85 KB），修改時注意只改動必要部分
- `修改規畫建議書.md` 記錄了課程重組的完整規劃
