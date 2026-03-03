# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 專案概述

GAS 自動化 Line Bot 實戰教學課程，包含 7 個 Part（Part 0-6）的互動式 HTML 簡報與 Google Apps Script 程式碼。部署於 GitHub Pages。

- **GitHub Pages**: `https://chatgpt3a01.github.io/GAS-Automation-LineBot/簡報/index.html`
- **所有回應使用繁體中文**

## 開發與部署

此專案為純靜態 HTML，無建置流程、無套件管理、無測試框架。修改後直接 git push 到 master 即可透過 GitHub Pages 生效。

```bash
git add -A && git commit -m "描述" && git push origin master
```

## 架構

### 簡報系統

每個 Part 是獨立的單檔 HTML（含內嵌 CSS/JS），共用 `簡報/password-gate.js` 做密碼保護。Day 1 共 8 個 Part（Part 0-7）。

- **入口**: `簡報/index.html`（課程總覽頁）
- **投影片結構**: 每個 `.slide` 元素為一頁，透過 `opacity` + `z-index` 控制切換
- **導航**: 頂部 nav-bar、左右箭頭、底部 indicators、鍵盤左右鍵
- **每個 Part 最後一頁**: 固定三按鈕（下載程式碼 / 下一個實作 / 回到總覽）

### 檔案對應

| Part | 簡報 | GAS 程式碼 |
|------|------|-----------|
| 0 | `Day1_Part0_GAS環境與LineBot建立.html` | `00_LineBot_共用模組.gs` |
| 1 | `Day1_Part1_RAG知識庫.html` | `01_RAG知識庫_LineBot.gs` |
| 2 | `Day1_Part2_LineBot訊息收集器.html` | `02_LineBot訊息收集器.gs` |
| 3 | `Day1_Part3_第一個GAS.html` | `03_第一個GAS_LineBot.gs` |
| 4 | `Day1_Part4_HTML表單.html` | `04_HTML表單_主程式.gs` + `04_HTML表單_前端.html` |
| 5 | `Day1_Part5_寫入試算表.html` | `05_寫入試算表_LineBot.gs` |
| 6 | `Day1_Part6_天氣預報.html` | `06_天氣預報_LineBot.gs` |
| 7 | `Day1_Part7_自動化評量.html` | `07_自動化評量_LineBot.gs` |

### 密碼閘門（password-gate.js）

所有 8 個 HTML 檔案在 `</body>` 前載入 `password-gate.js`。

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
- 配置常數：`LINE_TOKEN`、`LINE_USER_ID`、`SPREADSHEET_ID` 等
- 區塊用 `====` 分隔線
- 共用推播函式：`pushText`、`pushImage`、`pushLine`
- 回覆函式：`replyText`、`replyLine`
- 測試函式：`testPush`、`testSystem`

## 注意事項

- `參考/`、`截圖說明/`、`兩天GAS教學規劃.md` 在 `.gitignore` 中，不會推送
- `GAS程式碼/` 下有 4 個 zip 檔（Part 3-6 打包版），也在 `.gitignore` 中排除（但目前為 untracked 狀態）
- 簡報 HTML 檔案很大（60-95 KB），修改時注意只改動必要部分
