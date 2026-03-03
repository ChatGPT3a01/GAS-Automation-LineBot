# GAS 自動化 Line Bot 實戰教學

<div align="center">

<h3>Google Apps Script + Line Bot 從零到上線</h3>

<p>
<img src="https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white" alt="GAS">
<img src="https://img.shields.io/badge/Line%20Bot-00C300?style=for-the-badge&logo=line&logoColor=white" alt="Line Bot">
<img src="https://img.shields.io/badge/Vibe%20Coding-FF6B6B?style=for-the-badge&logo=sparkles&logoColor=white" alt="Vibe Coding">
</p>

<p><strong>8 個實作專案 | 122+ 張投影片 | 完整 GAS 程式碼</strong></p>

</div>

---

## 課程簡介

本教材帶領學員從零開始，透過 **Google Apps Script (GAS)** 搭配 **Line Bot**，完成 7 個完整的自動化實作專案。每個專案都附有互動式簡報與完整程式碼，並融入 **Vibe Coding（氣氛編碼）** 教學法，讓學員體驗用自然語言引導 AI 生成程式碼的全新開發方式。

## 課程大綱（Day 1）

| 單元 | 主題 | 簡報 | 程式碼 |
|:---:|------|:----:|:------:|
| Part 0 | GAS 環境建置與 Line Bot 建立 | [開啟簡報](簡報/Day1_Part0_GAS環境與LineBot建立.html) | [下載](GAS程式碼/00_LineBot_共用模組.gs) |
| Part 1 | RAG 知識庫 Line Bot | [開啟簡報](簡報/Day1_Part1_RAG知識庫.html) | [下載](GAS程式碼/01_RAG知識庫_LineBot.gs) |
| Part 2 | Line Bot 訊息收集器 | [開啟簡報](簡報/Day1_Part2_LineBot訊息收集器.html) | [下載](GAS程式碼/02_LineBot訊息收集器.gs) |
| Part 3 | 第一個 GAS 程式 | [開啟簡報](簡報/Day1_Part3_第一個GAS.html) | [下載](GAS程式碼/03_第一個GAS_LineBot.gs) |
| Part 4 | HTML 表單應用 | [開啟簡報](簡報/Day1_Part4_HTML表單.html) | [主程式](GAS程式碼/04_HTML表單_主程式.gs) / [前端](GAS程式碼/04_HTML表單_前端.html) |
| Part 5 | 寫入試算表 | [開啟簡報](簡報/Day1_Part5_寫入試算表.html) | [下載](GAS程式碼/05_寫入試算表_LineBot.gs) |
| Part 6 | 天氣預報 Line Bot | [開啟簡報](簡報/Day1_Part6_天氣預報.html) | [下載](GAS程式碼/06_天氣預報_LineBot.gs) |
| Part 7 | 自動化評量系統 | [開啟簡報](簡報/Day1_Part7_自動化評量.html) | [下載](GAS程式碼/07_自動化評量_LineBot.gs) |

## 專案架構

```
GAS-Automation-LineBot/
├── README.md                    # 本檔案
├── 作者資訊.png                  # 作者介紹
├── Day1_教材.md                 # Day 1 完整教材文件
│
├── 簡報/                        # 互動式 HTML 簡報
│   ├── index.html               # 課程總覽頁
│   ├── Day1_Part0_GAS環境與LineBot建立.html
│   ├── Day1_Part1_RAG知識庫.html
│   ├── Day1_Part2_LineBot訊息收集器.html
│   ├── Day1_Part3_第一個GAS.html
│   ├── Day1_Part4_HTML表單.html
│   ├── Day1_Part5_寫入試算表.html
│   ├── Day1_Part6_天氣預報.html
│   └── Day1_Part7_自動化評量.html
│
└── GAS程式碼/                    # Google Apps Script 原始碼
    ├── 00_LineBot_共用模組.gs     # Line Bot 共用函式庫
    ├── 01_RAG知識庫_LineBot.gs    # RAG 知識庫實作
    ├── 02_LineBot訊息收集器.gs    # 訊息收集器實作
    ├── 03_第一個GAS_LineBot.gs    # 基礎 GAS 實作
    ├── 04_HTML表單_主程式.gs      # HTML 表單後端
    ├── 04_HTML表單_前端.html      # HTML 表單前端
    ├── 05_寫入試算表_LineBot.gs   # 試算表寫入實作
    ├── 06_天氣預報_LineBot.gs     # 天氣預報實作
    └── 07_自動化評量_LineBot.gs   # 自動化評量系統
```

## 使用方式

### 瀏覽簡報
1. 點擊上方表格中的「開啟簡報」連結
2. 使用鍵盤左右鍵 `←` `→` 或畫面兩側箭頭切換投影片
3. 每個 Part 最後一頁有連結可前往下一個實作

### 使用 GAS 程式碼
1. 開啟 [Google Apps Script](https://script.google.com)
2. 建立新專案
3. 將對應的 `.gs` 程式碼貼入編輯器
4. 依照簡報說明設定必要的環境變數（如 Line Channel Access Token）
5. 部署為網路應用程式

## 技術棧

- **Google Apps Script** — 雲端自動化腳本
- **Line Messaging API** — 聊天機器人互動
- **Google Sheets API** — 試算表資料存取
- **Google Drive API** — 雲端檔案管理
- **UrlFetchApp** — 外部 API 串接（天氣預報、AI 模型）
- **HtmlService** — 網頁表單與前端介面

## 學習目標

完成本課程後，學員將能夠：

- 建立並部署 Google Apps Script 專案
- 串接 Line Bot Messaging API 收發訊息
- 使用 Google Sheets 作為資料庫
- 建立 HTML 網頁表單與後端互動
- 串接外部 API（天氣預報、AI 模型）
- 運用 Vibe Coding 方法讓 AI 協助開發

---

## 👨‍🏫 關於作者

<div align="center">

### 曾慶良 主任（阿亮老師）

<img src="作者資訊.png" width="600" alt="作者資訊">

<br>

<table>
<tr>
<td width="50%">

**📌 現任職務**

🎓 新興科技推廣中心主任<br>
🎓 教育部學科中心研究教師<br>
🎓 臺北市資訊教育輔導員

</td>
<td width="50%">

**🏆 獲獎紀錄**

🥇 2025年 SETEAM教學專業講師認證<br>
🥇 2024年 教育部人工智慧講師認證<br>
🥇 2022、2023年 指導學生XR專題競賽特優<br>
🥇 2022年 VR教材開發教師組特優<br>
🥇 2019年 百大資訊人才獎<br>
🥇 2018、2019年 親子天下創新100教師<br>
🥇 2018年 臺北市特殊優良教師<br>
🥇 2017年 教育部行動學習優等

</td>
</tr>
</table>

<br>

### 📞 聯絡方式

[![YouTube](https://img.shields.io/badge/YouTube-@Liang--yt02-red?style=for-the-badge&logo=youtube)](https://www.youtube.com/@Liang-yt02)
[![Facebook](https://img.shields.io/badge/Facebook-3A科技研究社-blue?style=for-the-badge&logo=facebook)](https://www.facebook.com/groups/2754139931432955)
[![Email](https://img.shields.io/badge/Email-3a01chatgpt@gmail.com-green?style=for-the-badge&logo=gmail)](mailto:3a01chatgpt@gmail.com)

</div>

## 📜 授權聲明

**© 2026 阿亮老師 版權所有**

本專案僅供「阿亮老師課程學員」學習使用。

### ⚠️ 禁止事項

- ❌ 禁止修改本專案內容
- ❌ 禁止轉傳或散布
- ❌ 禁止商業使用
- ❌ 禁止未經授權之任何形式使用

如有任何授權需求，請聯繫作者。

---

<div align="center">

<br>

## 🚀 立即開始學習

<br>

<a href="https://chatgpt3a01.github.io/GAS-Automation-LineBot/簡報/index.html">
<img src="https://img.shields.io/badge/%F0%9F%8E%93_%E9%BB%9E%E6%AD%A4%E9%80%B2%E5%85%A5%E8%AA%B2%E7%A8%8B%E7%B0%A1%E5%A0%B1-1b5e20?style=for-the-badge&labelColor=1b5e20&color=4caf50&logoColor=white" alt="進入課程簡報" height="60">
</a>

<br><br>

<table>
<tr>
<td align="center" width="140">
<a href="https://chatgpt3a01.github.io/GAS-Automation-LineBot/簡報/Day1_Part0_GAS環境與LineBot建立.html">
<img src="https://img.shields.io/badge/Part_0-環境建置-2e7d32?style=flat-square" alt="Part 0"><br>
<sub>GAS 環境 + Line Bot</sub>
</a>
</td>
<td align="center" width="140">
<a href="https://chatgpt3a01.github.io/GAS-Automation-LineBot/簡報/Day1_Part1_RAG知識庫.html">
<img src="https://img.shields.io/badge/Part_1-RAG知識庫-2e7d32?style=flat-square" alt="Part 1"><br>
<sub>知識庫 API</sub>
</a>
</td>
<td align="center" width="140">
<a href="https://chatgpt3a01.github.io/GAS-Automation-LineBot/簡報/Day1_Part2_LineBot訊息收集器.html">
<img src="https://img.shields.io/badge/Part_2-訊息收集-2e7d32?style=flat-square" alt="Part 2"><br>
<sub>訊息收集器</sub>
</a>
</td>
<td align="center" width="140">
<a href="https://chatgpt3a01.github.io/GAS-Automation-LineBot/簡報/Day1_Part3_第一個GAS.html">
<img src="https://img.shields.io/badge/Part_3-第一個GAS-2e7d32?style=flat-square" alt="Part 3"><br>
<sub>GAS 報表</sub>
</a>
</td>
</tr>
<tr>
<td align="center" width="140">
<a href="https://chatgpt3a01.github.io/GAS-Automation-LineBot/簡報/Day1_Part4_HTML表單.html">
<img src="https://img.shields.io/badge/Part_4-HTML表單-2e7d32?style=flat-square" alt="Part 4"><br>
<sub>表單應用</sub>
</a>
</td>
<td align="center" width="140">
<a href="https://chatgpt3a01.github.io/GAS-Automation-LineBot/簡報/Day1_Part5_寫入試算表.html">
<img src="https://img.shields.io/badge/Part_5-寫入試算表-2e7d32?style=flat-square" alt="Part 5"><br>
<sub>簽到系統</sub>
</a>
</td>
<td align="center" width="140">
<a href="https://chatgpt3a01.github.io/GAS-Automation-LineBot/簡報/Day1_Part6_天氣預報.html">
<img src="https://img.shields.io/badge/Part_6-天氣預報-2e7d32?style=flat-square" alt="Part 6"><br>
<sub>每日推播</sub>
</a>
</td>
<td align="center" width="140">
<a href="https://chatgpt3a01.github.io/GAS-Automation-LineBot/簡報/Day1_Part7_自動化評量.html">
<img src="https://img.shields.io/badge/Part_7-自動化評量-2e7d32?style=flat-square" alt="Part 7"><br>
<sub>AI 自動評語</sub>
</a>
</td>
</tr>
</table>

<br>

</div>

---

<div align="center">

## 🌟 喜歡這個專案嗎？

如果這個工具對您有幫助，請給我們一個 ⭐ Star！

<br>

**Made with ❤️ by 阿亮老師**

<br>

[⬆️ 回到頂部](#gas-自動化-line-bot-實戰教學)

---

© 2026 阿亮老師 版權所有

</div>
