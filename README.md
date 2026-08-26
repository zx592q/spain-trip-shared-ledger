# 西班牙之旅共用帳本

GitHub Pages 網站，透過 Google Apps Script 連接 Google Sheets 共用資料。

## 部署

1. 在 Google Sheet 的「擴充功能 → Apps Script」貼入 `Code.gs`。
2. 部署為網頁應用程式，執行身分選「我」，存取權限選可使用帳本的所有人。
3. 將 `/exec` 網址填入 `config.js`。
4. 在 GitHub Settings → Pages 選擇 main branch 與 root。
