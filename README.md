# Schooler 講師預計錄製排程

講師填寫課程資訊與每堂預計交付日期，送出後寫入 Google 試算表「講師回覆錄製日與檢核日」，並可下載 PDF 排程總表。

- 網站：GitHub Pages（本 repo 根目錄的 `index.html`）
- 後端：Google Apps Script 專案「Schooler 課程錄製排程網站」（`apps-script/Code.gs` 為程式碼備份；Apps Script 的舊網址只會引導到本網站）

## 運作方式

1. `index.html` 以 `fetch` POST（`text/plain` 的 JSON）到 Apps Script 網頁應用程式的 `/exec` 網址。
2. Apps Script `doPost` 驗證資料後寫入：
   - **課程總表**：一列，課程狀態「待開始」、來源「網站」，補充說明寫入「內部備註」
   - **堂次排程**：每堂一列，堂次從 0 起算（總時數 10 → 第 0–10 堂，共 11 堂），檢核狀態「尚未檢核」
3. **工作檢核日**：交付日的下一個工作日；交付日為週五時取第 3 個工作日，跳過週末與「國定假日」分頁中停班的日期。

## 修改後端

在 Apps Script 編輯器修改程式碼後，到「部署 → 管理部署作業 → 編輯 → 版本：建立新版本 → 部署」，網址不變。

## 修改網站

直接編輯 `index.html` 並推送到 `main`，GitHub Pages 會自動更新。
