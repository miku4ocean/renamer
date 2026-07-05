# HANDOFF — Renamer
更新：2026-07-05／claude

## 目前目標
Google Apps Script 批次更名工具，功能完整，維護與文件完善階段。

## 狀態
- 已完成：BatchRename.js、Code.js、FileOperations.js 核心邏輯；多樣命名規則（取代、序號、日期、大小寫）；`docs/` 批次更名規則說明
- 進行中：工作區乾淨，無未 commit 修改
- 驗收現況：未驗證（需部署至 GAS 環境手動測試）

## 下一步（接手的人從這裡開始）
1. 安裝 clasp：`npm install -g @google/clasp`，登入後 `clasp push`
2. 在 Google Sheets 綁定此 Script，手動執行選單項目確認功能
3. 若需新規則，在 `src/BatchRename.js` 照現有模式新增

## 地雷（別踩）
- GAS 有執行時間上限（6 分鐘），批次大量檔案需分批呼叫
- `templates/` 目錄存放試算表範本，勿刪除（使用者需匯入作為起點）
- `appsscript.json` 的 `oauthScopes` 需精確，過寬會觸發 Google Workspace 安全審查

## 主辦權
單線／待分派
