# Renamer — 薄索引
跨平台規則正本：`~/.agents/institution/`（先讀 core/PRINCIPLES.md，照其指示附版本標記）。

## 專案專屬
- Build/test 指令：無 npm；直接用 `clasp push` 部署至 Google Apps Script；`tests/` 下有 GAS 單元測試
- 架構一句話：Google Apps Script（V8 執行環境）批次重命名工具，以 Google Sheets 為操作介面，操作 Google Drive 檔案
- 本專案禁區：不得在 `src/` 中引入 npm 套件（GAS 環境無 npm）；勿動 `appsscript.json` 的 timeZone 設定
