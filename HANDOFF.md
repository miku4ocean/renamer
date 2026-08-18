# HANDOFF — Renamer
更新：2026-08-18／claude（修上一輪發現的 3 個真 bug，用 Node.js vm 載入 src/*.js 跑 fixture 驗證，未動 Google 帳號）

## 目前目標
Google Apps Script 批次更名工具，功能完整，維護與文件完善階段。

## 狀態
- 已完成：BatchRename.js、Code.js、FileOperations.js 核心邏輯；多樣命名規則（取代、序號、日期、大小寫）；`docs/` 批次更名規則說明
- 本次修完上一輪發現的 3 個真 bug（見下），29 個既有測試斷言全過，新增 7 個
- 驗收現況：仍未在真實 GAS 環境跑過（需部署手動測試）；本次以 Node.js vm 載入 `src/*.js`（Utilities/Session 打樁）跑 fixture 驗證邏輯層

## 本次修掉的 3 個真 bug

1. **重新命名靠檔名重查，不是靠已抓到的 File ID（TOCTOU + 改錯檔）**——已修。
   - `populateFileList()`（FileOperations.js）現在把 `file.id` 一起寫進「檔名變更區」的第 8 欄（H欄／檔案ID），跟掃描時 `getFilesFromFolder()` 抓到的 ID 是同一份。
   - `executeRenaming()` 改成直接讀 H 欄的 ID 呼叫 `DriveApp.getFileById(fileId)`，不再用檔名/資料夾名稱重新搜尋 Drive；缺 ID 時該筆直接報錯（`缺少檔案 ID，請重新執行「讀取資料夾檔案」`），不會誤改到別的同名檔案。
   - `applyRulesToExistingFiles()` 也同步保留/回寫 H 欄，避免重套規則時把 ID 弄丟。
   - 已完全刪除易出錯的 `getFileIdByNameAndPath()`（原本會 fallback 成 `DriveApp.getFoldersByName()` 掃全 Drive 同名資料夾）。
   - 連帶更新的文件（讓新架構有文件可查）：`docs/setup-guide.md`、`docs/api-reference.md`、`docs/architecture.mmd/.svg/.html`、`templates/filelist-sheet-template.md`、`templates/google-sheets-formatting-guide.md`、`templates/renamer-template.html`、`templates/create_excel_template.py`、`templates/檔名變更區.csv`，全部補上 H欄／檔案ID 的說明（建議隱藏、不可刪除）。

2. **正則表達式注入／ReDoS**——已修。
   - 新增 `escapeRegExp()`（FileOperations.js），把使用者輸入的「部分取代」尋找字串在丟進 `new RegExp()` 前先跳脫特殊字元。
   - `FileOperations.js` 的 `applyReplaceTextRule()` 與 `BatchRename.js` 的 `applyReplaceText()`（死碼但同款漏洞）都已套用；後者靠 GAS 專案共用全域作用域直接呼叫前者定義的 `escapeRegExp`。
   - 驗證：`report.v1.final` 用「.」取代「_」現在正確回傳 `report_v1_final`（跳脫前會整串被吃光）。

3. **測試套件本身壞的 + 生產路徑零測試**——已修。
   - `tests/test-functions.js` 的 `.hidden-file` 案例預期值改成跟程式碼實際行為一致（`expectedName: '.hidden-file', expectedExt: ''`）。
   - 新增 `testProductionApplyRenameRule()`（涵蓋 `FileOperations.js` 的生產路徑 `applyRenameRule`：取代文字含正則特殊字元、完全取代、新增序號、大小寫轉換）與 `testEscapeRegExp()`，並掛進 `runTests()` 主流程。之前 `runTests()` 只測到沒人呼叫的死碼 `applyBatchRename`（BatchRename.js）。

## 順便確認沒問題的項目
- 無 `eval`/`new Function`，`src/` 沒有 Node.js 或瀏覽器限定 API（`require`/`fetch`/`document`/`window` 皆無），GAS 相容性乾淨
- 中文檔名、emoji（astral surrogate pair）檔名跑過大小寫轉換規則不會被截斷/corrupt
- `applyBatchRename`（BatchRename.js）整組函式仍是死碼（`Code.js`/`FileOperations.js` 沒人呼叫），本次沒有處理「刪除死碼」這件事，留給下一輪評估

## 下一步（接手的人從這裡開始）
1. 安裝 clasp：`npm install -g @google/clasp`，登入後 `clasp push`（**注意：目錄下無 `.clasp.json`，需先 `clasp create`/`clasp clone` 建立**，progress.md 已記錄此缺口）
2. 在 Google Sheets 綁定此 Script，手動執行選單項目確認功能——**尤其要驗證新的 H 欄（檔案ID）**：讀取資料夾檔案後 H 欄應自動填入 Drive 檔案 ID，重新命名後該 ID 對應的檔案要正確被改名
3. 評估是否直接刪除死碼 `applyBatchRename`（BatchRename.js）及其專屬輔助函式，或保留作為未來重構的參考
4. 若需新規則，在 `src/BatchRename.js` 照現有模式新增（但先評估併入 `FileOperations.js`，見 progress.md J 段）

## 地雷（別踩）
- GAS 有執行時間上限（6 分鐘），批次大量檔案需分批呼叫
- `templates/` 目錄存放試算表範本，勿刪除（使用者需匯入作為起點）；本次已同步更新其中的欄位結構，記得若再改 schema 要一併更新這批文件
- `appsscript.json` 的 `oauthScopes` 需精確，過寬會觸發 Google Workspace 安全審查
- H 欄（檔案ID）是重新命名機制的核心，任何改動「檔名變更區」欄位配置的程式碼都要連帶檢查 `executeRenaming`/`applyRulesToExistingFiles`/`populateFileList` 三者的欄位索引是否還對得上

## 主辦權
單線／待分派
