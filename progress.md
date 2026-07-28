# Renamer — 專案進度報告

> 本檔依實際讀取到的 `HANDOFF.md`、`AGENTS.md`、`CLAUDE.md`、`README.md`、`docs/`、`templates/`、`src/`、`tests/`、`appsscript.json` 撰寫。查無佐證之處一律標註「未確認」，不臆測。

## A. 專案名稱
Renamer

## B. 專案路徑
`/Users/leonalin/Code/Renamer`

## C. 專案簡介
Renamer 是一個以 **Google Apps Script（V8 執行環境）** 為基礎的雲端批次檔案重新命名工具。它沒有獨立的網頁前端或後端伺服器，而是把 **Google Sheets 活頁簿當成操作介面**：使用者在「指令區」工作表填入 Google Drive 資料夾與命名規則，工具透過 `DriveApp`／`SpreadsheetApp` 兩個 Apps Script 內建服務讀取檔案、產生預覽、最後直接對使用者自己的 Google Drive 檔案執行更名或複製更名。

## D. 專案開發目的
依 `README.md` 與 `HANDOFF.md`：目標是做一個「不需寫程式、只要會用 Google Sheets 就能批次改檔名」的工具，鎖定 Google Workspace／Google Drive 的使用情境，透過選單操作取代逐一手動改檔名。`HANDOFF.md` 記錄目前階段為「功能完整，維護與文件完善階段」。

## E. 解決使用者痛點
- 大量檔案需要逐一手動重新命名，耗時且容易出錯。
- 命名規則（前綴、序號、日期等）若靠手動輸入，難以保持一致。
- 一般批次改名工具「按下去就全部改完」，缺乏改名前的檢視與修正機會，容易誤改、誤覆蓋檔案。
- 不想額外安裝軟體或申請新的雲端服務帳號——直接沿用既有的 Google 帳號與 Drive 權限。

## F. 專案功能細項介紹
- **選單系統**（`src/Code.js` `onOpen()`）：在 Google Sheets 頂部建立「檔案重新命名工具」自訂選單，含四個項目：
  - 讀取資料夾檔案（`loadFolderFiles()`）：依「指令區」B2 的資料夾 ID／URL 掃描 Google Drive 資料夾（排除 Google 試算表本身），寫入「檔名變更區」。
  - 套用批次規則（`applyBatchRules()`）：對「檔名變更區」既有資料重新套用目前設定的命名規則，不重新掃描 Drive。
  - 開始重新命名（`startRenaming()`）：跳出 YES/NO 確認對話框，確認後才真正呼叫 Drive 執行更名／複製。
  - 清除所有資料（`clearAllData()`）：確認後清空「檔名變更區」第 2 列以下的資料。
- **五種命名規則**（`src/BatchRename.js` / `src/FileOperations.js`）：
  - 新增文字（前綴／後綴／前後皆加）
  - 取代文字（完全取代／部分取代，支援正規表示式取代）
  - 大小寫轉換（全部大寫／全部小寫／首字母大寫）
  - 新增序號（前綴序號／後綴序號／插入序號，可設起始數字與位數）
  - 格式化日期（依檔案最後修改時間，支援 `YYYY-MM-DD`／`YYYYMMDD`／`DD-MM-YYYY`）
- **雙重操作模式**：原位置更名（直接改名） vs 複製後更名（複製到目標資料夾並命名新檔，保留原檔）。
- **人工預覽機制**：「檔名變更區」C 欄（變更後檔名）可手動編輯，模板文件並規劃了重複檔名／不合法字元的條件格式警示（`templates/filelist-sheet-template.md`）。
- **測試套件**（`tests/test-functions.js`）：`runTests()` 涵蓋檔名工具函式、五種命名規則、資料夾 ID 擷取；另有 `runPerformanceTest()`（1000 筆檔案效能測試）與 `testErrorHandling()`，須在 Apps Script 編輯器手動執行。

## G. 專案規格及 RPD

**技術棧**
- Google Apps Script，執行環境 V8（`appsscript.json` `runtimeVersion: "V8"`）
- 時區設定：`Asia/Taipei`
- 內建服務：`SpreadsheetApp`、`DriveApp`、`Utilities`（無 `enabledAdvancedServices`）
- 無前端框架、無 npm 套件、無資料庫、無自建伺服器、無第三方 API（`AGENTS.md` 明文禁止在 `src/` 引入 npm 套件，因 GAS 環境無 npm）

**部署方式／指令**
- 建議用 `clasp`：`npm install -g @google/clasp`，`clasp push` 部署程式碼到 Google Apps Script 專案（`HANDOFF.md`「下一步」第 1 項）。
- 專案目錄下**未發現** `.clasp.json`，代表尚未設定實際的 clasp 部署綁定（未確認是否曾在別處設定過）。
- 亦可手動複製 `src/*.js` 內容貼進 Apps Script 線上編輯器（`README.md`／`docs/setup-guide.md` 描述的路徑）。

**埠／服務位址**
不適用——本工具無獨立伺服器或本機服務，執行環境由 Google Apps Script 平台代管。

**資料流**
1. 使用者在「指令區」工作表填入來源／目標資料夾 ID、命名模式、參數、操作類型。
2. 選單「讀取資料夾檔案」→ `getFilesFromFolder()` 呼叫 `DriveApp` 掃描來源資料夾 → `populateFileList()` 寫入「檔名變更區」A/B/E/F/G 欄，並依目前規則預先產生 C/D 欄建議值。
3. 使用者可於「檔名變更區」C 欄手動調整變更後檔名，或用選單「套用批次規則」重新套用命名規則。
4. 使用者確認「檔名變更區」內容無誤後，執行選單「開始重新命名」→ 二次確認對話框 → `executeRenaming()` 依 C 欄逐列呼叫 `renameFile()` 或 `copyAndRenameFile()`，實際寫入 Google Drive。
5. 執行結果（成功／失敗筆數與錯誤訊息）以對話框回報給使用者。

**執行限制（HANDOFF「地雷」）**
- GAS 單次執行時間上限 6 分鐘，大量檔案須分批呼叫（程式碼中未見自動分批機制，見 I 段）。
- `appsscript.json` 的 `oauthScopes` 需精確設定，範圍過寬會觸發 Google Workspace 安全審查（目前 `appsscript.json` 未列出 `oauthScopes` 欄位，未確認實際授權範圍如何決定）。
- `templates/` 目錄的試算表範本不可刪除，是使用者建立活頁簿的起點。

## H. 目前已完成項目
- `src/Code.js`：選單建立與四個選單事件函數（`onOpen`、`loadFolderFiles`、`applyBatchRules`、`startRenaming`、`clearAllData`）皆已完整實作，含錯誤處理與確認對話框。
- `src/FileOperations.js`：資料夾 ID 解析、Drive 檔案掃描、寫入工作表、五種命名規則套用、更名／複製執行邏輯、依檔名+路徑反查 File ID 的容錯邏輯，皆已完成。
- `src/BatchRename.js`：命名規則的純函數版本（`applyBatchRename` 及各規則子函數），與 `FileOperations.js` 內近似邏輯並存，供 `tests/test-functions.js` 單元測試使用。
- 文件完整：`README.md`、`docs/setup-guide.md`、`docs/user-manual.md`、`docs/api-reference.md`、`docs/batch-rename-rules-guide.md`，以及 `templates/` 下的工作表模板說明與 CSV／HTML 範本。
- `tests/test-functions.js`：涵蓋核心函式的測試案例已撰寫完成。
- 本次新增：`docs/architecture.html`／`.svg`／`.mmd` 三份架構圖，`mockup/` 三頁線框稿＋索引頁，本 `progress.md`。

## I. 尚待完成項目
- **實機驗證**：`HANDOFF.md` 明載「驗收現況：未驗證（需部署至 GAS 環境手動測試）」——目前所有功能僅通過程式碼審視與 `tests/test-functions.js` 的邏輯測試，尚未在真正的 Google Sheets + Drive 環境跑過一次完整流程。
- **clasp 部署綁定**：目錄下未見 `.clasp.json`，代表尚未設定好可一鍵 `clasp push` 的部署設定，需按 `HANDOFF.md` 下一步第 1 項完成。
- **大量檔案分批處理**：`docs/setup-guide.md`「效能最佳化」建議「分批處理檔案」，但 `src/FileOperations.js` 的 `executeRenaming()` 是單一迴圈跑完整份清單，未見自動偵測執行時間並分批（超過 6 分鐘上限時的行為未確認）。
- **文件與程式碼不同步**：`docs/api-reference.md` 仍將 `startRenaming()` 標記為「狀態：開發中」，但 `src/Code.js` 中該函數已是完整實作（含確認對話框與錯誤處理），文件內容已過時，需更新。
- **`oauthScopes` 未列於 `appsscript.json`**：`AGENTS.md` 提醒此設定需精確、範圍過寬會觸發安全審查，但目前 `appsscript.json` 內容只有 `timeZone`、`dependencies`、`exceptionLogging`、`runtimeVersion`，未見明確的 `oauthScopes` 欄位（未確認 GAS 是否會依程式內實際呼叫的服務自動推斷）。
- 主辦權標示「單線／待分派」（`HANDOFF.md`），代表目前沒有指定的持續維護者。

## J. 系統優化或增加功能建議
- 在 `executeRenaming()` 中加入依耗時自動分批（例如每 N 筆檢查一次 `Utilities` 執行時間或改用時間驅動觸發器續跑），避免超過 GAS 6 分鐘上限時整批中斷在不可預期的位置。
- 「檔名變更區」目前只用條件格式**視覺提示**重複檔名／不合法字元，執行重新命名前可考慮加入**硬性阻擋**（例如偵測到重複或不合法字元時直接跳出錯誤、拒絕執行該筆），降低誤觸風險。
- 補上操作紀錄／復原機制：目前更名為即時生效且無內建復原功能，`docs/user-manual.md` 僅建議使用者自行截圖或備份；可考慮增加一個「操作紀錄」工作表，記錄每次執行的原檔名／新檔名對照，方便事後人工回復。
- 同步更新 `docs/api-reference.md`，移除 `startRenaming()` 的「開發中」標記，並反映目前實際簽章與行為。
- 建立 `.clasp.json`（或在文件中提供範本）並在 README／HANDOFF 補充「如何驗證已部署成功」的具體檢查步驟，降低新接手者的上手成本。
- `src/FileOperations.js` 與 `src/BatchRename.js` 存在兩套幾乎重複的命名規則實作（一套給實際執行用、一套給單元測試用），未來可評估合併為單一來源，避免兩邊邏輯日後修改時不同步。
