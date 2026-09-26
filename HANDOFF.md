# HANDOFF — Renamer
更新：2026-09-26／Claude Opus 5.5（依規劃檔 `plan-90/Renamer.md` 依序做完 WP1–WP4，
均為先紅後綠：先用 node:vm 探針重現 bug，修完再驗證，未動 Google 帳號、未執行 clasp push）

## 目前目標
Google Apps Script 批次更名工具，功能完整，維護與文件完善階段。程式碼層預估已到規劃檔說的
90%（WP1–WP4 全數完成）；再往上需要 clasp push 到真實 GAS 環境手動測試，屬下面「卡點」。

## 本機測試（本輪新增，之前沒有）
```
npm test
```
`tests/run-node.mjs` 用 `node:vm` 把 `src/Code.js`、`src/FileOperations.js`、
`tests/test-functions.js` 載進同一個假環境執行 `runTests()`：
- 打樁 `Utilities.formatDate`／`Session.getScriptTimeZone`（固定 Asia/Taipei）
- 提供 `makeSheet(rows)`：假 Sheet，支援 `getRange`（A1 或數字座標兩種呼叫方式）、
  `getValues`/`setValues`/`setValue`/`clearContent`、`getLastRow`/`getLastColumn`
- 提供假 `DriveApp`（`_registerFile`/`_registerFolder`/`_reset`，記錄 `setName`／`makeCopy`
  呼叫到 `DriveApp._calls`）與假 `SpreadsheetApp`（`_setSheet(name, sheet)`）
- `runTests()` 原本把測試失敗吞掉只 `console.error`，已改成重新拋出，`npm test` 會用
  exit code 正確反映測試結果（0＝全過，1＝有失敗）

**重要限制**：`tests/test-functions.js` 裡涉及 `makeSheet`／假 `DriveApp`／假
`SpreadsheetApp` 的測試（`testPopulateFileListSerialNumbering`、
`testApplyRulesToExistingFilesSequentialNumbering`、`testExecuteRenamingIdempotency`）
開頭都有 `typeof ... === 'undefined'` 的偵測，如果整份檔案直接貼進真正的 GAS
指令碼編輯器執行（沒有這些假物件，全域的 `DriveApp`/`SpreadsheetApp` 是真正的服務），
會印出「⏭️ 跳過」直接略過，不會誤觸真實 Drive/Sheets 操作，也不會因為呼叫到不存在
的方法而噴錯誤。`testFindDuplicateTargets` 是純函數測試，兩邊都會正常執行。

## 本輪做完的 4 個工作包（WP1–WP4）

### WP1：補進 Node 測試入口
新增 `tests/run-node.mjs`、`package.json`（只有 test script，無 dependencies）。之前
`runTests()` 測試失敗會被吞掉，已修成重新拋出。

### WP2：修序號 index 與取代規則的 3 個真 bug
1. `populateFileList()`（`src/FileOperations.js`）：`files.map(file => ...)` 沒把
   `index` 傳給 `applyRenameRule`，導致「讀取資料夾檔案」流程下新增序號規則全部拿到
   同一個號碼（001_...），已改成 `files.map((file, index) => ...)`。
2. `applyReplaceTextRule()`：原本用 `new RegExp(escapeRegExp(...))` + `String#replace`，
   有三個問題：取代文字裡的 `$&`/`$1` 會被當成正則特殊樣式、尋找字串為空時會在每個字元
   間插入、參數裡有多個 `→` 時第二個以後的內容會被丟掉。改用
   `baseName.split(literalFindText).join(replaceText)`：只切第一個 `→`（後面的 `→`
   保留在取代文字裡）、空尋找字串原樣回傳、取代文字一律當字面值處理，順帶徹底解決
   ReDoS／正則注入疑慮。`escapeRegExp` 保留給既有測試用，不再被生產路徑呼叫。
3. `applyRulesToExistingFiles()`：遇到空白列 `continue` 時序號會跳號，改成獨立計數器
   `seq`，只在真正處理到的列遞增。

`docs/batch-rename-rules-guide.md` 已補上取代規則行為細節（字面比對、只切第一個箭頭、
空尋找字串處理）。

### WP3：executeRenaming 冪等化＋I 欄執行結果＋時間預算
「檔名變更區」新增 I 欄（執行結果），H 欄仍是檔案 ID，欄位改動已同步檢查
`populateFileList`／`applyRulesToExistingFiles`／`executeRenaming` 三處（8 欄改 9 欄）。
`executeRenaming` 每處理完一列立即 `setValue` 寫入 I 欄（`✓ 已更名 yyyy-MM-dd HH:mm`／
`✓ 已複製 <新檔案ID>`／`✗ <錯誤訊息>`），開跑時 I 欄已是 `✓` 開頭的列直接跳過（冪等，
重跑不會重複複製／重新命名）。新增 `options.now`／`options.budgetMs`（預設 5 分鐘）算
deadline，超時就停止並回傳 `{timedOut:true, message:"已處理 N 列，剩餘 M 列，請再執行
一次「開始重新命名」"}`，不當成錯誤丟出。**回傳值型別已改變**：`executeRenaming` 現在
回傳物件 `{timedOut, successCount, errorCount, processedCount, errors, message}`，不再
是單純的數字，`Code.js` 的 `startRenaming()` 已同步更新依 `timedOut`/`errorCount` 顯示
不同的 alert。`templates/` 下 4 個檔案已同步補上 I 欄說明。

### WP4：執行前預檢重複目標檔名
新增 `findDuplicateTargets(data, operationType)`（純函數）：依「變更後檔名」分組（不分
大小寫），排除已完成（I 欄 ✓）的列，回傳達 2 筆以上的重複組。`startRenaming()` 在原本
的確認對話框前新增兩道檢查：C 欄空白直接擋下（不給確認）；有重複目標檔名則列出前 5 組
用 YES_NO 再次確認。

## 沒做的部分（規劃檔標 [H]／[D]／[X]，本輪刻意跳過）
- **WP5**（clasp 部署前準備）：新增 `.claspignore`／`.clasp.json.example` 未做。
- **WP6**（清除 27 處過期的 `BatchRename` 文件參照）：`HANDOFF.md`（本檔已在改寫時順便
  清掉自己的舊參照，其餘檔案未動）、`progress.md`、`docs/setup-guide.md`、
  `docs/api-reference.md`、`docs/architecture.mmd/.html/.svg`、`mockup/sheet-command.html`、
  `tests/test-functions.js` 的 `testErrorHandling`/`generateTestReport` 裡仍有
  `BatchRename`／`batchRenameLogic` 字樣未清。`grep -rn "BatchRename" . --exclude-dir=.git`
  目前還有多筆。
- **WP7**（`appsscript.json` 明確宣告 `oauthScopes`）：規劃檔標示「預設不做，由使用者
  決定」，本輪沒有動 `appsscript.json`。

## 卡點（跳出程式碼層，需要人／帳號才能做）
- `npm i -g @google/clasp`、`clasp login`、`clasp create --type sheets`（或 clone 既有
  腳本拿 scriptId）、`clasp push`——本專案沒有 `.clasp.json`，且未安裝 clasp。
- 在真實 Google Sheets／Drive 手動跑一遍：讀取→套用→重新命名、複製後更名，驗證 H 欄
  ID、**I 欄執行結果**、超時續跑三件事在真實環境下的行為（本輪只在 node:vm 假環境驗證）。
- `appsscript.json` 的 `oauthScopes` 要不要明確宣告（WP7），等實機測試時一起處理。

## 地雷（別踩）
- GAS 有執行時間上限（6 分鐘）；`executeRenaming` 已用 `options.budgetMs` 處理，預設
  5 分鐘，留 1 分鐘餘裕。
- `templates/` 目錄存放試算表範本，勿刪除；本輪已同步更新其中的欄位結構（新增 I 欄），
  若再改 schema 要一併更新這批文件。
- `appsscript.json` 的 `oauthScopes` 需精確，過寬會觸發 Google Workspace 安全審查。
- **欄位索引地雷**：任何改動「檔名變更區」欄位配置的程式碼，都要同時檢查
  `populateFileList`／`applyRulesToExistingFiles`／`executeRenaming`／
  `findDuplicateTargets` 這四處的欄位索引是否還對得上（目前是 9 欄：A~I）。
- `tests/run-node.mjs` 是本機測試跑者，不是給 clasp push 的（WP5 的 `.claspignore` 要
  排除它，還沒做）。
- `executeRenaming` 的回傳值已經是物件不是數字，如果之後有其他地方直接呼叫
  `executeRenaming()` 並把回傳值當數字用，會壞掉。

## 下一步（接手的人從這裡開始）
1. 先跑 `npm test` 確認綠燈（無需 npm install，純 Node 內建模組）。
2. 做 WP5：`.claspignore`（排除 `docs/**`/`mockup/**`/`templates/**`/
   `tests/run-node.mjs`/`*.md`/`package.json`）、`.clasp.json.example`、`.gitignore`
   加 `.clasp.json`。
3. 做 WP6：清掉 `progress.md`/`docs/setup-guide.md`/`docs/api-reference.md`/
   `docs/architecture.*`/`mockup/sheet-command.html`/`tests/test-functions.js` 裡
   殘留的 `BatchRename` 參照（`BatchRename.js` 本身在更早的輪次已經整組刪除，這些只是
   文件沒跟上）。
4. 安裝 clasp、`clasp create`/`clasp clone` 取得 scriptId、`clasp push`，在真實
   Google Sheets 綁定此腳本，手動驗證 I 欄與超時續跑機制。
5. 若要動 WP7 的 `oauthScopes`，建議跟第 4 步的實機驗證一起做。

## 主辦權
單線／待分派
