# HANDOFF — Renamer
更新：2026-08-18／claude（靜態邏輯驗證，用 Node.js vm 載入 src/*.js 跑 fixture，未動 Google 帳號）

## 目前目標
Google Apps Script 批次更名工具，功能完整，維護與文件完善階段。

## 狀態
- 已完成：BatchRename.js、Code.js、FileOperations.js 核心邏輯；多樣命名規則（取代、序號、日期、大小寫）；`docs/` 批次更名規則說明
- 進行中：工作區乾淨，無未 commit 修改
- 驗收現況：未驗證（需部署至 GAS 環境手動測試）；本次以 Node.js vm 載入 `src/*.js`（Utilities/Session 打樁）跑 fixture 驗證邏輯層，發現 3 個真 bug（見下）

## 本次驗證發現的真 bug（未修，等分派）
1. **重新命名靠檔名重查，不是靠已抓到的 File ID**：`getFilesFromFolder()`（FileOperations.js:27）明明抓了 `file.getId()`，但 `populateFileList()` 寫入工作表的 7 欄完全沒存這個 id；`executeRenaming()` 真正執行時改呼叫 `getFileIdByNameAndPath(originalName, originalPath)`（FileOperations.js:337-374）用**檔名**在 Drive 上重新搜尋，來源資料夾若查無結果還會 fallback 成 `DriveApp.getFoldersByName(folderName)` **掃全 Drive同名資料夾**。使用者 Drive 內只要有兩個同名資料夾（例如兩個「2024」），或來源資料夾內本身就有同名檔案，就可能改到完全不是使用者選的那個檔案。且 `sourceFolderId` 是執行當下重讀「指令區」B2，若讀清單後又改了來源資料夾，也會用新資料夾查舊清單——TOCTOU。
2. **正則表達式注入／ReDoS，用 `(a+)+$` 這種樣式對長字串實測直接掛住 2 分鐘以上（process 被強制 kill）**：`applyReplaceTextRule()`（FileOperations.js:135-144，`BatchRename.js` 的 `applyReplaceText()` 也同款）把使用者填的「部分取代」尋找字串直接丟進 `new RegExp(findText, 'g')`，沒有 escape。除了 ReDoS，一般使用者想字面取代「.」這種常見符號（例如去掉檔名裡的句點）也會被當萬用字元，整個檔名被吃光（實測 `report.v1.final` 取代「.」→ 結果變成空字串）。批次跑遇到這種輸入，單筆就可能拖到 GAS 6 分鐘上限，拖垮整批。
3. **`tests/test-functions.js` 內建的測試案例本身就會失敗**：`testFileNameUtilities()` 對 `.hidden-file` 的預期是 `expectedName: ''／expectedExt: '.hidden-file'`，但 `getNameWithoutExtension`/`getFileExtension`（BatchRename.js:141-149）對 `lastIndexOf('.') === 0` 的情況判斷式是 `> 0`（不含 0），實際回傳是「整串當檔名、副檔名空字串」，跟測試期望相反。代表 `runTests()` 從沒有真的在 GAS 編輯器跑過一次到底（progress.md 說「測試案例已撰寫完成」但沒說「有跑過」）。

## 順便確認沒問題的項目
- 無 `eval`/`new Function`，`src/` 沒有 Node.js 或瀏覽器限定 API（`require`/`fetch`/`document`/`window` 皆無），GAS 相容性乾淨
- 中文檔名、emoji（astral surrogate pair）檔名跑過大小寫轉換規則不會被截斷/corrupt
- `applyBatchRename`（BatchRename.js）整組函式在 `Code.js`/`FileOperations.js` 都沒人呼叫，是死碼——且 `tests/test-functions.js` 只測這組死碼，**完全沒測到真正在跑的 `FileOperations.js applyRenameRule` 那條production路徑**，上面 bug 2 的 ReDoS/正則注入就是活在沒被測到的那條路徑上
- 半形/全形冒號打錯（如「前綴:X」打成半形冒號）、序號參數格式錯、尋找字串含多個「→」被靜默截斷——這些「靜默 no-op」問題 progress.md 已點名要加硬性驗證，此次用 fixture 逐一重現確認存在，未新增獨立條目

## 下一步（接手的人從這裡開始）
1. 安裝 clasp：`npm install -g @google/clasp`，登入後 `clasp push`（**注意：目錄下無 `.clasp.json`，需先 `clasp create`/`clasp clone` 建立**，progress.md 已記錄此缺口）
2. 修 bug 1（改存/傳 File ID，不要靠檔名重查）與 bug 2（regex 特殊字元先 escape，或提供獨立的「literal/regex」切換）優先權最高，之後才是文件同步等其他項目
3. 在 Google Sheets 綁定此 Script，手動執行選單項目確認功能
4. 若需新規則，在 `src/BatchRename.js` 照現有模式新增（但先評估併入 `FileOperations.js`，見 progress.md J 段）

## 地雷（別踩）
- GAS 有執行時間上限（6 分鐘），批次大量檔案需分批呼叫；上面 bug 2 的 ReDoS 輸入會讓單筆處理時間暴增，更容易撞到這個上限
- `templates/` 目錄存放試算表範本，勿刪除（使用者需匯入作為起點）
- `appsscript.json` 的 `oauthScopes` 需精確，過寬會觸發 Google Workspace 安全審查

## 主辦權
單線／待分派
