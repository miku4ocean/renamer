# Renamer API 參考文件

## 概述

本文件詳細說明 Renamer 工具的所有函數、參數和回傳值，適合開發者進行客製化或除錯使用。

## 核心函數 (Code.js)

### onOpen()

建立自訂選單，當 Google Sheets 開啟時自動執行。

```javascript
function onOpen()
```

**功能**: 
- 建立「檔案重新命名工具」選單
- 新增選單項目：讀取資料夾檔案、開始重新命名、清除所有資料

**回傳值**: 無

**範例**:
```javascript
// 此函數會在 Google Sheets 開啟時自動執行
// 無需手動呼叫
```

### loadFolderFiles()

從指定的 Google Drive 資料夾讀取檔案資訊。

```javascript
function loadFolderFiles()
```

**功能**:
- 從「指令區」工作表讀取資料夾設定
- 掃描資料夾中的所有檔案（排除 Google Sheets）
- 將檔案資訊填入「檔名變更區」工作表

**錯誤處理**:
- 工作表不存在：顯示錯誤訊息
- 資料夾 ID 無效：提示設定資料夾 ID
- 存取權限不足：顯示權限錯誤

**回傳值**: 無

**相依函數**: `getFolderIdFromSheet()`, `getFilesFromFolder()`, `populateFileList()`

### startRenaming()

執行批次檔案重新命名操作。

```javascript
function startRenaming()
```

**狀態**: 開發中

**功能**: 
- 讀取重新命名規則
- 執行批次重新命名
- 更新檔案狀態

### clearAllData()

清除「檔名變更區」工作表的所有檔案資料。

```javascript
function clearAllData()
```

**功能**:
- 顯示確認對話框
- 清除第2列開始的所有資料
- 保留標題列

**使用者互動**: 需要使用者確認才會執行

## 檔案操作函數 (FileOperations.js)

### getFolderIdFromSheet(sheet)

從工作表儲存格提取 Google Drive 資料夾 ID。

```javascript
function getFolderIdFromSheet(sheet)
```

**參數**:
- `sheet` {Sheet}: Google Sheets 工作表物件

**回傳值**:
- `{string|null}`: 資料夾 ID 或 null（如果未設定）

**支援格式**:
- 資料夾 ID：`1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms`
- 完整 URL：`https://drive.google.com/drive/folders/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms`

**範例**:
```javascript
const commandSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('指令區');
const folderId = getFolderIdFromSheet(commandSheet);
console.log(folderId); // "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
```

### getFilesFromFolder(folderId)

取得指定資料夾中的所有檔案資訊。

```javascript
function getFilesFromFolder(folderId)
```

**參數**:
- `folderId` {string}: Google Drive 資料夾 ID

**回傳值**:
- `{Array<FileInfo>}`: 檔案資訊陣列

**FileInfo 物件結構**:
```javascript
{
  id: string,           // 檔案 ID
  name: string,         // 檔案名稱
  path: string,         // 檔案路徑
  mimeType: string,     // MIME 類型
  size: number,         // 檔案大小（bytes）
  lastModified: Date    // 最後修改時間
}
```

**錯誤處理**:
- 拋出 Error 如果無法存取資料夾

**範例**:
```javascript
try {
  const files = getFilesFromFolder('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms');
  console.log(`找到 ${files.length} 個檔案`);
} catch (error) {
  console.error('無法讀取資料夾:', error.message);
}
```

### populateFileList(sheet, files)

將檔案資訊填入 Google Sheets 工作表。

```javascript
function populateFileList(sheet, files)
```

**參數**:
- `sheet` {Sheet}: 目標工作表
- `files` {Array<FileInfo>}`: 檔案資訊陣列

**功能**:
- 清除現有資料（保留標題列）
- 填入新的檔案資訊
- 自動調整資料範圍

**資料對應**:
| 欄位 | 來源 |
|------|------|
| A | file.name |
| B | file.path |
| C | file.name（預設） |
| D | file.path（預設） |
| E | file.mimeType |
| F | file.size |
| G | file.lastModified |

### renameFile(fileId, newName)

重新命名單一檔案。

```javascript
function renameFile(fileId, newName)
```

**參數**:
- `fileId` {string}: 檔案 ID
- `newName` {string}: 新檔案名稱

**回傳值**:
- `{boolean}`: 成功回傳 true

**錯誤處理**:
- 拋出 Error 如果重新命名失敗

**範例**:
```javascript
try {
  const success = renameFile('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms', 'new_name.pdf');
  console.log('重新命名成功:', success);
} catch (error) {
  console.error('重新命名失敗:', error.message);
}
```

### copyAndRenameFile(fileId, targetFolderId, newName)

複製檔案到目標資料夾並重新命名。

```javascript
function copyAndRenameFile(fileId, targetFolderId, newName)
```

**參數**:
- `fileId` {string}: 來源檔案 ID
- `targetFolderId` {string}: 目標資料夾 ID
- `newName` {string}: 新檔案名稱

**回傳值**:
- `{string}`: 新建立檔案的 ID

**錯誤處理**:
- 拋出 Error 如果複製失敗

## 批次重新命名函數 (BatchRename.js)

### generateBatchRenameOptions()

產生所有可用的重新命名選項。

```javascript
function generateBatchRenameOptions()
```

**回傳值**:
- `{Object}`: 重新命名選項物件

**選項結構**:
```javascript
{
  addText: {
    name: '新增文字',
    options: ['前綴', '後綴', '前後皆加']
  },
  replaceText: {
    name: '取代文字',
    options: ['完全取代', '部分取代']
  },
  changeCase: {
    name: '大小寫轉換',
    options: ['全部大寫', '全部小寫', '首字母大寫']
  },
  addNumbers: {
    name: '新增序號',
    options: ['前綴序號', '後綴序號', '插入序號']
  },
  formatDate: {
    name: '格式化日期',
    options: ['YYYY-MM-DD', 'YYYYMMDD', 'DD-MM-YYYY']
  }
}
```

### applyBatchRename(files, renameConfig)

套用批次重新命名規則到檔案陣列。

```javascript
function applyBatchRename(files, renameConfig)
```

**參數**:
- `files` {Array<FileInfo>}: 檔案資訊陣列
- `renameConfig` {RenameConfig}: 重新命名設定

**RenameConfig 物件**:
```javascript
{
  type: string,        // 'addText', 'replaceText', 'changeCase', 'addNumbers', 'formatDate'
  // 其他屬性依據 type 而定
}
```

**回傳值**:
- `{Array<RenamedFile>}`: 處理後的檔案陣列

**RenamedFile 物件**:
```javascript
{
  ...FileInfo,         // 原始檔案資訊
  newName: string,     // 新檔案名稱
  renamed: boolean     // 是否有變更
}
```

### 輔助函數

#### applyAddText(baseName, config, extension)

套用「新增文字」規則。

**參數**:
- `baseName` {string}: 檔名主體（不含副檔名）
- `config` {Object}: 設定物件
- `extension` {string}: 副檔名

**config 物件屬性**:
- `text` {string}: 要新增的文字
- `position` {string}: 位置（'前綴', '後綴', '前後皆加'）

#### applyReplaceText(baseName, config, extension)

套用「取代文字」規則。

**參數**:
- `baseName` {string}: 檔名主體
- `config` {Object}: 設定物件
- `extension` {string}: 副檔名

**config 物件屬性**:
- `findText` {string}: 要尋找的文字
- `replaceText` {string}: 取代的文字
- `option` {string}: 選項（'完全取代', '部分取代'）

#### applyChangeCase(baseName, config, extension)

套用「大小寫轉換」規則。

**參數**:
- `baseName` {string}: 檔名主體
- `config` {Object}: 設定物件
- `extension` {string}: 副檔名

**config 物件屬性**:
- `option` {string}: 轉換選項（'全部大寫', '全部小寫', '首字母大寫'）

#### applyAddNumbers(baseName, config, extension, index)

套用「新增序號」規則。

**參數**:
- `baseName` {string}: 檔名主體
- `config` {Object}: 設定物件
- `extension` {string}: 副檔名
- `index` {number}: 檔案索引

**config 物件屬性**:
- `startNumber` {number}: 起始數字（預設：1）
- `digits` {number}: 位數（預設：2）
- `position` {string}: 位置（'前綴序號', '後綴序號', '插入序號'）

#### applyFormatDate(baseName, config, extension, lastModified)

套用「格式化日期」規則。

**參數**:
- `baseName` {string}: 檔名主體
- `config` {Object}: 設定物件
- `extension` {string}: 副檔名
- `lastModified` {Date}: 檔案最後修改時間

**config 物件屬性**:
- `format` {string}: 日期格式（'YYYY-MM-DD', 'YYYYMMDD', 'DD-MM-YYYY'）

#### getNameWithoutExtension(filename)

取得檔名主體（不含副檔名）。

```javascript
function getNameWithoutExtension(filename)
```

**參數**:
- `filename` {string}: 完整檔名

**回傳值**:
- `{string}`: 檔名主體

**範例**:
```javascript
const baseName = getNameWithoutExtension('document.pdf');
console.log(baseName); // "document"
```

#### getFileExtension(filename)

取得檔案副檔名。

```javascript
function getFileExtension(filename)
```

**參數**:
- `filename` {string}: 完整檔名

**回傳值**:
- `{string}`: 副檔名（包含點號）

**範例**:
```javascript
const extension = getFileExtension('document.pdf');
console.log(extension); // ".pdf"
```

## 錯誤代碼

### 系統錯誤

- **SHEET_NOT_FOUND**: 找不到指定的工作表
- **FOLDER_NOT_FOUND**: 找不到指定的資料夾
- **PERMISSION_DENIED**: 權限不足
- **INVALID_FOLDER_ID**: 無效的資料夾 ID

### 檔案操作錯誤

- **RENAME_FAILED**: 檔案重新命名失敗
- **COPY_FAILED**: 檔案複製失敗
- **FILE_NOT_FOUND**: 找不到指定檔案
- **INVALID_FILENAME**: 無效的檔案名稱

### 設定錯誤

- **INVALID_CONFIG**: 無效的設定參數
- **MISSING_PARAMETER**: 缺少必要參數
- **INVALID_RENAME_MODE**: 無效的重新命名模式

## 除錯建議

### 記錄功能

使用 `console.log()` 進行除錯：

```javascript
console.log('檔案數量:', files.length);
console.log('設定參數:', renameConfig);
```

### 錯誤追蹤

檢查 Google Apps Script 的執行記錄：
1. 開啟 Apps Script 編輯器
2. 點擊「執行」→「檢視執行記錄」
3. 查看詳細的錯誤訊息

### 測試建議

1. **小批次測試**: 先用少量檔案測試
2. **分步執行**: 分別測試每個功能模組
3. **參數驗證**: 確認所有參數格式正確

## 版本資訊

- **API 版本**: 1.0.0
- **Google Apps Script 版本**: V8
- **相容性**: Google Workspace
- **時區**: Asia/Taipei

---

*API 文件版本: 1.0.0 | 最後更新: 2024年3月*