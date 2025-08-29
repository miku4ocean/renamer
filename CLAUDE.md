# CLAUDE.md

這個檔案提供 Claude Code (claude.ai/code) 在操作此專案時的指導原則。

## 專案概述

**Renamer** 是一個基於 Google Apps Script 的雲端檔案批次重新命名工具，專為在 Google Drive 中進行高效檔案管理而設計。

### 核心功能
- 批次讀取 Google Drive 資料夾中的檔案
- 提供多種檔案重新命名方案（類似 Mac 的批次更名功能）
- 支援原位置更名或複製到新位置後更名
- 透過 Google Sheets 介面進行操作管理
- 完整的操作記錄與還原機制

## 專案架構

### 技術棧
- **平台**: Google Apps Script (基於 V8 執行環境)
- **介面**: Google Sheets (作為使用者操作介面)
- **儲存**: Google Drive (檔案存取與管理)
- **時區**: Asia/Taipei

### 檔案結構
```
renamer/
├── src/                        # 主要程式碼
│   ├── Code.js                # 主要入口點與選單功能
│   ├── FileOperations.js      # 檔案操作核心功能
│   └── BatchRename.js         # 批次重新命名邏輯
├── templates/                  # Google Sheets 模板
│   ├── command-sheet-template.md    # 指令區工作表模板
│   └── filelist-sheet-template.md  # 檔名變更區工作表模板
├── docs/                       # 文件與說明
│   ├── setup-guide.md         # 安裝與設定指南
│   ├── user-manual.md         # 使用者手冊
│   └── api-reference.md       # API 參考文件
├── tests/                      # 測試檔案
│   └── test-functions.js      # 單元測試
├── appsscript.json            # Apps Script 專案設定
└── README.md                  # 專案說明
```

## 工作流程設計

### 1. 初始設定階段
1. 在 Google Drive 建立目標資料夾（A 資料夾）
2. 在 A 資料夾中建立同名的 Google Sheets 檔案
3. 設定工作表結構（指令區 + 檔名變更區）
4. 部署 Google Apps Script 代碼

### 2. 檔案讀取階段
- 透過 `loadFolderFiles()` 掃描指定資料夾
- 自動排除 Google Sheets 檔案避免遞迴
- 收集檔案基本資訊：名稱、路徑、類型、大小、修改時間

### 3. 批次重新命名設定階段
支援以下重新命名模式：
- **新增文字**: 前綴、後綴、前後皆加
- **取代文字**: 完全取代、部分取代（支援正規表達式）
- **大小寫轉換**: 全大寫、全小寫、首字母大寫
- **新增序號**: 前綴序號、後綴序號、插入序號
- **格式化日期**: YYYY-MM-DD、YYYYMMDD、DD-MM-YYYY

### 4. 執行重新命名階段
- **原位置更名**: 直接修改原檔案名稱
- **複製後更名**: 複製到 B 資料夾並重新命名
- 提供預覽功能確認變更
- 支援批次確認與個別調整

### 5. 記錄與管理階段
- 完整記錄所有操作歷史
- 支援操作復原（限原位置更名）
- 提供操作統計與報告

## Google Sheets 工作表設計

### 指令區工作表
| 欄位 | 說明 | 範例 |
|------|------|------|
| A1 | 標題：檔案重新命名指令區 | |
| A2 | 來源資料夾 | 資料夾 ID 或 URL |
| A3 | 目標資料夾（複製模式用） | 資料夾 ID 或 URL |
| A4 | 重新命名模式 | 新增文字/取代文字/等 |
| A5 | 模式參數 | 依選擇模式而定 |
| A6 | 操作類型 | 原位置更名/複製後更名 |

### 檔名變更區工作表
| 欄位 | 說明 |
|------|------|
| A | 原檔名 |
| B | 原檔案路徑 |
| C | 變更後檔名 |
| D | 變更後檔案路徑 |
| E | 檔案類型 |
| F | 檔案大小 |
| G | 最後修改時間 |

## 核心功能模組

### Code.js - 主要入口點
- `onOpen()`: 建立自訂選單
- `loadFolderFiles()`: 讀取資料夾檔案
- `startRenaming()`: 開始重新命名程序
- `clearAllData()`: 清除所有資料

### FileOperations.js - 檔案操作
- `getFolderIdFromSheet()`: 從工作表取得資料夾 ID
- `getFilesFromFolder()`: 取得資料夾中的所有檔案
- `populateFileList()`: 填充檔案清單到工作表
- `renameFile()`: 重新命名檔案
- `copyAndRenameFile()`: 複製並重新命名檔案

### BatchRename.js - 批次重新命名邏輯
- `generateBatchRenameOptions()`: 產生重新命名選項
- `applyBatchRename()`: 套用批次重新命名規則
- 各種命名規則的實作函數

## 安全性考量

- 僅能存取已授權的 Google Drive 資料夾
- 支援操作預覽避免意外更改
- 提供操作記錄便於追蹤與復原
- 資料驗證防止無效操作

## 部署與使用

### 部署步驟
1. 開啟 Google Apps Script (script.google.com)
2. 建立新專案並上傳代碼檔案
3. 設定必要權限（Drive 存取）
4. 建立 Google Sheets 並匯入模板

### 使用流程
1. 在目標資料夾建立 Google Sheets
2. 透過自訂選單讀取檔案
3. 設定重新命名規則
4. 預覽並確認變更
5. 執行批次重新命名

## 錯誤處理

- 權限不足：提示使用者授權
- 資料夾不存在：驗證資料夾 ID
- 檔名衝突：提供解決方案選項
- 網路錯誤：重試機制

## 擴展性設計

- 模組化程式架構便於功能擴展
- 支援自訂命名規則
- 可整合其他 Google Workspace 服務
- 支援大量檔案處理最佳化

## 版本控制

- 使用 Git 進行版本控制
- 分支策略：main（主要版本）+ develop（開發版本）
- 標準化提交訊息格式
- 自動化測試與部署流程

## 測試策略

- 單元測試：個別函數功能驗證
- 整合測試：完整工作流程測試
- 使用者測試：實際使用場景驗證
- 效能測試：大量檔案處理測試