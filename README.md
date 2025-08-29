# Renamer

一個基於 Google Apps Script 的雲端檔案批次重新命名工具，透過 Google Sheets 介面提供直觀的檔案管理功能。

## ✨ 特色功能

- 🔍 **智慧檔案掃描** - 自動讀取 Google Drive 資料夾中的所有檔案
- 🏷️ **多樣化命名規則** - 支援文字新增、取代、大小寫轉換、序號和日期格式化
- 📋 **視覺化操作界面** - 透過 Google Sheets 進行直觀的檔案管理
- 🔄 **雙重操作模式** - 支援原位置更名或複製到新位置後更名  
- 📊 **完整操作記錄** - 詳細的檔案資訊追蹤與操作歷史
- ⚡ **批次高效處理** - 一次處理大量檔案，節省時間

## 🚀 快速開始

### 環境需求

- Google 帳號
- Google Drive 存取權限
- 基本的 Google Sheets 操作能力

### 安裝步驟

1. **建立 Google Apps Script 專案**
   ```
   前往 https://script.google.com/
   建立新專案，命名為 "Renamer"
   ```

2. **上傳程式碼**
   ```
   複製 /src/ 目錄下的所有 .js 檔案內容
   分別建立對應的 .gs 檔案並貼上程式碼
   ```

3. **設定權限**
   ```
   執行 onOpen() 函數進行授權
   允許 Google Drive 檔案存取權限
   ```

4. **建立工作表**
   ```
   建立新的 Google Sheets
   設定「指令區」和「檔名變更區」工作表
   ```

詳細安裝說明請參考：[📖 安裝指南](docs/setup-guide.md)

## 📋 使用方法

### 基本操作流程

1. **設定來源資料夾**
   - 在「指令區」工作表輸入 Google Drive 資料夾 ID 或 URL

2. **讀取檔案**
   - 使用選單「檔案重新命名工具」→「讀取資料夾檔案」

3. **設定重新命名規則**
   - 選擇重新命名模式（新增文字、取代文字等）
   - 設定對應參數

4. **執行重新命名**
   - 預覽變更結果
   - 執行批次處理

### 支援的重新命名模式

| 模式 | 說明 | 範例 |
|------|------|------|
| 🏷️ 新增文字 | 在檔名前後新增文字 | `document.pdf` → `IMG_document.pdf` |
| 🔄 取代文字 | 替換檔名中的特定文字 | `draft_v1.docx` → `final_v1.docx` |
| 📝 大小寫轉換 | 轉換檔名大小寫格式 | `MyFile.PDF` → `myfile.pdf` |
| 🔢 新增序號 | 為檔案新增連續序號 | `photo.jpg` → `001_photo.jpg` |
| 📅 格式化日期 | 根據修改時間新增日期 | `report.xlsx` → `2024-03-15_report.xlsx` |

## 📁 專案結構

```
renamer/
├── src/                          # 核心程式碼
│   ├── Code.js                   # 主要入口點與選單功能
│   ├── FileOperations.js         # 檔案操作核心功能
│   └── BatchRename.js           # 批次重新命名邏輯
├── templates/                    # Google Sheets 模板
│   ├── command-sheet-template.md      # 指令區工作表模板
│   └── filelist-sheet-template.md    # 檔名變更區工作表模板
├── docs/                         # 說明文件
│   ├── setup-guide.md           # 安裝設定指南
│   ├── user-manual.md          # 詳細使用手冊
│   └── api-reference.md        # API 參考文件
├── tests/                       # 測試檔案
├── appsscript.json             # Apps Script 專案設定
├── CLAUDE.md                   # Claude 專案指南
└── README.md                   # 專案說明
```

## 📖 文件導覽

- 🛠️ [安裝設定指南](docs/setup-guide.md) - 詳細的安裝步驟和環境設定
- 📘 [使用者手冊](docs/user-manual.md) - 完整的功能說明和操作教學
- 🔧 [API 參考文件](docs/api-reference.md) - 開發者技術參考
- 📋 [工作表模板](templates/) - Google Sheets 工作表設定模板

## 🎯 使用場景

### 📸 相片檔案整理
批次為旅遊照片新增日期前綴，便於時間順序管理。

### 📄 文件版本控制  
將草稿文件批次改為正式版本，支援複製到新位置保留原檔。

### 🔢 檔案序號化
為散亂檔案新增連續序號，建立有序的檔案結構。

### 📝 命名格式統一
批次轉換檔名大小寫，統一專案檔案命名格式。

## ⚠️ 注意事項

- 📋 執行前請備份重要檔案
- 🚫 避免檔名包含特殊字元：`/ \\ : * ? \" < > |`
- ⚡ 建議單次處理不超過 1000 個檔案
- 🔒 確保對目標資料夾有適當權限

## 🤝 貢獻指南

歡迎提交問題報告、功能建議或程式碼改進：

1. Fork 此專案
2. 建立功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交變更 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request

## 📝 授權條款

此專案採用 MIT 授權條款 - 詳見 [LICENSE](LICENSE) 檔案

## 🆘 技術支援

如果您遇到問題或需要協助：

- 📖 查看 [常見問題](docs/user-manual.md#故障排除)
- 🐛 提交 [Issue](https://github.com/your-username/renamer/issues)
- 📧 聯繫開發團隊

## 🎉 版本歷程

- **v1.0.0** (2024-03-29)
  - 🎯 初版發布
  - ✅ 核心檔案操作功能
  - 📋 Google Sheets 介面整合
  - 🏷️ 多種重新命名模式支援

---

<p align=\"center\">
  <strong>讓檔案管理變得更簡單 🚀</strong><br>
  使用 Renamer 輕鬆處理大量檔案重新命名需求
</p>