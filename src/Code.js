function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('檔案重新命名工具')
    .addItem('讀取資料夾檔案', 'loadFolderFiles')
    .addItem('套用批次規則', 'applyBatchRules')
    .addItem('開始重新命名', 'startRenaming')
    .addSeparator()
    .addItem('清除所有資料', 'clearAllData')
    .addToUi();
}

function loadFolderFiles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const commandSheet = sheet.getSheetByName('指令區');
  const fileListSheet = sheet.getSheetByName('檔名變更區');
  
  if (!commandSheet || !fileListSheet) {
    SpreadsheetApp.getUi().alert('錯誤', '找不到必要的工作表，請確認已建立「指令區」和「檔名變更區」工作表。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  try {
    const folderId = getFolderIdFromSheet(commandSheet);
    if (!folderId) {
      SpreadsheetApp.getUi().alert('錯誤', '請先在指令區設定資料夾ID或路徑。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const files = getFilesFromFolder(folderId);
    populateFileList(fileListSheet, files);
    
    SpreadsheetApp.getUi().alert('完成', `成功讀取 ${files.length} 個檔案！`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('錯誤', `讀取檔案時發生錯誤：${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function startRenaming() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const commandSheet = sheet.getSheetByName('指令區');
  const fileListSheet = sheet.getSheetByName('檔名變更區');
  
  if (!commandSheet || !fileListSheet) {
    ui.alert('錯誤', '找不到必要的工作表，請確認已建立「指令區」和「檔名變更區」工作表。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const renameConfig = getRenameConfigFromSheet(commandSheet);
    if (!renameConfig) {
      ui.alert('錯誤', '請先設定重新命名參數。', ui.ButtonSet.OK);
      return;
    }
    
    const result = ui.alert('確認操作', 
      `準備執行以下操作：\n模式：${renameConfig.mode}\n參數：${renameConfig.parameter}\n操作類型：${renameConfig.operationType}\n\n確定要繼續嗎？`, 
      ui.ButtonSet.YES_NO);
    
    if (result === ui.Button.YES) {
      const successCount = executeRenaming(fileListSheet, renameConfig);
      ui.alert('完成', `成功處理 ${successCount} 個檔案！`, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('錯誤', `重新命名時發生錯誤：${error.message}`, ui.ButtonSet.OK);
  }
}

function applyBatchRules() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const commandSheet = sheet.getSheetByName('指令區');
  const fileListSheet = sheet.getSheetByName('檔名變更區');
  
  if (!commandSheet || !fileListSheet) {
    ui.alert('錯誤', '找不到必要的工作表。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const renameConfig = getRenameConfigFromSheet(commandSheet);
    if (!renameConfig || !renameConfig.mode) {
      ui.alert('錯誤', '請先在指令區設定重新命名模式和參數。', ui.ButtonSet.OK);
      return;
    }
    
    const lastRow = fileListSheet.getLastRow();
    if (lastRow < 2) {
      ui.alert('錯誤', '沒有檔案資料，請先讀取資料夾檔案。', ui.ButtonSet.OK);
      return;
    }
    
    applyRulesToExistingFiles(fileListSheet, renameConfig);
    ui.alert('完成', '已套用批次重新命名規則！', ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('錯誤', `套用規則時發生錯誤：${error.message}`, ui.ButtonSet.OK);
  }
}

function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('確認清除', '確定要清除所有資料嗎？此操作無法復原。', ui.ButtonSet.YES_NO);
  
  if (result === ui.Button.YES) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const fileListSheet = sheet.getSheetByName('檔名變更區');
    
    if (fileListSheet && fileListSheet.getLastRow() > 1) {
      fileListSheet.getRange(2, 1, fileListSheet.getLastRow() - 1, fileListSheet.getLastColumn()).clearContent();
      ui.alert('完成', '已清除所有檔案資料！', ui.ButtonSet.OK);
    }
  }
}