function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('檔案重新命名工具')
    .addItem('讀取資料夾檔案', 'loadFolderFiles')
    .addItem('開始重新命名', 'startRenaming')
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
  SpreadsheetApp.getUi().alert('提示', '重新命名功能開發中...', SpreadsheetApp.getUi().ButtonSet.OK);
}

function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('確認清除', '確定要清除所有資料嗎？此操作無法復原。', ui.ButtonSet.YES_NO);
  
  if (result === ui.Button.YES) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const fileListSheet = sheet.getSheetByName('檔名變更區');
    
    if (fileListSheet) {
      fileListSheet.getRange(2, 1, fileListSheet.getLastRow() - 1, fileListSheet.getLastColumn()).clearContent();
      ui.alert('完成', '已清除所有檔案資料！', ui.ButtonSet.OK);
    }
  }
}