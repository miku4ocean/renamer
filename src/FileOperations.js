function getFolderIdFromSheet(sheet) {
  const folderCell = sheet.getRange('B2');
  const folderValue = folderCell.getValue();
  
  if (!folderValue) {
    return null;
  }
  
  if (typeof folderValue === 'string') {
    const urlMatch = folderValue.match(/[-\w]{25,}/);
    return urlMatch ? urlMatch[0] : folderValue;
  }
  
  return folderValue;
}

function getFilesFromFolder(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = [];
    const fileIterator = folder.getFiles();
    
    while (fileIterator.hasNext()) {
      const file = fileIterator.next();
      if (file.getMimeType() !== 'application/vnd.google-apps.spreadsheet') {
        files.push({
          id: file.getId(),
          name: file.getName(),
          path: folder.getName() + '/' + file.getName(),
          mimeType: file.getMimeType(),
          size: file.getSize(),
          lastModified: file.getLastUpdated()
        });
      }
    }
    
    return files;
  } catch (error) {
    throw new Error(`無法存取資料夾：${error.message}`);
  }
}

function populateFileList(sheet, files) {
  if (files.length === 0) {
    return;
  }
  
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  
  const commandSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('指令區');
  const renameConfig = getRenameConfigFromSheet(commandSheet);
  
  const data = files.map(file => {
    let newName = file.name;
    let newPath = file.path;
    
    if (renameConfig && renameConfig.mode && renameConfig.parameter) {
      try {
        newName = applyRenameRule(file.name, renameConfig.mode, renameConfig.parameter, file.lastModified);
        
        if (renameConfig.operationType === '複製後更名' && renameConfig.targetFolderId) {
          try {
            const targetFolder = DriveApp.getFolderById(renameConfig.targetFolderId);
            newPath = targetFolder.getName() + '/' + newName;
          } catch (error) {
            newPath = '目標資料夾/' + newName;
          }
        } else {
          const pathParts = file.path.split('/');
          pathParts[pathParts.length - 1] = newName;
          newPath = pathParts.join('/');
        }
      } catch (error) {
        console.log(`套用重新命名規則失敗 ${file.name}: ${error.message}`);
      }
    }
    
    return [
      file.name,
      file.path, 
      newName,
      newPath,
      file.mimeType,
      file.size,
      file.lastModified
    ];
  });
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
}

function applyRenameRule(fileName, mode, parameter, lastModified, index = 0) {
  const nameWithoutExt = getNameWithoutExtension(fileName);
  const extension = getFileExtension(fileName);
  
  switch (mode) {
    case '新增文字':
      return applyAddTextRule(nameWithoutExt, parameter, extension);
    
    case '取代文字':
      return applyReplaceTextRule(nameWithoutExt, parameter, extension);
    
    case '大小寫轉換':
      return applyCaseChangeRule(nameWithoutExt, parameter, extension);
    
    case '新增序號':
      return applyNumberRule(nameWithoutExt, parameter, extension, index);
    
    case '格式化日期':
      return applyDateRule(nameWithoutExt, parameter, extension, lastModified);
    
    default:
      return fileName;
  }
}

function applyAddTextRule(baseName, parameter, extension) {
  if (parameter.startsWith('前綴：')) {
    const prefix = parameter.replace('前綴：', '');
    return prefix + baseName + extension;
  } else if (parameter.startsWith('後綴：')) {
    const suffix = parameter.replace('後綴：', '');
    return baseName + suffix + extension;
  } else if (parameter.startsWith('前後皆加：')) {
    const text = parameter.replace('前後皆加：', '');
    return text + baseName + text + extension;
  }
  return baseName + extension;
}

function applyReplaceTextRule(baseName, parameter, extension) {
  if (parameter.startsWith('完全取代：')) {
    const newName = parameter.replace('完全取代：', '');
    return newName + extension;
  } else if (parameter.includes('→')) {
    const [findText, replaceText] = parameter.split('→');
    return baseName.replace(new RegExp(findText.replace('部分取代：', ''), 'g'), replaceText) + extension;
  }
  return baseName + extension;
}

function applyCaseChangeRule(baseName, parameter, extension) {
  switch (parameter) {
    case '全部大寫':
      return baseName.toUpperCase() + extension;
    case '全部小寫':
      return baseName.toLowerCase() + extension;
    case '首字母大寫':
      return baseName.charAt(0).toUpperCase() + baseName.slice(1).toLowerCase() + extension;
    default:
      return baseName + extension;
  }
}

function applyNumberRule(baseName, parameter, extension, index) {
  const match = parameter.match(/(\d+),(\d+)/);
  if (match) {
    const startNumber = parseInt(match[1]);
    const digits = parseInt(match[2]);
    const number = String(startNumber + index).padStart(digits, '0');
    
    if (parameter.startsWith('前綴序號：')) {
      return number + '_' + baseName + extension;
    } else if (parameter.startsWith('後綴序號：')) {
      return baseName + '_' + number + extension;
    } else if (parameter.startsWith('插入序號：')) {
      return baseName + '(' + number + ')' + extension;
    }
  }
  return baseName + extension;
}

function applyDateRule(baseName, parameter, extension, lastModified) {
  const date = new Date(lastModified);
  let dateStr = '';
  
  if (parameter.includes('YYYY-MM-DD')) {
    dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } else if (parameter.includes('YYYYMMDD')) {
    dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
  } else if (parameter.includes('DD-MM-YYYY')) {
    dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  } else {
    dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  
  return dateStr + '_' + baseName + extension;
}

function renameFile(fileId, newName) {
  try {
    const file = DriveApp.getFileById(fileId);
    file.setName(newName);
    return true;
  } catch (error) {
    throw new Error(`重新命名檔案失敗：${error.message}`);
  }
}

function copyAndRenameFile(fileId, targetFolderId, newName) {
  try {
    const file = DriveApp.getFileById(fileId);
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    const copiedFile = file.makeCopy(newName, targetFolder);
    return copiedFile.getId();
  } catch (error) {
    throw new Error(`複製並重新命名檔案失敗：${error.message}`);
  }
}

function getRenameConfigFromSheet(sheet) {
  const mode = sheet.getRange('B4').getValue();
  const parameter = sheet.getRange('B5').getValue();
  const operationType = sheet.getRange('B6').getValue();
  const targetFolder = sheet.getRange('B3').getValue();
  
  if (!mode || !operationType) {
    return null;
  }
  
  return {
    mode: mode,
    parameter: parameter || '',
    operationType: operationType,
    targetFolderId: targetFolder ? getFolderIdFromValue(targetFolder) : null
  };
}

function getFolderIdFromValue(value) {
  if (!value) return null;
  
  if (typeof value === 'string') {
    const urlMatch = value.match(/[-\w]{25,}/);
    return urlMatch ? urlMatch[0] : value;
  }
  
  return value;
}

function executeRenaming(fileListSheet, renameConfig) {
  const lastRow = fileListSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('沒有檔案可以處理');
  }
  
  const data = fileListSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  let successCount = 0;
  let errorCount = 0;
  const errors = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const [originalName, originalPath, newName] = row;
    
    if (!originalName || !newName || originalName === newName) {
      continue;
    }
    
    try {
      const fileId = getFileIdByNameAndPath(originalName, originalPath);
      
      if (renameConfig.operationType === '原位置更名') {
        renameFile(fileId, newName);
      } else if (renameConfig.operationType === '複製後更名') {
        if (!renameConfig.targetFolderId) {
          throw new Error('複製後更名需要設定目標資料夾');
        }
        copyAndRenameFile(fileId, renameConfig.targetFolderId, newName);
      }
      
      successCount++;
      
    } catch (error) {
      errorCount++;
      errors.push(`${originalName}: ${error.message}`);
      
      if (errors.length < 5) {
        console.log(`處理檔案 ${originalName} 時發生錯誤: ${error.message}`);
      }
    }
  }
  
  if (errorCount > 0) {
    const errorMessage = `成功: ${successCount}, 失敗: ${errorCount}\n前幾個錯誤:\n${errors.slice(0, 3).join('\n')}`;
    throw new Error(errorMessage);
  }
  
  return successCount;
}

function applyRulesToExistingFiles(fileListSheet, renameConfig) {
  const lastRow = fileListSheet.getLastRow();
  if (lastRow < 2) return;
  
  const data = fileListSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const newData = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const [originalName, originalPath, , , mimeType, size, lastModified] = row;
    
    if (!originalName) continue;
    
    try {
      let newName = applyRenameRule(originalName, renameConfig.mode, renameConfig.parameter, lastModified, i);
      let newPath = originalPath;
      
      if (renameConfig.operationType === '複製後更名' && renameConfig.targetFolderId) {
        try {
          const targetFolder = DriveApp.getFolderById(renameConfig.targetFolderId);
          newPath = targetFolder.getName() + '/' + newName;
        } catch (error) {
          newPath = '目標資料夾/' + newName;
        }
      } else {
        const pathParts = originalPath.split('/');
        pathParts[pathParts.length - 1] = newName;
        newPath = pathParts.join('/');
      }
      
      newData.push([originalName, originalPath, newName, newPath, mimeType, size, lastModified]);
    } catch (error) {
      newData.push(row);
      console.log(`處理檔案 ${originalName} 時發生錯誤: ${error.message}`);
    }
  }
  
  if (newData.length > 0) {
    fileListSheet.getRange(2, 1, newData.length, 7).setValues(newData);
  }
}

function getFileIdByNameAndPath(fileName, filePath) {
  try {
    const commandSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('指令區');
    const sourceFolderId = getFolderIdFromSheet(commandSheet);
    
    if (sourceFolderId) {
      const sourceFolder = DriveApp.getFolderById(sourceFolderId);
      const files = sourceFolder.getFilesByName(fileName);
      
      while (files.hasNext()) {
        const file = files.next();
        if (file.getName() === fileName) {
          return file.getId();
        }
      }
    }
    
    const pathParts = filePath.split('/');
    const folderName = pathParts[0];
    const folders = DriveApp.getFoldersByName(folderName);
    
    while (folders.hasNext()) {
      const folder = folders.next();
      const files = folder.getFilesByName(fileName);
      
      while (files.hasNext()) {
        const file = files.next();
        if (file.getName() === fileName) {
          return file.getId();
        }
      }
    }
    
    throw new Error(`找不到檔案: ${fileName}`);
    
  } catch (error) {
    throw new Error(`無法找到檔案 ${fileName}: ${error.message}`);
  }
}