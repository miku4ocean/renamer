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
  
  const data = files.map((file, index) => {
    let newName = file.name;
    let newPath = file.path;

    if (renameConfig && renameConfig.mode && renameConfig.parameter) {
      try {
        newName = applyRenameRule(file.name, renameConfig.mode, renameConfig.parameter, file.lastModified, index);
        
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
      file.lastModified,
      file.id,
      '' // I 欄／執行結果：這是新讀進來的一輪，之前留下的 ✓／✗ 紀錄一律作廢清空
    ];
  });
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
}

function getNameWithoutExtension(filename) {
  const lastDotIndex = filename.lastIndexOf('.');
  return lastDotIndex > 0 ? filename.substring(0, lastDotIndex) : filename;
}

function getFileExtension(filename) {
  const lastDotIndex = filename.lastIndexOf('.');
  return lastDotIndex > 0 ? filename.substring(lastDotIndex) : '';
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
    // 只切第一個箭頭：「部分取代：-→x→y」的取代文字要保留成 "x→y"，後面的箭頭不是分隔符
    const arrowIndex = parameter.indexOf('→');
    const findPart = parameter.substring(0, arrowIndex);
    const replaceText = parameter.substring(arrowIndex + 1);
    const literalFindText = findPart.replace('部分取代：', '');

    if (literalFindText === '') {
      // 尋找字串為空：沒有東西可比對，原樣傳回，不能讓 split('')/join() 在每個字元間插入
      return baseName + extension;
    }

    // 用 split/join 做「字面字串」取代，不建 RegExp：
    // 1) 取代文字裡的 $&、$1 等不會被當成正則特殊語法，只當一般字面字串
    // 2) 尋找字串裡的正則特殊字元（. * + ? 等）也不用跳脫，天生就是字面比對
    // 3) 沒有正則就沒有 ReDoS 風險
    return baseName.split(literalFindText).join(replaceText) + extension;
  }
  return baseName + extension;
}

// 保留給既有測試（testEscapeRegExp）使用；applyReplaceTextRule 已改用 split/join 做字面取代，
// 不再需要靠這支函式跳脫特殊字元建 RegExp。
function escapeRegExp(string) {
  return String(string).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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

// GAS 單次執行有 6 分鐘上限，大量檔案可能來不及跑完。executeRenaming 因此設計成可以安全重跑：
// - 每一列處理完立即把結果寫進 I 欄（執行結果），不等到最後才一次寫回，這樣就算執行途中被
//   GAS 強制中止，已經完成的列也有紀錄留在試算表上
// - 開跑時，I 欄已經是「✓ 開頭」的列會直接跳過，不會重新執行一次（複製後更名不會再多複製一份；
//   原位置更名也不會對已經改完名的檔案再做一次無意義的 setName）
// - options.now／options.budgetMs 讓測試可以注入假時鐘，不需要真的等 5 分鐘
function executeRenaming(fileListSheet, renameConfig, options) {
  options = options || {};
  const now = options.now || function() { return Date.now(); };
  const budgetMs = options.budgetMs || (5 * 60 * 1000);
  const deadline = now() + budgetMs;

  const lastRow = fileListSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('沒有檔案可以處理');
  }

  const totalRows = lastRow - 1;
  const data = fileListSheet.getRange(2, 1, totalRows, 9).getValues();
  let successCount = 0;
  let errorCount = 0;
  let processedCount = 0;
  let skippedDone = 0;
  let noChangeCount = 0;
  let timedOut = false;
  const errors = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const [originalName, , newName, , , , , fileId, execResult] = row;
    const resultCellRow = i + 2; // 對應試算表實際列號（第 1 列是標題）

    // 冪等：這一列先前已經成功完成過，直接跳過，不重複改名／複製
    if (typeof execResult === 'string' && execResult.indexOf('✓') === 0) {
      skippedDone++;
      continue;
    }

    if (now() > deadline) {
      timedOut = true;
      break;
    }

    if (!originalName || !newName || originalName === newName) {
      noChangeCount++;
      continue;
    }

    processedCount++;

    try {
      if (!fileId) {
        throw new Error('缺少檔案 ID，請重新執行「讀取資料夾檔案」');
      }

      if (renameConfig.operationType === '原位置更名') {
        renameFile(fileId, newName);
        const timestamp = Utilities.formatDate(new Date(now()), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        fileListSheet.getRange(resultCellRow, 9).setValue(`✓ 已更名 ${timestamp}`);
      } else if (renameConfig.operationType === '複製後更名') {
        if (!renameConfig.targetFolderId) {
          throw new Error('複製後更名需要設定目標資料夾');
        }
        const newFileId = copyAndRenameFile(fileId, renameConfig.targetFolderId, newName);
        fileListSheet.getRange(resultCellRow, 9).setValue(`✓ 已複製 ${newFileId}`);
      }

      successCount++;

    } catch (error) {
      errorCount++;
      errors.push(`${originalName}: ${error.message}`);
      fileListSheet.getRange(resultCellRow, 9).setValue(`✗ ${error.message}`);

      if (errors.length < 5) {
        console.log(`處理檔案 ${originalName} 時發生錯誤: ${error.message}`);
      }
    }
  }

  if (timedOut) {
    const doneSoFar = skippedDone + noChangeCount + processedCount;
    const remaining = totalRows - doneSoFar;
    return {
      timedOut: true,
      successCount: successCount,
      errorCount: errorCount,
      processedCount: processedCount,
      errors: errors,
      message: `已處理 ${processedCount} 列，剩餘 ${remaining} 列，請再執行一次「開始重新命名」`
    };
  }

  return {
    timedOut: false,
    successCount: successCount,
    errorCount: errorCount,
    processedCount: processedCount,
    errors: errors,
    message: errorCount > 0
      ? `成功: ${successCount}, 失敗: ${errorCount}\n前幾個錯誤:\n${errors.slice(0, 3).join('\n')}`
      : `成功處理 ${successCount} 個檔案！`
  };
}

function applyRulesToExistingFiles(fileListSheet, renameConfig) {
  const lastRow = fileListSheet.getLastRow();
  if (lastRow < 2) return;

  const data = fileListSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const newData = [];
  let seq = 0; // 獨立的序號計數器，只在真正處理到的（非空白）列遞增，跟迴圈索引 i 脫鉤

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const [originalName, originalPath, , , mimeType, size, lastModified, fileId] = row;

    if (!originalName) continue;

    const currentSeq = seq;
    seq++;

    try {
      let newName = applyRenameRule(originalName, renameConfig.mode, renameConfig.parameter, lastModified, currentSeq);
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

      // I 欄（執行結果）在這裡一律清空：重新套用規則代表「變更後檔名」是新的一輪，
      // 之前留下的 ✓／✗ 紀錄已經不對應現在這個檔名了
      newData.push([originalName, originalPath, newName, newPath, mimeType, size, lastModified, fileId, '']);
    } catch (error) {
      newData.push([originalName, originalPath, row[2], row[3], mimeType, size, lastModified, fileId, '']);
      console.log(`處理檔案 ${originalName} 時發生錯誤: ${error.message}`);
    }
  }

  if (newData.length > 0) {
    fileListSheet.getRange(2, 1, newData.length, 9).setValues(newData);
  }
}
