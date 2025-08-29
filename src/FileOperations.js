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
  
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  
  const data = files.map(file => [
    file.name,
    file.path,
    file.name,
    file.path,
    file.mimeType,
    file.size,
    file.lastModified
  ]);
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
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