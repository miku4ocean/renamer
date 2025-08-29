function generateBatchRenameOptions() {
  return {
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
  };
}

function applyBatchRename(files, renameConfig) {
  const renamedFiles = [];
  
  files.forEach((file, index) => {
    let newName = file.name;
    const nameWithoutExt = getNameWithoutExtension(newName);
    const extension = getFileExtension(newName);
    
    switch (renameConfig.type) {
      case 'addText':
        newName = applyAddText(nameWithoutExt, renameConfig, extension);
        break;
      case 'replaceText':
        newName = applyReplaceText(nameWithoutExt, renameConfig, extension);
        break;
      case 'changeCase':
        newName = applyChangeCase(nameWithoutExt, renameConfig, extension);
        break;
      case 'addNumbers':
        newName = applyAddNumbers(nameWithoutExt, renameConfig, extension, index);
        break;
      case 'formatDate':
        newName = applyFormatDate(nameWithoutExt, renameConfig, extension, file.lastModified);
        break;
      default:
        newName = file.name;
    }
    
    renamedFiles.push({
      ...file,
      newName: newName,
      renamed: newName !== file.name
    });
  });
  
  return renamedFiles;
}

function applyAddText(baseName, config, extension) {
  const text = config.text || '';
  
  switch (config.position) {
    case '前綴':
      return text + baseName + extension;
    case '後綴':
      return baseName + text + extension;
    case '前後皆加':
      return text + baseName + text + extension;
    default:
      return baseName + extension;
  }
}

function applyReplaceText(baseName, config, extension) {
  const findText = config.findText || '';
  const replaceText = config.replaceText || '';
  
  if (config.option === '完全取代') {
    return replaceText + extension;
  } else {
    return baseName.replace(new RegExp(findText, 'g'), replaceText) + extension;
  }
}

function applyChangeCase(baseName, config, extension) {
  switch (config.option) {
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

function applyAddNumbers(baseName, config, extension, index) {
  const startNumber = config.startNumber || 1;
  const digits = config.digits || 2;
  const number = String(startNumber + index).padStart(digits, '0');
  
  switch (config.position) {
    case '前綴序號':
      return number + '_' + baseName + extension;
    case '後綴序號':
      return baseName + '_' + number + extension;
    case '插入序號':
      return baseName + '(' + number + ')' + extension;
    default:
      return baseName + extension;
  }
}

function applyFormatDate(baseName, config, extension, lastModified) {
  const date = new Date(lastModified);
  let dateStr = '';
  
  switch (config.format) {
    case 'YYYY-MM-DD':
      dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      break;
    case 'YYYYMMDD':
      dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
      break;
    case 'DD-MM-YYYY':
      dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MM-yyyy');
      break;
    default:
      dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  
  return dateStr + '_' + baseName + extension;
}

function getNameWithoutExtension(filename) {
  const lastDotIndex = filename.lastIndexOf('.');
  return lastDotIndex > 0 ? filename.substring(0, lastDotIndex) : filename;
}

function getFileExtension(filename) {
  const lastDotIndex = filename.lastIndexOf('.');
  return lastDotIndex > 0 ? filename.substring(lastDotIndex) : '';
}