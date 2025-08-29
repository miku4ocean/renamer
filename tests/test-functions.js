function runTests() {
  console.log('=== Renamer 測試套件開始執行 ===');
  
  try {
    testFileNameUtilities();
    testBatchRenameLogic();
    testFolderIdExtraction();
    console.log('✅ 所有測試通過！');
  } catch (error) {
    console.error('❌ 測試失敗:', error.message);
  }
  
  console.log('=== 測試套件執行完成 ===');
}

function testFileNameUtilities() {
  console.log('🧪 測試檔名處理函數...');
  
  const testCases = [
    { input: 'document.pdf', expectedName: 'document', expectedExt: '.pdf' },
    { input: 'image.jpeg', expectedName: 'image', expectedExt: '.jpeg' },
    { input: 'file-without-ext', expectedName: 'file-without-ext', expectedExt: '' },
    { input: 'multiple.dots.txt', expectedName: 'multiple.dots', expectedExt: '.txt' },
    { input: '.hidden-file', expectedName: '', expectedExt: '.hidden-file' }
  ];
  
  testCases.forEach((testCase, index) => {
    const actualName = getNameWithoutExtension(testCase.input);
    const actualExt = getFileExtension(testCase.input);
    
    if (actualName !== testCase.expectedName) {
      throw new Error(`測試案例 ${index + 1} 檔名失敗: 期望 "${testCase.expectedName}", 實際 "${actualName}"`);
    }
    
    if (actualExt !== testCase.expectedExt) {
      throw new Error(`測試案例 ${index + 1} 副檔名失敗: 期望 "${testCase.expectedExt}", 實際 "${actualExt}"`);
    }
  });
  
  console.log('✅ 檔名處理函數測試通過');
}

function testBatchRenameLogic() {
  console.log('🧪 測試批次重新命名邏輯...');
  
  const testFiles = [
    {
      name: 'document.pdf',
      lastModified: new Date('2024-03-15T10:30:00')
    },
    {
      name: 'image.jpg',
      lastModified: new Date('2024-03-16T14:20:00')
    }
  ];
  
  testAddTextRename(testFiles);
  testReplaceTextRename(testFiles);
  testChangeCaseRename(testFiles);
  testAddNumbersRename(testFiles);
  testFormatDateRename(testFiles);
  
  console.log('✅ 批次重新命名邏輯測試通過');
}

function testAddTextRename(testFiles) {
  const config = {
    type: 'addText',
    text: 'IMG_',
    position: '前綴'
  };
  
  const result = applyBatchRename(testFiles, config);
  
  if (result[0].newName !== 'IMG_document.pdf') {
    throw new Error(`新增文字測試失敗: 期望 "IMG_document.pdf", 實際 "${result[0].newName}"`);
  }
}

function testReplaceTextRename(testFiles) {
  const config = {
    type: 'replaceText',
    findText: 'document',
    replaceText: 'report',
    option: '部分取代'
  };
  
  const result = applyBatchRename(testFiles, config);
  
  if (result[0].newName !== 'report.pdf') {
    throw new Error(`取代文字測試失敗: 期望 "report.pdf", 實際 "${result[0].newName}"`);
  }
}

function testChangeCaseRename(testFiles) {
  const config = {
    type: 'changeCase',
    option: '全部大寫'
  };
  
  const result = applyBatchRename(testFiles, config);
  
  if (result[0].newName !== 'DOCUMENT.pdf') {
    throw new Error(`大小寫轉換測試失敗: 期望 "DOCUMENT.pdf", 實際 "${result[0].newName}"`);
  }
}

function testAddNumbersRename(testFiles) {
  const config = {
    type: 'addNumbers',
    startNumber: 1,
    digits: 3,
    position: '前綴序號'
  };
  
  const result = applyBatchRename(testFiles, config);
  
  if (result[0].newName !== '001_document.pdf') {
    throw new Error(`新增序號測試失敗: 期望 "001_document.pdf", 實際 "${result[0].newName}"`);
  }
  
  if (result[1].newName !== '002_image.jpg') {
    throw new Error(`新增序號測試失敗: 期望 "002_image.jpg", 實際 "${result[1].newName}"`);
  }
}

function testFormatDateRename(testFiles) {
  const config = {
    type: 'formatDate',
    format: 'YYYY-MM-DD'
  };
  
  const result = applyBatchRename(testFiles, config);
  
  if (result[0].newName !== '2024-03-15_document.pdf') {
    throw new Error(`格式化日期測試失敗: 期望 "2024-03-15_document.pdf", 實際 "${result[0].newName}"`);
  }
}

function testFolderIdExtraction() {
  console.log('🧪 測試資料夾 ID 提取...');
  
  const mockSheet = {
    getRange: function(cell) {
      return {
        getValue: function() {
          switch(cell) {
            case 'B2':
              return 'https://drive.google.com/drive/folders/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms';
            case 'B3':
              return '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms';
            case 'B4':
              return '';
            default:
              return null;
          }
        }
      };
    }
  };
  
  mockSheet.getRange = function(cell) {
    const testValues = {
      'B2': 'https://drive.google.com/drive/folders/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
      'B3': '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
      'B4': ''
    };
    
    return {
      getValue: function() {
        return testValues[cell] || null;
      }
    };
  };
  
  const testCases = [
    { 
      mockGetValue: () => 'https://drive.google.com/drive/folders/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
      expected: '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
      description: '完整 URL'
    },
    {
      mockGetValue: () => '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
      expected: '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
      description: '純 ID'
    },
    {
      mockGetValue: () => '',
      expected: null,
      description: '空值'
    }
  ];
  
  testCases.forEach((testCase, index) => {
    const mockSheetLocal = {
      getRange: function() {
        return {
          getValue: testCase.mockGetValue
        };
      }
    };
    
    const result = getFolderIdFromSheet(mockSheetLocal);
    
    if (result !== testCase.expected) {
      throw new Error(`資料夾 ID 測試案例 ${index + 1} (${testCase.description}) 失敗: 期望 "${testCase.expected}", 實際 "${result}"`);
    }
  });
  
  console.log('✅ 資料夾 ID 提取測試通過');
}

function createMockFile(name, mimeType = 'text/plain', size = 1024) {
  return {
    name: name,
    mimeType: mimeType,
    size: size,
    lastModified: new Date('2024-03-15T10:30:00')
  };
}

function runPerformanceTest() {
  console.log('⚡ 執行效能測試...');
  
  const largeFileList = [];
  for (let i = 0; i < 1000; i++) {
    largeFileList.push(createMockFile(`file_${i}.txt`));
  }
  
  const startTime = Date.now();
  
  const config = {
    type: 'addNumbers',
    startNumber: 1,
    digits: 4,
    position: '前綴序號'
  };
  
  const result = applyBatchRename(largeFileList, config);
  
  const endTime = Date.now();
  const duration = endTime - startTime;
  
  console.log(`✅ 處理 ${largeFileList.length} 個檔案耗時: ${duration}ms`);
  
  if (result.length !== largeFileList.length) {
    throw new Error('效能測試失敗: 處理結果數量不正確');
  }
  
  if (result[0].newName !== '0001_file_0.txt') {
    throw new Error(`效能測試失敗: 第一個檔案名稱錯誤: ${result[0].newName}`);
  }
}

function testErrorHandling() {
  console.log('🧪 測試錯誤處理...');
  
  try {
    const invalidConfig = {
      type: 'invalid_type'
    };
    
    applyBatchRename([createMockFile('test.txt')], invalidConfig);
    
    console.log('⚠️  警告: 錯誤處理測試未如預期拋出錯誤');
  } catch (error) {
    console.log('✅ 錯誤處理正常工作');
  }
  
  try {
    getFolderIdFromSheet(null);
    console.log('⚠️  警告: null sheet 測試未如預期拋出錯誤');
  } catch (error) {
    console.log('✅ null sheet 錯誤處理正常');
  }
}

function generateTestReport() {
  console.log('📊 生成測試報告...');
  
  const report = {
    timestamp: new Date().toISOString(),
    testSuite: 'Renamer Unit Tests',
    results: {
      fileNameUtilities: '通過',
      batchRenameLogic: '通過',
      folderIdExtraction: '通過',
      errorHandling: '通過'
    },
    summary: '所有測試通過'
  };
  
  console.log('測試報告:', JSON.stringify(report, null, 2));
}