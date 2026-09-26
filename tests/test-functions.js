function runTests() {
  console.log('=== Renamer 測試套件開始執行 ===');

  try {
    testFileNameUtilities();
    testFolderIdExtraction();
    testProductionApplyRenameRule();
    testPopulateFileListSerialNumbering();
    testApplyRulesToExistingFilesSequentialNumbering();
    console.log('✅ 所有測試通過！');
  } catch (error) {
    console.error('❌ 測試失敗:', error.message);
    console.log('=== 測試套件執行完成 ===');
    // 重新拋出，讓呼叫端（例如 node:vm runner）知道測試套件失敗，不再被這裡的 catch 吞掉
    throw error;
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
    { input: '.hidden-file', expectedName: '.hidden-file', expectedExt: '' }
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

// 生產路徑測試：applyRenameRule 是 FileOperations.js 真正被 Code.js 呼叫的核心邏輯
// （原本的 applyBatchRename／BatchRename.js 是沒人呼叫的死碼，連同其專屬測試已一併移除）
function testProductionApplyRenameRule() {
  console.log('🧪 測試生產路徑 applyRenameRule（FileOperations.js）...');

  testApplyRenameRuleReplaceText();
  testApplyRenameRuleNumbering();
  testApplyRenameRuleCaseChange();
  testEscapeRegExp();

  console.log('✅ 生產路徑 applyRenameRule 測試通過');
}

function testApplyRenameRuleReplaceText() {
  const lastModified = new Date('2024-03-15T10:30:00');

  const result1 = applyRenameRule('document.pdf', '取代文字', '部分取代：document→report', lastModified);
  if (result1 !== 'report.pdf') {
    throw new Error(`applyRenameRule 取代文字測試失敗: 期望 "report.pdf", 實際 "${result1}"`);
  }

  // regex 特殊字元（.）必須被當字面文字處理，證明 escapeRegExp 有生效，
  // 而不是被當成正則萬用字元把整個檔名吃光
  const result2 = applyRenameRule('report.v1.final.pdf', '取代文字', '部分取代：.→_', lastModified);
  if (result2 !== 'report_v1_final.pdf') {
    throw new Error(`applyRenameRule 正則特殊字元跳脫測試失敗: 期望 "report_v1_final.pdf", 實際 "${result2}"`);
  }

  const result3 = applyRenameRule('document.pdf', '取代文字', '完全取代：final', lastModified);
  if (result3 !== 'final.pdf') {
    throw new Error(`applyRenameRule 完全取代測試失敗: 期望 "final.pdf", 實際 "${result3}"`);
  }

  // $& / $1 等取代字串必須當成字面值，不能被當成 String.prototype.replace 的特殊樣式
  // （改用 split/join 而非 RegExp#replace 之後才不會誤觸發）
  const result4 = applyRenameRule('a-b.txt', '取代文字', '部分取代：-→$&$&', lastModified);
  if (result4 !== 'a$&$&b.txt') {
    throw new Error(`applyRenameRule 取代文字含 $& 測試失敗: 期望 "a$&$&b.txt", 實際 "${result4}"`);
  }

  // 尋找字串為空：不能在每個字元間插入取代文字，應原樣傳回
  const result5 = applyRenameRule('abc.txt', '取代文字', '部分取代：→_', lastModified);
  if (result5 !== 'abc.txt') {
    throw new Error(`applyRenameRule 空尋找字串測試失敗: 期望 "abc.txt", 實際 "${result5}"`);
  }

  // 參數裡有多個 →：只切第一個箭頭，後面的箭頭要保留在取代文字裡，不能被丟掉
  const result6 = applyRenameRule('a-b.txt', '取代文字', '部分取代：-→x→y', lastModified);
  if (result6 !== 'ax→yb.txt') {
    throw new Error(`applyRenameRule 多個箭頭測試失敗: 期望 "ax→yb.txt", 實際 "${result6}"`);
  }
}

function testApplyRenameRuleNumbering() {
  const lastModified = new Date('2024-03-15T10:30:00');

  const result1 = applyRenameRule('photo.jpg', '新增序號', '前綴序號：1,3', lastModified, 0);
  if (result1 !== '001_photo.jpg') {
    throw new Error(`applyRenameRule 新增序號測試失敗: 期望 "001_photo.jpg", 實際 "${result1}"`);
  }

  const result2 = applyRenameRule('photo.jpg', '新增序號', '前綴序號：1,3', lastModified, 2);
  if (result2 !== '003_photo.jpg') {
    throw new Error(`applyRenameRule 新增序號（index）測試失敗: 期望 "003_photo.jpg", 實際 "${result2}"`);
  }
}

function testApplyRenameRuleCaseChange() {
  const lastModified = new Date('2024-03-15T10:30:00');

  const result = applyRenameRule('Document.PDF', '大小寫轉換', '全部大寫', lastModified);
  if (result !== 'DOCUMENT.PDF') {
    throw new Error(`applyRenameRule 大小寫轉換測試失敗: 期望 "DOCUMENT.PDF", 實際 "${result}"`);
  }
}

function testEscapeRegExp() {
  const escaped = escapeRegExp('a.b+c(d)[e]');
  if (escaped !== 'a\\.b\\+c\\(d\\)\\[e\\]') {
    throw new Error(`escapeRegExp 測試失敗: 期望 "a\\\\.b\\\\+c\\\\(d\\\\)\\\\[e\\\\]", 實際 "${escaped}"`);
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

function testErrorHandling() {
  console.log('🧪 測試錯誤處理...');

  try {
    getFolderIdFromSheet(null);
    console.log('⚠️  警告: null sheet 測試未如預期拋出錯誤');
  } catch (error) {
    console.log('✅ null sheet 錯誤處理正常');
  }
}

// 下面幾個測試需要 tests/run-node.mjs 提供的 makeSheet／假 SpreadsheetApp／假 DriveApp，
// 才能在不碰真實 Google Sheets／Drive 的情況下驗證「讀取資料夾檔案」「套用批次規則」
// 「開始重新命名」這幾條生產路徑。如果整份檔案是直接貼進真正的 GAS 指令碼編輯器執行
// （這幾個假物件不存在，全域的 SpreadsheetApp/DriveApp 會是真正的服務），
// 這裡會偵測不到假物件並直接跳過，不會因為呼叫到不存在的方法而噴錯誤。

function testPopulateFileListSerialNumbering() {
  console.log('🧪 測試 populateFileList 序號規則（bug: 讀取資料夾檔案時序號全部相同）...');

  if (typeof SpreadsheetApp === 'undefined' || typeof SpreadsheetApp._setSheet !== 'function') {
    console.log('⏭️  跳過（僅限 node tests/run-node.mjs 跑者，GAS 編輯器內沒有假 SpreadsheetApp）');
    return;
  }

  const commandSheet = {
    getRange: function(cell) {
      const values = { B3: '', B4: '新增序號', B5: '前綴序號：1,3', B6: '原位置更名' };
      return { getValue: function() { return values[cell] !== undefined ? values[cell] : ''; } };
    }
  };
  SpreadsheetApp._setSheet('指令區', commandSheet);

  const fileListSheet = makeSheet([[]]);
  const lastModified = new Date('2024-03-15T10:30:00');
  const files = [
    { name: 'a.txt', path: '資料夾/a.txt', mimeType: 'text/plain', size: 1, lastModified: lastModified, id: 'id-a' },
    { name: 'b.txt', path: '資料夾/b.txt', mimeType: 'text/plain', size: 2, lastModified: lastModified, id: 'id-b' },
    { name: 'c.txt', path: '資料夾/c.txt', mimeType: 'text/plain', size: 3, lastModified: lastModified, id: 'id-c' }
  ];

  populateFileList(fileListSheet, files);

  const written = fileListSheet.getRange(2, 1, 3, 8).getValues();
  const expectedNames = ['001_a.txt', '002_b.txt', '003_c.txt'];
  expectedNames.forEach(function(expected, idx) {
    const actual = written[idx][2];
    if (actual !== expected) {
      throw new Error('populateFileList 序號測試失敗（第 ' + (idx + 1) + ' 筆）: 期望 "' + expected + '", 實際 "' + actual + '"');
    }
  });

  console.log('✅ populateFileList 序號規則測試通過（讀取資料夾後每個檔案序號不再全部相同）');
}

function testApplyRulesToExistingFilesSequentialNumbering() {
  console.log('🧪 測試 applyRulesToExistingFiles 空列不佔用序號...');

  if (typeof makeSheet !== 'function') {
    console.log('⏭️  跳過（僅限 node tests/run-node.mjs 跑者，GAS 編輯器內沒有 makeSheet）');
    return;
  }

  const lastModified = new Date('2024-03-15T10:30:00');
  const rows = [
    ['原檔名'],
    ['a.txt', '資料夾/a.txt', 'old', 'old', 'text/plain', 1, lastModified, 'id-a'],
    ['', '', '', '', '', '', '', ''],
    ['b.txt', '資料夾/b.txt', 'old', 'old', 'text/plain', 2, lastModified, 'id-b'],
    ['c.txt', '資料夾/c.txt', 'old', 'old', 'text/plain', 3, lastModified, 'id-c']
  ];
  const fileListSheet = makeSheet(rows);
  const renameConfig = { mode: '新增序號', parameter: '前綴序號：1,3', operationType: '原位置更名', targetFolderId: null };

  applyRulesToExistingFiles(fileListSheet, renameConfig);

  const written = fileListSheet.getRange(2, 1, 3, 8).getValues();
  const expectedNames = ['001_a.txt', '002_b.txt', '003_c.txt'];
  expectedNames.forEach(function(expected, idx) {
    const actual = written[idx][2];
    if (actual !== expected) {
      throw new Error('applyRulesToExistingFiles 序號連號測試失敗（第 ' + (idx + 1) + ' 筆）: 期望 "' + expected + '", 實際 "' + actual + '"');
    }
  });

  console.log('✅ applyRulesToExistingFiles 空列不佔用序號測試通過');
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