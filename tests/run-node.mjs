#!/usr/bin/env node
// 本機測試入口：用 node:vm 把 GAS 原始碼載入一個假的執行環境跑 tests/test-functions.js 的 runTests()。
// 只用 Node 內建模組，不需要 npm install；src/、tests/test-functions.js 不可反過來依賴這支檔案
// （它們必須能整段貼進真正的 GAS 編輯器執行，所以不能用 import/export 或其他 node 專屬 API）。
import vm from 'node:vm';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const TIME_ZONE = 'Asia/Taipei';

// ---- Utilities / Session 打樁：固定時區 Asia/Taipei ----

function getTaipeiParts(date) {
  const fmt = new Intl.DateTimeFormat('en-US', {
    timeZone: TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  });
  const parts = {};
  for (const p of fmt.formatToParts(date)) {
    parts[p.type] = p.value === '24' && p.type === 'hour' ? '00' : p.value;
  }
  return parts;
}

function formatDate(date, _timeZone, pattern) {
  const p = getTaipeiParts(date);
  switch (pattern) {
    case 'yyyy-MM-dd':
      return `${p.year}-${p.month}-${p.day}`;
    case 'yyyyMMdd':
      return `${p.year}${p.month}${p.day}`;
    case 'dd-MM-yyyy':
      return `${p.day}-${p.month}-${p.year}`;
    case 'yyyy-MM-dd HH:mm':
      return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}`;
    default:
      // 保底：手動代換常見 token，不支援的 pattern 至少不會整包壞掉
      return pattern
        .replace(/yyyy/g, p.year)
        .replace(/MM/g, p.month)
        .replace(/dd/g, p.day)
        .replace(/HH/g, p.hour)
        .replace(/mm/g, p.minute)
        .replace(/ss/g, p.second);
  }
}

const Utilities = { formatDate };
const Session = { getScriptTimeZone: () => TIME_ZONE };

// ---- makeSheet(rows)：假 Sheet，支援 populateFileList/executeRenaming 等會用到的 API ----

function colLetterToIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function parseA1(a1) {
  const m = /^([A-Z]+)(\d+)$/.exec(a1);
  if (!m) throw new Error(`makeSheet：無法解析 A1 表示法 "${a1}"`);
  return { row: parseInt(m[2], 10), col: colLetterToIndex(m[1]) };
}

function makeSheet(rows) {
  // rows：二維陣列，rows[0] 對應試算表第 1 列（標題列），1-indexed 換算成 row-1
  const data = (rows || []).map(r => r.slice());

  function ensureRow(idx) {
    while (data.length <= idx) data.push([]);
    return data[idx];
  }

  return {
    _data: data,
    getLastRow() {
      // 找出最後一列「有內容」的列號（GAS 的 getLastRow 語意）
      for (let i = data.length - 1; i >= 0; i--) {
        if (data[i] && data[i].some(v => v !== '' && v !== undefined && v !== null)) {
          return i + 1;
        }
      }
      return 0;
    },
    getLastColumn() {
      return data.reduce((max, r) => Math.max(max, r.length), 0);
    },
    getRange(...args) {
      if (args.length === 1 && typeof args[0] === 'string') {
        const { row, col } = parseA1(args[0]);
        return {
          getValue() {
            const r = data[row - 1];
            const v = r ? r[col - 1] : undefined;
            return v === undefined ? '' : v;
          },
          setValue(v) {
            ensureRow(row - 1)[col - 1] = v;
          }
        };
      }
      const [row, col, numRows = 1, numCols = 1] = args;
      return {
        getValues() {
          const out = [];
          for (let i = 0; i < numRows; i++) {
            const r = data[row - 1 + i] || [];
            const line = [];
            for (let j = 0; j < numCols; j++) {
              const v = r[col - 1 + j];
              line.push(v === undefined ? '' : v);
            }
            out.push(line);
          }
          return out;
        },
        setValues(values) {
          for (let i = 0; i < values.length; i++) {
            const r = ensureRow(row - 1 + i);
            for (let j = 0; j < values[i].length; j++) {
              r[col - 1 + j] = values[i][j];
            }
          }
        },
        setValue(v) {
          // 單一儲存格（numRows/numCols 皆為 1）也支援 setValue
          ensureRow(row - 1)[col - 1] = v;
        },
        clearContent() {
          for (let i = 0; i < numRows; i++) {
            const r = data[row - 1 + i];
            if (!r) continue;
            for (let j = 0; j < numCols; j++) r[col - 1 + j] = '';
          }
        }
      };
    }
  };
}

// ---- DriveApp 假物件：記錄 setName／makeCopy 呼叫，供 WP2–WP4 測試使用 ----

function makeFakeDriveApp() {
  const calls = { setName: [], makeCopy: [], getFileById: [], getFolderById: [] };
  const files = new Map();
  const folders = new Map();
  let copySeq = 0;

  return {
    _calls: calls,
    _reset() {
      calls.setName.length = 0;
      calls.makeCopy.length = 0;
      calls.getFileById.length = 0;
      calls.getFolderById.length = 0;
      files.clear();
      folders.clear();
      copySeq = 0;
    },
    _registerFile(id, attrs) {
      files.set(id, Object.assign({ id, name: id, mimeType: 'application/octet-stream', size: 0, lastModified: new Date() }, attrs));
    },
    _registerFolder(id, name, fileIds) {
      folders.set(id, { id, name, fileIds: fileIds || [] });
    },
    getFileById(id) {
      calls.getFileById.push(id);
      const rec = files.get(id);
      if (!rec) throw new Error(`找不到檔案：${id}`);
      return {
        getId: () => rec.id,
        getName: () => rec.name,
        setName(newName) {
          calls.setName.push({ id, newName });
          rec.name = newName;
        },
        makeCopy(newName, targetFolder) {
          copySeq += 1;
          const newId = `copy-${copySeq}-${id}`;
          calls.makeCopy.push({ sourceId: id, newName, targetFolderId: targetFolder && targetFolder._id });
          files.set(newId, Object.assign({}, rec, { id: newId, name: newName }));
          return { getId: () => newId, getName: () => newName };
        }
      };
    },
    getFolderById(id) {
      calls.getFolderById.push(id);
      const rec = folders.get(id);
      if (!rec) throw new Error(`找不到資料夾：${id}`);
      return {
        _id: id,
        getName: () => rec.name,
        getFiles() {
          let idx = 0;
          return {
            hasNext: () => idx < rec.fileIds.length,
            next: () => {
              const f = files.get(rec.fileIds[idx]);
              idx += 1;
              return {
                getId: () => f.id,
                getName: () => f.name,
                getMimeType: () => f.mimeType,
                getSize: () => f.size,
                getLastUpdated: () => f.lastModified
              };
            }
          };
        }
      };
    }
  };
}

// ---- SpreadsheetApp 假物件：populateFileList 會用 getActiveSpreadsheet().getSheetByName('指令區') 拿設定 ----

function makeFakeSpreadsheetApp() {
  const sheetsByName = {};
  return {
    _setSheet(name, sheet) {
      sheetsByName[name] = sheet;
    },
    getActiveSpreadsheet() {
      return {
        getSheetByName(name) {
          return sheetsByName[name] || null;
        }
      };
    }
  };
}

const DriveApp = makeFakeDriveApp();
const SpreadsheetApp = makeFakeSpreadsheetApp();

// ---- 載入 src/*.js 與 tests/test-functions.js 進同一個 vm context ----

const context = {
  console,
  Utilities,
  Session,
  DriveApp,
  SpreadsheetApp,
  makeSheet,
  Date,
  Math,
  JSON,
  RegExp,
  parseInt,
  parseFloat,
  String,
  Number,
  Boolean,
  Array,
  Object,
  isNaN
};
vm.createContext(context);

const filesToLoad = ['src/Code.js', 'src/FileOperations.js', 'tests/test-functions.js'];
for (const rel of filesToLoad) {
  const full = path.join(ROOT, rel);
  const code = fs.readFileSync(full, 'utf8');
  vm.runInContext(code, context, { filename: full });
}

let ok = true;
try {
  vm.runInContext('runTests()', context, { filename: 'run-node.mjs::runTests()' });
} catch (error) {
  console.error('❌ runTests() 拋出例外，測試套件視為失敗：', error && error.message ? error.message : error);
  ok = false;
}

process.exit(ok ? 0 : 1);
