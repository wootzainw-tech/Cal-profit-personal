/**
 * ExpenseTracker Pro - Google Apps Script (Realtime-safe version)
 * Deploy as Web App: Execute as Me, Who has access: Anyone
 */

function getTargetSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  var files = DriveApp.getFilesByName('ExpenseTracker_Data');
  if (files.hasNext()) return SpreadsheetApp.open(files.next());

  return SpreadsheetApp.create('ExpenseTracker_Data');
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureSheets_(ss) {
  var txSheet = ss.getSheetByName('Transactions');
  if (!txSheet) txSheet = ss.insertSheet('Transactions');

  var txHeaders = ['ID', 'Date', 'Type', 'Category ID', 'Category/Sub', 'Amount', 'Note'];
  if (txSheet.getLastRow() === 0) {
    txSheet.appendRow(txHeaders);
  } else {
    // รองรับชีตเก่าที่มี 6 คอลัมน์ โดยเพิ่ม Category ID เป็นคอลัมน์ D
    var currentHeaders = txSheet.getRange(1, 1, 1, Math.max(txSheet.getLastColumn(), 1)).getValues()[0];
    if (String(currentHeaders[3] || '') !== 'Category ID') {
      txSheet.insertColumnAfter(3);
      txSheet.getRange(1, 1, 1, 7).setValues([txHeaders]);
    } else {
      txSheet.getRange(1, 1, 1, 7).setValues([txHeaders]);
    }
  }
  txSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');

  var configSheet = ss.getSheetByName('Config');
  if (!configSheet) configSheet = ss.insertSheet('Config');
  if (configSheet.getLastRow() === 0) configSheet.appendRow(['Key', 'Value']);
  configSheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
  configSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');

  return { txSheet: txSheet, configSheet: configSheet };
}

function setup() {
  var ss = getTargetSpreadsheet();
  var sheets = ensureSheets_(ss);
  migrateMissingCategoryIds_(sheets.txSheet);
  Logger.log('ตั้งค่าเรียบร้อย Spreadsheet ID: ' + ss.getId());
}

function inferCategoryId_(sub, note, type) {
  if (String(type || '').toLowerCase() === 'income') return 'income';

  var text = (String(sub || '') + ' ' + String(note || '')).toLowerCase().trim();
  if (!text) return '';

  var prefix = text.match(/^\s*([1-7])(?:\.|\s|$)/);
  if (prefix) return Number(prefix[1]);

  var groups = {
    1: ['อาหาร', 'ข้าว', 'กาแฟ', 'เครื่องดื่ม', 'ของกิน', 'ขนม'],
    2: ['เบียร์', 'ปาร์ตี้', 'สังสรรค์'],
    3: ['ยา', 'เวชภัณฑ์', 'ของใช้', 'เบ็ดเตล็ด', 'อื่นๆ', 'อื่น ๆ'],
    4: ['เงินเก็บ', 'เก็บส่วนตัว', 'เก็บเที่ยว', 'กองกลาง', 'ออม'],
    5: ['ค่าไฟ', 'ไฟฟ้า', 'ค่าน้ำ', 'น้ำประปา'],
    6: ['น้ำมัน', 'ทางด่วน'],
    7: ['อินเทอร์เน็ต', 'ค่าเน็ต', 'เน็ตบ้าน', 'เน็ตมือถือ', 'wifi', 'wi-fi']
  };

  for (var id in groups) {
    for (var i = 0; i < groups[id].length; i++) {
      if (text.indexOf(groups[id][i]) !== -1) return Number(id);
    }
  }
  return '';
}

function migrateMissingCategoryIds_(txSheet) {
  if (!txSheet || txSheet.getLastRow() < 2) return;
  var range = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 7);
  var values = range.getValues();
  var changed = false;

  values.forEach(function(row) {
    if (row[0] === '' || row[3] !== '') return;
    var inferred = inferCategoryId_(row[4], row[6], row[2]);
    if (inferred !== '') {
      row[3] = inferred;
      changed = true;
    }
  });

  if (changed) range.setValues(values);
}

function doGet(e) {
  try {
    var ss = getTargetSpreadsheet();
    var sheets = ensureSheets_(ss);
    migrateMissingCategoryIds_(sheets.txSheet);

    var txData = sheets.txSheet.getDataRange().getValues();
    var transactions = txData.length > 1 ? txData.slice(1)
      .filter(function(row) { return row[0] !== ''; })
      .map(function(row) {
        var rawCatId = row[3];
        var inferredCatId = rawCatId === '' ? inferCategoryId_(row[4], row[6], row[2]) : rawCatId;
        var catId = inferredCatId === 'income' ? 'income' : (inferredCatId === '' ? null : Number(inferredCatId));
        return {
          id: String(row[0] || ''),
          date: row[1] instanceof Date ? row[1].toISOString() : String(row[1] || ''),
          type: String(row[2] || 'expense'),
          catId: catId,
          sub: String(row[4] || ''),
          amount: Number(row[5] || 0),
          note: String(row[6] || '')
        };
      }) : [];

    var configData = sheets.configSheet.getDataRange().getValues();
    var config = {};
    if (configData.length > 1) {
      configData.slice(1).forEach(function(row) {
        if (!row[0]) return;
        try {
          config[row[0]] = JSON.parse(row[1]);
        } catch (err) {
          config[row[0]] = row[1];
        }
      });
    }

    return jsonOutput_({
      status: 'success',
      transactions: transactions,
      config: config,
      serverTime: new Date().toISOString()
    });
  } catch (err) {
    return jsonOutput_({ status: 'error', message: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (!e || !e.postData || !e.postData.contents) {
      return jsonOutput_({ status: 'error', message: 'No post data' });
    }

    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var ss = getTargetSpreadsheet();
    var sheets = ensureSheets_(ss);

    if (action === 'upsertTransaction') {
      var tx = data.transaction;
      if (!tx || !tx.id) return jsonOutput_({ status: 'error', message: 'Transaction ID is required' });

      var rows = sheets.txSheet.getDataRange().getValues();
      var targetRow = -1;
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(tx.id)) {
          targetRow = i + 1;
          break;
        }
      }

      var values = [[
        String(tx.id),
        String(tx.date || new Date().toISOString()),
        String(tx.type || 'expense'),
        tx.catId === undefined || tx.catId === null ? '' : tx.catId,
        String(tx.sub || ''),
        Number(tx.amount || 0),
        String(tx.note || '')
      ]];

      if (targetRow > 0) {
        sheets.txSheet.getRange(targetRow, 1, 1, 7).setValues(values);
      } else {
        sheets.txSheet.getRange(sheets.txSheet.getLastRow() + 1, 1, 1, 7).setValues(values);
      }

      SpreadsheetApp.flush();
      return jsonOutput_({ status: 'success', message: 'Transaction saved' });
    }

    if (action === 'saveConfig') {
      var config = data.config || {};
      var configRows = sheets.configSheet.getDataRange().getValues();
      var rowByKey = {};
      for (var r = 1; r < configRows.length; r++) rowByKey[String(configRows[r][0])] = r + 1;

      Object.keys(config).forEach(function(key) {
        var value = JSON.stringify(config[key]);
        if (rowByKey[key]) {
          sheets.configSheet.getRange(rowByKey[key], 1, 1, 2).setValues([[key, value]]);
        } else {
          sheets.configSheet.appendRow([key, value]);
        }
      });

      SpreadsheetApp.flush();
      return jsonOutput_({ status: 'success', message: 'Config saved' });
    }

    if (action === 'delete') {
      var txRows = sheets.txSheet.getDataRange().getValues();
      var targetId = String(data.id || '');
      for (var j = 1; j < txRows.length; j++) {
        if (String(txRows[j][0]) === targetId) {
          sheets.txSheet.deleteRow(j + 1);
          return jsonOutput_({ status: 'success', deleted: true });
        }
      }
      return jsonOutput_({ status: 'success', deleted: false });
    }

    if (action === 'clearAll') {
      sheets.txSheet.clearContents();
      sheets.configSheet.clearContents();
      ensureSheets_(ss);
      return jsonOutput_({ status: 'success', message: 'All data cleared' });
    }

    // รองรับเว็บเวอร์ชันเก่า โดย MERGE ทีละ ID แทน clearContents()
    if (action === 'sync') {
      (data.transactions || []).forEach(function(tx) {
        var postEvent = { postData: { contents: JSON.stringify({ action: 'upsertTransaction', transaction: tx }) } };
        // ไม่เรียก doPost ซ้ำเพราะติด lock; ทำ merge โดยตรงผ่าน helper แบบง่าย
        var currentRows = sheets.txSheet.getDataRange().getValues();
        var foundRow = -1;
        for (var k = 1; k < currentRows.length; k++) {
          if (String(currentRows[k][0]) === String(tx.id)) { foundRow = k + 1; break; }
        }
        var rowValue = [[String(tx.id), String(tx.date || ''), String(tx.type || 'expense'), tx.catId == null ? '' : tx.catId, String(tx.sub || ''), Number(tx.amount || 0), String(tx.note || '')]];
        if (foundRow > 0) sheets.txSheet.getRange(foundRow, 1, 1, 7).setValues(rowValue);
        else sheets.txSheet.getRange(sheets.txSheet.getLastRow() + 1, 1, 1, 7).setValues(rowValue);
      });
      if (data.config) {
        var oldConfigRows = sheets.configSheet.getDataRange().getValues();
        var oldRowByKey = {};
        for (var q = 1; q < oldConfigRows.length; q++) oldRowByKey[String(oldConfigRows[q][0])] = q + 1;
        Object.keys(data.config).forEach(function(key) {
          var encoded = JSON.stringify(data.config[key]);
          if (oldRowByKey[key]) sheets.configSheet.getRange(oldRowByKey[key], 1, 1, 2).setValues([[key, encoded]]);
          else sheets.configSheet.appendRow([key, encoded]);
        });
      }
      return jsonOutput_({ status: 'success', message: 'Legacy data merged safely' });
    }

    return jsonOutput_({ status: 'error', message: 'Invalid action parameter' });
  } catch (err) {
    return jsonOutput_({ status: 'error', message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}
