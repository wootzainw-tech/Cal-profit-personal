/**
 * Google Apps Script for ExpenseTracker Pro
 * Connects directly to Google Sheet 'ExpenseTracker_Data' in Google Drive.
 * Handles both transaction records and configuration options (budgets, subcategories, extra incomes).
 */

function getTargetSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    return ss;
  }
  
  // Look for existing ExpenseTracker_Data file in Google Drive
  var files = DriveApp.getFilesByName('ExpenseTracker_Data');
  if (files.hasNext()) {
    var file = files.next();
    return SpreadsheetApp.open(file);
  }
  
  // Create a new spreadsheet if not found
  var newSs = SpreadsheetApp.create('ExpenseTracker_Data');
  var txSheet = newSs.getActiveSheet();
  txSheet.setName('Transactions');
  txSheet.appendRow(['ID', 'Date', 'Type', 'Category/Sub', 'Amount', 'Note']);
  var headerRange = txSheet.getRange(1, 1, 1, 6);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#0284c7');
  headerRange.setFontColor('#ffffff');

  var configSheet = newSs.insertSheet('Config');
  configSheet.appendRow(['Key', 'Value']);
  var configHeader = configSheet.getRange(1, 1, 1, 2);
  configHeader.setFontWeight('bold');
  configHeader.setBackground('#0284c7');
  configHeader.setFontColor('#ffffff');

  return newSs;
}

// Authorization setup function
function setup() {
  var ss = getTargetSpreadsheet();
  Logger.log("ตั้งค่าการเชื่อมต่อ Google Sheet สำเร็จ! Spreadsheet ID: " + ss.getId());
}

function doGet(e) {
  try {
    var ss = getTargetSpreadsheet();
    var txSheet = ss.getSheetByName('Transactions') || ss.insertSheet('Transactions');
    var configSheet = ss.getSheetByName('Config') || ss.insertSheet('Config');

    // Fetch transactions
    var txData = txSheet.getDataRange().getValues();
    var transactions = [];
    if (txData.length > 1) {
      transactions = txData.slice(1).map(function(row) {
        return {
          id: String(row[0] || ''),
          date: row[1] || '',
          type: row[2] || 'expense',
          sub: row[3] || '',
          amount: Number(row[4] || 0),
          note: row[5] || ''
        };
      });
    }

    // Fetch configurations
    var configData = configSheet.getDataRange().getValues();
    var config = {};
    if (configData.length > 1) {
      configData.slice(1).forEach(function(row) {
        if (row[0] && row[1]) {
          try {
            config[row[0]] = JSON.parse(row[1]);
          } catch(err) {
            config[row[0]] = row[1];
          }
        }
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      transactions: transactions,
      config: config
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "No post data" })).setMimeType(ContentService.MimeType.JSON);
    }

    var ss = getTargetSpreadsheet();
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    var action = data.action;

    if (action === 'sync') {
      // Sync Transactions
      var txSheet = ss.getSheetByName('Transactions') || ss.insertSheet('Transactions');
      txSheet.clearContents();
      txSheet.appendRow(['ID', 'Date', 'Type', 'Category/Sub', 'Amount', 'Note']);
      
      var headerRange = txSheet.getRange(1, 1, 1, 6);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#0284c7');
      headerRange.setFontColor('#ffffff');

      if (data.transactions && Array.isArray(data.transactions)) {
        data.transactions.forEach(function(tx) {
          txSheet.appendRow([
            tx.id || '',
            tx.date || '',
            tx.type || 'expense',
            tx.sub || '',
            tx.amount || 0,
            tx.note || ''
          ]);
        });
      }

      // Sync Config
      if (data.config) {
        var configSheet = ss.getSheetByName('Config') || ss.insertSheet('Config');
        configSheet.clearContents();
        configSheet.appendRow(['Key', 'Value']);
        
        var configHeader = configSheet.getRange(1, 1, 1, 2);
        configHeader.setFontWeight('bold');
        configHeader.setBackground('#0284c7');
        configHeader.setFontColor('#ffffff');

        Object.keys(data.config).forEach(function(key) {
          configSheet.appendRow([key, JSON.stringify(data.config[key])]);
        });
      }

      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'Data synced successfully' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'delete') {
      var txSheet = ss.getSheetByName('Transactions');
      if (txSheet) {
        var rows = txSheet.getDataRange().getValues();
        var targetId = String(data.id);
        var deleted = false;

        for (var i = 1; i < rows.length; i++) {
          if (String(rows[i][0]) === targetId) {
            txSheet.deleteRow(i + 1);
            deleted = true;
            break;
          }
        }

        return ContentService.createTextOutput(JSON.stringify({ 
          status: 'success', 
          deleted: deleted,
          message: deleted ? 'Row deleted' : 'ID not found' 
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'error', 
      message: 'Invalid action parameter' 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'error', 
      message: err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}