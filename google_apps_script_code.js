/**
 * Google Apps Script for ExpenseTracker Pro
 */

function getTargetSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  return sheet;
}

// ฟังก์ชันสำหรับกด "เรียกใช้" (Run) ในหน้านี้เพื่อขอสิทธิ์การใช้งานครั้งแรก
function setup() {
  var sheet = getTargetSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', 'Date', 'Type', 'Category/Sub', 'Amount', 'Note']);
    var headerRange = sheet.getRange(1, 1, 1, 6);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#0284c7');
    headerRange.setFontColor('#ffffff');
  }
  Logger.log("ตั้งค่าการเชื่อมต่อ Google Sheet สำเร็จ!");
}

function doGet(e) {
  try {
    var sheet = getTargetSheet();
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        transactions: [] 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var rows = data.slice(1);
    var transactions = rows.map(function(row) {
      return {
        id: String(row[0] || ''),
        date: row[1] || '',
        type: row[2] || 'expense',
        sub: row[3] || '',
        amount: Number(row[4] || 0),
        note: row[5] || ''
      };
    });

    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      transactions: transactions 
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
    // ป้องกัน Error กรณีไม่มีข้อมูล e ส่งเข้ามา
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: 'No data received. Please test via Web App.' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = getTargetSheet();
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    var action = data.action;

    if (action === 'sync') {
      sheet.clearContents();
      sheet.appendRow(['ID', 'Date', 'Type', 'Category/Sub', 'Amount', 'Note']);
      
      var headerRange = sheet.getRange(1, 1, 1, 6);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#0284c7');
      headerRange.setFontColor('#ffffff');

      if (data.transactions && Array.isArray(data.transactions)) {
        data.transactions.forEach(function(tx) {
          sheet.appendRow([
            tx.id || '',
            tx.date || '',
            tx.type || 'expense',
            tx.sub || '',
            tx.amount || 0,
            tx.note || ''
          ]);
        });
      }

      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'Data synced successfully' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'delete') {
      var rows = sheet.getDataRange().getValues();
      var targetId = String(data.id);
      var deleted = false;

      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === targetId) {
          sheet.deleteRow(i + 1);
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