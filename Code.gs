/**
 * Backend ฟรีสำหรับ "ระบบจองรถออนไลน์"
 * ทำหน้าที่เป็นฐานข้อมูล key-value โดยใช้ Google Sheet แท็บชื่อ "KV"
 * คอลัมน์ A = key, คอลัมน์ B = value (เก็บเป็นข้อความ JSON)
 *
 * วิธีใช้: วางโค้ดนี้ใน Extensions > Apps Script ของชีตที่มีแท็บชื่อ "KV"
 * แล้ว Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone
 * คัดลอก URL ที่ได้ไปใส่ในไฟล์ HTML ตรง APPS_SCRIPT_URL
 */

function getSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KV');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('KV');
    sheet.appendRow(['key', 'value']);
  }
  return sheet;
}

function findRow_(sheet, key) {
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) return i + 1; // 1-indexed row number
  }
  return -1;
}

function doGet(e) {
  var key = e.parameter.key;
  if (key === '__ping__') {
    return ContentService.createTextOutput(JSON.stringify({ version: 'v2-fixed-2026-08-24' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sheet = getSheet_();
  var row = findRow_(sheet, key);
  var value = row === -1 ? 'null' : sheet.getRange(row, 2).getValue();
  return ContentService.createTextOutput(value)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var key = body.key;
  var value = body.value;
  var sheet = getSheet_();
  var row = findRow_(sheet, key);
  if (row === -1) {
    sheet.appendRow([key, value]);
  } else {
    sheet.getRange(row, 2).setValue(value);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
