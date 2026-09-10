/**
 * TNBES Energy Audit Dashboard — Google Sheets backend
 * -------------------------------------------------------
 * Simpan skrip ini di Google Apps Script (script.google.com) yang
 * dipautkan kepada Google Sheet anda, kemudian "Deploy > New deployment"
 * sebagai Web App. Salin URL Web App yang dihasilkan dan tampal pada
 * pemalar GOOGLE_SHEETS_WEBAPP_URL dalam fail app.html.
 *
 * Cara kerja: setiap "key" (contohnya senarai projek, data lantai, dsb.)
 * disimpan sebagai SATU baris dalam sheet bernama "Data", dengan lajur:
 *   A: Key | B: Value (JSON string) | C: UpdatedAt
 */

const SHEET_NAME = 'Data';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Key', 'Value', 'UpdatedAt']);
  }
  return sheet;
}

function findRow_(sheet, key) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return i + 1; // 1-indexed row number
  }
  return -1;
}

function doGet(e) {
  e = e || { parameter: {} };
  const action = e.parameter.action;
  const sheet = getSheet_();

  if (action === 'get') {
    const key = e.parameter.key;
    const row = findRow_(sheet, key);
    if (row === -1) {
      return jsonOut_({ ok: true, value: null });
    }
    const value = sheet.getRange(row, 2).getValue();
    return jsonOut_({ ok: true, value: String(value) });
  }

  if (action === 'list') {
    const data = sheet.getDataRange().getValues();
    const keys = [];
    for (let i = 1; i < data.length; i++) keys.push(data[i][0]);
    return jsonOut_({ ok: true, keys: keys });
  }

  if (action === 'dump_all') {
    // Pulangkan SEMUA key-value dalam SATU panggilan (untuk Python compile tool)
    const data = sheet.getDataRange().getValues();
    const out = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) out[data[i][0]] = String(data[i][1]);
    }
    return jsonOut_({ ok: true, data: out });
  }

  return jsonOut_({ ok: false, error: 'Unknown action' });
}

function doPost(e) {
  e = e || {};
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'Invalid JSON body' });
  }

  const sheet = getSheet_();

  if (body.action === 'set') {
    const key = body.key;
    const value = body.value;
    const row = findRow_(sheet, key);
    const now = new Date();
    if (row === -1) {
      sheet.appendRow([key, value, now]);
    } else {
      sheet.getRange(row, 2, 1, 2).setValues([[value, now]]);
    }
    return jsonOut_({ ok: true });
  }

  if (body.action === 'delete') {
    const row = findRow_(sheet, body.key);
    if (row !== -1) sheet.deleteRow(row);
    return jsonOut_({ ok: true });
  }

  return jsonOut_({ ok: false, error: 'Unknown action' });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
