/**
 * MTP MLEB — Pencatatan Pengunjung Google Apps Script
 *
 * Setup instructions:
 * 1. Create a new Google Sheet for MTP MLEB data.
 * 2. In the Sheet, open Extensions > Apps Script.
 * 3. Paste this entire file into the Apps Script editor.
 * 4. Save the project.
 * 5. Click Deploy > New deployment.
 * 6. Select deployment type: Web app.
 * 7. Set "Execute as" to "Me".
 * 8. Set "Who has access" to "Anyone".
 * 9. Click Deploy, authorize access if prompted, and copy the Web App URL.
 * 10. Paste the Web App URL into the Pengaturan tab of index.html.
 *
 * The web app receives JSON POST requests from index.html using fetch mode: no-cors.
 */

function doPost(e) {
  try {
    var payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (payload.type === 'pengunjung') {
      appendPengunjung(spreadsheet, payload);
    } else if (payload.type === 'brosur') {
      appendBrosur(spreadsheet, payload);
    } else if (payload.type === 'sesi') {
      appendSesi(spreadsheet, payload);
    } else {
      throw new Error('Unknown payload type: ' + payload.type);
    }

    return jsonResponse({ status: 'ok' });
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.message });
  }
}

function appendPengunjung(spreadsheet, payload) {
  var sheetName = 'MTP_' + String(payload.tanggal || '').replace(/\//g, '-');
  var sheet = getOrCreateSheet(spreadsheet, sheetName, [
    'Tanggal',
    'Waktu',
    'Nama Pengunjung',
    'No HP',
    'Jumlah Orang',
    'Kategori',
    'Catatan'
  ]);

  sheet.appendRow([
    payload.tanggal || '',
    payload.waktu || '',
    payload.nama_pengunjung || '',
    payload.no_hp || '',
    payload.jumlah_orang || '',
    payload.kategori || '',
    payload.catatan || ''
  ]);
}

function appendBrosur(spreadsheet, payload) {
  var sheet = getOrCreateSheet(spreadsheet, 'Brosur', [
    'Tanggal',
    'Nama Tim',
    'Dibawa',
    'Sisa',
    'Tersebar'
  ]);

  sheet.appendRow([
    payload.tanggal || '',
    payload.nama_tim || '',
    payload.dibawa || '',
    payload.sisa || '',
    payload.tersebar || ''
  ]);
}

function appendSesi(spreadsheet, payload) {
  var sheet = getOrCreateSheet(spreadsheet, 'Rekap Sesi', [
    'Tanggal',
    'Waktu Mulai',
    'Waktu Selesai',
    'Total Entri'
  ]);

  sheet.appendRow([
    payload.tanggal || '',
    payload.waktu_mulai || '',
    payload.waktu_selesai || '',
    payload.total_entri || 0
  ]);
}

function getOrCreateSheet(spreadsheet, name, headers) {
  var sheet = spreadsheet.getSheetByName(name);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  return sheet;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
