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

var SHEETS = {
  pengunjung: 'Per Pengunjung',
  brosur: 'Rekap Brosur',
  sesi: 'Sesi'
};

var HEADERS = {
  pengunjung: [
    'Tanggal',
    'Jam',
    'Nama Pengunjung',
    'Nomor HP',
    'Jumlah Orang',
    'Kategori Interaksi',
    'Catatan'
  ],
  brosur: [
    'Tanggal',
    'Jam',
    'Brosur Dibawa',
    'Brosur Sisa',
    'Brosur Tersebar',
    'Catatan'
  ],
  sesi: [
    'Tanggal',
    'Jam Mulai',
    'Jam Selesai',
    'Total Pengunjung',
    'Total Brosur',
    'Catatan'
  ]
};

function doPost(e) {
  try {
    var payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    ensureRequiredSheets(spreadsheet);

    if (payload.type === 'pengunjung') {
      appendPengunjung(spreadsheet, payload);
    } else if (payload.type === 'brosur') {
      appendBrosur(spreadsheet, payload);
    } else if (payload.type === 'sesi') {
      upsertSesi(spreadsheet, payload);
    } else {
      throw new Error('Unknown payload type: ' + payload.type);
    }

    return jsonResponse({ status: 'ok' });
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.message });
  }
}

function appendPengunjung(spreadsheet, payload) {
  var sheet = ensureSheet(spreadsheet, SHEETS.pengunjung, HEADERS.pengunjung);

  sheet.appendRow([
    payload.tanggal || currentDate(spreadsheet),
    payload.waktu || currentTime(spreadsheet),
    payload.nama_pengunjung || '',
    payload.no_hp || '',
    cleanNumber(payload.jumlah_orang, 1),
    payload.kategori || '',
    payload.catatan || ''
  ]);
}

function appendBrosur(spreadsheet, payload) {
  var sheet = ensureSheet(spreadsheet, SHEETS.brosur, HEADERS.brosur);
  var dibawa = cleanNumber(payload.dibawa, 0);
  var sisa = cleanNumber(payload.sisa, 0);
  var tersebar = payload.tersebar !== undefined && payload.tersebar !== null && payload.tersebar !== ''
    ? cleanNumber(payload.tersebar, dibawa - sisa)
    : dibawa - sisa;

  sheet.appendRow([
    payload.tanggal || currentDate(spreadsheet),
    payload.waktu || currentTime(spreadsheet),
    dibawa,
    sisa,
    tersebar,
    payload.catatan || ''
  ]);
}

function upsertSesi(spreadsheet, payload) {
  var sheet = ensureSheet(spreadsheet, SHEETS.sesi, HEADERS.sesi);
  var tanggal = payload.tanggal || currentDate(spreadsheet);
  var sessionKey = normalizeDateKey(tanggal);
  var totals = getSessionTotals(spreadsheet, sessionKey);
  var existingRows = findRowsByDateKey(sheet, sessionKey);
  var targetRow = existingRows.length ? existingRows[0] : null;
  var jamMulai = payload.waktu_mulai || currentTime(spreadsheet);

  if (targetRow) {
    var existingJamMulai = sheet.getRange(targetRow, 2).getValue();
    if (existingJamMulai) {
      jamMulai = existingJamMulai;
    }
  }

  var rowData = [
    tanggal,
    jamMulai,
    payload.waktu_selesai || currentTime(spreadsheet),
    totals.totalPengunjung,
    totals.totalBrosur,
    payload.catatan || ''
  ];

  if (targetRow) {
    sheet.getRange(targetRow, 1, 1, HEADERS.sesi.length).setValues([rowData]);
    deleteDuplicateSessionRows(sheet, existingRows);
  } else {
    sheet.appendRow(rowData);
  }
}

function getSessionTotals(spreadsheet, sessionKey) {
  var pengunjungSheet = ensureSheet(spreadsheet, SHEETS.pengunjung, HEADERS.pengunjung);
  var brosurSheet = ensureSheet(spreadsheet, SHEETS.brosur, HEADERS.brosur);

  return {
    totalPengunjung: countRowsByDateKey(pengunjungSheet, sessionKey),
    totalBrosur: sumRowsByDateKey(brosurSheet, sessionKey, 5)
  };
}

function countRowsByDateKey(sheet, dateKey) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var total = 0;

  values.forEach(function(row) {
    if (normalizeDateKey(row[0]) === dateKey) {
      total += 1;
    }
  });

  return total;
}

function sumRowsByDateKey(sheet, dateKey, amountColumn) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  var values = sheet.getRange(2, 1, lastRow - 1, Math.max(amountColumn, 1)).getValues();
  var total = 0;

  values.forEach(function(row) {
    if (normalizeDateKey(row[0]) === dateKey) {
      total += cleanNumber(row[amountColumn - 1], 0);
    }
  });

  return total;
}

function findRowsByDateKey(sheet, dateKey) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var rows = [];

  values.forEach(function(row, index) {
    if (normalizeDateKey(row[0]) === dateKey) {
      rows.push(index + 2);
    }
  });

  return rows;
}

function deleteDuplicateSessionRows(sheet, rows) {
  if (rows.length <= 1) return;

  for (var index = rows.length - 1; index >= 1; index -= 1) {
    sheet.deleteRow(rows[index]);
  }
}

function ensureRequiredSheets(spreadsheet) {
  ensureSheet(spreadsheet, SHEETS.pengunjung, HEADERS.pengunjung);
  ensureSheet(spreadsheet, SHEETS.brosur, HEADERS.brosur);
  ensureSheet(spreadsheet, SHEETS.sesi, HEADERS.sesi);
}

function ensureSheet(spreadsheet, name, headers) {
  var sheet = spreadsheet.getSheetByName(name);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }

  var maxColumns = sheet.getMaxColumns();
  if (maxColumns < headers.length) {
    sheet.insertColumnsAfter(maxColumns, headers.length - maxColumns);
  } else if (maxColumns > headers.length) {
    sheet.deleteColumns(headers.length + 1, maxColumns - headers.length);
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  return sheet;
}

function normalizeDateKey(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var text = String(value).trim();
  var slashDate = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashDate) {
    return slashDate[3] + '-' + pad(slashDate[2]) + '-' + pad(slashDate[1]);
  }

  var dashDate = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (dashDate) {
    return dashDate[1] + '-' + pad(dashDate[2]) + '-' + pad(dashDate[3]);
  }

  return text;
}

function currentDate(spreadsheet) {
  return Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd/MM/yyyy');
}

function currentTime(spreadsheet) {
  return Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'HH:mm');
}

function cleanNumber(value, fallback) {
  var number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function pad(value) {
  return String(value).padStart(2, '0');
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
