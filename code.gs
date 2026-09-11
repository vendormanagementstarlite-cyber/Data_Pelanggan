// ============================================
// KONFIGURASI
// ============================================
const SPREADSHEET_ID = '1pl7JndJD6067Dt34xN5hpGcsF_KY28oJ1wcSWOh3Pr4';
const SHEET_LOGIN = 'Data Login Mitra';
const SHEET_DATA = 'Data Mitra';

// Kolom di sheet "Data Login Mitra"
const LOGIN_COL = {
  NAMA_MITRA: 0, // A
  USERNAME: 1,   // B
  PASSWORD: 2    // C
};

// Header kolom di sheet "Data Mitra" (harus persis sama urutannya dengan gsheet)
const DATA_HEADERS = [
  'Nama Pelanggan',
  'Nomor Telepon',
  'Alamat',
  'Nama Mitra',
  'Provinsi',
  'Kota',
  'Kecamatan',
  'Kelurahan',
  'Stasiun',
  'Longitude',
  'Latitude',
  'Status',
  'Remarks'
];

// Index kolom "Nama Mitra" di sheet Data Mitra (untuk filter)
const DATA_NAMA_MITRA_INDEX = 3; // kolom D

// ============================================
// ENTRY POINT
// ============================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portal Data Mitra')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// LOGIN
// ============================================
function login(username, password) {
  username = (username || '').toString().trim();
  password = (password || '').toString().trim();

  if (!username || !password) {
    return { success: false, message: 'Username dan password wajib diisi.' };
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_LOGIN);
  if (!sheet) {
    return { success: false, message: 'Sheet "' + SHEET_LOGIN + '" tidak ditemukan.' };
  }

  const values = sheet.getDataRange().getValues();
  // baris 0 = header, mulai dari baris 1
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowUsername = (row[LOGIN_COL.USERNAME] || '').toString().trim();
    const rowPassword = (row[LOGIN_COL.PASSWORD] || '').toString().trim();
    const namaMitra = (row[LOGIN_COL.NAMA_MITRA] || '').toString().trim();

    if (rowUsername.toLowerCase() === username.toLowerCase() && rowPassword === password) {
      return {
        success: true,
        namaMitra: namaMitra,
        username: rowUsername
      };
    }
  }

  return { success: false, message: 'Username atau password salah.' };
}

// ============================================
// AMBIL DATA SESUAI MITRA YANG LOGIN
// ============================================
function getDataMitra(namaMitra) {
  namaMitra = (namaMitra || '').toString().trim();
  if (!namaMitra) {
    return { success: false, message: 'Nama mitra tidak valid.', headers: DATA_HEADERS, rows: [] };
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DATA);
  if (!sheet) {
    return { success: false, message: 'Sheet "' + SHEET_DATA + '" tidak ditemukan.', headers: DATA_HEADERS, rows: [] };
  }

  const values = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowNamaMitra = (row[DATA_NAMA_MITRA_INDEX] || '').toString().trim();
    if (rowNamaMitra.toLowerCase() === namaMitra.toLowerCase()) {
      // ambil hanya sejumlah kolom sesuai DATA_HEADERS, dan format angka koordinat
      const cleanedRow = DATA_HEADERS.map((h, idx) => {
        let val = row[idx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        }
        return val === null || val === undefined ? '' : val.toString();
      });
      rows.push(cleanedRow);
    }
  }

  return { success: true, headers: DATA_HEADERS, rows: rows };
}
